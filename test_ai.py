# -*- coding: utf-8 -*-
"""AI 채팅의 계산 부분 테스트 (화면 없음, 인터넷 · API 키 없음 — 가짜 서버를 내 PC 에 띄워서 확인).
실행: python test_ai.py"""
import json
import shutil
import tempfile
import unittest
from pathlib import Path

import ai_client
import ai_commands as ac
import ai_mock_server
import data_editor as ed

HERE = Path(__file__).resolve().parent
SECRET = "sk-test-must-not-leak-0123456789"


class Base(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="ai_test_"))
        for name in ("보금자리론_신청_2026_10.xlsx", "주택연금_가입현황_정제전.xlsx"):
            shutil.copy2(HERE / "samples_original" / name, self.dir / name)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def loan(self):
        src = ed.open_source(self.dir / "보금자리론_신청_2026_10.xlsx")
        return src.load(src.tables()[0])

    def pension(self):
        src = ed.open_source(self.dir / "주택연금_가입현황_정제전.xlsx")
        return src.load(src.tables()[0])


class 보내는내용(Base):
    def test_기본은_열_이름과_요약만_보내고_칸의_값은_보내지_않는다(self):
        s = self.loan()
        text = ac.describe(s, "신청내역")
        for name in s.names:
            self.assertIn(name, text)
        self.assertIn("240", text)                                   # 행 수
        first = s.rows[s.ids()[0]]
        self.assertNotIn(str(first[s.col("신청번호")]), text)          # 칸의 값은 없다
        self.assertNotIn("X99", text)
        with_rows = ac.describe(s, "신청내역", samples=3)
        self.assertIn(str(first[s.col("신청번호")]), with_rows)        # 켰을 때만 예시 행이 들어간다

    def test_안내문에_쓸_수_있는_명령이_모두_적혀_있다(self):
        prompt = ac.system_prompt(self.loan(), "신청내역")
        for op in ac.OPS:
            self.assertIn(f'"{op}"', prompt)
        self.assertIn("JSON", prompt)


class 답읽기(unittest.TestCase):
    def test_코드_울타리나_앞뒤_말이_있어도_읽는다(self):
        plan = {"reply": "X99 를 ON1 로 바꿉니다.", "actions": [{"op": "replace", "column": "지점코드", "find": "X99", "to": "ON1"}]}
        raw = json.dumps(plan, ensure_ascii=False)
        for text in (raw, "```json\n" + raw + "\n```", "네, 이렇게 하겠습니다.\n" + raw + "\n끝."):
            reply, actions = ac.parse_reply(text)
            self.assertEqual(reply, plan["reply"])
            self.assertEqual(actions, plan["actions"])

    def test_JSON_이_아니면_말로만_받는다(self):
        reply, actions = ac.parse_reply("어느 열을 말씀하시는지 알려 주세요.")
        self.assertEqual(reply, "어느 열을 말씀하시는지 알려 주세요.")
        self.assertEqual(actions, [])


class 명령실행(Base):
    def test_조건에_맞는_칸_바꾸기와_한_번에_되돌리기(self):
        s = self.loan()
        n_x99 = len(s.search("X99", "지점코드"))
        actions = [{"op": "replace", "column": "지점코드", "find": "X99", "to": "ON1"},
                   {"op": "set", "column": "처리상태", "value": "재심사",
                    "where": [{"column": "처리상태", "is": "=", "value": "반려"}, {"column": "신청금액(만원)", "is": ">=", "value": 30000}]}]
        lines = ac.preview(s, actions)
        self.assertEqual(len(lines), 2)
        self.assertIn(f"{n_x99}", lines[0])                           # 몇 칸이 바뀔지 미리 알려 준다
        self.assertEqual(s.change_count(), 0)                         # 미리 보기는 아무것도 바꾸지 않는다
        result = ac.run(s, actions)
        self.assertTrue(result.changed)
        self.assertEqual(len(s.search("X99", "지점코드")), 0)
        ci, st = s.col("신청금액(만원)"), s.col("처리상태")
        self.assertTrue(all(row[st] != "반려" for row in s.rows.values() if row[ci] >= 30000))
        self.assertTrue(any(row[st] == "반려" for row in s.rows.values()))       # 조건 밖은 그대로
        self.assertTrue(s.undo())                                     # 두 명령이 한 번에 되돌아간다
        self.assertEqual(s.change_count(), 0)

    def test_행_지우기_계산_열_전처리_명령(self):
        s = self.loan()
        before = len(s.rows)
        n = sum(1 for row in s.rows.values() if row[s.col("고객구분")] == "일반")
        result = ac.run(s, [
            {"op": "delete_rows", "where": [{"column": "고객구분", "is": "=", "value": "일반"}]},
            {"op": "compute", "target": "신청금액(억원)", "left": "신청금액(만원)", "operator": "/", "right": 10000},
            {"op": "rename_column", "column": "대출기간(년)", "to": "기간"},
        ])
        self.assertEqual(len(s.rows), before - n)
        self.assertIn("신청금액(억원)", s.names)
        self.assertIn("기간", s.names)
        row = s.rows[s.ids()[0]]
        self.assertAlmostEqual(row[s.col("신청금액(억원)")], row[s.col("신청금액(만원)")] / 10000)
        self.assertEqual(len(result.lines), 3)
        s.undo()
        self.assertEqual((len(s.rows), s.change_count()), (before, 0))

        p = self.pension()
        ac.run(p, [{"op": "strip"}, {"op": "drop_duplicates"},
                   {"op": "convert", "column": "월지급액", "kind": "number"},
                   {"op": "fill_empty", "column": "월지급액", "how": "median"}])
        self.assertEqual(len(p.rows), 60)
        self.assertEqual(p.empty_count("월지급액"), 0)
        self.assertEqual(p.columns[p.col("월지급액")].kind, "int")
        first = p.ids()[0]
        self.assertTrue(p.undo())                    # 칸 고치기 + 행 삭제 + 종류 바꾸기가 섞여 있어도 한 번에 제자리로
        self.assertEqual((len(p.rows), p.change_count(), p.columns[p.col("월지급액")].kind), (62, 0, "str"))
        self.assertEqual(p.ids()[0], first)
        self.assertEqual(p.ids(), sorted(p.ids()))   # 되살아난 행이 원래 자리에

    def test_보기와_집계는_표를_바꾸지_않는다(self):
        s = self.loan()
        result = ac.run(s, [
            {"op": "show", "where": [{"column": "상품명", "is": "contains", "value": "디딤돌"}]},
            {"op": "sort", "column": "신청금액(만원)", "descending": True},
            {"op": "stats", "column": "신청금액(만원)", "how": "mean", "by": "고객구분"},
            {"op": "stats", "column": "신청번호", "how": "count", "where": [{"column": "처리상태", "is": "=", "value": "승인"}]},
        ])
        self.assertFalse(result.changed)
        self.assertEqual(s.change_count(), 0)
        ci, pi = s.col("신청금액(만원)"), s.col("상품명")
        self.assertTrue(result.view and all("디딤돌" in s.rows[rid][pi] for rid in result.view))
        amounts = [s.rows[rid][ci] for rid in result.view]
        self.assertEqual(amounts, sorted(amounts, reverse=True))
        approved = sum(1 for row in s.rows.values() if row[s.col("처리상태")] == "승인")
        text = "\n".join(result.lines)
        self.assertIn(f"{approved:,}", text)
        for group in {row[s.col("고객구분")] for row in s.rows.values()}:
            self.assertIn(group, text)                                # 집단별 평균이 한 줄씩

    def test_잘못된_명령은_아무것도_바꾸지_않고_이유를_말한다(self):
        s = self.loan()
        for bad in ([{"op": "set", "column": "없는열", "value": 1}],
                    [{"op": "날려버려"}],
                    [{"op": "delete_rows"}],                                       # 조건 없는 삭제는 받지 않는다
                    [{"op": "set", "column": "신청금액(만원)", "value": "많이"}],     # 숫자 열에 글자
                    [{"op": "replace", "column": "지점코드", "find": "X99", "to": "ON1"},
                     {"op": "set", "column": "없는열", "value": 1}]):                # 뒤 명령이 틀리면 앞 명령도 취소
            with self.assertRaises(ValueError):
                ac.run(s, bad)
            self.assertEqual(s.change_count(), 0, bad)
            self.assertEqual(len(s.rows), 240)


class 통신(unittest.TestCase):
    """세 회사의 주소 · 머리말 · 본문 모양이 맞는지, 가짜 서버로 확인한다."""

    def ask(self, provider, server, model="test-model"):
        return ai_client.ask(provider, model, SECRET, "안내문", [{"role": "user", "content": "먼저 한 말"},
                                                              {"role": "assistant", "content": "먼저 한 답"},
                                                              {"role": "user", "content": "질문"}], base_url=server.url, timeout=10)

    def test_제미나이(self):
        server = ai_mock_server.start(lambda q: "답:" + q)
        try:
            self.assertEqual(self.ask("gemini", server, "gemini-flash-latest"), "답:질문")
            r = server.requests[0]
            self.assertTrue(r["path"].endswith("/v1beta/models/gemini-flash-latest:generateContent"))
            self.assertEqual(r["headers"]["x-goog-api-key"], SECRET)
            self.assertNotIn(SECRET, r["path"])                       # 키를 주소에 붙이지 않는다
            self.assertEqual(r["body"]["system_instruction"]["parts"][0]["text"], "안내문")
            self.assertEqual([c["role"] for c in r["body"]["contents"]], ["user", "model", "user"])
        finally:
            server.stop()

    def test_지피티와_직접_지정(self):
        for provider in ("openai", "custom"):
            server = ai_mock_server.start(lambda q: "답:" + q)
            try:
                self.assertEqual(self.ask(provider, server), "답:질문")
                r = server.requests[0]
                self.assertTrue(r["path"].endswith("/chat/completions"))
                self.assertEqual(r["headers"]["authorization"], "Bearer " + SECRET)
                self.assertEqual([m["role"] for m in r["body"]["messages"]], ["system", "user", "assistant", "user"])
                self.assertEqual(r["body"]["model"], "test-model")
            finally:
                server.stop()

    def test_클로드(self):
        try:
            import anthropic  # noqa
        except ImportError:
            self.skipTest("anthropic 라이브러리 없음")
        server = ai_mock_server.start(lambda q: "답:" + q)
        try:
            self.assertEqual(self.ask("claude", server, "claude-opus-5-5"), "답:질문")
            r = server.requests[-1]
            self.assertTrue(r["path"].split("?")[0].endswith("/v1/messages"))
            sent = r["headers"].get("x-api-key") or r["headers"].get("authorization", "")     # 라이브러리 판에 따라 머리말 이름이 다르다
            self.assertIn(SECRET, sent)
            self.assertEqual(r["body"]["system"], "안내문")
            self.assertEqual(r["body"]["model"], "claude-opus-5-5")
            self.assertEqual([m["role"] for m in r["body"]["messages"]], ["user", "assistant", "user"])
        finally:
            server.stop()

    def test_오류는_우리말로_알려_주고_키를_드러내지_않는다(self):
        for status, word in ((401, "API 키"), (429, "한도"), (500, "서버")):
            for provider in ("gemini", "openai"):
                server = ai_mock_server.start(lambda q: "", status=status)
                try:
                    with self.assertRaises(ai_client.AIError) as cm:
                        self.ask(provider, server)
                    self.assertIn(word, str(cm.exception))
                    self.assertNotIn(SECRET, str(cm.exception))
                finally:
                    server.stop()
        with self.assertRaises(ai_client.AIError) as cm:              # 서버가 없을 때
            ai_client.ask("openai", "m", SECRET, "s", [{"role": "user", "content": "q"}], base_url="http://127.0.0.1:9", timeout=3)
        self.assertIn("연결", str(cm.exception))
        with self.assertRaises(ai_client.AIError):                     # 키가 비었을 때는 보내지도 않는다
            ai_client.ask("gemini", "m", "", "s", [{"role": "user", "content": "q"}])
        with self.assertRaises(ai_client.AIError) as cm:              # 키에 한글이 섞였을 때
            ai_client.ask("openai", "m", "sk-한글키", "s", [{"role": "user", "content": "q"}])
        self.assertIn("한글", str(cm.exception))
        for provider, url in (("gemini", None), ("custom", "http://한글주소.없음/v1")):       # 모델 이름 · 주소에 한글
            with self.assertRaises(ai_client.AIError):
                ai_client.ask(provider, "모델" if url is None else "m", "k", "s", [{"role": "user", "content": "q"}], base_url=url, timeout=3)

    def test_환경_변수에서_키_찾기(self):
        import os
        old = os.environ.get("OPENAI_API_KEY")
        os.environ["OPENAI_API_KEY"] = "abc"
        try:
            self.assertEqual(ai_client.env_key("openai"), "abc")
        finally:
            if old is None:
                del os.environ["OPENAI_API_KEY"]
            else:
                os.environ["OPENAI_API_KEY"] = old
        self.assertEqual(set(ai_client.PROVIDERS), {"gemini", "openai", "claude", "custom"})


if __name__ == "__main__":
    unittest.main(verbosity=2)
