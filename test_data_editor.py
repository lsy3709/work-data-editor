# -*- coding: utf-8 -*-
"""data_editor.py 의 핵심 기능 테스트 (화면 없이 실행).
실행: python test_data_editor.py
"""
import datetime as dt
import shutil
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
import data_editor as ed  # noqa: E402

SAMPLE = HERE / "samples_original"


def has_access():
    try:
        import pyodbc
        return any("Access" in d and "accdb" in d for d in pyodbc.drivers())
    except Exception:
        return False


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="편집기_"))
        for f in SAMPLE.iterdir():
            shutil.copy2(f, self.tmp / f.name)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def open(self, name, table=None):
        src = ed.open_source(self.tmp / name)
        table = table or src.tables()[0]
        return src, table, src.load(table)


class 값변환(unittest.TestCase):
    def test_종류_추정(self):
        self.assertEqual(ed.kind_of([1, 2, None]), "int")
        self.assertEqual(ed.kind_of([1, 2.5]), "float")
        self.assertEqual(ed.kind_of([dt.datetime(2026, 10, 1)]), "date")
        self.assertEqual(ed.kind_of(["a", 1]), "str")
        self.assertEqual(ed.kind_of([None, None]), "str")

    def test_글자를_값으로(self):
        c = ed.Column("금액", "int")
        self.assertEqual(ed.parse_text("12,300", c), 12300)
        self.assertIsNone(ed.parse_text("  ", c))
        with self.assertRaises(ValueError):
            ed.parse_text("만원", c)
        d = ed.Column("신청일", "date", as_datetime=True)
        for t in ("2026-10-05", "2026.10.05", "2026/10/5", "20261005"):
            self.assertEqual(ed.parse_text(t, d), dt.datetime(2026, 10, 5))
        with self.assertRaises(ValueError):
            ed.parse_text("10월 5일쯤", d)

    def test_값을_글자로(self):
        self.assertEqual(ed.to_text(None), "")
        self.assertEqual(ed.to_text(dt.datetime(2026, 10, 5)), "2026-10-05")
        self.assertEqual(ed.to_text(12300.0), "12300")
        self.assertEqual(ed.to_text(1.5), "1.5")


class 엑셀(Base):
    def test_읽기(self):
        src, table, s = self.open("보금자리론_신청_2026_10.xlsx")
        self.assertEqual(table, "신청내역")
        self.assertEqual(len(s.rows), 240)
        self.assertEqual(s.names[:3], ["신청번호", "신청일", "지점코드"])
        kinds = {c.name: c.kind for c in s.columns}
        self.assertEqual(kinds["신청금액(만원)"], "int")
        self.assertEqual(kinds["신청일"], "date")
        self.assertEqual(s.change_count(), 0)

    def test_수정_추가_삭제_되돌리기(self):
        _, _, s = self.open("보금자리론_신청_2026_10.xlsx")
        rid = s.ids()[0]
        s.set_cell(rid, "처리상태", "승인완료")
        s.set_cell(rid, "신청금액(만원)", "9,900")
        self.assertEqual(s.get(rid, "신청금액(만원)"), 9900)
        with self.assertRaises(ValueError):
            s.set_cell(rid, "신청금액(만원)", "많이")
        new = s.add_row({"신청번호": "HF202610-99999", "신청일": "2026-10-31", "신청금액(만원)": "5000"})
        gone = s.ids()[1:3]
        s.delete_rows(gone)
        self.assertEqual(len(s.rows), 239)
        d = s.diff()
        self.assertEqual((len(d["added"]), len(d["deleted"])), (1, 2))
        self.assertEqual(s.status(new), "added")
        self.assertEqual(s.status(rid), "changed")
        s.undo()                                  # 삭제 취소
        self.assertEqual(len(s.rows), 241)
        self.assertEqual(s.ids()[1:3], gone)      # 원래 자리로
        s.undo(); s.undo(); s.undo()              # 추가 · 수정 2건 취소
        self.assertEqual(s.change_count(), 0)
        self.assertFalse(s.undo())

    def test_찾기와_일괄바꾸기(self):
        _, _, s = self.open("보금자리론_신청_2026_10.xlsx")
        hold = s.search("보류", "처리상태")
        self.assertEqual(len(hold) + len(s.search("반려", "처리상태")), 72)
        self.assertEqual(len(s.search("x99")), 13)             # 전체 열, 대소문자 무시
        n = s.replace("처리상태", "보류", "심사중", whole=True)
        self.assertEqual(n, len(hold))
        self.assertEqual(len(s.search("보류", "처리상태")), 0)
        s.undo()                                               # 한 번에 되돌림
        self.assertEqual(len(s.search("보류", "처리상태")), len(hold))
        n = s.replace("상품명", "보금자리론", "보금자리", whole=False)   # 일부 글자 바꾸기
        self.assertGreater(n, 0)
        self.assertEqual(len(s.search("아낌e보금자리론", "상품명")), 0)
        before = s.change_count()
        with self.assertRaises(ValueError):                    # 숫자 열에 글자 → 아무것도 안 바뀜
            s.replace("대출기간(년)", "30", "삼십", whole=True)
        self.assertEqual(s.change_count(), before)
        n = s.fill("고객구분", "확인필요", s.search("X99", "지점코드"))
        self.assertEqual(n, 13)

    def test_저장하면_백업_서식_다른시트_유지(self):
        src, table, s = self.open("S1_보금자리론_취합.xlsx", "월별요약")
        self.assertEqual(len(s.rows), 3)
        s.set_cell(s.ids()[0], "승인건수", "91")
        s.delete_rows([s.ids()[2]])
        backup = src.save(table, s)
        self.assertTrue(backup.exists())
        self.assertEqual(backup.parent.name, "백업")
        s2 = src.load("월별요약")
        self.assertEqual(len(s2.rows), 2)
        self.assertEqual(s2.get(s2.ids()[0], "승인건수"), 91)
        self.assertEqual(len(src.load("전체내역").rows), 630)   # 다른 시트는 그대로
        self.assertEqual(len(ed.open_source(backup).load("월별요약").rows), 3)

    def test_저장하면_머리글_서식과_추가행(self):
        from openpyxl import load_workbook
        src, table, s = self.open("2026_10_부산지사.xlsx")
        n0 = len(s.rows)
        s.add_row({"신청번호": "HF-NEW", "신청금액(만원)": "12345", "신청일": "2026-10-30"})
        src.save(table, s)
        ws = load_workbook(self.tmp / "2026_10_부산지사.xlsx").active
        self.assertTrue(ws.cell(1, 1).font.bold)
        self.assertEqual(ws.max_row, n0 + 2)
        col = [c.value for c in ws[1]].index("신청금액(만원)") + 1
        self.assertEqual(ws.cell(n0 + 2, col).value, 12345)
        self.assertEqual(ws.cell(n0 + 2, col).number_format, "#,##0")   # 윗줄 서식 이어받음


class 메모리(Base):
    def test_원본은_고친_행만_따로_보관(self):
        _, _, s = self.open("보금자리론_신청_2026_10.xlsx")
        self.assertEqual(s.kept_originals(), 0)            # 열기만 하면 사본을 만들지 않는다
        ids = s.ids()
        before = s.get(ids[0], "처리상태")
        s.set_cell(ids[0], "처리상태", "검토완료")
        s.set_cell(ids[0], "신청금액(만원)", "1")
        self.assertEqual(s.kept_originals(), 1)            # 같은 행을 두 번 고쳐도 한 번만
        self.assertEqual(s.original_row(ids[0])[s.col("처리상태")], before)
        s.delete_rows([ids[5]])
        self.assertEqual(s.kept_originals(), 2)
        d = s.diff()
        self.assertEqual((len(d["updated"]), len(d["deleted"]), len(d["added"])), (2, 1, 0))
        self.assertEqual(s.status(ids[0]), "changed")
        self.assertEqual(s.status(ids[1]), "")
        s.undo(); s.undo(); s.undo()
        self.assertEqual(s.change_count(), 0)
        self.assertEqual(s.status(ids[0]), "")

    def test_기준열은_고치기_전_값으로_찾는다(self):
        _, _, s = self.open("보금자리론_신청_2026_10.xlsx")
        s.set_cell(s.ids()[0], "신청번호", s.get(s.ids()[1], "신청번호"))   # 화면에서는 겹치게 고쳤지만
        self.assertEqual(ed.find_key(s), "신청번호")                        # 원본 기준으로는 여전히 고유
        self.assertEqual(len(list(s.original_values("신청번호"))), 240)

    def test_되돌리기_기록은_한도까지만(self):
        _, _, s = self.open("보금자리론_신청_2026_10.xlsx")
        rid = s.ids()[0]
        for i in range(ed.UNDO_LIMIT + 50):
            s.set_cell(rid, "신청금액(만원)", str(i + 1))
        n = 0
        while s.undo():
            n += 1
        self.assertEqual(n, ed.UNDO_LIMIT)


class 백그라운드작업(Base):
    def wait(self, jobs, seconds=30):
        import time
        end = time.time() + seconds
        while jobs.pending and time.time() < end:
            jobs.poll()
            time.sleep(0.005)
        self.assertEqual(jobs.pending, 0, "작업이 제때 끝나지 않음")

    def test_결과는_부른_쪽_스레드에서_받는다(self):
        import threading
        jobs = ed.Jobs()
        seen = {}

        def work():
            seen["work"] = threading.get_ident()
            return 42

        def done(result, error):
            seen["done"] = threading.get_ident()
            seen["result"], seen["error"] = result, error

        jobs.start(work, done)
        self.assertEqual(jobs.pending, 1)
        self.wait(jobs)
        self.assertEqual((seen["result"], seen["error"]), (42, None))
        self.assertNotEqual(seen["work"], threading.get_ident())   # 일은 다른 스레드에서
        self.assertEqual(seen["done"], threading.get_ident())      # 결과 처리는 부른 스레드에서

    def test_오류도_결과처럼_전달된다(self):
        jobs = ed.Jobs()
        got = []
        jobs.start(lambda: ed.open_source(self.tmp / "없는파일.xlsx"), lambda r, e: got.append((r, e)))
        self.wait(jobs)
        self.assertIsNone(got[0][0])
        self.assertIsInstance(got[0][1], ValueError)

    def test_여러_파일을_동시에_읽어도_결과가_같다(self):
        names = ["보금자리론_신청_2026_10.xlsx", "S1_보금자리론_취합.xlsx", "2026_10_부산지사.xlsx"]
        if has_access():
            names.append("업무DB.accdb")
        names = names * 3                                           # 같은 파일도 여러 스레드가 함께 읽는다

        def read(name):
            src = ed.open_source(self.tmp / name)
            return list(src.load(src.tables()[0]).rows.values())

        expected = {name: read(name) for name in set(names)}
        jobs = ed.Jobs()
        got = []
        for name in names:
            jobs.start(lambda n=name: (n, read(n)), lambda r, e: got.append((r, e)))
        self.wait(jobs, 60)
        self.assertEqual(len(got), len(names))
        for result, error in got:
            self.assertIsNone(error)
            self.assertEqual(result[1], expected[result[0]])


@unittest.skipUnless(has_access(), "pyodbc 또는 Access 드라이버 없음")
class 액세스(Base):
    def test_읽기_수정_저장(self):
        src, table, s = self.open("업무DB.accdb")
        self.assertEqual(table, "신청내역")
        self.assertEqual(len(s.rows), 240)
        self.assertEqual(ed.find_key(s), "신청번호")
        ids = s.ids()
        key0, key1 = s.get(ids[0], "신청번호"), s.get(ids[1], "신청번호")
        s.set_cell(ids[0], "처리상태", "검토완료")
        s.set_cell(ids[0], "신청금액(만원)", "7777")
        s.delete_rows([ids[1]])
        s.add_row({"신청번호": "HF202610-99999", "신청일": "2026-10-31", "지점코드": "S01",
                   "상품명": "보금자리론", "신청금액(만원)": "5000", "처리상태": "승인"})
        backup = src.save(table, s, key="신청번호")
        self.assertTrue(backup.exists())
        s2 = src.load(table)
        self.assertEqual(len(s2.rows), 240)
        by = {s2.get(r, "신청번호"): r for r in s2.ids()}
        self.assertEqual(s2.get(by[key0], "처리상태"), "검토완료")
        self.assertEqual(int(s2.get(by[key0], "신청금액(만원)")), 7777)
        self.assertNotIn(key1, by)
        self.assertIn("HF202610-99999", by)
        self.assertEqual(ed.to_text(s2.get(by["HF202610-99999"], "신청일")), "2026-10-31")

    def test_기준열이_중복이면_저장_거부(self):
        src, table, s = self.open("업무DB.accdb")
        s.set_cell(s.ids()[0], "상품명", "X")
        with self.assertRaises(ValueError):
            src.save(table, s, key="처리상태")
        self.assertEqual(len(src.load(table).search("X", "상품명", exact=True)), 0)

    def test_중간에_실패하면_전부_취소(self):
        import pyodbc
        src, table, s = self.open("업무DB.accdb")
        ids = s.ids()
        s.set_cell(ids[0], "처리상태", "바뀜")
        s.set_cell(ids[1], "처리상태", "바뀜")
        # 다른 프로그램이 그 사이에 한 건을 지운 상황을 흉내 낸다
        con = pyodbc.connect(src.conn_str(), autocommit=True)
        con.execute("DELETE FROM [신청내역] WHERE [신청번호]=?", s.get(ids[1], "신청번호"))
        con.close()
        with self.assertRaises(RuntimeError):
            src.save(table, s, key="신청번호")
        self.assertEqual(len(src.load(table).search("바뀜", "처리상태")), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
