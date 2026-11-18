# -*- coding: utf-8 -*-
"""내 API 키로 AI 채팅이 실제로 되는지 확인한다 (진짜 AI 서버에 묻는다 — 몇 원쯤의 사용료가 들 수 있다).

실행:  python ai_live_check.py              환경 변수에 키가 있는 AI 를 모두 확인
       python ai_live_check.py gemini       그 AI 만 (gemini · openai · claude)
       python ai_live_check.py gemini 모델이름

키는 환경 변수에서 읽는다:  GEMINI_API_KEY · OPENAI_API_KEY · ANTHROPIC_API_KEY
보내는 것은 연습용 가상 표의 열 이름과 아래 질문 세 개뿐이다. 칸의 값은 보내지 않는다.
"""
import shutil
import sys
import tempfile
import time
from pathlib import Path

import ai_client
import ai_commands as ac
import data_editor as ed

HERE = Path(__file__).resolve().parent


def check_provider(provider, model=None):
    key = ai_client.env_key(provider)
    name = ai_client.PROVIDERS[provider]["name"]
    if not key:
        print(f"SKIP {name}: 환경 변수 {ai_client.PROVIDERS[provider]['env'][0]} 가 없음")
        return None
    model = model or ai_client.PROVIDERS[provider]["models"][0]
    work = Path(tempfile.mkdtemp(prefix="ai_live_"))
    shutil.copy2(HERE / "samples_original" / "보금자리론_신청_2026_10.xlsx", work / "표.xlsx")
    source = ed.open_source(work / "표.xlsx")
    table = source.tables()[0]
    s = source.load(table)
    amount, state, branch = s.col("신청금액(만원)"), s.col("처리상태"), s.col("지점코드")
    x99 = sum(1 for row in s.rows.values() if row[branch] == "X99")
    big = sum(1 for row in s.rows.values() if row[amount] >= 30000 and row[state] == "반려")
    ok = True
    history = []

    def ask(question):
        started = time.perf_counter()
        raw = ai_client.ask(provider, model, key, ac.system_prompt(s, table), history + [{"role": "user", "content": question}])
        history.extend([{"role": "user", "content": question}, {"role": "assistant", "content": raw}])
        reply, actions = ac.parse_reply(raw)
        print(f"  Q. {question}  ({time.perf_counter() - started:.1f}초)")
        print(f"     AI: {reply}")
        print(f"     명령: {[a.get('op') for a in actions]}")
        return actions

    def report(cond, text):
        nonlocal ok
        ok = ok and bool(cond)
        print(f"     {'OK  ' if cond else 'FAIL'} {text}")

    print(f"=== {name} · {model}")
    try:
        actions = ask("지점코드가 X99 인 것을 ON1 로 바꿔 줘")
        result = ac.run(s, actions) if actions else None
        report(result and result.changed and not any(row[branch] == "X99" for row in s.rows.values()) and s.change_count() == x99,
               f"X99 {x99}칸이 ON1 로 바뀜")
        actions = ask("신청금액이 3억 원 이상이면서 반려된 건은 처리상태를 재심사로 바꿔 줘")
        result = ac.run(s, actions) if actions else None
        left = sum(1 for row in s.rows.values() if row[amount] >= 30000 and row[state] == "반려")
        report(result and left == 0 and sum(1 for row in s.rows.values() if row[state] == "재심사") == big,
               f"'3억 원'을 30,000만원으로 알아듣고 {big}건을 재심사로 바꿈")
        actions = ask("고객구분별로 신청금액 평균을 알려 줘")
        result = ac.run(s, actions) if actions else None
        report(result and not result.changed and any("별" in line for line in result.lines), "집계 명령을 만들고 표는 바꾸지 않음")
        actions = ask("오늘 점심 메뉴 추천해 줘")
        report(actions == [], "표와 상관없는 말에는 명령을 만들지 않음")

        shutil.copy2(HERE / "samples_original" / "주택연금_가입현황_정제전.xlsx", work / "정제전.xlsx")
        raw_source = ed.open_source(work / "정제전.xlsx")
        s, table = raw_source.load(raw_source.tables()[0]), raw_source.tables()[0]       # 여기부터는 금액 · 날짜가 글자로 적힌 표
        history.clear()
        actions = ask("전처리 해 줘")
        result = ac.run(s, actions) if actions else None
        report(result and result.changed and len(s.rows) == 60 and s.columns[s.col("월지급액")].kind in ("int", "float"),
               "'전처리 해 줘'만으로 되묻지 않고 계획을 세움 (겹친 행 삭제 → 60행, 월지급액이 숫자)")
    except ai_client.AIError as e:
        print("     FAIL AI 호출 오류:", str(e).replace(key, "***"))
        ok = False
    except ValueError as e:
        print("     FAIL 명령 실행 오류:", e)
        ok = False
    finally:
        shutil.rmtree(work, ignore_errors=True)
    print(f"{'PASS' if ok else 'FAIL'} {name}")
    return ok


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    wanted = [sys.argv[1]] if len(sys.argv) > 1 else ["gemini", "openai", "claude"]
    results = [check_provider(p, sys.argv[2] if len(sys.argv) > 2 else None) for p in wanted]
    done = [r for r in results if r is not None]
    print("RESULT", "확인한 AI 없음 (키가 없음)" if not done else ("PASS" if all(done) else "FAIL"))
    sys.exit(0 if all(done) else 1)
