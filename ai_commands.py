"""
AI 채팅의 가운데 부분 — AI 에게 무엇을 보낼지, 받은 답을 어떻게 표에 적용할지 (화면 없음).

생각은 이렇다.
  1) AI 에게는 표의 '모양'(열 이름 · 종류 · 빈칸 수)만 보낸다. 칸의 값은 보내지 않는다.
  2) AI 는 "무엇을 할지"를 정해진 명령(JSON)으로 답한다. 표를 직접 만지지 못한다.
  3) 프로그램이 그 명령을 내 PC 에서 실행한다. 실행 전에 미리 보여 주고, 실행 뒤에는 되돌릴 수 있다.

    diagnose(session)                                    # 손볼 곳 세기 (값 없이 개수와 판단만)
    actions, notes = suggest(session)                    # AI 없이 만드는 기본 전처리 명령
    prompt = system_prompt(session, "신청내역")          # AI 에게 주는 안내문 (표의 모양 + 진단)
    reply, actions = parse_reply(AI 의 답)               # 말 + 명령들
    lines = preview(session, actions)                    # 적용하면 어떻게 되는지 (표는 그대로)
    result = run(session, actions)                       # 실제로 적용 (한 번의 [되돌리기]로 돌아감)
"""
import json
import re
import statistics
from dataclasses import dataclass, field

import data_editor as ed

OPS = ["set", "replace", "fill_empty", "drop_empty_rows", "delete_rows", "drop_duplicates", "strip", "convert", "scale",
       "compute", "rename_column", "delete_columns", "add_column", "show", "sort", "stats"]
READ_ONLY = {"show", "sort", "stats"}           # 표를 바꾸지 않는 명령 (보기 · 집계)
IS = {"=": "=", "==": "=", "eq": "=", "equals": "=", "is": "=", "!=": "!=", "<>": "!=", "ne": "!=", "not": "!=",
      ">": ">", ">=": ">=", "<": "<", "<=": "<=", "contains": "contains", "includes": "contains", "starts": "starts",
      "startswith": "starts", "ends": "ends", "endswith": "ends", "empty": "empty", "not_empty": "not_empty", "in": "in"}
HOW = {"count": "건수", "sum": "합계", "mean": "평균", "median": "중앙값", "min": "최솟값", "max": "최댓값"}
FILL = {"median": "중앙값", "mean": "평균", "mode": "가장 흔한 값", "value": "넣어 준 값", "above": "윗칸 값"}


@dataclass
class Result:
    lines: list = field(default_factory=list)   # 한 일을 한 줄씩
    changed: bool = False                       # 표가 바뀌었나
    view: list = None                           # 화면에 보여 줄 행번호들 (show · sort 를 했을 때만)


# =====================================================================
# 1. AI 에게 보내는 글
# =====================================================================
LOOKS = {"number": "값이 모두 숫자로 읽힘 → 숫자로 바꿀 수 있음", "date": "값이 모두 날짜로 읽힘 → 날짜로 바꿀 수 있음",
         "code": "숫자처럼 보이지만 번호 · 코드로 보임(앞자리 0 또는 열 이름) → 글자로 둘 것"}
_DATE_SHAPE = re.compile(r"^\d{4}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{1,2}")
_CODE_NAME = ("번호", "코드", "전화", "우편", "id", "no")
_DATE_NAME = ("일자", "날짜", "date", "일시")


def diagnose(session):
    """표를 내 PC 에서 훑어 '손볼 곳'을 센다. 칸의 값은 담지 않고 개수와 판단만 돌려준다.
    {"columns": {열: {"kind", "empty", "spaces", "looks", "unreadable"}}, "duplicates": 겹친 행 수, "spaces": 앞뒤 빈칸 칸 수}
    looks: "number"(글자인데 모두 숫자로 읽힘) · "date" · "code"(숫자 모양이지만 번호 · 코드) · None"""
    s = session
    report = {"columns": {}, "spaces": 0}
    for ci, col in enumerate(s.columns):
        filled = [row[ci] for row in s.rows.values() if not ed.is_empty(row[ci])]
        info = {"kind": col.kind, "empty": len(s.rows) - len(filled), "spaces": 0, "looks": None, "unreadable": 0}
        if col.kind == "str" and filled:
            info["spaces"] = sum(1 for v in filled if isinstance(v, str) and v != v.strip())
            texts = [str(v).strip() for v in filled]
            numbers = sum(1 for x in texts if ed.as_number(x) not in (None, ed.FAIL))
            dates = sum(1 for x in texts if ed.as_date(x) not in (None, ed.FAIL))
            low = col.name.lower()
            date_named = col.name.endswith("일") or any(w in low for w in _DATE_NAME)
            if dates == len(texts) and (date_named or any(_DATE_SHAPE.match(x) for x in texts)):
                info["looks"] = "date"
            elif numbers == len(texts):
                zero = any(len(x) > 1 and x[0] == "0" and x.isdigit() for x in texts)
                info["looks"] = "code" if zero or any(w in low for w in _CODE_NAME) else "number"
            elif numbers >= 0.9 * len(texts):
                info["looks"], info["unreadable"] = "number?", len(texts) - numbers
            elif dates >= 0.9 * len(texts) and date_named:
                info["looks"], info["unreadable"] = "date?", len(texts) - dates
        report["spaces"] += info["spaces"]
        report["columns"][col.name] = info
    seen = set()
    for row in s.rows.values():                      # 앞뒤 빈칸을 지웠다고 치고 겹친 행을 센다
        seen.add(tuple("" if ed.is_empty(v) else ed.to_text(v).strip() for v in row))
    report["duplicates"] = len(s.rows) - len(seen)
    return report


def suggest(session):
    """AI 없이, 진단만으로 만드는 '기본 전처리' 명령들과 덧붙이는 말. (명령 목록, 말 목록)"""
    d = diagnose(session)
    actions, notes = [], []
    if d["spaces"]:
        actions.append({"op": "strip"})
    if d["duplicates"]:
        actions.append({"op": "drop_duplicates"})
    numeric = []
    for name, info in d["columns"].items():
        if info["looks"] in ("number", "date"):
            actions.append({"op": "convert", "column": name, "kind": info["looks"]})
        if info["looks"] == "number" or info["kind"] in ("int", "float"):
            numeric.append(name)
        if info["looks"] in ("number?", "date?"):
            word = "숫자" if info["looks"] == "number?" else "날짜"
            notes.append(f"[{name}] 은 대부분 {word}로 읽히지만 {info['unreadable']:,}칸이 읽히지 않아 그대로 두었습니다. 그 칸을 먼저 고치세요.")
        if info["looks"] == "code":
            notes.append(f"[{name}] 은 번호 · 코드로 보여 글자로 두었습니다.")
    for name in numeric:
        if d["columns"][name]["empty"]:
            actions.append({"op": "fill_empty", "column": name, "how": "median"})
    for name, info in d["columns"].items():
        if info["empty"] and name not in numeric:
            notes.append(f"[{name}] 의 빈칸 {info['empty']:,}칸은 채우지 않았습니다(숫자가 아니라 지어낼 값이 없음). "
                         "필요하면 '가장 흔한 값으로 채워 줘' 또는 '빈칸 있는 행 지워 줘'라고 하세요.")
    return actions, notes


def describe(session, table, samples=0):
    """표의 모양을 글로 적는다. samples 가 0 이면 칸의 값은 한 글자도 넣지 않는다 (진단은 개수와 판단뿐이다)."""
    d = diagnose(session)
    lines = [f"표 이름: {table}", f"행 수: {len(session.rows):,}",
             f"진단: 겹친 행 {d['duplicates']:,}개(앞뒤 빈칸을 지운 뒤 기준) · 앞뒤 빈칸이 있는 칸 {d['spaces']:,}개",
             "열 (이름 | 종류 | 빈칸 수 | 진단):"]
    for name in session.names:
        info = session.column_summary(name)
        look = d["columns"][name]
        line = f"- {name} | {ed.KIND_NAME[info['kind']]} | 빈칸 {info['empty']:,}"
        hints = []
        if look["looks"] in LOOKS:
            hints.append(LOOKS[look["looks"]])
        elif look["looks"]:
            word = "숫자" if look["looks"] == "number?" else "날짜"
            hints.append(f"대부분 {word}로 읽히나 {look['unreadable']:,}칸은 읽히지 않음")
        if look["spaces"]:
            hints.append(f"앞뒤 빈칸 {look['spaces']:,}칸")
        if hints:
            line += " | 진단: " + ", ".join(hints)
        if samples:
            line += " | 값 예시: " + ", ".join(info["examples"][:4])
        lines.append(line)
    if samples:
        lines.append(f"앞의 {samples}행 (예시):")
        for rid in session.ids()[:samples]:
            lines.append("  " + json.dumps({n: ed.to_text(v) for n, v in zip(session.names, session.rows[rid])}, ensure_ascii=False))
    else:
        lines.append("(칸의 값은 보내지 않았습니다. 값이 필요하면 사용자에게 물어보세요.)")
    return "\n".join(lines)


def system_prompt(session, table, samples=0):
    return f"""너는 '업무 데이터 편집기'라는 표 편집 프로그램 안에서 일하는 도우미다.
사용자가 우리말로 표를 어떻게 고치고 싶은지 말하면, 프로그램이 실행할 명령으로 바꿔 준다.
너는 표의 값을 볼 수 없고 직접 고칠 수도 없다. 명령만 만든다. 실행은 프로그램이 사용자 확인을 받은 뒤에 한다.

[지금 열려 있는 표]
{describe(session, table, samples)}

[답하는 형식] 반드시 아래 모양의 JSON 객체 하나만 답한다. 다른 글이나 코드 울타리를 붙이지 않는다.
{{"reply": "사용자에게 하는 말 (한두 문장, 우리말)", "actions": [명령, 명령, ...]}}
- 할 일이 없으면 "actions" 를 빈 배열로 두고 "reply" 로 말한다. 되묻는 것은 정말 정할 수 없을 때만 한다.
- 열 이름은 위 목록에 있는 것을 글자 그대로 쓴다. 없는 열을 지어내지 않는다.

[명령] "op" 로 종류를 정한다.
- {{"op": "set", "column": 열, "value": 값, "where": [조건...]}}  조건에 맞는 행의 그 열을 값으로 바꾼다 (값이 null 이면 비운다)
- {{"op": "replace", "column": 열, "find": 찾을글자, "to": 바꿀글자, "mode": "cell" 또는 "part"}}  cell = 칸 전체가 같을 때, part = 일부 글자만
- {{"op": "fill_empty", "column": 열, "how": "median"|"mean"|"mode"|"value"|"above", "value": 값}}  빈칸 채우기 (value 는 how 가 "value" 일 때만)
- {{"op": "drop_empty_rows", "columns": [열...]}}  그 열에 빈칸이 있는 행 삭제 (columns 를 빼면 모든 열)
- {{"op": "delete_rows", "where": [조건...]}}  조건에 맞는 행 삭제 (조건이 꼭 있어야 한다)
- {{"op": "drop_duplicates", "columns": [열...]}}  겹친 행 삭제 (columns 를 빼면 모든 값이 같은 행)
- {{"op": "strip", "columns": [열...]}}  글자 앞뒤 빈칸 지우기 (columns 를 빼면 모든 열)
- {{"op": "convert", "column": 열, "kind": "number"|"date"|"str"|"bool", "on_error": "stop"|"empty"}}  열의 종류 바꾸기
- {{"op": "scale", "column": 열, "by": 숫자, "divide": true|false}}  숫자 열에 곱하기(나누려면 divide: true)
- {{"op": "compute", "target": 결과열, "left": 열, "operator": "+"|"-"|"*"|"/", "right": 열 또는 숫자}}  계산해서 결과열에 넣기 (결과열이 없으면 새로 만든다)
- {{"op": "rename_column", "column": 열, "to": 새이름}}
- {{"op": "delete_columns", "columns": [열...]}}
- {{"op": "add_column", "name": 새열, "kind": "str"|"int"|"float"|"date", "value": 값, "after": 열}}
- {{"op": "show", "where": [조건...]}}  조건에 맞는 행만 화면에 보이기 (표는 바꾸지 않는다)
- {{"op": "sort", "column": 열, "descending": true|false}}  화면의 행 순서 바꾸기 (표는 바꾸지 않는다)
- {{"op": "stats", "column": 열, "how": "count"|"sum"|"mean"|"median"|"min"|"max", "by": 묶을열, "where": [조건...]}}  집계. 프로그램이 계산해 사용자에게 보여 준다 (너에게는 결과가 가지 않는다)

[조건] {{"column": 열, "is": 비교, "value": 값}} 의 배열. 모두 맞아야 한다. 하나만 맞아도 되게 하려면 명령에 "match": "any" 를 넣는다.
비교: "=", "!=", ">", ">=", "<", "<=", "contains"(포함), "starts"(시작), "ends"(끝), "empty"(빈칸), "not_empty", "in"(value 가 배열)

[막연한 요청] "전처리 해 줘", "정제해 줘", "정리해 줘", "널값(결측치) 처리해 줘", "숫자로 바꿔 줘" 처럼 무엇을 · 어느 열에 할지 다 말하지 않아도 되묻지 말고, 위 [지금 열려 있는 표]의 진단을 보고 바로 명령을 만든다. 사용자는 실행 전에 미리 보기를 확인하고 되돌릴 수도 있으니 합리적인 기본값으로 정하면 된다.
- 기본 전처리 순서: ① 앞뒤 빈칸이 있으면 strip ② 겹친 행이 있으면 drop_duplicates ③ 진단이 "숫자로 바꿀 수 있음"인 열은 convert number, "날짜로 바꿀 수 있음"인 열은 convert date ④ 숫자 열(바꾼 열 포함)의 빈칸은 fill_empty median
- "번호 · 코드로 보임"인 열은 숫자로 바꾸지 않는다. "n칸은 읽히지 않음"인 열은 바꾸지 말고 reply 에서 알려 준다.
- 글자 열의 빈칸은 지어내지 않는다. 사용자가 "널값 제거"처럼 지우라고 했으면 drop_empty_rows, 채우라고만 했으면 숫자 열은 median · 글자 열은 mode 를 쓴다.
- 열을 말하지 않았으면 진단에 해당하는 열 모두에 적용한다. 무엇을 기본값으로 골랐는지 reply 에 한두 문장으로 적는다.
- 진단에 손볼 곳이 하나도 없을 때만 "actions" 를 비우고 "손볼 곳이 없습니다"라고 답한다.

[주의]
- "평균이 얼마야?" 처럼 값을 묻는 말에는 스스로 답을 지어내지 말고 stats 명령을 만든다.
- 지우는 일(delete_rows, delete_columns)은 사용자가 분명히 말했을 때만 한다.
- 표와 상관없는 질문에는 "actions" 를 비우고 짧게 답한다."""


# =====================================================================
# 2. AI 의 답 읽기
# =====================================================================
def parse_reply(text):
    """AI 의 답에서 (사용자에게 하는 말, 명령 목록)을 꺼낸다. JSON 이 아니면 말로만 받는다."""
    text = (text or "").strip()
    start = text.find("{")
    if start >= 0:
        try:
            data, _ = json.JSONDecoder().raw_decode(text[start:])
        except ValueError:
            data = None
        if isinstance(data, dict) and ("actions" in data or "reply" in data):
            actions = data.get("actions") or []
            if isinstance(actions, dict):
                actions = [actions]
            return str(data.get("reply") or "").strip(), [a for a in actions if isinstance(a, dict)]
    return re.sub(r"^```\w*\n?|\n?```$", "", text).strip(), []


def changes_data(actions):
    return any(a.get("op") not in READ_ONLY for a in actions)


# =====================================================================
# 3. 명령 실행
# =====================================================================
def preview(session, actions):
    """적용하면 무슨 일이 생기는지 줄글로 돌려준다. 실제로 해 본 뒤 곧바로 되돌리므로 표는 그대로다."""
    result = run(session, actions)
    if result.changed:
        session.undo()
    return result.lines


def run(session, actions):
    """명령들을 차례로 실행한다. 하나라도 틀리면 앞서 한 것까지 모두 취소하고 ValueError 를 낸다.
    여러 명령이 한 번의 [되돌리기]로 돌아가게 기록을 하나로 묶는다."""
    if not actions:
        raise ValueError("실행할 명령이 없습니다.")
    undo = session._undo
    mark = undo[-1] if undo else None
    result = Result()
    try:
        for n, action in enumerate(actions, 1):
            op = action.get("op")
            if op not in OPS:
                raise ValueError(f"모르는 명령입니다: {op}")
            try:
                result.lines.append(_DO[op](session, action, result))
            except ValueError as e:
                raise ValueError(f"{n}번째 명령({op}): {e}") from None
            except (KeyError, TypeError, AttributeError) as e:
                raise ValueError(f"{n}번째 명령({op})에 필요한 값이 빠졌거나 모양이 틀립니다: {e}") from None
    except ValueError:
        while undo and undo[-1] is not mark:          # 앞서 한 것까지 모두 취소
            session.undo()
        raise
    steps = []
    while undo and undo[-1] is not mark:
        steps.append(undo.pop())
    if steps:
        undo.append([op for step in reversed(steps) for op in step])
        result.changed = True
    if result.view is not None:
        result.view = [rid for rid in result.view if rid in session.rows]
    return result


# ---------------- 도움 함수 ----------------
def _col(s, name):
    if not isinstance(name, str) or name not in s.names:
        raise ValueError(f"[{name}] 이라는 열이 없습니다. 있는 열: {', '.join(s.names)}")
    return s.col(name)


def _cols(s, action):
    names = action.get("columns")
    if names is None and action.get("column") is not None:
        names = [action["column"]]
    if isinstance(names, str):
        names = [names]
    for name in names or []:
        _col(s, name)
    return names or None


def _test(s, cond):
    """조건 하나를 '행 → 참/거짓' 함수로 만든다."""
    ci = _col(s, cond.get("column"))
    col = s.columns[ci]
    how = IS.get(str(cond.get("is", "=")).strip().lower())
    if how is None:
        raise ValueError(f"모르는 비교입니다: {cond.get('is')}")
    value = cond.get("value")
    if how == "empty":
        return lambda row: ed.is_empty(row[ci])
    if how == "not_empty":
        return lambda row: not ed.is_empty(row[ci])
    if how == "in":
        wanted = {str(v) for v in (value if isinstance(value, list) else [value])}
        return lambda row: ed.to_text(row[ci]) in wanted
    if how in ("contains", "starts", "ends"):
        word = str(value)
        check = {"contains": lambda t: word in t, "starts": lambda t: t.startswith(word), "ends": lambda t: t.endswith(word)}[how]
        return lambda row: check(ed.to_text(row[ci]))
    if col.kind in ("int", "float"):
        target, read = ed.as_number(value), ed.as_number
    elif col.kind == "date":
        target, read = ed.as_date(value), ed.as_date
    else:
        target, read = ("" if value is None else str(value)), lambda v: ed.to_text(v)
    if target is ed.FAIL or (target is None and how not in ("=", "!=")):
        raise ValueError(f"[{col.name}] 은 {ed.KIND_NAME[col.kind]} 열인데 비교할 값이 맞지 않습니다: {value!r}")
    compare = {"=": lambda a: a == target, "!=": lambda a: a != target, ">": lambda a: a > target,
               ">=": lambda a: a >= target, "<": lambda a: a < target, "<=": lambda a: a <= target}[how]

    def test(row):
        v = read(row[ci])
        if v is None or v is ed.FAIL:
            return how == "!=" and target is not None
        return compare(v)
    return test


def _match(s, action):
    """where 조건에 맞는 행번호들. 조건이 없으면 모든 행."""
    where = action.get("where") or []
    if isinstance(where, dict):
        where = [where]
    tests = [_test(s, cond) for cond in where]
    if not tests:
        return s.ids()
    pick = any if str(action.get("match", "all")).lower() == "any" else all
    return [rid for rid, row in s.rows.items() if pick(t(row) for t in tests)]


def _where_text(action):
    where = action.get("where") or []
    if isinstance(where, dict):
        where = [where]
    if not where:
        return "모든 행"
    joiner = " 또는 " if str(action.get("match", "all")).lower() == "any" else " 그리고 "
    words = {"=": "=", "!=": "≠", "contains": "포함", "starts": "시작", "ends": "끝", "empty": "빈칸", "not_empty": "빈칸 아님", "in": "중 하나"}
    parts = []
    for cond in where:
        how = IS.get(str(cond.get("is", "=")).strip().lower(), "?")
        tail = "" if how in ("empty", "not_empty") else f" {cond.get('value')!r}"
        parts.append(f"[{cond.get('column')}] {words.get(how, how)}{tail}")
    return joiner.join(parts)


def _value(col, value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        value = "예" if value else "아니오"
    return ed.parse_text(str(value), col)


def _num(x):
    if isinstance(x, float) and not x.is_integer():
        return f"{x:,.2f}"
    return f"{int(x):,}"


# ---------------- 명령 하나하나 ----------------
def _set(s, a, result):
    ci = _col(s, a["column"])
    new = _value(s.columns[ci], a.get("value"))
    rids = _match(s, a)
    n = s._change([(rid, ci, s.rows[rid][ci], new) for rid in rids if not ed.same(s.rows[rid][ci], new)])
    return f"[{a['column']}] 을(를) {ed.to_text(new) or '빈칸'} (으)로 — {_where_text(a)} {len(rids):,}행 중 {n:,}칸 바뀜"


def _replace(s, a, result):
    ci = _col(s, a["column"])
    col, find, to = s.columns[ci], str(a["find"]), "" if a.get("to") is None else str(a["to"])
    part = str(a.get("mode", "cell")).lower() == "part"
    plan = []
    for rid in _match(s, a):
        old = s.rows[rid][ci]
        text = ed.to_text(old)
        if (find in text) if part else (text == find):
            plan.append((rid, ci, old, _value(col, text.replace(find, to) if part else to)))
    n = s._change(plan)
    return f"[{a['column']}] 의 '{find}' → '{to}' ({'일부 글자' if part else '칸 전체'}) — {n:,}칸"


def _fill(s, a, result):
    how = str(a.get("how", "median")).lower()
    n = sum(s.fill_empty(name, how, a.get("value")) for name in _cols(s, a))
    return f"{_names(_cols(s, a))} 의 빈칸을 {FILL.get(how, how)}(으)로 채움 — {n:,}칸"


def _drop_empty(s, a, result):
    names = _cols(s, a)
    return f"{_names(names) if names else '어느 열이든'} 빈칸이 있는 행 삭제 — {s.drop_empty_rows(names):,}행"


def _delete_rows(s, a, result):
    if not a.get("where"):
        raise ValueError("조건 없이 모든 행을 지우는 것은 받지 않습니다. 어떤 행을 지울지 조건을 말해 주세요.")
    n = s.delete_rows(_match(s, a))
    return f"{_where_text(a)} 인 행 삭제 — {n:,}행"


def _drop_dup(s, a, result):
    names = _cols(s, a)
    return f"{_names(names) + ' 이(가) 같은' if names else '모든 값이 같은'} 겹친 행 삭제 — {s.drop_duplicates(names):,}행"


def _strip(s, a, result):
    names = _cols(s, a)
    return f"{_names(names) if names else '모든 열'} 의 앞뒤 빈칸 지우기 — {s.strip_text(names):,}칸"


def _convert(s, a, result):
    kind = str(a["kind"]).lower()
    kind = {"int": "number", "float": "number", "text": "str", "string": "str"}.get(kind, kind)
    on_error = "empty" if str(a.get("on_error", "stop")).lower() == "empty" else "stop"
    changed = failed = 0
    for name in _cols(s, a):
        c, f = s.convert(name, kind, on_error=on_error)
        changed, failed = changed + c, failed + f
    return f"{_names(_cols(s, a))} 을(를) {ed.KIND_NAME.get(kind, kind)}(으)로 — 값이 바뀐 칸 {changed:,}" + (f", 못 바꿔 비운 칸 {failed:,}" if failed else "")


def _scale(s, a, result):
    by = ed.as_number(a.get("by"))
    if by in (None, ed.FAIL, 0):
        raise ValueError(f"곱하거나 나눌 수가 맞지 않습니다: {a.get('by')!r}")
    divide = bool(a.get("divide"))
    n = sum(s.scale(name, by, divide=divide) for name in _cols(s, a))
    return f"{_names(_cols(s, a))} 을(를) {_num(by)} (으)로 {'나누기' if divide else '곱하기'} — {n:,}칸"


def _compute(s, a, result):
    li = _col(s, a["left"])
    if s.columns[li].kind not in ("int", "float"):
        raise ValueError(f"[{a['left']}] 은 숫자 열이 아닙니다. 먼저 숫자로 바꾸세요.")
    sign = str(a.get("operator", "")).strip()
    work = {"+": lambda x, y: x + y, "-": lambda x, y: x - y, "*": lambda x, y: x * y, "×": lambda x, y: x * y,
            "/": lambda x, y: x / y if y else None, "÷": lambda x, y: x / y if y else None}.get(sign)
    if work is None:
        raise ValueError(f"모르는 계산입니다: {sign!r} (+ - * / 중 하나)")
    right = a.get("right")
    if isinstance(right, str) and right in s.names:
        ri = s.col(right)
        if s.columns[ri].kind not in ("int", "float"):
            raise ValueError(f"[{right}] 은 숫자 열이 아닙니다.")
        get = lambda row: row[ri]                    # noqa: E731
    else:
        number = ed.as_number(right)
        if number in (None, ed.FAIL):
            raise ValueError(f"오른쪽 값이 열 이름도 숫자도 아닙니다: {right!r}")
        get = lambda row: number                     # noqa: E731
    target = str(a["target"]).strip()
    if target not in s.names:
        s.add_column(target, kind="float")
    ti = s.col(target)
    li = s.col(a["left"])                            # 열을 끼워 넣었으면 자리가 밀렸을 수 있다
    if isinstance(right, str) and right in s.names:
        ri = s.col(right)
    before = s._retype(ti, "float") if s.columns[ti].kind != "float" else []
    plan = []
    for rid, row in s.rows.items():
        x, y = row[li], get(row)
        new = None if ed.is_empty(x) or ed.is_empty(y) else work(x, y)
        if new is not None:
            new = round(new, 10)
        if not ed.same(row[ti], new) or type(row[ti]) is not type(new):
            plan.append((rid, ti, row[ti], new))
    n = s._change(plan, before)
    return f"[{target}] = [{a['left']}] {sign} {('[' + right + ']') if isinstance(right, str) and right in s.names else right} — {n:,}칸"


def _rename(s, a, result):
    _col(s, a["column"])
    s.rename_column(a["column"], str(a["to"]))
    return f"열 이름 [{a['column']}] → [{str(a['to']).strip()}]"


def _delete_cols(s, a, result):
    names = _cols(s, a)
    if not names:
        raise ValueError("지울 열을 알려 주세요.")
    return f"{_names(names)} 열 삭제 — {s.delete_columns(names)}개"


def _add_col(s, a, result):
    kind = str(a.get("kind", "str")).lower()
    kind = {"number": "float", "text": "str", "string": "str"}.get(kind, kind)
    if kind not in ("str", "int", "float", "date", "bool"):
        raise ValueError(f"모르는 종류입니다: {kind}")
    after = a.get("after")
    if after is not None:
        _col(s, after)
    s.add_column(str(a["name"]), kind=kind, value=a.get("value"), after=after)
    return f"[{str(a['name']).strip()}] 열 추가 ({ed.KIND_NAME[kind]})" + (f", 모든 행에 {a['value']!r}" if a.get("value") not in (None, "") else "")


def _show(s, a, result):
    result.view = _match(s, a)
    return f"{_where_text(a)} 만 보기 — {len(result.view):,}행"


def _sort(s, a, result):
    ci = _col(s, a["column"])
    down = bool(a.get("descending"))
    rids = result.view if result.view is not None else s.ids()
    filled = [rid for rid in rids if not ed.is_empty(s.rows[rid][ci])]
    empty = [rid for rid in rids if ed.is_empty(s.rows[rid][ci])]
    try:
        filled.sort(key=lambda rid: s.rows[rid][ci], reverse=down)
    except TypeError:                                # 종류가 섞인 열은 글자로 견준다
        filled.sort(key=lambda rid: ed.to_text(s.rows[rid][ci]), reverse=down)
    result.view = filled + empty
    return f"[{a['column']}] {'큰 것부터' if down else '작은 것부터'} 정렬 (빈칸은 맨 뒤)"


def _stats(s, a, result):
    how = str(a.get("how", "count")).lower()
    if how not in HOW:
        raise ValueError(f"모르는 집계입니다: {how}")
    rids = _match(s, a)
    ci = _col(s, a["column"]) if a.get("column") else None
    if how != "count":
        if ci is None:
            raise ValueError("어느 열을 집계할지 알려 주세요.")
        if s.columns[ci].kind not in ("int", "float"):
            raise ValueError(f"[{a['column']}] 은 숫자 열이 아니어서 {HOW[how]}을(를) 낼 수 없습니다. 건수(count)만 됩니다.")
    bi = _col(s, a["by"]) if a.get("by") else None

    def figure(group):
        values = [s.rows[rid][ci] for rid in group if not ed.is_empty(s.rows[rid][ci])] if ci is not None else group
        if how == "count":
            return f"{len(values):,}"
        if not values:
            return "값 없음"
        return _num({"sum": sum, "mean": statistics.fmean, "median": statistics.median, "min": min, "max": max}[how](values))

    head = f"{HOW[how]}" + (f" [{a['column']}]" if ci is not None else "") + f" — {_where_text(a)} {len(rids):,}행"
    if bi is None:
        return f"{head}: {figure(rids)}"
    groups = {}
    for rid in rids:
        groups.setdefault(ed.to_text(s.rows[rid][bi]) or "(빈칸)", []).append(rid)
    rows = [f"    {key}: {figure(group)} ({len(group):,}행)" for key, group in sorted(groups.items())[:40]]
    more = f"\n    … 외 {len(groups) - 40}개 묶음" if len(groups) > 40 else ""
    return f"{head}, [{a['by']}] 별\n" + "\n".join(rows) + more


def _names(names):
    shown = ", ".join(f"[{n}]" for n in names[:3])
    return shown + (f" 외 {len(names) - 3}개" if len(names) > 3 else "")


_DO = {"set": _set, "replace": _replace, "fill_empty": _fill, "drop_empty_rows": _drop_empty, "delete_rows": _delete_rows,
       "drop_duplicates": _drop_dup, "strip": _strip, "convert": _convert, "scale": _scale, "compute": _compute,
       "rename_column": _rename, "delete_columns": _delete_cols, "add_column": _add_col, "show": _show, "sort": _sort,
       "stats": _stats}
