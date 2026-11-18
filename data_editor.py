# -*- coding: utf-8 -*-
"""
업무 데이터 편집기 — 엑셀(.xlsx) 시트와 Access(.accdb) 테이블을 열어 찾고, 고치고, 저장하는 프로그램

실행:  python data_editor.py            (또는 파일을 더블클릭)
       python data_editor.py 파일1 파일2 …   (그 파일들을 탭으로 바로 열기)

필요한 것:  pip install openpyxl            (엑셀)
            pip install pyodbc              (Access 를 열 때만. Access 드라이버도 있어야 함)

코드 구성 (위에서 아래로 읽으면 됩니다)
  1. 값 다루기      : 화면의 글자 ↔ 숫자·날짜 바꾸기
  2. EditSession    : 표 한 장을 메모리에 올려 수정·추가·삭제·되돌리기·찾기·일괄 바꾸기
  3. ExcelSource    : 엑셀 파일 읽기 / 저장 (저장 전에 백업)
  4. AccessSource   : Access 파일 읽기 / 저장 (바뀐 곳만 UPDATE·INSERT·DELETE, 실패하면 전부 취소)
  5. Jobs           : 파일 읽기·저장을 작업 스레드에서 실행 (그동안 화면이 멈추지 않음)
  6. TablePane      : 탭 하나 — 파일 하나를 열어 고치는 화면 (tkinter)
  7. App            : 창 — 탭 여러 개를 담는다
회귀 분석(머신러닝)은 analysis.py(계산) · analysis_window.py(화면)에 따로 있고, [회귀 분석] 단추로 엽니다.
모든 저장은 원본 옆 '백업' 폴더에 복사본을 먼저 만든 뒤에 합니다.
"""
import datetime as dt
import gc
import queue
import re
import shutil
import sys
import threading
import time
from collections import deque
from copy import copy
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

TITLE = "업무 데이터 편집기"


# =====================================================================
# 1. 값 다루기
# =====================================================================
@dataclass
class Column:
    name: str
    kind: str = "str"            # int / float / date / bool / str
    as_datetime: bool = False    # 날짜를 datetime 으로 저장할지 (엑셀·Access 는 보통 True)


def kind_of(values):
    """열에 들어 있는 값들을 보고 종류를 정한다. 섞여 있으면 글자(str)."""
    vals = [v for v in values if v is not None and v != ""]
    if not vals:
        return "str"
    if all(isinstance(v, bool) for v in vals):
        return "bool"
    if all(isinstance(v, int) and not isinstance(v, bool) for v in vals):
        return "int"
    if all(isinstance(v, (int, float, Decimal)) and not isinstance(v, bool) for v in vals):
        return "float"
    if all(isinstance(v, dt.date) for v in vals):
        return "date"
    return "str"


def to_text(v):
    """값 → 화면에 보이는 글자."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "예" if v else "아니오"
    if isinstance(v, dt.datetime):
        if (v.hour, v.minute, v.second) == (0, 0, 0):
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, (float, Decimal)):
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    return str(v)


_DATE = re.compile(r"^(\d{4})[-./]?(\d{1,2})[-./]?(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$")


def parse_text(text, column):
    """화면에서 입력한 글자 → 그 열의 종류에 맞는 값. 맞지 않으면 ValueError."""
    t = (text or "").strip()
    if t == "":
        return None
    kind = column.kind
    if kind in ("int", "float"):
        try:
            num = float(t.replace(",", ""))
        except ValueError:
            raise ValueError(f"[{column.name}] 열은 숫자만 넣을 수 있습니다: '{t}'") from None
        if kind == "int":
            if not num.is_integer():
                raise ValueError(f"[{column.name}] 열은 정수만 넣을 수 있습니다: '{t}'")
            return int(num)
        return num
    if kind == "date":
        m = _DATE.match(t)
        try:
            if not m:
                raise ValueError
            y, mo, d, h, mi, s = (int(x) if x else 0 for x in m.groups())
            value = dt.datetime(y, mo, d, h, mi, s)
        except ValueError:
            raise ValueError(f"[{column.name}] 열은 날짜(예: 2026-10-05)만 넣을 수 있습니다: '{t}'") from None
        return value if column.as_datetime else value.date()
    if kind == "bool":
        low = t.lower()
        if low in ("예", "y", "yes", "true", "1", "o"):
            return True
        if low in ("아니오", "n", "no", "false", "0", "x"):
            return False
        raise ValueError(f"[{column.name}] 열은 예/아니오만 넣을 수 있습니다: '{t}'")
    return t


def share_text(row, pool):
    """한 줄을 목록으로 만들되, 같은 글자는 하나만 두고 함께 쓴다 (pool = 지금까지 나온 글자들).
    '승인'·'보금자리론' 처럼 같은 글자가 수만 번 나오는 표에서 메모리가 크게 준다."""
    share = pool.setdefault
    return [share(v, v) if type(v) is str else v for v in row]


def same(a, b):
    """두 값이 같은지 (None 과 빈 글자는 같다고 본다)."""
    if a in (None, "") and b in (None, ""):
        return True
    return a == b


# =====================================================================
# 2. EditSession — 표 한 장을 메모리에서 고친다
# =====================================================================
UNDO_LIMIT = 200                 # 되돌리기는 최근 200번까지만 기억한다 (기록이 끝없이 쌓이지 않게)


class EditSession:
    """메모리를 아끼는 방법: 원본 표를 통째로 한 벌 더 두지 않는다.
    처음 읽은 행은 번호가 0 ~ _n0-1 이고, 그중 '손댄 행'의 처음 모습만 _orig 에 적어 둔다.
    그래서 5만 행을 열어 10행을 고치면 사본은 10행뿐이다."""

    def __init__(self, columns, rows):
        self.columns = list(columns)
        self.names = [c.name for c in self.columns]
        self.rows = {i: (r if isinstance(r, list) else list(r)) for i, r in enumerate(rows)}   # 행번호(rid) → 값 목록
        self._n0 = len(self.rows)                                 # 처음 읽은 행 수 (이 번호 이상은 새로 추가한 행)
        self._orig = {}                                           # 손댄 행의 처음 모습: rid → 값 튜플
        self._next = self._n0
        self._undo = deque(maxlen=UNDO_LIMIT)                     # 되돌리기용 기록

    # ---- 읽기 ----
    def ids(self):
        return list(self.rows)

    def col(self, name):
        return name if isinstance(name, int) else self.names.index(name)

    def get(self, rid, name):
        return self.rows[rid][self.col(name)]

    # ---- 원본 (처음 읽었을 때의 모습) ----
    def _keep(self, rid):
        """이 행을 처음 건드릴 때 한 번만 처음 모습을 적어 둔다."""
        if rid < self._n0 and rid not in self._orig:
            self._orig[rid] = tuple(self.rows[rid])

    def original_row(self, rid):
        return self._orig[rid] if rid in self._orig else tuple(self.rows[rid])

    def original_values(self, name):
        """처음 읽었을 때 그 열에 있던 값들 (지운 행 포함)."""
        ci = self.col(name)
        return (self._orig[rid][ci] if rid in self._orig else self.rows[rid][ci] for rid in range(self._n0))

    def kept_originals(self):
        return len(self._orig)

    # ---- 고치기 ----
    def set_cell(self, rid, name, text):
        ci = self.col(name)
        new = parse_text(text, self.columns[ci])
        old = self.rows[rid][ci]
        if same(old, new):
            return False
        self._keep(rid)
        self.rows[rid][ci] = new
        self._undo.append([("set", rid, ci, old)])
        return True

    def add_row(self, values):
        row = [None] * len(self.columns)
        for name, text in values.items():
            ci = self.col(name)
            row[ci] = parse_text(text, self.columns[ci])
        rid = self._next
        self._next += 1
        self.rows[rid] = row
        self._undo.append([("del", rid)])
        return rid

    def delete_rows(self, rids):
        order = list(self.rows)
        ops = [("ins", rid, order.index(rid), self.rows[rid]) for rid in rids if rid in self.rows]
        for _, rid, _, _ in ops:
            self._keep(rid)
            del self.rows[rid]
        if ops:
            self._undo.append(ops)
        return len(ops)

    def replace(self, name, find, repl, rids=None, whole=True):
        """한 열에서 찾아 바꾸기. whole=True 면 칸 전체가 같을 때만, False 면 일부 글자만 바꾼다.
        하나라도 값이 맞지 않으면 아무것도 바꾸지 않고 ValueError."""
        ci = self.col(name)
        if not whole and find == "":
            raise ValueError("찾을 글자를 입력하세요.")
        plan = []
        for rid in (self.ids() if rids is None else rids):
            old = self.rows[rid][ci]
            text = to_text(old)
            if whole:
                if text != find.strip():
                    continue
                new_text = repl
            else:
                if find not in text:
                    continue
                new_text = text.replace(find, repl)
            new = parse_text(new_text, self.columns[ci])
            if not same(old, new):
                plan.append((rid, old, new))
        return self._apply(ci, plan)

    def fill(self, name, text, rids):
        """고른 행들의 한 열을 모두 같은 값으로 채운다."""
        ci = self.col(name)
        new = parse_text(text, self.columns[ci])
        plan = [(rid, self.rows[rid][ci], new) for rid in rids if not same(self.rows[rid][ci], new)]
        return self._apply(ci, plan)

    def _apply(self, ci, plan):
        for rid, _, new in plan:
            self._keep(rid)
            self.rows[rid][ci] = new
        if plan:
            self._undo.append([("set", rid, ci, old) for rid, old, _ in plan])
        return len(plan)

    def undo(self):
        """마지막 작업 하나를 되돌린다. 되돌릴 것이 없으면 False."""
        if not self._undo:
            return False
        ops = self._undo.pop()
        for op in reversed(ops):
            if op[0] == "set":
                _, rid, ci, old = op
                self.rows[rid][ci] = old
            elif op[0] == "del":
                del self.rows[op[1]]
        inserts = sorted((op for op in ops if op[0] == "ins"), key=lambda op: op[2])
        if inserts:
            items = list(self.rows.items())
            for _, rid, pos, row in inserts:
                items.insert(pos, (rid, row))
            self.rows = dict(items)
        return True

    # ---- 찾기 ----
    def search(self, text, name=None, exact=False):
        t = (text or "").strip().lower()
        if not t:
            return self.ids()
        cols = range(len(self.columns)) if name in (None, "") else [self.col(name)]
        found = []
        for rid, row in self.rows.items():
            for ci in cols:
                cell = to_text(row[ci]).lower()
                if (cell == t) if exact else (t in cell):
                    found.append(rid)
                    break
        return found

    # ---- 무엇이 바뀌었나 ----
    def diff(self):
        """손댄 행(_orig)과 새 행만 살펴본다 — 표가 커도 고친 만큼만 시간이 든다."""
        added = [rid for rid in self.rows if rid >= self._n0]
        deleted = [rid for rid in sorted(self._orig) if rid not in self.rows]
        updated = []
        for rid in sorted(self._orig):
            if rid in self.rows:
                for ci, (old, new) in enumerate(zip(self._orig[rid], self.rows[rid])):
                    if not same(old, new):
                        updated.append((rid, ci, old, new))
        return {"added": added, "deleted": deleted, "updated": updated}

    def change_count(self):
        d = self.diff()
        return len(d["added"]) + len(d["deleted"]) + len(d["updated"])

    def status(self, rid):
        if rid >= self._n0:
            return "added"
        if rid in self._orig and any(not same(o, n) for o, n in zip(self._orig[rid], self.rows[rid])):
            return "changed"
        return ""

    def describe_changes(self):
        """바뀐 내용을 사람이 읽는 문장 목록으로."""
        d = self.diff()
        lines = []
        for rid, ci, old, new in d["updated"]:
            label = to_text(self._orig[rid][0])
            lines.append(f"수정  {label}  [{self.names[ci]}]  {to_text(old) or '(빈칸)'} → {to_text(new) or '(빈칸)'}")
        for rid in d["added"]:
            lines.append("추가  " + " | ".join(to_text(v) for v in self.rows[rid]))
        for rid in d["deleted"]:
            lines.append("삭제  " + " | ".join(to_text(v) for v in self._orig[rid]))
        return lines


def find_key(session):
    """행을 구분할 수 있는 열(값이 비어 있지 않고 서로 다 다른 첫 열)의 이름. 없으면 None."""
    for name in session.names:
        vals = list(session.original_values(name))
        if vals and all(v is not None for v in vals) and len(set(vals)) == len(vals):
            return name
    return None


def make_backup(path):
    """원본 옆 '백업' 폴더에 '이름_날짜_시각' 으로 복사한다."""
    path = Path(path)
    folder = path.parent / "백업"
    folder.mkdir(exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    target = folder / f"{path.stem}_{stamp}{path.suffix}"
    n = 1
    while target.exists():
        n += 1
        target = folder / f"{path.stem}_{stamp}_{n}{path.suffix}"
    shutil.copy2(path, target)
    return target


# =====================================================================
# 3. ExcelSource — 엑셀
# =====================================================================
class ExcelSource:
    label = "엑셀"
    needs_key = False

    def __init__(self, path):
        self.path = Path(path)

    def tables(self):
        from openpyxl import load_workbook
        wb = load_workbook(self.path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names

    def load(self, name):
        """첫 줄을 열 이름으로 보고 시트 전체를 읽는다.
        read_only=True 는 칸마다 서식 객체를 만들지 않고 값만 한 줄씩 흘려 읽어서 메모리를 훨씬 덜 쓴다."""
        from openpyxl import load_workbook
        wb = load_workbook(self.path, read_only=True)
        try:
            ws = wb[name]
            ws.reset_dimensions()                    # 파일에 적힌 '표 크기' 정보가 틀려도 실제 있는 줄을 다 읽는다
            pool = {}
            rows = [share_text(r, pool) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
        while rows and all(v is None for v in rows[-1]):
            rows.pop()
        if not rows:
            raise ValueError(f"'{name}' 시트가 비어 있습니다.")
        header = rows[0]
        ncol = max((i + 1 for i, v in enumerate(header) if v is not None), default=0)
        if ncol == 0:
            raise ValueError(f"'{name}' 시트의 첫 줄에 열 이름이 없습니다.")
        names = [str(h) if h is not None else f"열{i + 1}" for i, h in enumerate(header[:ncol])]
        del rows[0]
        data = [r if len(r) == ncol else (r + [None] * ncol)[:ncol] for r in rows]   # 길이가 다른 줄만 맞춘다
        del rows
        columns = []
        for ci, col_name in enumerate(names):
            vals = [r[ci] for r in data]
            columns.append(Column(col_name, kind_of(vals),
                                  as_datetime=any(isinstance(v, dt.datetime) for v in vals) or not vals))
        return EditSession(columns, data)

    def save(self, name, session, key=None):
        """고친 표를 그 시트에 다시 쓴다. 다른 시트와 머리글 서식은 그대로 둔다. 백업 파일 경로를 돌려준다."""
        from openpyxl import load_workbook
        backup = make_backup(self.path)
        wb = load_workbook(self.path)
        ws = wb[name]
        old_last = ws.max_row
        data = list(session.rows.values())
        for r, row in enumerate(data, start=2):
            for c, value in enumerate(row, start=1):
                cell = ws.cell(r, c)
                cell.value = value
                if r > old_last and r > 2:                      # 새로 생긴 줄은 윗줄 서식을 이어받는다
                    cell._style = copy(ws.cell(r - 1, c)._style)
        new_last = len(data) + 1
        if old_last > new_last:
            ws.delete_rows(new_last + 1, old_last - new_last)
        try:
            wb.save(self.path)
        except PermissionError:
            raise RuntimeError("엑셀에서 이 파일이 열려 있으면 저장할 수 없습니다. 엑셀을 닫고 다시 저장하세요.") from None
        finally:
            wb.close()
        return backup


# =====================================================================
# 4. AccessSource — Access
# =====================================================================
class AccessSource:
    label = "Access"
    needs_key = True             # 어느 행을 고칠지 알려 주는 '기준 열'이 필요하다

    def __init__(self, path):
        self.path = Path(path)

    def conn_str(self):
        return r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" + f"DBQ={self.path};"

    def _connect(self):
        try:
            import pyodbc
        except ImportError:
            raise RuntimeError("Access 파일을 열려면 pyodbc 가 필요합니다.\n터미널에서:  pip install pyodbc") from None
        try:
            return pyodbc.connect(self.conn_str(), autocommit=False)
        except pyodbc.Error as e:
            raise RuntimeError("Access 에 연결하지 못했습니다.\n"
                               "· Access Database Engine(드라이버)이 설치되어 있는지\n"
                               "· 드라이버와 Python 의 비트 수(64/32)가 같은지 확인하세요.\n\n" + str(e)[:300]) from None

    def tables(self):
        con = self._connect()
        try:
            names = [r.table_name for r in con.cursor().tables(tableType="TABLE")]
        finally:
            con.close()
        return [n for n in names if not n.startswith("MSys")]

    def load(self, name):
        con = self._connect()
        try:
            cur = con.cursor()
            cur.execute(f"SELECT * FROM {q(name)}")
            desc = cur.description
            pool = {}
            data = [share_text([float(v) if isinstance(v, Decimal) else v for v in row], pool)
                    for row in cur]                  # 한 줄씩 받아 바로 바꾼다
        finally:
            con.close()
        kinds = {int: "int", float: "float", Decimal: "float", str: "str", bool: "bool",
                 dt.datetime: "date", dt.date: "date"}
        columns = []
        for ci, d in enumerate(desc):
            kind = kinds.get(d[1]) or kind_of([r[ci] for r in data])
            columns.append(Column(d[0], kind, as_datetime=d[1] is not dt.date))
        return EditSession(columns, data)

    def save(self, name, session, key=None):
        """바뀐 곳만 Access 에 반영한다. 하나라도 실패하면 전부 취소(rollback)한다."""
        if key not in session.names:
            raise ValueError("기준 열을 고르세요. (행을 구분하는 열, 예: 신청번호)")
        ki = session.col(key)
        keys = list(session.original_values(ki))
        if any(k is None for k in keys) or len(set(keys)) != len(keys):
            raise ValueError(f"[{key}] 열은 값이 비어 있거나 겹쳐서 기준 열로 쓸 수 없습니다.\n"
                             "값이 행마다 다른 열(예: 신청번호)을 고르세요.")
        d = session.diff()
        backup = make_backup(self.path)
        con = self._connect()
        try:
            cur = con.cursor()
            table = q(name)

            def run(sql, params, what):
                cur.execute(sql, params)
                if cur.rowcount != 1:
                    raise RuntimeError(f"{what}: 대상 행을 찾지 못했습니다. 다른 사람이 먼저 고쳤을 수 있습니다.\n"
                                       "아무것도 저장하지 않았습니다. 파일을 다시 열어 확인하세요.")

            for rid in d["deleted"]:
                k = session.original_row(rid)[ki]
                run(f"DELETE FROM {table} WHERE {q(key)}=?", [k], f"삭제({k})")

            by_row = {}
            for rid, ci, _, new in d["updated"]:
                by_row.setdefault(rid, []).append((ci, new))
            for rid, cells in by_row.items():
                sets, params = [], []
                for ci, new in cells:
                    if new is None:
                        sets.append(f"{q(session.names[ci])}=NULL")
                    else:
                        sets.append(f"{q(session.names[ci])}=?")
                        params.append(new)
                k = session.original_row(rid)[ki]
                run(f"UPDATE {table} SET {', '.join(sets)} WHERE {q(key)}=?", params + [k], f"수정({k})")

            for rid in d["added"]:
                row = session.rows[rid]
                cols = [ci for ci, v in enumerate(row) if v is not None]
                if not cols:
                    continue
                sql = (f"INSERT INTO {table} ({', '.join(q(session.names[ci]) for ci in cols)}) "
                       f"VALUES ({', '.join('?' for _ in cols)})")
                cur.execute(sql, [row[ci] for ci in cols])
            con.commit()
        except RuntimeError:
            con.rollback()
            raise
        except Exception as e:                       # 드라이버가 낸 오류(형식 불일치, 중복 키 등)
            con.rollback()
            raise RuntimeError("Access 저장 중 오류가 나서 전부 취소했습니다.\n\n" + str(e)[:300]) from None
        finally:
            con.close()
        return backup


def q(name):
    """테이블·열 이름을 SQL 에 넣을 때 [이름] 으로 감싼다."""
    if "]" in name:
        raise ValueError(f"이름에 ] 가 들어 있어 사용할 수 없습니다: {name}")
    return f"[{name}]"


def open_source(path):
    path = Path(path)
    if not path.exists():
        raise ValueError(f"파일이 없습니다: {path}")
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return ExcelSource(path)
    if ext in (".accdb", ".mdb"):
        return AccessSource(path)
    raise ValueError("엑셀(.xlsx) 또는 Access(.accdb) 파일만 열 수 있습니다.")


# =====================================================================
# 5. Jobs — 오래 걸리는 일(파일 읽기·저장)을 다른 스레드에서 한다
# =====================================================================
class Jobs:
    """화면이 멈추지 않게 일을 '작업 스레드'에 맡기고, 결과는 부른 쪽(화면) 스레드에서 받는다.

    규칙 두 가지만 지키면 안전하다.
      1) work 함수 안에서는 화면(tkinter 위젯)을 건드리지 않는다 — tkinter 는 만든 스레드에서만 써야 한다.
      2) 결과 처리(done)는 poll() 을 부른 스레드에서 실행된다. 화면은 poll() 을 주기적으로 부르기만 하면 된다.
    스레드끼리 주고받는 통로는 queue.Queue 하나뿐이고, 이 통로는 여러 스레드가 동시에 써도 안전하다."""

    def __init__(self):
        self._results = queue.Queue()
        self.pending = 0                       # 아직 결과를 받지 않은 작업 수 (화면 스레드에서만 바꾼다)

    def start(self, work, done):
        """work() 를 다른 스레드에서 실행한다. 끝나면 poll() 때 done(결과, 오류) 가 불린다."""
        self.pending += 1

        def run():
            try:
                out = (work(), None)
            except Exception as e:             # 오류도 결과처럼 화면 스레드로 넘긴다
                out = (None, e)
            self._results.put((done, out))

        threading.Thread(target=run, daemon=True).start()

    def poll(self):
        """끝난 작업의 결과를 처리한다. 처리한 개수를 돌려준다."""
        n = 0
        while True:
            try:
                done, (result, error) = self._results.get_nowait()
            except queue.Empty:
                return n
            self.pending -= 1
            n += 1
            done(result, error)


# =====================================================================
# 6. TablePane — 탭 하나 (파일 하나를 열어 고치는 화면)
# =====================================================================
import tkinter as tk                                   # noqa: E402
from tkinter import filedialog, messagebox, ttk        # noqa: E402

FONT = ("맑은 고딕", 10)
ALL = "전체 열"


class TablePane:
    MAX_SHOW = 5000                 # 한 번에 화면에 그리는 최대 행 수

    def __init__(self, app, source):
        self.app = app
        self.root = app.root
        self.source = source
        self.table = None
        self.session = None
        self.view = []              # 지금 화면에 보이는 행번호들
        self.sort = (None, False)
        self.editor = None
        self.last_col = 0
        self.busy = ""              # 읽는 중·저장하는 중이면 그 안내 글 (이 동안에는 고칠 수 없다)
        self.saving = False
        self.closed = False
        self.last_backup = None     # 마지막으로 저장할 때 만든 백업 파일
        self.windows = []           # 이 탭에서 연 회귀 분석 창들
        self.frame = ttk.Frame(app.notebook)
        self._build()

    # ---------------- 화면 만들기 ----------------
    def _build(self):
        top = ttk.Frame(self.frame, padding=(10, 8, 10, 2))
        top.pack(fill="x")
        where = f"{self.source.path.parent.name} / {self.source.path.name}"     # 폴더 / 파일 (이름이 같은 파일을 구분)
        self.file_label = ttk.Label(top, text=f"{self.source.label} · {where}", foreground="#111")
        self.file_label.pack(side="left")
        self.key_box = ttk.Combobox(top, state="readonly", width=16, font=FONT)
        self.key_label = ttk.Label(top, text="기준 열")
        self.table_box = ttk.Combobox(top, state="readonly", width=22, font=FONT)
        self.table_box.pack(side="right")
        self.table_box.bind("<<ComboboxSelected>>", self.on_table_change)
        self.table_label = ttk.Label(top, text="테이블" if self.source.needs_key else "시트")
        self.table_label.pack(side="right", padx=(10, 4))

        bar = ttk.Frame(self.frame, padding=(10, 4, 10, 6))
        bar.pack(fill="x")
        ttk.Label(bar, text="찾기").pack(side="left")
        self.find_col = ttk.Combobox(bar, state="readonly", width=12, values=[ALL], font=FONT)
        self.find_col.set(ALL)
        self.find_col.pack(side="left", padx=4)
        self.find_entry = ttk.Entry(bar, width=14, font=FONT)
        self.find_entry.pack(side="left")
        self.find_entry.bind("<Return>", lambda e: self.do_search())
        ttk.Button(bar, text="찾기", command=self.do_search).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text="전체 보기", command=self.show_all).pack(side="left", padx=4)
        ttk.Button(bar, text="저장", command=self.save).pack(side="right")
        ttk.Button(bar, text="내보내기", command=self.export_dialog).pack(side="right", padx=4)
        ttk.Button(bar, text="변경 내역", command=self.show_changes).pack(side="right")
        ttk.Button(bar, text="되돌리기", command=self.do_undo).pack(side="right", padx=4)
        ttk.Button(bar, text="일괄 바꾸기", command=self.replace_dialog).pack(side="right")
        ttk.Button(bar, text="행 삭제", command=self.delete_selected).pack(side="right", padx=4)
        ttk.Button(bar, text="행 추가", command=self.add_dialog).pack(side="right")

        self.status = ttk.Label(self.frame, text="", padding=(10, 4), foreground="#333")
        self.status.pack(fill="x", side="bottom")

        box = ttk.Frame(self.frame, padding=(10, 0, 10, 0))
        box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(box, show="headings", selectmode="extended")
        ys = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side="right", fill="y")
        xs.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        self.tree.tag_configure("changed", background="#FFF3C4")
        self.tree.tag_configure("added", background="#DDF5E0")
        self.tree.bind("<Double-1>", self.begin_edit)
        self.tree.bind("<Button-1>", self.remember_col, add="+")
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.cover = ttk.Label(box, justify="center", foreground="#666", font=(FONT[0], 12))

    # ---------------- 파일 · 표 열기 (작업 스레드에서 읽는다) ----------------
    def set_busy(self, text, saving=False):
        self.busy = text
        self.saving = bool(text) and saving
        if text and self.session is None:           # 처음 여는 중이면 표 자리에 안내를 띄운다
            self.cover.configure(text=text)
            self.cover.place(relx=0.5, rely=0.45, anchor="center")
        else:
            self.cover.place_forget()
        self.update_status()

    def ready(self):
        """고칠 수 있는 상태인가 (표가 열려 있고, 읽기·저장 중이 아님)."""
        if self.session is None:
            return False
        if self.busy:
            self.root.bell()
            return False
        return True

    def load(self, name=None):
        """시트/테이블을 작업 스레드에서 읽어 온다. name 이 없으면 첫 번째 것."""
        source = self.source

        def work():                                  # ← 작업 스레드 (화면을 건드리지 않는다)
            tables = source.tables()
            if not tables:
                raise ValueError("이 파일에는 시트/테이블이 없습니다.")
            target = name or tables[0]
            return tables, target, source.load(target)

        def done(result, error):                     # ← 화면 스레드
            if self.closed:
                return
            self.set_busy("")
            if error:
                self.app.show_error(str(error))
                if self.session is None:             # 처음부터 못 열었으면 빈 탭을 닫는다
                    self.app.close_pane(self, ask=False)
                else:
                    self.table_box.set(self.table)
                return
            self.show(*result)

        self.cancel_edit()
        self.set_busy(f"불러오는 중…  {source.path.name}" + (f" [{name}]" if name else ""))
        self.app.run_job(work, done)

    def show(self, tables, name, session):
        """읽어 온 표를 화면에 올린다."""
        self.session = session
        self.table = name
        self.table_box["values"] = tables
        self.table_box.set(name)
        s = self.session
        self.find_col["values"] = [ALL] + s.names
        self.find_col.set(ALL)
        self.find_entry.delete(0, "end")
        for w in (self.key_box, self.key_label, self.table_box, self.table_label):
            w.pack_forget()
        if self.source.needs_key:                    # Access 일 때만 '기준 열' 을 보여 준다
            self.key_box["values"] = s.names
            self.key_box.set(find_key(s) or "")
            self.key_box.pack(side="right")
            self.key_label.pack(side="right", padx=(14, 4))
        self.table_box.pack(side="right")
        self.table_label.pack(side="right", padx=(10, 4))
        ids = [f"c{i}" for i in range(len(s.columns))]
        self.tree.configure(columns=ids)
        sample = [s.rows[rid] for rid in s.ids()[:200]]
        for ci, col in enumerate(s.columns):
            longest = max([len(col.name) * 2] + [sum(2 if ord(ch) > 255 else 1 for ch in to_text(r[ci])) for r in sample])
            self.tree.heading(ids[ci], text=col.name, command=lambda c=ci: self.sort_by(c))
            self.tree.column(ids[ci], width=min(max(longest * 8 + 24, 70), 280), stretch=False,
                             anchor="e" if col.kind in ("int", "float") else "w")
        self.cover.place_forget()
        self.sort = (None, False)
        self.view = s.ids()
        self.refresh()

    def on_table_change(self, event=None):
        name = self.table_box.get()
        if name == self.table:
            return
        if self.busy or not self.confirm_discard():
            self.table_box.set(self.table or "")
            return
        self.load(name)

    # ---------------- 표 그리기 ----------------
    def refresh(self):
        self.cancel_edit()
        self.tree.delete(*self.tree.get_children())
        s = self.session
        if s is None:
            return
        self.view = [rid for rid in self.view if rid in s.rows]
        for rid in self.view[:self.MAX_SHOW]:
            self.tree.insert("", "end", iid=str(rid), values=[to_text(v) for v in s.rows[rid]],
                             tags=(s.status(rid),))
        self.update_status()

    def refresh_row(self, rid):
        s = self.session
        self.tree.item(str(rid), values=[to_text(v) for v in s.rows[rid]], tags=(s.status(rid),))
        self.update_status()

    def update_status(self):
        s = self.session
        if self.busy:
            text = self.busy
        elif s is None:
            text = ""
        else:
            n = s.change_count()
            text = f"전체 {len(s.rows):,}건 · 표시 {len(self.view):,}건"
            if len(self.view) > self.MAX_SHOW:
                text += f" (화면에는 앞 {self.MAX_SHOW:,}건만)"
            text += f" · 저장 안 한 변경 {n:,}곳" if n else " · 변경 없음"
        self.status["text"] = text
        self.app.update_title(self)

    def is_dirty(self):
        return self.session is not None and self.session.change_count() > 0

    # ---------------- 찾기 · 정렬 ----------------
    def do_search(self, text=None, column=None):
        if self.session is None:
            return
        text = self.find_entry.get() if text is None else text
        column = self.find_col.get() if column is None else column
        self.view = self.session.search(text, None if column == ALL else column)
        self._apply_sort()
        self.refresh()

    def show_all(self):
        if self.session is None:
            return
        self.find_entry.delete(0, "end")
        self.find_col.set(ALL)
        self.view = self.session.ids()
        self._apply_sort()
        self.refresh()

    def sort_by(self, ci):
        old_ci, old_desc = self.sort
        self.sort = (ci, (not old_desc) if old_ci == ci else False)
        self._apply_sort()
        self.refresh()

    def _apply_sort(self):
        ci, desc = self.sort
        if ci is None or self.session is None:
            return
        s = self.session
        text_sort = s.columns[ci].kind == "str"

        def key(rid):
            v = s.rows[rid][ci]
            return (v is None, to_text(v) if text_sort or v is None else v)

        try:
            self.view.sort(key=key, reverse=desc)
        except TypeError:                      # 종류가 섞인 열은 글자로 비교
            self.view.sort(key=lambda rid: to_text(s.rows[rid][ci]), reverse=desc)

    # ---------------- 칸 고치기 ----------------
    def remember_col(self, event):
        col = self.tree.identify_column(event.x)
        if col:
            self.last_col = int(col[1:]) - 1

    def begin_edit(self, event):
        if not self.ready() or self.tree.identify("region", event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not iid or not col:
            return
        ci = int(col[1:]) - 1
        box = self.tree.bbox(iid, col)
        if not box:
            return
        self.cancel_edit()
        x, y, w, h = box
        entry = ttk.Entry(self.tree, font=FONT)
        entry.insert(0, to_text(self.session.rows[int(iid)][ci]))
        entry.select_range(0, "end")
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.bind("<Return>", lambda e: self.commit_edit())
        entry.bind("<Escape>", lambda e: self.cancel_edit())
        entry.bind("<FocusOut>", lambda e: self.commit_edit())
        self.editor = (entry, int(iid), ci)

    def commit_edit(self):
        if not self.editor:
            return
        entry, rid, ci = self.editor
        self.editor = None
        text = entry.get()
        entry.destroy()
        if self.busy or rid not in self.session.rows:
            return
        try:
            if self.session.set_cell(rid, ci, text):
                self.refresh_row(rid)
        except ValueError as e:
            messagebox.showwarning(TITLE, str(e))

    def cancel_edit(self):
        if self.editor:
            entry = self.editor[0]
            self.editor = None
            entry.destroy()

    # ---------------- 행 추가 · 삭제 · 되돌리기 ----------------
    def add_dialog(self):
        if not self.ready():
            return
        s = self.session
        win = self._dialog("행 추가")
        entries = {}
        for i, col in enumerate(s.columns):
            hint = {"int": "숫자", "float": "숫자", "date": "날짜 2026-10-05", "bool": "예/아니오"}.get(col.kind, "")
            ttk.Label(win, text=col.name).grid(row=i, column=0, sticky="e", padx=(14, 6), pady=3)
            e = ttk.Entry(win, width=28, font=FONT)
            e.grid(row=i, column=1, pady=3)
            ttk.Label(win, text=hint, foreground="#888").grid(row=i, column=2, sticky="w", padx=(6, 14))
            entries[col.name] = e

        def ok():
            if not self.ready():
                return
            try:
                rid = self.add_row({name: e.get() for name, e in entries.items()})
            except ValueError as err:
                messagebox.showwarning(TITLE, str(err), parent=win)
                return
            win.destroy()
            self.tree.see(str(rid))
            self.tree.selection_set(str(rid))

        ttk.Button(win, text="추가", command=ok).grid(row=len(s.columns), column=1, sticky="e", pady=12)
        next(iter(entries.values())).focus_set()
        win.bind("<Return>", lambda e: ok())

    def add_row(self, values):
        rid = self.session.add_row(values)
        self.view.append(rid)
        self.refresh()
        return rid

    def delete_selected(self, ask=True):
        if not self.ready():
            return 0
        rids = [int(i) for i in self.tree.selection()]
        if not rids:
            messagebox.showinfo(TITLE, "삭제할 행을 먼저 고르세요. (Ctrl·Shift 로 여러 행)")
            return 0
        if ask and not messagebox.askyesno(TITLE, f"고른 {len(rids)}행을 삭제할까요?\n(저장 전에는 [되돌리기]로 되살릴 수 있습니다)"):
            return 0
        n = self.session.delete_rows(rids)
        self.refresh()
        return n

    def do_undo(self):
        if self.session is None:
            return
        if self.editor:
            self.cancel_edit()
            return
        if not self.ready():
            return
        before = set(self.session.ids())
        if self.session.undo():
            self.view += [rid for rid in self.session.ids() if rid not in before]   # 되살아난 행만 다시 보이게
            order = {rid: i for i, rid in enumerate(self.session.ids())}
            if self.sort[0] is None:
                self.view.sort(key=order.get)
            self.refresh()

    # ---------------- 일괄 바꾸기 ----------------
    def replace_dialog(self):
        if not self.ready():
            return
        s = self.session
        win = self._dialog("일괄 바꾸기")
        pad = {"padx": 12, "pady": 4}
        ttk.Label(win, text="열").grid(row=0, column=0, sticky="e", **pad)
        col = ttk.Combobox(win, state="readonly", values=s.names, width=24, font=FONT)
        col.set(s.names[min(self.last_col, len(s.names) - 1)])
        col.grid(row=0, column=1, sticky="w", **pad)
        mode = tk.StringVar(value="whole")
        ttk.Radiobutton(win, text="칸 전체가 같을 때 바꾸기", variable=mode, value="whole").grid(row=1, column=1, sticky="w", padx=12)
        ttk.Radiobutton(win, text="칸 안의 일부 글자만 바꾸기", variable=mode, value="part").grid(row=2, column=1, sticky="w", padx=12)
        ttk.Radiobutton(win, text="찾지 않고 모두 같은 값으로 채우기", variable=mode, value="fill").grid(row=3, column=1, sticky="w", padx=12)
        ttk.Label(win, text="찾을 값").grid(row=4, column=0, sticky="e", **pad)
        find = ttk.Entry(win, width=26, font=FONT)
        find.grid(row=4, column=1, sticky="w", **pad)
        ttk.Label(win, text="바꿀 값").grid(row=5, column=0, sticky="e", **pad)
        repl = ttk.Entry(win, width=26, font=FONT)
        repl.grid(row=5, column=1, sticky="w", **pad)
        scope = tk.StringVar(value="view")
        nsel = len(self.tree.selection())
        ttk.Label(win, text="범위").grid(row=6, column=0, sticky="e", **pad)
        f = ttk.Frame(win)
        f.grid(row=6, column=1, sticky="w", padx=12)
        ttk.Radiobutton(f, text=f"화면에 보이는 {len(self.view):,}행", variable=scope, value="view").pack(anchor="w")
        ttk.Radiobutton(f, text=f"고른 {nsel:,}행", variable=scope, value="sel").pack(anchor="w")

        def ok():
            if not self.ready():
                return
            try:
                n = self.apply_replace(col.get(), find.get(), repl.get(), mode.get(), scope.get())
            except ValueError as err:
                messagebox.showwarning(TITLE, str(err), parent=win)
                return
            win.destroy()
            messagebox.showinfo(TITLE, f"{n:,}칸을 바꿨습니다." + ("" if n else "\n(조건에 맞는 칸이 없습니다)"))

        ttk.Button(win, text="바꾸기", command=ok).grid(row=7, column=1, sticky="e", padx=12, pady=12)
        find.focus_set()

    def apply_replace(self, column, find, repl, mode="whole", scope="view"):
        rids = [int(i) for i in self.tree.selection()] if scope == "sel" else list(self.view)
        if mode == "fill":
            n = self.session.fill(column, repl, rids)
        else:
            n = self.session.replace(column, find, repl, rids, whole=(mode == "whole"))
        self.refresh()
        return n

    # ---------------- 변경 내역 · 저장 · 내보내기 ----------------
    def show_changes(self):
        if self.session is None:
            return
        lines = self.session.describe_changes()
        win = self._dialog(f"변경 내역 — 저장하면 반영될 {len(lines):,}건")
        text = tk.Text(win, width=100, height=24, font=FONT, wrap="none")
        text.pack(fill="both", expand=True, padx=10, pady=10)
        text.insert("1.0", "\n".join(lines) if lines else "바뀐 내용이 없습니다.")
        text.configure(state="disabled")

    def save(self, ask=True):
        """작업 스레드에서 저장하고 다시 읽는다. 저장을 시작했으면 True.
        끝나면 last_backup 에 백업 파일 경로가 들어간다. 저장하는 동안 이 탭은 고칠 수 없다(다른 탭은 쓸 수 있다)."""
        if self.session is None:
            return False
        self.commit_edit()
        if not self.ready():
            return False
        s = self.session
        d = s.diff()
        if not (d["added"] or d["deleted"] or d["updated"]):
            if ask:
                messagebox.showinfo(TITLE, "바뀐 내용이 없습니다.")
            return False
        summary = f"수정 {len(d['updated']):,}칸 · 추가 {len(d['added']):,}행 · 삭제 {len(d['deleted']):,}행"
        if ask and not messagebox.askyesno(
                TITLE, f"{self.source.path.name} [{self.table}] 에 저장할까요?\n\n{summary}\n\n저장 전에 '백업' 폴더에 복사본을 만듭니다."):
            return False
        source, table = self.source, self.table
        key = self.key_box.get() if source.needs_key else None

        def work():                                  # ← 작업 스레드. 그동안 이 탭은 busy 라서 s 가 바뀌지 않는다
            backup = source.save(table, s, key=key)
            return backup, source.tables(), source.load(table)

        def done(result, error):                     # ← 화면 스레드
            if self.closed:
                return
            self.set_busy("")
            if error:
                self.app.show_error(str(error))
                return
            backup, tables, session = result
            self.last_backup = backup
            self.show(tables, table, session)
            if ask:
                messagebox.showinfo(TITLE, f"저장했습니다. ({summary})\n\n백업: 백업\\{backup.name}")

        self.last_backup = None
        self.set_busy(f"저장하는 중…  {source.path.name} [{table}]", saving=True)
        self.app.run_job(work, done)
        return True

    def export_dialog(self):
        if self.session is None:
            return
        path = filedialog.asksaveasfilename(title="화면에 보이는 행을 새 엑셀로 내보내기", defaultextension=".xlsx",
                                            filetypes=[("엑셀", "*.xlsx")], initialfile=f"{self.table}_내보내기.xlsx")
        if path:
            try:
                n = self.export_view(path)
            except Exception as e:
                messagebox.showerror(TITLE, str(e))
                return
            messagebox.showinfo(TITLE, f"{n:,}행을 내보냈습니다.\n{path}")

    def export_view(self, path):
        from openpyxl import Workbook
        from openpyxl.cell import WriteOnlyCell
        from openpyxl.styles import Font
        wb = Workbook(write_only=True)               # 한 줄씩 바로 파일로 흘려 써서 메모리를 덜 쓴다
        ws = wb.create_sheet(str(self.table)[:31])
        ws.freeze_panes = "A2"
        head = []
        for name in self.session.names:
            cell = WriteOnlyCell(ws, value=name)
            cell.font = Font(bold=True)
            head.append(cell)
        ws.append(head)
        for rid in self.view:
            ws.append(self.session.rows[rid])
        wb.save(path)
        return len(self.view)

    # ---------------- 그 밖 ----------------
    def _dialog(self, title):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.resizable(False, False)
        win.bind("<Escape>", lambda e: win.destroy())
        return win

    def confirm_discard(self):
        if not self.is_dirty():
            return True
        return messagebox.askyesno(TITLE, f"{self.source.path.name} 에 저장하지 않은 변경이 있습니다. 버리고 계속할까요?")

    def open_analysis(self):
        """이 탭의 표로 회귀 분석 창을 연다 (analysis_window.py). 분석 라이브러리가 없으면 설치 방법을 알려 준다."""
        if not self.ready():
            return None
        try:
            import analysis_window
        except ImportError as e:
            messagebox.showinfo(TITLE, "회귀 분석에는 pandas · scikit-learn · matplotlib 이 필요합니다.\n"
                                       "터미널에서:  pip install -r requirements.txt\n\n" + f"({e})")
            return None
        window = analysis_window.AnalysisWindow(self)
        self.windows.append(window)
        return window

    def close(self):
        """탭을 닫을 때 표 데이터와 화면을 놓아 준다."""
        self.closed = True
        for window in list(self.windows):
            window.close()
        self.cancel_edit()
        self.session = None
        self.view = []
        self.frame.destroy()


# =====================================================================
# 7. App — 창 (탭 여러 개를 담는다)
# =====================================================================
class App:
    def __init__(self, root):
        self.root = root
        self.panes = []             # 열려 있는 탭들
        self.jobs = Jobs()
        self._polling = False
        root.title(TITLE)
        root.geometry("1240x720")
        root.minsize(900, 500)
        self._build()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Control-z>", lambda e: self._on("do_undo"))
        root.bind("<Control-s>", lambda e: self._on("save"))
        root.bind("<Control-f>", lambda e: self.current and self.current.find_entry.focus_set())
        root.bind("<Control-o>", lambda e: self.choose_file())
        root.bind("<Control-w>", lambda e: self.close_pane())

    def _build(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=26, font=FONT)
        style.configure("Treeview.Heading", font=(FONT[0], 10, "bold"))
        style.configure("TButton", font=FONT, padding=(8, 3))
        style.configure("TLabel", font=FONT)
        style.configure("TCheckbutton", font=FONT)
        style.configure("TRadiobutton", font=FONT)
        style.configure("TNotebook.Tab", font=FONT, padding=(12, 4))

        top = ttk.Frame(self.root, padding=(10, 8, 10, 4))
        top.pack(fill="x")
        ttk.Button(top, text="파일 열기", command=self.choose_file).pack(side="left")
        ttk.Button(top, text="탭 닫기", command=self.close_pane).pack(side="left", padx=6)
        ttk.Button(top, text="회귀 분석", command=lambda: self._on("open_analysis")).pack(side="left")
        ttk.Label(top, foreground="#777",
                  text="여러 파일을 한 번에 고를 수 있습니다 · Ctrl+Tab 탭 이동 · Ctrl+W 탭 닫기").pack(side="left", padx=6)
        self.job_label = ttk.Label(top, text="", foreground="#B45309")
        self.job_label.pack(side="right")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=(0, 4))
        self.notebook.enable_traversal()             # Ctrl+Tab / Ctrl+Shift+Tab
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.update_title())
        self.notebook.bind("<Button-2>", self._middle_click)
        self.notebook.bind("<Button-3>", self._tab_menu)
        self.hint = ttk.Label(self.root, justify="center", foreground="#666", font=(FONT[0], 12),
                              text="[파일 열기]로 시작하세요 — 파일마다 탭이 하나씩 열립니다\n\n"
                                   "칸을 더블클릭하면 고칠 수 있습니다 · 노란 줄 = 고친 행 · 초록 줄 = 새 행\n"
                                   "저장하면 원본 옆 '백업' 폴더에 복사본을 먼저 남깁니다")
        self.hint.place(relx=0.5, rely=0.48, anchor="center")

    # ---------------- 지금 보고 있는 탭 ----------------
    @property
    def current(self):
        if not self.panes:
            return None
        selected = self.notebook.select()
        for pane in self.panes:
            if str(pane.frame) == selected:
                return pane
        return None

    def __getattr__(self, name):
        """app.session, app.tree, app.save() 처럼 쓰면 '지금 보고 있는 탭'의 것을 돌려준다."""
        if name.startswith("_") or name in ("root", "panes", "notebook", "jobs", "hint", "job_label"):
            raise AttributeError(name)
        pane = self.current
        if pane is None:
            raise AttributeError(f"열려 있는 탭이 없습니다: {name}")
        return getattr(pane, name)

    def _on(self, method):
        pane = self.current
        if pane is not None:
            getattr(pane, method)()

    # ---------------- 파일 열기 · 탭 닫기 ----------------
    def choose_file(self):
        paths = filedialog.askopenfilenames(
            title="엑셀 또는 Access 파일 열기 (여러 개를 함께 고를 수 있습니다)",
            filetypes=[("엑셀 · Access", "*.xlsx *.xlsm *.accdb *.mdb"), ("모든 파일", "*.*")])
        self.open_paths(paths)

    def open_paths(self, paths):
        """여러 파일을 탭으로 연다. 파일마다 작업 스레드가 따로 읽으므로 동시에 열린다."""
        return [pane for pane in (self.open_path(p) for p in paths) if pane is not None]

    def open_path(self, path, ask=True):
        """파일을 새 탭으로 연다(읽기는 작업 스레드). 이미 열려 있으면 그 탭으로 간다. 탭을 돌려준다.
        ask 는 예전 버전과 맞추려고 남긴 것으로, 새 탭은 기존 작업을 버리지 않으므로 쓰지 않는다."""
        try:
            source = open_source(path)
        except Exception as e:
            self.show_error(str(e))
            return None
        same_file = source.path.resolve()
        for pane in self.panes:                      # 같은 파일을 두 탭에서 고치면 저장이 서로 덮어쓴다
            if pane.source.path.resolve() == same_file:
                self.notebook.select(pane.frame)
                return pane
        pane = TablePane(self, source)
        self.panes.append(pane)
        self.notebook.add(pane.frame, text=source.path.name)
        self.notebook.select(pane.frame)
        self.hint.place_forget()
        pane.load()
        return pane

    def close_pane(self, pane=None, ask=True):
        pane = pane or self.current
        if pane is None or pane not in self.panes:
            return False
        if pane.saving:
            messagebox.showinfo(TITLE, "저장하는 중입니다. 저장이 끝난 뒤에 닫으세요.")
            return False
        if ask and not pane.confirm_discard():
            return False
        self.panes.remove(pane)
        self.notebook.forget(pane.frame)
        pane.close()
        gc.collect()                                 # 닫은 탭의 표 데이터를 바로 돌려받는다
        if not self.panes:
            self.hint.place(relx=0.5, rely=0.48, anchor="center")
        self.update_title()
        return True

    def _tab_at(self, event):
        try:
            index = self.notebook.index(f"@{event.x},{event.y}")
        except tk.TclError:
            return None
        frame = self.notebook.tabs()[index]
        return next((p for p in self.panes if str(p.frame) == frame), None)

    def _middle_click(self, event):
        pane = self._tab_at(event)
        if pane:
            self.close_pane(pane)

    def _tab_menu(self, event):
        pane = self._tab_at(event)
        if not pane:
            return
        menu = tk.Menu(self.root, tearoff=0, font=FONT)
        menu.add_command(label="이 탭 닫기", command=lambda: self.close_pane(pane))
        menu.add_command(label="다른 탭 모두 닫기",
                         command=lambda: [self.close_pane(p) for p in list(self.panes) if p is not pane])
        menu.tk_popup(event.x_root, event.y_root)

    # ---------------- 작업 스레드 ----------------
    def run_job(self, work, done):
        """work 는 작업 스레드에서, done 은 화면 스레드에서 실행된다."""
        self.jobs.start(work, done)
        self._show_jobs()
        if not self._polling:
            self._polling = True
            self.root.after(30, self._poll)

    def _poll(self):
        self.jobs.poll()
        self._show_jobs()
        if self.jobs.pending:
            self.root.after(30, self._poll)
        else:
            self._polling = False

    def _show_jobs(self):
        n = self.jobs.pending
        self.job_label["text"] = f"작업 중 {n}개 (창은 계속 쓸 수 있습니다)" if n else ""

    def wait_idle(self, seconds=120):
        """진행 중인 읽기·저장이 모두 끝날 때까지 기다린다 (자동 점검·스크립트용). 다 끝났으면 True."""
        end = time.time() + seconds
        while self.jobs.pending and time.time() < end:
            self.jobs.poll()
            self.root.update()
            time.sleep(0.01)
        self._show_jobs()
        self.root.update()
        return self.jobs.pending == 0

    # ---------------- 그 밖 ----------------
    def show_error(self, text):
        messagebox.showerror(TITLE, text)

    def update_title(self, pane=None):
        """탭 이름(● = 저장 안 한 변경, … = 읽는 중·저장 중)과 창 제목을 맞춘다."""
        for p in ([pane] if pane else self.panes):
            if p in self.panes:
                mark = "… " if p.busy else ("● " if p.is_dirty() else "")
                self.notebook.tab(p.frame, text=mark + p.source.path.name)
        cur = self.current
        if cur is None:
            self.root.title(TITLE)
        else:
            mark = "● " if cur.is_dirty() else ""
            self.root.title(f"{mark}{TITLE} — {cur.source.path.name}" + (f" [{cur.table}]" if cur.table else ""))

    def on_close(self):
        if any(p.saving for p in self.panes):
            messagebox.showinfo(TITLE, "저장하는 중인 탭이 있습니다. 저장이 끝난 뒤에 닫으세요.")
            return
        dirty = [p.source.path.name for p in self.panes if p.is_dirty()]
        if dirty and not messagebox.askyesno(
                TITLE, f"저장하지 않은 변경이 있는 탭 {len(dirty)}개:\n\n" + "\n".join(dirty) + "\n\n버리고 끝낼까요?"):
            return
        self.root.destroy()


def main():
    try:                                            # 고해상도 화면에서 글자가 흐리지 않게 (Windows)
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    app = App(root)
    app.open_paths(sys.argv[1:])                    # 파일을 여러 개 넘기면 탭으로 모두 연다
    root.mainloop()


if __name__ == "__main__":
    main()
