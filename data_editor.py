# -*- coding: utf-8 -*-
"""
업무 데이터 편집기 — 엑셀(.xlsx) 시트와 Access(.accdb) 테이블을 열어 찾고, 고치고, 저장하는 프로그램

실행:  python data_editor.py            (또는 파일을 더블클릭)
       python data_editor.py 파일경로   (그 파일을 바로 열기)

필요한 것:  pip install openpyxl            (엑셀)
            pip install pyodbc              (Access 를 열 때만. Access 드라이버도 있어야 함)

코드 구성 (위에서 아래로 읽으면 됩니다)
  1. 값 다루기      : 화면의 글자 ↔ 숫자·날짜 바꾸기
  2. EditSession    : 표 한 장을 메모리에 올려 수정·추가·삭제·되돌리기·찾기·일괄 바꾸기
  3. ExcelSource    : 엑셀 파일 읽기 / 저장 (저장 전에 백업)
  4. AccessSource   : Access 파일 읽기 / 저장 (바뀐 곳만 UPDATE·INSERT·DELETE, 실패하면 전부 취소)
  5. App            : 화면 (tkinter)
모든 저장은 원본 옆 '백업' 폴더에 복사본을 먼저 만든 뒤에 합니다.
"""
import datetime as dt
import re
import shutil
import sys
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


def same(a, b):
    """두 값이 같은지 (None 과 빈 글자는 같다고 본다)."""
    if a in (None, "") and b in (None, ""):
        return True
    return a == b


# =====================================================================
# 2. EditSession — 표 한 장을 메모리에서 고친다
# =====================================================================
class EditSession:
    def __init__(self, columns, rows):
        self.columns = list(columns)
        self.names = [c.name for c in self.columns]
        self.rows = {i: list(r) for i, r in enumerate(rows)}      # 행번호(rid) → 값 목록
        self.original = {rid: tuple(r) for rid, r in self.rows.items()}
        self._next = len(self.rows)
        self._undo = []                                           # 되돌리기용 기록

    # ---- 읽기 ----
    def ids(self):
        return list(self.rows)

    def col(self, name):
        return name if isinstance(name, int) else self.names.index(name)

    def get(self, rid, name):
        return self.rows[rid][self.col(name)]

    # ---- 고치기 ----
    def set_cell(self, rid, name, text):
        ci = self.col(name)
        new = parse_text(text, self.columns[ci])
        old = self.rows[rid][ci]
        if same(old, new):
            return False
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
        added = [rid for rid in self.rows if rid not in self.original]
        deleted = [rid for rid in self.original if rid not in self.rows]
        updated = []
        for rid, row in self.rows.items():
            if rid in self.original:
                for ci, (old, new) in enumerate(zip(self.original[rid], row)):
                    if not same(old, new):
                        updated.append((rid, ci, old, new))
        return {"added": added, "deleted": deleted, "updated": updated}

    def change_count(self):
        d = self.diff()
        return len(d["added"]) + len(d["deleted"]) + len(d["updated"])

    def status(self, rid):
        if rid not in self.original:
            return "added"
        if any(not same(o, n) for o, n in zip(self.original[rid], self.rows[rid])):
            return "changed"
        return ""

    def describe_changes(self):
        """바뀐 내용을 사람이 읽는 문장 목록으로."""
        d = self.diff()
        lines = []
        for rid, ci, old, new in d["updated"]:
            label = to_text(self.original[rid][0])
            lines.append(f"수정  {label}  [{self.names[ci]}]  {to_text(old) or '(빈칸)'} → {to_text(new) or '(빈칸)'}")
        for rid in d["added"]:
            lines.append("추가  " + " | ".join(to_text(v) for v in self.rows[rid]))
        for rid in d["deleted"]:
            lines.append("삭제  " + " | ".join(to_text(v) for v in self.original[rid]))
        return lines


def find_key(session):
    """행을 구분할 수 있는 열(값이 비어 있지 않고 서로 다 다른 첫 열)의 이름. 없으면 None."""
    for ci, name in enumerate(session.names):
        vals = [row[ci] for row in session.original.values()]
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
        """첫 줄을 열 이름으로 보고 시트 전체를 읽는다."""
        from openpyxl import load_workbook
        wb = load_workbook(self.path)
        rows = [list(r) for r in wb[name].iter_rows(values_only=True)]
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
        data = [(r + [None] * ncol)[:ncol] for r in rows[1:]]
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
            data = [[float(v) if isinstance(v, Decimal) else v for v in row] for row in cur.fetchall()]
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
        keys = [row[ki] for row in session.original.values()]
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
                k = session.original[rid][ki]
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
                k = session.original[rid][ki]
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
# 5. App — 화면
# =====================================================================
import tkinter as tk                                   # noqa: E402
from tkinter import filedialog, messagebox, ttk        # noqa: E402

FONT = ("맑은 고딕", 10)
ALL = "전체 열"


class App:
    MAX_SHOW = 5000                 # 한 번에 화면에 그리는 최대 행 수

    def __init__(self, root):
        self.root = root
        self.source = None
        self.table = None
        self.session = None
        self.view = []              # 지금 화면에 보이는 행번호들
        self.sort = (None, False)
        self.editor = None
        self.last_col = 0
        root.title(TITLE)
        root.geometry("1240x720")
        root.minsize(900, 500)
        self._build()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Control-z>", lambda e: self.do_undo())
        root.bind("<Control-s>", lambda e: self.save())
        root.bind("<Control-f>", lambda e: self.find_entry.focus_set())
        root.bind("<Control-o>", lambda e: self.choose_file())

    # ---------------- 화면 만들기 ----------------
    def _build(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=26, font=FONT)
        style.configure("Treeview.Heading", font=(FONT[0], 10, "bold"))
        style.configure("TButton", font=FONT, padding=(8, 3))
        style.configure("TLabel", font=FONT)
        style.configure("TCheckbutton", font=FONT)
        style.configure("TRadiobutton", font=FONT)

        top = ttk.Frame(self.root, padding=(10, 8, 10, 2))
        top.pack(fill="x")
        ttk.Button(top, text="파일 열기", command=self.choose_file).pack(side="left")
        self.file_label = ttk.Label(top, text="엑셀(.xlsx) 또는 Access(.accdb) 파일을 여세요", foreground="#555")
        self.file_label.pack(side="left", padx=10)
        self.key_box = ttk.Combobox(top, state="readonly", width=16, font=FONT)
        self.key_label = ttk.Label(top, text="기준 열")
        self.table_box = ttk.Combobox(top, state="readonly", width=22, font=FONT)
        self.table_box.pack(side="right")
        self.table_box.bind("<<ComboboxSelected>>", self.on_table_change)
        self.table_label = ttk.Label(top, text="시트")
        self.table_label.pack(side="right", padx=(10, 4))

        bar = ttk.Frame(self.root, padding=(10, 4, 10, 6))
        bar.pack(fill="x")
        ttk.Label(bar, text="찾기").pack(side="left")
        self.find_col = ttk.Combobox(bar, state="readonly", width=14, values=[ALL], font=FONT)
        self.find_col.set(ALL)
        self.find_col.pack(side="left", padx=4)
        self.find_entry = ttk.Entry(bar, width=20, font=FONT)
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

        self.status = ttk.Label(self.root, text="", padding=(10, 4), foreground="#333")
        self.status.pack(fill="x", side="bottom")

        box = ttk.Frame(self.root, padding=(10, 0, 10, 0))
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
        self.hint = ttk.Label(box, justify="center", foreground="#666", font=(FONT[0], 12),
                              text="[파일 열기]로 시작하세요\n\n칸을 더블클릭하면 고칠 수 있습니다 · 노란 줄 = 고친 행 · 초록 줄 = 새 행\n"
                                   "저장하면 원본 옆 '백업' 폴더에 복사본을 먼저 남깁니다")
        self.hint.place(relx=0.5, rely=0.45, anchor="center")

    # ---------------- 파일 · 표 열기 ----------------
    def choose_file(self):
        path = filedialog.askopenfilename(
            title="엑셀 또는 Access 파일 열기",
            filetypes=[("엑셀 · Access", "*.xlsx *.xlsm *.accdb *.mdb"), ("모든 파일", "*.*")])
        if path:
            self.open_path(path)

    def open_path(self, path, ask=True):
        if ask and not self.confirm_discard():
            return False
        try:
            source = open_source(path)
            tables = source.tables()
            if not tables:
                raise ValueError("이 파일에는 시트/테이블이 없습니다.")
            self.source = source
            self.table_box["values"] = tables
            self.table_label["text"] = "테이블" if source.needs_key else "시트"
            self.load_table(tables[0])
        except Exception as e:
            messagebox.showerror(TITLE, str(e))
            return False
        return True

    def load_table(self, name):
        self.session = self.source.load(name)
        self.table = name
        self.table_box.set(name)
        s = self.session
        self.file_label.configure(text=f"{self.source.label} · {self.source.path.name}", foreground="#111")
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
        sample = list(s.rows.values())[:200]
        for ci, col in enumerate(s.columns):
            longest = max([len(col.name) * 2] + [sum(2 if ord(ch) > 255 else 1 for ch in to_text(r[ci])) for r in sample])
            self.tree.heading(ids[ci], text=col.name, command=lambda c=ci: self.sort_by(c))
            self.tree.column(ids[ci], width=min(max(longest * 8 + 24, 70), 280), stretch=False,
                             anchor="e" if col.kind in ("int", "float") else "w")
        self.hint.place_forget()
        self.sort = (None, False)
        self.view = s.ids()
        self.refresh()

    def on_table_change(self, event=None):
        name = self.table_box.get()
        if name == self.table:
            return
        if not self.confirm_discard():
            self.table_box.set(self.table)
            return
        try:
            self.load_table(name)
        except Exception as e:
            messagebox.showerror(TITLE, str(e))
            self.table_box.set(self.table)

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
        if s is None:
            self.status["text"] = ""
            return
        n = s.change_count()
        text = f"전체 {len(s.rows):,}건 · 표시 {len(self.view):,}건"
        if len(self.view) > self.MAX_SHOW:
            text += f" (화면에는 앞 {self.MAX_SHOW:,}건만)"
        text += f" · 저장 안 한 변경 {n:,}곳" if n else " · 변경 없음"
        self.status["text"] = text
        self.root.title(("● " if n else "") + f"{TITLE} — {self.source.path.name} [{self.table}]")

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
        if self.session is None or self.tree.identify("region", event.x, event.y) != "cell":
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
        if self.session is None:
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
        if self.session is None:
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
        before = set(self.session.ids())
        if self.session.undo():
            self.view += [rid for rid in self.session.ids() if rid not in before]   # 되살아난 행만 다시 보이게
            order = {rid: i for i, rid in enumerate(self.session.ids())}
            if self.sort[0] is None:
                self.view.sort(key=order.get)
            self.refresh()

    # ---------------- 일괄 바꾸기 ----------------
    def replace_dialog(self):
        if self.session is None:
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
        """저장하고 백업 파일 경로를 돌려준다(저장하지 않았으면 None)."""
        if self.session is None:
            return None
        self.commit_edit()
        s = self.session
        d = s.diff()
        if not (d["added"] or d["deleted"] or d["updated"]):
            if ask:
                messagebox.showinfo(TITLE, "바뀐 내용이 없습니다.")
            return None
        summary = f"수정 {len(d['updated']):,}칸 · 추가 {len(d['added']):,}행 · 삭제 {len(d['deleted']):,}행"
        if ask and not messagebox.askyesno(
                TITLE, f"{self.source.path.name} [{self.table}] 에 저장할까요?\n\n{summary}\n\n저장 전에 '백업' 폴더에 복사본을 만듭니다."):
            return None
        try:
            key = self.key_box.get() if self.source.needs_key else None
            backup = self.source.save(self.table, s, key=key)
            self.load_table(self.table)
        except Exception as e:
            messagebox.showerror(TITLE, str(e))
            return None
        if ask:
            messagebox.showinfo(TITLE, f"저장했습니다. ({summary})\n\n백업: 백업\\{backup.name}")
        return backup

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
        from openpyxl.styles import Font
        wb = Workbook()
        ws = wb.active
        ws.title = str(self.table)[:31]
        ws.append(self.session.names)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for rid in self.view:
            ws.append(self.session.rows[rid])
        ws.freeze_panes = "A2"
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
        if self.session is None or self.session.change_count() == 0:
            return True
        return messagebox.askyesno(TITLE, "저장하지 않은 변경이 있습니다. 버리고 계속할까요?")

    def on_close(self):
        if self.confirm_discard():
            self.root.destroy()


def main():
    try:                                            # 고해상도 화면에서 글자가 흐리지 않게 (Windows)
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    app = App(root)
    if len(sys.argv) > 1:
        app.open_path(sys.argv[1], ask=False)
    root.mainloop()


if __name__ == "__main__":
    main()
