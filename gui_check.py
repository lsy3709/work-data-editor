# -*- coding: utf-8 -*-
"""화면을 실제로 띄워 프로그램이 제대로 움직이는지 자동으로 확인한다 (약 10초, 창이 잠깐 떴다 닫힘).
실행: python gui_check.py            마지막 줄에 RESULT PASS 가 나오면 정상
      python gui_check.py 폴더       화면 캡처도 그 폴더에 저장 (Pillow 가 있을 때)
"""
import shutil
import sys
import tempfile
import tkinter as tk
from pathlib import Path

import data_editor as ed

HERE = Path(__file__).resolve().parent
SHOTS = Path(sys.argv[1]) if len(sys.argv) > 1 else None
OUT = Path(tempfile.mkdtemp(prefix="gui_check_"))        # 시험용 사본을 여기에 만든다 (samples 폴더는 건드리지 않음)
DATA = OUT / "data"
shutil.copytree(HERE / "samples_original", DATA)
if SHOTS:
    SHOTS.mkdir(parents=True, exist_ok=True)

try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass


def shot(root, name):
    if not SHOTS:
        return
    try:
        from PIL import ImageGrab
    except ImportError:
        return
    root.lift(); root.attributes("-topmost", True); root.update(); root.after(300); root.update()
    x, y = root.winfo_rootx(), root.winfo_rooty()
    ImageGrab.grab((x, y, x + root.winfo_width(), y + root.winfo_height())).save(SHOTS / name)


def check(cond, msg):
    print(("OK  " if cond else "FAIL"), msg)
    if not cond:
        global failed
        failed = True


failed = False
root = tk.Tk()
root.geometry("1240x720+40+40")
app = ed.App(root)
root.update()
shot(root, "0_start.png")

# ---------- 엑셀 ----------
xlsx = DATA / "보금자리론_신청_2026_10.xlsx"
check(app.open_path(xlsx, ask=False), "엑셀 열기")
root.update()
check(len(app.tree.get_children()) == 240, "240행 표시")
app.do_search("X99", "지점코드"); root.update()
check(len(app.tree.get_children()) == 13, "지점코드 X99 찾기 = 13행")
n = app.apply_replace("지점코드", "X99", "ON1", "whole", "view"); root.update()
check(n == 13, "일괄 바꾸기 13칸")
rid = app.add_row({"신청번호": "HF202610-99999", "신청일": "2026-10-31", "지점코드": "ON1",
                   "상품명": "보금자리론", "신청금액(만원)": "12,000", "대출기간(년)": "30",
                   "고객구분": "일반", "처리상태": "심사중"})
root.update()
check(app.session.status(rid) == "added", "행 추가")
shot(root, "1_excel_edit.png")
app.tree.selection_set(app.tree.get_children()[0])
check(app.delete_selected(ask=False) == 1, "행 삭제")
app.do_undo(); root.update()
check(len(app.tree.get_children()) == 14, "되돌리기로 삭제 취소")
lines = app.session.describe_changes()
check(len(lines) == 14, f"변경 내역 14건 (실제 {len(lines)})")
backup = app.save(ask=False); root.update()
check(backup is not None and backup.exists(), "엑셀 저장 + 백업")
check(app.session.change_count() == 0 and len(app.session.rows) == 241, "다시 읽으니 241행, 변경 0")
check(len(app.session.search("ON1", "지점코드")) == 14, "저장된 값 확인")
app.sort_by(app.session.col("신청금액(만원)")); app.sort_by(app.session.col("신청금액(만원)")); root.update()
first = app.session.rows[int(app.tree.get_children()[0])][app.session.col("신청금액(만원)")]
check(first == max(r[app.session.col("신청금액(만원)")] for r in app.session.rows.values()), "금액 내림차순 정렬")
n = app.export_view(OUT / "내보내기_시험.xlsx")
check(n == 241, "내보내기 241행")

# 칸 더블클릭 편집 (실제 이벤트)
app.show_all(); root.update()
iid = app.tree.get_children()[2]
bx = app.tree.bbox(iid, "#8")
ev = type("E", (), {"x": bx[0] + 5, "y": bx[1] + 5})()
app.begin_edit(ev); root.update()
check(app.editor is not None, "더블클릭하면 입력 칸이 열림")
app.editor[0].delete(0, "end"); app.editor[0].insert(0, "검토완료")
shot(root, "2_cell_edit.png")
app.commit_edit(); root.update()
check(app.session.get(int(iid), "처리상태") == "검토완료", "칸 수정 반영")
app.do_undo()

# 여러 시트 파일
check(app.open_path(DATA / "S1_보금자리론_취합.xlsx", ask=False), "여러 시트 엑셀 열기")
check(list(app.table_box["values"]) == ["전체내역", "월별요약", "지점별요약"], "시트 목록")
app.table_box.set("월별요약"); app.on_table_change(); root.update()
check(len(app.tree.get_children()) == 3, "월별요약 3행")

# ---------- Access ----------
try:
    import pyodbc  # noqa
    have = True
except ImportError:
    have = False
if have:
    check(app.open_path(DATA / "업무DB.accdb", ask=False), "Access 열기")
    root.update()
    check(app.key_box.get() == "신청번호", "기준 열 자동 선택 = 신청번호")
    check(len(app.tree.get_children()) == 240, "Access 240행")
    app.do_search("반려", "처리상태"); root.update()
    cnt = len(app.tree.get_children())
    n = app.apply_replace("처리상태", "", "재심사", "fill", "view"); root.update()
    check(n == cnt and cnt > 0, f"반려 {cnt}건을 재심사로 채우기")
    shot(root, "3_access_edit.png")
    backup = app.save(ask=False); root.update()
    check(backup is not None, "Access 저장 + 백업")
    check(len(app.session.search("재심사", "처리상태")) == cnt and app.session.change_count() == 0, "Access 저장값 확인")
else:
    print("SKIP Access (pyodbc 없음)")

root.destroy()
shutil.rmtree(OUT, ignore_errors=True)
print("RESULT", "FAIL" if failed else "PASS")
sys.exit(1 if failed else 0)
