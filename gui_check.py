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
    try:                                             # 마우스 포인터가 찍히지 않게 화면 구석으로 보낸다
        ctypes.windll.user32.SetCursorPos(ctypes.windll.user32.GetSystemMetrics(0) - 2, ctypes.windll.user32.GetSystemMetrics(1) - 2)
    except Exception:
        pass
    root.lift(); root.attributes("-topmost", True); root.update(); root.after(400); root.update()
    x, y = root.winfo_rootx(), root.winfo_rooty()
    ImageGrab.grab((x, y, x + root.winfo_width(), y + root.winfo_height())).save(SHOTS / name)


def check(cond, msg):
    print(("OK  " if cond else "FAIL"), msg)
    if not cond:
        global failed
        failed = True


def open_file(path):
    """파일을 탭으로 열고, 작업 스레드가 다 읽을 때까지 기다린다."""
    pane = app.open_path(path)
    return bool(pane) and app.wait_idle() and pane.session is not None


def save_now():
    """저장하고(작업 스레드) 끝나면 백업 파일 경로를 돌려준다."""
    pane = app.current
    if not pane.save(ask=False):
        return None
    app.wait_idle()
    return pane.last_backup


failed = False
root = tk.Tk()
root.geometry("1240x720+40+40")
app = ed.App(root)
app.show_error = lambda text: print("ERROR", text)      # 점검 중에는 오류 창 대신 글자로
root.update()
check(app.current is None and not app.panes, "처음에는 탭이 없음")
shot(root, "0_start.png")

# ---------- 엑셀 ----------
xlsx = DATA / "보금자리론_신청_2026_10.xlsx"
check(open_file(xlsx), "엑셀 열기")
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
backup = save_now(); root.update()
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
check(open_file(DATA / "S1_보금자리론_취합.xlsx"), "여러 시트 엑셀 열기 (두 번째 탭)")
check(list(app.table_box["values"]) == ["전체내역", "월별요약", "지점별요약"], "시트 목록")
app.table_box.set("월별요약"); app.on_table_change(); app.wait_idle()
check(len(app.tree.get_children()) == 3, "월별요약 3행")

# ---------- Access ----------
try:
    import pyodbc  # noqa
    have = True
except ImportError:
    have = False
if have:
    check(open_file(DATA / "업무DB.accdb"), "Access 열기 (세 번째 탭)")
    root.update()
    check(app.key_box.get() == "신청번호", "기준 열 자동 선택 = 신청번호")
    check(len(app.tree.get_children()) == 240, "Access 240행")
    app.do_search("반려", "처리상태"); root.update()
    cnt = len(app.tree.get_children())
    n = app.apply_replace("처리상태", "", "재심사", "fill", "view"); root.update()
    check(n == cnt and cnt > 0, f"반려 {cnt}건을 재심사로 채우기")
    shot(root, "3_access_edit.png")
    backup = save_now(); root.update()
    check(backup is not None, "Access 저장 + 백업")
    check(len(app.session.search("재심사", "처리상태")) == cnt and app.session.change_count() == 0, "Access 저장값 확인")
else:
    print("SKIP Access (pyodbc 없음)")

# ---------- 여러 탭 ----------
import threading  # noqa: E402

n_tabs = len(app.panes)
check(n_tabs == (3 if have else 2) and len(app.notebook.tabs()) == n_tabs, f"파일마다 탭 하나씩 = {n_tabs}개")
first = app.panes[0]
again = app.open_path(xlsx); app.wait_idle()
check(again is first and len(app.panes) == n_tabs and app.current is first, "이미 연 파일은 새 탭 없이 그 탭으로 이동")
second = app.panes[1]
check(second.table == "월별요약" and len(second.session.rows) == 3, "다른 탭의 시트·내용은 그대로")
first.session.set_cell(first.session.ids()[0], "처리상태", "탭1만"); first.refresh()
check(first.is_dirty() and not second.is_dirty(), "한 탭에서 고쳐도 다른 탭은 변경 없음")
check(app.notebook.tab(first.frame, "text").startswith("● ") and
      not app.notebook.tab(second.frame, "text").startswith("● "), "고친 탭 이름에만 ● 표시")
app.notebook.select(second.frame); root.update()
check(app.current is second and "S1_보금자리론_취합" in root.title(), "탭을 바꾸면 창 제목도 바뀜")
shot(root, "4_tabs.png")

# 여러 파일을 한꺼번에 열기 — 작업 스레드가 동시에 읽는다
extra = [DATA / "2026_10_부산지사.xlsx", DATA / "지점코드표.xlsx"]
before_threads = threading.active_count()
opened = app.open_paths(extra)
peak_jobs, peak_threads = app.jobs.pending, threading.active_count()
check(len(opened) == 2 and all(p.busy and p.session is None for p in opened), "여는 즉시 탭이 생기고 '불러오는 중' 상태")
root.update()
check(app.current is not None and app.current.tree.winfo_exists(), "읽는 동안에도 화면은 살아 있음")
check(app.wait_idle() and all(p.session is not None and not p.busy for p in opened),
      f"{len(extra)}개 파일 동시에 열기 (동시 작업 {peak_jobs}개, 스레드 {before_threads}→{peak_threads})")
check(peak_jobs == len(extra) and peak_threads > before_threads, "파일마다 작업 스레드가 따로 돎")
check(threading.active_count() == before_threads, "끝나면 작업 스레드가 남지 않음")

# 저장하는 동안에는 그 탭만 잠기고, 닫을 수 없다
pane = opened[0]
app.notebook.select(pane.frame); root.update()
pane.session.set_cell(pane.session.ids()[0], pane.session.names[-1], "저장시험")
check(pane.save(ask=False) and pane.saving, "저장은 작업 스레드에서 시작")
check(not pane.ready() and pane.delete_selected(ask=False) == 0, "저장 중인 탭은 고칠 수 없음")
check(first.ready(), "다른 탭은 계속 고칠 수 있음")
app.wait_idle()
check(not pane.saving and pane.last_backup is not None and not pane.is_dirty(), "저장이 끝나면 다시 쓸 수 있음")

# ---------- 회귀 분석 (pandas · scikit-learn · matplotlib 이 있을 때) ----------
try:
    import analysis_window  # noqa: F401
    have_ml = True
except ImportError as e:
    have_ml = False
    print("SKIP 회귀 분석 (라이브러리 없음):", e)
if have_ml:
    check(open_file(DATA / "주택_감정가_분석연습.xlsx"), "분석 연습 파일 열기")
    house = app.current
    win = house.open_analysis()
    win.error = lambda text: print("ERROR", text)
    root.update()
    target, features = win.roles()
    check(target == "감정가(만원)" and len(features) == 7 and "물건번호" not in features, "① 변수 추천: 감정가를 7개 열로 예측")
    intro = win.model_text.get("1.0", "end")
    check("이름의 뜻" in intro and "역사" in intro and "부스팅" in intro, "④ 모델 소개(어원 · 역사) 표시")
    win.set_roles("지역", ["층"])
    win.target_box.set("지역")
    errors = []
    win.error = errors.append
    win.run(); app.wait_idle()
    check(len(errors) == 1 and "숫자 열이 아니어서" in errors[0] and win.result is None, "글자 열을 예측하려 하면 이유를 알려 줌")
    win.error = lambda text: print("ERROR", text)
    win.use_suggestion()
    win.set_model("linear")
    win.set_outlier_method("none")
    check(win.run() and win.busy and not win.run(), "분석은 작업 스레드에서 (도는 동안 다시 누르면 무시)")
    app.wait_idle(); root.update()
    plain = win.result.metrics["r2_test"]
    check(win.result.key == "linear" and len(win.clean.X) == 596, f"② 정제: 604행 → 596행, 선형 회귀 R² {plain:.3f}")
    check(win.canvas is not None and len(win.fig.axes) == 4, "⑤ 그래프 4장")
    win.set_outlier_method("combo")
    win.run(); app.wait_idle(); root.update()
    n_out = int(win.outliers.mask.sum())
    better = win.result.metrics["r2_test"]
    check(better > plain and 9 <= n_out < 30, f"③ 이상 탐지로 {n_out}행을 빼니 R² {plain:.3f} → {better:.3f}")
    check(len(win.outlier_tree.get_children()) == n_out, "이상 행 목록에 이유와 함께 표시")
    check(win.run_auto(), "자동 분석 시작")
    root.update()
    check(house.ready() and first.ready(), "분석하는 동안에도 편집기 탭은 쓸 수 있음")
    app.wait_idle(180); root.update()
    best = win.table[0]
    check(len(win.table) == 9 and win.result.key == best["key"] and best["r2"] > 0.93,
          f"자동 분석: 9개 모델 비교, 1위 {best['name']} (R² {best['r2']:.3f})")
    check(len(win.compare_tree.get_children()) == 9 and "자동 분석이 한 일" in win.summary.get("1.0", "end"), "모델 비교표 · 요약 표시")

    # 통계 검정 (t 검정 · 분산분석 · 회귀 계수 검정)
    by = {t.feature: t for t in win.tests}
    check(len(win.tests) == 7 and len(win.test_tree.get_children()) == 7, "통계 검정: 설명하는 열 7개를 하나씩 검정")
    check("상관" in by["전용면적(㎡)"].method and by["전용면적(㎡)"].significant, "숫자 열 = 상관 t 검정 (전용면적: 유의함)")
    check("분산분석" in by["지역"].method and len(by["지역"].pairs) == 15, "세 집단 이상 = 분산분석 + 사후 t 검정 (지역: 15쌍)")
    check(not by["세대수"].significant and "유의하지 않음" in win.test_tree.item("6", "values")[-1], "관계없는 열은 '유의하지 않음' (세대수)")
    rows = {r["name"]: r for r in win.ols.rows}
    check(len(win.ols_tree.get_children()) == len(win.ols.rows) and abs(rows["전용면적(㎡)"]["t"]) > 10 and rows["전용면적(㎡)"]["significant"],
          f"회귀 계수 t 검정: 전용면적 t = {rows['전용면적(㎡)']['t']:.1f}, 조정 R² {win.ols.adj_r2:.3f}")
    win.test_tree.selection_set("0"); root.update()
    detail = win.test_text.get("1.0", "end")
    check("서울" in detail and "평균" in detail and "≠" in detail, "검정 줄을 누르면 집단별 평균과 둘씩 견준 결과가 나옴")
    check("tests" in win.figures and len(win.figures["tests"].axes) == 6, "검정 그래프 6장 (산점도 · 상자그림 · t 값 · p 값)")

    # 학습 과정 (진행 막대 · 기록 · 학습 곡선)
    log = win.log_text.get("1.0", "end")
    check(all(word in log for word in ("정제", "이상한 행", "모델 비교 중 (9/9)", "통계 검정", "학습 곡선", "끝")) and log.count("초") > 15,
          f"학습 과정 기록 {len(log.splitlines())}줄 (정제 → 이상 탐지 → 모델 9개 → 검정 → 학습 곡선)")
    check(win.progress["value"] == 100 and len(win._rows) == 9, "진행 막대 100%, 모델이 끝날 때마다 비교표에 한 줄씩")
    check(len(win.curve["sizes"]) == 5 and win.curve["valid"][-1] > win.curve["valid"][0] and "learn" in win.figures,
          f"학습 곡선: 행을 {win.curve['sizes'][0]} → {win.curve['sizes'][-1]}개로 늘리면 시험 R² {win.curve['valid'][0]:.3f} → {win.curve['valid'][-1]:.3f}")

    for index, name in ((0, "a1_summary"), (1, "a2_graph"), (2, "a3_tests"), (3, "a4_test_graph"), (4, "a5_learning"),
                        (5, "a6_compare"), (6, "a7_outliers"), (7, "a8_model_info")):
        win.tabs.select(index); root.update()
        shot(win.win, name + ".png")
    win.tabs.select(0); root.update()
    win.compare_tree.selection_set("forest"); win.use_compared_model(); app.wait_idle(); root.update()
    check(win.result.key == "forest" and len(win.table) == 9, "비교표에서 고른 모델로 다시 분석 (비교표는 유지)")
    files = win.export(OUT / "분석결과.xlsx")
    check(len(files) == 4 and all(f.exists() and f.stat().st_size > 5000 for f in files) and files[1].suffix == ".png",
          "결과 내보내기: 엑셀(검정 · 계수 시트 포함) + 그림 3장")
    shown = win.show_outliers_in_editor(); root.update()
    check(shown == int(win.outliers.mask.sum()) and len(house.tree.get_children()) == shown and app.current is house,
          f"이상 행 {shown}개만 편집기에서 보기")
    rid = int(house.tree.get_children()[0])
    house.session.set_cell(rid, "감정가(만원)", "50000"); house.refresh()
    win.set_model("boosting"); win.run(); app.wait_idle()
    check(float(win.df.at[rid, "감정가(만원)"]) == 50000, "편집기에서 고친 값이 다음 분석에 반영됨")
    check(app.close_pane(house, ask=False) and not win.win.winfo_exists(), "탭을 닫으면 분석 창도 닫힘")

    # ---------- 수업 샘플 데이터로 정제 · 분석 · 검정 ----------
    app.notebook.select(first.frame); root.update()                    # 보금자리론_신청_2026_10.xlsx (앞에서 고친 탭)
    loan = first.open_analysis()
    loan.error = lambda text: print("ERROR", text)
    target, features = loan.roles()
    check(target == "신청금액(만원)" and "신청번호" not in features and "대출기간(년)" in features, "수업 샘플(보금자리론 신청): 신청금액을 예측하도록 추천")
    loan.run_auto(); app.wait_idle(180); root.update()
    lt = {t.feature: t for t in loan.tests}
    yes = [t.feature for t in loan.tests if t.significant]
    check(loan.result is not None and len(loan.tests) == len(features) and not lt["대출기간(년)"].significant,
          f"수업 샘플 검정: {len(loan.tests)}개 열 중 유의 {len(yes)}개 {yes}, 1위 모델 R² {loan.table[0]['r2']:.3f} (관계가 약한 데이터)")
    check(loan.table[0]["r2"] < 0.5, "관계가 약하면 어떤 모델도 잘 못 맞힌다는 것이 그대로 나옴")
    loan.tabs.select(2); root.update()
    shot(loan.win, "b1_loan_tests.png")
    loan.tabs.select(3); root.update()
    shot(loan.win, "b2_loan_test_graph.png")
    loan.close()

    check(open_file(DATA / "주택연금_가입현황_정제전.xlsx"), "수업 샘플(주택연금 정제전) 열기 — 숫자 · 날짜가 글자로 적힌 표")
    pension = app.current
    pw = pension.open_analysis()
    pw.error = lambda text: print("ERROR", text)
    target, features = pw.roles()
    check(target == "월지급액" and "성명" not in features, "글자로 적힌 금액을 숫자로 알아보고 월지급액을 예측하도록 추천")
    pw.set_model("linear"); pw.run(); app.wait_idle(120); root.update()
    report = " ".join(pw.clean.report)
    check(pw.result is not None and "글자로 적힌 숫자" in report and "날짜" in report and pw.clean.dropped["duplicate"] == 2,
          f"정제: 금액 · 날짜 바로잡기, 겹친 행 2개 제거 → {len(pw.clean.X)}행")
    check(len(pw.tests) == len(pw.clean.features) and pw.ols is not None, "정제한 표로 검정 · 회귀 계수 검정까지")
    pw.tabs.select(4); root.update()
    shot(pw.win, "b3_pension_learning.png")
    pw.tabs.select(0); root.update()
    shot(pw.win, "b4_pension_summary.png")
    check(app.close_pane(pension, ask=False), "분석한 탭 닫기")

# ---------------- 복사본 만들기 · 전처리 · 다른 이름으로 저장 ----------------
PRE = OUT / "pre"
PRE.mkdir()
for name in ("주택연금_가입현황_정제전.xlsx", "주택_감정가_분석연습.xlsx", "업무DB.accdb"):
    shutil.copy2(HERE / "samples_original" / name, PRE / name)
raw = PRE / "주택연금_가입현황_정제전.xlsx"
before_bytes = raw.read_bytes()
check(open_file(raw), "전처리할 표 열기 — 숫자 · 날짜가 모두 글자로 적힌 주택연금 표 (62행)")
origin = app.current
tabs = len(app.panes)
copy = app.duplicate(ask=False)
app.wait_idle()
root.update()
check(copy is not None and copy.source.path.name == "주택연금_가입현황_정제전_복사본.xlsx" and copy.source.path.exists()
      and len(app.panes) == tabs + 1 and app.current is copy and len(copy.session.rows) == 62,
      "[복사본 만들기] 원본 옆에 _복사본 파일을 만들어 새 탭으로 열기")
w = copy.open_preprocess()
w.quiet = True
root.update()
check(w is not None and copy.open_preprocess() is w, "[전처리] 창 열기 (같은 탭에서 다시 누르면 같은 창)")
check(all(copy.session.column_summary(n)["kind"] == "str" for n in copy.session.names), "처음에는 모든 열이 글자")
w.select_empty()
check(set(w.selected()) == {"가입일", "월지급액", "관할지점"}, "[빈칸 있는 열 고르기] 빈칸이 있는 열 3개가 골라짐")
shot(w.win, "c1_preprocess_before.png")

w.select([])
check(w.strip() and copy.session.get(copy.session.ids()[0], "성명") == "장재훈", "앞뒤 빈칸 지우기 (' 장재훈 ' → '장재훈')")
check(w.drop_duplicates() and len(copy.session.rows) == 60 and len(copy.tree.get_children()) == 60, "겹친 행 삭제 → 60행, 표 화면도 60행")
w.select(["월지급액", "출생연도"])
check(w.convert("number") and copy.session.columns[copy.session.col("월지급액")].kind == "int"
      and copy.session.get(copy.session.ids()[1], "월지급액") == 2430000,
      "종류 바꾸기: 글자 → 숫자 ('2,430,000원' → 2430000), 두 열을 한 번에")
check(str(copy.tree.column("c5", "anchor")) == "e", "숫자가 된 열은 표에서 오른쪽 맞춤으로 바뀜")
w.select(["가입일"])
check(w.convert("date") and copy.session.columns[copy.session.col("가입일")].kind == "date", "종류 바꾸기: 글자 → 날짜 ('2026/07/15' · '2026.01.03' 섞인 표기)")
w.select(["관할지점"])
check(not w.fill("mean") and "숫자 열이 아니어서" in w.last_error, "글자 열을 평균으로 채우려 하면 이유를 알려 주고 아무것도 바꾸지 않음")
check(w.fill("mode") and copy.session.empty_count("관할지점") == 0, "글자 열 빈칸은 가장 흔한 값으로")
w.select(["월지급액"])
check(w.fill("median") and copy.session.empty_count("월지급액") == 0, "숫자 열 빈칸은 중앙값으로")
check(w.scale(divide=True, factor=10000) and copy.session.get(copy.session.ids()[1], "월지급액") == 243, "숫자 계산: 10000 으로 나누기 (원 → 만원)")
check(w.rename("월지급액(만원)") and "월지급액(만원)" in copy.session.names and copy.tree.heading("c5", "text").startswith("월지급액(만원)"),
      "열 이름 바꾸기 — 표 머리글도 바뀜")
w.select(["가입일"])
check(w.drop_empty() and len(copy.session.rows) == 59, "빈칸 있는 행 삭제 (가입일이 빈 1행) → 59행")
w.select(["출생연도"])
check(w.add_column("확인", "str", "미확인") and copy.session.names.index("확인") == 3 and len(copy.tree["columns"]) == 8, "열 추가 (고른 열 바로 뒤) → 8열")
w.select(["가입번호"])
check(w.delete_columns() and "가입번호" not in copy.session.names and len(copy.tree["columns"]) == 7, "고른 열 삭제 → 7열")
check(w.undo() and copy.session.names[0] == "가입번호" and len(copy.tree["columns"]) == 8, "[되돌리기] 지운 열이 값과 함께 돌아옴")
check(copy.session.structure_changed and "저장 안 한 변경" in copy.status.cget("text"), "상태 줄에 저장 안 한 변경이 표시됨")
w.select(["월지급액(만원)", "출생연도", "가입일"])
shot(w.win, "c2_preprocess_after.png")
shot(root, "c3_preprocess_table.png")

clean = PRE / "주택연금_정리본.xlsx"
check(copy.save_as(clean, ask=False) and app.wait_idle() and clean.exists(), "[다른 이름으로 저장] 정리한 표를 새 파일에")
root.update()
check(copy.source.path == clean and "주택연금_정리본.xlsx" in copy.file_label.cget("text") and not copy.is_dirty()
      and copy.session.names == ["가입번호", "성명", "출생연도", "확인", "가입일", "지급방식", "월지급액(만원)", "관할지점"]
      and len(copy.session.rows) == 59, "저장한 뒤 탭이 새 파일을 가리키고, 다시 읽은 표가 고친 그대로 (8열 · 59행)")
kinds = {n: copy.session.column_summary(n)["kind"] for n in copy.session.names}
check(kinds["월지급액(만원)"] in ("int", "float") and kinds["출생연도"] == "int" and kinds["가입일"] == "date",
      f"다시 읽어도 숫자 · 날짜로 읽힘 ({kinds['월지급액(만원)']} · {kinds['출생연도']} · {kinds['가입일']})")
check(raw.read_bytes() == before_bytes and not (PRE / "백업").exists(), "원본 파일은 한 글자도 바뀌지 않음 (백업 폴더도 만들지 않음)")
check("주택연금 정리본" not in w.win.title() and "주택연금_정리본.xlsx" in w.win.title(), "전처리 창 제목도 새 파일 이름으로")
other = PRE / "주택연금_가입현황_정제전_복사본.xlsx"
check(origin.save_as(other, ask=False) and app.wait_idle() and origin.source.path == other, "바뀐 것이 없는 표도 다른 이름으로 저장 가능 (다른 탭이 쓰지 않는 파일 이름)")
check(not origin.save_as(clean, ask=False), "다른 탭에서 열어 둔 파일 이름으로는 저장하지 않음")
w.close()
check(w not in copy.windows, "전처리 창 닫기")

check(open_file(PRE / "주택_감정가_분석연습.xlsx"), "숫자 표(주택 감정가 604행) 열기")
house = app.current
hw = house.open_preprocess()
hw.quiet = True
hw.select_empty()
check(hw.fill("mean") and all(house.session.empty_count(n) == 0 for n in house.session.names)
      and house.session.columns[house.session.col("건축연수(년)")].kind == "float",
      "빈칸 있는 열 3개를 평균으로 한 번에 채움 (정수 열이 소수 평균으로 채워지면 실수 열이 됨)")
check(hw.undo() and house.session.empty_count("건축연수(년)") == 10 and house.session.change_count() == 0
      and house.session.columns[house.session.col("건축연수(년)")].kind == "int",
      "여러 열에 한 일도 [되돌리기] 한 번으로 모두 되돌아감")
check(hw.fill("median") and hw.drop_duplicates() and len(house.session.rows) == 600, "중앙값으로 채우고 겹친 행 4개 삭제 → 600행")
shot(hw.win, "c4_preprocess_numbers.png")
backup = save_now()
check(backup is not None and backup.exists() and not house.is_dirty() and len(house.session.rows) == 600, "[저장] 원본 엑셀에 저장 (백업 만들고 600행)")
hw.close()

if not have:
    print("SKIP Access 표를 엑셀로 저장 (pyodbc 없음)")
else:
    check(open_file(PRE / "업무DB.accdb"), "Access 표 열기 (전처리용 사본)")
    if app.current.source.needs_key:
        acc = app.current
        aw = acc.open_preprocess()
        aw.quiet = True
        col1 = acc.session.names[1]
        aw.select([col1])
        check(aw.rename(col1 + "_새이름") and acc.session.structure_changed, "Access 표의 열 이름 바꾸기")
        check(not acc.save(ask=False), "열 구조를 바꾼 Access 표는 [저장]을 막고 다른 이름으로 저장을 안내")
        out = PRE / "신청내역_정리.xlsx"
        check(acc.save_as(out, ask=False) and app.wait_idle() and out.exists() and acc.source.path == out
              and acc.session.names[1] == col1 + "_새이름" and not acc.source.needs_key,
              "Access 표를 [다른 이름으로 저장] → 엑셀 파일, 탭이 그 엑셀을 가리킴")
        root.update()
        check(not acc.key_box.winfo_ismapped() and acc.table_label.cget("text") == "시트", "엑셀이 되었으니 '기준 열' 상자가 사라짐")
        aw.close()
        app.close_pane(acc, ask=False)
for p in (house, copy, origin):
    app.close_pane(p, ask=False)
shot(root, "c5_toolbar.png")

# ---------------- AI 채팅 (가짜 AI 서버로 — 인터넷과 API 키 없이) ----------------
import json
import os

import ai_mock_server

AI_DIR = OUT / "ai"
AI_DIR.mkdir()
shutil.copy2(HERE / "samples_original" / "보금자리론_신청_2026_10.xlsx", AI_DIR / "보금자리론_신청_2026_10.xlsx")
ANSWERS = {
    "X99": {"reply": "지점코드 열의 X99 를 ON1 로 바꿉니다.",
            "actions": [{"op": "replace", "column": "지점코드", "find": "X99", "to": "ON1", "mode": "cell"}]},
    "재심사": {"reply": "신청금액이 3억 이상이고 반려된 건을 재심사로 바꾸고, 금액(억원) 열을 만듭니다.",
            "actions": [{"op": "set", "column": "처리상태", "value": "재심사",
                         "where": [{"column": "신청금액(만원)", "is": ">=", "value": 30000}, {"column": "처리상태", "is": "=", "value": "반려"}]},
                        {"op": "compute", "target": "신청금액(억원)", "left": "신청금액(만원)", "operator": "/", "right": 10000}]},
    "평균": {"reply": "고객구분별 평균 신청금액을 계산합니다.",
           "actions": [{"op": "stats", "column": "신청금액(만원)", "how": "mean", "by": "고객구분"}]},
    "보여": {"reply": "디딤돌 상품만 금액이 큰 순서로 보여 드립니다.",
           "actions": [{"op": "show", "where": [{"column": "상품명", "is": "contains", "value": "디딤돌"}]},
                       {"op": "sort", "column": "신청금액(만원)", "descending": True}]},
    "없는": {"reply": "등급 열을 A 로 채웁니다.", "actions": [{"op": "set", "column": "등급", "value": "A"}]},
    "지워": {"reply": "일반 고객 행을 지웁니다.", "actions": [{"op": "delete_rows", "where": [{"column": "고객구분", "is": "=", "value": "일반"}]}]},
}


def fake_ai(question):
    for word, answer in ANSWERS.items():
        if word in question:
            return json.dumps(answer, ensure_ascii=False)
    return json.dumps({"reply": "표와 관련된 일을 말씀해 주세요.", "actions": []}, ensure_ascii=False)


def say(window, text):
    """AI 창에 말을 보내고 답이 처리될 때까지 기다린다."""
    sent = window.send(text)
    app.wait_idle(60)
    root.update()
    return sent


mock = ai_mock_server.start(fake_ai)
check(open_file(AI_DIR / "보금자리론_신청_2026_10.xlsx"), "AI 로 고칠 표 열기 (보금자리론 신청 240행)")
ai_pane = app.current
aw = ai_pane.open_ai()
aw.quiet = True
root.update()
check(aw is not None and ai_pane.open_ai() is aw, "[AI 채팅] 창 열기 (같은 탭에서 다시 누르면 같은 창)")
check(set(aw.keys) == {"gemini", "openai", "claude", "custom"} and aw.model_box.get() != "", "AI 네 가지(Gemini · GPT · Claude · 직접 지정)와 기본 모델 이름")
aw.choose("claude")
check("claude-opus-5-5" in aw.model_box["values"] and not aw.url_row.winfo_ismapped(), "AI 를 바꾸면 모델 목록이 바뀜 (주소 칸은 직접 지정일 때만)")
aw.choose("custom", model="mock-model", key="mock-key-1234", url=mock.url + "/v1")
root.update()
check(aw.url_row.winfo_ismapped() and aw.key_entry.cget("show") == "•", "직접 지정: 주소 칸이 보이고, 키는 화면에 가려서 표시")

check(say(aw, "지점코드가 X99 인 것을 ON1 로 바꿔 줘") and aw.pending is not None, "말을 보내면 AI 가 명령을 만들어 옴")
plan_text = aw.plan_box.get("1.0", "end")
check("13칸" in plan_text and ai_pane.session.change_count() == 0 and str(aw.apply_button["state"]) == "normal",
      "적용 전 미리 보기: '13칸'이라고 알려 주고 표는 아직 그대로")
sent = mock.requests[-1]
system_text = sent["body"]["messages"][0]["content"]
first_id = str(ai_pane.session.rows[ai_pane.session.ids()[0]][0])
check(all(n in system_text for n in ai_pane.session.names) and first_id not in json.dumps(sent["body"], ensure_ascii=False),
      "AI 에게는 열 이름만 가고 칸의 값은 가지 않음")
check(sent["headers"].get("authorization") == "Bearer mock-key-1234" and "mock-key-1234" not in json.dumps(sent["body"], ensure_ascii=False)
      and "mock-key-1234" not in aw.show_payload(), "키는 머리말로만 보내고, '보내는 내용 보기'에도 나오지 않음")
shot(aw.win, "d1_ai_preview.png")
check(aw.apply() and len(ai_pane.session.search("X99", "지점코드")) == 0 and ai_pane.session.change_count() == 13
      and str(aw.apply_button["state"]) == "disabled", "[적용] → 13칸이 바뀜")
check(aw.undo() and ai_pane.session.change_count() == 0, "[되돌리기] → 처음 그대로")

check(say(aw, "3억 이상이면서 반려된 건은 재심사로 바꾸고 억원 열도 만들어 줘") and aw.pending and len(aw.last_lines) == 2, "명령 두 개짜리 요청")
check(aw.apply() and "신청금액(억원)" in ai_pane.session.names and len(ai_pane.tree["columns"]) == 9, "[적용] → 조건에 맞는 칸을 바꾸고 계산 열이 생김 (표 화면도 9열)")
check("사용자가 적용했다" in aw.note, "다음 말에 '직전 명령을 적용했다'는 것을 AI 에게 알려 줌")
check(aw.undo() and "신청금액(억원)" not in ai_pane.session.names and ai_pane.session.change_count() == 0, "명령 두 개가 [되돌리기] 한 번으로 돌아감")

check(say(aw, "고객구분별 평균 신청금액을 알려 줘") and aw.pending is None and "[고객구분] 별" in aw.last_lines[0],
      "집계는 확인 없이 바로 계산해 보여 줌 (표는 바꾸지 않음)")
check("직전" not in mock.requests[-1]["body"]["messages"][-1]["content"] or "되돌렸다" in mock.requests[-1]["body"]["messages"][-1]["content"],
      "집계 결과(숫자)는 AI 에게 보내지 않음")
check(say(aw, "디딤돌만 금액 큰 순서로 보여 줘") and 0 < len(ai_pane.view) < 240 and len(ai_pane.tree.get_children()) == len(ai_pane.view)
      and ai_pane.session.change_count() == 0, f"보기 · 정렬: 표 화면에 {len(ai_pane.view)}행만 큰 금액부터")
shot(aw.win, "d2_ai_stats.png")
shot(root, "d3_ai_table.png")

check(say(aw, "없는 열을 채워 줘") and aw.pending is None and "등급" in aw.last_error and ai_pane.session.change_count() == 0,
      "AI 가 없는 열을 말하면 실행하지 않고 이유를 알려 줌")
check(say(aw, "일반 고객은 지워 줘") and aw.pending is not None and aw.discard() and len(ai_pane.session.rows) == 240 and "버렸다" in aw.note,
      "[버리기] → 표는 그대로, AI 에게는 버렸다고 알려 줌")
check(say(aw, "안녕") and aw.pending is None and ai_pane.session.change_count() == 0, "표와 상관없는 말에는 명령 없이 말로만 답함")

aw.samples.set(True)
check(first_id in aw.show_payload(), "'예시 값도 보내기'를 켰을 때만 앞 3행이 보내는 내용에 들어감")
aw.samples.set(False)
mock.stop()
check(say(aw, "X99 를 바꿔 줘") and "연결" in aw.last_error, "AI 서버에 연결이 안 되면 우리말로 알려 줌")
saved = os.environ.pop("OPENAI_API_KEY", None)
aw.choose("openai")
check(aw.key_entry.get() == "" and say(aw, "X99 를 바꿔 줘") and "API 키가 없습니다" in aw.last_error, "키를 넣지 않으면 보내지 않고 알려 줌")
if saved:
    os.environ["OPENAI_API_KEY"] = saved
aw.choose("custom")
check(aw.key_entry.get() == "mock-key-1234" and app.ai_state["keys"]["custom"] == "mock-key-1234", "AI 를 오갈 때 넣어 둔 키 · 모델을 기억 (메모리에만)")
leaked = [p.name for p in list(HERE.glob("*")) + list(AI_DIR.glob("*")) if p.is_file() and p.suffix in (".json", ".ini", ".cfg", ".env", ".txt")
          and "mock-key-1234" in p.read_text(encoding="utf-8", errors="ignore")]
check(not leaked, "키를 어떤 파일에도 쓰지 않음")
aw.close()
check(aw not in ai_pane.windows and app.close_pane(ai_pane, ask=False), "AI 창 닫기 · 탭 닫기")

# 없는 파일 · 탭 닫기
check(app.open_path(DATA / "없는파일.xlsx") is None, "없는 파일은 탭을 만들지 않음")
count = len(app.panes)
check(app.close_pane(pane, ask=False) and len(app.panes) == count - 1 and pane.closed, "탭 닫기")
check(app.close_pane(first, ask=False) and first.session is None, "고친 탭도 닫으면 데이터를 놓아 줌")
for p in list(app.panes):
    app.close_pane(p, ask=False)
root.update()
check(app.current is None and root.title() == ed.TITLE and app.hint.winfo_ismapped(), "모든 탭을 닫으면 처음 화면")

root.destroy()
shutil.rmtree(OUT, ignore_errors=True)
print("RESULT", "FAIL" if failed else "PASS")
sys.exit(1 if failed else 0)
