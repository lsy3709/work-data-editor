"""
전처리 창 — 분석하기 전에 표를 한꺼번에 다듬는다.

왼쪽에서 열을 고르고(여러 개 가능), 오른쪽에서 할 일을 누른다.
  ① 빈칸      채우기(중앙값 · 평균 · 가장 흔한 값 · 직접 입력 · 윗칸 값) / 빈칸 있는 행 삭제
  ② 종류      글자 → 숫자 · 날짜 · 글자 · 예/아니오 로 한꺼번에 바꾸기 / 숫자에 곱하기 · 나누기
  ③ 글자 · 행  앞뒤 빈칸 지우기 / 겹친 행 삭제
  ④ 열        이름 바꾸기 · 삭제 · 추가

계산은 모두 data_editor.EditSession 이 한다. 이 파일은 화면만 맡는다.
무엇을 하든 원본 파일은 [저장]을 누르기 전까지 그대로이고, [되돌리기]로 한 단계씩 되돌아간다.
"""
import tkinter as tk
from tkinter import messagebox, ttk

from data_editor import FONT, KIND_NAME, TITLE, to_text

FILLS = [("중앙값", "median"), ("평균", "mean"), ("가장 흔한 값", "mode"), ("직접 입력", "value"), ("윗칸 값", "above")]
KINDS = [("숫자", "number"), ("날짜", "date"), ("글자", "str"), ("예/아니오", "bool")]
HELP = {
    "median": "크기순으로 줄 세웠을 때 한가운데 값. 유난히 큰 값이 섞여 있어도 덜 흔들립니다.",
    "mean": "모두 더해서 개수로 나눈 값. 유난히 큰 값이 있으면 그쪽으로 끌려갑니다.",
    "mode": "가장 자주 나온 값. 글자 열(지역 · 유형)에도 쓸 수 있습니다.",
    "value": "오른쪽 칸에 적은 값으로 채웁니다. (예: 0, 미확인)",
    "above": "바로 위 칸의 값을 이어받습니다. 엑셀에서 병합을 풀었을 때 생긴 빈칸에 씁니다.",
}


class PreprocessWindow:
    def __init__(self, pane):
        self.pane = pane
        self.app = pane.app
        self.quiet = False             # True 면 안내 창을 띄우지 않는다 (자동 점검용)
        self.last_error = ""
        self.log = []
        self.win = tk.Toplevel(pane.root)
        self.win.geometry("1340x720")
        self.win.minsize(980, 600)
        self.win.protocol("WM_DELETE_WINDOW", self.close)
        self._build()
        self.reload()

    @property
    def session(self):
        return self.pane.session

    # ---------------- 화면 만들기 ----------------
    def _build(self):
        bottom = ttk.Frame(self.win, padding=(10, 4, 10, 8))
        bottom.pack(side="bottom", fill="x")
        ttk.Button(bottom, text="닫기", command=self.close).pack(side="right")
        ttk.Button(bottom, text="저장", command=self.pane.save).pack(side="right", padx=6)
        ttk.Button(bottom, text="다른 이름으로 저장", command=self.pane.save_as).pack(side="right")
        ttk.Button(bottom, text="되돌리기", command=self.undo).pack(side="right", padx=6)
        self.message = ttk.Label(bottom, text="", foreground="#333")
        self.message.pack(side="left")

        right = ttk.Frame(self.win, padding=(6, 8, 10, 0))
        right.pack(side="right", fill="y")
        left = ttk.Frame(self.win, padding=(10, 8, 6, 0))
        left.pack(side="left", fill="both", expand=True)

        # ---- 왼쪽: 열 목록 ----
        self.summary = ttk.Label(left, text="", foreground="#111")
        self.summary.pack(anchor="w")
        ttk.Label(left, foreground="#666",
                  text="다듬을 열을 고르세요. Ctrl · Shift 를 누른 채 클릭하면 여러 열을 고를 수 있습니다.").pack(anchor="w", pady=(2, 4))
        box = ttk.Frame(left)
        box.pack(fill="both", expand=True)
        heads = [("name", "열", 140, "w"), ("kind", "종류", 60, "w"), ("empty", "빈칸", 50, "e"),
                 ("unique", "서로 다른 값", 84, "e"), ("stat", "최소 ~ 최대 · 중앙값", 170, "w"), ("example", "값 예시", 240, "w")]
        self.cols = ttk.Treeview(box, show="headings", selectmode="extended", columns=[h[0] for h in heads], height=12)
        for key, text, width, anchor in heads:
            self.cols.heading(key, text=text)
            self.cols.column(key, width=width, anchor=anchor, stretch=key == "example")
        ys = ttk.Scrollbar(box, orient="vertical", command=self.cols.yview)
        self.cols.configure(yscrollcommand=ys.set)
        ys.pack(side="right", fill="y")
        self.cols.pack(fill="both", expand=True)
        self.cols.tag_configure("empty", background="#FFF3C4")          # 빈칸이 있는 열
        self.cols.bind("<<TreeviewSelect>>", lambda e: self._on_select())
        row = ttk.Frame(left)
        row.pack(fill="x", pady=4)
        ttk.Button(row, text="빈칸 있는 열 고르기", command=self.select_empty).pack(side="left")
        ttk.Button(row, text="모두 고르기", command=lambda: self.select(self.session.names)).pack(side="left", padx=6)
        ttk.Button(row, text="선택 풀기", command=lambda: self.select([])).pack(side="left")
        self.picked = ttk.Label(row, text="", foreground="#666")
        self.picked.pack(side="left", padx=10)
        ttk.Label(left, text="작업 기록").pack(anchor="w", pady=(4, 2))
        self.log_box = tk.Text(left, height=8, font=FONT, state="disabled", wrap="word", background="#FAFAFA")
        self.log_box.pack(fill="x")

        # ---- 오른쪽: 할 일 ----
        f1 = ttk.LabelFrame(right, text=" ① 빈칸 ", padding=8)
        f1.pack(fill="x")
        r = ttk.Frame(f1)
        r.pack(fill="x")
        ttk.Label(r, text="채우는 값").pack(side="left")
        self.fill_how = ttk.Combobox(r, state="readonly", width=12, values=[t for t, _ in FILLS], font=FONT)
        self.fill_how.set("중앙값")
        self.fill_how.pack(side="left", padx=4)
        self.fill_how.bind("<<ComboboxSelected>>", lambda e: self._show_help())
        self.fill_value = ttk.Entry(r, width=10, font=FONT)
        self.fill_value.pack(side="left")
        ttk.Button(r, text="빈칸 채우기", command=self.fill).pack(side="left", padx=(6, 0))
        self.fill_help = ttk.Label(f1, text="", foreground="#666", wraplength=380, justify="left")
        self.fill_help.pack(anchor="w", pady=(4, 6))
        r = ttk.Frame(f1)
        r.pack(fill="x")
        ttk.Button(r, text="빈칸 있는 행 삭제", command=self.drop_empty).pack(side="left")
        ttk.Label(r, text="고른 열에 빈칸이 있는 행 (안 고르면 모든 열)", foreground="#666").pack(side="left", padx=6)

        f2 = ttk.LabelFrame(right, text=" ② 종류 바꾸기 · 숫자 계산 ", padding=8)
        f2.pack(fill="x", pady=(8, 0))
        r = ttk.Frame(f2)
        r.pack(fill="x")
        ttk.Label(r, text="바꿀 종류").pack(side="left")
        self.kind_box = ttk.Combobox(r, state="readonly", width=10, values=[t for t, _ in KINDS], font=FONT)
        self.kind_box.set("숫자")
        self.kind_box.pack(side="left", padx=4)
        ttk.Button(r, text="종류 바꾸기", command=self.convert).pack(side="left")
        self.to_empty = tk.BooleanVar(value=False)
        ttk.Checkbutton(f2, text="못 바꾸는 값은 빈칸으로 두고 계속 (끄면 아무것도 바꾸지 않고 알려 줌)",
                        variable=self.to_empty).pack(anchor="w", pady=(4, 0))
        ttk.Label(f2, foreground="#666", wraplength=380, justify="left",
                  text="'2,430,000원' · ' 1,140,000 ' · '12%' 같은 글자를 숫자로, '2026/07/15' · '20260613' 을 날짜로 읽습니다.").pack(anchor="w", pady=(2, 6))
        r = ttk.Frame(f2)
        r.pack(fill="x")
        ttk.Label(r, text="숫자 열에").pack(side="left")
        self.factor = ttk.Entry(r, width=10, font=FONT)
        self.factor.insert(0, "10000")
        self.factor.pack(side="left", padx=4)
        ttk.Button(r, text="곱하기", command=lambda: self.scale(False)).pack(side="left")
        ttk.Button(r, text="나누기", command=lambda: self.scale(True)).pack(side="left", padx=4)
        ttk.Label(r, text="(원 → 만원 = 10000 으로 나누기)", foreground="#666").pack(side="left")

        f3 = ttk.LabelFrame(right, text=" ③ 글자 · 행 ", padding=8)
        f3.pack(fill="x", pady=(8, 0))
        r = ttk.Frame(f3)
        r.pack(fill="x")
        ttk.Button(r, text="앞뒤 빈칸 지우기", command=self.strip).pack(side="left")
        ttk.Button(r, text="겹친 행 삭제", command=self.drop_duplicates).pack(side="left", padx=6)
        ttk.Label(f3, foreground="#666", wraplength=380, justify="left",
                  text="열을 고르면 그 열만 봅니다. 겹친 행은 먼저 나온 행을 남깁니다. (' 김하나' 와 '김하나' 는 빈칸부터 지워야 같아집니다)").pack(anchor="w", pady=(4, 0))

        f4 = ttk.LabelFrame(right, text=" ④ 열 ", padding=8)
        f4.pack(fill="x", pady=(8, 0))
        r = ttk.Frame(f4)
        r.pack(fill="x")
        ttk.Label(r, text="새 이름").pack(side="left")
        self.new_name = ttk.Entry(r, width=16, font=FONT)
        self.new_name.pack(side="left", padx=4)
        ttk.Button(r, text="이름 바꾸기", command=self.rename).pack(side="left")
        ttk.Button(r, text="고른 열 삭제", command=self.delete_columns).pack(side="left", padx=6)
        r = ttk.Frame(f4)
        r.pack(fill="x", pady=(6, 0))
        ttk.Label(r, text="새 열").pack(side="left")
        self.add_name = ttk.Entry(r, width=12, font=FONT)
        self.add_name.pack(side="left", padx=4)
        self.add_kind = ttk.Combobox(r, state="readonly", width=8, values=["글자", "정수", "실수", "날짜"], font=FONT)
        self.add_kind.set("글자")
        self.add_kind.pack(side="left")
        ttk.Label(r, text="값").pack(side="left", padx=(6, 2))
        self.add_value = ttk.Entry(r, width=8, font=FONT)
        self.add_value.pack(side="left")
        ttk.Button(r, text="열 추가", command=self.add_column).pack(side="left", padx=(6, 0))
        ttk.Label(f4, foreground="#666", wraplength=380, justify="left",
                  text="열을 하나 골라 두면 그 열 바로 뒤에 넣습니다. 열 구조를 바꾼 Access 표는 [다른 이름으로 저장]으로 엑셀에 저장합니다.").pack(anchor="w", pady=(4, 0))
        self._show_help()

    # ---------------- 보여 주기 ----------------
    def reload(self):
        """표가 바뀌었을 때 열 목록과 요약을 다시 그린다 (고른 열은 이름으로 기억해 둔다)."""
        if not self.win.winfo_exists() or self.session is None:
            return
        s = self.session
        self.win.title(f"전처리 — {self.pane.source.path.name} [{self.pane.table}]")
        chosen = set(self.selected())
        self.cols.delete(*self.cols.get_children())
        total_empty = 0
        for i, name in enumerate(s.names):
            info = s.column_summary(name)
            total_empty += info["empty"]
            stat = ""
            if "min" in info:
                stat = f"{to_text(info['min'])} ~ {to_text(info['max'])} · {_num(info['median'])}"
            self.cols.insert("", "end", iid=f"n{i}", tags=("empty",) if info["empty"] else (),
                             values=(name, KIND_NAME[info["kind"]], f"{info['empty']:,}" if info["empty"] else "",
                                     f"{info['unique']:,}", stat, "  ·  ".join(info["examples"])))
        self.select([n for n in s.names if n in chosen])
        self.summary["text"] = (f"행 {len(s.rows):,}개 · 열 {len(s.names)}개 · 빈칸 {total_empty:,}칸"
                                f" · 아직 저장하지 않은 변경 {s.change_count():,}건")

    def selected(self):
        names = self.session.names if self.session else []
        return [names[int(iid[1:])] for iid in self.cols.selection() if int(iid[1:]) < len(names)]

    def select(self, names):
        wanted = set(names)
        self.cols.selection_set([f"n{i}" for i, n in enumerate(self.session.names) if n in wanted])
        self._on_select()

    def select_empty(self):
        self.select([n for n in self.session.names if self.session.empty_count(n)])

    def _on_select(self):
        picked = self.selected()
        self.picked["text"] = f"고른 열 {len(picked)}개" if picked else "고른 열 없음"
        self.new_name.delete(0, "end")
        if len(picked) == 1:
            self.new_name.insert(0, picked[0])

    def _show_help(self):
        how = dict(FILLS)[self.fill_how.get()]
        self.fill_help["text"] = HELP[how]
        self.fill_value.configure(state="normal" if how == "value" else "disabled")

    def _say(self, text, error=False):
        self.last_error = text if error else ""
        self.message.configure(text=text.split("\n")[0], foreground="#B00020" if error else "#1B5E20")
        self.log.append(("✗ " if error else "✓ ") + text.replace("\n", " "))
        self.log_box.configure(state="normal")
        self.log_box.insert("end", self.log[-1] + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        if error and not self.quiet:
            messagebox.showwarning(TITLE, text, parent=self.win)

    # ---------------- 실행 ----------------
    def _run(self, label, work, need=0):
        """work(고른 열들) 을 실행하고 표 화면과 이 창을 다시 그린다. work 는 결과 문장을 돌려준다."""
        self.pane.commit_edit()
        if not self.pane.ready():
            return False
        names = self.selected()
        if need and len(names) < need:
            self._say(f"{label}: 왼쪽에서 열을 먼저 고르세요.", error=True)
            return False
        if need == 1 and label in ("이름 바꾸기",) and len(names) != 1:
            self._say(f"{label}: 열을 하나만 고르세요.", error=True)
            return False
        done = len(self.session._undo)
        try:
            text = work(names)
        except ValueError as e:
            while len(self.session._undo) > done:        # 여러 열 중 하나에서 멈췄으면 앞서 한 것도 되돌린다
                self.session.undo()
            self.pane.after_preprocess()
            self.reload()
            self._say(f"{label}: {e}", error=True)
            return False
        self._merge_undo(done)
        self.pane.after_preprocess()
        self.reload()
        self._say(f"{label} — {text}")
        return True

    def _merge_undo(self, done):
        """여러 열에 한 번에 한 일은 [되돌리기] 한 번으로 되돌아가게 기록을 하나로 묶는다."""
        undo = self.session._undo
        extra = len(undo) - done
        if extra > 1:
            steps = [undo.pop() for _ in range(extra)][::-1]
            undo.append([op for step in steps for op in step])

    def fill(self, how=None, value=None):
        how = how or dict(FILLS)[self.fill_how.get()]
        value = self.fill_value.get() if value is None else value

        def work(names):
            n = sum(self.session.fill_empty(name, how, value) for name in names)
            return f"{_names(names)} 의 빈칸 {n:,}칸을 {dict((k, t) for t, k in FILLS)[how]}(으)로 채웠습니다."
        return self._run("빈칸 채우기", work, need=1)

    def drop_empty(self):
        def work(names):
            n = self.session.drop_empty_rows(names or None)
            return f"{_names(names) if names else '모든 열'} 에 빈칸이 있는 행 {n:,}개를 지웠습니다."
        return self._run("빈칸 있는 행 삭제", work)

    def convert(self, kind=None, to_empty=None):
        kind = kind or dict(KINDS)[self.kind_box.get()]
        on_error = "empty" if (self.to_empty.get() if to_empty is None else to_empty) else "stop"

        def work(names):
            changed = failed = 0
            for name in names:
                c, f = self.session.convert(name, kind, on_error=on_error)
                changed, failed = changed + c, failed + f
            tail = f" (못 바꾼 {failed:,}칸은 빈칸)" if failed else ""
            return f"{_names(names)} 을(를) {KIND_NAME[kind]}(으)로 바꿨습니다. 값이 바뀐 칸 {changed:,}개{tail}"
        return self._run("종류 바꾸기", work, need=1)

    def scale(self, divide=False, factor=None):
        def work(names):
            text = (self.factor.get() if factor is None else str(factor)).replace(",", "").strip()
            try:
                number = float(text)
            except ValueError:
                raise ValueError(f"'{text}' 은(는) 숫자가 아닙니다.") from None
            if number == 0:
                raise ValueError("0 으로는 곱하거나 나눌 수 없습니다.")
            n = sum(self.session.scale(name, number, divide=divide) for name in names)
            return f"{_names(names)} 에 {text} 을(를) {'나눴' if divide else '곱했'}습니다. 바뀐 칸 {n:,}개"
        return self._run("숫자 계산", work, need=1)

    def strip(self):
        def work(names):
            n = self.session.strip_text(names or None)
            return f"{_names(names) if names else '모든 열'} 에서 앞뒤 빈칸이 있던 {n:,}칸을 다듬었습니다."
        return self._run("앞뒤 빈칸 지우기", work)

    def drop_duplicates(self):
        def work(names):
            n = self.session.drop_duplicates(names or None)
            return f"{_names(names) + ' 이(가) 같은' if names else '모든 값이 같은'} 겹친 행 {n:,}개를 지웠습니다."
        return self._run("겹친 행 삭제", work)

    def rename(self, new=None):
        def work(names):
            target = self.new_name.get() if new is None else new
            self.session.rename_column(names[0], target)
            return f"[{names[0]}] → [{target.strip()}]"
        return self._run("이름 바꾸기", work, need=1)

    def delete_columns(self):
        def work(names):
            n = self.session.delete_columns(names)
            return f"{_names(names)} 열 {n}개를 지웠습니다."
        return self._run("열 삭제", work, need=1)

    def add_column(self, name=None, kind=None, value=None):
        kinds = {"글자": "str", "정수": "int", "실수": "float", "날짜": "date"}

        def work(names):
            new = self.add_name.get() if name is None else name
            k = kind or kinds[self.add_kind.get()]
            v = self.add_value.get() if value is None else value
            self.session.add_column(new, kind=k, value=v, after=names[-1] if names else None)
            return f"[{new.strip()}] 열을 {('[' + names[-1] + '] 뒤') if names else '맨 끝'}에 넣었습니다."
        return self._run("열 추가", work)

    def undo(self):
        self.pane.commit_edit()
        if not self.pane.ready():
            return False
        if not self.session.undo():
            self._say("되돌릴 작업이 없습니다.", error=False)
            return False
        self.pane.after_preprocess()
        self.reload()
        self._say("마지막 작업을 되돌렸습니다.")
        return True

    def close(self):
        if self in self.pane.windows:
            self.pane.windows.remove(self)
        if self.win.winfo_exists():
            self.win.destroy()


def _names(names):
    shown = ", ".join(f"[{n}]" for n in names[:3])
    return shown + (f" 외 {len(names) - 3}개" if len(names) > 3 else "")


def _num(v):
    return to_text(int(v) if float(v).is_integer() else round(v, 2))
