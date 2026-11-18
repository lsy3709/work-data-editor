# -*- coding: utf-8 -*-
"""
회귀 분석 · 통계 검정 창 — 편집기의 탭 하나에서 [회귀 분석] 을 누르면 뜬다.

왼쪽에서 ① 변수 → ② 정제 → ③ 이상 탐지 → ④ 모델 순서로 고르고 [분석 실행],
또는 [자동 분석] 으로 한 번에. 오른쪽에 ⑤ 결과가 탭으로 나온다:
  결과 요약 · 그래프 · 통계 검정 · 검정 그래프 · 학습 과정 · 모델 비교 · 이상 행 · 모델 소개
계산은 analysis.py · stat_tests.py 가 작업 스레드에서 하고, 이 파일은 화면만 맡는다.
도는 동안에는 진행 막대와 [학습 과정] 탭의 기록이 실시간으로 채워진다.
"""
import gc
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import analysis as an
import stat_tests as st
from data_editor import FONT, TITLE, to_text

FILL = {"중앙값으로 채우기": "median", "평균으로 채우기": "mean", "빈칸이 있는 행 빼기": "drop"}
OUTLIER_ORDER = ["combo", "model", "iqr", "zscore", "iforest", "none"]


class AnalysisWindow:
    def __init__(self, pane):
        self.pane = pane
        self.app = pane.app
        self.df = an.to_frame(pane.session)
        self.info = {p["name"]: p for p in an.profile(self.df)}
        self.clean = self.outliers = self.result = self.table = None
        self.tests = self.ols = self.stats = self.curve = None
        self.log = []
        self.busy = False
        self._events = []              # 작업 스레드가 쌓는 진행 기록 (화면 스레드가 꺼내서 보여 준다)
        self._rows = []                # 작업 스레드가 쌓는 모델 비교 결과 (끝나는 대로 한 줄씩)
        self._fraction = 0.0
        self._started = 0.0
        self.canvases = {}
        self.figures = {}
        self.win = tk.Toplevel(pane.root)
        self.win.title(f"회귀 분석 · 통계 검정 — {pane.source.path.name} [{pane.table}]")
        self.win.geometry("1320x840")
        self.win.minsize(1040, 660)
        self.win.protocol("WM_DELETE_WINDOW", self.close)
        self._build()
        self.use_suggestion()
        self.show_model_info()

    # ---------------- 화면 만들기 ----------------
    def _build(self):
        left = ttk.Frame(self.win, padding=(10, 8, 6, 8))
        left.pack(side="left", fill="y")

        box = ttk.LabelFrame(left, text=" ① 변수 지정 ", padding=8)
        box.pack(fill="x")
        ttk.Label(box, text="예측할 열 (종속변수 · 숫자)").pack(anchor="w")
        numeric = [name for name, p in self.info.items() if p["kind"] == "숫자"]
        self.target_box = ttk.Combobox(box, state="readonly", values=numeric, width=30, font=FONT)
        self.target_box.pack(fill="x", pady=(2, 6))
        ttk.Label(box, text="설명하는 열 (독립변수 · 여러 개, 클릭해서 고르기)").pack(anchor="w")
        self.feature_list = tk.Listbox(box, selectmode="multiple", exportselection=False, height=9, font=FONT,
                                       activestyle="none")
        self.feature_names = list(self.df.columns)
        for name in self.feature_names:
            p = self.info[name]
            extra = f"빈칸 {p['missing']}" if p["missing"] else ""
            self.feature_list.insert("end", f"{name}   ({p['kind']}{' · ' + extra if extra else ''})")
        self.feature_list.pack(fill="x", pady=2)
        ttk.Button(box, text="추천대로 고르기", command=self.use_suggestion).pack(anchor="e", pady=(2, 0))

        box = ttk.LabelFrame(left, text=" ② 데이터 정제 ", padding=8)
        box.pack(fill="x", pady=(8, 0))
        ttk.Label(box, text="빈칸 처리 (글자로 적힌 숫자 · 날짜는 자동으로 바로잡음)", wraplength=300).pack(anchor="w")
        self.fill_box = ttk.Combobox(box, state="readonly", values=list(FILL), width=30, font=FONT)
        self.fill_box.set("중앙값으로 채우기")
        self.fill_box.pack(fill="x", pady=2)
        self.dedup = tk.BooleanVar(value=True)
        ttk.Checkbutton(box, text="모든 칸이 똑같은 겹친 행 빼기", variable=self.dedup).pack(anchor="w")

        box = ttk.LabelFrame(left, text=" ③ 이상 탐지 (찾은 행은 학습 · 검정에서 뺌) ", padding=8)
        box.pack(fill="x", pady=(8, 0))
        self.outlier_names = {an.OUTLIER_METHODS[k][0]: k for k in OUTLIER_ORDER}
        self.outlier_box = ttk.Combobox(box, state="readonly", values=list(self.outlier_names), width=30, font=FONT)
        self.outlier_box.set(an.OUTLIER_METHODS["combo"][0])
        self.outlier_box.pack(fill="x", pady=2)
        self.outlier_box.bind("<<ComboboxSelected>>", lambda e: self.show_outlier_info())
        self.outlier_help = ttk.Label(box, wraplength=300, foreground="#555", justify="left")
        self.outlier_help.pack(anchor="w")
        self.show_outlier_info()

        box = ttk.LabelFrame(left, text=" ④ 모델 선택 ", padding=8)
        box.pack(fill="x", pady=(8, 0))
        self.model_names = {f"{m.name}  [{m.family}]": key for key, m in an.MODELS.items()}
        self.model_box = ttk.Combobox(box, state="readonly", values=list(self.model_names), width=30, font=FONT)
        self.model_box.current(list(an.MODELS).index("boosting"))
        self.model_box.pack(fill="x", pady=2)
        self.model_box.bind("<<ComboboxSelected>>", lambda e: self.show_model_info())

        run = ttk.Frame(left)
        run.pack(fill="x", pady=(12, 0))
        ttk.Button(run, text="분석 실행", command=self.run).pack(side="left", fill="x", expand=True)
        ttk.Button(run, text="자동 분석 (9개 모델 비교)", command=self.run_auto).pack(side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Button(left, text="결과 내보내기 (엑셀 + 그림)", command=self.export_dialog).pack(fill="x", pady=(6, 0))
        self.progress = ttk.Progressbar(left, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(10, 0))
        self.status = ttk.Label(left, text="① ~ ④ 를 고르고 [분석 실행], 잘 모르겠으면 [자동 분석]",
                                wraplength=310, foreground="#B45309", justify="left")
        self.status.pack(anchor="w", pady=(6, 0))

        self.tabs = ttk.Notebook(self.win)
        self.tabs.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=8)
        self.summary = self._text_tab("⑤ 결과 요약")
        self.graph_frame = self._figure_tab("그래프", "분석을 실행하면 그래프 4장이 나옵니다")

        frame = ttk.Frame(self.tabs)                                  # ---- 통계 검정 ----
        self.tabs.add(frame, text="통계 검정")
        self.test_head = ttk.Label(frame, padding=(6, 6), foreground="#555", wraplength=900, justify="left",
                                   text="분석을 실행하면 설명하는 열마다 예측할 열과 관계가 있는지 검정합니다. "
                                        "숫자 열 = 상관 t 검정, 두 집단 = 독립표본 t 검정, 세 집단 이상 = 분산분석 + 사후 t 검정.")
        self.test_head.pack(anchor="w")
        cols = ("설명하는 열", "종류", "검정 방법", "통계량", "자유도", "p 값", "효과 크기", "판정")
        self.test_tree = ttk.Treeview(frame, columns=cols, show="headings", height=7)
        for col, width in zip(cols, (130, 50, 220, 110, 90, 90, 150, 90)):
            self.test_tree.heading(col, text=col)
            self.test_tree.column(col, width=width, anchor="w" if col in cols[:3] or col == "판정" else "e", stretch=False)
        self.test_tree.tag_configure("yes", background="#FEE2E2")
        self.test_tree.pack(fill="x", padx=6)
        self.test_tree.bind("<<TreeviewSelect>>", lambda e: self.show_test_detail())
        self.test_text = tk.Text(frame, font=FONT, wrap="word", height=7, padx=10, pady=6, spacing3=3, state="disabled")
        self.test_text.pack(fill="both", expand=True, padx=6, pady=6)
        self.ols_head = ttk.Label(frame, padding=(6, 0), foreground="#555", wraplength=900, justify="left",
                                  text="다중 회귀의 계수 t 검정 — 모든 설명 열을 함께 넣었을 때, 다른 열이 같아도 이 열이 유의한가 (t = 계수 ÷ 표준오차)")
        self.ols_head.pack(anchor="w")
        cols = ("열", "계수", "표준오차", "t", "p 값", "95% 신뢰구간", "판정")
        self.ols_tree = ttk.Treeview(frame, columns=cols, show="headings", height=8)
        for col, width in zip(cols, (200, 120, 110, 80, 90, 230, 90)):
            self.ols_tree.heading(col, text=col)
            self.ols_tree.column(col, width=width, anchor="w" if col in ("열", "판정") else "e", stretch=False)
        self.ols_tree.tag_configure("yes", background="#FEE2E2")
        self.ols_tree.pack(fill="x", padx=6, pady=(2, 6))

        self.test_graph_frame = self._figure_tab("검정 그래프", "분석을 실행하면 검정 결과 그래프가 나옵니다")

        frame = ttk.Frame(self.tabs)                                  # ---- 학습 과정 ----
        self.tabs.add(frame, text="학습 과정")
        ttk.Label(frame, padding=(6, 6), foreground="#555",
                  text="진행 기록 (도는 동안 실시간으로 채워집니다) — 걸린 시간 · 한 일").pack(anchor="w")
        self.log_text = tk.Text(frame, font=FONT, wrap="none", height=10, padx=10, pady=6, state="disabled")
        self.log_text.pack(fill="x", padx=6)
        self.learn_frame = ttk.Frame(frame)
        self.learn_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.learn_hint = ttk.Label(self.learn_frame, foreground="#666", font=(FONT[0], 12),
                                    text="분석이 끝나면 학습 곡선이 나옵니다")
        self.learn_hint.place(relx=0.5, rely=0.45, anchor="center")

        frame = ttk.Frame(self.tabs)                                  # ---- 모델 비교 ----
        self.tabs.add(frame, text="모델 비교")
        ttk.Label(frame, padding=(6, 6), foreground="#555", wraplength=900, justify="left",
                  text="[자동 분석] 을 하면 모델이 하나 끝날 때마다 한 줄씩 채워집니다. 줄을 더블클릭하면 그 모델로 다시 분석합니다. "
                       "R² 는 데이터를 5조각으로 나눠 돌아가며 시험한 평균입니다.").pack(anchor="w")
        cols = ("순위", "모델", "계열", "교차 검증 R²", "흔들림(±)", "RMSE", "시간(초)")
        self.compare_tree = ttk.Treeview(frame, columns=cols, show="headings", height=12)
        for col, width in zip(cols, (50, 190, 80, 110, 90, 110, 80)):
            self.compare_tree.heading(col, text=col)
            self.compare_tree.column(col, width=width, anchor="w" if col in ("모델", "계열") else "e", stretch=False)
        self.compare_tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.compare_tree.bind("<Double-1>", lambda e: self.use_compared_model())

        frame = ttk.Frame(self.tabs)                                  # ---- 이상 행 ----
        self.tabs.add(frame, text="이상 행")
        bar = ttk.Frame(frame, padding=(6, 6))
        bar.pack(fill="x")
        self.outlier_count = ttk.Label(bar, text="분석을 실행하면 이상하다고 본 행과 이유가 나옵니다", foreground="#555")
        self.outlier_count.pack(side="left")
        ttk.Button(bar, text="편집기에서 이 행들만 보기", command=self.show_outliers_in_editor).pack(side="right")
        cols = ("행", "실제 값", "예측 값", "이유")
        self.outlier_tree = ttk.Treeview(frame, columns=cols, show="headings")
        for col, width in zip(cols, (150, 110, 110, 620)):
            self.outlier_tree.heading(col, text=col)
            self.outlier_tree.column(col, width=width, anchor="e" if "값" in col else "w", stretch=(col == "이유"))
        self.outlier_tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.model_text = self._text_tab("모델 소개")

    def _text_tab(self, title):
        frame = ttk.Frame(self.tabs)
        self.tabs.add(frame, text=title)
        text = tk.Text(frame, font=FONT, wrap="word", padx=12, pady=10, spacing1=2, spacing3=4, state="disabled")
        scroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        text.tag_configure("head", font=(FONT[0], 12, "bold"), spacing3=8)
        text.tag_configure("label", font=(FONT[0], 10, "bold"), foreground="#1D4ED8")
        return text

    def _figure_tab(self, title, hint):
        frame = ttk.Frame(self.tabs)
        self.tabs.add(frame, text=title)
        label = ttk.Label(frame, text=hint, foreground="#666", font=(FONT[0], 12))
        label.place(relx=0.5, rely=0.45, anchor="center")
        frame.hint = label
        return frame

    def _show_figure(self, name, frame, fig):
        """그림(matplotlib Figure)을 그 자리에 올린다. 전에 있던 그림은 치운다."""
        import matplotlib
        matplotlib.use("TkAgg", force=False)
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        if name in self.canvases:
            self.canvases.pop(name).get_tk_widget().destroy()
            self.figures.pop(name, None)
            gc.collect()        # 옛 그림(Tk 이미지)을 지금, 화면 스레드에서 치운다 — 작업 스레드에서 치워지면 안 된다
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)
        canvas.draw()
        self.canvases[name], self.figures[name] = canvas, fig

    @staticmethod
    def _write(text, parts):
        """parts = [(글, 꾸밈 이름 또는 None), …] 을 Text 칸에 다시 쓴다."""
        text.configure(state="normal")
        text.delete("1.0", "end")
        for chunk, tag in parts:
            text.insert("end", chunk, tag or ())
        text.configure(state="disabled")

    @property
    def canvas(self):                                    # 예전 이름 (그래프 탭의 그림)
        return self.canvases.get("main")

    @property
    def fig(self):
        return self.figures.get("main")

    # ---------------- ① ~ ④ 고르기 ----------------
    def use_suggestion(self):
        target, features = an.suggest(self.df)
        self.set_roles(target, features)

    def set_roles(self, target, features):
        """예측할 열과 설명하는 열을 고른다 (화면의 선택도 맞춘다)."""
        self.target_box.set(target or "")
        self.feature_list.selection_clear(0, "end")
        for name in features:
            self.feature_list.selection_set(self.feature_names.index(name))

    def roles(self):
        return self.target_box.get(), [self.feature_names[i] for i in self.feature_list.curselection()]

    def set_model(self, key):
        self.model_box.current(list(an.MODELS).index(key))
        self.show_model_info()

    def set_outlier_method(self, key):
        self.outlier_box.set(an.OUTLIER_METHODS[key][0])
        self.show_outlier_info()

    def model_key(self):
        return self.model_names[self.model_box.get()]

    def show_outlier_info(self):
        self.outlier_help["text"] = an.OUTLIER_METHODS[self.outlier_names[self.outlier_box.get()]][1]

    def show_model_info(self):
        m = an.MODELS[self.model_key()]
        self._write(self.model_text, [
            (f"{m.name}  ·  {m.english}\n", "head"), (f"계열: {m.family}\n\n", None),
            ("이름의 뜻 (어원)\n", "label"), (m.origin + "\n\n", None),
            ("역사\n", "label"), (m.history + "\n\n", None),
            ("하는 일\n", "label"), (m.what + "\n\n", None),
            ("언제 쓰나\n", "label"), (m.when + "\n\n", None),
            ("다른 모델은 왼쪽 ④ 에서 바꿔 보세요. 전체 안내는 docs/ml_guide.md 에 있습니다.", None)])

    # ---------------- 실행 (계산은 작업 스레드에서) ----------------
    def _prepare(self):
        """실행해도 되는지 보고, 편집기에서 고친 내용까지 반영해 표를 다시 읽는다."""
        if self.busy:
            self.win.bell()
            return False
        if self.pane.closed or self.pane.session is None:
            self.error("편집기에서 이 파일의 탭이 닫혔습니다. 분석 창을 닫고 다시 여세요.")
            return False
        self.df = an.to_frame(self.pane.session)
        return True

    def _start(self, work, done, first):
        self.busy = True
        self._events, self._rows, self._fraction = [], [], 0.0
        self._started = time.perf_counter()
        self._shown_events = self._shown_rows = 0
        self._write(self.log_text, [])
        self.compare_tree.delete(*self.compare_tree.get_children())
        self._say(first, 0.01)
        self._tick()

        def finish(result, error):
            self.busy = False
            if not self.win.winfo_exists():
                return
            if error:
                self._say("멈춤: " + str(error).splitlines()[0])
                self._flush()
                self.status["text"] = "분석하지 못했습니다."
                self.progress["value"] = 0
                self.error(str(error))
                return
            self._say("끝", 1.0)
            self._flush()
            done(result)

        self.app.run_job(work, finish)

    def _say(self, text, fraction=None):                 # ← 작업 스레드에서도 불린다 (목록에 적어 두기만 한다)
        self._events.append((time.perf_counter() - self._started, text))
        if fraction is not None:
            self._fraction = fraction

    def _flush(self):
        """작업 스레드가 적어 둔 진행 기록 · 모델 결과를 화면에 옮긴다 (화면 스레드에서만)."""
        events, rows = self._events, self._rows
        if self._shown_events < len(events):
            self.log_text.configure(state="normal")
            for seconds, text in events[self._shown_events:]:
                self.log_text.insert("end", f"{seconds:7.2f}초   {text}\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
            self._shown_events = len(events)
            self.status["text"] = events[-1][1] + (" …" if self.busy else "")
        for row in rows[self._shown_rows:]:
            self.compare_tree.insert("", "end", iid=row["key"], values=(
                "…", row["name"], row["family"], f"{row['r2']:.3f}", f"{row['std']:.3f}", f"{row['rmse']:,.0f}", f"{row['seconds']:.1f}"))
        self._shown_rows = len(rows)
        self.progress["value"] = self._fraction * 100

    def _tick(self):
        if self.win.winfo_exists():
            self._flush()
            if self.busy:
                self.win.after(80, self._tick)

    def _extras(self, c, o, result):
        """학습 뒤에 덧붙이는 계산: 통계 검정 · 기술통계 · 학습 곡선 (작업 스레드에서)."""
        self._say("설명하는 열마다 통계 검정 (t 검정 · 분산분석)", 0.86)
        tests = st.test_all(c, exclude=o.mask)
        self._say("다중 회귀의 계수 t 검정", 0.89)
        try:
            fit = st.ols(c, exclude=o.mask)
        except ValueError:
            fit = None
        stats = st.describe(c, exclude=o.mask)
        self._say(f"학습 곡선 그리는 중 · {result.name}", 0.92)
        curve = an.learning(c, result.key, exclude=o.mask, model=result.model)
        return tests, fit, stats, curve

    def run(self, keep_table=False):
        """고른 설정대로 정제 → 이상 탐지 → 학습 → 통계 검정. 시작했으면 True. (keep_table: 모델 비교표는 그대로 둔다)"""
        if not self._prepare():
            return False
        target, features = self.roles()
        missing, dedup = FILL[self.fill_box.get()], self.dedup.get()
        method, key, df = self.outlier_names[self.outlier_box.get()], self.model_key(), self.df
        kept = self.table if keep_table else None

        def work():
            self._say("데이터를 정제하는 중", 0.05)
            c = an.clean(df, target, features, missing=missing, drop_duplicates=dedup)
            for line in c.report:
                self._say("   " + line)
            self._say(f"이상한 행을 찾는 중 · {an.OUTLIER_METHODS[method][0]}", 0.2)
            o = an.find_outliers(c, method)
            self._say("   " + o.report[0])
            self._say(f"{an.MODELS[key].name} 학습 중", 0.45)
            result = an.train(c, key, exclude=o.mask)
            self._say(f"   시험 R² {result.metrics['r2_test']:.3f} · 평균 오차율 {result.metrics['mape']:.1f}%")
            return c, o, result, self._extras(c, o, result)

        def done(out):
            self.clean, self.outliers, self.result, extras = out
            self.tests, self.ols, self.stats, self.curve = extras
            self.table, self.log = kept, []
            self.show_result()

        self._start(work, done, "준비 중")
        return True

    def run_auto(self):
        """자동 분석: 변수는 화면에서 고른 것을 쓰고(없으면 추천), 나머지는 규칙대로 한 번에."""
        if not self._prepare():
            return False
        target, features = self.roles()
        df = self.df

        def work():
            a = an.auto(df, target or None, features or None,
                        progress=lambda text, fraction=None: self._say(text, None if fraction is None else fraction * 0.85),
                        on_row=lambda row: (self._rows.append(row),
                                            self._say(f"   {row['name']}: R² {row['r2']:.3f} ±{row['std']:.3f} ({row['seconds']:.1f}초)")))
            return a, self._extras(a.clean, a.outliers, a.best)

        def done(out):
            a, extras = out
            self.clean, self.outliers, self.result, self.table, self.log = a.clean, a.outliers, a.best, a.table, a.log
            self.tests, self.ols, self.stats, self.curve = extras
            self.set_roles(a.clean.target, a.clean.features)
            self.fill_box.set("중앙값으로 채우기")
            self.set_outlier_method(a.outliers.method)
            self.set_model(a.best.key)
            self.show_result()

        self._start(work, done, "자동 분석을 시작하는 중")
        return True

    def use_compared_model(self):
        picked = self.compare_tree.selection()
        if picked and self.table and not self.busy:
            self.set_model(picked[0])
            self.run(keep_table=True)                    # 비교표는 그대로 두고 그 모델로 다시 학습

    # ---------------- ⑤ 결과 보여 주기 ----------------
    def show_result(self):
        c, o, r = self.clean, self.outliers, self.result
        parts = []
        if self.log:
            parts += [("자동 분석이 한 일\n", "head")] + [(line + "\n", None) for line in self.log] + [("\n", None)]
        lines = an.summary_lines(c, o, r, self.table)
        parts += [(lines[0] + "\n", "head")] + [(line + "\n", None) for line in lines[1:]]
        if self.tests:
            yes = [t.feature for t in self.tests if t.significant]
            parts += [("\n통계 검정 (자세한 내용은 [통계 검정] 탭)\n", "label"),
                      (f"  설명하는 열 {len(self.tests)}개 중 {len(yes)}개가 통계적으로 유의" + (": " + ", ".join(yes) if yes else "") + "\n", None)]
        self._write(self.summary, parts)

        self.graph_frame.hint.place_forget()
        self._show_figure("main", self.graph_frame, an.figure(c, o, r, self.table))
        self.show_tests()
        self.learn_hint.place_forget()
        self._show_figure("learn", self.learn_frame, an.figure_learning(r, self.curve))

        self.compare_tree.delete(*self.compare_tree.get_children())
        for i, row in enumerate(self.table or [], 1):
            self.compare_tree.insert("", "end", iid=row["key"], values=(
                i, row["name"], row["family"], f"{row['r2']:.3f}", f"{row['std']:.3f}", f"{row['rmse']:,.0f}", f"{row['seconds']:.1f}"))
        if self.table and r.key in self.compare_tree.get_children():
            self.compare_tree.selection_set(r.key)

        self.outlier_tree.delete(*self.outlier_tree.get_children())
        first = self.df.columns[0]
        rids = [rid for rid in o.mask.index[o.mask]]
        for rid in rids:
            self.outlier_tree.insert("", "end", iid=str(rid), values=(
                to_text(self.df.at[rid, first]), f"{c.y[rid]:,.0f}", f"{r.predictions[rid]:,.0f}", o.reasons[rid]))
        self.outlier_count["text"] = f"{len(c.X):,}행 중 {len(rids):,}행을 이상하다고 보고 학습에서 뺐습니다 — 지우지는 않습니다"
        m = r.metrics
        self.status["text"] = (f"끝: {r.name} · 시험 R² {m['r2_test']:.3f} ({an.grade(m['r2_test'])}) · "
                               f"평균 오차율 {m['mape']:.1f}% · 이상 {len(rids)}행 · {time.perf_counter() - self._started:.1f}초")
        self.progress["value"] = 100
        self.tabs.select(0)

    def show_tests(self):
        """[통계 검정] · [검정 그래프] 탭을 채운다."""
        c, tests, fit = self.clean, self.tests or [], self.ols
        yes = [t for t in tests if t.significant]
        self.test_head["text"] = (f"예측할 열 [{c.target}] 과(와) 설명하는 열 {len(tests)}개의 관계 — {len(yes)}개가 통계적으로 유의 (p < {st.ALPHA}). "
                                  "분홍 줄 = 유의. 줄을 누르면 아래에 풀이와 집단별 평균이 나옵니다. 이상 탐지에서 뺀 행은 검정에서도 뺍니다.")
        self.test_tree.delete(*self.test_tree.get_children())
        for i, t in enumerate(tests):
            self.test_tree.insert("", "end", iid=str(i), tags=("yes",) if t.significant else (), values=(
                t.feature, t.kind, t.method, f"{t.stat_name} = {t.stat:,.3f}", t.df, st.p_short(t.p),
                st.effect_text(t.effect_name, t.effect), "유의함" if t.significant else "유의하지 않음"))
        self.show_test_detail()
        self.ols_tree.delete(*self.ols_tree.get_children())
        if fit:
            self.ols_head["text"] = (f"다중 회귀의 계수 t 검정 — {fit.n:,}행 · R² {fit.r2:.3f} (조정 R² {fit.adj_r2:.3f}) · 모델 전체 F = {fit.f:,.1f}, "
                                     f"{st.p_text(fit.f_p)}.  다른 열이 같을 때 이 열이 1 늘면 [{c.target}] 이 '계수'만큼 달라집니다 (t = 계수 ÷ 표준오차).  {fit.note}")
            for i, row in enumerate(fit.rows):
                self.ols_tree.insert("", "end", iid=str(i), tags=("yes",) if row["significant"] else (), values=(
                    row["name"], f"{row['coef']:,.2f}", f"{row['se']:,.2f}", f"{row['t']:.2f}", st.p_short(row["p"]),
                    f"{row['low']:,.1f} ~ {row['high']:,.1f}", "유의함" if row["significant"] else "유의하지 않음"))
        self.test_graph_frame.hint.place_forget()
        self._show_figure("tests", self.test_graph_frame, st.figure_tests(c, tests, fit, exclude=self.outliers.mask))

    def show_test_detail(self):
        tests = self.tests or []
        picked = self.test_tree.selection()
        chosen = [tests[int(picked[0])]] if picked else tests
        parts = []
        for t in chosen:
            parts.append((("● " if t.significant else "○ ") + t.sentence + "\n", None))
            if picked:
                for g in t.groups:
                    parts.append((f"      {g['name']}:  {g['n']:,}행 · 평균 {g['mean']:,.1f} · 표준편차 {g['std']:,.1f}\n", None))
                for q in t.pairs:
                    mark = "≠" if q["p_adj"] < st.ALPHA else "≈"
                    parts.append((f"      {q['a']} {mark} {q['b']}:  차이 {q['diff']:,.1f} · t = {q['t']:.2f} · 보정한 {st.p_text(q['p_adj'])}\n", None))
        self._write(self.test_text, parts)

    def show_outliers_in_editor(self):
        """이상하다고 본 행만 편집기 탭에 남겨 보여 준다 (거기서 바로 고칠 수 있다). 보여 준 행 수를 돌려준다."""
        if self.outliers is None or self.pane.closed or self.pane.session is None:
            return 0
        rids = [rid for rid in self.outliers.mask.index[self.outliers.mask] if rid in self.pane.session.rows]
        self.pane.view = rids
        self.pane.refresh()
        self.app.notebook.select(self.pane.frame)
        self.app.root.lift()
        return len(rids)

    # ---------------- 내보내기 ----------------
    def export_dialog(self):
        if self.result is None:
            messagebox.showinfo(TITLE, "먼저 분석을 실행하세요.", parent=self.win)
            return
        path = filedialog.asksaveasfilename(parent=self.win, title="분석 결과를 엑셀로 내보내기", defaultextension=".xlsx",
                                            filetypes=[("엑셀", "*.xlsx")], initialfile=f"{self.pane.table}_분석결과_내보내기.xlsx")
        if path:
            try:
                files = self.export(path)
            except Exception as e:
                self.error(str(e))
                return
            messagebox.showinfo(TITLE, "내보냈습니다.\n\n" + "\n".join(str(f) for f in files), parent=self.win)

    def export(self, path):
        """엑셀(예측결과 · 모델비교 · 요약 · 통계검정 · 회귀계수)과 그림 3장(.png)을 만든다. 만든 파일들을 돌려준다."""
        path = Path(path)
        an.export(path, self.df, self.clean, self.outliers, self.result, self.table, tests=self.tests, ols=self.ols)
        files = [path]
        for name, tail in (("main", ""), ("tests", "_검정"), ("learn", "_학습")):
            if name in self.figures:
                picture = path.with_name(path.stem + tail + ".png")
                self.figures[name].savefig(picture, dpi=130)
                files.append(picture)
        return files

    # ---------------- 그 밖 ----------------
    def error(self, text):
        messagebox.showwarning(TITLE, text, parent=self.win)

    def close(self):
        if self in self.pane.windows:
            self.pane.windows.remove(self)
        if self.win.winfo_exists():
            self.win.destroy()
        self.canvases.clear()
        self.figures.clear()
        gc.collect()            # 그림이 쥐고 있던 Tk 이미지를 화면 스레드에서 치운다
