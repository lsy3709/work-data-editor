"""
AI 채팅 창 — 우리말로 "이렇게 고쳐 줘"라고 하면 AI 가 명령을 만들고, 프로그램이 내 PC 에서 실행한다.

  · AI 에게 보내는 것: 열 이름 · 종류 · 빈칸 수와 내가 친 말. 칸의 값은 보내지 않는다 ('예시 값도 보내기'를 켰을 때만 앞 3행).
  · 표를 바꾸는 명령은 '적용할 내용'에 미리 보여 주고 [적용]을 눌러야 실행된다. 실행한 뒤에는 [되돌리기]로 돌아간다.
  · 보기 · 정렬 · 집계처럼 표를 바꾸지 않는 명령은 바로 실행한다. 집계 결과는 내 PC 에서 계산하고 AI 에게 보내지 않는다.
  · API 키는 이 창에 넣거나 환경 변수에서 읽는다. 프로그램을 끄면 사라지고 파일에 저장하지 않는다.

계산은 ai_commands.py, 통신은 ai_client.py 가 한다. 이 파일은 화면만 맡는다.
"""
import tkinter as tk
from tkinter import messagebox, ttk

import ai_client
import ai_commands as ac
from data_editor import FONT, TITLE

HISTORY = 12            # AI 에게 다시 보내는 지난 대화 수 (너무 길어지지 않게)
SAMPLES = 3             # '예시 값도 보내기'를 켰을 때 보내는 행 수
HELLO = ("무엇을 할지 우리말로 적어 보세요. 예)\n"
         "  · 전처리 해 줘  (프로그램이 표를 진단해서 AI 가 알아서 계획을 세웁니다)\n"
         "  · 지점코드가 X99 인 것을 ON1 로 바꿔 줘\n"
         "  · 신청금액이 3억 이상이고 반려된 건만 보여 줘\n"
         "  · 고객구분별 평균 신청금액을 알려 줘\n"
         "  · 월지급액을 숫자로 바꾸고 빈칸은 중앙값으로 채워 줘\n"
         "표를 바꾸는 일은 아래 '적용할 내용'에 먼저 보여 드리고, [적용]을 눌러야 실행합니다.\n"
         "키가 없으면 위의 [전처리 제안 (AI 없이)]을 눌러 보세요. 프로그램이 스스로 기본 전처리를 골라 줍니다.")


class AIWindow:
    is_ai = True

    def __init__(self, pane):
        self.pane = pane
        self.app = pane.app
        self.state = self.app.ai_state               # 고른 AI · 모델 · 넣은 키 (프로그램이 켜져 있는 동안만)
        self.ask = ai_client.ask                     # 점검할 때 가짜 함수로 바꿔 끼울 수 있다
        self.quiet = False                           # True 면 안내 창을 띄우지 않는다 (자동 점검용)
        self.busy = False
        self.history = []                            # AI 와 주고받은 말
        self.note = ""                               # 다음 말 앞에 붙여 AI 에게 알려 줄 것 (직전 명령을 적용했는지)
        self.pending = None                          # [적용]을 기다리는 명령들
        self.last_error = ""
        self.last_lines = []
        self.win = tk.Toplevel(pane.root)
        self.win.geometry("860x760")
        self.win.minsize(720, 600)
        self.win.protocol("WM_DELETE_WINDOW", self.close)
        self._build()
        self._on_provider()
        self._title()
        self._write("sys", HELLO)

    @property
    def session(self):
        return self.pane.session

    # ---------------- 화면 만들기 ----------------
    def _build(self):
        top = ttk.LabelFrame(self.win, text=" AI 고르기 ", padding=8)
        top.pack(fill="x", padx=10, pady=(8, 4))
        row = ttk.Frame(top)
        row.pack(fill="x")
        ttk.Label(row, text="AI").pack(side="left")
        self.keys = list(ai_client.PROVIDERS)
        self.provider_box = ttk.Combobox(row, state="readonly", width=22, font=FONT,
                                         values=[ai_client.PROVIDERS[k]["name"] for k in self.keys])
        self.provider_box.current(self.keys.index(self.state.get("provider", "gemini")))
        self.provider_box.pack(side="left", padx=(4, 12))
        self.provider_box.bind("<<ComboboxSelected>>", lambda e: self._on_provider())
        ttk.Label(row, text="모델").pack(side="left")
        self.model_box = ttk.Combobox(row, width=26, font=FONT)          # 목록에 없는 이름도 직접 적을 수 있다
        self.model_box.pack(side="left", padx=4)
        ttk.Label(row, text="(이름을 직접 고쳐 써도 됩니다)", foreground="#666").pack(side="left")
        row = ttk.Frame(top)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="API 키").pack(side="left")
        self.key_entry = ttk.Entry(row, width=34, show="•", font=FONT)
        self.key_entry.pack(side="left", padx=4)
        self.key_note = ttk.Label(row, text="", foreground="#666")
        self.key_note.pack(side="left", padx=4)
        self.url_row = ttk.Frame(top)
        ttk.Label(self.url_row, text="주소").pack(side="left")
        self.url_entry = ttk.Entry(self.url_row, width=46, font=FONT)
        self.url_entry.pack(side="left", padx=4)
        ttk.Label(self.url_row, text="OpenAI 와 같은 모양으로 받는 서버 (…/v1)", foreground="#666").pack(side="left")
        row = ttk.Frame(top)
        row.pack(fill="x", pady=(6, 0))
        self.samples = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, variable=self.samples,
                        text=f"예시 값도 보내기 (열마다 값 예시와 앞 {SAMPLES}행 — 가짜 데이터일 때만 켜세요)").pack(side="left")
        ttk.Button(row, text="AI 에게 보내는 내용 보기", command=self.show_payload).pack(side="right")
        ttk.Button(row, text="전처리 제안 (AI 없이)", command=self.suggest).pack(side="right", padx=6)

        bottom = ttk.Frame(self.win, padding=(10, 4, 10, 8))
        bottom.pack(side="bottom", fill="x")
        ttk.Button(bottom, text="닫기", command=self.close).pack(side="right")
        ttk.Button(bottom, text="되돌리기", command=self.undo).pack(side="right", padx=6)
        ttk.Button(bottom, text="대화 지우기", command=self.clear).pack(side="right")
        self.status = ttk.Label(bottom, text="", foreground="#333")
        self.status.pack(side="left")

        ask = ttk.Frame(self.win, padding=(10, 0, 10, 0))
        ask.pack(side="bottom", fill="x")
        self.send_button = ttk.Button(ask, text="보내기", command=self.send)
        self.send_button.pack(side="right", padx=(6, 0), fill="y")
        self.input = tk.Text(ask, height=3, font=FONT, wrap="word")
        self.input.pack(fill="x")
        self.input.bind("<Return>", self._enter)
        ttk.Label(self.win, text="Enter 보내기 · Shift+Enter 줄 바꿈", foreground="#666",
                  padding=(10, 2, 10, 0)).pack(side="bottom", anchor="w")

        plan = ttk.LabelFrame(self.win, text=" 적용할 내용 (미리 보기 — 아직 표는 그대로입니다) ", padding=8)
        plan.pack(side="bottom", fill="x", padx=10, pady=4)
        buttons = ttk.Frame(plan)
        buttons.pack(side="right", fill="y", padx=(8, 0))
        self.apply_button = ttk.Button(buttons, text="적용", command=self.apply, state="disabled")
        self.apply_button.pack(fill="x")
        self.discard_button = ttk.Button(buttons, text="버리기", command=self.discard, state="disabled")
        self.discard_button.pack(fill="x", pady=(4, 0))
        self.plan_box = tk.Text(plan, height=5, font=FONT, wrap="word", state="disabled", background="#FFFBEA")
        self.plan_box.pack(fill="x")

        box = ttk.Frame(self.win, padding=(10, 0, 10, 0))
        box.pack(fill="both", expand=True)
        self.chat = tk.Text(box, font=FONT, wrap="word", state="disabled", background="#FAFAFA", padx=8, pady=6)
        ys = ttk.Scrollbar(box, orient="vertical", command=self.chat.yview)
        self.chat.configure(yscrollcommand=ys.set)
        ys.pack(side="right", fill="y")
        self.chat.pack(fill="both", expand=True)
        self.chat.tag_configure("me", foreground="#0B3D91", font=(FONT[0], FONT[1], "bold"), spacing1=8)
        self.chat.tag_configure("ai", foreground="#111", spacing1=4)
        self.chat.tag_configure("done", foreground="#1B5E20", lmargin1=14, lmargin2=14)
        self.chat.tag_configure("sys", foreground="#666", lmargin1=14, lmargin2=14)
        self.chat.tag_configure("err", foreground="#B00020", lmargin1=14, lmargin2=14)

    def _title(self):
        self.win.title(f"AI 채팅 — {self.pane.source.path.name} [{self.pane.table}]")

    # ---------------- AI 고르기 ----------------
    @property
    def provider(self):
        return self.keys[self.provider_box.current()]

    def _remember(self):
        """지금 화면에 적힌 모델 · 키 · 주소를 기억해 둔다 (메모리에만)."""
        old = self.state.get("provider")
        if old:
            self.state.setdefault("models", {})[old] = self.model_box.get().strip()
            self.state.setdefault("keys", {})[old] = self.key_entry.get().strip()
            if old == "custom":
                self.state["url"] = self.url_entry.get().strip()

    def _on_provider(self):
        if self.model_box.get() or self.key_entry.get():
            self._remember()
        provider = self.provider
        info = ai_client.PROVIDERS[provider]
        self.state["provider"] = provider
        self.model_box["values"] = info["models"]
        self.model_box.set(self.state.get("models", {}).get(provider) or (info["models"][0] if info["models"] else ""))
        self.key_entry.delete(0, "end")
        self.key_entry.insert(0, self.state.get("keys", {}).get(provider, ""))
        if ai_client.env_key(provider):
            self.key_note["text"] = f"비워 두면 환경 변수 {info['env'][0]} 의 키를 씁니다"
        else:
            self.key_note["text"] = f"키 받는 곳: {info['get_key']}"
        if provider == "custom":
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, self.state.get("url") or info["url"])
            self.url_row.pack(fill="x", pady=(6, 0), after=self.key_entry.master)
        else:
            self.url_row.pack_forget()

    def choose(self, provider, model=None, key=None, url=None):
        """코드로 AI 를 고른다 (자동 점검 · 노트북용)."""
        self.provider_box.current(self.keys.index(provider))
        self._on_provider()
        if model is not None:
            self.model_box.set(model)
        if key is not None:
            self.key_entry.delete(0, "end")
            self.key_entry.insert(0, key)
        if url is not None:
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, url)

    # ---------------- 대화 ----------------
    def _write(self, tag, text):
        self.chat.configure(state="normal")
        self.chat.insert("end", text.rstrip("\n") + "\n", tag)
        self.chat.see("end")
        self.chat.configure(state="disabled")

    def _enter(self, event):
        if event.state & 0x1:                        # Shift+Enter 는 줄 바꿈
            return None
        self.send()
        return "break"

    def _request(self, text):
        """AI 에게 보낼 (안내문, 대화)를 만든다."""
        system = ac.system_prompt(self.session, self.pane.table, SAMPLES if self.samples.get() else 0)
        said = (f"[프로그램 알림] {self.note}\n\n" if self.note else "") + text
        return system, self.history[-HISTORY:] + [{"role": "user", "content": said}]

    def send(self, text=None):
        """AI 에게 보낸다. 보냈으면 True. 답은 작업 스레드에서 받아 화면 스레드에서 처리한다."""
        if self.busy or self.session is None:
            return False
        text = (self.input.get("1.0", "end") if text is None else text).strip()
        if not text:
            return False
        if self.pending:
            self.discard(say=False)
        self._remember()
        provider, model = self.provider, self.model_box.get().strip()
        key = self.key_entry.get().strip() or ai_client.env_key(provider)
        url = self.url_entry.get().strip() if provider == "custom" else None
        system, messages = self._request(text)
        self.input.delete("1.0", "end")
        self._write("me", "나: " + text)
        self._busy(f"{ai_client.PROVIDERS[provider]['name']} 에게 묻는 중…")
        ask = self.ask

        def work():                                  # ← 작업 스레드 (그동안 창은 멈추지 않는다)
            return ask(provider, model, key, system, messages, base_url=url)

        def done(raw, error):                        # ← 화면 스레드
            if not self.win.winfo_exists():
                return
            self._busy("")
            if error:
                self._error(str(error) if isinstance(error, ai_client.AIError) else f"AI 호출 중 문제가 생겼습니다: {error}")
                return
            self.note = ""
            self.history += [messages[-1], {"role": "assistant", "content": raw}]
            try:
                self._answer(raw)
            except Exception as e:                   # 어떤 답이 와도 창이 죽지 않게
                self._error(f"답을 처리하는 중 예상하지 못한 문제가 생겼습니다: {type(e).__name__}: {e}")

        self.app.run_job(work, done)
        return True

    def _answer(self, raw):
        reply, actions = ac.parse_reply(raw)
        if reply:
            self._write("ai", "AI: " + reply)
        if not actions:
            return
        if self.pane.busy:
            self._error("표를 읽거나 저장하는 중이라 지금은 실행할 수 없습니다. 끝난 뒤 다시 말해 주세요.")
            return
        try:
            lines = ac.preview(self.session, actions)
        except ValueError as e:
            self.note = f"직전에 만든 명령은 실행할 수 없어 버렸다. 이유: {e}"
            self._error(f"AI 가 만든 명령을 실행할 수 없습니다 — {e}\n다시 말해 보세요. (표는 그대로입니다)")
            return
        if not ac.changes_data(actions):             # 보기 · 정렬 · 집계는 표를 바꾸지 않으니 바로 한다
            self._run(actions)
            return
        self._offer(actions, lines)

    def _offer(self, actions, lines):
        """표를 바꾸는 명령을 '적용할 내용'에 올려놓고 [적용]을 기다린다."""
        self.pending = actions
        self.last_lines = lines
        self._plan("\n".join(f"{n}. {line}" for n, line in enumerate(lines, 1)))
        self._write("sys", f"→ 아직 실행하지 않았습니다. 아래 '적용할 내용' {len(lines)}가지를 확인하고 [적용]을 누르세요.")
        self.apply_button.focus_set()

    def suggest(self):
        """AI 를 부르지 않고, 프로그램이 표를 진단해 기본 전처리를 골라 '적용할 내용'에 올린다. 올렸으면 True."""
        if self.busy or self.session is None:
            return False
        self.pane.commit_edit()
        if not self.pane.ready():
            return False
        if self.pending:
            self.discard(say=False)
        actions, notes = ac.suggest(self.session)
        self._write("me", "나: [전처리 제안] — AI 없이 프로그램이 표를 살펴봅니다")
        if not actions:
            self._write("ai", "프로그램: 기본 전처리로 손볼 곳을 찾지 못했습니다 (앞뒤 빈칸 · 겹친 행 · 글자로 적힌 숫자와 날짜 · 숫자 열의 빈칸이 없음).")
            for note in notes:
                self._write("sys", note)
            return False
        self._write("ai", "프로그램: 표를 진단해 기본 전처리를 골랐습니다 — 앞뒤 빈칸 지우기 → 겹친 행 삭제 → 글자로 적힌 숫자 · 날짜 바꾸기 → 숫자 열 빈칸을 중앙값으로.")
        for note in notes:
            self._write("sys", note)
        self._offer(actions, ac.preview(self.session, actions))
        self.note = "프로그램이 AI 없이 기본 전처리 명령을 만들어 사용자에게 보여 주었다."
        return True

    def apply(self):
        """미리 보여 준 명령을 실제로 실행한다."""
        if not self.pending or self.busy:
            return False
        actions, self.pending = self.pending, None
        self._plan("")
        return self._run(actions)

    def discard(self, say=True):
        if not self.pending:
            return False
        self.pending = None
        self._plan("")
        self.note = "직전에 만든 명령은 사용자가 적용하지 않고 버렸다."
        if say:
            self._write("sys", "버렸습니다. 표는 그대로입니다.")
        return True

    def _run(self, actions):
        self.pane.commit_edit()
        if not self.pane.ready():
            return False
        try:
            result = ac.run(self.session, actions)
        except ValueError as e:
            self.note = f"직전 명령은 실행 중 오류가 나서 취소됐다. 이유: {e}"
            self._error(f"실행하지 못했습니다 — {e}\n(표는 그대로입니다)")
            return False
        self.last_lines = result.lines
        if result.changed:
            self.pane.after_preprocess()
            self.note = "직전에 만든 명령을 사용자가 적용했다."
        if result.view is not None:                  # 보기 · 정렬
            self.pane.view = result.view
            self.pane.refresh()
        for window in self.pane.windows:             # 열려 있는 전처리 창도 새 표로
            if hasattr(window, "reload"):
                window.reload()
        for line in result.lines:
            self._write("done", "✓ " + line)
        if result.changed:
            self._write("sys", "적용했습니다. 마음에 들지 않으면 [되돌리기]를 누르세요. 파일은 [저장]을 눌러야 바뀝니다.")
        self._title()
        return True

    def undo(self):
        self.pane.commit_edit()
        if not self.pane.ready() or self.session is None:
            return False
        if not self.session.undo():
            self._write("sys", "되돌릴 작업이 없습니다.")
            return False
        self.pane.after_preprocess()
        for window in self.pane.windows:
            if hasattr(window, "reload"):
                window.reload()
        self.note = "사용자가 직전 변경을 되돌렸다."
        self._write("sys", "마지막 변경을 되돌렸습니다.")
        return True

    def clear(self):
        self.history, self.note, self.pending = [], "", None
        self._plan("")
        self.chat.configure(state="normal")
        self.chat.delete("1.0", "end")
        self.chat.configure(state="disabled")
        self._write("sys", HELLO)

    # ---------------- 작은 것들 ----------------
    def _busy(self, text):
        self.busy = bool(text)
        self.status["text"] = text
        self.send_button.configure(state="disabled" if text else "normal")

    def _plan(self, text):
        self.plan_box.configure(state="normal")
        self.plan_box.delete("1.0", "end")
        self.plan_box.insert("1.0", text)
        self.plan_box.configure(state="disabled")
        state = "normal" if text else "disabled"
        self.apply_button.configure(state=state)
        self.discard_button.configure(state=state)

    def _error(self, text):
        self.last_error = text
        self._write("err", "✗ " + text)

    def show_payload(self):
        """지금 [보내기]를 누르면 AI 에게 가는 글을 그대로 보여 준다 (키는 여기에 없다)."""
        if self.session is None:
            return None
        system, messages = self._request(self.input.get("1.0", "end").strip() or "(여기에 내가 친 말이 들어갑니다)")
        text = "■ 안내문 (표의 모양과 규칙)\n\n" + system + "\n\n■ 대화\n\n" + "\n\n".join(
            f"[{'나' if m['role'] == 'user' else 'AI'}] {m['content']}" for m in messages)
        if self.quiet:
            return text
        top = tk.Toplevel(self.win)
        top.title("AI 에게 보내는 내용")
        top.geometry("760x640")
        box = tk.Text(top, font=FONT, wrap="word", padx=8, pady=6)
        box.insert("1.0", text)
        box.configure(state="disabled")
        box.pack(fill="both", expand=True)
        return text

    def close(self):
        self._remember() if self.win.winfo_exists() else None
        if self in self.pane.windows:
            self.pane.windows.remove(self)
        if self.win.winfo_exists():
            self.win.destroy()


def ask_once(pane, text):                            # 노트북 · 스크립트에서 한 번만 물어볼 때
    window = pane.open_ai()
    return window.send(text) if window else False


__all__ = ["AIWindow", "ask_once", "messagebox", "TITLE"]
