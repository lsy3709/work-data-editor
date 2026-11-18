"""
AI(LLM)에게 묻고 답을 글자로 받아 오는 부분 — 화면도, 표도 모른다.

    text = ask("gemini", "gemini-flash-latest", 키, 안내문, [{"role": "user", "content": "질문"}])

회사마다 주소 · 머리말 · 본문 모양이 달라서 이 파일이 그 차이를 감춘다.
  gemini  Google Gemini        REST (파이썬 기본 라이브러리만 씀)
  openai  OpenAI GPT           REST, Chat Completions 모양
  claude  Anthropic Claude     공식 라이브러리 anthropic (pip install anthropic)
  custom  직접 지정             OpenAI 와 같은 모양으로 받는 서버면 무엇이든 (사내 모델, Ollama, LM Studio …)

API 키는 인자로만 받고 이 파일은 어디에도 저장하지 않는다. 오류 글에도 키가 섞이지 않게 지운다.
"""
import json
import os
import socket
import urllib.error
import urllib.request

PROVIDERS = {
    "gemini": {"name": "Gemini (Google)", "env": ["GEMINI_API_KEY", "GOOGLE_API_KEY"],
               "url": "https://generativelanguage.googleapis.com",
               "models": ["gemini-flash-latest", "gemini-pro-latest", "gemini-flash-lite-latest"],
               "get_key": "aistudio.google.com → Get API key"},
    "openai": {"name": "GPT (OpenAI)", "env": ["OPENAI_API_KEY"],
               "url": "https://api.openai.com/v1",
               "models": ["gpt-6.2-astra", "gpt-6.2-sol", "gpt-6.1-terra", "gpt-6-luna"],
               "get_key": "platform.openai.com → API keys"},
    "claude": {"name": "Claude (Anthropic)", "env": ["ANTHROPIC_API_KEY"],
               "url": "",
               "models": ["claude-opus-5-5", "claude-sonnet-5-5", "claude-haiku-5-5", "claude-fable-5-2"],
               "get_key": "platform.claude.com → API keys"},
    "custom": {"name": "직접 지정 (OpenAI 호환)", "env": ["CUSTOM_LLM_API_KEY"],
               "url": "http://localhost:11434/v1",
               "models": [],
               "get_key": "쓰는 서버의 안내를 따르세요 (내 PC 의 Ollama 는 키가 없어도 됩니다)"},
}
# 안전 장치가 답을 거절했을 때 다른 모델로 한 번 더 시도하는 기능(fallbacks)을 켜 두는 Claude 모델
CLAUDE_FALLBACK = {"claude-opus-5-5", "claude-opus-5", "claude-sonnet-5-5", "claude-fable-5-2", "claude-fable-5-1"}


class AIError(Exception):
    """AI 호출이 실패했을 때. 글은 사용자에게 그대로 보여 줄 수 있는 우리말이다."""


def env_key(provider):
    """환경 변수에 넣어 둔 API 키. 없으면 빈 글자."""
    for name in PROVIDERS[provider]["env"]:
        if os.environ.get(name):
            return os.environ[name]
    return ""


def ask(provider, model, key, system, messages, base_url=None, timeout=90):
    """AI 에게 묻고 답 글자를 돌려준다. messages = [{"role": "user" 또는 "assistant", "content": 글자}, …]"""
    if provider not in PROVIDERS:
        raise AIError(f"모르는 AI 종류입니다: {provider}")
    model = (model or "").strip()
    key = (key or "").strip()
    if not model:
        raise AIError("모델 이름을 입력하세요.")
    if not key and provider != "custom":
        raise AIError(f"API 키가 없습니다. 키를 입력하거나 환경 변수 {PROVIDERS[provider]['env'][0]} 에 넣어 두세요.")
    if not model.isascii() or any(ch.isspace() for ch in model):
        raise AIError(f"모델 이름에 한글이나 빈칸이 있습니다: '{model}'. 영문 모델 이름을 그대로 적으세요 (예: {(PROVIDERS[provider]['models'] or ['모델이름'])[0]}).")
    if not key.isascii() or any(ch.isspace() for ch in key):
        raise AIError("API 키에 한글이나 빈칸이 섞여 있습니다. 키를 다시 복사해 붙여 넣으세요.")
    base = (base_url or PROVIDERS[provider]["url"]).rstrip("/")
    try:
        if provider == "gemini":
            text = _gemini(base, model, key, system, messages, timeout)
        elif provider == "claude":
            text = _claude(base, model, key, system, messages, timeout)
        else:
            text = _openai(base, model, key, system, messages, timeout, json_mode=provider == "openai")
    except AIError as e:
        raise AIError(_hide(str(e), key)) from None
    if not (text or "").strip():
        raise AIError("AI 가 빈 답을 보냈습니다. 다시 물어보거나 모델을 바꿔 보세요.")
    return text


# ---------------- 회사별 ----------------
def _gemini(base, model, key, system, messages, timeout):
    body = {"system_instruction": {"parts": [{"text": system}]},
            "contents": [{"role": "model" if m["role"] == "assistant" else "user", "parts": [{"text": m["content"]}]} for m in messages],
            "generationConfig": {"responseMimeType": "application/json"}}
    data = _post(f"{base}/v1beta/models/{model}:generateContent", {"x-goog-api-key": key}, body, timeout)
    candidates = data.get("candidates") or []
    if not candidates:
        reason = (data.get("promptFeedback") or {}).get("blockReason", "")
        raise AIError("Gemini 가 답하지 않았습니다" + (f" (차단 사유: {reason})" if reason else "") + ". 질문을 바꿔 보세요.")
    parts = (candidates[0].get("content") or {}).get("parts") or []
    text = "".join(p.get("text", "") for p in parts if not p.get("thought"))
    if not text and candidates[0].get("finishReason") not in (None, "STOP"):
        raise AIError(f"Gemini 가 답을 끝내지 못했습니다 (사유: {candidates[0].get('finishReason')}).")
    return text


def _openai(base, model, key, system, messages, timeout, json_mode):
    body = {"model": model, "messages": [{"role": "system", "content": system}] + [{"role": m["role"], "content": m["content"]} for m in messages]}
    if json_mode:
        body["response_format"] = {"type": "json_object"}
    data = _post(f"{base}/chat/completions", {"Authorization": f"Bearer {key}"} if key else {}, body, timeout)
    choices = data.get("choices") or []
    if not choices:
        raise AIError("AI 가 답하지 않았습니다. 모델 이름과 주소를 확인하세요.")
    message = choices[0].get("message") or {}
    if message.get("refusal"):
        raise AIError("AI 가 이 요청에는 답하지 않겠다고 했습니다: " + str(message["refusal"])[:200])
    content = message.get("content")
    if isinstance(content, list):                          # 조각으로 나눠 보내는 서버
        content = "".join(part.get("text", "") for part in content if isinstance(part, dict))
    return content or ""


def _claude(base, model, key, system, messages, timeout):
    try:
        import anthropic
    except ImportError:
        raise AIError("Claude 를 쓰려면 anthropic 라이브러리가 필요합니다. 터미널에서:  pip install anthropic") from None
    client = anthropic.Anthropic(api_key=key, base_url=base or None, timeout=timeout, max_retries=1)
    request = dict(model=model, max_tokens=16000, system=system, messages=[{"role": m["role"], "content": m["content"]} for m in messages])
    try:
        response = None
        if model in CLAUDE_FALLBACK:                       # 거절되면 서버가 다른 모델로 한 번 더 시도하게 한다
            try:
                response = client.beta.messages.create(betas=["server-side-fallback-2026-07-01"], fallbacks="default", **request)
            except (anthropic.BadRequestError, TypeError):
                response = None                            # 이 기능을 받지 않는 계정 · 라이브러리면 보통 요청으로
        if response is None:
            response = client.messages.create(**request)
    except anthropic.AuthenticationError:
        raise AIError(_status_text(401, "")) from None
    except anthropic.PermissionDeniedError as e:
        raise AIError(_status_text(403, getattr(e, "message", ""))) from None
    except anthropic.NotFoundError as e:
        raise AIError(_status_text(404, getattr(e, "message", ""))) from None
    except anthropic.RateLimitError as e:
        raise AIError(_status_text(429, getattr(e, "message", ""))) from None
    except anthropic.APIStatusError as e:
        raise AIError(_status_text(e.status_code, getattr(e, "message", ""))) from None
    except anthropic.APIConnectionError:
        raise AIError(_status_text(0, "")) from None
    if response.stop_reason == "refusal":
        raise AIError("Claude 가 이 요청에는 답하지 않겠다고 했습니다. 질문을 바꿔 보세요.")
    text = "".join(block.text for block in response.content if block.type == "text")
    if response.stop_reason == "max_tokens" and not text:
        raise AIError("Claude 의 답이 너무 길어 중간에 끊겼습니다. 요청을 나눠서 해 보세요.")
    return text


# ---------------- 공통 ----------------
def _post(url, headers, body, timeout):
    request = urllib.request.Request(url, data=json.dumps(body, ensure_ascii=False).encode("utf-8"), method="POST",
                                     headers={"Content-Type": "application/json", **headers})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read()
    except urllib.error.HTTPError as e:
        raise AIError(_status_text(e.code, _detail(e.read()))) from None
    except (UnicodeError, ValueError):
        raise AIError("주소나 모델 이름에 쓸 수 없는 글자가 있습니다. 영문 주소 · 모델 이름인지 확인하세요.") from None
    except (urllib.error.URLError, socket.timeout, TimeoutError, ConnectionError, OSError):
        raise AIError(_status_text(0, "")) from None
    try:
        return json.loads(raw)
    except ValueError:
        raise AIError("AI 서버가 알 수 없는 모양으로 답했습니다. 주소가 맞는지 확인하세요.") from None


def _detail(raw):
    """오류 응답에서 사람이 읽을 한 줄을 꺼낸다."""
    try:
        error = json.loads(raw).get("error")
        text = error.get("message", "") if isinstance(error, dict) else str(error or "")
    except (ValueError, AttributeError):
        text = raw.decode("utf-8", "replace") if isinstance(raw, bytes) else str(raw)
    return " ".join(text.split())[:300]


def _status_text(code, detail):
    low = detail.lower()
    if code == 0:
        text = "AI 서버에 연결하지 못했습니다. 인터넷 연결 · 주소 · 회사 방화벽을 확인하세요."
    elif code in (401, 403) or "api key" in low or "api_key" in low:
        text = "API 키가 맞지 않거나 이 모델을 쓸 권한이 없습니다. 키를 다시 확인하세요."
    elif code == 404:
        text = "모델 이름이나 주소를 찾을 수 없습니다. 모델 이름을 확인하세요."
    elif code == 429:
        text = "사용 한도를 넘었습니다. 잠시 뒤 다시 하거나, 그 서비스의 요금제 · 한도를 확인하세요."
    elif code >= 500:
        text = "AI 서버에 문제가 있습니다. 잠시 뒤 다시 해 보세요."
    else:
        text = "요청을 받아 주지 않았습니다. 모델 이름을 확인하세요."
    return text + (f"\n(오류 {code}: {detail})" if detail and code else "")


def _hide(text, key):
    return text.replace(key, "***") if key else text
