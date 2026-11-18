"""
가짜 AI 서버 — 인터넷과 API 키 없이 AI 채팅을 시험할 때 쓴다 (테스트 · 화면 자동 점검용).

내 PC 안(127.0.0.1)에서만 뜨고, 세 회사의 응답 모양을 흉내 낸다.
  POST /v1beta/models/<모델>:generateContent   Gemini 모양
  POST /v1/chat/completions                    OpenAI 모양 (직접 지정도 이 모양)
  POST /v1/messages                            Claude 모양

쓰는 법:
    server = ai_mock_server.start(lambda question: '{"reply": "네", "actions": []}')
    ... server.url 을 '주소'로 넣어 호출 ...
    server.requests    # 받은 요청들 (경로 · 머리말 · 본문) — 무엇을 보냈는지 검사할 때
    server.stop()
"""
import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer


class MockServer:
    def __init__(self, answer, status=200):
        self.answer = answer            # 마지막 사용자 말(글자)을 받아 답(글자)을 돌려주는 함수
        self.status = status            # 200 이 아니면 그 오류 코드로 답한다
        self.requests = []
        owner = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):          # 터미널에 접속 기록을 찍지 않는다
                pass

            def do_POST(self):
                body = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
                owner.requests.append({"path": self.path, "headers": {k.lower(): v for k, v in self.headers.items()}, "body": body})
                if owner.status != 200:
                    return self._send(owner.status, {"error": {"message": "가짜 서버가 일부러 낸 오류", "type": "mock_error"}})
                path = self.path.split("?")[0]
                if ":generateContent" in path:
                    text = owner.answer(body["contents"][-1]["parts"][0]["text"])
                    reply = {"candidates": [{"content": {"role": "model", "parts": [{"text": text}]}, "finishReason": "STOP"}]}
                elif path.endswith("/chat/completions"):
                    text = owner.answer(body["messages"][-1]["content"])
                    reply = {"choices": [{"index": 0, "message": {"role": "assistant", "content": text}, "finish_reason": "stop"}]}
                elif path.endswith("/messages"):
                    last = body["messages"][-1]["content"]
                    text = owner.answer(last if isinstance(last, str) else last[0]["text"])
                    reply = {"id": "msg_mock", "type": "message", "role": "assistant", "model": body.get("model", "mock"),
                             "content": [{"type": "text", "text": text}], "stop_reason": "end_turn", "stop_sequence": None,
                             "usage": {"input_tokens": 1, "output_tokens": 1}}
                else:
                    return self._send(404, {"error": {"message": "모르는 주소: " + path}})
                self._send(200, reply)

            def _send(self, status, data):
                raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)

        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)       # 0 = 비어 있는 포트를 알아서
        self.url = f"http://127.0.0.1:{self.httpd.server_address[1]}"
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()

    def stop(self):
        self.httpd.shutdown()
        self.httpd.server_close()


def start(answer, status=200):
    return MockServer(answer, status)
