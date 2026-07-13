"""Vercel callback endpoint for n8n async workflows."""

from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler
from typing import Any

MAX_CALLBACK_BYTES = 25 * 1024 * 1024


def _json_bytes(payload: dict[str, Any]) -> bytes:
    return json.dumps(payload).encode("utf-8")


class handler(BaseHTTPRequestHandler):
    def _send_json(self, status_code: int, payload: dict[str, Any]) -> None:
        response_bytes = _json_bytes(payload)
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(response_bytes)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(response_bytes)

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS, GET")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self) -> None:
        self._send_json(200, {"ok": True, "message": "n8n callback endpoint is ready"})

    def do_POST(self) -> None:
        content_length = int(self.headers.get("content-length", "0") or "0")
        if content_length > MAX_CALLBACK_BYTES:
            self._send_json(413, {"error": "Callback payload is too large."})
            return

        body = self.rfile.read(content_length) if content_length > 0 else b""
        parsed: Any = None
        if body:
            body_text = body.decode("utf-8", errors="replace")
            try:
                parsed = json.loads(body_text)
            except json.JSONDecodeError:
                parsed = {"raw": body_text}

        request_id = ""
        if isinstance(parsed, dict):
            request_id = str(
                parsed.get("request_id")
                or parsed.get("requestId")
                or parsed.get("id")
                or ""
            ).strip()

        self._send_json(
            200,
            {
                "ok": True,
                "received": True,
                "request_id": request_id,
            },
        )
