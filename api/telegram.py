"""Telegram webhook endpoint for the CRM assistant bot."""

from __future__ import annotations

import json
import os
from http.server import BaseHTTPRequestHandler
from typing import Any, Mapping


MAX_UPDATE_BYTES = 1024 * 1024


def _dashboard_url(dashboard_url: str | None = None) -> str:
    return (dashboard_url or os.environ.get("PUBLIC_APP_URL") or "").rstrip("/")


def _dashboard_line(dashboard_url: str | None = None) -> str:
    url = _dashboard_url(dashboard_url)
    if url:
        return f"Panel: {url}"
    return "Panel linki icin PUBLIC_APP_URL env degiskenini ayarlayin."


def _start_message(dashboard_url: str | None = None) -> str:
    return "\n".join(
        [
            "Merhaba! CRM TR Asistan botuna hos geldiniz.",
            "",
            "Bu bot CRM + PowerBI rapor panelini kullanmaniza yardim eder.",
            _dashboard_line(dashboard_url),
            "",
            "Komutlar:",
            "/programa - CRM Output Report aciklamasi",
            "/programb - Country Split Report aciklamasi",
            "/columns - Gerekli Excel kolonlari",
            "/link - Rapor paneli linki",
            "/help - Yardim",
        ]
    )


PROGRAM_A_MESSAGE = "\n".join(
    [
        "Program A - CRM Output Report",
        "",
        "Tek bir CRM Output sayfasi uretir.",
        "Status, call-attempt ve campaign pivotlarini ayni workbook icine ekler.",
        "Program A icin pivot table name alani zorunludur.",
    ]
)


PROGRAM_B_MESSAGE = "\n".join(
    [
        "Program B - Country Split Report",
        "",
        "Main Report sayfasi ve Country kolonuna gore ayri ulke sayfalari uretir.",
        "Ulke sayfalari Main Report'a formul baglantilariyla baglanir.",
        "Program B pivot adlarini otomatik kullanir.",
    ]
)


COLUMNS_MESSAGE = "\n".join(
    [
        "Gerekli Excel kolonlari:",
        "",
        "CRM dosyalari:",
        "Customer Type, ID, Created, Name, Department, Status, Country, Campaign, "
        "Sub-Campaign, Placement, Assigned to",
        "",
        "PowerBI raporu:",
        "Account No, Brand, Last 10 Comments, Voip Calls Attempts Cnt",
    ]
)


def _document_message(dashboard_url: str | None = None) -> str:
    return "\n".join(
        [
            "Excel dosyasi aldim, ancak rapor uretimi web panelinden yapiliyor.",
            "Lutfen PowerBI ve CRM dosyalarini panelde yukleyin.",
            _dashboard_line(dashboard_url),
        ]
    )


def _fallback_message(dashboard_url: str | None = None) -> str:
    return "\n".join(
        [
            "Size CRM + PowerBI rapor paneli konusunda yardim edebilirim.",
            "Komutlari gormek icin /help yazin.",
            _dashboard_line(dashboard_url),
        ]
    )


def _extract_message(update: Mapping[str, Any]) -> Mapping[str, Any] | None:
    for key in ("message", "edited_message", "channel_post", "edited_channel_post"):
        value = update.get(key)
        if isinstance(value, Mapping):
            return value
    return None


def _command_from_text(text: str) -> str:
    first_word = text.strip().split(maxsplit=1)[0]
    return first_word.split("@", maxsplit=1)[0].lower()


def build_reply(
    update: Mapping[str, Any],
    dashboard_url: str | None = None,
) -> dict[str, Any] | None:
    """Build a Telegram webhook reply payload for one incoming update."""

    message = _extract_message(update)
    if message is None:
        return None

    chat = message.get("chat")
    if not isinstance(chat, Mapping):
        return None

    chat_id = chat.get("id")
    if chat_id is None:
        return None

    if "document" in message:
        text = _document_message(dashboard_url)
    else:
        raw_text = str(message.get("text") or "").strip()
        command = _command_from_text(raw_text) if raw_text.startswith("/") else ""

        if command in {"/start", "/help"}:
            text = _start_message(dashboard_url)
        elif command in {"/programa", "/program_a", "/a"}:
            text = PROGRAM_A_MESSAGE
        elif command in {"/programb", "/program_b", "/b"}:
            text = PROGRAM_B_MESSAGE
        elif command in {"/columns", "/kolonlar"}:
            text = COLUMNS_MESSAGE
        elif command in {"/link", "/panel"}:
            text = _dashboard_line(dashboard_url)
        else:
            text = _fallback_message(dashboard_url)

    return {
        "method": "sendMessage",
        "chat_id": chat_id,
        "text": text,
        "disable_web_page_preview": True,
    }


def is_authorized(headers: Mapping[str, str], secret: str | None = None) -> bool:
    expected_secret = secret if secret is not None else os.environ.get("TELEGRAM_WEBHOOK_SECRET")
    if not expected_secret:
        return True
    return headers.get("X-Telegram-Bot-Api-Secret-Token") == expected_secret


def _json_bytes(payload: dict[str, Any]) -> bytes:
    return json.dumps(payload).encode("utf-8")


class handler(BaseHTTPRequestHandler):
    def _send_json(self, status_code: int, payload: dict[str, Any]) -> None:
        body = _json_bytes(payload)
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        self._send_json(
            200,
            {
                "ok": True,
                "service": "telegram-webhook",
            },
        )

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self) -> None:
        if not is_authorized(self.headers):
            self._send_json(401, {"ok": False, "error": "Unauthorized"})
            return

        try:
            content_length = int(self.headers.get("content-length", "0") or "0")
        except ValueError:
            self._send_json(400, {"ok": False, "error": "Invalid content length"})
            return

        if content_length <= 0:
            self._send_json(400, {"ok": False, "error": "Empty update"})
            return
        if content_length > MAX_UPDATE_BYTES:
            self._send_json(413, {"ok": False, "error": "Update too large"})
            return

        try:
            raw_body = self.rfile.read(content_length)
            update = json.loads(raw_body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            self._send_json(400, {"ok": False, "error": "Invalid JSON"})
            return

        reply = build_reply(update)
        if reply is None:
            self.send_response(204)
            self.end_headers()
            return

        self._send_json(200, reply)
