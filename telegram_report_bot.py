"""Standalone Telegram bot for the Report app workflow.

This script reproduces the Report tab behavior (Program A / Program B):
- Upload one PowerBI report (.xlsx)
- Upload one or more CRM files (.xlsx), each with a platform name
- Optional PowerBI sheet name
- Optional CRM sheet name
- Program A requires a pivot table label
- Optional output filename

It then generates the same workbook output as the web app by calling:
- program_a_report.build_output
- program_b_country_report.build_output

Important:
- This script is standalone and not wired into web deployment.
- Run it manually in an environment with TELEGRAM_BOT_TOKEN set.
"""

from __future__ import annotations

import os
import re
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import requests

import program_a_report
import program_b_country_report

PROGRAM_A = "program_a"
PROGRAM_B = "program_b"
PROGRAM_A_OUTPUT_FILENAME = "crm_powerbi_output.xlsx"
PROGRAM_B_OUTPUT_FILENAME = "crm_country_report.xlsx"

TELEGRAM_API_ROOT = "https://api.telegram.org"
POLL_TIMEOUT_SECONDS = 40
NETWORK_TIMEOUT_SECONDS = 90
RETRY_DELAY_SECONDS = 2


def _safe_filename(raw_name: str | None, fallback: str) -> str:
    name = Path(raw_name or fallback).name
    name = re.sub(r"[^A-Za-z0-9_.-]+", "_", name).strip("._")
    return name or fallback


def _output_filename(raw_name: str | None, fallback: str) -> str:
    name = _safe_filename(raw_name, fallback)
    if not name.lower().endswith(".xlsx"):
        name = f"{name}.xlsx"
    return name


def _is_skip(text: str | None) -> bool:
    return (text or "").strip().lower() in {"/skip", "skip", "-"}


@dataclass
class CrmUpload:
    file_id: str
    filename: str
    platform: str


@dataclass
class ChatSession:
    stage: str = "await_program"
    program: str | None = None
    powerbi_file_id: str | None = None
    powerbi_filename: str | None = None
    powerbi_sheet: str | None = None
    pivot_name: str | None = None
    crm_uploads: list[CrmUpload] = field(default_factory=list)
    pending_crm_file_id: str | None = None
    pending_crm_filename: str | None = None
    crm_sheet: str | None = None
    output_file: str | None = None


class TelegramReportBot:
    def __init__(self, token: str) -> None:
        self.token = token
        self.base_url = f"{TELEGRAM_API_ROOT}/bot{token}"
        self.file_url = f"{TELEGRAM_API_ROOT}/file/bot{token}"
        self.http = requests.Session()
        self.sessions: dict[int, ChatSession] = {}

    def _api(self, method: str, *, data: dict[str, Any] | None = None, files: dict[str, Any] | None = None) -> dict[str, Any]:
        url = f"{self.base_url}/{method}"
        response = self.http.post(
            url,
            data=data,
            files=files,
            timeout=NETWORK_TIMEOUT_SECONDS,
        )
        response.raise_for_status()
        payload = response.json()
        if not payload.get("ok"):
            raise RuntimeError(f"Telegram API error ({method}): {payload}")
        return payload

    def send_message(self, chat_id: int, text: str) -> None:
        self._api("sendMessage", data={"chat_id": str(chat_id), "text": text})

    def send_document(self, chat_id: int, path: Path, caption: str) -> None:
        with path.open("rb") as handle:
            self._api(
                "sendDocument",
                data={"chat_id": str(chat_id), "caption": caption},
                files={"document": (path.name, handle, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            )

    def download_file(self, file_id: str, target_path: Path) -> None:
        file_info = self._api("getFile", data={"file_id": file_id})["result"]
        file_path = file_info["file_path"]
        download_url = f"{self.file_url}/{file_path}"
        with self.http.get(download_url, timeout=NETWORK_TIMEOUT_SECONDS, stream=True) as response:
            response.raise_for_status()
            with target_path.open("wb") as output:
                for chunk in response.iter_content(chunk_size=1024 * 64):
                    if chunk:
                        output.write(chunk)

    def reset_session(self, chat_id: int) -> ChatSession:
        session = ChatSession()
        self.sessions[chat_id] = session
        return session

    def session(self, chat_id: int) -> ChatSession:
        return self.sessions.setdefault(chat_id, ChatSession())

    def _help_text(self) -> str:
        return (
            "Report Bot (same workflow as Report app)\n\n"
            "Commands:\n"
            "/start - reset and start\n"
            "/program_a - Program A (CRM Output Report)\n"
            "/program_b - Program B (Country Split Report)\n"
            "/done - finish CRM uploads\n"
            "/skip - skip optional field\n"
            "/cancel - cancel current session\n"
        )

    def _prompt_program(self, chat_id: int) -> None:
        self.send_message(
            chat_id,
            "Choose report program:\n/program_a for Program A\n/program_b for Program B",
        )

    def _expect_powerbi(self, chat_id: int) -> None:
        self.send_message(chat_id, "Upload the PowerBI report file (.xlsx).")

    def _expect_powerbi_sheet(self, chat_id: int) -> None:
        self.send_message(
            chat_id,
            "Optional: send PowerBI sheet name, or /skip to use active sheet.",
        )

    def _expect_pivot_name(self, chat_id: int) -> None:
        self.send_message(chat_id, "Program A: send pivot table name (required).")

    def _expect_crm_or_done(self, chat_id: int) -> None:
        self.send_message(
            chat_id,
            "Upload a CRM file (.xlsx). After each upload, send its platform name.\n"
            "When finished uploading all CRM files, send /done.",
        )

    def _expect_platform(self, chat_id: int) -> None:
        self.send_message(chat_id, "Send platform name for the CRM file you just uploaded.")

    def _expect_crm_sheet(self, chat_id: int) -> None:
        self.send_message(
            chat_id,
            "Optional: send CRM sheet name for all CRM files, or /skip.",
        )

    def _expect_output_name(self, chat_id: int, default_name: str) -> None:
        self.send_message(
            chat_id,
            f"Optional: send output filename, or /skip for default ({default_name}).",
        )

    def _set_program(self, chat_id: int, session: ChatSession, program: str) -> None:
        session.program = program
        session.stage = "await_powerbi"
        self.send_message(
            chat_id,
            f"Selected {'Program A' if program == PROGRAM_A else 'Program B'}.",
        )
        self._expect_powerbi(chat_id)

    def _document_meta(self, message: dict[str, Any]) -> tuple[str, str] | None:
        document = message.get("document")
        if not document:
            return None
        file_id = document.get("file_id")
        filename = document.get("file_name") or "upload.xlsx"
        if not file_id:
            return None
        return file_id, filename

    def _ensure_xlsx(self, filename: str) -> bool:
        return filename.lower().endswith(".xlsx")

    def _handle_generation(self, chat_id: int, session: ChatSession) -> None:
        if not session.program:
            raise ValueError("Program selection is missing.")
        if not session.powerbi_file_id or not session.powerbi_filename:
            raise ValueError("PowerBI file is missing.")
        if not session.crm_uploads:
            raise ValueError("At least one CRM file is required.")
        if session.program == PROGRAM_A and not session.pivot_name:
            raise ValueError("Pivot table name is required for Program A.")

        default_output = (
            PROGRAM_B_OUTPUT_FILENAME if session.program == PROGRAM_B else PROGRAM_A_OUTPUT_FILENAME
        )
        output_name = _output_filename(session.output_file, default_output)

        self.send_message(chat_id, "Processing files... this may take some time for large uploads.")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            powerbi_path = tmp / _safe_filename(session.powerbi_filename, "powerbi.xlsx")
            self.download_file(session.powerbi_file_id, powerbi_path)

            crm_paths: list[Path] = []
            platforms: list[str] = []
            for index, crm in enumerate(session.crm_uploads, start=1):
                crm_path = tmp / _safe_filename(crm.filename, f"crm_{index}.xlsx")
                self.download_file(crm.file_id, crm_path)
                crm_paths.append(crm_path)
                platforms.append(crm.platform)

            output_path = tmp / output_name
            common_args = {
                "powerbi_report": powerbi_path,
                "crm_files": crm_paths,
                "platforms": platforms,
                "output_file": output_path,
                "powerbi_sheet": session.powerbi_sheet,
                "crm_sheet": session.crm_sheet,
            }
            if session.program == PROGRAM_B:
                program_b_country_report.build_output(**common_args)
            else:
                program_a_report.build_output(**common_args, pivot_name=session.pivot_name or "")

            self.send_document(
                chat_id,
                output_path,
                caption=f"Done. Generated {output_path.name}",
            )

    def handle_message(self, message: dict[str, Any]) -> None:
        chat = message.get("chat") or {}
        chat_id = chat.get("id")
        if not isinstance(chat_id, int):
            return

        text = (message.get("text") or "").strip()
        text_lower = text.lower()

        if text_lower in {"/start", "/help"}:
            session = self.reset_session(chat_id)
            self.send_message(chat_id, self._help_text())
            self._prompt_program(chat_id)
            session.stage = "await_program"
            return

        if text_lower == "/cancel":
            self.reset_session(chat_id)
            self.send_message(chat_id, "Session cancelled.")
            self._prompt_program(chat_id)
            return

        session = self.session(chat_id)

        if text_lower == "/program_a":
            self._set_program(chat_id, session, PROGRAM_A)
            return
        if text_lower == "/program_b":
            self._set_program(chat_id, session, PROGRAM_B)
            return

        if session.stage == "await_program":
            self._prompt_program(chat_id)
            return

        doc_meta = self._document_meta(message)

        if session.stage == "await_powerbi":
            if not doc_meta:
                self._expect_powerbi(chat_id)
                return
            file_id, filename = doc_meta
            if not self._ensure_xlsx(filename):
                self.send_message(chat_id, "PowerBI file must be .xlsx.")
                return
            session.powerbi_file_id = file_id
            session.powerbi_filename = filename
            session.stage = "await_powerbi_sheet"
            self._expect_powerbi_sheet(chat_id)
            return

        if session.stage == "await_powerbi_sheet":
            session.powerbi_sheet = None if _is_skip(text) else (text or None)
            if session.program == PROGRAM_A:
                session.stage = "await_pivot_name"
                self._expect_pivot_name(chat_id)
            else:
                session.stage = "await_crm_or_done"
                self._expect_crm_or_done(chat_id)
            return

        if session.stage == "await_pivot_name":
            if not text:
                self._expect_pivot_name(chat_id)
                return
            session.pivot_name = text
            session.stage = "await_crm_or_done"
            self._expect_crm_or_done(chat_id)
            return

        if session.stage == "await_crm_or_done":
            if text_lower == "/done":
                if not session.crm_uploads:
                    self.send_message(chat_id, "Please upload at least one CRM file before /done.")
                    return
                session.stage = "await_crm_sheet"
                self._expect_crm_sheet(chat_id)
                return

            if not doc_meta:
                self._expect_crm_or_done(chat_id)
                return

            file_id, filename = doc_meta
            if not self._ensure_xlsx(filename):
                self.send_message(chat_id, "CRM file must be .xlsx.")
                return
            session.pending_crm_file_id = file_id
            session.pending_crm_filename = filename
            session.stage = "await_platform_for_crm"
            self._expect_platform(chat_id)
            return

        if session.stage == "await_platform_for_crm":
            if not text:
                self._expect_platform(chat_id)
                return
            if not session.pending_crm_file_id or not session.pending_crm_filename:
                session.stage = "await_crm_or_done"
                self._expect_crm_or_done(chat_id)
                return
            session.crm_uploads.append(
                CrmUpload(
                    file_id=session.pending_crm_file_id,
                    filename=session.pending_crm_filename,
                    platform=text,
                )
            )
            session.pending_crm_file_id = None
            session.pending_crm_filename = None
            session.stage = "await_crm_or_done"
            self.send_message(chat_id, "CRM file + platform saved.")
            self._expect_crm_or_done(chat_id)
            return

        if session.stage == "await_crm_sheet":
            session.crm_sheet = None if _is_skip(text) else (text or None)
            default_output = (
                PROGRAM_B_OUTPUT_FILENAME if session.program == PROGRAM_B else PROGRAM_A_OUTPUT_FILENAME
            )
            session.stage = "await_output_name"
            self._expect_output_name(chat_id, default_output)
            return

        if session.stage == "await_output_name":
            session.output_file = None if _is_skip(text) else (text or None)
            session.stage = "processing"
            try:
                self._handle_generation(chat_id, session)
                self.send_message(chat_id, "Session complete. Send /start to run another report.")
            except Exception as exc:
                self.send_message(chat_id, f"Error: {exc}")
                self.send_message(chat_id, "Send /start to try again.")
            finally:
                self.reset_session(chat_id)
            return

        self.send_message(chat_id, "Send /start to begin.")

    def run(self) -> None:
        self.send_message_to_stdout("Bot started. Waiting for updates...")
        offset = 0
        while True:
            try:
                response = self.http.get(
                    f"{self.base_url}/getUpdates",
                    params={"timeout": POLL_TIMEOUT_SECONDS, "offset": offset},
                    timeout=POLL_TIMEOUT_SECONDS + 10,
                )
                response.raise_for_status()
                payload = response.json()
                if not payload.get("ok"):
                    raise RuntimeError(f"Telegram getUpdates failed: {payload}")

                for update in payload.get("result", []):
                    update_id = update.get("update_id")
                    if isinstance(update_id, int):
                        offset = update_id + 1
                    message = update.get("message")
                    if message:
                        self.handle_message(message)
            except KeyboardInterrupt:
                self.send_message_to_stdout("Bot stopped.")
                return
            except Exception as exc:
                self.send_message_to_stdout(f"Polling error: {exc}")
                time.sleep(RETRY_DELAY_SECONDS)

    @staticmethod
    def send_message_to_stdout(text: str) -> None:
        print(text, flush=True)


def main() -> int:
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    if not token:
        print("Missing TELEGRAM_BOT_TOKEN environment variable.", flush=True)
        return 1

    bot = TelegramReportBot(token)
    bot.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
