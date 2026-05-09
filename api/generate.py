"""Vercel serverless endpoint for generating CRM/PowerBI workbooks."""

from __future__ import annotations

import cgi
import json
import re
import sys
import tempfile
from email.message import Message
from http.server import BaseHTTPRequestHandler
from io import BytesIO
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from report_generator import build_output  # noqa: E402


DEFAULT_OUTPUT_FILENAME = "crm_powerbi_output.xlsx"
MAX_UPLOAD_BYTES = 45 * 1024 * 1024


def _read_static_file(filename: str) -> bytes:
    return (PROJECT_ROOT / filename).read_bytes()


def _json_bytes(payload: dict[str, Any]) -> bytes:
    return json.dumps(payload).encode("utf-8")


def _field(form: cgi.FieldStorage, name: str) -> cgi.FieldStorage | None:
    if name not in form:
        return None
    value = form[name]
    if isinstance(value, list):
        return value[0] if value else None
    return value


def _field_text(form: cgi.FieldStorage, name: str) -> str:
    value = _field(form, name)
    if value is None:
        return ""
    raw = value.value
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="replace")
    return str(raw or "").strip()


def _optional_text(form: cgi.FieldStorage, name: str) -> str | None:
    value = _field_text(form, name)
    return value or None


def _safe_filename(raw_name: str | None, fallback: str) -> str:
    name = Path(raw_name or fallback).name
    name = re.sub(r"[^A-Za-z0-9_.-]+", "_", name).strip("._")
    return name or fallback


def _output_filename(raw_name: str) -> str:
    name = _safe_filename(raw_name, DEFAULT_OUTPUT_FILENAME)
    if not name.lower().endswith(".xlsx"):
        name = f"{name}.xlsx"
    return name


def _save_upload(
    field: cgi.FieldStorage | None,
    directory: Path,
    fallback_name: str,
) -> Path:
    if field is None or not field.filename:
        raise ValueError(f"Please upload {fallback_name}.")

    filename = _safe_filename(field.filename, fallback_name)
    path = directory / filename
    data = field.file.read()
    if not data:
        raise ValueError(f"{filename} is empty.")
    path.write_bytes(data)
    return path


def _has_upload(field: cgi.FieldStorage | None) -> bool:
    return bool(field is not None and field.filename)


def _parse_form(handler: BaseHTTPRequestHandler) -> cgi.FieldStorage:
    content_type = handler.headers.get("content-type", "")
    if not content_type.lower().startswith("multipart/form-data"):
        raise ValueError("Upload form must use multipart/form-data.")

    content_length = int(handler.headers.get("content-length", "0") or "0")
    if content_length <= 0:
        raise ValueError("No upload data was received.")
    if content_length > MAX_UPLOAD_BYTES:
        raise ValueError("Upload is too large. Please keep the total upload under 45 MB.")

    body = handler.rfile.read(content_length)
    headers = Message()
    headers["content-type"] = content_type
    headers["content-length"] = str(len(body))
    environ = {
        "REQUEST_METHOD": "POST",
        "CONTENT_TYPE": content_type,
        "CONTENT_LENGTH": str(len(body)),
    }
    return cgi.FieldStorage(
        fp=BytesIO(body),
        headers=headers,
        environ=environ,
        keep_blank_values=True,
    )


class handler(BaseHTTPRequestHandler):
    def _send_static_file(self, filename: str, content_type: str) -> None:
        try:
            content = _read_static_file(filename)
        except FileNotFoundError:
            self.send_error(404, "File not found")
            return

        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self) -> None:
        if self.path in {"/", "/index.html"}:
            self._send_static_file("index.html", "text/html; charset=utf-8")
            return

        if self.path == "/styles.css":
            self._send_static_file("styles.css", "text/css; charset=utf-8")
            return

        self.send_error(404, "Not found")

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self) -> None:
        try:
            form = _parse_form(self)
            pivot_name = _field_text(form, "pivot_name")
            if not pivot_name:
                raise ValueError("Pivot table name is required.")

            crm_count_raw = _field_text(form, "crm_count")
            try:
                crm_count = int(crm_count_raw)
            except ValueError as exc:
                raise ValueError("At least one CRM file is required.") from exc

            if crm_count < 1:
                raise ValueError("At least one CRM file is required.")

            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                powerbi_path = _save_upload(
                    _field(form, "powerbi_report"),
                    tmp_path,
                    "PowerBI report",
                )

                crm_files: list[Path] = []
                platforms: list[str] = []
                for index in range(crm_count):
                    crm_field = _field(form, f"crm_file_{index}")
                    platform = _field_text(form, f"platform_{index}")
                    has_upload = _has_upload(crm_field)
                    if not has_upload and not platform:
                        continue
                    if not has_upload:
                        raise ValueError(
                            f"Please upload CRM file #{index + 1}."
                        )
                    if not platform:
                        raise ValueError(
                            f"Platform name for CRM file #{index + 1} is required."
                        )
                    crm_path = _save_upload(
                        crm_field,
                        tmp_path,
                        f"CRM file #{index + 1}",
                    )
                    crm_files.append(crm_path)
                    platforms.append(platform)

                if not crm_files:
                    raise ValueError("Please upload at least one CRM file.")

                output_filename = _output_filename(_field_text(form, "output_file"))
                output_path = tmp_path / output_filename
                build_output(
                    powerbi_report=powerbi_path,
                    crm_files=crm_files,
                    platforms=platforms,
                    pivot_name=pivot_name,
                    output_file=output_path,
                    powerbi_sheet=_optional_text(form, "powerbi_sheet"),
                    crm_sheet=_optional_text(form, "crm_sheet"),
                )
                workbook_bytes = output_path.read_bytes()

        except Exception as exc:
            self.send_response(400)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(_json_bytes({"error": str(exc)}))
            return

        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{output_filename}"',
        )
        self.send_header("Content-Length", str(len(workbook_bytes)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(workbook_bytes)
