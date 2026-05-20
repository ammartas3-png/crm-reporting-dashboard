"""Vercel serverless endpoint for generating CRM/PowerBI workbooks."""

from __future__ import annotations

import cgi
import json
import re
import sys
import tempfile
import zipfile
from email.message import Message
from http.server import BaseHTTPRequestHandler
from io import BytesIO
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import cr_maker  # noqa: E402
import lead_splitter  # noqa: E402
import program_a_report  # noqa: E402
import program_b_country_report  # noqa: E402


APP_REPORT = "report"
APP_LEAD_SPLITTER = "lead_splitter"
APP_CR = "cr"
PROGRAM_A = "program_a"
PROGRAM_B = "program_b"
PROGRAM_A_OUTPUT_FILENAME = "crm_powerbi_output.xlsx"
PROGRAM_B_OUTPUT_FILENAME = "crm_country_report.xlsx"
LEAD_ZIP_FILENAME = "lead_splitter_outputs.zip"
MAX_UPLOAD_BYTES = 45 * 1024 * 1024
XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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


def _output_filename(raw_name: str, fallback: str) -> str:
    name = _safe_filename(raw_name, fallback)
    if not name.lower().endswith(".xlsx"):
        name = f"{name}.xlsx"
    return name


def _zip_filename(raw_name: str, fallback: str) -> str:
    name = _safe_filename(raw_name, fallback)
    if not name.lower().endswith(".zip"):
        name = f"{name}.zip"
    return name


def _app_from_form(form: cgi.FieldStorage) -> str:
    app = _field_text(form, "app") or APP_REPORT
    if app not in {APP_REPORT, APP_LEAD_SPLITTER, APP_CR}:
        raise ValueError("Please select Report, Lead Splitter, or CR.")
    return app


def _program_from_form(form: cgi.FieldStorage) -> str:
    program = _field_text(form, "program") or PROGRAM_A
    if program not in {PROGRAM_A, PROGRAM_B}:
        raise ValueError("Please select Program A or Program B.")
    return program


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
            app = _app_from_form(form)

            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                if app == APP_REPORT:
                    program = _program_from_form(form)
                    pivot_name = _field_text(form, "pivot_name")
                    if program == PROGRAM_A and not pivot_name:
                        raise ValueError("Pivot table name is required for Program A.")

                    crm_count_raw = _field_text(form, "crm_count")
                    try:
                        crm_count = int(crm_count_raw)
                    except ValueError as exc:
                        raise ValueError("At least one CRM file is required.") from exc

                    if crm_count < 1:
                        raise ValueError("At least one CRM file is required.")

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
                            raise ValueError(f"Please upload CRM file #{index + 1}.")
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

                    default_output = (
                        PROGRAM_B_OUTPUT_FILENAME
                        if program == PROGRAM_B
                        else PROGRAM_A_OUTPUT_FILENAME
                    )
                    response_filename = _output_filename(
                        _field_text(form, "output_file"),
                        default_output,
                    )
                    output_path = tmp_path / response_filename
                    common_args = {
                        "powerbi_report": powerbi_path,
                        "crm_files": crm_files,
                        "platforms": platforms,
                        "output_file": output_path,
                        "powerbi_sheet": _optional_text(form, "powerbi_sheet"),
                        "crm_sheet": _optional_text(form, "crm_sheet"),
                    }
                    if program == PROGRAM_B:
                        program_b_country_report.build_output(**common_args)
                    else:
                        program_a_report.build_output(
                            **common_args,
                            pivot_name=pivot_name,
                        )
                    response_bytes = output_path.read_bytes()
                    response_content_type = XLSX_CONTENT_TYPE
                elif app == APP_LEAD_SPLITTER:
                    lead_input = _save_upload(
                        _field(form, "lead_input"),
                        tmp_path,
                        "Lead splitter input file",
                    )
                    lead_output_raw = _field_text(form, "lead_output_file")
                    aff_output_raw = _field_text(form, "aff_output_file")

                    lead_output_name = (
                        _output_filename(lead_output_raw, "lead_splitter_output.xlsx")
                        if lead_output_raw
                        else None
                    )
                    aff_output_name = (
                        _output_filename(aff_output_raw, "aff_by_status.xlsx")
                        if aff_output_raw
                        else None
                    )

                    generated_paths = lead_splitter.build_outputs(
                        input_path=lead_input,
                        output_dir=tmp_path,
                        lead_output_name=lead_output_name,
                        aff_output_name=aff_output_name,
                    )
                    response_filename = _zip_filename(
                        _field_text(form, "lead_bundle_file"),
                        LEAD_ZIP_FILENAME,
                    )
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
                        for generated in generated_paths:
                            zip_file.write(generated, arcname=generated.name)
                    response_bytes = zip_buffer.getvalue()
                    response_content_type = "application/zip"
                else:
                    cr_input = _save_upload(
                        _field(form, "cr_input"),
                        tmp_path,
                        "CR input file",
                    )
                    response_filename = _output_filename(
                        _field_text(form, "cr_output_file"),
                        cr_maker.default_output_filename(),
                    )
                    cr_output_path = tmp_path / response_filename
                    cr_maker.process(cr_input, cr_output_path)
                    response_bytes = cr_output_path.read_bytes()
                    response_content_type = XLSX_CONTENT_TYPE

        except Exception as exc:
            self.send_response(400)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(_json_bytes({"error": str(exc)}))
            return

        self.send_response(200)
        self.send_header("Content-Type", response_content_type)
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{response_filename}"',
        )
        self.send_header("Content-Length", str(len(response_bytes)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(response_bytes)
