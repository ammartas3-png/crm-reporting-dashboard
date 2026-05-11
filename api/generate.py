"""Vercel serverless endpoint for generating CRM/PowerBI workbooks."""

from __future__ import annotations

import cgi
import base64
import hashlib
import hmac
import json
import os
import re
import sys
import tempfile
import time
import uuid
from email.message import Message
from http.server import BaseHTTPRequestHandler
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlparse
from urllib.request import Request, urlopen

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import program_a_report  # noqa: E402
import program_b_country_report  # noqa: E402


PROGRAM_A = "program_a"
PROGRAM_B = "program_b"
PROGRAM_A_OUTPUT_FILENAME = "crm_powerbi_output.xlsx"
PROGRAM_B_OUTPUT_FILENAME = "crm_country_report.xlsx"
MAX_JSON_BYTES = 1024 * 1024
BLOB_API_URL = "https://vercel.com/api/blob"
BLOB_API_VERSION = "12"
XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ALLOWED_UPLOAD_CONTENT_TYPES = [
    XLSX_CONTENT_TYPE,
    "application/octet-stream",
]


def _read_static_file(filename: str) -> bytes:
    return (PROJECT_ROOT / filename).read_bytes()


def _json_bytes(payload: dict[str, Any]) -> bytes:
    return json.dumps(payload).encode("utf-8")


def _send_json(
    request_handler: BaseHTTPRequestHandler,
    payload: dict[str, Any],
    status: int = 200,
) -> None:
    data = _json_bytes(payload)
    request_handler.send_response(status)
    request_handler.send_header("Content-Type", "application/json; charset=utf-8")
    request_handler.send_header("Content-Length", str(len(data)))
    request_handler.send_header("Access-Control-Allow-Origin", "*")
    request_handler.end_headers()
    request_handler.wfile.write(data)


def _blob_token() -> str:
    token = os.environ.get("BLOB_READ_WRITE_TOKEN", "").strip()
    if not token:
        raise ValueError(
            "Vercel Blob is not configured. Create/connect a Vercel Blob store "
            "for this project so BLOB_READ_WRITE_TOKEN is available."
        )
    return token


def _blob_store_id(token: str) -> str:
    parts = token.split("_")
    if len(parts) < 4 or not parts[3]:
        raise ValueError("Invalid BLOB_READ_WRITE_TOKEN.")
    return parts[3]


def _blob_path(filename: str, prefix: str) -> str:
    safe_name = _safe_filename(filename, "upload.xlsx")
    return f"crm-reporting-dashboard/{prefix}/{uuid.uuid4().hex}/{safe_name}"


def _generate_blob_client_token(pathname: str) -> str:
    read_write_token = _blob_token()
    valid_until_ms = int((time.time() + 60 * 60) * 1000)
    token_payload = {
        "pathname": pathname,
        "access": "public",
        "addRandomSuffix": True,
        "allowedContentTypes": ALLOWED_UPLOAD_CONTENT_TYPES,
        "validUntil": valid_until_ms,
    }

    payload = base64.b64encode(
        json.dumps(token_payload, separators=(",", ":")).encode("utf-8")
    ).decode("utf-8")
    signature = hmac.new(
        read_write_token.encode("utf-8"),
        payload.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()
    encoded_token = base64.b64encode(f"{signature}.{payload}".encode("utf-8")).decode(
        "utf-8"
    )
    return f"vercel_blob_client_{_blob_store_id(read_write_token)}_{encoded_token}"


def _parse_json_request(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    content_length = int(handler.headers.get("content-length", "0") or "0")
    if content_length <= 0:
        raise ValueError("No request data was received.")
    if content_length > MAX_JSON_BYTES:
        raise ValueError("Request metadata is too large.")

    raw = handler.rfile.read(content_length)
    try:
        payload = json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError("Invalid JSON request.") from exc
    if not isinstance(payload, dict):
        raise ValueError("Invalid JSON request.")
    return payload


def _is_blob_url(url: str) -> bool:
    parsed = urlparse(url)
    hostname = parsed.hostname or ""
    return parsed.scheme == "https" and hostname.endswith(".blob.vercel-storage.com")


def _download_blob_file(blob: dict[str, Any], directory: Path, fallback_name: str) -> Path:
    url = str(blob.get("downloadUrl") or blob.get("url") or "")
    if not _is_blob_url(url):
        raise ValueError(f"Invalid Blob URL for {fallback_name}.")

    filename = _safe_filename(
        str(blob.get("filename") or blob.get("pathname") or ""),
        fallback_name,
    )
    path = directory / filename

    request = Request(url)
    try:
        with urlopen(request, timeout=120) as response:
            path.write_bytes(response.read())
    except (HTTPError, URLError) as exc:
        raise ValueError(f"Unable to download {fallback_name} from Vercel Blob.") from exc

    if path.stat().st_size == 0:
        raise ValueError(f"{filename} is empty.")
    return path


def _upload_output_to_blob(output_path: Path, output_filename: str) -> dict[str, Any]:
    token = _blob_token()
    pathname = _blob_path(output_filename, "outputs")
    url = f"{BLOB_API_URL}/?{urlencode({'pathname': pathname})}"
    request = Request(
        url,
        data=output_path.read_bytes(),
        method="PUT",
        headers={
            "authorization": f"Bearer {token}",
            "x-api-version": BLOB_API_VERSION,
            "x-vercel-blob-access": "public",
            "x-content-type": XLSX_CONTENT_TYPE,
            "x-add-random-suffix": "1",
        },
    )
    try:
        with urlopen(request, timeout=120) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, json.JSONDecodeError) as exc:
        raise ValueError("Unable to upload the generated workbook to Vercel Blob.") from exc

    payload["filename"] = output_filename
    return payload


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


def _program_from_form(form: cgi.FieldStorage) -> str:
    program = _field_text(form, "program") or PROGRAM_A
    if program not in {PROGRAM_A, PROGRAM_B}:
        raise ValueError("Please select Report Generator or Bulk Country Reports.")
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


def _program_from_value(value: Any) -> str:
    program = str(value or PROGRAM_A).strip()
    if program not in {PROGRAM_A, PROGRAM_B}:
        raise ValueError("Please select Report Generator or Bulk Country Reports.")
    return program


def _build_from_files(
    *,
    program: str,
    pivot_name: str,
    powerbi_path: Path,
    crm_files: list[Path],
    platforms: list[str],
    output_path: Path,
) -> None:
    common_args = {
        "powerbi_report": powerbi_path,
        "crm_files": crm_files,
        "platforms": platforms,
        "output_file": output_path,
        "powerbi_sheet": None,
        "crm_sheet": None,
    }
    if program == PROGRAM_B:
        program_b_country_report.build_output(**common_args)
    else:
        program_a_report.build_output(
            **common_args,
            pivot_name=pivot_name,
        )


def _handle_blob_token_request(handler: BaseHTTPRequestHandler) -> None:
    try:
        payload = _parse_json_request(handler)
        filename = _safe_filename(str(payload.get("filename") or ""), "upload.xlsx")
        content_type = str(payload.get("contentType") or XLSX_CONTENT_TYPE)
        if content_type not in ALLOWED_UPLOAD_CONTENT_TYPES:
            content_type = XLSX_CONTENT_TYPE
        pathname = _blob_path(filename, "inputs")
        response = {
            "pathname": pathname,
            "clientToken": _generate_blob_client_token(pathname),
            "contentType": content_type,
        }
    except Exception as exc:
        _send_json(handler, {"error": str(exc)}, status=400)
        return

    _send_json(handler, response)


def _handle_blob_generate_request(handler: BaseHTTPRequestHandler) -> None:
    try:
        payload = _parse_json_request(handler)
        program = _program_from_value(payload.get("program"))
        pivot_name = str(payload.get("pivot_name") or "").strip()
        if program == PROGRAM_A and not pivot_name:
            raise ValueError("Pivot table name is required for Report Generator.")

        powerbi_blob = payload.get("powerbi_blob")
        crm_blobs = payload.get("crm_blobs")
        if not isinstance(powerbi_blob, dict):
            raise ValueError("Please upload the PowerBI report.")
        if not isinstance(crm_blobs, list) or not crm_blobs:
            raise ValueError("Please upload at least one CRM file.")

        default_output = (
            PROGRAM_B_OUTPUT_FILENAME if program == PROGRAM_B else PROGRAM_A_OUTPUT_FILENAME
        )
        output_filename = _output_filename(str(payload.get("output_file") or ""), default_output)

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            powerbi_path = _download_blob_file(
                powerbi_blob,
                tmp_path,
                "PowerBI report.xlsx",
            )

            crm_files: list[Path] = []
            platforms: list[str] = []
            for index, item in enumerate(crm_blobs):
                if not isinstance(item, dict):
                    continue
                platform = str(item.get("platform") or "").strip()
                blob = item.get("blob")
                if not platform:
                    raise ValueError(
                        f"Platform name for CRM file #{index + 1} is required."
                    )
                if not isinstance(blob, dict):
                    raise ValueError(f"Please upload CRM file #{index + 1}.")
                crm_files.append(
                    _download_blob_file(
                        blob,
                        tmp_path,
                        f"CRM file #{index + 1}.xlsx",
                    )
                )
                platforms.append(platform)

            if not crm_files:
                raise ValueError("Please upload at least one CRM file.")

            output_path = tmp_path / output_filename
            _build_from_files(
                program=program,
                pivot_name=pivot_name,
                powerbi_path=powerbi_path,
                crm_files=crm_files,
                platforms=platforms,
                output_path=output_path,
            )
            output_blob = _upload_output_to_blob(output_path, output_filename)
    except Exception as exc:
        _send_json(handler, {"error": str(exc)}, status=400)
        return

    _send_json(handler, {"output": output_blob})


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
        path = urlparse(self.path).path
        if path == "/api/blob-token":
            _handle_blob_token_request(self)
            return

        content_type = self.headers.get("content-type", "")
        if content_type.lower().startswith("application/json"):
            _handle_blob_generate_request(self)
            return

        try:
            form = _parse_form(self)
            program = _program_from_form(form)

            pivot_name = _field_text(form, "pivot_name")
            if program == PROGRAM_A and not pivot_name:
                raise ValueError("Pivot table name is required for Report Generator.")

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

                default_output = (
                    PROGRAM_B_OUTPUT_FILENAME
                    if program == PROGRAM_B
                    else PROGRAM_A_OUTPUT_FILENAME
                )
                output_filename = _output_filename(
                    _field_text(form, "output_file"),
                    default_output,
                )
                output_path = tmp_path / output_filename
                _build_from_files(
                    program=program,
                    pivot_name=pivot_name,
                    powerbi_path=powerbi_path,
                    crm_files=crm_files,
                    platforms=platforms,
                    output_path=output_path,
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
