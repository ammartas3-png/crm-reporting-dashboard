"""Chunked upload endpoint for large monthly comments workbooks."""

from __future__ import annotations

import cgi
import json
import sys
from email.message import Message
from http.server import BaseHTTPRequestHandler
from io import BytesIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from api.file_uploads import (  # noqa: E402
    CHUNK_UPLOAD_MAX_BYTES,
    cleanup_upload,
    save_chunk,
)


def _json_bytes(payload: dict) -> bytes:
    return json.dumps(payload).encode("utf-8")


def _field_text(form: cgi.FieldStorage, name: str) -> str:
    field = form.getfirst(name)
    return "" if field is None else str(field).strip()


def _field_int(form: cgi.FieldStorage, name: str) -> int:
    value = _field_text(form, name)
    if not value:
        raise ValueError(f"{name} is required.")
    return int(value)


def _parse_form(handler: BaseHTTPRequestHandler) -> cgi.FieldStorage:
    content_type = handler.headers.get("content-type", "")
    if not content_type.lower().startswith("multipart/form-data"):
        raise ValueError("Upload form must use multipart/form-data.")

    content_length = int(handler.headers.get("content-length", "0") or "0")
    if content_length <= 0:
        raise ValueError("No upload data was received.")
    if content_length > CHUNK_UPLOAD_MAX_BYTES + 512 * 1024:
        raise ValueError(
            "Upload chunk is too large. The monthly comments file is uploaded in smaller pieces automatically."
        )

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
    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self) -> None:
        upload_id = ""
        try:
            form = _parse_form(self)
            upload_id = _field_text(form, "upload_id")
            chunk_index = _field_int(form, "chunk_index")
            total_chunks = _field_int(form, "total_chunks")
            filename = _field_text(form, "filename") or "monthly_comments.xlsx"
            chunk_field = form["chunk"] if "chunk" in form else None
            if chunk_field is None or not getattr(chunk_field, "file", None):
                raise ValueError("Upload chunk file is required.")

            chunk_bytes = chunk_field.file.read()
            if not chunk_bytes:
                raise ValueError("Upload chunk file is empty.")

            output_path = save_chunk(
                upload_id,
                chunk_index,
                total_chunks,
                filename,
                chunk_bytes,
                pivot_name=_field_text(form, "pivot_name") or None,
                program=_field_text(form, "program") or None,
                crm_count=_field_text(form, "crm_count") or None,
            )
            response = {
                "ok": True,
                "upload_id": upload_id,
                "chunk_index": chunk_index,
                "total_chunks": total_chunks,
                "complete": output_path is not None,
            }
            response_bytes = _json_bytes(response)
            status_code = 200
        except Exception as exc:
            if upload_id:
                try:
                    cleanup_upload(upload_id)
                except Exception:
                    pass
            response_bytes = _json_bytes({"error": str(exc)})
            status_code = 400

        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(response_bytes)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(response_bytes)
