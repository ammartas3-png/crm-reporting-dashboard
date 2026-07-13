"""Tests for report form field resolution in api.generate."""

from __future__ import annotations

import tempfile
import unittest
import zipfile
from email.message import Message
from io import BytesIO
from pathlib import Path
from unittest.mock import Mock

import cgi

from api import generate


def _multipart_form(
    fields: dict[str, str],
    files: list[tuple[str, str, bytes]] | None = None,
) -> cgi.FieldStorage:
    boundary = "----WebKitFormBoundaryTest"
    body = BytesIO()
    for name, value in fields.items():
        body.write(
            (
                f"--{boundary}\r\n"
                f'Content-Disposition: form-data; name="{name}"\r\n\r\n'
                f"{value}\r\n"
            ).encode("utf-8")
        )
    for name, filename, data in files or []:
        body.write(
            (
                f"--{boundary}\r\n"
                f'Content-Disposition: form-data; name="{name}"; filename="{filename}"\r\n'
                "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
            ).encode("utf-8")
        )
        body.write(data)
        body.write(b"\r\n")
    body.write(f"--{boundary}--\r\n".encode("utf-8"))
    payload = body.getvalue()
    content_type = f"multipart/form-data; boundary={boundary}"
    headers = Message()
    headers["content-type"] = content_type
    headers["content-length"] = str(len(payload))
    return cgi.FieldStorage(
        fp=BytesIO(payload),
        headers=headers,
        environ={
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": content_type,
            "CONTENT_LENGTH": str(len(payload)),
        },
        keep_blank_values=True,
    )


def _minimal_xlsx() -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types></Types>')
    return buffer.getvalue()


class GenerateFormFieldTests(unittest.TestCase):
    def test_field_text_reads_pivot_name_from_multipart(self) -> None:
        form = _multipart_form(
            {
                "program": "program_c",
                "pivot_name": "ZA July",
                "monthly_comments_upload_id": "upload-123",
            }
        )
        self.assertEqual(generate._field_text(form, "pivot_name"), "ZA July")

    def test_query_fallback_supplies_missing_pivot_name(self) -> None:
        form = _multipart_form({"program": "program_c"})
        handler = Mock(path="/api/generate?pivot_name=ZA%20July", headers={})
        pivot_name = generate._field_text_with_query_fallback(handler, form, "pivot_name")
        self.assertEqual(pivot_name, "ZA July")

    def test_header_fallback_supplies_missing_pivot_name(self) -> None:
        form = _multipart_form({"program": "program_c"})
        handler = Mock(path="/api/generate")
        handler.headers = {"X-Report-Pivot-Name": "ZA%20July"}
        pivot_name = generate._resolve_pivot_name(handler, form)
        self.assertEqual(pivot_name, "ZA July")

    def test_upload_metadata_supplies_missing_pivot_name(self) -> None:
        form = _multipart_form(
            {
                "program": "program_c",
                "monthly_comments_upload_id": "upload-123",
            }
        )
        handler = Mock(path="/api/generate")
        handler.headers = {}
        pivot_name = generate._resolve_pivot_name(
            handler,
            form,
            {"pivot_name": "ZA July", "program": "program_c"},
        )
        self.assertEqual(pivot_name, "ZA July")

    def test_monthly_upload_id_forces_program_c(self) -> None:
        form = _multipart_form(
            {
                "program": "program_a",
                "monthly_comments_upload_id": "upload-123",
            }
        )
        handler = Mock(path="/api/generate", headers={})
        program = generate._resolve_report_program(handler, form, "upload-123", {})
        self.assertEqual(program, generate.PROGRAM_C)

    def test_query_fallback_supplies_program_when_form_missing(self) -> None:
        form = _multipart_form({"pivot_name": "ZA July"})
        handler = Mock(path="/api/generate?program=program_c", headers={})
        program = generate._resolve_report_program(handler, form, "", {})
        self.assertEqual(program, generate.PROGRAM_C)

    def test_discover_crm_indices_finds_uploaded_files(self) -> None:
        form = _multipart_form(
            {"platform_0": "Fintana"},
            [("crm_file_0", "ZA July.xlsx", _minimal_xlsx())],
        )
        self.assertEqual(generate._discover_crm_indices(form), [0])

    def test_collect_crm_uploads_works_without_crm_count_field(self) -> None:
        form = _multipart_form(
            {"platform_0": "Fintana"},
            [("crm_file_0", "ZA July.xlsx", _minimal_xlsx())],
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            crm_files, platforms = generate._collect_crm_uploads(
                form,
                Path(temp_dir),
            )
            self.assertEqual(len(crm_files), 1)
            self.assertEqual(platforms, ["Fintana"])
            self.assertTrue(crm_files[0].exists())

    def test_monthly_upload_id_from_query_string(self) -> None:
        form = _multipart_form({})
        handler = Mock(
            path="/api/generate?monthly_comments_upload_id=upload-abc-123",
            headers={},
        )
        upload_id = generate._resolve_monthly_comments_upload_id(handler, form)
        self.assertEqual(upload_id, "upload-abc-123")

    def test_monthly_upload_id_from_header(self) -> None:
        form = _multipart_form({})
        handler = Mock(path="/api/generate", headers={})
        handler.headers = {
            "X-Report-Monthly-Comments-Upload-Id": "upload-header-123",
        }
        upload_id = generate._resolve_monthly_comments_upload_id(handler, form)
        self.assertEqual(upload_id, "upload-header-123")


if __name__ == "__main__":
    unittest.main()
