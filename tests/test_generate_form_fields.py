"""Tests for report form field resolution in api.generate."""

from __future__ import annotations

import unittest
from email.message import Message
from io import BytesIO
from unittest.mock import Mock

import cgi

from api import generate


def _multipart_form(fields: dict[str, str]) -> cgi.FieldStorage:
    boundary = "----WebKitFormBoundaryTest"
    parts: list[str] = []
    for name, value in fields.items():
        parts.append(
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'
            f"{value}\r\n"
        )
    parts.append(f"--{boundary}--\r\n")
    body = "".join(parts).encode("utf-8")
    content_type = f"multipart/form-data; boundary={boundary}"
    headers = Message()
    headers["content-type"] = content_type
    headers["content-length"] = str(len(body))
    return cgi.FieldStorage(
        fp=BytesIO(body),
        headers=headers,
        environ={
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": content_type,
            "CONTENT_LENGTH": str(len(body)),
        },
        keep_blank_values=True,
    )


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


if __name__ == "__main__":
    unittest.main()
