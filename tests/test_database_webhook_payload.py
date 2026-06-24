from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from api import generate


def _write_workbook(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


class DatabaseWebhookPayloadTests(unittest.TestCase):
    def test_build_webhook_records_formats_comments_as_pipe_blocks(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "database_input.xlsx"
            headers = [
                "Brand",
                "Account No",
                "Client Name",
                "Customer Status",
                "Country",
                "Current Assigned Agent",
                "Current Agent Office",
                "Last 10 Comments",
            ]
            rows = [
                [
                    "BrandA",
                    1001,
                    "Client A",
                    "Lead",
                    "TR",
                    "Zawar Bh",
                    "IN Office",
                    (
                        "2026-06-23 00:48 - na vm;\\n"
                        "2026-06-22 23:57 - cb na vm;\\n"
                        "2026-06-22 23:56 - 28 years\\n"
                        "Accountant\\n"
                        "No Exp\\n"
                        "Intro and explained all details..."
                    ),
                ],
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)

            self.assertEqual(len(records), 1)
            record = records[0]
            self.assertEqual(record["brand"], "BrandA")
            self.assertEqual(record["account no"], 1001)
            self.assertEqual(record["country"], "TR")
            self.assertEqual(record["Agent"], "Zawar Bh")
            self.assertEqual(record["Current Agent Office"], "IN Office")
            self.assertEqual(
                record["last 10 comments"],
                (
                    "|| 2026-06-23 00:48 | Zawar Bh | na vm; "
                    "|| 2026-06-22 23:57 | Zawar Bh | cb na vm; "
                    "|| 2026-06-22 23:56 | Zawar Bh | 28 years\n"
                    "Accountant\n"
                    "No Exp\n"
                    "Intro and explained all details..."
                ),
            )

    def test_build_webhook_records_keeps_existing_pipe_and_plain_text_blocks(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "database_input.xlsx"
            headers = [
                "Brand",
                "Account No",
                "Client Name",
                "Customer Status",
                "Country",
                "Current Assigned Agent",
                "Current Agent Office",
                "Last 10 Comments",
            ]
            rows = [
                [
                    "BrandA",
                    1001,
                    "Client A",
                    "Lead",
                    "TR",
                    "Luca Na",
                    "TR Office",
                    (
                        "2026-06-22 21:18 | Luca Na | fw to vm;\\n"
                        "No timestamp line should stay;"
                    ),
                ]
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)
            self.assertEqual(len(records), 1)
            self.assertEqual(
                records[0]["last 10 comments"],
                (
                    "|| 2026-06-22 21:18 | Luca Na | fw to vm; "
                    "|| No timestamp line should stay;"
                ),
            )

    def test_build_webhook_records_uses_row_agent_when_dash_format_has_no_agent(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "database_input.xlsx"
            headers = [
                "Brand",
                "Account No",
                "Client Name",
                "Customer Status",
                "Country",
                "Current Assigned Agent",
                "Current Agent Office",
                "Last 10 Comments",
            ]
            rows = [
                [
                    "BrandA",
                    1001,
                    "Client A",
                    "Lead",
                    "TR",
                    "Luca Na",
                    "TR Office",
                    (
                        "2026-06-22 22:50 - navm;\\n"
                        "2026-06-21 20:17 - na;"
                    ),
                ]
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)
            self.assertEqual(len(records), 1)
            self.assertEqual(
                records[0]["last 10 comments"],
                (
                    "|| 2026-06-22 22:50 | Luca Na | navm; "
                    "|| 2026-06-21 20:17 | Luca Na | na;"
                ),
            )

    def test_build_webhook_records_decodes_double_escaped_newlines(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "database_input.xlsx"
            headers = [
                "Brand",
                "Account No",
                "Client Name",
                "Customer Status",
                "Country",
                "Current Assigned Agent",
                "Current Agent Office",
                "Last 10 Comments",
            ]
            rows = [
                [
                    "BrandA",
                    1001,
                    "Client A",
                    "Lead",
                    "TR",
                    "Luca Na",
                    "TR Office",
                    # Includes double-escaped \\n that should become real newlines.
                    "2026-06-22 22:50 - navm;\\\\n2026-06-21 20:17 - na;",
                ]
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)
            self.assertEqual(len(records), 1)
            self.assertEqual(
                records[0]["last 10 comments"],
                (
                    "|| 2026-06-22 22:50 | Luca Na | navm; "
                    "|| 2026-06-21 20:17 | Luca Na | na;"
                ),
            )


if __name__ == "__main__":
    unittest.main()
