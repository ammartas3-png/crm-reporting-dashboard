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
    def test_build_webhook_records_formats_comments_with_agent_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "database_input.xlsx"
            headers = [
                "Brand",
                "Account No",
                "Client Name",
                "Customer Status",
                "Country",
                "Current Assigned Agent",
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
                    (
                        "2026-06-22 21:18 | Luca Na | vm;\n"
                        "2026-06-22 21:16 - In Progress;"
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
            self.assertEqual(record["Agent"], "Luca Na")
            self.assertEqual(
                record["last 10 comments"],
                (
                    "|| 2026-06-22 21:18 | Luca Na | vm;\n"
                    "|| 2026-06-22 21:16 | Luca Na | In Progress;"
                ),
            )

    def test_build_webhook_records_keeps_rows_that_were_previous_filtered(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "database_input.xlsx"
            headers = [
                "Brand",
                "Account No",
                "Client Name",
                "Customer Status",
                "Country",
                "Current Assigned Agent",
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
                    "2026-06-22 21:18 | Luca Na | fw to vm;",
                ]
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)
            self.assertEqual(len(records), 1)
            self.assertEqual(
                records[0]["last 10 comments"],
                "|| 2026-06-22 21:18 | Luca Na | fw to vm;",
            )

    def test_build_webhook_records_formats_escaped_newline_entries(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "database_input.xlsx"
            headers = [
                "Brand",
                "Account No",
                "Client Name",
                "Customer Status",
                "Country",
                "Current Assigned Agent",
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
                    (
                        "2026-06-22 22:50 - navm;\\n"
                        "2026-06-21 20:17 - na;\\n"
                        "plain non timestamp note;"
                    ),
                ]
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)
            self.assertEqual(len(records), 1)
            self.assertEqual(
                records[0]["last 10 comments"],
                (
                    "|| 2026-06-22 22:50 | Luca Na | navm;\n"
                    "|| 2026-06-21 20:17 | Luca Na | na;\n"
                    "plain non timestamp note;"
                ),
            )


if __name__ == "__main__":
    unittest.main()
