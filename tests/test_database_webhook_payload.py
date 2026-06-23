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
    def test_build_webhook_records_includes_agent_and_refines_comments(self) -> None:
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
                        "2026-06-22 21:16 | Luca Na | vm;\n"
                        "2026-06-22 21:16 | Luca Na | In Progress;"
                    ),
                ],
                [
                    "BrandB",
                    2002,
                    "Client B",
                    "Potential",
                    "DE",
                    "Marta K",
                    (
                        "2026-06-22 20:00 | Marta K | Email follow up;\n"
                        "2026-06-22 19:00 | Marta K | Client asked for callback;"
                    ),
                ],
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)

            self.assertEqual(len(records), 1)
            record = records[0]
            self.assertEqual(record["brand"], "BrandB")
            self.assertEqual(record["account no"], 2002)
            self.assertEqual(record["country"], "DE")
            self.assertEqual(record["Agent"], "Marta K")
            self.assertIn("NA", record["last 10 comments"])
            self.assertIn("2026-06-22 19:00 - Client asked for callback;", record["last 10 comments"])

    def test_build_webhook_records_errors_when_all_rows_ignored(self) -> None:
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

            with self.assertRaisesRegex(ValueError, "No usable rows were found"):
                generate._build_database_check_webhook_records(input_path)

    def test_build_webhook_records_ignores_escaped_newline_na_vm_block(self) -> None:
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
                        "2026-06-20 0:12 - na;\\n"
                        "2026-06-19 13:18 - navm;\\n"
                        "2026-06-18 21:50 - na;\\n"
                        "2026-06-18 14:40 - na;\\n"
                        "2026-06-17 22:19 - nadb;\\n"
                        "2026-06-16 15:43 - nadb;\\n"
                        "2026-06-16 8:25 - nadb;\\n"
                        "2026-06-15 10:51 - Na;"
                    ),
                ]
            ]
            _write_workbook(input_path, headers, rows)

            with self.assertRaisesRegex(ValueError, "No usable rows were found"):
                generate._build_database_check_webhook_records(input_path)


if __name__ == "__main__":
    unittest.main()
