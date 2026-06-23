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
                    "2026-06-22 21:18 | Luca Na | switched off;",
                ]
            ]
            _write_workbook(input_path, headers, rows)

            with self.assertRaisesRegex(ValueError, "No usable rows were found"):
                generate._build_database_check_webhook_records(input_path)


if __name__ == "__main__":
    unittest.main()
