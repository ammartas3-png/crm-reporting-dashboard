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
    def test_build_webhook_records_normalizes_dash_format_to_pipe(self) -> None:
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
                        "2026-06-22 14:36 - pu said i dont have funds;\n"
                        "2026-06-22 14:40 | Zawar Bh | already normalized;\n"
                        "No timestamp line should stay;"
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
                    "2026-06-22 14:36 | Zawar Bh | pu said i dont have funds;\n"
                    "2026-06-22 14:40 | Zawar Bh | already normalized;\n"
                    "No timestamp line should stay;"
                ),
            )

    def test_build_webhook_records_keeps_non_action_comments(self) -> None:
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
                    "2026-06-22 21:18 - fw to vm;",
                ]
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)
            self.assertEqual(len(records), 1)
            self.assertEqual(
                records[0]["last 10 comments"],
                "2026-06-22 21:18 | Luca Na | fw to vm;",
            )

    def test_build_webhook_records_preserves_escaped_newline_order(self) -> None:
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
                        "2026-06-21 20:17 - na;\\n"
                        "2026-06-20 00:12 - in progress;"
                    ),
                ]
            ]
            _write_workbook(input_path, headers, rows)

            records = generate._build_database_check_webhook_records(input_path)
            self.assertEqual(len(records), 1)
            self.assertEqual(
                records[0]["last 10 comments"],
                (
                    "2026-06-22 22:50 | Luca Na | navm;\n"
                    "2026-06-21 20:17 | Luca Na | na;\n"
                    "2026-06-20 00:12 | Luca Na | in progress;"
                ),
            )


if __name__ == "__main__":
    unittest.main()
