from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from report_generator import (
    CRM_COLUMNS,
    OUTPUT_COLUMNS,
    POWERBI_COLUMNS,
    build_output,
    read_powerbi_lookup,
)


def _write_workbook(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


class ReportGeneratorTests(unittest.TestCase):
    def test_build_output_merges_comments_and_call_attempts(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "output.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [
                    [
                        123,
                        "BrandA",
                        "| L1 called client ; | NA ; | VM ; | email follow up ;",
                        3,
                    ],
                ],
            )
            _write_workbook(
                crm,
                CRM_COLUMNS,
                [
                    [
                        "Lead",
                        123,
                        "2026-05-09",
                        "Jane Doe",
                        "Sales",
                        "Potential",
                        "TR",
                        "Campaign A",
                        "Sub A",
                        "Placement A",
                        "Agent 1",
                    ],
                    [
                        "Depositor",
                        456,
                        "2026-05-09",
                        "John Doe",
                        "Sales",
                        "Potential",
                        "TR",
                        "Campaign B",
                        "Sub B",
                        "Placement B",
                        "Agent 2",
                    ],
                ],
            )

            build_output(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                pivot_name="Status Pivot",
                output_file=output,
            )

            workbook = load_workbook(output, data_only=False)
            worksheet = workbook["CRM Output"]
            headers = [cell.value for cell in worksheet[1]]
            self.assertEqual(headers, OUTPUT_COLUMNS)

            comments_col = OUTPUT_COLUMNS.index("Comments") + 1
            attempts_col = OUTPUT_COLUMNS.index("Call Attempts") + 1
            status_col = OUTPUT_COLUMNS.index("Status") + 1

            self.assertEqual(worksheet.cell(2, comments_col).value, "NA VM x2 // called client")
            self.assertEqual(worksheet.cell(2, attempts_col).value, 3)
            self.assertEqual(worksheet.cell(3, status_col).value, "Telemarketing")
            self.assertEqual(worksheet.cell(3, attempts_col).value, 1)

    def test_missing_powerbi_columns_reports_file_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "bad_powerbi.xlsx"
            _write_workbook(path, ["Account No"], [[1]])

            with self.assertRaisesRegex(ValueError, "bad_powerbi.xlsx is missing required columns"):
                read_powerbi_lookup(path)


if __name__ == "__main__":
    unittest.main()
