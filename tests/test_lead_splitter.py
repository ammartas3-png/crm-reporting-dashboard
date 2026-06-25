from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

import lead_splitter


def _write_input_with_header_row_3(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["meta"] * len(headers))
    worksheet.append(["meta"] * len(headers))
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


class LeadSplitterAffCrTests(unittest.TestCase):
    def test_aff_totals_use_strict_ftd_divided_by_leads(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_path = root / "lead_input.xlsx"
            output_dir = root / "out"

            headers = [
                "Ignore0",
                "Desk",
                "Agent",
                "Ignore3",
                "CID",
                "Status",
                "Campaign",
                "Ignore7",
                "Country",
                "Ignore9",
                "Ignore10",
                "Ignore11",
                "Ignore12",
                "Assigned",
                "FTD",
            ]
            rows = [
                ["", "TR1-EN", "Agent A", "", "CID-1", "No Answer", "Aff A", "", "Switzerland", "", "", "", "", "0", "1"],
            ]
            _write_input_with_header_row_3(input_path, headers, rows)

            outputs = lead_splitter.build_outputs(
                input_path=input_path,
                output_dir=output_dir,
                generate_lead=False,
                generate_aff=True,
            )
            aff_output = outputs["aff"]
            workbook = load_workbook(aff_output, data_only=False)
            sheet = workbook["AFF by Status"]

            def _cr_for_label(label: str, label_column: int, cr_column: int) -> float:
                for row_idx in range(2, sheet.max_row + 1):
                    if str(sheet.cell(row_idx, label_column).value or "").strip() == label:
                        return float(sheet.cell(row_idx, cr_column).value or 0.0)
                self.fail(f"Could not find label '{label}' in column {label_column}.")

            # CH table is written in columns A:G.
            self.assertEqual(_cr_for_label("Aff A Total", 3, 7), 0.0)
            self.assertEqual(_cr_for_label("TR Total", 2, 7), 0.0)
            self.assertEqual(_cr_for_label("CH Total", 1, 7), 0.0)


if __name__ == "__main__":
    unittest.main()
