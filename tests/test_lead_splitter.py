from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

import lead_splitter


def _write_lead_input(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["meta"] * len(headers))
    worksheet.append(["meta"] * len(headers))
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


class LeadSplitterByCountriesTests(unittest.TestCase):
    def test_build_outputs_generates_countries_workbook_with_telemarketing(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_path = root / "lead_input.xlsx"
            output_dir = root / "outputs"

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
                ["", "TR1-EN", "Alice", "", "1001", "New", "Camp A", "", "Switzerland", "", "", "", "", "1", "0"],
                ["", "TR1-EN", "Alice", "", "1002", "Call Again", "Camp A", "", "Switzerland", "", "", "", "", "1", "1"],
                ["", "FR1-FR", "Bob", "", "1003", "No Answer", "Camp B", "", "France", "", "", "", "", "1", "0"],
                ["", "TR1-BD", "Chad", "", "1004", "No Answer", "Camp C", "", "India", "", "", "", "", "1", "0"],
            ]
            _write_lead_input(input_path, headers, rows)

            outputs = lead_splitter.build_outputs(
                input_path=input_path,
                output_dir=output_dir,
                generate_lead=False,
                generate_aff=False,
                generate_countries=True,
            )

            self.assertEqual(set(outputs.keys()), {"countries"})
            output_path = outputs["countries"]
            self.assertTrue(output_path.exists())

            workbook = load_workbook(output_path, data_only=False)
            worksheet = workbook["Lead splitter by countries"]
            self.assertEqual(
                [worksheet.cell(1, column).value for column in range(1, 8)],
                ["Desk", "Country", "Campaign", "Status", "Leads", "FTD", "CR"],
            )

            desk_labels: list[str] = []
            for desk_col in (1, 9, 17):
                for row_idx in range(2, worksheet.max_row + 1):
                    value = str(worksheet.cell(row_idx, desk_col).value or "").strip()
                    if value:
                        desk_labels.append(value)
            self.assertFalse(any("BD" in label for label in desk_labels), "BD desk should be merged into IN.")
            self.assertTrue(any(label == "IN" or label.startswith("IN Total") for label in desk_labels))

            telemarketing_found = False
            for status_col in (4, 12, 20):
                leads_col = status_col + 1
                ftd_col = status_col + 2
                cr_col = status_col + 3
                for row_idx in range(2, worksheet.max_row + 1):
                    if str(worksheet.cell(row_idx, status_col).value or "").strip() != "Telemarketing":
                        continue
                    telemarketing_found = True
                    self.assertEqual(worksheet.cell(row_idx, leads_col).value, 1)
                    self.assertEqual(worksheet.cell(row_idx, ftd_col).value, 1)
                    self.assertEqual(float(worksheet.cell(row_idx, cr_col).value or 0), 1.0)
                    break
                if telemarketing_found:
                    break

            self.assertTrue(telemarketing_found, "Expected a Telemarketing status row in countries output.")


if __name__ == "__main__":
    unittest.main()
