from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

import program_b_country_report
from report_generator import (
    CRM_COLUMNS,
    OUTPUT_COLUMNS,
    POWERBI_COLUMNS,
    PROGRAM_A_OUTPUT_COLUMNS,
    build_output_files,
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
                [*CRM_COLUMNS, "Date of Birth"],
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
                        "1990-01-01",
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
                        "1988-10-10",
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
            self.assertEqual(headers, PROGRAM_A_OUTPUT_COLUMNS)

            comments_col = PROGRAM_A_OUTPUT_COLUMNS.index("Comments") + 1
            attempts_col = PROGRAM_A_OUTPUT_COLUMNS.index("Call Attempts") + 1
            status_col = PROGRAM_A_OUTPUT_COLUMNS.index("Status") + 1
            dob_col = PROGRAM_A_OUTPUT_COLUMNS.index("Date of birth") + 1

            self.assertEqual(worksheet.cell(2, comments_col).value, "NA VM x2 // called client")
            self.assertEqual(worksheet.cell(2, attempts_col).value, 3)
            self.assertEqual(worksheet.cell(3, status_col).value, "Telemarketing")
            self.assertEqual(worksheet.cell(3, attempts_col).value, 1)
            self.assertEqual(str(worksheet.cell(2, dob_col).value), "1990-01-01")
            self.assertTrue(
                any(
                    isinstance(cell.value, str)
                    and "COUNTIF(" in cell.value
                    and "$G$2:$G$3" in cell.value
                    for row in worksheet.iter_rows(
                        min_row=1,
                        max_row=worksheet.max_row,
                        min_col=1,
                        max_col=worksheet.max_column,
                    )
                    for cell in row
                )
            )

    def test_build_output_files_splits_m_inhousemedia_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "My_report.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [
                    [111, "BrandA", "| first ;", 1],
                    [222, "BrandA", "| second ;", 2],
                ],
            )
            _write_workbook(
                crm,
                [*CRM_COLUMNS, "Date of Birth"],
                [
                    [
                        "Lead",
                        111,
                        "2026-05-09",
                        "Jane Doe",
                        "Sales",
                        "Potential",
                        "TR",
                        "Campaign A",
                        "Sub A",
                        "Placement A",
                        "Agent 1",
                        "1990-01-01",
                    ],
                    [
                        "Lead",
                        222,
                        "2026-05-09",
                        "John Doe",
                        "Sales",
                        "Call Again",
                        "TR",
                        "M-Inhousemedia Alpha",
                        "Sub B",
                        "Placement B",
                        "Agent 2",
                        "1989-09-09",
                    ],
                ],
            )

            outputs = build_output_files(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                pivot_name="Status Pivot",
                output_file=output,
            )

            output_names = sorted(path.name for path in outputs)
            self.assertEqual(
                output_names,
                ["My_report_M-Inhousemedia.xlsx", "My_report_general.xlsx"],
            )

            workbook_general = load_workbook(root / "My_report_general.xlsx", data_only=False)
            ws_general = workbook_general["CRM Output"]
            self.assertTrue(
                any(
                    cell.value == "Status Pivot"
                    for row in ws_general.iter_rows(
                        min_row=1,
                        max_row=ws_general.max_row,
                        min_col=1,
                        max_col=ws_general.max_column,
                    )
                    for cell in row
                )
            )
            general_status_cell = next(
                cell
                for row in ws_general.iter_rows(
                    min_row=1,
                    max_row=ws_general.max_row,
                    min_col=1,
                    max_col=ws_general.max_column,
                )
                for cell in row
                if cell.value == "Status Pivot"
            )
            self.assertEqual(general_status_cell.column, 1)
            self.assertTrue(
                any(
                    isinstance(cell.value, str)
                    and cell.value.startswith("=")
                    and "COUNTIFS(" in cell.value
                    for row in ws_general.iter_rows(
                        min_row=1,
                        max_row=ws_general.max_row,
                        min_col=1,
                        max_col=ws_general.max_column,
                    )
                    for cell in row
                )
            )

            workbook_inhouse = load_workbook(root / "My_report_M-Inhousemedia.xlsx", data_only=False)
            ws_inhouse = workbook_inhouse["CRM Output"]
            self.assertTrue(
                any(
                    cell.value == "Status Pivot"
                    for row in ws_inhouse.iter_rows(
                        min_row=1,
                        max_row=ws_inhouse.max_row,
                        min_col=1,
                        max_col=ws_inhouse.max_column,
                    )
                    for cell in row
                )
            )
            self.assertFalse(
                any(
                    isinstance(cell.value, str)
                    and cell.value.startswith("=")
                    and "COUNTIFS(" in cell.value
                    for row in ws_inhouse.iter_rows(
                        min_row=1,
                        max_row=ws_inhouse.max_row,
                        min_col=1,
                        max_col=ws_inhouse.max_column,
                    )
                    for cell in row
                )
            )

            status_pivot_fill = None
            for row in ws_inhouse.iter_rows(
                min_row=1,
                max_row=ws_inhouse.max_row,
                min_col=1,
                max_col=ws_inhouse.max_column,
            ):
                matching = next((cell for cell in row if cell.value == "Status Pivot"), None)
                if matching:
                    status_pivot_fill = matching.fill
                    break
            self.assertIsNotNone(status_pivot_fill)
            self.assertTrue(
                (status_pivot_fill.fgColor.rgb or "").endswith("FFDBB7")
                or (status_pivot_fill.start_color.rgb or "").endswith("FFDBB7")
            )

    def test_build_output_files_keeps_single_file_when_split_toggle_off(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "My_report.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [
                    [111, "BrandA", "| first ;", 1],
                    [222, "BrandA", "| second ;", 2],
                ],
            )
            _write_workbook(
                crm,
                [*CRM_COLUMNS, "Date of Birth"],
                [
                    [
                        "Lead",
                        111,
                        "2026-05-09",
                        "Jane Doe",
                        "Sales",
                        "Potential",
                        "TR",
                        "Campaign A",
                        "Sub A",
                        "Placement A",
                        "Agent 1",
                        "1990-01-01",
                    ],
                    [
                        "Lead",
                        222,
                        "2026-05-09",
                        "John Doe",
                        "Sales",
                        "Call Again",
                        "TR",
                        "M-Inhousemedia Alpha",
                        "Sub B",
                        "Placement B",
                        "Agent 2",
                        "1989-09-09",
                    ],
                ],
            )

            outputs = build_output_files(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                pivot_name="Status Pivot",
                output_file=output,
                separate_m_inhousemedia=False,
            )
            self.assertEqual([path.name for path in outputs], ["My_report.xlsx"])

            workbook = load_workbook(root / "My_report.xlsx", data_only=False)
            ws = workbook["CRM Output"]
            campaign_col = PROGRAM_A_OUTPUT_COLUMNS.index("Campaign") + 1
            campaigns = [
                str(ws.cell(row, campaign_col).value or "")
                for row in range(2, min(ws.max_row, 20) + 1)
            ]
            self.assertIn("Campaign A", campaigns)
            self.assertIn("M-Inhousemedia Alpha", campaigns)

    def test_build_output_files_splits_by_department_code(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "My_report.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [
                    [111, "BrandA", "| first ;", 1],
                    [222, "BrandA", "| second ;", 2],
                    [333, "BrandA", "| third ;", 3],
                ],
            )
            _write_workbook(
                crm,
                [*CRM_COLUMNS, "Date of Birth"],
                [
                    [
                        "Lead",
                        111,
                        "2026-05-09",
                        "Jane Doe",
                        "HQ / TR1 / EN / Opening / Murat",
                        "Potential",
                        "TR",
                        "Campaign A",
                        "Sub A",
                        "Placement A",
                        "Agent 1",
                        "1990-01-01",
                    ],
                    [
                        "Lead",
                        222,
                        "2026-05-09",
                        "John Doe",
                        "HQ / CY1 / EN-TR / Opening / Housse",
                        "Call Again",
                        "CY",
                        "Campaign B",
                        "Sub B",
                        "Placement B",
                        "Agent 2",
                        "1989-09-09",
                    ],
                    [
                        "Lead",
                        333,
                        "2026-05-09",
                        "Alex Doe",
                        "Sales",
                        "Potential",
                        "TR",
                        "Campaign C",
                        "Sub C",
                        "Placement C",
                        "Agent 3",
                        "1988-08-08",
                    ],
                ],
            )

            outputs = build_output_files(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                pivot_name="Status Pivot",
                output_file=output,
                separate_m_inhousemedia=False,
                separate_department=True,
            )

            self.assertEqual([path.name for path in outputs], ["My_report.xlsx"])

            workbook = load_workbook(root / "My_report.xlsx", data_only=False)
            self.assertEqual(sorted(workbook.sheetnames), ["CY", "TR", "general"])

            ws_tr = workbook["TR"]
            department_col = PROGRAM_A_OUTPUT_COLUMNS.index("Department") + 1
            departments = {
                str(ws_tr.cell(row, department_col).value or "")
                for row in range(2, ws_tr.max_row + 1)
                if ws_tr.cell(row, department_col).value
            }
            self.assertEqual(departments, {"HQ / TR1 / EN / Opening / Murat"})

    def test_build_output_files_splits_by_department_and_m_inhousemedia(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "My_report.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [
                    [111, "BrandA", "| first ;", 1],
                    [222, "BrandA", "| second ;", 2],
                ],
            )
            _write_workbook(
                crm,
                [*CRM_COLUMNS, "Date of Birth"],
                [
                    [
                        "Lead",
                        111,
                        "2026-05-09",
                        "Jane Doe",
                        "HQ / TR1 / EN / Opening / Murat",
                        "Potential",
                        "TR",
                        "Campaign A",
                        "Sub A",
                        "Placement A",
                        "Agent 1",
                        "1990-01-01",
                    ],
                    [
                        "Lead",
                        222,
                        "2026-05-09",
                        "John Doe",
                        "HQ / TR1 / EN / Opening / Murat",
                        "Call Again",
                        "TR",
                        "M-Inhousemedia Alpha",
                        "Sub B",
                        "Placement B",
                        "Agent 2",
                        "1989-09-09",
                    ],
                ],
            )

            outputs = build_output_files(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                pivot_name="Status Pivot",
                output_file=output,
                separate_m_inhousemedia=True,
                separate_department=True,
            )

            self.assertEqual(
                sorted(path.name for path in outputs),
                ["My_report_M-Inhousemedia.xlsx", "My_report_general.xlsx"],
            )

            workbook_general = load_workbook(root / "My_report_general.xlsx", data_only=False)
            workbook_inhouse = load_workbook(root / "My_report_M-Inhousemedia.xlsx", data_only=False)
            self.assertEqual(workbook_general.sheetnames, ["TR"])
            self.assertEqual(workbook_inhouse.sheetnames, ["TR"])

    def test_missing_powerbi_columns_reports_file_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "bad_powerbi.xlsx"
            _write_workbook(path, ["Account No"], [[1]])

            with self.assertRaisesRegex(ValueError, "bad_powerbi.xlsx is missing required columns"):
                read_powerbi_lookup(path)


class ProgramBCountryReportTests(unittest.TestCase):
    def test_build_output_creates_main_and_country_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "country_output.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [
                    [123, "BrandA", "| NA ; | VM ;", 0],
                    [456, "BrandA", "| L1 interested ;", 5],
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
                        "Lead",
                        456,
                        "2026-05-09",
                        "Max Doe",
                        "Sales",
                        "Call Again",
                        "DE",
                        "Campaign B",
                        "Sub B",
                        "Placement B",
                        "Agent 2",
                    ],
                ],
            )

            program_b_country_report.build_output(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                output_file=output,
            )

            workbook = load_workbook(output, data_only=False)
            self.assertIn("Main Report", workbook.sheetnames)
            self.assertIn("DE", workbook.sheetnames)
            self.assertIn("TR", workbook.sheetnames)

            main = workbook["Main Report"]
            headers = [cell.value for cell in main[1]]
            self.assertEqual(headers, OUTPUT_COLUMNS)

            attempts_col = OUTPUT_COLUMNS.index("Call Attempts") + 1
            self.assertEqual(main.cell(3, attempts_col).value, 1)
            self.assertEqual(main.cell(2, attempts_col).value, 5)

            de_sheet = workbook["DE"]
            self.assertEqual([cell.value for cell in de_sheet[1]], OUTPUT_COLUMNS)
            self.assertTrue(str(de_sheet.cell(2, 1).value).startswith("=IFERROR("))
            self.assertEqual(de_sheet.cell(5, 6).value, "DE")


if __name__ == "__main__":
    unittest.main()
