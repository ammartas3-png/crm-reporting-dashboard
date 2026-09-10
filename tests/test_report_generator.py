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
    _extract_day_code,
    build_output_files,
    build_output,
    extract_comments,
    extract_monthly_comments,
    read_powerbi_lookup,
)


def _write_workbook(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


def _write_multi_sheet_workbook(
    path: Path,
    sheets: dict[str, tuple[list[str], list[list[object]]]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for sheet_name, (headers, rows) in sheets.items():
        worksheet = workbook.create_sheet(sheet_name)
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

    def test_extract_comments_keeps_multiline_body(self) -> None:
        raw = (
            "2026-09-10 6:56 | Alain Mo | no exp\n"
            "has id\n"
            "hospitality\n"
            "31\n"
            "gambia\n"
            "2000 AED;\n"
            "2026-09-10 6:43 | Alain Mo | In progress;"
        )
        self.assertEqual(
            extract_comments(raw),
            [
                "In progress",
                "no exp has id hospitality 31 gambia 2000 AED",
            ],
        )

    def test_extract_comments_preserves_semicolon_inside_body(self) -> None:
        raw = (
            "2026-09-10 6:56 | Alain Mo | deposited 2000; will call back tomorrow;\n"
            "2026-09-10 6:43 | Alain Mo | second note;"
        )
        self.assertEqual(
            extract_comments(raw),
            [
                "second note",
                "deposited 2000; will call back tomorrow",
            ],
        )

    def test_extract_comments_preserves_pipe_inside_body(self) -> None:
        raw = "2026-09-10 6:56 | Alain Mo | call at 3|4 pm;"
        self.assertEqual(extract_comments(raw), ["call at 3|4 pm"])

    def test_extract_monthly_comments_strips_timestamp_and_semicolon(self) -> None:
        raw = (
            "2026-07-02 14:13 | v3: pu/ intro/ no exp/ age 37/ hu;\n"
            "2026-07-03 10:00 | NA;\n"
            "2026-07-04 11:00 | VM;"
        )
        comments = extract_monthly_comments(raw)
        self.assertEqual(
            comments,
            [
                "VM",
                "NA",
                "v3: pu/ intro/ no exp/ age 37/ hu",
            ],
        )

    def test_build_output_files_uses_monthly_comments_lookup(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            monthly = root / "monthly_comments.xlsx"
            crm = root / "crm.xlsx"
            output = root / "output.xlsx"

            _write_multi_sheet_workbook(
                monthly,
                {
                    "BrandA": (
                        ["CID", "Call Attempts", "Comments"],
                        [
                            [
                                123,
                                13,
                                (
                                    "2026-07-02 14:13 | v3: pu/ intro/ no exp/ age 37/ hu;\n"
                                    "2026-07-03 10:00 | NA;\n"
                                    "2026-07-04 11:00 | VM;\n"
                                    "2026-07-05 12:00 | Potential;"
                                ),
                            ],
                        ],
                    ),
                },
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
                        "Call Again",
                        "TR",
                        "Campaign A",
                        "Sub A",
                        "Placement A",
                        "Agent 1",
                        "1990-01-01",
                    ],
                ],
            )

            outputs = build_output_files(
                monthly_comments_report=monthly,
                crm_files=[crm],
                platforms=["BrandA"],
                pivot_name="Status Pivot",
                output_file=output,
            )
            self.assertEqual([path.name for path in outputs], ["output.xlsx"])

            workbook = load_workbook(root / "output.xlsx", data_only=False)
            ws = workbook["CRM Output"]
            comments_col = PROGRAM_A_OUTPUT_COLUMNS.index("Comments") + 1
            attempts_col = PROGRAM_A_OUTPUT_COLUMNS.index("Call Attempts") + 1

            self.assertEqual(ws.cell(2, attempts_col).value, 13)
            self.assertEqual(
                ws.cell(2, comments_col).value,
                "NA VM x2 // v3: pu/ intro/ no exp/ age 37/ hu",
            )

    def test_build_output_files_splits_by_department_into_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "My_report.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [[111, "BrandA", "| first ;", 1], [222, "BrandA", "| second ;", 2]],
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
            self.assertEqual(sorted(workbook.sheetnames), ["CY", "TR"])

    def test_extract_day_code_handles_us_and_iso_formats(self) -> None:
        self.assertEqual(_extract_day_code("7/22/2026  3:57:54 PM"), "22")
        self.assertEqual(_extract_day_code("7/2/2026  3:57:54 PM"), "02")
        self.assertEqual(_extract_day_code("2026-05-09"), "09")
        self.assertIsNone(_extract_day_code(""))
        self.assertIsNone(_extract_day_code(None))

    def test_build_output_files_splits_by_days_into_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "My_report.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [[111, "BrandA", "| first ;", 1], [222, "BrandA", "| second ;", 2]],
            )
            _write_workbook(
                crm,
                [*CRM_COLUMNS, "Date of Birth"],
                [
                    [
                        "Lead",
                        111,
                        "7/1/2026  3:57:54 PM",
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
                        "7/22/2026  9:00:00 AM",
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
                ],
            )

            outputs = build_output_files(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                pivot_name="Status Pivot",
                output_file=output,
                separate_m_inhousemedia=False,
                separate_by_days=True,
            )

            self.assertEqual([path.name for path in outputs], ["My_report.xlsx"])
            workbook = load_workbook(root / "My_report.xlsx", data_only=False)
            self.assertEqual(workbook.sheetnames, ["Main Report", "01", "22"])

            status_idx = PROGRAM_A_OUTPUT_COLUMNS.index("Status") + 1
            comments_idx = PROGRAM_A_OUTPUT_COLUMNS.index("Comments") + 1
            cb_idx = PROGRAM_A_OUTPUT_COLUMNS.index("CB") + 1
            id_idx = PROGRAM_A_OUTPUT_COLUMNS.index("ID") + 1

            main = workbook["Main Report"]
            self.assertEqual([cell.value for cell in main[1]], PROGRAM_A_OUTPUT_COLUMNS)

            day = workbook["01"]
            # Non-linked columns stay static; the linked ones are formulas.
            self.assertEqual(day.cell(2, id_idx).value, 111)
            for linked_idx in (status_idx, cb_idx, comments_idx):
                formula = day.cell(2, linked_idx).value
                self.assertTrue(str(formula).startswith("="))
                self.assertIn("'Main Report'!", str(formula))
                self.assertIn("MATCH(111", str(formula))

    def test_build_output_files_days_and_departments_make_separate_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "My_report.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [[111, "BrandA", "| first ;", 1], [222, "BrandA", "| second ;", 2]],
            )
            _write_workbook(
                crm,
                [*CRM_COLUMNS, "Date of Birth"],
                [
                    [
                        "Lead",
                        111,
                        "7/1/2026  3:57:54 PM",
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
                        "7/22/2026  9:00:00 AM",
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
                separate_by_days=True,
            )

            self.assertEqual(
                sorted(path.name for path in outputs),
                ["My_report_CY.xlsx", "My_report_TR.xlsx"],
            )
            tr_workbook = load_workbook(root / "My_report_TR.xlsx", data_only=False)
            self.assertEqual(tr_workbook.sheetnames, ["Main Report", "01"])
            cy_workbook = load_workbook(root / "My_report_CY.xlsx", data_only=False)
            self.assertEqual(cy_workbook.sheetnames, ["Main Report", "22"])

    def test_build_output_files_applies_kyc_comments_to_telemarketing(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "kyc_report.xlsx"

            _write_workbook(powerbi, POWERBI_COLUMNS, [[111, "Fintana", "| x ;", 1]])
            _write_workbook(
                crm,
                [*CRM_COLUMNS, "Date of Birth"],
                [
                    [
                        "Depositor",
                        12345,
                        "2026-07-01",
                        "Jane Doe",
                        "Sales",
                        "",
                        "TR",
                        "Campaign A",
                        "Sub A",
                        "Placement A",
                        "Agent 1",
                        "1990-01-01",
                    ],
                    [
                        "Depositor",
                        999,
                        "2026-07-02",
                        "John Doe",
                        "Sales",
                        "",
                        "TR",
                        "Campaign B",
                        "Sub B",
                        "Placement B",
                        "Agent 2",
                        "1989-09-09",
                    ],
                ],
            )

            lookup = {("12345", "fintana"): "KYC done"}
            build_output_files(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["Fintana"],
                pivot_name="Status Pivot",
                output_file=output,
                separate_m_inhousemedia=False,
                telemarketing_kyc_lookup=lookup,
            )

            workbook = load_workbook(output, data_only=False)
            ws = workbook["CRM Output"]
            id_col = PROGRAM_A_OUTPUT_COLUMNS.index("ID") + 1
            comments_col = PROGRAM_A_OUTPUT_COLUMNS.index("Comments") + 1
            status_col = PROGRAM_A_OUTPUT_COLUMNS.index("Status") + 1

            comment_by_id = {}
            for row_idx in range(2, ws.max_row + 1):
                row_id = ws.cell(row_idx, id_col).value
                if row_id in (None, ""):
                    continue
                comment_by_id[str(row_id)] = (
                    ws.cell(row_idx, status_col).value,
                    ws.cell(row_idx, comments_col).value,
                )

            self.assertEqual(comment_by_id["12345"], ("Telemarketing", "KYC done"))
            self.assertEqual(comment_by_id["999"], ("Telemarketing", "No KYC found"))

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

    def test_build_output_separates_by_days(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "country_output.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [
                    [123, "BrandA", "| NA ;", 0],
                    [456, "BrandA", "| VM ;", 5],
                ],
            )
            _write_workbook(
                crm,
                CRM_COLUMNS,
                [
                    [
                        "Lead",
                        123,
                        "7/1/2026  3:57:54 PM",
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
                        "7/22/2026  9:00:00 AM",
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

            outputs = program_b_country_report.build_output(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                output_file=output,
                separate_by_days=True,
            )

            self.assertEqual([path.name for path in outputs], [output.name])
            workbook = load_workbook(output, data_only=False)
            self.assertEqual(workbook.sheetnames, ["Main Report", "01", "22"])

            status_idx = OUTPUT_COLUMNS.index("Status") + 1
            day = workbook["01"]
            formula = day.cell(2, status_idx).value
            self.assertTrue(str(formula).startswith("="))
            self.assertIn("'Main Report'!", str(formula))

    def test_build_output_days_and_departments_make_separate_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            powerbi = root / "powerbi.xlsx"
            crm = root / "crm.xlsx"
            output = root / "country_output.xlsx"

            _write_workbook(
                powerbi,
                POWERBI_COLUMNS,
                [
                    [123, "BrandA", "| NA ;", 0],
                    [456, "BrandA", "| VM ;", 5],
                ],
            )
            _write_workbook(
                crm,
                CRM_COLUMNS,
                [
                    [
                        "Lead",
                        123,
                        "7/1/2026  3:57:54 PM",
                        "Jane Doe",
                        "HQ / TR1 / EN / Opening / Murat",
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
                        "7/22/2026  9:00:00 AM",
                        "Max Doe",
                        "HQ / CY1 / EN-TR / Opening / Housse",
                        "Call Again",
                        "CY",
                        "Campaign B",
                        "Sub B",
                        "Placement B",
                        "Agent 2",
                    ],
                ],
            )

            outputs = program_b_country_report.build_output(
                powerbi_report=powerbi,
                crm_files=[crm],
                platforms=["BrandA"],
                output_file=output,
                separate_department=True,
                separate_by_days=True,
            )

            self.assertEqual(
                sorted(path.name for path in outputs),
                ["country_output_CY.xlsx", "country_output_TR.xlsx"],
            )
            tr_workbook = load_workbook(root / "country_output_TR.xlsx", data_only=False)
            self.assertEqual(tr_workbook.sheetnames, ["Main Report", "01"])
            cy_workbook = load_workbook(root / "country_output_CY.xlsx", data_only=False)
            self.assertEqual(cy_workbook.sheetnames, ["Main Report", "22"])


if __name__ == "__main__":
    unittest.main()
