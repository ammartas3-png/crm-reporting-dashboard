from __future__ import annotations

import os
import re
from functools import lru_cache
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

GCC_COUNTRIES = {
    "Saudi Arabia",
    "United Arab Emirates",
    "Qatar",
    "Kuwait",
    "Bahrain",
    "Oman",
}

NAME_FIXES_SPREADSHEET_ID = "1wqF8cCsPI8jFcxQ-nwMRGzwvcxP-ynfJyOXGSJEhk-E"
NAME_FIXES_SHEET_NAME = "Lead Splitter"
NAME_FIXES_SERVICE_ACCOUNT_EMAIL = "matservice@mitservice.iam.gserviceaccount.com"
NAME_FIXES_ENV_KEYS = ("gmail", "GMAIL", "GOOGLE_PRIVATE_KEY")
NAME_FIXES_FETCH_TIMEOUT_SECONDS = float(
    os.environ.get("LEAD_SPLITTER_NAME_FIX_TIMEOUT_SECONDS", "1.25")
)


def make_border(color: str = "D0D0D0") -> Border:
    side = Side(style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def _get_google_private_key() -> str:
    for env_key in NAME_FIXES_ENV_KEYS:
        private_key = os.environ.get(env_key, "").strip()
        if private_key:
            return private_key.replace("\\n", "\n")
    raise ValueError(
        "Missing Google private key. Set environment variable 'gmail' with the service account private key."
    )


def load_name_data_from_sheet() -> tuple[dict[str, str], set[str]]:
    import httplib2
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as build_google_service
    from google_auth_httplib2 import AuthorizedHttp

    credentials_info = {
        "type": "service_account",
        "client_email": NAME_FIXES_SERVICE_ACCOUNT_EMAIL,
        "private_key": _get_google_private_key(),
        "token_uri": "https://oauth2.googleapis.com/token",
    }
    credentials = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    service = build_google_service(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )
    request = service.spreadsheets().values().get(
        spreadsheetId=NAME_FIXES_SPREADSHEET_ID,
        range=f"{NAME_FIXES_SHEET_NAME}!A:D",
    )
    auth_http = AuthorizedHttp(
        credentials,
        http=httplib2.Http(timeout=NAME_FIXES_FETCH_TIMEOUT_SECONDS),
    )
    response = request.execute(http=auth_http, num_retries=0)
    rows = response.get("values", [])
    if not rows:
        return {}, set()

    header = [str(item).strip().casefold() for item in rows[0]]
    wrong_idx = header.index("wrong names") if "wrong names" in header else 0
    right_idx = header.index("right names") if "right names" in header else 1
    eng_agents_idx = header.index("eng agents") if "eng agents" in header else 3

    fixes: dict[str, str] = {}
    eng_agents: set[str] = set()
    for row in rows[1:]:
        wrong_name = row[wrong_idx].strip() if len(row) > wrong_idx else ""
        right_name = row[right_idx].strip() if len(row) > right_idx else ""
        if wrong_name and right_name:
            fixes[wrong_name] = right_name

        eng_agent_name = row[eng_agents_idx].strip() if len(row) > eng_agents_idx else ""
        if eng_agent_name:
            eng_agents.add(eng_agent_name)
    return fixes, eng_agents


@lru_cache(maxsize=1)
def _cached_name_data() -> tuple[tuple[tuple[str, str], ...], tuple[str, ...]]:
    name_fixes, eng_agents = load_name_data_from_sheet()
    return tuple(name_fixes.items()), tuple(sorted(eng_agents))


def get_name_data() -> tuple[dict[str, str], set[str]]:
    try:
        name_fix_items, eng_agents = _cached_name_data()
        return dict(name_fix_items), set(eng_agents)
    except Exception:
        return {}, set()


def clean_agent_name(name):
    if not isinstance(name, str):
        return name
    name = re.sub(r"\s*\([^)]*\)\s*$", "", name).strip()
    name = re.sub(r"(CY|AE|IN)$", "", name).strip()
    return name


def is_pool_agent(name) -> bool:
    if not isinstance(name, str):
        return False
    clean = name.strip()
    return clean.startswith("BI pool") or clean.startswith("Pool")


def fix_name(name, name_fixes: dict[str, str]):
    return name_fixes.get(name, name)


def get_desk2(desk) -> str:
    if not isinstance(desk, str):
        return str(desk)
    parts = desk.split("-")
    return parts[1] if len(parts) >= 2 else desk


def get_office(desk) -> str:
    if not isinstance(desk, str) or len(desk) < 2:
        return ""
    return str(desk)[:2].upper()


def _pivot_lane_for_desk(desk_name: str) -> str:
    desk_upper = str(desk_name or "").strip().upper()
    if desk_upper.startswith("FR"):
        return "right"
    if desk_upper == "EN" or (desk_upper.startswith("EN") and desk_upper.endswith("SK")):
        return "middle"
    return "left"


def _rgb_interp(start: tuple[int, int, int], end: tuple[int, int, int], t: float) -> str:
    t = max(0.0, min(1.0, t))
    r = round(start[0] + (end[0] - start[0]) * t)
    g = round(start[1] + (end[1] - start[1]) * t)
    b = round(start[2] + (end[2] - start[2]) * t)
    return f"{r:02X}{g:02X}{b:02X}"


def _cr_fill_for_ratio(cr_value: float) -> PatternFill:
    # Mirror Excel-style red->yellow->green color scale.
    if cr_value <= 0.1:
        color = _rgb_interp((248, 105, 107), (255, 235, 132), cr_value / 0.1)
    else:
        color = _rgb_interp((255, 235, 132), (99, 190, 123), (cr_value - 0.1) / 0.1)
    return PatternFill("solid", start_color=color, end_color=color)


def _flag_is_one(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().eq("1").astype(int)


def _cr(leads, ftd) -> float:
    # Keep CR flexible when leads are zero: 1 FTD -> 100%, 2 FTD -> 200%, etc.
    if leads > 0:
        return ftd / leads
    return float(ftd) if ftd > 0 else 0.0


def build_pivot(wb, df, n_col, o_col, b_col, c_col, i_col) -> None:
    ws = wb.create_sheet("Pivot")

    df = df.copy()
    df["_DESK2"] = df[b_col].apply(get_desk2)
    df["_N1"] = _flag_is_one(df[n_col])
    df["_O1"] = _flag_is_one(df[o_col])

    agg = df.groupby(["_DESK2", i_col, c_col], sort=True).agg(
        Assigned=("_N1", "sum"), FTD=("_O1", "sum")
    ).reset_index()
    header_fill = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    country_total_fill = PatternFill("solid", start_color="D9EAF7", end_color="D9EAF7")
    desk_total_fill = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
    white_fill = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    bold_font = Font(bold=True, color="000000", name="Arial", size=10)
    normal_font = Font(bold=False, color="000000", name="Arial", size=10)
    black_border = make_border("000000")

    section_specs = {
        "left": {"start_col": 1, "country_header": "Country"},
        "middle": {"start_col": 8, "country_header": "Country"},
        "right": {"start_col": 15, "country_header": "New Country"},
    }
    lane_rows = {lane: 2 for lane in section_specs}

    def write_header(start_col: int, country_header: str) -> None:
        headers = ["DESK", country_header, "Agent", "Leads", "FTD", "CR"]
        for offset, header in enumerate(headers):
            cell = ws.cell(row=1, column=start_col + offset, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = black_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 20

    def write_row(
        start_col: int,
        row_i: int,
        desk_val: str,
        country_val: str,
        agent_val: str,
        leads: int,
        ftd: int,
        *,
        is_total: bool = False,
        total_fill: PatternFill | None = None,
    ) -> None:
        cr_value = _cr(leads, ftd)
        values = [desk_val, country_val, agent_val, leads, ftd, cr_value]
        row_fill = total_fill if is_total and total_fill is not None else white_fill
        for offset, value in enumerate(values):
            col_i = start_col + offset
            cell = ws.cell(row=row_i, column=col_i, value=value)
            cell.border = black_border
            cell.font = bold_font if is_total else normal_font

            if offset <= 2:
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if offset == 5:
                cell.number_format = "0%"

            if (not is_total) and offset == 5:
                cell.fill = _cr_fill_for_ratio(cr_value)
            else:
                cell.fill = row_fill

        ws.row_dimensions[row_i].height = 16

    for lane_name, lane_spec in section_specs.items():
        write_header(lane_spec["start_col"], lane_spec["country_header"])

    desks = sorted(agg["_DESK2"].dropna().unique().tolist(), key=lambda x: str(x))
    desks_by_lane: dict[str, list[str]] = {"left": [], "middle": [], "right": []}
    for desk in desks:
        desks_by_lane[_pivot_lane_for_desk(str(desk))].append(str(desk))

    for lane_name, lane_desks in desks_by_lane.items():
        start_col = section_specs[lane_name]["start_col"]
        current_row = lane_rows[lane_name]

        for desk_index, desk_name in enumerate(lane_desks):
            desk_df = agg[agg["_DESK2"] == desk_name].copy()
            country_totals = desk_df.groupby(i_col)["Assigned"].sum().sort_values(ascending=False)
            country_order = country_totals.index.tolist()
            first_desk_row = True

            for country in country_order:
                country_df = desk_df[desk_df[i_col] == country].copy()
                country_df = country_df.sort_values(
                    by=["Assigned", "FTD", c_col],
                    ascending=[False, False, True],
                )
                first_country_row = True

                for _, item in country_df.iterrows():
                    write_row(
                        start_col,
                        current_row,
                        desk_name if first_desk_row else "",
                        str(country) if first_country_row else "",
                        str(item[c_col]) if pd.notna(item[c_col]) else "",
                        int(item["Assigned"]),
                        int(item["FTD"]),
                    )
                    first_desk_row = False
                    first_country_row = False
                    current_row += 1

                country_assigned = int(country_df["Assigned"].sum())
                country_ftd = int(country_df["FTD"].sum())
                write_row(
                    start_col,
                    current_row,
                    "",
                    f"{country} Total",
                    "",
                    country_assigned,
                    country_ftd,
                    is_total=True,
                    total_fill=country_total_fill,
                )
                current_row += 1

            desk_assigned = int(desk_df["Assigned"].sum())
            desk_ftd = int(desk_df["FTD"].sum())
            write_row(
                start_col,
                current_row,
                f"{desk_name} Total",
                "",
                "",
                desk_assigned,
                desk_ftd,
                is_total=True,
                total_fill=desk_total_fill,
            )
            current_row += 1

            # Keep one fully empty separator row between desk blocks.
            if desk_index < len(lane_desks) - 1:
                current_row += 1

        lane_rows[lane_name] = current_row

    column_widths = {
        1: 12,
        2: 22,
        3: 17,
        4: 8,
        5: 6,
        6: 6,
        7: 3,
        8: 12,
        9: 22,
        10: 17,
        11: 8,
        12: 6,
        13: 6,
        14: 3,
        15: 12,
        16: 22,
        17: 17,
        18: 8,
        19: 6,
        20: 6,
    }
    for col_index, width in column_widths.items():
        ws.column_dimensions[get_column_letter(col_index)].width = width


AFF_BORDER = make_border("C8C8C8")
AFF_BORDER_DARK = make_border("0D2340")
AFF_FILL_HEADER = PatternFill("solid", start_color="17375E", end_color="17375E")
AFF_FILL_CAMP = PatternFill("solid", start_color="2E75B6", end_color="2E75B6")
AFF_FILL_OFFICE = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
AFF_FILL_ALT = PatternFill("solid", start_color="F5FBFF", end_color="F5FBFF")
AFF_FILL_NONE = PatternFill(fill_type=None)
AFF_FILL_GRAND = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
AFF_FONT_HDR = Font(bold=True, color="FFFFFF", name="Arial", size=10)
AFF_FONT_BOLD = Font(bold=True, name="Arial", size=10)
AFF_FONT_BOLD_W = Font(bold=True, color="FFFFFF", name="Arial", size=10)
AFF_FONT_NORM = Font(bold=False, name="Arial", size=10)


def _aff_write_headers(ws, col_offset, headers) -> None:
    for j, header in enumerate(headers):
        cell = ws.cell(row=1, column=col_offset + j + 1, value=header)
        cell.font = AFF_FONT_HDR
        cell.fill = AFF_FILL_HEADER
        cell.border = AFF_BORDER_DARK
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 20


def _share(part, whole) -> float:
    return (part / whole) if whole > 0 else 0.0


def _aff_write_row(ws, row_i, col_offset, vals, font, fill, skip_fill=0) -> None:
    for j, val in enumerate(vals):
        cell = ws.cell(row=row_i, column=col_offset + j + 1, value=val)
        n_cols = len(vals)
        pct_col = n_cols - 1
        num_start = n_cols - 3
        if j < skip_fill:
            cell.font = AFF_FONT_NORM
            cell.fill = AFF_FILL_NONE
        else:
            cell.font = font
            cell.fill = fill
        cell.border = AFF_BORDER
        if j >= num_start:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        if j == pct_col:
            cell.number_format = "0%"
    ws.row_dimensions[row_i].height = 15


def _write_standard_table(ws, data_df, col_offset, country_label, campaign_col, status_col) -> None:
    _aff_write_headers(
        ws,
        col_offset,
        ["Country", "Office", "Campaign", "Status", "LEADS", "FTD", "CR%"],
    )
    row_i = 2
    first_country = True

    for office, office_df in data_df.groupby("_OFFICE", sort=True):
        first_office = True
        for campaign, campaign_df in office_df.groupby(campaign_col, sort=True):
            first_campaign = True
            alt = 0
            campaign_leads = int(campaign_df["_N1"].sum())

            status_agg = (
                campaign_df.groupby(status_col, sort=True)
                .agg(Leads=("_N1", "sum"), FTD=("_O1", "sum"))
                .reset_index()
            )

            for status_value, leads_value, ftd_value in status_agg.itertuples(
                index=False, name=None
            ):
                leads = int(leads_value)
                ftd = int(ftd_value)
                vals = [
                    country_label if first_country else "",
                    office if first_office else "",
                    campaign if first_campaign else "",
                    status_value,
                    leads,
                    ftd,
                    _share(leads, campaign_leads),
                ]
                _aff_write_row(
                    ws,
                    row_i,
                    col_offset,
                    vals,
                    AFF_FONT_NORM,
                    AFF_FILL_ALT if alt % 2 == 0 else AFF_FILL_NONE,
                )
                first_country = False
                first_office = False
                first_campaign = False
                alt += 1
                row_i += 1

            campaign_ftd = int(campaign_df["_O1"].sum())
            _aff_write_row(
                ws,
                row_i,
                col_offset,
                ["", "", f"{campaign} Total", "", campaign_leads, campaign_ftd, _cr(campaign_leads, campaign_ftd)],
                AFF_FONT_BOLD_W,
                AFF_FILL_CAMP,
                skip_fill=2,
            )
            row_i += 1

        office_leads = int(office_df["_N1"].sum())
        office_ftd = int(office_df["_O1"].sum())
        _aff_write_row(
            ws,
            row_i,
            col_offset,
            ["", f"{office} Total", "", "", office_leads, office_ftd, _cr(office_leads, office_ftd)],
            AFF_FONT_BOLD,
            AFF_FILL_OFFICE,
        )
        row_i += 1

    grand_leads = int(data_df["_N1"].sum())
    grand_ftd = int(data_df["_O1"].sum())
    _aff_write_row(
        ws,
        row_i,
        col_offset,
        [f"{country_label} Total", "", "", "", grand_leads, grand_ftd, _cr(grand_leads, grand_ftd)],
        AFF_FONT_BOLD_W,
        AFF_FILL_GRAND,
    )


def _write_gcc_table(ws, data_df, col_offset, campaign_col, country_col, status_col) -> None:
    _aff_write_headers(
        ws,
        col_offset,
        ["Country", "Office", "Campaign", "Country", "Status", "LEADS", "FTD", "CR%"],
    )
    row_i = 2
    first_region = True

    for office, office_df in data_df.groupby("_OFFICE", sort=True):
        first_office = True
        for campaign, campaign_df in office_df.groupby(campaign_col, sort=True):
            first_campaign = True
            campaign_leads = int(campaign_df["_N1"].sum())

            for country, country_df in campaign_df.groupby(country_col, sort=True):
                first_country = True
                alt = 0

                status_agg = (
                    country_df.groupby(status_col, sort=True)
                    .agg(Leads=("_N1", "sum"), FTD=("_O1", "sum"))
                    .reset_index()
                )

                for status_value, leads_value, ftd_value in status_agg.itertuples(
                    index=False, name=None
                ):
                    leads = int(leads_value)
                    ftd = int(ftd_value)
                    vals = [
                        "GCC EN" if first_region else "",
                        office if first_office else "",
                        campaign if first_campaign else "",
                        country if first_country else "",
                        status_value,
                        leads,
                        ftd,
                        _share(leads, campaign_leads),
                    ]
                    _aff_write_row(
                        ws,
                        row_i,
                        col_offset,
                        vals,
                        AFF_FONT_NORM,
                        AFF_FILL_ALT if alt % 2 == 0 else AFF_FILL_NONE,
                    )
                    first_region = False
                    first_office = False
                    first_campaign = False
                    first_country = False
                    alt += 1
                    row_i += 1

            campaign_ftd = int(campaign_df["_O1"].sum())
            _aff_write_row(
                ws,
                row_i,
                col_offset,
                ["", "", f"{campaign} Total", "", "", campaign_leads, campaign_ftd, _cr(campaign_leads, campaign_ftd)],
                AFF_FONT_BOLD_W,
                AFF_FILL_CAMP,
                skip_fill=2,
            )
            row_i += 1

        office_leads = int(office_df["_N1"].sum())
        office_ftd = int(office_df["_O1"].sum())
        _aff_write_row(
            ws,
            row_i,
            col_offset,
            ["", f"{office} Total", "", "", "", office_leads, office_ftd, _cr(office_leads, office_ftd)],
            AFF_FONT_BOLD,
            AFF_FILL_OFFICE,
        )
        row_i += 1

    grand_leads = int(data_df["_N1"].sum())
    grand_ftd = int(data_df["_O1"].sum())
    _aff_write_row(
        ws,
        row_i,
        col_offset,
        ["GCC EN Total", "", "", "", "", grand_leads, grand_ftd, _cr(grand_leads, grand_ftd)],
        AFF_FONT_BOLD_W,
        AFF_FILL_GRAND,
    )


def build_aff_by_status(df, output_path, campaign_col, country_col, desk_col, status_col, n_col, o_col) -> None:
    data = df.copy()
    data["_OFFICE"] = data[desk_col].apply(get_office)
    data["_DESK2"] = data[desk_col].apply(get_desk2)
    data["_N1"] = _flag_is_one(data[n_col])
    data["_O1"] = _flag_is_one(data[o_col])
    # AFF rule: any row with O column = 1 is treated as Telemarketing.
    data.loc[data["_O1"] == 1, status_col] = "Telemarketing"

    ch_df = data[data[country_col].apply(lambda x: str(x).strip() == "Switzerland")].copy()
    sg_df = data[data[country_col].apply(lambda x: str(x).strip() == "Singapore")].copy()
    gcc_df = data[
        data[country_col].apply(lambda x: str(x).strip() in GCC_COUNTRIES) & (data["_DESK2"] == "EN")
    ].copy()

    wb = Workbook()
    ws = wb.active
    ws.title = "AFF by Status"

    _write_standard_table(
        ws,
        ch_df,
        col_offset=0,
        country_label="CH",
        campaign_col=campaign_col,
        status_col=status_col,
    )
    _write_standard_table(
        ws,
        sg_df,
        col_offset=8,
        country_label="SG",
        campaign_col=campaign_col,
        status_col=status_col,
    )
    _write_gcc_table(
        ws,
        gcc_df,
        col_offset=16,
        campaign_col=campaign_col,
        country_col=country_col,
        status_col=status_col,
    )

    col_widths = [
        10,
        9,
        22,
        18,
        8,
        7,
        7,
        1.5,
        10,
        9,
        22,
        18,
        8,
        7,
        7,
        1.5,
        10,
        9,
        22,
        20,
        18,
        8,
        7,
        7,
    ]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = "A2"
    wb.save(output_path)


def build_outputs(
    input_path: Path,
    output_dir: Path,
    lead_output_name: str | None = None,
    aff_output_name: str | None = None,
    generate_lead: bool = True,
    generate_aff: bool = True,
) -> dict[str, Path]:
    today = datetime.now()
    output_dir.mkdir(parents=True, exist_ok=True)

    lead_name = lead_output_name or f"Lead Splitter - {today.strftime('%d-%m')}.xlsx"
    lead_output_path = output_dir / lead_name

    name_fixes, eng_agents = get_name_data()

    df = pd.read_excel(input_path, header=2, dtype=str)
    df.columns = [str(c) for c in df.columns]
    cols = df.columns.tolist()

    def col(idx):
        return cols[idx] if idx < len(cols) else None

    b_col = col(1)
    c_col = col(2)
    e_col = col(4)
    f_col = col(5)
    i_col = col(8)
    n_col = col(13)
    o_col = col(14)

    required_columns = {
        "Desk": b_col,
        "Agent": c_col,
        "CID": e_col,
        "Status": f_col,
        "Country": i_col,
        "Assigned": n_col,
        "FTD": o_col,
    }
    missing_required = [label for label, column_name in required_columns.items() if column_name is None]
    if missing_required:
        raise ValueError(
            "Lead Splitter input is missing required columns: "
            + ", ".join(missing_required)
        )

    campaign_col = next((c for c in df.columns if str(c).strip().lower() == "campaign"), None)
    assigned_text = df[n_col].fillna("").astype(str).str.strip()
    ftd_text = df[o_col].fillna("").astype(str).str.strip()
    df = df.loc[~(assigned_text.eq("") & ftd_text.eq(""))].copy()

    cid_key = df[e_col].fillna("").astype(str)
    duplicate_cid_mask = cid_key.duplicated(keep=False)
    assigned_is_one = df[n_col].fillna("").astype(str).str.strip().eq("1")
    ftd_is_one = df[o_col].fillna("").astype(str).str.strip().eq("1")

    # Keep the FTD row for duplicated CIDs and normalize Assigned=1 there.
    df.loc[duplicate_cid_mask & ftd_is_one, n_col] = "1"
    df = df.loc[~(duplicate_cid_mask & assigned_is_one & ~ftd_is_one)].copy()

    agent_raw = df[c_col].fillna("").astype(str).str.strip()
    pool_mask = agent_raw.str.startswith("BI pool") | agent_raw.str.startswith("Pool")
    df = df.loc[~pool_mask].copy()

    cleaned_agents = (
        df[c_col]
        .fillna("")
        .astype(str)
        .str.replace(r"\s*\([^)]*\)\s*$", "", regex=True)
        .str.strip()
        .str.replace(r"(CY|AE|IN)$", "", regex=True)
        .str.strip()
        .replace(name_fixes)
    )
    df[c_col] = cleaned_agents

    country_text = df[i_col].fillna("").astype(str).str.strip()
    agent_text = df[c_col].fillna("").astype(str).str.strip()
    df.loc[country_text.eq("Bangladesh"), b_col] = "TR1-IN"
    df.loc[country_text.eq("Malaysia") & ~agent_text.isin(eng_agents), b_col] = "TR1-MY"

    outputs: dict[str, Path] = {}

    if generate_lead:
        header_row = list(df.columns)

        wb_new = Workbook()
        ws_data = wb_new.active
        ws_data.title = "Data"
        ws_data.append(header_row)

        hdr_fill = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
        hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for cell in ws_data[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = hdr_align

        export_rows = df.where(pd.notna(df), None)

        for row_data in export_rows.itertuples(index=False, name=None):
            ws_data.append(row_data)

        width_sample = export_rows.head(5000)
        for col_idx, column_name in enumerate(df.columns, 1):
            col_values = (
                width_sample.iloc[:, col_idx - 1].dropna().astype(str).str.len()
            )
            max_len = max(
                int(col_values.max()) if not col_values.empty else 0,
                len(str(column_name)),
            )
            ws_data.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)

        ws_data.freeze_panes = "A2"
        build_pivot(wb_new, df, n_col, o_col, b_col, c_col, i_col)
        wb_new.save(lead_output_path)
        outputs["lead"] = lead_output_path

    if generate_aff:
        if campaign_col is None:
            raise ValueError("Campaign column was not found, AFF output cannot be created.")
        aff_name = aff_output_name or f"AFF BY status- SG - CH - GCC - {today.strftime('%d-%m')}.xlsx"
        aff_output_path = output_dir / aff_name
        build_aff_by_status(df, aff_output_path, campaign_col, i_col, b_col, f_col, n_col, o_col)
        outputs["aff"] = aff_output_path

    return outputs


def main() -> None:
    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    input_path = script_dir / "report.xlsx"
    outputs = build_outputs(input_path=input_path, output_dir=script_dir)
    print("\n".join([f"Generated: {path.name}" for path in outputs.values()]))


if __name__ == "__main__":
    main()
