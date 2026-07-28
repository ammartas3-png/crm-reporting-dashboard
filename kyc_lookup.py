"""Enrich Telemarketing rows with KYC comments from a monthly Google Sheet.

For every output row whose Status is ``Telemarketing`` the report programs look
the lead up in a Google Sheet (one tab per month, e.g. ``JULY``) by matching the
lead ID against column B and the brand/platform against column G. When both
match, the value from column I is written into the output row's Comments cell.

Column B in the sheet is stored like ``ACC12345`` and may carry a trailing
symbol (e.g. ``ACC12345#``). Matching strips the trailing symbol and the ``ACC``
prefix before comparing to the plain numeric ID used in the CRM output.
"""

from __future__ import annotations

import os
import re
from datetime import datetime, timezone
from typing import Any, Iterable

KYC_SPREADSHEET_ID = "1aODXtjBqEEqfee8W0mpqRZ2SOv3r3Yv4fbDlk1tYSVw"
KYC_SERVICE_ACCOUNT_EMAIL = "matservice@mitservice.iam.gserviceaccount.com"
KYC_PRIVATE_KEY_ENV_KEYS = ("google_secret", "GOOGLE_SECRET")
KYC_SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

NO_KYC_FOUND_TEXT = "No KYC found"
KYC_MONTH_SHEET_NOT_FOUND_WARNING = "the KYC month sheet not found"

# Spreadsheet columns (0-based indexes within an A:I read).
_ID_COLUMN_INDEX = 1  # column B
_BRAND_COLUMN_INDEX = 6  # column G
_COMMENT_COLUMN_INDEX = 8  # column I

MONTH_TITLES = [
    "JANUARY",
    "FEBRUARY",
    "MARCH",
    "APRIL",
    "MAY",
    "JUNE",
    "JULY",
    "AUGUST",
    "SEPTEMBER",
    "OCTOBER",
    "NOVEMBER",
    "DECEMBER",
]


class KycMonthSheetNotFound(Exception):
    """Raised when the current month's tab is missing from the KYC spreadsheet."""


def month_sheet_title(now: datetime | None = None) -> str:
    current = now or datetime.now(timezone.utc)
    return MONTH_TITLES[current.month - 1]


def _normalize_id(value: Any) -> str:
    """Return the plain numeric ID, stripping an ``ACC`` prefix and trailing symbol."""
    if isinstance(value, bool):
        return ""
    if isinstance(value, float):
        value = int(value) if value.is_integer() else value

    text = str(value if value is not None else "").strip()
    if not text:
        return ""

    # Drop a leading ACC prefix (optionally followed by spaces).
    text = re.sub(r"^\s*acc\s*", "", text, flags=re.IGNORECASE)
    # Drop trailing non-alphanumeric symbol(s) at the end of the number.
    text = re.sub(r"[^0-9A-Za-z]+$", "", text)
    # Normalize a trailing ".0" that appears when numbers arrive as floats.
    text = re.sub(r"\.0+$", "", text)
    return text.strip().casefold()


def _normalize_brand(value: Any) -> str:
    return str(value if value is not None else "").strip().casefold()


def kyc_row_key(id_value: Any, brand_value: Any) -> tuple[str, str]:
    return (_normalize_id(id_value), _normalize_brand(brand_value))


def build_lookup_from_values(values: Iterable[list[Any]]) -> dict[tuple[str, str], str]:
    """Build a ``{(id, brand): comment}`` map from raw sheet rows (A:I)."""
    lookup: dict[tuple[str, str], str] = {}
    for row in values or []:
        id_raw = row[_ID_COLUMN_INDEX] if len(row) > _ID_COLUMN_INDEX else ""
        brand_raw = row[_BRAND_COLUMN_INDEX] if len(row) > _BRAND_COLUMN_INDEX else ""
        comment_raw = row[_COMMENT_COLUMN_INDEX] if len(row) > _COMMENT_COLUMN_INDEX else ""

        key = kyc_row_key(id_raw, brand_raw)
        if not key[0]:
            continue
        lookup[key] = str(comment_raw if comment_raw is not None else "").strip()
    return lookup


def apply_kyc_comments(
    rows: list[dict[str, Any]],
    lookup: dict[tuple[str, str], str] | None,
) -> None:
    """Fill Telemarketing rows' Comments from the KYC lookup (in place).

    A matched row receives the lookup value; a Telemarketing row with no match
    receives ``No KYC found``. When ``lookup`` is ``None`` (sheet unavailable)
    the rows are left untouched.
    """
    if lookup is None:
        return

    for row in rows:
        status = str(row.get("Status", "") or "").strip().casefold()
        if status != "telemarketing":
            continue
        key = kyc_row_key(row.get("ID"), row.get("Platform"))
        row["Comments"] = lookup.get(key, NO_KYC_FOUND_TEXT)
        row["_comments_yellow"] = False


def _get_private_key() -> str:
    for env_key in KYC_PRIVATE_KEY_ENV_KEYS:
        raw = os.environ.get(env_key)
        if raw and raw.strip():
            return raw.strip().replace("\\n", "\n")
    raise ValueError(
        "Missing Google private key. Set environment variable 'google_secret' "
        "with the KYC service account private key."
    )


def _build_service() -> Any:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as build_google_service

    service_account_info = {
        "type": "service_account",
        "client_email": KYC_SERVICE_ACCOUNT_EMAIL,
        "private_key": _get_private_key(),
        "token_uri": "https://oauth2.googleapis.com/token",
    }
    credentials = service_account.Credentials.from_service_account_info(
        service_account_info,
        scopes=KYC_SCOPES,
    )
    return build_google_service("sheets", "v4", credentials=credentials, cache_discovery=False)


def fetch_kyc_comment_lookup(
    now: datetime | None = None,
    *,
    spreadsheet_id: str = KYC_SPREADSHEET_ID,
    service: Any | None = None,
) -> dict[tuple[str, str], str]:
    """Read the current month's KYC tab and build the comment lookup.

    Raises :class:`KycMonthSheetNotFound` when the month tab does not exist.
    """
    sheets_service = service or _build_service()
    title = month_sheet_title(now)

    metadata = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(title))",
    ).execute()
    titles = [
        sheet.get("properties", {}).get("title", "")
        for sheet in metadata.get("sheets", [])
    ]
    if title not in titles:
        raise KycMonthSheetNotFound(title)

    quoted = "'" + title.replace("'", "''") + "'"
    response = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{quoted}!A:I",
    ).execute()
    return build_lookup_from_values(response.get("values", []))
