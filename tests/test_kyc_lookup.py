from __future__ import annotations

import unittest
from datetime import datetime, timezone

import kyc_lookup


class KycNormalizationTests(unittest.TestCase):
    def test_month_sheet_title_is_uppercase_english(self) -> None:
        self.assertEqual(
            kyc_lookup.month_sheet_title(datetime(2026, 7, 15, tzinfo=timezone.utc)),
            "JULY",
        )
        self.assertEqual(
            kyc_lookup.month_sheet_title(datetime(2026, 1, 2, tzinfo=timezone.utc)),
            "JANUARY",
        )

    def test_normalize_id_strips_acc_prefix_and_trailing_symbol(self) -> None:
        self.assertEqual(kyc_lookup.kyc_row_key("ACC12345#", "Fintana")[0], "12345")
        self.assertEqual(kyc_lookup.kyc_row_key("ACC12345", "Fintana")[0], "12345")
        self.assertEqual(kyc_lookup.kyc_row_key("acc12345*", "Fintana")[0], "12345")
        self.assertEqual(kyc_lookup.kyc_row_key(12345, "Fintana")[0], "12345")
        self.assertEqual(kyc_lookup.kyc_row_key(12345.0, "Fintana")[0], "12345")
        self.assertEqual(kyc_lookup.kyc_row_key("12345", "Fintana")[0], "12345")

    def test_brand_matching_is_case_insensitive(self) -> None:
        self.assertEqual(
            kyc_lookup.kyc_row_key("ACC1", "EnvessaMarkets"),
            kyc_lookup.kyc_row_key("1", "envessamarkets"),
        )

    def test_build_lookup_from_values_uses_columns_b_g_i(self) -> None:
        values = [
            ["A", "ACC12345#", "C", "D", "E", "F", "Fintana", "H", "KYC done"],
            ["A", "ACC67890", "C", "D", "E", "F", "EnvessaMarkets", "H", "Approved"],
            ["A", "", "C", "D", "E", "F", "Fintana", "H", "ignored (no id)"],
        ]
        lookup = kyc_lookup.build_lookup_from_values(values)
        self.assertEqual(lookup[("12345", "fintana")], "KYC done")
        self.assertEqual(lookup[("67890", "envessamarkets")], "Approved")
        self.assertNotIn(("", "fintana"), lookup)

    def test_apply_kyc_comments_fills_matches_and_no_kyc(self) -> None:
        rows = [
            {"Status": "Telemarketing", "ID": 12345, "Platform": "Fintana", "Comments": ""},
            {"Status": "Telemarketing", "ID": 999, "Platform": "Fintana", "Comments": ""},
            {"Status": "Potential", "ID": 1, "Platform": "Fintana", "Comments": "keep"},
        ]
        lookup = {("12345", "fintana"): "KYC done"}
        kyc_lookup.apply_kyc_comments(rows, lookup)

        self.assertEqual(rows[0]["Comments"], "KYC done")
        self.assertEqual(rows[1]["Comments"], kyc_lookup.NO_KYC_FOUND_TEXT)
        self.assertEqual(rows[2]["Comments"], "keep")

    def test_apply_kyc_comments_noop_when_lookup_none(self) -> None:
        rows = [
            {"Status": "Telemarketing", "ID": 12345, "Platform": "Fintana", "Comments": ""},
        ]
        kyc_lookup.apply_kyc_comments(rows, None)
        self.assertEqual(rows[0]["Comments"], "")


if __name__ == "__main__":
    unittest.main()
