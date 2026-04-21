#!/usr/bin/env python3
"""
CRM + PowerBI report automation utility.

This script:
1) enriches CRM records with customer numbers from a PowerBI export,
2) supports exact and fuzzy comment matching,
3) produces lead call counts, AFF/Status ratios, and call-frequency tables.
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd
import yaml
from rapidfuzz import fuzz, process


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automate CRM report generation using CRM and PowerBI files."
    )
    parser.add_argument("--crm", required=True, help="Path to CRM export (csv/xlsx).")
    parser.add_argument(
        "--powerbi", required=True, help="Path to PowerBI export (csv/xlsx)."
    )
    parser.add_argument(
        "--config",
        required=True,
        help="Path to YAML config with column mappings and match settings.",
    )
    parser.add_argument(
        "--output",
        default="report_output.xlsx",
        help="Output Excel path. Default: report_output.xlsx",
    )
    return parser.parse_args()


def read_table(path: str) -> pd.DataFrame:
    file_path = Path(path)
    suffix = file_path.suffix.lower()

    if suffix == ".csv":
        return pd.read_csv(file_path)
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(file_path)

    raise ValueError(
        f"Unsupported file type: {suffix}. Use csv/xlsx/xls for {file_path.name}"
    )


def normalize_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def load_config(path: str) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as file:
        config = yaml.safe_load(file) or {}
    return config


def validate_config(config: dict[str, Any]) -> None:
    required_paths = [
        ("columns", dict),
        ("columns.crm", dict),
        ("columns.powerbi", dict),
    ]
    for key_path, expected_type in required_paths:
        node = config
        for part in key_path.split("."):
            if part not in node:
                raise ValueError(f"Missing config section: '{key_path}'")
            node = node[part]
        if not isinstance(node, expected_type):
            raise ValueError(f"Config section '{key_path}' must be {expected_type.__name__}")

    required_crm_cols = ["lead", "customer_no", "comment", "aff", "status"]
    required_power_cols = ["lead", "customer_no", "comment"]

    for col_key in required_crm_cols:
        if col_key not in config["columns"]["crm"]:
            raise ValueError(f"Missing config key: columns.crm.{col_key}")
    for col_key in required_power_cols:
        if col_key not in config["columns"]["powerbi"]:
            raise ValueError(f"Missing config key: columns.powerbi.{col_key}")


def validate_required_columns(df: pd.DataFrame, required_cols: list[str], label: str) -> None:
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"{label} missing required columns: {missing}")


def find_best_match(
    crm_comment_norm: str,
    candidates: pd.DataFrame,
    power_customer_col: str,
    fuzzy_threshold: int,
) -> tuple[Any, float, str]:
    candidates = candidates[
        candidates["comment_norm"].astype(bool) & candidates[power_customer_col].notna()
    ]
    if candidates.empty:
        return None, 0.0, "not_found"

    exact = candidates[candidates["comment_norm"] == crm_comment_norm]
    if not exact.empty:
        customer_no = exact.iloc[0][power_customer_col]
        return customer_no, 100.0, "exact_comment"

    choice_map = {idx: text for idx, text in candidates["comment_norm"].items()}
    match = process.extractOne(
        crm_comment_norm,
        choice_map,
        scorer=fuzz.token_set_ratio,
        score_cutoff=fuzzy_threshold,
    )
    if not match:
        return None, 0.0, "not_found"

    _, score, matched_idx = match
    customer_no = candidates.loc[matched_idx, power_customer_col]
    return customer_no, float(score), "fuzzy_comment"


def enrich_crm_with_customer_numbers(
    crm_df: pd.DataFrame,
    power_df: pd.DataFrame,
    config: dict[str, Any],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    crm_cols = config["columns"]["crm"]
    power_cols = config["columns"]["powerbi"]
    match_cfg = config.get("match", {})

    fuzzy_threshold = int(match_cfg.get("fuzzy_threshold", 80))
    lead_gate = bool(match_cfg.get("lead_gate", True))

    crm_lead_col = crm_cols["lead"]
    crm_customer_col = crm_cols["customer_no"]
    crm_comment_col = crm_cols["comment"]
    power_lead_col = power_cols["lead"]
    power_customer_col = power_cols["customer_no"]
    power_comment_col = power_cols["comment"]

    validate_required_columns(
        crm_df,
        [crm_lead_col, crm_customer_col, crm_comment_col],
        "CRM file",
    )
    validate_required_columns(
        power_df,
        [power_lead_col, power_customer_col, power_comment_col],
        "PowerBI file",
    )

    crm = crm_df.copy()
    power = power_df.copy()

    crm["lead_norm"] = crm[crm_lead_col].map(normalize_text)
    crm["comment_norm"] = crm[crm_comment_col].map(normalize_text)
    power["lead_norm"] = power[power_lead_col].map(normalize_text)
    power["comment_norm"] = power[power_comment_col].map(normalize_text)

    power_by_lead: dict[str, pd.DataFrame] = {
        lead: group.copy() for lead, group in power.groupby("lead_norm", dropna=False)
    }

    match_rows: list[dict[str, Any]] = []
    matched_customers: list[Any] = []
    match_scores: list[float] = []
    match_methods: list[str] = []

    for _, crm_row in crm.iterrows():
        current_customer = crm_row[crm_customer_col]
        if pd.notna(current_customer) and str(current_customer).strip():
            matched_customers.append(current_customer)
            match_scores.append(100.0)
            match_methods.append("already_present")
            match_rows.append(
                {
                    "crm_lead": crm_row[crm_lead_col],
                    "crm_comment": crm_row[crm_comment_col],
                    "resolved_customer_no": current_customer,
                    "score": 100.0,
                    "method": "already_present",
                }
            )
            continue

        crm_comment_norm = crm_row["comment_norm"]
        crm_lead_norm = crm_row["lead_norm"]
        if not crm_comment_norm:
            matched_customers.append(None)
            match_scores.append(0.0)
            match_methods.append("empty_comment")
            match_rows.append(
                {
                    "crm_lead": crm_row[crm_lead_col],
                    "crm_comment": crm_row[crm_comment_col],
                    "resolved_customer_no": None,
                    "score": 0.0,
                    "method": "empty_comment",
                }
            )
            continue

        if lead_gate and crm_lead_norm in power_by_lead:
            candidates = power_by_lead[crm_lead_norm]
        else:
            candidates = power

        customer_no, score, method = find_best_match(
            crm_comment_norm=crm_comment_norm,
            candidates=candidates,
            power_customer_col=power_customer_col,
            fuzzy_threshold=fuzzy_threshold,
        )
        matched_customers.append(customer_no)
        match_scores.append(score)
        match_methods.append(method)

        match_rows.append(
            {
                "crm_lead": crm_row[crm_lead_col],
                "crm_comment": crm_row[crm_comment_col],
                "resolved_customer_no": customer_no,
                "score": score,
                "method": method,
            }
        )

    crm["resolved_customer_no"] = matched_customers
    crm["match_score"] = match_scores
    crm["match_method"] = match_methods
    crm["customer_no_original"] = crm[crm_customer_col]
    crm[crm_customer_col] = crm[crm_customer_col].where(
        crm[crm_customer_col].notna(), crm["resolved_customer_no"]
    )

    match_audit = pd.DataFrame(match_rows)
    return crm, match_audit


def build_summary_tables(crm_enriched: pd.DataFrame, config: dict[str, Any]) -> dict[str, pd.DataFrame]:
    crm_cols = config["columns"]["crm"]
    lead_col = crm_cols["lead"]
    aff_col = crm_cols["aff"]
    status_col = crm_cols["status"]

    validate_required_columns(
        crm_enriched,
        [lead_col, aff_col, status_col],
        "Enriched CRM",
    )

    lead_call_counts = (
        crm_enriched.groupby(lead_col, dropna=False)
        .size()
        .reset_index(name="call_count")
        .sort_values("call_count", ascending=False)
    )

    call_count_distribution = (
        lead_call_counts["call_count"]
        .value_counts()
        .rename_axis("call_count")
        .reset_index(name="lead_count")
        .sort_values("call_count")
    )

    aff_status_counts = (
        crm_enriched.groupby([aff_col, status_col], dropna=False)
        .size()
        .reset_index(name="call_count")
    )
    aff_totals = aff_status_counts.groupby(aff_col)["call_count"].transform("sum")
    aff_status_counts["ratio_within_aff"] = (
        aff_status_counts["call_count"] / aff_totals
    ).round(4)

    lead_aff_status_calls = (
        crm_enriched.groupby([lead_col, aff_col, status_col], dropna=False)
        .size()
        .reset_index(name="call_count")
    )

    call_frequency_by_aff_status = (
        lead_aff_status_calls.groupby([aff_col, status_col, "call_count"], dropna=False)
        .size()
        .reset_index(name="lead_count")
        .sort_values([aff_col, status_col, "call_count"])
    )

    return {
        "lead_call_counts": lead_call_counts,
        "call_count_distribution": call_count_distribution,
        "aff_status_ratios": aff_status_counts,
        "call_frequency_by_aff_status": call_frequency_by_aff_status,
    }


def write_output_excel(
    output_path: str,
    crm_enriched: pd.DataFrame,
    match_audit: pd.DataFrame,
    summaries: dict[str, pd.DataFrame],
) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        crm_enriched.to_excel(writer, sheet_name="crm_enriched", index=False)
        match_audit.to_excel(writer, sheet_name="match_audit", index=False)
        for sheet_name, df in summaries.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)


def main() -> None:
    args = parse_args()
    config = load_config(args.config)
    validate_config(config)

    crm_df = read_table(args.crm)
    power_df = read_table(args.powerbi)

    crm_enriched, match_audit = enrich_crm_with_customer_numbers(
        crm_df=crm_df,
        power_df=power_df,
        config=config,
    )
    summaries = build_summary_tables(crm_enriched=crm_enriched, config=config)
    lead_counts_map = dict(zip(summaries["lead_call_counts"][config["columns"]["crm"]["lead"]], summaries["lead_call_counts"]["call_count"]))
    crm_enriched["lead_call_count"] = crm_enriched[config["columns"]["crm"]["lead"]].map(lead_counts_map)

    write_output_excel(
        output_path=args.output,
        crm_enriched=crm_enriched,
        match_audit=match_audit,
        summaries=summaries,
    )

    resolved_count = crm_enriched["resolved_customer_no"].notna().sum()
    total_rows = len(crm_enriched)
    print(
        f"Done. Output saved to {args.output} | "
        f"Customer number resolved rows: {resolved_count}/{total_rows}"
    )


if __name__ == "__main__":
    main()
