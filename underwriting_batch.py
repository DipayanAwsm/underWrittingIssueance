"""
Batch underwriting issuance analysis (no Streamlit).

This script:
- Loads a CSV/Excel underwriting data file
- Performs the same analysis as the Streamlit dashboard / notebook
- Writes an Excel workbook with multiple sheets:
  Summary, Status_Counts, Case_Type_Classification, TAT_Buckets,
  Median_TAT_by_CaseType, Median_TAT_by_Hold_Bucket, Median_TAT_by_HoldCategory,
  Monthly_Statistics, Location_Analysis, Vehicle_Analysis, Holding_Time_Stats,
  Original_Data, Processed_Data

Usage (from project root):

    python underwriting_batch.py \
        --input data/auto_issuance_synthetic_1year_10000rows.csv \
        --output underwriting_report_batch.xlsx
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import List, Any

import numpy as np
import pandas as pd


def parse_separated_values(value: Any, separator: str | None = None) -> List[str]:
    """Parse values that may be separated by ', ' or '|' into a list."""
    if pd.isna(value) or value == "":
        return []

    text = str(value).strip()
    if not text:
        return []

    if separator:
        separators = [separator]
    else:
        if "|" in text:
            separators = ["|"]
        elif ", " in text:
            separators = [", "]
        else:
            return [text]

    for sep in separators:
        if sep in text:
            return [v.strip() for v in text.split(sep) if v.strip()]

    return [text]


def calculate_aging(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate aging in days for non-completed cases."""
    df = df.copy()
    df["createDateTime"] = pd.to_datetime(df.get("createDateTime"), errors="coerce")
    df["completedDateTime"] = pd.to_datetime(df.get("completedDateTime"), errors="coerce")

    mask = df["completedDateTime"].isna()
    df.loc[mask, "Aging_Days"] = (datetime.now() - df.loc[mask, "createDateTime"]).dt.days
    df.loc[~mask, "Aging_Days"] = np.nan
    return df


def classify_case_type(on_hold_reasons: Any) -> str:
    """Classify case as straight-through, one-touch, or multi-hold."""
    reasons = parse_separated_values(on_hold_reasons)
    if len(reasons) == 0:
        return "Straight Through"
    if len(reasons) == 1:
        return "One Touch"
    return f"Multi Hold ({len(reasons)} touches)"


def calculate_holding_times(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate holding time for each onHoldReasonDescriptionsHistory entry."""
    df = df.copy()
    holding_times_list: List[List[int]] = []

    for _, row in df.iterrows():
        on_hold_dates = parse_separated_values(row.get("onHoldDatesHistory", ""))
        off_hold_dates = parse_separated_values(row.get("offHoldDatesHistory", ""))

        on_parsed = [
            pd.to_datetime(d, errors="coerce") for d in on_hold_dates
        ]
        on_parsed = [d for d in on_parsed if pd.notna(d)]

        off_parsed = [
            pd.to_datetime(d, errors="coerce") for d in off_hold_dates
        ]
        off_parsed = [d for d in off_parsed if pd.notna(d)]

        hold_times: List[int] = []
        for i, on_date in enumerate(on_parsed):
            if i < len(off_parsed):
                off_date = off_parsed[i]
                if pd.notna(on_date) and pd.notna(off_date):
                    hold_times.append((off_date - on_date).days)
            else:
                # No off date -> from on_date to now
                if pd.notna(on_date):
                    hold_times.append((datetime.now() - on_date).days)

        holding_times_list.append(hold_times)

    df["HoldingTimes"] = holding_times_list
    df["TotalHoldingTime"] = df["HoldingTimes"].apply(lambda x: sum(x) if x else 0)
    df["NumberOfTouches"] = df["HoldingTimes"].apply(lambda x: len(x) if x else 0)
    return df


def calculate_tat(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate Turnaround Time (TAT) in days for completed cases."""
    df = df.copy()
    df["createDateTime"] = pd.to_datetime(df.get("createDateTime"), errors="coerce")
    df["completedDateTime"] = pd.to_datetime(df.get("completedDateTime"), errors="coerce")

    mask = df["completedDateTime"].notna()
    df.loc[mask, "TAT_Days"] = (
        df.loc[mask, "completedDateTime"] - df.loc[mask, "createDateTime"]
    ).dt.days
    df.loc[~mask, "TAT_Days"] = np.nan
    return df


def create_tat_bucket(tat_days: Any) -> str | None:
    """Create TAT buckets: 0-5, 5-7, 7+ days."""
    if pd.isna(tat_days):
        return None
    if tat_days <= 5:
        return "0-5 days"
    if tat_days <= 7:
        return "5-7 days"
    return "7+ days"


def load_data(path: Path) -> pd.DataFrame:
    """Load CSV or Excel file."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, low_memory=False)
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(path, engine="openpyxl")
    raise ValueError(f"Unsupported file type: {suffix}")


def run_analysis(input_path: Path, output_path: Path) -> None:
    """Run full analysis and write Excel workbook."""
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    print(f"Loading data from: {input_path}")
    original_df = load_data(input_path)
    print(f"Rows loaded: {len(original_df)}")

    # Processed copy
    df = original_df.copy()

    # Core calculations
    df = calculate_aging(df)
    if "onHoldReasonDescriptionsHistory" in df.columns:
        df["CaseType"] = df["onHoldReasonDescriptionsHistory"].apply(classify_case_type)
    else:
        df["CaseType"] = "Unknown"

    df = calculate_holding_times(df)
    df = calculate_tat(df)
    df["TAT_Bucket"] = df["TAT_Days"].apply(create_tat_bucket)

    # Month / seasonality
    if "createDateTime" in df.columns:
        df["createDateTime"] = pd.to_datetime(df["createDateTime"], errors="coerce")
        df["Month"] = df["createDateTime"].dt.to_period("M")
        df["Month_Str"] = df["Month"].astype(str)

    # HoldCategory: Straight / Single / Multi
    df["HoldCategory"] = df["CaseType"].apply(
        lambda x: "Straight Through"
        if x == "Straight Through"
        else "Single Hold"
        if x == "One Touch"
        else "Multi Hold"
    )

    # Derived subsets and stats
    status_counts = (
        df["statusDescription"].value_counts()
        if "statusDescription" in df.columns
        else pd.Series(dtype=int)
    )

    aging_data = df[df["Aging_Days"].notna()].copy()
    case_type_counts = df["CaseType"].value_counts()

    straight_through = df[df["CaseType"] == "Straight Through"]
    one_touch = df[df["CaseType"] == "One Touch"]
    multi_hold = df[df["CaseType"].str.contains("Multi Hold", na=False)]

    holding_data = df[df["TotalHoldingTime"] > 0].copy()
    tat_data = df[df["TAT_Days"].notna()].copy()

    tat_bucket_counts = df["TAT_Bucket"].value_counts(dropna=True)
    tat_by_case_type = df.groupby("CaseType")["TAT_Days"].median().sort_values(ascending=False)

    tat_by_hold_bucket_raw = df.groupby(["HoldCategory", "TAT_Bucket"])["TAT_Days"].median().reset_index()
    if not tat_by_hold_bucket_raw.empty:
        tat_by_hold_bucket = tat_by_hold_bucket_raw.pivot(
            index="TAT_Bucket", columns="HoldCategory", values="TAT_Days"
        )
    else:
        tat_by_hold_bucket = pd.DataFrame()

    tat_by_hold_category = df.groupby("HoldCategory")["TAT_Days"].median().sort_values(ascending=False)

    # Seasonality
    if "Month_Str" in df.columns:
        monthly_stats = df.groupby("Month_Str").agg(
            {
                "requestId": "count",
                "completedDateTime": lambda x: x.notna().sum(),
                "TAT_Days": "median",
                "TotalHoldingTime": "median",
                "NumberOfTouches": "mean",
            }
        ).rename(
            columns={
                "requestId": "Total_Cases",
                "completedDateTime": "Completed_Cases",
                "TAT_Days": "Median_TAT",
                "TotalHoldingTime": "Median_Holding_Time",
                "NumberOfTouches": "Avg_Touches",
            }
        )

        monthly_stats["Completion_Rate"] = (
            monthly_stats["Completed_Cases"] / monthly_stats["Total_Cases"] * 100
        ).round(2)

        holding_rate = df.groupby("Month_Str")["TotalHoldingTime"].apply(
            lambda x: (x > 0).sum() / len(x) * 100
        ).round(2)
        monthly_stats["Holding_Rate"] = holding_rate
    else:
        monthly_stats = pd.DataFrame()

    # Location / vehicle analysis
    if "numberOfLocations" in df.columns and "NumberOfVehicles" in df.columns:
        lv_data = df[
            (df["numberOfLocations"].notna()) & (df["NumberOfVehicles"].notna())
        ].copy()
    else:
        lv_data = pd.DataFrame()

    if not lv_data.empty:
        location_analysis = lv_data.groupby("numberOfLocations").agg(
            {
                "TAT_Days": "median",
                "NumberOfTouches": "mean",
                "TotalHoldingTime": "median",
                "requestId": "count",
            }
        ).rename(
            columns={
                "TAT_Days": "Median_TAT",
                "NumberOfTouches": "Avg_Touches",
                "TotalHoldingTime": "Median_Holding_Time",
                "requestId": "Count",
            }
        )

        vehicle_analysis = lv_data.groupby("NumberOfVehicles").agg(
            {
                "TAT_Days": "median",
                "NumberOfTouches": "mean",
                "TotalHoldingTime": "median",
                "requestId": "count",
            }
        ).rename(
            columns={
                "TAT_Days": "Median_TAT",
                "NumberOfTouches": "Avg_Touches",
                "TotalHoldingTime": "Median_Holding_Time",
                "requestId": "Count",
            }
        )
    else:
        location_analysis = pd.DataFrame()
        vehicle_analysis = pd.DataFrame()

    # Holding time stats
    all_holding_times_flat: List[int] = []
    if not holding_data.empty:
        for times in holding_data["HoldingTimes"]:
            all_holding_times_flat.extend(times)

    highest_per_case = (
        holding_data["HoldingTimes"].apply(lambda x: max(x) if x else 0)
        if not holding_data.empty
        else pd.Series(dtype=float)
    )
    lowest_per_case = (
        holding_data["HoldingTimes"].apply(lambda x: min(x) if x else 0)
        if not holding_data.empty
        else pd.Series(dtype=float)
    )

    # Write Excel
    print(f"Writing Excel report to: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Summary sheet
        summary_data = {
            "Metric": [
                "Total Cases",
                "Completed Cases",
                "Median TAT (days)",
                "Median Aging (days)",
                "Straight Through Cases",
                "One Touch Cases",
                "Multi Hold Cases",
                "Median Holding Time (All)",
                "Median of Highest Holding Time",
                "Median of Lowest Holding Time",
                "Average Number of Touches",
                "Highest Holding Time (Overall)",
                "Lowest Holding Time (Overall)",
            ],
            "Value": [
                len(df),
                df["completedDateTime"].notna().sum()
                if "completedDateTime" in df.columns
                else 0,
                df["TAT_Days"].median()
                if "TAT_Days" in df.columns
                and not df["TAT_Days"].isna().all()
                else 0,
                df["Aging_Days"].median()
                if "Aging_Days" in df.columns
                and not df["Aging_Days"].isna().all()
                else 0,
                len(straight_through),
                len(one_touch),
                len(multi_hold),
                pd.Series(all_holding_times_flat).median()
                if all_holding_times_flat
                else 0,
                highest_per_case.median() if not highest_per_case.empty else 0,
                lowest_per_case.median() if not lowest_per_case.empty else 0,
                holding_data["NumberOfTouches"].mean()
                if not holding_data.empty
                else 0,
                max(all_holding_times_flat) if all_holding_times_flat else 0,
                min(all_holding_times_flat) if all_holding_times_flat else 0,
            ],
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="Summary", index=False)

        # Status counts
        if not status_counts.empty:
            sc_df = (
                status_counts.to_frame("Count")
                .reset_index()
                .rename(columns={"index": "Status"})
            )
            sc_df.to_excel(writer, sheet_name="Status_Counts", index=False)

        # Case type classification
        ctc_df = (
            case_type_counts.to_frame("Count")
            .reset_index()
            .rename(columns={"index": "Case_Type"})
        )
        ctc_df.to_excel(writer, sheet_name="Case_Type_Classification", index=False)

        # TAT buckets
        if not tat_bucket_counts.empty:
            tbc_df = (
                tat_bucket_counts.to_frame("Count")
                .reset_index()
                .rename(columns={"index": "TAT_Bucket"})
            )
            tbc_df.to_excel(writer, sheet_name="TAT_Buckets", index=False)

        # Median TAT by Case Type
        tat_by_case_type_df = (
            tat_by_case_type.reset_index()
            .rename(columns={"TAT_Days": "Median_TAT_Days"})
        )
        tat_by_case_type_df.to_excel(
            writer, sheet_name="Median_TAT_by_CaseType", index=False
        )

        # Median TAT by Hold Category and TAT Bucket
        if not tat_by_hold_bucket.empty:
            tat_by_hold_bucket.to_excel(
                writer, sheet_name="Median_TAT_by_Hold_Bucket"
            )

        # Median TAT by Hold Category
        tat_by_hold_category_df = (
            tat_by_hold_category.to_frame("Median_TAT_Days")
            .reset_index()
            .rename(columns={"index": "Hold_Category"})
        )
        tat_by_hold_category_df.to_excel(
            writer, sheet_name="Median_TAT_by_HoldCategory", index=False
        )

        # Monthly statistics
        if not monthly_stats.empty:
            monthly_stats.to_excel(writer, sheet_name="Monthly_Statistics")

        # Location / Vehicle analysis
        if not location_analysis.empty:
            location_analysis.to_excel(writer, sheet_name="Location_Analysis")
        if not vehicle_analysis.empty:
            vehicle_analysis.to_excel(writer, sheet_name="Vehicle_Analysis")

        # Holding Time Stats
        if all_holding_times_flat:
            holding_stats = {
                "Metric": [
                    "Median Holding Time (All)",
                    "Median of Highest Holding Time",
                    "Median of Lowest Holding Time",
                    "Highest Holding Time (Overall)",
                    "Lowest Holding Time (Overall)",
                    "Average Number of Touches",
                ],
                "Value": [
                    pd.Series(all_holding_times_flat).median(),
                    highest_per_case.median() if not highest_per_case.empty else 0,
                    lowest_per_case.median() if not lowest_per_case.empty else 0,
                    max(all_holding_times_flat),
                    min(all_holding_times_flat),
                    holding_data["NumberOfTouches"].mean()
                    if not holding_data.empty
                    else 0,
                ],
            }
            pd.DataFrame(holding_stats).to_excel(
                writer, sheet_name="Holding_Time_Stats", index=False
            )

        # Original data
        original_df.to_excel(writer, sheet_name="Original_Data", index=False)

        # Processed data (essential columns)
        essential_cols = [
            "requestId",
            "statusDescription",
            "createDateTime",
            "completedDateTime",
            "onHoldReasonDescriptionsHistory",
            "onHoldDatesHistory",
            "offHoldDatesHistory",
            "requestTypeDescription",
            "onHoldReasonDescription",
            "numberOfLocations",
            "NumberOfVehicles",
            "CaseType",
            "HoldCategory",
            "Aging_Days",
            "TAT_Days",
            "TAT_Bucket",
            "TotalHoldingTime",
            "NumberOfTouches",
            "Month_Str",
        ]
        available_cols = [c for c in essential_cols if c in df.columns]
        df[available_cols].to_excel(writer, sheet_name="Processed_Data", index=False)

    print("Done.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Batch underwriting issuance analysis (no Streamlit)."
    )
    parser.add_argument(
        "--input",
        type=str,
        required=False,
        default="data/auto_issuance_synthetic_1year_10000rows.csv",
        help="Path to input CSV/Excel file.",
    )
    parser.add_argument(
        "--output",
        type=str,
        required=False,
        default="underwriting_report_batch.xlsx",
        help="Path to output Excel report.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    in_path = Path(args.input)
    out_path = Path(args.output)
    run_analysis(in_path, out_path)


