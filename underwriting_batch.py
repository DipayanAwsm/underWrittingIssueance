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
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio


def parse_separated_values(value: Any, separator: str | None = None) -> List[str]:
    """Parse values that may be separated by ',' (with or without space) or '|' into a list."""
    if pd.isna(value) or value == "":
        return []

    text = str(value).strip()
    if not text:
        return []

    if separator:
        separators = [separator]
    else:
        # Try to detect separator: pipe, comma-space, then plain comma
        if "|" in text:
            separators = ["|"]
        elif ", " in text:
            separators = [", "]
        elif "," in text:
            separators = [","]
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

        # Seasonality by case type: Straight Through vs Multi Hold
        straight_mask = df["CaseType"] == "Straight Through"
        multi_mask = df["CaseType"].str.contains("Multi Hold", na=False)

        if straight_mask.any():
            monthly_straight_stats = df[straight_mask].groupby("Month_Str").agg(
                {
                    "requestId": "count",
                    "completedDateTime": lambda x: x.notna().sum(),
                }
            ).rename(
                columns={
                    "requestId": "Total_Cases_ST",
                    "completedDateTime": "Completed_Cases_ST",
                }
            )
        else:
            monthly_straight_stats = pd.DataFrame()

        if multi_mask.any():
            monthly_multi_stats = df[multi_mask].groupby("Month_Str").agg(
                {
                    "requestId": "count",
                    "completedDateTime": lambda x: x.notna().sum(),
                }
            ).rename(
                columns={
                    "requestId": "Total_Cases_MH",
                    "completedDateTime": "Completed_Cases_MH",
                }
            )
        else:
            monthly_multi_stats = pd.DataFrame()
    else:
        monthly_stats = pd.DataFrame()
        monthly_straight_stats = pd.DataFrame()
        monthly_multi_stats = pd.DataFrame()

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

    # Top 10 request types (overall)
    if "requestTypeDescription" in df.columns:
        top_request_types = df["requestTypeDescription"].value_counts().head(10)
    else:
        top_request_types = pd.Series(dtype=int)

    # Top 10 hold reasons (overall, flattened history)
    all_hold_reasons: List[str] = []
    if "onHoldReasonDescriptionsHistory" in df.columns:
        for value in df["onHoldReasonDescriptionsHistory"]:
            all_hold_reasons.extend(parse_separated_values(value))
    if all_hold_reasons:
        top_hold_reasons = pd.Series(all_hold_reasons).value_counts().head(10)
    else:
        top_hold_reasons = pd.Series(dtype=int)

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

    # ------------------------------------------------------------------
    # Create consolidated interactive charts in a single HTML file
    # ------------------------------------------------------------------
    chart_base = output_path.with_suffix("")  # e.g. underwriting_report_batch
    figs: List[go.Figure] = []

    # 0) High-level summary bar (shown first)
    summary_x = [
        "Total Cases",
        "Completed Cases",
        "Straight Through Cases",
        "Multi Hold Cases",
    ]
    summary_y = [
        len(df),
        df["completedDateTime"].notna().sum()
        if "completedDateTime" in df.columns
        else 0,
        len(straight_through),
        len(multi_hold),
    ]
    fig_summary = px.bar(
        x=summary_x,
        y=summary_y,
        labels={"x": "Metric", "y": "Count"},
        title="Overall Summary: Total, Completed, Straight Through, Multi Hold",
    )
    fig_summary.update_xaxes(tickangle=0)
    figs.append(fig_summary)

    # 1) Status Count (bar)
    if not status_counts.empty:
        sc_df = status_counts.to_frame("Count").reset_index()
        # After value_counts on 'statusDescription', the index column is named 'statusDescription'
        # Normalize to a generic 'Status' column for plotting
        if "statusDescription" in sc_df.columns:
            sc_df = sc_df.rename(columns={"statusDescription": "Status"})

        fig_status = px.bar(
            sc_df,
            x="Status",
            y="Count",
            title="Status Count",
            labels={"Status": "Status", "Count": "Count"},
        )
        fig_status.update_xaxes(tickangle=45)
        figs.append(fig_status)

    # 2) Aging distribution (histogram)
    if not aging_data.empty:
        fig_aging = px.histogram(
            aging_data,
            x="Aging_Days",
            nbins=30,
            title="Distribution of Aging Days",
            labels={"Aging_Days": "Aging (Days)", "count": "Frequency"},
        )
        figs.append(fig_aging)

    # 3) Case type distribution (pie)
    if not case_type_counts.empty:
        ctc_df = (
            case_type_counts.to_frame("Count")
            .reset_index()
            .rename(columns={"index": "CaseType"})
        )
        fig_case_type = px.pie(
            ctc_df,
            values="Count",
            names="CaseType",
            title="Case Type Distribution",
        )
        figs.append(fig_case_type)

    # 4) Holding time distribution (histogram)
    if not holding_data.empty:
        fig_holding = px.histogram(
            holding_data,
            x="TotalHoldingTime",
            nbins=30,
            title="Distribution of Total Holding Time (Days)",
            labels={"TotalHoldingTime": "Holding Time (Days)", "count": "Frequency"},
        )
        figs.append(fig_holding)

    # 5) TAT distribution (histogram)
    if not tat_data.empty:
        fig_tat = px.histogram(
            tat_data,
            x="TAT_Days",
            nbins=30,
            title="Distribution of TAT (Days)",
            labels={"TAT_Days": "TAT (Days)", "count": "Frequency"},
        )
        figs.append(fig_tat)

    # 6) TAT Bucket distribution (bar)
    if not tat_bucket_counts.empty:
        tbc_df = (
            tat_bucket_counts.to_frame("Count")
            .reset_index()
            .rename(columns={"index": "TAT_Bucket"})
        )
        fig_tat_bucket = px.bar(
            tbc_df,
            x="TAT_Bucket",
            y="Count",
            title="TAT Bucket Distribution",
            labels={"TAT_Bucket": "TAT Bucket", "Count": "Count"},
        )
        figs.append(fig_tat_bucket)

        # For each TAT bucket, show top Request Types and Write-Out Reasons
        for bucket in sorted(tbc_df["TAT_Bucket"].dropna().unique()):
            bucket_df = df[df["TAT_Bucket"] == bucket].copy()

            # Top request types for this bucket
            if "requestTypeDescription" in bucket_df.columns:
                trt_bucket = (
                    bucket_df["requestTypeDescription"]
                    .value_counts()
                    .head(10)
                )
                if not trt_bucket.empty:
                    trt_bucket_df = (
                        trt_bucket.to_frame("Count")
                        .reset_index()
                        .rename(columns={"index": "requestTypeDescription"})
                    )
                    fig_trt_bucket = px.bar(
                        trt_bucket_df,
                        x="Count",
                        y="requestTypeDescription",
                        orientation="h",
                        title=f"Top 10 Request Types (TAT Bucket = {bucket})",
                        labels={
                            "Count": "Count",
                            "requestTypeDescription": "Request Type",
                        },
                    )
                    figs.append(fig_trt_bucket)

            # Top write-out reasons for this bucket (history column preferred)
            all_wr_reasons: List[str] = []
            col_hist = None
            if "writeOutReasonDescriptionsHistory" in bucket_df.columns:
                col_hist = "writeOutReasonDescriptionsHistory"
            elif "writeOutReasonDescription" in bucket_df.columns:
                col_hist = "writeOutReasonDescription"

            if col_hist is not None:
                for val in bucket_df[col_hist]:
                    all_wr_reasons.extend(parse_separated_values(val))

                if all_wr_reasons:
                    wr_series = (
                        pd.Series(all_wr_reasons).value_counts().head(10)
                    )
                    if not wr_series.empty:
                        wr_df = (
                            wr_series.to_frame("Count")
                            .reset_index()
                            .rename(columns={"index": "WriteOutReason"})
                        )
                        fig_wr_bucket = px.bar(
                            wr_df,
                            x="Count",
                            y="WriteOutReason",
                            orientation="h",
                            title=f"Top 10 Write-Out Reasons (TAT Bucket = {bucket})",
                            labels={
                                "Count": "Count",
                                "WriteOutReason": "Write-Out Reason",
                            },
                        )
                        figs.append(fig_wr_bucket)

    # 7) Median TAT by Case Type (bar)
    if not tat_by_case_type.empty:
        fig_tat_case = px.bar(
            tat_by_case_type.reset_index(),
            x="CaseType",
            y="TAT_Days",
            title="Median TAT by Case Type",
            labels={"CaseType": "Case Type", "TAT_Days": "Median TAT (Days)"},
        )
        fig_tat_case.update_xaxes(tickangle=45)
        figs.append(fig_tat_case)

    # 8) Median TAT by HoldCategory & TAT_Bucket (grouped bar)
    if not tat_by_hold_bucket.empty:
        fig_tat_hold_bucket = px.bar(
            tat_by_hold_bucket.reset_index(),
            x="TAT_Bucket",
            y=[c for c in tat_by_hold_bucket.columns if c != "TAT_Bucket"],
            barmode="group",
            title="Median TAT by Hold Category and TAT Bucket",
            labels={"value": "Median TAT (Days)", "TAT_Bucket": "TAT Bucket"},
        )
        figs.append(fig_tat_hold_bucket)

    # 9) Seasonality: completion & holding rates (overall + by case type)
    if not monthly_stats.empty:
        ms_reset = monthly_stats.reset_index().rename(columns={"index": "Month_Str"})

        fig_completion = px.line(
            ms_reset,
            x="Month_Str",
            y="Completion_Rate",
            title="Monthly Completion Rate (%) - All Cases",
            labels={"Month_Str": "Month", "Completion_Rate": "Completion Rate (%)"},
        )
        fig_completion.update_xaxes(tickangle=45)
        figs.append(fig_completion)

        fig_holding_rate = px.line(
            ms_reset,
            x="Month_Str",
            y="Holding_Rate",
            title="Monthly Holding Rate (%) - All Cases",
            labels={"Month_Str": "Month", "Holding_Rate": "Holding Rate (%)"},
        )
        fig_holding_rate.update_xaxes(tickangle=45)
        figs.append(fig_holding_rate)

        # Seasonality for Straight Through cases (counts per month)
        if not monthly_straight_stats.empty:
            ms_st_reset = monthly_straight_stats.reset_index()
            fig_straight_seasonality = px.line(
                ms_st_reset,
                x="Month_Str",
                y="Total_Cases_ST",
                title="Seasonality - Straight Through Cases (Monthly Volume)",
                labels={"Month_Str": "Month", "Total_Cases_ST": "Straight Through Cases"},
            )
            fig_straight_seasonality.update_xaxes(tickangle=45)
            figs.append(fig_straight_seasonality)

        # Seasonality for Multi Hold cases (counts per month)
        if not monthly_multi_stats.empty:
            ms_mh_reset = monthly_multi_stats.reset_index()
            fig_multi_seasonality = px.line(
                ms_mh_reset,
                x="Month_Str",
                y="Total_Cases_MH",
                title="Seasonality - Multi Hold Cases (Monthly Volume)",
                labels={"Month_Str": "Month", "Total_Cases_MH": "Multi Hold Cases"},
            )
            fig_multi_seasonality.update_xaxes(tickangle=45)
            figs.append(fig_multi_seasonality)

    # 10) Location & Vehicle analysis charts
    if not location_analysis.empty:
        fig_loc = px.bar(
            location_analysis.reset_index(),
            x="numberOfLocations",
            y="Median_TAT",
            title="Median TAT by Number of Locations",
            labels={
                "numberOfLocations": "Number of Locations",
                "Median_TAT": "Median TAT (Days)",
            },
        )
        figs.append(fig_loc)

    if not vehicle_analysis.empty:
        fig_veh = px.bar(
            vehicle_analysis.reset_index(),
            x="NumberOfVehicles",
            y="Median_TAT",
            title="Median TAT by Number of Vehicles",
            labels={
                "NumberOfVehicles": "Number of Vehicles",
                "Median_TAT": "Median TAT (Days)",
            },
        )
        figs.append(fig_veh)

    if not lv_data.empty:
        fig_corr_touches = px.scatter(
            lv_data,
            x="NumberOfVehicles",
            y="NumberOfTouches",
            trendline="ols",
            title="Number of Vehicles vs Number of Touches",
        )
        figs.append(fig_corr_touches)

        fig_corr_hold = px.scatter(
            lv_data,
            x="NumberOfVehicles",
            y="TotalHoldingTime",
            trendline="ols",
            title="Number of Vehicles vs Holding Time",
        )
        figs.append(fig_corr_hold)

    # 11) Top 10 request types (overall)
    if not top_request_types.empty:
        trt_df = (
            top_request_types.to_frame("Count")
            .reset_index()
            .rename(columns={"index": "requestTypeDescription"})
        )
        fig_trt = px.bar(
            trt_df,
            x="Count",
            y="requestTypeDescription",
            orientation="h",
            title="Top 10 Request Types (Overall)",
            labels={"Count": "Count", "requestTypeDescription": "Request Type"},
        )
        figs.append(fig_trt)

    # 12) Top 10 hold reasons (overall)
    if not top_hold_reasons.empty:
        thr_df = (
            top_hold_reasons.to_frame("Count")
            .reset_index()
            .rename(columns={"index": "HoldReason"})
        )
        fig_thr = px.bar(
            thr_df,
            x="Count",
            y="HoldReason",
            orientation="h",
            title="Top 10 Hold Reasons (Overall)",
            labels={"Count": "Count", "HoldReason": "Hold Reason"},
        )
        figs.append(fig_thr)

    # 13) Highest vs Lowest holding times per case (box plots)
    if not holding_data.empty and not highest_per_case.empty:
        fig_high_low = go.Figure()
        fig_high_low.add_trace(
            go.Box(
                y=highest_per_case,
                name="Highest Holding Time per Case",
                boxmean="sd",
            )
        )
        if not lowest_per_case.empty:
            fig_high_low.add_trace(
                go.Box(
                    y=lowest_per_case,
                    name="Lowest Holding Time per Case",
                    boxmean="sd",
                )
            )
        fig_high_low.update_layout(
            title="Highest vs Lowest Holding Times per Case",
            yaxis_title="Holding Time (Days)",
        )
        figs.append(fig_high_low)

    # 14) Consolidated HTML file with all charts
    if figs:
        html_path_all = chart_base.with_name(f"{chart_base.name}_all_charts.html")
        parts = []
        for idx, fig in enumerate(figs):
            parts.append(
                pio.to_html(
                    fig,
                    full_html=False,
                    include_plotlyjs=True if idx == 0 else False,
                )
            )

        html_content = (
            "<html><head><meta charset='utf-8'><title>Underwriting Charts</title></head>"
            "<body>\n"
            + "\n<hr/>\n".join(parts)
            + "\n</body></html>"
        )
        html_path_all.write_text(html_content, encoding="utf-8")
        print(f"All interactive charts written to single HTML file: {html_path_all}")

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


