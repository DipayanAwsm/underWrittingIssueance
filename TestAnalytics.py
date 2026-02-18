import io
import re
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(page_title="Auto Issuance Prescriptive Analytics", layout="wide")

MISSING_HISTORY_TOKENS = {"", "nan", "none", "null", "na", "n/a", "-", "[]", "0"}


def normalize_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    normalized_map = {normalize_name(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_name(candidate)
        if key in normalized_map:
            return normalized_map[key]
    return None


def parse_history_cell(value: object) -> List[str]:
    if isinstance(value, (list, tuple)):
        cleaned = []
        for item in value:
            text = str(item).strip()
            if text.lower() in MISSING_HISTORY_TOKENS:
                continue
            if text:
                cleaned.append(text)
        return cleaned
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text or text.lower() in MISSING_HISTORY_TOKENS:
        return []
    text = text.strip("[]")
    parts = []
    for part in re.split(r"\s*,\s*", text):
        item = part.strip()
        if not item:
            continue
        if item.lower() in MISSING_HISTORY_TOKENS:
            continue
        parts.append(item)
    return parts


def parse_dt_value(value: object) -> pd.Timestamp:
    if value is None:
        return pd.NaT
    text = str(value).strip()
    if not text or text.lower() in MISSING_HISTORY_TOKENS:
        return pd.NaT
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def calculate_hold_metrics(
    on_hold_text: object,
    off_hold_text: object,
    reason_text: object,
) -> Tuple[float, int, int, bool]:
    on_values = parse_history_cell(on_hold_text)
    off_values = parse_history_cell(off_hold_text)
    reason_values = parse_history_cell(reason_text)

    total_hold_days = 0.0
    for idx, on_value in enumerate(on_values):
        on_dt = parse_dt_value(on_value)
        if pd.isna(on_dt):
            continue

        off_dt = parse_dt_value(off_values[idx]) if idx < len(off_values) else pd.NaT
        if pd.isna(off_dt):
            # Missing off-hold means no valid hold interval to count.
            continue

        hold_days = (off_dt - on_dt).total_seconds() / 86400
        if hold_days < 0:
            continue
        total_hold_days += hold_days

    reason_count = len(reason_values)
    multihold_count = reason_count + 1
    straight_through = reason_count == 0
    return total_hold_days, reason_count, multihold_count, straight_through


@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, file_name: str, delimiter: str) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)
    lower_name = file_name.lower()

    if lower_name.endswith((".xlsx", ".xls")):
        return pd.read_excel(bio, dtype=str)

    if delimiter == "auto":
        return pd.read_csv(bio, sep=None, engine="python", dtype=str, on_bad_lines="skip")
    if delimiter == "tab":
        return pd.read_csv(bio, sep="\t", dtype=str, on_bad_lines="skip")
    return pd.read_csv(bio, sep=delimiter, dtype=str, on_bad_lines="skip")


@st.cache_data(show_spinner=False)
def prepare_data(raw_df: pd.DataFrame, today_str: str) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    today = pd.Timestamp(today_str)
    df = raw_df.copy()
    df.columns = [str(col).strip() for col in df.columns]

    create_col = find_column(df, ["createDateTime", "create_date_time"])
    completed_col = find_column(df, ["completedDateTime", "completed_date_time"])
    received_col = find_column(df, ["receivedDateTime", "received_date_time"])
    status_col = find_column(df, ["statusDescription", "status_description"])
    request_type_col = find_column(df, ["requestTypeDescription", "requestTypeCode", "requestType"])
    request_col = find_column(df, ["requestId", "request_id"])
    on_hold_col = find_column(df, ["onHoldDatesHistory"])
    off_hold_col = find_column(df, ["offHoldDatesHistory"])
    hold_reason_col = find_column(df, ["onHoldReasonDescriptionsHistory"])
    hold_reason_current_col = find_column(df, ["onHoldReasonDescription"])
    underwriter_col = find_column(df, ["underwriterName", "underwriter"])
    analyst_col = find_column(df, ["accountAnalystName", "accountAnalyst"])
    locations_col = find_column(df, ["numberOfLocations"])
    cancellation_col = find_column(
        df,
        ["cancellationReason", "reasonForCancellationDescription", "cancellationTypeDescription"],
    )

    if request_col is None:
        request_col = "__generated_request_id"
        df[request_col] = [f"REQ_{idx + 1}" for idx in range(len(df))]

    date_cols = [col for col in [create_col, completed_col, received_col] if col is not None]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    if status_col is None:
        df["__status"] = "Unknown"
        status_col = "__status"
    df["status_value"] = df[status_col].fillna("Unknown").astype(str).str.strip()
    df["status_value"] = df["status_value"].replace("", "Unknown")

    if request_type_col is not None:
        df["request_type_value"] = df[request_type_col].fillna("Unknown").astype(str).str.strip()
        df["request_type_value"] = df["request_type_value"].replace("", "Unknown")
    else:
        df["request_type_value"] = "Unknown"

    if create_col is not None:
        df["create_month_dt"] = df[create_col].dt.to_period("M").dt.to_timestamp()
        df["create_month"] = df["create_month_dt"].dt.strftime("%Y-%m")
    else:
        df["create_month_dt"] = pd.NaT
        df["create_month"] = "Unknown"

    if all(col is not None for col in [on_hold_col, off_hold_col, hold_reason_col]):
        hold_metrics = df.apply(
            lambda row: calculate_hold_metrics(
                row[on_hold_col], row[off_hold_col], row[hold_reason_col]
            ),
            axis=1,
            result_type="expand",
        )
        hold_metrics.columns = [
            "total_hold_days",
            "hold_reason_count",
            "multihold_count",
            "straight_through",
        ]
        df = pd.concat([df, hold_metrics], axis=1)
    else:
        reason_count = (
            df[hold_reason_col].apply(lambda x: len(parse_history_cell(x)))
            if hold_reason_col is not None
            else pd.Series([0] * len(df))
        )
        df["total_hold_days"] = 0.0
        df["hold_reason_count"] = reason_count
        df["multihold_count"] = df["hold_reason_count"] + 1
        df["straight_through"] = df["hold_reason_count"] == 0

    df["total_hold_days"] = pd.to_numeric(df["total_hold_days"], errors="coerce").fillna(0.0).clip(lower=0.0)
    df["hold_reason_count"] = pd.to_numeric(df["hold_reason_count"], errors="coerce").fillna(0).astype(int)
    df["multihold_count"] = pd.to_numeric(df["multihold_count"], errors="coerce").fillna(1).astype(int)

    df["case_type"] = np.where(df["straight_through"], "StraightThrough", "MultiHold")
    df["touches_count"] = pd.to_numeric(df["multihold_count"], errors="coerce").fillna(1).clip(lower=1)

    if completed_col is not None:
        # Completed means completedDateTime exists and is a valid parsed datetime.
        df["is_completed"] = df[completed_col].notna()
    else:
        df["is_completed"] = False

    df["gross_tat_days"] = np.nan
    df["tat_days"] = np.nan
    if create_col is not None and completed_col is not None:
        gross_tat_days = (df[completed_col] - df[create_col]).dt.total_seconds() / 86400
        gross_tat_days = gross_tat_days.where(gross_tat_days >= 0, np.nan)
        df["gross_tat_days"] = gross_tat_days

        # Net TAT definition from user: completed - created - holding time.
        tat_days = (gross_tat_days - df["total_hold_days"]).clip(lower=0)
        df["tat_days"] = tat_days

    df["tat_band"] = pd.cut(
        df["tat_days"],
        bins=[0, 4, 7, np.inf],
        labels=["1-4 days", "5-7 days", "7+ days"],
        include_lowest=True,
    )

    df["aging_open_days"] = np.nan
    if create_col is not None:
        open_mask = ~df["is_completed"]
        aging_days = (today - df.loc[open_mask, create_col]).dt.total_seconds() / 86400
        aging_days = aging_days.where(aging_days >= 0, np.nan)
        df.loc[open_mask, "aging_open_days"] = aging_days

    if locations_col is not None:
        df["number_of_locations_numeric"] = pd.to_numeric(df[locations_col], errors="coerce")
    else:
        df["number_of_locations_numeric"] = np.nan

    if underwriter_col is not None:
        df["underwriter_value"] = df[underwriter_col].fillna("Unassigned").astype(str).str.strip()
        df["underwriter_value"] = df["underwriter_value"].replace("", "Unassigned")
    else:
        df["underwriter_value"] = "Unassigned"

    if analyst_col is not None:
        df["analyst_value"] = df[analyst_col].fillna("Unassigned").astype(str).str.strip()
        df["analyst_value"] = df["analyst_value"].replace("", "Unassigned")
    else:
        df["analyst_value"] = "Unassigned"

    if cancellation_col is not None:
        df["cancellation_value"] = df[cancellation_col].fillna("Unknown").astype(str).str.strip()
        df["cancellation_value"] = df["cancellation_value"].replace("", "Unknown")
    else:
        df["cancellation_value"] = "Unknown"

    if hold_reason_current_col is not None:
        df["on_hold_reason_value"] = df[hold_reason_current_col].fillna("No hold reason").astype(str).str.strip()
        df["on_hold_reason_value"] = df["on_hold_reason_value"].replace("", "No hold reason")
    elif hold_reason_col is not None:
        df["on_hold_reason_value"] = df[hold_reason_col].apply(
            lambda x: parse_history_cell(x)[0] if len(parse_history_cell(x)) > 0 else "No hold reason"
        )
    else:
        df["on_hold_reason_value"] = "No hold reason"

    metadata = {
        "create_col": create_col,
        "completed_col": completed_col,
        "received_col": received_col,
        "status_col": status_col,
        "request_type_col": request_type_col,
        "request_col": request_col,
        "on_hold_col": on_hold_col,
        "off_hold_col": off_hold_col,
        "hold_reason_col": hold_reason_col,
        "hold_reason_current_col": hold_reason_current_col,
        "underwriter_col": underwriter_col,
        "analyst_col": analyst_col,
        "locations_col": locations_col,
        "cancellation_col": cancellation_col,
    }
    return df, metadata


@st.cache_data(show_spinner=False)
def build_hold_segments(df: pd.DataFrame, metadata: Dict[str, Optional[str]]) -> pd.DataFrame:
    on_hold_col = metadata.get("on_hold_col")
    off_hold_col = metadata.get("off_hold_col")
    reason_col = metadata.get("hold_reason_col")
    request_col = metadata.get("request_col")

    if not all([on_hold_col, off_hold_col, reason_col, request_col]):
        return pd.DataFrame(
            columns=[
                "request_id",
                "create_month",
                "status_value",
                "hold_reason",
                "hold_days",
                "underwriter_value",
                "analyst_value",
                "request_type_value",
            ]
        )

    rows = []
    subset = df[
        [
            request_col,
            on_hold_col,
            off_hold_col,
            reason_col,
            "create_month",
            "status_value",
            "underwriter_value",
            "analyst_value",
            "request_type_value",
        ]
    ].copy()

    for _, row in subset.iterrows():
        on_values = parse_history_cell(row[on_hold_col])
        off_values = parse_history_cell(row[off_hold_col])
        reason_values = parse_history_cell(row[reason_col])

        for idx, on_value in enumerate(on_values):
            on_dt = parse_dt_value(on_value)
            if pd.isna(on_dt):
                continue

            off_dt = parse_dt_value(off_values[idx]) if idx < len(off_values) else pd.NaT
            if pd.isna(off_dt):
                # Missing off-hold means no valid hold interval to count.
                continue

            hold_days = (off_dt - on_dt).total_seconds() / 86400
            if hold_days < 0:
                continue

            reason = reason_values[idx] if idx < len(reason_values) else "Unspecified"
            reason = reason if reason else "Unspecified"

            rows.append(
                {
                    "request_id": row[request_col],
                    "create_month": row["create_month"],
                    "status_value": row["status_value"],
                    "hold_reason": reason,
                    "hold_days": hold_days,
                    "underwriter_value": row["underwriter_value"],
                    "analyst_value": row["analyst_value"],
                    "request_type_value": row["request_type_value"],
                }
            )

    return pd.DataFrame(rows)


def apply_filters(
    df: pd.DataFrame,
    metadata: Dict[str, Optional[str]],
    date_range: Optional[Tuple[pd.Timestamp, pd.Timestamp]],
    statuses: List[str],
    case_types: List[str],
    underwriters: List[str],
    analysts: List[str],
) -> pd.DataFrame:
    filtered = df.copy()

    create_col = metadata.get("create_col")
    if create_col is not None and date_range is not None:
        start_dt, end_dt = date_range
        mask = filtered[create_col].between(start_dt, end_dt, inclusive="both")
        filtered = filtered[mask]

    if statuses:
        filtered = filtered[filtered["status_value"].isin(statuses)]
    if case_types:
        filtered = filtered[filtered["case_type"].isin(case_types)]
    if underwriters:
        filtered = filtered[filtered["underwriter_value"].isin(underwriters)]
    if analysts:
        filtered = filtered[filtered["analyst_value"].isin(analysts)]

    return filtered


def format_metric(value: float, decimals: int = 2) -> str:
    if value is None or pd.isna(value):
        return "NA"
    return f"{value:.{decimals}f}"


def summarize_role_performance(
    filtered_df: pd.DataFrame,
    completed_tat_df: pd.DataFrame,
    owner_col: str,
    tat_target: float,
) -> pd.DataFrame:
    role_all = (
        filtered_df.groupby(owner_col, as_index=False)
        .agg(
            total_cases=("case_type", "size"),
            completed_cases=("is_completed", "sum"),
            open_cases=("is_completed", lambda s: (~s).sum()),
            avg_hold_days=("total_hold_days", "mean"),
            avg_touches=("touches_count", "mean"),
            avg_open_aging=("aging_open_days", "mean"),
        )
    )

    if not completed_tat_df.empty:
        role_tat = completed_tat_df.groupby(owner_col, as_index=False).agg(
            avg_tat_days=("tat_days", "mean"),
            p90_tat_days=("tat_days", lambda s: s.quantile(0.9)),
            pct_over_target=("tat_days", lambda s: (s > tat_target).mean() * 100),
            tat_cases=("tat_days", "size"),
        )
        role_perf = role_all.merge(role_tat, on=owner_col, how="left")
    else:
        role_perf = role_all.copy()
        role_perf["avg_tat_days"] = np.nan
        role_perf["p90_tat_days"] = np.nan
        role_perf["pct_over_target"] = np.nan
        role_perf["tat_cases"] = 0

    role_perf["completed_cases"] = role_perf["completed_cases"].fillna(0).astype(int)
    role_perf["open_cases"] = role_perf["open_cases"].fillna(0).astype(int)
    role_perf["tat_cases"] = role_perf["tat_cases"].fillna(0).astype(int)
    return role_perf.sort_values("total_cases", ascending=False)


def build_owner_driver_tables(
    owner_cases: pd.DataFrame,
    owner_tat: pd.DataFrame,
    owner_hold_segments: pd.DataFrame,
    tat_target: float,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if not owner_hold_segments.empty:
        reason_summary = (
            owner_hold_segments.groupby("hold_reason", as_index=False)
            .agg(total_hold_days=("hold_days", "sum"), hold_events=("hold_days", "size"))
            .sort_values("total_hold_days", ascending=False)
        )
    else:
        reason_summary = (
            owner_cases.groupby("on_hold_reason_value", as_index=False)
            .agg(total_hold_days=("total_hold_days", "sum"), hold_events=("case_type", "size"))
            .rename(columns={"on_hold_reason_value": "hold_reason"})
            .sort_values("total_hold_days", ascending=False)
        )

    request_summary = (
        owner_cases.groupby("request_type_value", as_index=False)
        .agg(
            total_cases=("case_type", "size"),
            open_cases=("is_completed", lambda s: (~s).sum()),
            avg_hold_days=("total_hold_days", "mean"),
            avg_touches=("touches_count", "mean"),
        )
        .sort_values("total_cases", ascending=False)
    )
    request_summary["open_cases"] = request_summary["open_cases"].fillna(0).astype(int)

    if not owner_tat.empty:
        request_tat = owner_tat.groupby("request_type_value", as_index=False).agg(
            avg_tat_days=("tat_days", "mean"),
            pct_over_target=("tat_days", lambda s: (s > tat_target).mean() * 100),
            tat_cases=("tat_days", "size"),
        )
        request_summary = request_summary.merge(request_tat, on="request_type_value", how="left")
    else:
        request_summary["avg_tat_days"] = np.nan
        request_summary["pct_over_target"] = np.nan
        request_summary["tat_cases"] = 0

    request_summary["tat_cases"] = request_summary["tat_cases"].fillna(0).astype(int)
    return reason_summary, request_summary


def build_owner_prescription(
    role_label: str,
    owner_name: str,
    owner_cases: pd.DataFrame,
    owner_tat: pd.DataFrame,
    owner_hold_segments: pd.DataFrame,
    all_cases: pd.DataFrame,
    all_tat: pd.DataFrame,
    tat_target: float,
) -> Tuple[List[str], List[str]]:
    story: List[str] = []
    actions: List[str] = []

    owner_avg_hold = owner_cases["total_hold_days"].mean()
    team_avg_hold = all_cases["total_hold_days"].mean()
    story.append(
        f"{role_label} '{owner_name}' avg hold time is {owner_avg_hold:.2f} days vs team {team_avg_hold:.2f}."
    )

    if owner_avg_hold > team_avg_hold + 0.5:
        actions.append(
            f"Reduce hold-time for {role_label} '{owner_name}' by pre-validating documents before assignment."
        )

    if not owner_tat.empty and not all_tat.empty:
        owner_avg_tat = owner_tat["tat_days"].mean()
        team_avg_tat = all_tat["tat_days"].mean()
        owner_over = (owner_tat["tat_days"] > tat_target).mean() * 100
        team_over = (all_tat["tat_days"] > tat_target).mean() * 100
        story.append(
            f"Avg TAT is {owner_avg_tat:.2f} days ({owner_over:.1f}% over {tat_target:.0f}d) vs team avg {team_avg_tat:.2f} ({team_over:.1f}% over)."
        )

        if owner_avg_tat > team_avg_tat + 1:
            actions.append(
                f"Rebalance complex workload for {role_label} '{owner_name}' and add a fast-lane for low-risk requests."
            )
        if owner_over > team_over + 10:
            actions.append(
                f"Run a daily exception queue for {role_label} '{owner_name}' focused on 7+ day risk cases."
            )

    if not owner_hold_segments.empty:
        reason_summary = (
            owner_hold_segments.groupby("hold_reason", as_index=False)
            .agg(total_hold_days=("hold_days", "sum"))
            .sort_values("total_hold_days", ascending=False)
        )
        if not reason_summary.empty:
            top_reason = reason_summary.iloc[0]
            reason_share = 100 * top_reason["total_hold_days"] / max(reason_summary["total_hold_days"].sum(), 1e-9)
            story.append(
                f"Top onHoldReasonDescription is '{top_reason['hold_reason']}' contributing {reason_share:.1f}% of hold days."
            )
            actions.append(
                f"Create a targeted SOP for '{top_reason['hold_reason']}' with defined owner and same-day SLA."
            )

    request_tat = (
        owner_tat.groupby("request_type_value", as_index=False)
        .agg(cases=("tat_days", "size"), over_target_pct=("tat_days", lambda s: (s > tat_target).mean() * 100))
        .sort_values(["over_target_pct", "cases"], ascending=[False, False])
    )
    high_risk_request = request_tat[request_tat["cases"] >= 3]
    if not high_risk_request.empty:
        req = high_risk_request.iloc[0]
        story.append(
            f"RequestType '{req['request_type_value']}' has {req['over_target_pct']:.1f}% TAT breach for this {role_label.lower()}."
        )
        actions.append(
            f"Use request-type based routing: prioritize '{req['request_type_value']}' with early checklist and senior review."
        )

    aged_open = owner_cases[(~owner_cases["is_completed"]) & (owner_cases["aging_open_days"] > tat_target)]
    if not aged_open.empty:
        actions.append(
            f"Clear aged open backlog ({len(aged_open)} cases) for {role_label} '{owner_name}' before accepting new non-urgent work."
        )

    return story, list(dict.fromkeys(actions))


def render_role_tab(
    role_label: str,
    owner_col: str,
    role_perf: pd.DataFrame,
    filtered_df: pd.DataFrame,
    completed_tat_df: pd.DataFrame,
    hold_segments: pd.DataFrame,
    tat_target: float,
    widget_key_prefix: str,
) -> None:
    st.markdown(f"**Holding time and TAT for {role_label}**")
    if role_perf.empty:
        st.info(f"No data available for {role_label.lower()} analysis.")
        return

    st.dataframe(role_perf.head(25), use_container_width=True)

    plot_df = role_perf.dropna(subset=["avg_tat_days"])
    if not plot_df.empty:
        fig_perf = px.scatter(
            plot_df,
            x="total_cases",
            y="avg_tat_days",
            size="avg_hold_days",
            color="pct_over_target",
            hover_name=owner_col,
            title=f"{role_label}: Volume vs Avg TAT (bubble size = hold time)",
        )
        fig_perf.update_layout(xaxis_title="Total Cases", yaxis_title="Avg Completed TAT (days)")
        st.plotly_chart(fig_perf, use_container_width=True)
    else:
        st.info("Not enough completedDateTime records for TAT comparison in this role.")

    owner_options = role_perf[owner_col].astype(str).tolist()
    selected_owner = st.selectbox(
        f"Select {role_label} for onHoldReasonDescription / requestTypeDescription analysis",
        options=owner_options,
        key=f"{widget_key_prefix}_owner_select",
    )

    owner_cases = filtered_df[filtered_df[owner_col] == selected_owner]
    owner_tat = completed_tat_df[completed_tat_df[owner_col] == selected_owner]
    if owner_col in hold_segments.columns:
        owner_hold_segments = hold_segments[hold_segments[owner_col] == selected_owner]
    else:
        owner_hold_segments = hold_segments.iloc[0:0].copy()

    reason_summary, request_summary = build_owner_driver_tables(
        owner_cases,
        owner_tat,
        owner_hold_segments,
        tat_target,
    )

    detail_col1, detail_col2 = st.columns(2)

    if not reason_summary.empty:
        top_reasons = reason_summary.head(10)
        fig_reason = px.bar(
            top_reasons,
            x="hold_reason",
            y="total_hold_days",
            color="hold_events",
            title=f"{role_label} '{selected_owner}': onHoldReasonDescription impact",
        )
        fig_reason.update_layout(xaxis_title="onHoldReasonDescription", yaxis_title="Total Hold Days")
        detail_col1.plotly_chart(fig_reason, use_container_width=True)
    else:
        detail_col1.info("No on-hold reason data available for this selection.")

    if not request_summary.empty:
        req_plot = request_summary.head(10).copy()
        req_plot["request_type_value"] = req_plot["request_type_value"].astype(str)
        fig_request = px.bar(
            req_plot,
            x="request_type_value",
            y="avg_hold_days",
            color="avg_tat_days",
            hover_data=["total_cases", "open_cases", "avg_touches", "pct_over_target", "tat_cases"],
            title=f"{role_label} '{selected_owner}': requestTypeDescription analysis",
        )
        fig_request.update_layout(xaxis_title="requestTypeDescription", yaxis_title="Avg Hold Days")
        detail_col2.plotly_chart(fig_request, use_container_width=True)
        st.dataframe(
            request_summary.head(15),
            use_container_width=True,
        )
    else:
        detail_col2.info("No request type data available for this selection.")

    story, actions = build_owner_prescription(
        role_label=role_label,
        owner_name=selected_owner,
        owner_cases=owner_cases,
        owner_tat=owner_tat,
        owner_hold_segments=owner_hold_segments,
        all_cases=filtered_df,
        all_tat=completed_tat_df,
        tat_target=tat_target,
    )

    st.markdown(f"**{role_label} Prescriptive Model**")
    if story:
        for line in story:
            st.markdown(f"- {line}")
    else:
        st.markdown("- Not enough data for role-specific story.")

    if actions:
        for idx, action in enumerate(actions, start=1):
            st.markdown(f"{idx}. {action}")
    else:
        st.markdown("1. No high-risk pattern detected for this role selection.")


def build_prescriptions(
    filtered_df: pd.DataFrame,
    hold_segments: pd.DataFrame,
    tat_target: float,
    scenario_reduction_pct: int,
) -> Tuple[List[str], List[str], Dict[str, float]]:
    insights: List[str] = []
    actions: List[str] = []

    completed = filtered_df[filtered_df["is_completed"]]
    completed_tat = completed[completed["tat_days"].notna()]
    open_cases = filtered_df[~filtered_df["is_completed"]]

    current_over_target_pct = np.nan
    simulated_over_target_pct = np.nan
    avg_tat_current = np.nan
    avg_tat_simulated = np.nan

    if not completed_tat.empty:
        avg_tat_current = completed_tat["tat_days"].mean()
        current_over_target_pct = (completed_tat["tat_days"] > tat_target).mean() * 100
        # With net TAT definition (gross - hold), reducing hold time does not change net TAT.
        simulated_tat = completed_tat["tat_days"].copy()
        avg_tat_simulated = simulated_tat.mean()
        simulated_over_target_pct = (simulated_tat > tat_target).mean() * 100

        insights.append(
            f"Completed-case average TAT is {avg_tat_current:.2f} days, with {current_over_target_pct:.1f}% above {tat_target:.0f} days."
        )
        insights.append(
            f"Because TAT is defined net of hold time, reducing hold by {scenario_reduction_pct}% does not directly change this TAT metric."
        )
        if current_over_target_pct > 25:
            actions.append("Prioritize a rapid-clearance lane for cases likely to cross 7+ days (daily queue with strict ownership).")

    if not hold_segments.empty:
        reason_summary = (
            hold_segments.groupby("hold_reason", as_index=False)
            .agg(total_hold_days=("hold_days", "sum"), hold_events=("hold_days", "size"))
            .sort_values("total_hold_days", ascending=False)
        )
        top_reason = reason_summary.iloc[0]
        reason_share = 100 * top_reason["total_hold_days"] / max(reason_summary["total_hold_days"].sum(), 1e-9)
        insights.append(
            f"Top hold reason is '{top_reason['hold_reason']}' contributing {top_reason['total_hold_days']:.1f} hold-days ({reason_share:.1f}% of all hold time)."
        )
        actions.append(
            f"Create a focused fix plan for '{top_reason['hold_reason']}' (upstream data checklist, fast-track routing, and SLA with owners)."
        )

    if not open_cases.empty:
        aged_over_target = open_cases[open_cases["aging_open_days"] > tat_target]
        aged_pct = 100 * len(aged_over_target) / len(open_cases)
        insights.append(
            f"{len(aged_over_target)} open cases ({aged_pct:.1f}%) are already older than {tat_target:.0f} days."
        )
        if len(aged_over_target) > 0:
            actions.append("Run a twice-daily triage for aged open cases and clear blockers before adding new work.")

    monthly_completed = (
        completed_tat.groupby("create_month", as_index=False)
        .agg(over_target_rate=("tat_days", lambda s: (s > tat_target).mean() * 100))
        .sort_values("create_month")
    )
    if len(monthly_completed) >= 4:
        latest = monthly_completed.iloc[-1]["over_target_rate"]
        previous_avg = monthly_completed.iloc[-4:-1]["over_target_rate"].mean()
        if latest > previous_avg + 5:
            insights.append(
                f"Recent seasonality shift detected: latest month over-target rate is {latest:.1f}% vs prior 3-month average {previous_avg:.1f}%."
            )
            actions.append("Pre-allocate extra underwriting and analyst capacity in the high-spike months shown in seasonality charts.")

    underwriter_perf = (
        completed_tat.groupby("underwriter_value", as_index=False)
        .agg(avg_tat_days=("tat_days", "mean"), completed_cases=("tat_days", "size"))
        .query("completed_cases >= 5")
        .sort_values("avg_tat_days", ascending=False)
    )
    if not underwriter_perf.empty:
        overall_avg = completed_tat["tat_days"].mean()
        outliers = underwriter_perf[underwriter_perf["avg_tat_days"] > overall_avg + 1.0]
        if not outliers.empty:
            names = ", ".join(outliers["underwriter_value"].head(3).tolist())
            insights.append(f"Underwriter variation is material; higher-TAT cohort includes: {names}.")
            actions.append("Use targeted coaching and workload rebalance for higher-TAT underwriters with sustained volume.")

    scenario_stats = {
        "current_over_target_pct": current_over_target_pct,
        "simulated_over_target_pct": simulated_over_target_pct,
        "avg_tat_current": avg_tat_current,
        "avg_tat_simulated": avg_tat_simulated,
    }
    return insights, actions, scenario_stats


st.title("Auto Insurance Issuance Prescriptive Analytics")
st.caption("Interactive analysis for TAT (target: 7 days), hold-time drivers, seasonality, and productivity actions.")

st.sidebar.header("Data Input")
uploaded_file = st.sidebar.file_uploader("Upload issuance data (.csv, .txt, .xlsx)", type=["csv", "txt", "xlsx", "xls"])

delimiter_map = {
    "Auto detect": "auto",
    "Comma (,)": ",",
    "Tab": "tab",
    "Pipe (|)": "|",
    "Semicolon (;)": ";",
}
delimiter_label = st.sidebar.selectbox("Delimiter (for text/CSV files)", list(delimiter_map.keys()))
delimiter = delimiter_map[delimiter_label]

tat_target = st.sidebar.number_input("TAT target (days)", min_value=1.0, value=7.0, step=0.5)

if uploaded_file is None:
    st.info("Upload your issuance dataset to begin analysis.")
    st.stop()

try:
    raw_df = load_data(uploaded_file.getvalue(), uploaded_file.name, delimiter)
except Exception as exc:
    st.error(f"Unable to read file: {exc}")
    st.stop()

if raw_df.empty:
    st.error("Uploaded file is empty.")
    st.stop()

today_str = pd.Timestamp.today().normalize().strftime("%Y-%m-%d")
df, metadata = prepare_data(raw_df, today_str)

missing_critical = []
for key, label in [
    ("create_col", "createDateTime"),
    ("status_col", "statusDescription"),
]:
    if metadata.get(key) is None:
        missing_critical.append(label)
if metadata.get("hold_reason_col") is None and metadata.get("hold_reason_current_col") is None:
    missing_critical.append("onHoldReasonDescription/onHoldReasonDescriptionsHistory")

if missing_critical:
    st.warning(
        "Some expected columns were not detected: "
        + ", ".join(missing_critical)
        + ". The app still runs with available fields."
    )

with st.expander("Detected columns and sample preview"):
    st.write(metadata)
    st.dataframe(df.head(10), use_container_width=True)

st.sidebar.header("Filters")

create_col = metadata.get("create_col")
date_range = None
if create_col is not None and df[create_col].notna().any():
    month_starts = sorted(pd.to_datetime(df["create_month_dt"].dropna().unique()))
    if month_starts:
        month_labels = [pd.Timestamp(month_val).strftime("%Y-%m") for month_val in month_starts]
        month_label_to_start = {
            label: pd.Timestamp(month_val) for label, month_val in zip(month_labels, month_starts)
        }
        if len(month_labels) == 1:
            only_month_start = month_label_to_start[month_labels[0]]
            only_month_end = only_month_start + pd.offsets.MonthEnd(1) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            date_range = (only_month_start, only_month_end)
            st.sidebar.caption(f"Create month range: {month_labels[0]} to {month_labels[0]}")
        else:
            selected_months = st.sidebar.select_slider(
                "Create month range",
                options=month_labels,
                value=(month_labels[0], month_labels[-1]),
            )
            if isinstance(selected_months, tuple) and len(selected_months) == 2:
                start_month_label, end_month_label = selected_months
                start_month = month_label_to_start[start_month_label]
                end_month_start = month_label_to_start[end_month_label]
                end_month = end_month_start + pd.offsets.MonthEnd(1) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                date_range = (start_month, end_month)

status_options = sorted(df["status_value"].dropna().unique().tolist())
selected_statuses = st.sidebar.multiselect("Status", options=status_options, default=status_options)

case_type_options = ["StraightThrough", "MultiHold"]
selected_case_types = st.sidebar.multiselect("Case Type", options=case_type_options, default=case_type_options)

underwriter_options = sorted(df["underwriter_value"].dropna().unique().tolist())
selected_underwriters = st.sidebar.multiselect("Underwriter", options=underwriter_options, default=underwriter_options)

analyst_options = sorted(df["analyst_value"].dropna().unique().tolist())
selected_analysts = st.sidebar.multiselect("Account Analyst", options=analyst_options, default=analyst_options)

filtered_df = apply_filters(
    df,
    metadata,
    date_range,
    selected_statuses,
    selected_case_types,
    selected_underwriters,
    selected_analysts,
)

if filtered_df.empty:
    st.warning("No rows match the selected filters.")
    st.stop()

with st.expander("Calculation audit (sample rows)"):
    audit_cols = [
        metadata.get("request_col"),
        metadata.get("create_col"),
        metadata.get("completed_col"),
        "is_completed",
        "gross_tat_days",
        "total_hold_days",
        "tat_days",
        "hold_reason_count",
        "touches_count",
        "case_type",
    ]
    audit_cols = [col for col in audit_cols if col in filtered_df.columns and col is not None]
    st.dataframe(filtered_df[audit_cols].head(100), use_container_width=True)

completed_df = filtered_df[filtered_df["is_completed"]]
completed_tat_df = completed_df[completed_df["tat_days"].notna()]
open_df = filtered_df[~filtered_df["is_completed"]]

col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
col1.metric("Total Cases", f"{len(filtered_df):,}")
col2.metric("Completed Cases", f"{len(completed_df):,}")
col3.metric("Open Cases", f"{len(open_df):,}")
col4.metric("Avg Completed TAT (days)", format_metric(completed_tat_df["tat_days"].mean()))
col5.metric("Avg Open Hold Time (days)", format_metric(open_df["total_hold_days"].mean()))
col6.metric("Avg Open Touches", format_metric(open_df["touches_count"].mean(), 1))
over_target_pct = (completed_tat_df["tat_days"] > tat_target).mean() * 100 if not completed_tat_df.empty else np.nan
col7.metric(f">% {tat_target:.0f}d TAT", f"{format_metric(over_target_pct, 1)}%")
completed_without_tat = len(completed_df) - len(completed_tat_df)
if completed_without_tat > 0:
    st.caption(
        f"{completed_without_tat} completed cases are excluded from TAT due to missing/invalid create-complete timestamps or negative duration."
    )

st.subheader("Seasonality and Case Flow")
season_col1, season_col2 = st.columns(2)

monthly_mix = (
    filtered_df.groupby(["create_month", "case_type"], as_index=False)
    .size()
    .rename(columns={"size": "cases"})
    .sort_values("create_month")
)

if not monthly_mix.empty:
    fig_mix = px.bar(
        monthly_mix,
        x="create_month",
        y="cases",
        color="case_type",
        barmode="group",
        title="Month-wise StraightThrough vs MultiHold (by createDateTime)",
    )
    fig_mix.update_layout(xaxis_title="Create Month", yaxis_title="Case Count")
    season_col1.plotly_chart(fig_mix, use_container_width=True)

top_statuses = filtered_df["status_value"].value_counts().head(10).index.tolist()
season_status = (
    filtered_df[filtered_df["status_value"].isin(top_statuses)]
    .groupby(["create_month", "status_value"], as_index=False)
    .size()
    .rename(columns={"size": "cases"})
)

if not season_status.empty:
    fig_heat = px.density_heatmap(
        season_status,
        x="create_month",
        y="status_value",
        z="cases",
        color_continuous_scale="YlOrRd",
        title="Seasonality: Month vs StatusDescription (Top statuses)",
    )
    fig_heat.update_layout(xaxis_title="Create Month", yaxis_title="StatusDescription")
    season_col2.plotly_chart(fig_heat, use_container_width=True)

st.subheader("Month-wise Status Analysis: StraightThrough vs MultiHold")
status_pick_default = top_statuses[:3] if len(top_statuses) >= 3 else top_statuses
status_pick = st.multiselect(
    "Select statuses for trend comparison",
    options=top_statuses,
    default=status_pick_default,
)

status_case_trend = (
    filtered_df[filtered_df["status_value"].isin(status_pick)]
    .groupby(["create_month", "status_value", "case_type"], as_index=False)
    .size()
    .rename(columns={"size": "cases"})
    .sort_values("create_month")
)
if not status_case_trend.empty:
    fig_status_line = px.line(
        status_case_trend,
        x="create_month",
        y="cases",
        color="status_value",
        line_dash="case_type",
        markers=True,
        title="Status trend split by StraightThrough and MultiHold",
    )
    fig_status_line.update_layout(xaxis_title="Create Month", yaxis_title="Case Count")
    st.plotly_chart(fig_status_line, use_container_width=True)

st.subheader("Turnaround and Aging")
tat_col1, tat_col2 = st.columns(2)

tat_band_order = ["1-4 days", "5-7 days", "7+ days"]
if not completed_tat_df.empty:
    tat_band_counts = (
        completed_tat_df["tat_band"]
        .value_counts()
        .reindex(tat_band_order, fill_value=0)
        .rename_axis("tat_band")
        .reset_index(name="cases")
    )
    fig_tat_band = px.bar(
        tat_band_counts,
        x="tat_band",
        y="cases",
        color="tat_band",
        category_orders={"tat_band": tat_band_order},
        title="Completed Case TAT Bands",
    )
    fig_tat_band.update_layout(showlegend=False, xaxis_title="TAT Band", yaxis_title="Cases")
    tat_col1.plotly_chart(fig_tat_band, use_container_width=True)
else:
    tat_col1.info("No completed cases with valid completedDateTime available for TAT band analysis.")

if not open_df.empty:
    fig_aging = px.histogram(
        open_df,
        x="aging_open_days",
        nbins=30,
        color="case_type",
        barmode="overlay",
        title="Aging Distribution for Non-Completed Cases",
    )
    fig_aging.update_layout(xaxis_title="Aging (days)", yaxis_title="Open cases")
    tat_col2.plotly_chart(fig_aging, use_container_width=True)
else:
    tat_col2.info("No open cases available for aging analysis.")

st.subheader("Completed Cases: Month-wise TAT")
if not completed_tat_df.empty:
    monthly_completed_tat = (
        completed_tat_df.groupby("create_month", as_index=False)
        .agg(
            completed_cases=("tat_days", "size"),
            avg_tat_days=("tat_days", "mean"),
            p90_tat_days=("tat_days", lambda s: s.quantile(0.9)),
        )
        .sort_values("create_month")
    )
    completed_melt = monthly_completed_tat.melt(
        id_vars=["create_month", "completed_cases"],
        value_vars=["avg_tat_days", "p90_tat_days"],
        var_name="metric",
        value_name="days",
    )
    completed_melt["metric"] = completed_melt["metric"].map(
        {"avg_tat_days": "Avg TAT", "p90_tat_days": "P90 TAT"}
    )
    fig_completed_tat = px.line(
        completed_melt,
        x="create_month",
        y="days",
        color="metric",
        markers=True,
        title="Monthly TAT Trend (Completed Cases Only)",
    )
    fig_completed_tat.update_layout(xaxis_title="Create Month", yaxis_title="Days")
    st.plotly_chart(fig_completed_tat, use_container_width=True)
else:
    st.info("No completed cases with valid completedDateTime available for month-wise TAT trend.")

st.subheader("Non-Completed Cases: Month-wise Hold Time and Touches")
open_month_col1, open_month_col2 = st.columns(2)

if not open_df.empty:
    monthly_open = (
        open_df.groupby("create_month", as_index=False)
        .agg(
            open_cases=("case_type", "size"),
            avg_open_hold_days=("total_hold_days", "mean"),
            p90_open_aging=("aging_open_days", lambda s: s.quantile(0.9)),
            avg_open_touches=("touches_count", "mean"),
            p90_open_touches=("touches_count", lambda s: s.quantile(0.9)),
        )
        .sort_values("create_month")
    )

    open_days_long = monthly_open.melt(
        id_vars=["create_month", "open_cases"],
        value_vars=["avg_open_hold_days", "p90_open_aging"],
        var_name="metric",
        value_name="value",
    )
    open_days_long["metric"] = open_days_long["metric"].map(
        {"avg_open_hold_days": "Avg Hold Time", "p90_open_aging": "P90 Aging"}
    )
    fig_open_days = px.line(
        open_days_long,
        x="create_month",
        y="value",
        color="metric",
        markers=True,
        title="Open Cases: Monthly Hold Time and Aging",
    )
    fig_open_days.update_layout(xaxis_title="Create Month", yaxis_title="Days")
    open_month_col1.plotly_chart(fig_open_days, use_container_width=True)

    open_touch_long = monthly_open.melt(
        id_vars=["create_month", "open_cases"],
        value_vars=["avg_open_touches", "p90_open_touches"],
        var_name="metric",
        value_name="touches",
    )
    open_touch_long["metric"] = open_touch_long["metric"].map(
        {"avg_open_touches": "Avg Touches", "p90_open_touches": "P90 Touches"}
    )
    fig_open_touches = px.line(
        open_touch_long,
        x="create_month",
        y="touches",
        color="metric",
        markers=True,
        title="Open Cases: Monthly Touch Trend",
    )
    fig_open_touches.update_layout(xaxis_title="Create Month", yaxis_title="Touches")
    open_month_col2.plotly_chart(fig_open_touches, use_container_width=True)

    open_touch_band = open_df.copy()
    open_touch_band["touch_band"] = pd.cut(
        open_touch_band["touches_count"],
        bins=[0, 1, 2, np.inf],
        labels=["1 touch", "2 touches", "3+ touches"],
        include_lowest=True,
    )
    touch_month_dist = (
        open_touch_band.groupby(["create_month", "touch_band"], as_index=False)
        .size()
        .rename(columns={"size": "cases"})
        .sort_values("create_month")
    )
    if not touch_month_dist.empty:
        fig_touch_dist = px.bar(
            touch_month_dist,
            x="create_month",
            y="cases",
            color="touch_band",
            barmode="stack",
            title="Open Cases: Monthly Touch Distribution",
        )
        fig_touch_dist.update_layout(xaxis_title="Create Month", yaxis_title="Open Case Count")
        st.plotly_chart(fig_touch_dist, use_container_width=True)
else:
    open_month_col1.info("No non-completed cases available for month-wise hold analysis.")
    open_month_col2.info("No non-completed cases available for month-wise touch analysis.")

st.subheader("Hold-Time Diagnostics")
hold_segments = build_hold_segments(filtered_df, metadata)
hold_col1, hold_col2 = st.columns(2)

if not hold_segments.empty:
    hold_reason_summary = (
        hold_segments.groupby("hold_reason", as_index=False)
        .agg(total_hold_days=("hold_days", "sum"), hold_events=("hold_days", "size"))
        .sort_values("total_hold_days", ascending=False)
    )
    top_hold_reasons = hold_reason_summary.head(10)
    fig_hold_reason = px.bar(
        top_hold_reasons,
        x="hold_reason",
        y="total_hold_days",
        color="hold_events",
        title="Top Hold Reasons by Total Hold Time",
    )
    fig_hold_reason.update_layout(xaxis_title="Hold Reason", yaxis_title="Total Hold Days")
    hold_col1.plotly_chart(fig_hold_reason, use_container_width=True)

    hold_month_reason = (
        hold_segments[hold_segments["hold_reason"].isin(top_hold_reasons["hold_reason"].head(5))]
        .groupby(["create_month", "hold_reason"], as_index=False)
        .agg(total_hold_days=("hold_days", "sum"))
        .sort_values("create_month")
    )
    fig_hold_month = px.area(
        hold_month_reason,
        x="create_month",
        y="total_hold_days",
        color="hold_reason",
        title="Monthly Hold-Time Contribution (Top reasons)",
    )
    fig_hold_month.update_layout(xaxis_title="Create Month", yaxis_title="Hold Days")
    hold_col2.plotly_chart(fig_hold_month, use_container_width=True)
else:
    hold_col1.info("No valid hold history found for hold-time breakdown.")
    hold_col2.info("No hold trend can be built without hold date history.")

st.subheader("Role-Based TAT, Holding Time and Driver Analysis")
underwriter_perf = summarize_role_performance(filtered_df, completed_tat_df, "underwriter_value", tat_target)
analyst_perf = summarize_role_performance(filtered_df, completed_tat_df, "analyst_value", tat_target)

tab_uw, tab_aa = st.tabs(["Underwriter", "Account Analyst"])

with tab_uw:
    render_role_tab(
        role_label="Underwriter",
        owner_col="underwriter_value",
        role_perf=underwriter_perf,
        filtered_df=filtered_df,
        completed_tat_df=completed_tat_df,
        hold_segments=hold_segments,
        tat_target=tat_target,
        widget_key_prefix="underwriter",
    )

with tab_aa:
    render_role_tab(
        role_label="Account Analyst",
        owner_col="analyst_value",
        role_perf=analyst_perf,
        filtered_df=filtered_df,
        completed_tat_df=completed_tat_df,
        hold_segments=hold_segments,
        tat_target=tat_target,
        widget_key_prefix="analyst",
    )

st.subheader("Complexity and Cancellation Analysis")
complex_col1, complex_col2 = st.columns(2)

loc_df = filtered_df.dropna(subset=["number_of_locations_numeric"]).copy()
if not loc_df.empty and not completed_tat_df.empty:
    loc_df["location_band"] = pd.cut(
        loc_df["number_of_locations_numeric"],
        bins=[-np.inf, 1, 5, 10, np.inf],
        labels=["1", "2-5", "6-10", "10+"],
    )
    loc_tat = (
        loc_df.dropna(subset=["tat_days"])
        .groupby("location_band", as_index=False)
        .agg(avg_tat_days=("tat_days", "mean"), cases=("tat_days", "size"))
    )
    if not loc_tat.empty:
        fig_loc = px.bar(
            loc_tat,
            x="location_band",
            y="avg_tat_days",
            color="cases",
            title="Complexity Check: Avg TAT by Number of Locations",
        )
        fig_loc.update_layout(xaxis_title="Number of Locations band", yaxis_title="Avg TAT (days)")
        complex_col1.plotly_chart(fig_loc, use_container_width=True)
else:
    complex_col1.info("Location-based complexity chart requires numberOfLocations and completed TAT.")

aged_open = open_df[open_df["aging_open_days"] > tat_target]
if not aged_open.empty:
    cancel_summary = (
        aged_open.groupby("cancellation_value", as_index=False)
        .size()
        .rename(columns={"size": "open_aged_cases"})
        .sort_values("open_aged_cases", ascending=False)
        .head(10)
    )
    fig_cancel = px.bar(
        cancel_summary,
        x="cancellation_value",
        y="open_aged_cases",
        title=f"Top Cancellation Reasons in Open Cases Older Than {tat_target:.0f} Days",
    )
    fig_cancel.update_layout(xaxis_title="Cancellation Reason", yaxis_title="Open aged cases")
    complex_col2.plotly_chart(fig_cancel, use_container_width=True)
else:
    complex_col2.info("No aged open cases above target for cancellation analysis.")

st.subheader("Prescriptive Recommendations")
scenario_reduction_pct = st.slider("What-if: reduce hold time by (%)", min_value=0, max_value=50, value=20, step=5)
insights, actions, scenario_stats = build_prescriptions(filtered_df, hold_segments, tat_target, scenario_reduction_pct)

presc_col1, presc_col2 = st.columns(2)

with presc_col1:
    st.markdown("**Story (what the data says)**")
    if insights:
        for line in insights:
            st.markdown(f"- {line}")
    else:
        st.markdown("- Not enough data after filters to generate insight narrative.")

with presc_col2:
    st.markdown("**What you should do next**")
    if actions:
        dedup_actions = list(dict.fromkeys(actions))
        for idx, action in enumerate(dedup_actions, start=1):
            st.markdown(f"{idx}. {action}")
    else:
        st.markdown("1. Expand date range or remove filters to generate stronger prescriptive actions.")

if not pd.isna(scenario_stats["current_over_target_pct"]) and not pd.isna(scenario_stats["simulated_over_target_pct"]):
    sim_col1, sim_col2 = st.columns(2)
    sim_col1.metric(
        f"Current > {tat_target:.0f}d TAT",
        f"{scenario_stats['current_over_target_pct']:.1f}%",
    )
    sim_col2.metric(
        f"Simulated > {tat_target:.0f}d TAT",
        f"{scenario_stats['simulated_over_target_pct']:.1f}%",
        delta=f"{scenario_stats['simulated_over_target_pct'] - scenario_stats['current_over_target_pct']:.1f} pts",
        delta_color="inverse",
    )

st.caption(
    "Assumptions: day-first date parsing is used; TAT(days)=completedDateTime-createDateTime-total holding time; missing/empty off-hold intervals are not counted."
)
