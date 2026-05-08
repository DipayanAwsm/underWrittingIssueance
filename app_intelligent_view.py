import re
import warnings
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="Underwriting Issuance - Intelligent View",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0f4c81;
        margin-bottom: 0.4rem;
    }
    .subtitle {
        color: #5c6770;
        margin-bottom: 1.3rem;
    }
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #ffffff 0%, #eef5ff 100%);
        border: 1px solid #cfe0ff;
        border-radius: 14px;
        padding: 0.8rem 0.9rem;
        box-shadow: 0 8px 16px rgba(15, 76, 129, 0.12);
    }
    div[data-testid="stMetricLabel"] p {
        color: #0f4c81;
        font-weight: 700;
        letter-spacing: 0.1px;
    }
    div[data-testid="stMetricValue"] {
        color: #0b253a;
        font-size: 1.7rem;
        font-weight: 800;
    }
    div[data-testid="stMetricDelta"] {
        font-weight: 650;
    }
    .hero-kpi {
        background: linear-gradient(150deg, #eef5ff 0%, #ffffff 100%);
        border: 1px solid #cfe0ff;
        border-radius: 14px;
        padding: 0.95rem 1rem;
        box-shadow: 0 8px 16px rgba(15, 76, 129, 0.12);
        min-height: 180px;
    }
    .hero-kpi-title {
        color: #0f4c81;
        font-weight: 700;
        font-size: 0.98rem;
        margin-bottom: 0.2rem;
    }
    .hero-kpi-value {
        color: #0b253a;
        font-weight: 800;
        font-size: 2rem;
        line-height: 1.2;
        margin-bottom: 0.45rem;
    }
    .hero-kpi-sub {
        color: #35556d;
        font-size: 0.9rem;
        margin-bottom: 0.2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def load_data(file_path_or_buffer):
    """Load CSV/Excel input."""
    try:
        if isinstance(file_path_or_buffer, str):
            if file_path_or_buffer.endswith(".csv"):
                return pd.read_csv(file_path_or_buffer, low_memory=False)
            if file_path_or_buffer.endswith((".xlsx", ".xls")):
                return pd.read_excel(file_path_or_buffer, engine="openpyxl")
            return None

        if file_path_or_buffer.name.endswith(".csv"):
            return pd.read_csv(file_path_or_buffer, low_memory=False)
        if file_path_or_buffer.name.endswith((".xlsx", ".xls")):
            return pd.read_excel(file_path_or_buffer, engine="openpyxl")
        return None
    except Exception as exc:
        st.error(f"Error loading file: {exc}")
        return None


@st.cache_data
def dataframe_to_csv_bytes(df):
    """Convert dataframe to CSV bytes for download."""
    export_df = df.copy()
    for list_col in ["Hold_Reasons_List", "Hold_Days_List"]:
        if list_col in export_df.columns:
            export_df[list_col] = export_df[list_col].apply(
                lambda v: "|".join(map(str, v)) if isinstance(v, list) else v
            )
    return export_df.to_csv(index=False).encode("utf-8")


def parse_separated_values(value):
    """Split values by |, comma-space, or comma."""
    if pd.isna(value):
        return []

    text = str(value).strip()
    if not text:
        return []

    if "|" in text:
        parts = text.split("|")
    elif ", " in text:
        parts = text.split(", ")
    elif "," in text:
        parts = text.split(",")
    else:
        return [text]

    return [p.strip() for p in parts if str(p).strip()]


def resolve_column(df, candidates):
    """Return the first matching column name from candidates."""
    for col in candidates:
        if col in df.columns:
            return col
    return None


def infer_case_type(hold_count):
    if hold_count <= 0:
        return "Straight Through"
    if hold_count == 1:
        return "One Touch"
    return f"Multi Hold ({int(hold_count)} touches)"


def create_tat_bucket(tat_days):
    if pd.isna(tat_days):
        return None
    if tat_days <= 4:
        return "1-4 days"
    if tat_days <= 7:
        return "5-7 days"
    return "7+ days"


def to_float_series(series):
    return pd.to_numeric(series, errors="coerce")


def parse_datetime_series(series):
    """Parse datetime safely and prefer day-first parse if it captures more values."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed_dayfirst = pd.to_datetime(series, errors="coerce", dayfirst=True)
        parsed_default = pd.to_datetime(series, errors="coerce")

    if parsed_dayfirst.notna().sum() > parsed_default.notna().sum():
        return parsed_dayfirst
    return parsed_default


def parse_datetime_value(value):
    """Parse a single datetime value with day-first preference."""
    if pd.isna(value):
        return pd.NaT
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        dt_dayfirst = pd.to_datetime(value, errors="coerce", dayfirst=True)
        dt_default = pd.to_datetime(value, errors="coerce")
    if pd.notna(dt_dayfirst):
        return dt_dayfirst
    return dt_default


def prepare_dataframe(df):
    """Prepare reusable analytic fields for the dashboard."""
    data = df.copy()

    create_col = resolve_column(data, ["createDateTime"])
    complete_col = resolve_column(data, ["completedDateTime"])
    tat_col = resolve_column(data, ["TAT", "Tat", "tat"])
    holds_col = resolve_column(data, ["No of Holds", "No_of_Holds", "numberOfHolds"])
    history_reason_col = resolve_column(data, ["onHoldReasonDescriptionsHistory"])
    hold_reason_col = resolve_column(data, ["onHoldReasonDescription"])
    on_hold_dates_col = resolve_column(data, ["onHoldDatesHistory"])
    off_hold_dates_col = resolve_column(data, ["offHoldDatesHistory"])

    if create_col:
        data["createDateTime"] = parse_datetime_series(data[create_col])
    else:
        data["createDateTime"] = pd.NaT

    if complete_col:
        data["completedDateTime"] = parse_datetime_series(data[complete_col])
    else:
        data["completedDateTime"] = pd.NaT

    tat_from_dates = (
        data["completedDateTime"] - data["createDateTime"]
    ).dt.total_seconds() / 86400.0

    if tat_col:
        fallback_tat = to_float_series(data[tat_col])
        data["TAT_Days"] = tat_from_dates.where(tat_from_dates.notna(), fallback_tat)
    else:
        data["TAT_Days"] = tat_from_dates

    data.loc[data["TAT_Days"] < 0, "TAT_Days"] = np.nan

    hold_counts = pd.Series(0, index=data.index, dtype="float64")

    if history_reason_col:
        hold_counts = data[history_reason_col].apply(lambda v: len(parse_separated_values(v))).astype("float64")

    if holds_col:
        numeric_holds = to_float_series(data[holds_col])
        hold_counts = hold_counts.where(hold_counts > 0, numeric_holds)

    if hold_reason_col:
        has_single_hold_reason = (
            data[hold_reason_col].fillna("").astype(str).str.strip().ne("")
        )
        hold_counts = hold_counts.where(hold_counts > 0, has_single_hold_reason.astype(int))

    data["Hold_Count"] = hold_counts.fillna(0).clip(lower=0)
    data["CaseType"] = data["Hold_Count"].apply(infer_case_type)
    data["TAT_Bucket"] = data["TAT_Days"].apply(create_tat_bucket)

    hold_reasons_list = []
    hold_days_list = []
    now_ts = pd.Timestamp.now()
    for _, row in data.iterrows():
        reasons = parse_separated_values(row.get(history_reason_col, "")) if history_reason_col else []
        if not reasons and hold_reason_col:
            primary_reason = row.get(hold_reason_col, "")
            if pd.notna(primary_reason) and str(primary_reason).strip():
                reasons = [str(primary_reason).strip()]

        on_hold_raw = parse_separated_values(row.get(on_hold_dates_col, "")) if on_hold_dates_col else []
        off_hold_raw = parse_separated_values(row.get(off_hold_dates_col, "")) if off_hold_dates_col else []

        on_hold_dates = [parse_datetime_value(v) for v in on_hold_raw]
        on_hold_dates = [d for d in on_hold_dates if pd.notna(d)]
        off_hold_dates = [parse_datetime_value(v) for v in off_hold_raw]
        off_hold_dates = [d for d in off_hold_dates if pd.notna(d)]

        aligned_reasons = []
        aligned_days = []
        event_count = max(len(reasons), len(on_hold_dates))
        for idx in range(event_count):
            on_dt = on_hold_dates[idx] if idx < len(on_hold_dates) else pd.NaT
            if pd.isna(on_dt):
                continue

            reason_val = reasons[idx] if idx < len(reasons) else (reasons[-1] if reasons else "Unknown")
            reason_val = str(reason_val).strip() if str(reason_val).strip() else "Unknown"

            off_dt = off_hold_dates[idx] if idx < len(off_hold_dates) else pd.NaT
            if pd.notna(off_dt):
                end_dt = off_dt
            elif pd.notna(row.get("completedDateTime")):
                end_dt = row.get("completedDateTime")
            else:
                end_dt = now_ts

            days = (end_dt - on_dt).total_seconds() / 86400.0
            if pd.notna(days) and days >= 0:
                aligned_reasons.append(reason_val)
                aligned_days.append(days)

        hold_reasons_list.append(aligned_reasons)
        hold_days_list.append(aligned_days)

    data["Hold_Reasons_List"] = hold_reasons_list
    data["Hold_Days_List"] = hold_days_list
    data["Total_Hold_Days"] = data["Hold_Days_List"].apply(lambda v: float(np.sum(v)) if v else 0.0)

    data["Month"] = data["createDateTime"].dt.to_period("M")
    data["Month_Str"] = data["Month"].astype(str)
    data.loc[data["createDateTime"].isna(), "Month_Str"] = np.nan

    if hold_reason_col:
        data["PrimaryHoldReason"] = data[hold_reason_col].fillna("Unknown").astype(str)
    elif history_reason_col:
        data["PrimaryHoldReason"] = data[history_reason_col].apply(
            lambda v: parse_separated_values(v)[0] if parse_separated_values(v) else "Unknown"
        )
    else:
        data["PrimaryHoldReason"] = "Unknown"

    return data


def apply_sidebar_filters(df):
    filtered = df.copy()

    st.sidebar.header("Data Source")
    st.sidebar.caption(f"Rows loaded: {len(df):,}")

    st.sidebar.header("Filters")

    if filtered["Month_Str"].notna().any():
        months = sorted(filtered["Month_Str"].dropna().unique().tolist())
        chosen_months = st.sidebar.multiselect(
            "Month-wise",
            options=months,
            default=months,
        )
        if chosen_months:
            filtered = filtered[filtered["Month_Str"].isin(chosen_months)]

    status_col = resolve_column(filtered, ["statusDescription"])
    if status_col:
        status_values = sorted(filtered[status_col].dropna().astype(str).unique().tolist())
        chosen_status = st.sidebar.multiselect(
            "Status",
            options=status_values,
            default=status_values,
        )
        if chosen_status:
            filtered = filtered[filtered[status_col].astype(str).isin(chosen_status)]

    req_col = resolve_column(filtered, ["requestTypeDescription"])
    if req_col:
        req_values = sorted(filtered[req_col].dropna().astype(str).unique().tolist())
        chosen_req = st.sidebar.multiselect(
            "Request Type",
            options=req_values,
            default=req_values,
        )
        if chosen_req:
            filtered = filtered[filtered[req_col].astype(str).isin(chosen_req)]

    lob_col = resolve_column(filtered, ["lineOfBusinessDescription"])
    if lob_col:
        lob_values = sorted(filtered[lob_col].dropna().astype(str).unique().tolist())
        chosen_lob = st.sidebar.multiselect(
            "Line Of Business",
            options=lob_values,
            default=lob_values,
        )
        if chosen_lob:
            filtered = filtered[filtered[lob_col].astype(str).isin(chosen_lob)]

    state_col = resolve_column(filtered, ["AgentBrokerStateCode"])
    if state_col:
        state_values = sorted(filtered[state_col].dropna().astype(str).unique().tolist())
        chosen_state = st.sidebar.multiselect(
            "Broker State",
            options=state_values,
            default=state_values,
        )
        if chosen_state:
            filtered = filtered[filtered[state_col].astype(str).isin(chosen_state)]

    case_type_filter = st.sidebar.radio(
        "Case Type",
        options=["All Cases", "Straight Through", "Hold Cases"],
        index=0,
    )
    if case_type_filter == "Straight Through":
        filtered = filtered[filtered["Hold_Count"] <= 0]
    elif case_type_filter == "Hold Cases":
        filtered = filtered[filtered["Hold_Count"] > 0]

    st.sidebar.caption(f"Rows after filters: {len(filtered):,}")
    return filtered


def top_counts(df, column_name, top_n=5):
    if column_name is None or column_name not in df.columns:
        return pd.DataFrame(columns=["Value", "Count"])

    series = (
        df[column_name]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
        .replace("", "Unknown")
    )
    counts = series.value_counts().head(top_n).reset_index()
    counts.columns = ["Value", "Count"]
    return counts


def build_top_table(df, col_candidates, title, top_n=5):
    col_name = resolve_column(df, col_candidates)
    table = top_counts(df, col_name, top_n=top_n)
    st.subheader(title)
    if table.empty:
        st.info("Column not available in current data.")
    else:
        st.dataframe(table, use_container_width=True, hide_index=True)


def build_top_table_with_tat(df, col_candidates, title, top_n=5):
    col_name = resolve_column(df, col_candidates)
    st.subheader(title)

    if col_name is None or col_name not in df.columns or "TAT_Days" not in df.columns:
        st.info("Column not available in current data.")
        return

    grouped = (
        df.assign(
            _group_val=df[col_name]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
            .replace("", "Unknown")
        )
        .groupby("_group_val")
        .agg(
            Number_of_Cases=("TAT_Days", "size"),
            Avg_TAT_Days=("TAT_Days", "mean"),
        )
        .reset_index()
        .sort_values("Number_of_Cases", ascending=False)
        .head(top_n)
    )

    grouped["Avg_TAT_Days"] = grouped["Avg_TAT_Days"].round(2)
    grouped = grouped.rename(columns={"_group_val": "Value"})
    st.dataframe(grouped, use_container_width=True, hide_index=True)


def build_hold_reason_impact(df):
    if "Hold_Reasons_List" not in df.columns or "Hold_Days_List" not in df.columns:
        return pd.DataFrame(columns=["onHoldReasonDescription", "Count", "Total_Hold_Days", "Avg_Hold_Days"])

    records = []
    for _, row in df.iterrows():
        reasons = row.get("Hold_Reasons_List", [])
        days = row.get("Hold_Days_List", [])
        if not isinstance(reasons, list) or not isinstance(days, list):
            continue
        for idx, reason in enumerate(reasons):
            if idx >= len(days):
                continue
            day_val = days[idx]
            if pd.isna(day_val):
                continue
            records.append({"onHoldReasonDescription": reason, "Hold_Days": float(day_val)})

    if not records:
        return pd.DataFrame(columns=["onHoldReasonDescription", "Count", "Total_Hold_Days", "Avg_Hold_Days"])

    reason_df = pd.DataFrame(records)
    impact_df = (
        reason_df.groupby("onHoldReasonDescription")
        .agg(
            Count=("Hold_Days", "size"),
            Total_Hold_Days=("Hold_Days", "sum"),
            Avg_Hold_Days=("Hold_Days", "mean"),
        )
        .reset_index()
        .sort_values("Total_Hold_Days", ascending=False)
    )
    impact_df["Total_Hold_Days"] = impact_df["Total_Hold_Days"].round(2)
    impact_df["Avg_Hold_Days"] = impact_df["Avg_Hold_Days"].round(2)
    return impact_df


def suggest_action_for_reason(reason):
    text = str(reason).lower()

    rules = [
        (r"document|missing|incomplete|info", "Add a pre-submission document checklist and broker self-validation gate."),
        (r"validation|error|edit", "Introduce automated validation at intake with immediate correction prompts."),
        (r"system|issue|outage|technical", "Create an IT fast-lane escalation and fallback manual processing SOP."),
        (r"underwriting|review|approval", "Set SLA-based underwriting queue triage with daily aging review."),
        (r"write[- ]?out", "Standardize write-out templates and approval pathways to reduce waiting time."),
        (r"premium|pricing|rate", "Publish pricing exception rules and assign same-day approvers for edge cases."),
        (r"claim|loss", "Pre-fetch claim/loss data before assignment to avoid downstream hold cycles."),
    ]

    for pattern, action in rules:
        if re.search(pattern, text):
            return action

    return "Run weekly root-cause review, assign owner, and track closure date for repeat reasons."


def show_tat_bucket_distribution(completed_df):
    bucket_order = ["1-4 days", "5-7 days", "7+ days"]
    color_map = {
        "1-4 days": "#2E7D32",
        "5-7 days": "#FB8C00",
        "7+ days": "#D32F2F",
    }

    bucket_counts = completed_df["TAT_Bucket"].value_counts().reindex(bucket_order, fill_value=0)
    bucket_df = bucket_counts.reset_index()
    bucket_df.columns = ["TAT_Bucket", "Cases"]

    fig = px.bar(
        bucket_df,
        x="TAT_Bucket",
        y="Cases",
        color="TAT_Bucket",
        title="TAT Bucket Distribution (Overall Completed Cases)",
        color_discrete_map=color_map,
        category_orders={"TAT_Bucket": bucket_order},
    )
    fig.update_layout(showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(bucket_df, use_container_width=True, hide_index=True)


def show_monthly_tat_buckets(completed_df):
    bucket_order = ["1-4 days", "5-7 days", "7+ days"]
    color_map = {
        "1-4 days": "#2E7D32",
        "5-7 days": "#FB8C00",
        "7+ days": "#D32F2F",
    }

    monthly_bucket = (
        completed_df.dropna(subset=["Month_Str", "TAT_Bucket"]) 
        .groupby(["Month_Str", "TAT_Bucket"]) 
        .size() 
        .reset_index(name="Cases")
    )
    monthly_bucket = monthly_bucket.sort_values("Month_Str")

    if monthly_bucket.empty:
        st.info("No monthly completed-case trend available.")
        return

    fig = px.area(
        monthly_bucket,
        x="Month_Str",
        y="Cases",
        color="TAT_Bucket",
        title="Primary TAT Buckets Over Time (Month-wise)",
        category_orders={"TAT_Bucket": bucket_order},
        color_discrete_map=color_map,
    )
    fig.update_xaxes(tickangle=45)
    st.plotly_chart(fig, use_container_width=True)


def render_tab_one(filtered_df):
    st.subheader("TAT Overview and 7+ Days Drilldown")

    left_col, right_col = st.columns([6, 4])
    tab1_df = filtered_df.copy()
    tab1_completed = tab1_df[tab1_df["TAT_Days"].notna()].copy()

    with left_col:
        with st.container(border=True):
            st.markdown("### TAT Overview")
            st.caption("Month-wise filter is applied from the left sidebar with other filters.")
            month_case_counts = (
                filtered_df.dropna(subset=["Month_Str"])
                .groupby("Month_Str")
                .size()
                .sort_index()
            )

            st.markdown("#### KPI Tiles")
            mean_tat = tab1_completed["TAT_Days"].mean()
            total_cases = len(tab1_df)

            current_count = len(tab1_df)
            delta_text = "No previous month"
            current_month_label = "Multiple/All"
            if len(month_case_counts) > 0:
                current_month_label = month_case_counts.index[-1]
                current_count = int(month_case_counts.iloc[-1])
                if len(month_case_counts) >= 2:
                    prev_month = month_case_counts.index[-2]
                    prev_count = int(month_case_counts.iloc[-2])
                    abs_change = current_count - prev_count
                    pct_change = (abs_change / prev_count * 100) if prev_count else np.nan
                    if pd.notna(pct_change):
                        delta_text = f"{abs_change:+,} ({pct_change:+.1f}%) vs {prev_month}"

            bucket_order = ["1-4 days", "5-7 days", "7+ days"]
            color_map = {
                "1-4 days": "#2E7D32",
                "5-7 days": "#FB8C00",
                "7+ days": "#D32F2F",
            }
            bucket_counts = (
                tab1_completed["TAT_Bucket"]
                .value_counts()
                .reindex(bucket_order, fill_value=0)
            )
            total_bucket_cases = int(bucket_counts.sum())

            tile1, tile2 = st.columns(2)
            with tile1:
                with st.container(border=True):
                    mean_tat_text = f"{mean_tat:.2f} days" if pd.notna(mean_tat) else "N/A"
                    st.markdown(
                        f"""
                        <div class="hero-kpi-title">Aggregated Mean TAT</div>
                        <div class="hero-kpi-value">{mean_tat_text}</div>
                        <div class="hero-kpi-sub"><b>Total Number of Cases:</b> {total_cases:,}</div>
                        <div class="hero-kpi-sub"><b>MoM Change:</b> {delta_text}</div>
                        <div class="hero-kpi-sub"><b>Reference Month:</b> {current_month_label}</div>
                        """,
                        unsafe_allow_html=True,
                    )

            with tile2:
                with st.container(border=True):
                    st.markdown("##### Tat bucket distribution")
                    if total_bucket_cases <= 0:
                        st.info("No completed cases available for bucket segmentation.")
                    else:
                        segment_df = pd.DataFrame(
                            {
                                "TAT_Bucket": bucket_order,
                                "Count": [int(bucket_counts.get(b, 0)) for b in bucket_order],
                            }
                        )
                        segment_df["Percentage"] = segment_df["Count"] / total_bucket_cases * 100
                        segment_df["Pct_Text"] = segment_df["Percentage"].map(lambda v: f"{v:.1f}%")

                        fig_segment = go.Figure()
                        for _, row in segment_df.iterrows():
                            bucket = row["TAT_Bucket"]
                            pct = float(row["Percentage"])
                            fig_segment.add_trace(
                                go.Bar(
                                    x=[pct],
                                    y=[""],
                                    orientation="h",
                                    name=bucket,
                                    marker_color=color_map[bucket],
                                    text=[f"{pct:.1f}%" if pct > 0 else ""],
                                    textposition="inside",
                                    insidetextanchor="middle",
                                    hovertemplate=f"TAT Bucket: {bucket}<br>Percentage: {pct:.2f}%<extra></extra>",
                                )
                            )
                        fig_segment.update_layout(
                            barmode="stack",
                            showlegend=True,
                            height=220,
                            margin=dict(l=4, r=4, t=8, b=4),
                            xaxis_title="",
                            yaxis_title="",
                            legend_title_text="TAT Bucket",
                            paper_bgcolor="rgba(0,0,0,0)",
                            plot_bgcolor="rgba(0,0,0,0)",
                        )
                        fig_segment.update_xaxes(range=[0, 100])
                        fig_segment.update_yaxes(showticklabels=False, showgrid=False, zeroline=False)
                        st.plotly_chart(fig_segment, use_container_width=True)

            tile3, tile4 = st.columns(2)
            with tile3:
                open_cases = int(tab1_df["completedDateTime"].isna().sum())
                st.metric("Open Cases", f"{open_cases:,}")
            with tile4:
                closed_cases = int(tab1_df["completedDateTime"].notna().sum())
                st.metric("Closed Cases", f"{closed_cases:,}")

            st.markdown("#### TAT Bucket Distribution Over Time")
            monthly_bucket = (
                tab1_completed.dropna(subset=["Month_Str", "TAT_Bucket"])
                .groupby(["Month_Str", "TAT_Bucket"])
                .size()
                .reset_index(name="Count")
                .sort_values("Month_Str")
            )
            if monthly_bucket.empty:
                st.info("No completed-case monthly bucket data available.")
            else:
                fig_bucket_time = px.bar(
                    monthly_bucket,
                    x="Month_Str",
                    y="Count",
                    color="TAT_Bucket",
                    barmode="stack",
                    color_discrete_map=color_map,
                    category_orders={"TAT_Bucket": bucket_order},
                    title="TAT Bucket Distribution by Month",
                )
                fig_bucket_time.update_xaxes(tickangle=45)
                st.plotly_chart(fig_bucket_time, use_container_width=True)

            st.markdown("#### On-Hold Reason Impact in Days")
            reason_impact = build_hold_reason_impact(tab1_df)
            metric_cols = st.columns(2)
            total_reason_instances = int(reason_impact["Count"].sum()) if not reason_impact.empty else 0
            total_hold_days = float(reason_impact["Total_Hold_Days"].sum()) if not reason_impact.empty else 0.0
            metric_cols[0].metric("Total On-Hold Reason Occurrences", f"{total_reason_instances:,}")
            metric_cols[1].metric("Total Hold Days Added/Taken", f"{total_hold_days:,.2f}")

            if reason_impact.empty:
                st.info("No on-hold reason day contribution data available.")
            else:
                st.dataframe(reason_impact.head(10), use_container_width=True, hide_index=True)

    with right_col:
        with st.container(border=True):
            st.markdown("### 7+ Days Drilldown")
            with st.expander("Open 7+ Days Drilldown", expanded=False):
                high_tat = tab1_completed[tab1_completed["TAT_Bucket"] == "7+ days"].copy()

                if high_tat.empty:
                    st.info("No 7+ day completed cases available for selected filters/month.")
                    return

                top_cols = st.columns(3)
                top_cols[0].metric("7+ Day Cases", f"{len(high_tat):,}")
                top_cols[1].metric("Max TAT", f"{high_tat['TAT_Days'].max():.1f} days")
                top_cols[2].metric("Avg Holds", f"{high_tat['Hold_Count'].mean():.2f}")

                build_top_table_with_tat(
                    high_tat,
                    ["onHoldReasonDescription", "PrimaryHoldReason"],
                    "Top 5 onHoldReasonDescription",
                )
                build_top_table_with_tat(
                    high_tat,
                    ["bgiDescription"],
                    "Top 5 bgiDescription",
                )
                build_top_table(high_tat, ["lineOfBusinessDescription"], "Top 5 lineOfBusinessDescription")
                build_top_table(high_tat, ["AgentBrokerStateCode"], "Top 5 AgentBrokerStateCode")

                st.subheader("Top 5 High TAT Cases (Decreasing Order)")
                state_col = resolve_column(high_tat, ["AgentBrokerStateCode"])
                broker_num_col = resolve_column(high_tat, ["agentBrokerNum"])

                show_cols = [c for c in [broker_num_col, "TAT_Days", state_col] if c and c in high_tat.columns]

                if show_cols:
                    case_table = high_tat[show_cols].sort_values("TAT_Days", ascending=False).head(5).copy()
                    rename_map = {
                        broker_num_col: "agentBrokerNum",
                        "TAT_Days": "TAT",
                        state_col: "AgentBrokerStateCode",
                    }
                    case_table = case_table.rename(columns=rename_map)
                    st.dataframe(case_table, use_container_width=True, hide_index=True)
                else:
                    st.info("Required columns for high-TAT case listing are unavailable.")


def render_tab_two(filtered_df, completed_df):
    st.subheader("Hold-Centric Intelligence")

    # (a) Number of holds and distribution over TAT bucket over time
    monthly_holds = (
        completed_df.dropna(subset=["Month_Str", "TAT_Bucket"])
        .groupby(["Month_Str", "TAT_Bucket"])
        .agg(
            Avg_Holds=("Hold_Count", "mean"),
            Cases=("TAT_Days", "size"),
        )
        .reset_index()
        .sort_values("Month_Str")
    )

    if not monthly_holds.empty:
        trend_col, heat_col = st.columns(2)
        with trend_col:
            with st.container(border=True):
                st.markdown("### Holds Trend by TAT Bucket")
                fig_holds_trend = px.line(
                    monthly_holds,
                    x="Month_Str",
                    y="Avg_Holds",
                    color="TAT_Bucket",
                    markers=True,
                    title="Average Number of Holds Over Time by TAT Bucket",
                )
                fig_holds_trend.update_xaxes(tickangle=45)
                st.plotly_chart(fig_holds_trend, use_container_width=True)

        with heat_col:
            with st.container(border=True):
                st.markdown("### Holds Distribution Heatmap")
                pivot = monthly_holds.pivot(
                    index="TAT_Bucket", columns="Month_Str", values="Avg_Holds"
                )
                pivot = pivot.reindex(["1-4 days", "5-7 days", "7+ days"])

                fig_heat = px.imshow(
                    pivot,
                    text_auto=".2f",
                    aspect="auto",
                    color_continuous_scale="YlOrRd",
                    title="Holds Distribution Over TAT Buckets by Month",
                    labels=dict(x="Month", y="TAT Bucket", color="Avg Holds"),
                )
                st.plotly_chart(fig_heat, use_container_width=True)
    else:
        st.info("Not enough completed case data for hold-vs-TAT monthly distribution.")

    st.markdown("---")

    # (b) Hold reason distribution for on-hold cases
    status_col = resolve_column(filtered_df, ["statusDescription"])
    on_hold_mask = filtered_df["Hold_Count"] > 0
    if status_col:
        on_hold_mask = on_hold_mask | filtered_df[status_col].astype(str).str.contains("hold", case=False, na=False)

    on_hold_cases = filtered_df[on_hold_mask].copy()
    reason_col = resolve_column(on_hold_cases, ["onHoldReasonDescription", "PrimaryHoldReason"])

    reason_col1, reason_col2 = st.columns([1.4, 1])
    with reason_col1:
        with st.container(border=True):
            st.subheader("On-Hold Cases: Hold Reason Distribution")
            top_reasons = top_counts(on_hold_cases, reason_col, top_n=15)
            if top_reasons.empty:
                st.info("Hold reason column is unavailable.")
            else:
                fig_reason = px.bar(
                    top_reasons,
                    x="Value",
                    y="Count",
                    title="Top Hold Reasons (On-Hold Cases)",
                    labels={"Value": "Hold Reason", "Count": "Cases"},
                )
                fig_reason.update_xaxes(tickangle=45)
                st.plotly_chart(fig_reason, use_container_width=True)

    with reason_col2:
        st.subheader("Top Brokers Handling Cases")
        broker_name_col = resolve_column(
            filtered_df,
            ["AgentBrokerName", "AgentBrokerName__2", "agentBrokerNum"],
        )
        top_brokers = top_counts(filtered_df, broker_name_col, top_n=10)
        if top_brokers.empty:
            st.info("Broker column unavailable.")
        else:
            st.dataframe(top_brokers, use_container_width=True, hide_index=True)

    st.markdown("---")

    # (c) Top accountAnalyst and raterFullName taking 7+ days
    high_tat = completed_df[completed_df["TAT_Bucket"] == "7+ days"].copy()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Top accountAnalyst Taking 7+ TAT Days")
        analyst_col = resolve_column(high_tat, ["accountAnalyst", "accountAnalystName"])
        top_analyst = top_counts(high_tat, analyst_col, top_n=10)
        if top_analyst.empty:
            st.info("accountAnalyst column unavailable.")
        else:
            st.dataframe(top_analyst, use_container_width=True, hide_index=True)

    with c2:
        st.subheader("Top UnderWriter Taking 7+ TAT Days")
        rater_col = resolve_column(high_tat, ["raterFullName"])
        top_rater = top_counts(high_tat, rater_col, top_n=10)
        if top_rater.empty:
            st.info("raterFullName column unavailable.")
        else:
            st.dataframe(top_rater, use_container_width=True, hide_index=True)

    st.markdown("---")

    # (d) Last month cases with highest TAT (top 5)
    st.subheader("Last Month: Highest TAT Cases (Top 5)")
    month_series = completed_df["Month_Str"].dropna().sort_values()
    if month_series.empty:
        st.info("No month data available for last-month high TAT cases.")
    else:
        last_month = month_series.iloc[-1]
        lm_df = completed_df[completed_df["Month_Str"] == last_month].copy()

        cols_requested = [
            resolve_column(lm_df, ["requestId"]),
            resolve_column(lm_df, ["processId"]),
            resolve_column(lm_df, ["AgentBrokerName", "AgentBrokerName__2"]),
            resolve_column(lm_df, ["accountAnalyst", "accountAnalystName"]),
            "TAT_Days",
            "TAT_Bucket",
        ]
        cols_requested = [c for c in cols_requested if c and c in lm_df.columns]

        if cols_requested:
            high_tat_lm = lm_df.sort_values("TAT_Days", ascending=False).head(5)[cols_requested].copy()
            rename_map = {
                "TAT_Days": "TAT",
                resolve_column(lm_df, ["requestId"]): "requestId",
                resolve_column(lm_df, ["processId"]): "processId",
                resolve_column(lm_df, ["AgentBrokerName", "AgentBrokerName__2"]): "AgentBrokerName",
                resolve_column(lm_df, ["accountAnalyst", "accountAnalystName"]): "accountAnalyst",
            }
            high_tat_lm = high_tat_lm.rename(columns=rename_map)
            st.caption(f"Last month in filtered data: {last_month}")
            st.dataframe(high_tat_lm, use_container_width=True, hide_index=True)
        else:
            st.info("Requested case-level columns not available.")

    st.markdown("---")

    # (e) Top on-hold reason and suggested actions
    st.subheader("Prescriptive Actions for Top Hold Reasons")
    reason_summary = top_counts(on_hold_cases, reason_col, top_n=10)
    if reason_summary.empty:
        st.info("Unable to generate hold reason actions because hold reason data is missing.")
    else:
        reason_summary["Recommended Action"] = reason_summary["Value"].apply(suggest_action_for_reason)
        reason_summary = reason_summary.rename(columns={"Value": "onHoldReasonDescription"})
        st.dataframe(reason_summary, use_container_width=True, hide_index=True)


def main():
    st.markdown('<div class="main-title">Underwriting Issuance - Intelligent View</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitle">Reference logic retained, view redesigned into two decision tabs with 7+ day deep-dive and hold intelligence.</div>',
        unsafe_allow_html=True,
    )

    st.sidebar.header("Input")
    data_source = st.sidebar.radio("Choose data source", ["Upload File", "Data Folder"], index=1)

    source_df = None

    if data_source == "Upload File":
        uploaded_file = st.sidebar.file_uploader("Upload CSV/Excel", type=["csv", "xlsx", "xls"])
        if uploaded_file is not None:
            source_df = load_data(uploaded_file)
    else:
        data_folder = Path("data")
        files = []
        if data_folder.exists():
            files = sorted(list(data_folder.glob("*.csv")) + list(data_folder.glob("*.xlsx")) + list(data_folder.glob("*.xls")))

        if files:
            selected_file = st.sidebar.selectbox("Select file from data folder", options=[f.name for f in files])
            if selected_file:
                source_df = load_data(str(data_folder / selected_file))
        else:
            st.sidebar.warning("No CSV/Excel files found in data folder.")

    if source_df is None or source_df.empty:
        st.info("Please upload a file or select one from the data folder.")
        return

    with st.spinner("Preparing data..."):
        prepared_df = prepare_dataframe(source_df)
        filtered_df = apply_sidebar_filters(prepared_df)

    if filtered_df.empty:
        st.warning("No records left after filters. Please widen filter selections.")
        return

    completed_df = filtered_df[filtered_df["TAT_Days"].notna()].copy()

    st.sidebar.header("CSV Report")
    report_filename = f"underwriting_filtered_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    st.sidebar.download_button(
        label="Download Filtered Report (CSV)",
        data=dataframe_to_csv_bytes(filtered_df),
        file_name=report_filename,
        mime="text/csv",
    )

    tab1, tab2 = st.tabs(["Tab 1 - TAT Overview", "Tab 2 - Hold Intelligence"])

    with tab1:
        render_tab_one(filtered_df)

    with tab2:
        render_tab_two(filtered_df, completed_df)


if __name__ == "__main__":
    main()
