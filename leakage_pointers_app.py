import io
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(page_title="Auto Issuance Leakage Pointers", layout="wide")

DEFAULT_FILE = Path("/Users/rituparnapaldas/Downloads/auto_issuance_synthetic_1year_10000rows.csv")
MISSING_TOKENS = {"", "nan", "none", "null", "na", "n/a", "-", "[]"}
DATE_TOKEN_PATTERN = re.compile(r"\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}")
TAT_BUCKET_ORDER = ["1-4 days", "5-7 days", "7+ days"]
TAT_BUCKET_COLORS = {
    "1-4 days": "#2ca02c",  # green
    "5-7 days": "#FFBF00",  # amber
    "7+ days": "#d62728",   # red
    "Unknown": "#9e9e9e",
}


def normalize_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    norm_map = {normalize_name(col): col for col in df.columns}
    for cand in candidates:
        key = normalize_name(cand)
        if key in norm_map:
            return norm_map[key]
    return None


def clean_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().strip("[]")
    if not text or text.lower() in MISSING_TOKENS:
        return ""
    return text


def clean_tokens(tokens: List[str]) -> List[str]:
    out: List[str] = []
    for token in tokens:
        item = str(token).strip()
        if not item or item.lower() in MISSING_TOKENS:
            continue
        out.append(item)
    return out


def parse_date_history_cell(value: object) -> List[str]:
    text = clean_text(value)
    if not text:
        return []
    if "|" in text:
        return clean_tokens(text.split("|"))
    if "," in text:
        matches = DATE_TOKEN_PATTERN.findall(text)
        if matches:
            return clean_tokens(matches)
        return clean_tokens(re.split(r"\s*,\s*", text))
    return clean_tokens([text])


def parse_reason_history_cell(value: object, expected_count: Optional[int] = None) -> List[str]:
    text = clean_text(value)
    if not text:
        return []
    if "|" in text:
        return clean_tokens(text.split("|"))
    if "," in text:
        if expected_count is not None and expected_count <= 1:
            return clean_tokens([text])
        parts = clean_tokens(re.split(r"\s*,\s*", text))
        if expected_count is not None and expected_count > 1 and len(parts) != expected_count:
            return clean_tokens([text])
        return parts
    return clean_tokens([text])


def parse_dt(value: object) -> pd.Timestamp:
    text = clean_text(value)
    if not text:
        return pd.NaT
    dt = pd.to_datetime(text, errors="coerce")
    if pd.isna(dt):
        dt = pd.to_datetime(text, errors="coerce", dayfirst=True)
    return dt


def parse_datetime_series(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce")
    reparsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return parsed.fillna(reparsed)


def short_reason(reason: str) -> str:
    text = str(reason).strip()
    if not text:
        return "Unspecified"
    return text.split(",")[0].strip() if "," in text else text


def calc_case_hold_metrics(
    on_hold_text: object,
    off_hold_text: object,
    reason_text: object,
    create_dt: pd.Timestamp,
) -> Tuple[float, int, int, bool, bool, bool, float]:
    on_values = parse_date_history_cell(on_hold_text)
    off_values = parse_date_history_cell(off_hold_text)
    reason_values = parse_reason_history_cell(reason_text, expected_count=len(on_values) if on_values else None)

    total_hold_days = 0.0
    first_on_dt = pd.NaT

    for idx, on_value in enumerate(on_values):
        on_dt = parse_dt(on_value)
        if pd.isna(on_dt):
            continue
        if pd.isna(first_on_dt) or on_dt < first_on_dt:
            first_on_dt = on_dt

        off_dt = parse_dt(off_values[idx]) if idx < len(off_values) else pd.NaT
        if pd.isna(off_dt):
            continue
        hold_days = (off_dt - on_dt).total_seconds() / 86400
        if hold_days < 0:
            continue
        total_hold_days += hold_days

    hold_reason_count = len(reason_values)
    touches_count = hold_reason_count + 1
    straight_through = hold_reason_count == 0
    multi_touch = touches_count > 1
    multi_hold = hold_reason_count >= 2

    first_hold_delay = np.nan
    if not pd.isna(first_on_dt) and not pd.isna(create_dt):
        delay_days = (first_on_dt - create_dt).total_seconds() / 86400
        if delay_days >= 0:
            first_hold_delay = delay_days

    return (
        total_hold_days,
        hold_reason_count,
        touches_count,
        straight_through,
        multi_touch,
        multi_hold,
        first_hold_delay,
    )


@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, file_name: str, delimiter: str) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)
    lower_name = file_name.lower()

    if lower_name.endswith((".xlsx", ".xls")):
        return pd.read_excel(bio, dtype=str)

    if delimiter == "auto":
        if lower_name.endswith(".csv"):
            return pd.read_csv(bio, sep=",", dtype=str, on_bad_lines="skip")
        return pd.read_csv(bio, sep=None, engine="python", dtype=str, on_bad_lines="skip")
    if delimiter == "tab":
        return pd.read_csv(bio, sep="\t", dtype=str, on_bad_lines="skip")
    return pd.read_csv(bio, sep=delimiter, dtype=str, on_bad_lines="skip")


@st.cache_data(show_spinner=False)
def prepare_data(raw_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Optional[str]]]:
    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    request_col = find_column(df, ["requestId", "request_id"])
    create_col = find_column(df, ["createDateTime", "create_date_time"])
    completed_col = find_column(df, ["completedDateTime", "completed_date_time"])
    status_col = find_column(df, ["statusDescription", "status_description"])
    request_desc_col = find_column(df, ["requestTypeDescription", "requestTypeCode", "requestType"])
    on_hold_col = find_column(df, ["onHoldDatesHistory"])
    off_hold_col = find_column(df, ["offHoldDatesHistory"])
    hold_reason_hist_col = find_column(df, ["onHoldReasonDescriptionsHistory"])
    writeout_reason_col = find_column(
        df,
        ["writeOutReasonDescription", "writeOutReasonDescriptionsHistory", "writeOutDescriptions"],
    )
    underwriter_col = find_column(df, ["underwriterName", "underwriter"])
    analyst_col = find_column(df, ["accountAnalystName", "accountAnalyst"])
    locations_col = find_column(df, ["numberOfLocations", "numberOfLocations__2"])
    vehicles_col = find_column(df, ["NumberOfVehicles", "numberOfVehicles"])

    if request_col is None:
        request_col = "__request_id"
        df[request_col] = [f"REQ_{idx+1}" for idx in range(len(df))]
    df["request_id"] = df[request_col].astype(str)

    if create_col is not None:
        df["create_dt"] = parse_datetime_series(df[create_col])
    else:
        df["create_dt"] = pd.NaT

    if completed_col is not None:
        df["completed_dt"] = parse_datetime_series(df[completed_col])
    else:
        df["completed_dt"] = pd.NaT

    df["status_value"] = (
        df[status_col].fillna("Unknown").astype(str).str.strip()
        if status_col is not None
        else "Unknown"
    )
    df["status_value"] = df["status_value"].replace("", "Unknown")

    df["request_desc_value"] = (
        df[request_desc_col].fillna("Unknown").astype(str).str.strip()
        if request_desc_col is not None
        else "Unknown"
    )
    df["request_desc_value"] = df["request_desc_value"].replace("", "Unknown")

    if hold_reason_hist_col is not None:
        df["on_hold_reason_value"] = df[hold_reason_hist_col].apply(
            lambda x: short_reason(
                parse_reason_history_cell(x, expected_count=1)[0]
                if len(parse_reason_history_cell(x, expected_count=1)) > 0
                else "No Hold Reason"
            )
        )
    else:
        df["on_hold_reason_value"] = "No Hold Reason"

    if writeout_reason_col is not None:
        df["writeout_reason_value"] = df[writeout_reason_col].apply(
            lambda x: short_reason(
                parse_reason_history_cell(x, expected_count=1)[0]
                if len(parse_reason_history_cell(x, expected_count=1)) > 0
                else "No Write-Out Reason"
            )
        )
    else:
        df["writeout_reason_value"] = "No Write-Out Reason"

    df["underwriter_value"] = (
        df[underwriter_col].fillna("Unassigned").astype(str).str.strip()
        if underwriter_col is not None
        else "Unassigned"
    )
    df["underwriter_value"] = df["underwriter_value"].replace("", "Unassigned")

    df["analyst_value"] = (
        df[analyst_col].fillna("Unassigned").astype(str).str.strip()
        if analyst_col is not None
        else "Unassigned"
    )
    df["analyst_value"] = df["analyst_value"].replace("", "Unassigned")

    df["number_of_locations"] = (
        pd.to_numeric(df[locations_col], errors="coerce")
        if locations_col is not None
        else np.nan
    )
    df["number_of_vehicles"] = (
        pd.to_numeric(df[vehicles_col], errors="coerce")
        if vehicles_col is not None
        else np.nan
    )

    if create_col is not None:
        df["create_month_dt"] = df["create_dt"].dt.to_period("M").dt.to_timestamp()
        df["create_month"] = df["create_month_dt"].dt.strftime("%Y-%m")
    else:
        df["create_month_dt"] = pd.NaT
        df["create_month"] = "Unknown"

    # Case-level hold metrics
    if all(col is not None for col in [on_hold_col, off_hold_col, hold_reason_hist_col]):
        hold_metrics = df.apply(
            lambda row: calc_case_hold_metrics(
                row[on_hold_col],
                row[off_hold_col],
                row[hold_reason_hist_col],
                row["create_dt"],
            ),
            axis=1,
            result_type="expand",
        )
        hold_metrics.columns = [
            "total_hold_days",
            "hold_reason_count",
            "touches_count",
            "straight_through",
            "multi_touch",
            "multi_hold",
            "first_hold_delay_days",
        ]
        df = pd.concat([df, hold_metrics], axis=1)
    else:
        df["total_hold_days"] = 0.0
        df["hold_reason_count"] = 0
        df["touches_count"] = 1
        df["straight_through"] = True
        df["multi_touch"] = False
        df["multi_hold"] = False
        df["first_hold_delay_days"] = np.nan

    df["total_hold_days"] = pd.to_numeric(df["total_hold_days"], errors="coerce").fillna(0.0).clip(lower=0)
    df["hold_reason_count"] = pd.to_numeric(df["hold_reason_count"], errors="coerce").fillna(0).astype(int)
    df["touches_count"] = pd.to_numeric(df["touches_count"], errors="coerce").fillna(1).clip(lower=1)
    df["case_type"] = np.where(df["straight_through"], "StraightThrough", "MultiTouch")

    # Completion and TAT
    df["is_completed"] = df["completed_dt"].notna()
    gross_tat = (df["completed_dt"] - df["create_dt"]).dt.total_seconds() / 86400
    gross_tat = gross_tat.where(gross_tat >= 0, np.nan)
    df["gross_tat_days"] = gross_tat
    net_tat = gross_tat - df["total_hold_days"]
    net_tat = net_tat.where(net_tat >= 0, np.nan)
    df["net_tat_days"] = net_tat

    df["tat_bucket"] = pd.cut(
        df["net_tat_days"],
        bins=[0, 4, 7, np.inf],
        labels=["1-4 days", "5-7 days", "7+ days"],
        include_lowest=True,
    )
    df["hold_bucket"] = pd.cut(
        df["total_hold_days"],
        bins=[-0.001, 0, 1, 3, 7, np.inf],
        labels=["0 days", "0-1", "1-3", "3-7", "7+"],
        include_lowest=True,
    )

    today = pd.Timestamp.today().normalize()
    open_mask = ~df["is_completed"]
    df["aging_open_days"] = np.nan
    df.loc[open_mask, "aging_open_days"] = (today - df.loc[open_mask, "create_dt"]).dt.total_seconds() / 86400
    df.loc[open_mask, "aging_open_days"] = df.loc[open_mask, "aging_open_days"].where(
        df.loc[open_mask, "aging_open_days"] >= 0, np.nan
    )

    # Hold-segment table for reason analysis
    hold_rows = []
    if all(col is not None for col in [on_hold_col, off_hold_col, hold_reason_hist_col]):
        base_cols = [request_col, "create_month", "is_completed", "underwriter_value", "analyst_value", "request_desc_value"]
        work_df = df[base_cols + [on_hold_col, off_hold_col, hold_reason_hist_col]]

        for _, row in work_df.iterrows():
            on_vals = parse_date_history_cell(row[on_hold_col])
            off_vals = parse_date_history_cell(row[off_hold_col])
            reason_vals = parse_reason_history_cell(
                row[hold_reason_hist_col], expected_count=len(on_vals) if on_vals else None
            )
            for idx, on_token in enumerate(on_vals):
                on_dt = parse_dt(on_token)
                off_dt = parse_dt(off_vals[idx]) if idx < len(off_vals) else pd.NaT
                if pd.isna(on_dt) or pd.isna(off_dt):
                    continue
                hold_days = (off_dt - on_dt).total_seconds() / 86400
                if hold_days < 0:
                    continue
                reason = reason_vals[idx] if idx < len(reason_vals) and reason_vals[idx] else "Unspecified"
                hold_rows.append(
                    {
                        "request_id": row[request_col],
                        "create_month": row["create_month"],
                        "is_completed": row["is_completed"],
                        "underwriter_value": row["underwriter_value"],
                        "analyst_value": row["analyst_value"],
                        "request_desc_value": row["request_desc_value"],
                        "hold_reason": reason,
                        "hold_reason_short": short_reason(reason),
                        "hold_days": hold_days,
                    }
                )

    hold_segments = pd.DataFrame(hold_rows)
    if not hold_segments.empty:
        hold_segments["request_id"] = hold_segments["request_id"].astype(str)

    metadata = {
        "request_col": request_col,
        "create_col": create_col,
        "completed_col": completed_col,
        "status_col": status_col,
        "request_desc_col": request_desc_col,
        "on_hold_col": on_hold_col,
        "off_hold_col": off_hold_col,
        "hold_reason_hist_col": hold_reason_hist_col,
        "writeout_reason_col": writeout_reason_col,
        "underwriter_col": underwriter_col,
        "analyst_col": analyst_col,
        "locations_col": locations_col,
        "vehicles_col": vehicles_col,
    }

    return df, hold_segments, metadata


def filter_data(
    df: pd.DataFrame,
    date_range: Optional[Tuple[pd.Timestamp, pd.Timestamp]],
    statuses: List[str],
    case_segments: List[str],
) -> pd.DataFrame:
    out = df.copy()
    if date_range is not None:
        start_dt, end_dt = date_range
        out = out[out["create_dt"].between(start_dt, end_dt, inclusive="both")]
    if statuses:
        out = out[out["status_value"].isin(statuses)]

    if case_segments:
        segment_mask = pd.Series(False, index=out.index)
        if "StraightThrough" in case_segments:
            segment_mask = segment_mask | out["straight_through"]
        if "MultiTouch" in case_segments:
            segment_mask = segment_mask | out["multi_touch"]
        if "MultiHold" in case_segments:
            segment_mask = segment_mask | out["multi_hold"]
        out = out[segment_mask]

    return out


def pct_stats(series: pd.Series) -> Tuple[float, float, float]:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return np.nan, np.nan, np.nan
    return float(s.median()), float(s.quantile(0.5)), float(s.quantile(0.9))


def tat_bucket_label(value: float) -> str:
    if value is None or pd.isna(value):
        return "Unknown"
    if value > 7:
        return "7+ days"
    if value >= 5:
        return "5-7 days"
    return "1-4 days"


def owner_summary(segment_df: pd.DataFrame, owner_col: str) -> pd.DataFrame:
    if segment_df.empty:
        return pd.DataFrame()
    return (
        segment_df.groupby(owner_col, as_index=False)
        .agg(
            cases=("request_id", "size"),
            completed_cases=("is_completed", "sum"),
            avg_hold_days=("total_hold_days", "mean"),
            median_hold_days=("total_hold_days", "median"),
            p50_hold_days=("total_hold_days", lambda s: s.quantile(0.5)),
            p90_hold_days=("total_hold_days", lambda s: s.quantile(0.9)),
            tat_cases=("net_tat_days", lambda s: s.notna().sum()),
            avg_tat_days=("net_tat_days", "mean"),
            p90_tat_days=("net_tat_days", lambda s: s.quantile(0.9)),
        )
        .sort_values("cases", ascending=False)
    )


def render_bucket_for_owner(segment_df: pd.DataFrame, owner_col: str, title: str, key_prefix: str) -> None:
    if segment_df.empty:
        st.info(f"No data for {title}.")
        return

    owners = (
        segment_df.groupby(owner_col, as_index=False)
        .agg(cases=("request_id", "size"))
        .sort_values("cases", ascending=False)[owner_col]
        .tolist()
    )
    if not owners:
        st.info(f"No owners found for {title}.")
        return

    selected_owner = st.selectbox(f"{title} - Select {owner_col}", owners, key=f"{key_prefix}_owner")
    owner_df = segment_df[segment_df[owner_col] == selected_owner]

    hold_counts = (
        owner_df["hold_bucket"].value_counts().reindex(["0 days", "0-1", "1-3", "3-7", "7+"], fill_value=0)
        .rename_axis("bucket")
        .reset_index(name="cases")
    )
    tat_counts = (
        owner_df[owner_df["net_tat_days"].notna()]["tat_bucket"]
        .value_counts()
        .reindex(TAT_BUCKET_ORDER, fill_value=0)
        .rename_axis("bucket")
        .reset_index(name="cases")
    )

    c1, c2 = st.columns(2)
    fig_hold = px.bar(hold_counts, x="bucket", y="cases", color="bucket", title=f"{selected_owner} - Holding Buckets")
    fig_hold.update_layout(showlegend=False, xaxis_title="Hold Bucket", yaxis_title="Cases")
    c1.plotly_chart(fig_hold, use_container_width=True)

    fig_tat = px.bar(
        tat_counts,
        x="bucket",
        y="cases",
        color="bucket",
        color_discrete_map=TAT_BUCKET_COLORS,
        category_orders={"bucket": TAT_BUCKET_ORDER},
        title=f"{selected_owner} - TAT Buckets",
    )
    fig_tat.update_layout(showlegend=False, xaxis_title="TAT Bucket", yaxis_title="Cases")
    c2.plotly_chart(fig_tat, use_container_width=True)


def render_owner_reason_graphs(
    segment_df: pd.DataFrame,
    segment_hold_segments: pd.DataFrame,
    owner_col: str,
    owner_label: str,
    key_prefix: str,
) -> None:
    if segment_df.empty:
        st.info(f"No {owner_label.lower()} reason analysis available.")
        return

    owners = (
        segment_df.groupby(owner_col, as_index=False)
        .agg(cases=("request_id", "size"))
        .sort_values("cases", ascending=False)[owner_col]
        .tolist()
    )
    if not owners:
        st.info(f"No {owner_label.lower()} values found.")
        return

    selected_owner = st.selectbox(
        f"{owner_label} - show most handled reasons",
        owners,
        key=f"{key_prefix}_reason_owner",
    )
    owner_cases = segment_df[segment_df[owner_col] == selected_owner].copy()

    owner_hold = pd.DataFrame()
    if not segment_hold_segments.empty:
        owner_hold = segment_hold_segments[segment_hold_segments[owner_col] == selected_owner].copy()

    if not owner_hold.empty:
        hold_reason_top = (
            owner_hold.groupby("hold_reason_short", as_index=False)
            .agg(hold_events=("hold_days", "size"), total_hold_days=("hold_days", "sum"))
            .sort_values(["hold_events", "total_hold_days"], ascending=[False, False])
            .head(10)
        )
    else:
        hold_reason_top = (
            owner_cases.groupby("on_hold_reason_value", as_index=False)
            .size()
            .rename(columns={"size": "hold_events", "on_hold_reason_value": "hold_reason_short"})
            .sort_values("hold_events", ascending=False)
            .head(10)
        )
        hold_reason_top["total_hold_days"] = np.nan

    writeout_top = (
        owner_cases.groupby("writeout_reason_value", as_index=False)
        .size()
        .rename(columns={"size": "cases"})
        .sort_values("cases", ascending=False)
        .head(10)
    )

    c1, c2 = st.columns(2)
    if not hold_reason_top.empty:
        fig_hold_reason = px.bar(
            hold_reason_top,
            x="hold_reason_short",
            y="hold_events",
            color="total_hold_days" if "total_hold_days" in hold_reason_top.columns else None,
            title=f"{selected_owner}: Most onHoldReasonDescription handled",
        )
        fig_hold_reason.update_layout(xaxis_title="onHoldReasonDescription", yaxis_title="Hold Events")
        c1.plotly_chart(fig_hold_reason, use_container_width=True)
    else:
        c1.info("No onHoldReasonDescription data for this owner.")

    if not writeout_top.empty:
        fig_writeout = px.bar(
            writeout_top,
            x="writeout_reason_value",
            y="cases",
            title=f"{selected_owner}: Most writeOutReasonDescription handled",
        )
        fig_writeout.update_layout(xaxis_title="writeOutReasonDescription", yaxis_title="Case Count")
        c2.plotly_chart(fig_writeout, use_container_width=True)
    else:
        c2.info("No writeOutReasonDescription data for this owner.")


st.title("Leakage Pointers - Auto Insurance Issuance")
st.caption("Focused app for leakage pointers, monthly flow, non-completed diagnostics, and multi-touch/multi-hold deep dives.")

st.sidebar.header("Input")
uploaded = st.sidebar.file_uploader("Upload file (.csv/.xlsx/.xls)", type=["csv", "xlsx", "xls"])
delimiter_map = {"Auto detect": "auto", "Comma": ",", "Tab": "tab", "Pipe": "|", "Semicolon": ";"}
delimiter_label = st.sidebar.selectbox("Delimiter (text files)", list(delimiter_map.keys()))
delimiter = delimiter_map[delimiter_label]

if uploaded is not None:
    raw_df = load_data(uploaded.getvalue(), uploaded.name, delimiter)
    source = uploaded.name
elif DEFAULT_FILE.exists():
    raw_df = pd.read_csv(DEFAULT_FILE, dtype=str, low_memory=False)
    source = str(DEFAULT_FILE)
else:
    st.error("Upload a file to continue.")
    st.stop()

if raw_df.empty:
    st.error("No data found in file.")
    st.stop()

df, hold_segments, metadata = prepare_data(raw_df)

with st.expander("Detected columns and source"):
    st.write({"source": source, **metadata})
    st.dataframe(df.head(10), use_container_width=True)

st.sidebar.header("Filters")
date_range = None
if df["create_month_dt"].notna().any():
    month_starts = sorted(pd.to_datetime(df["create_month_dt"].dropna().unique()))
    month_labels = [m.strftime("%Y-%m") for m in month_starts]
    label_to_month = {lab: m for lab, m in zip(month_labels, month_starts)}
    if len(month_labels) == 1:
        start_month = month_starts[0]
        end_month = start_month + pd.offsets.MonthEnd(1) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        date_range = (start_month, end_month)
        st.sidebar.caption(f"Month range: {month_labels[0]} to {month_labels[0]}")
    else:
        selected = st.sidebar.select_slider("Create month range", options=month_labels, value=(month_labels[0], month_labels[-1]))
        start_month = label_to_month[selected[0]]
        end_month = label_to_month[selected[1]] + pd.offsets.MonthEnd(1) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        date_range = (start_month, end_month)

status_options = sorted(df["status_value"].dropna().unique().tolist())
selected_statuses = st.sidebar.multiselect("StatusDescription", status_options, default=status_options)

case_segment_options = ["StraightThrough", "MultiTouch", "MultiHold"]
selected_case_segments = st.sidebar.multiselect(
    "Case Segment",
    case_segment_options,
    default=case_segment_options,
)

filtered = filter_data(df, date_range, selected_statuses, selected_case_segments)
if filtered.empty:
    st.warning("No rows match the selected filters.")
    st.stop()

if not hold_segments.empty:
    filtered_request_ids = set(filtered["request_id"].astype(str))
    hold_segments_filtered = hold_segments[hold_segments["request_id"].astype(str).isin(filtered_request_ids)].copy()
else:
    hold_segments_filtered = hold_segments.copy()

completed = filtered[filtered["is_completed"]].copy()
open_cases = filtered[~filtered["is_completed"]].copy()
multi_touch_df = filtered[filtered["multi_touch"]].copy()
multi_hold_df = filtered[filtered["multi_hold"]].copy()

# 1) Total/completed + month-wise status count
st.subheader("1. Totals and Month-wise StatusDescription Count")
c1, c2, c3 = st.columns(3)
c1.metric("Total Cases", f"{len(filtered):,}")
c2.metric("Completed Cases", f"{int(filtered['is_completed'].sum()):,}")
c3.metric("Open Cases", f"{int((~filtered['is_completed']).sum()):,}")

status_month_all = (
    filtered.groupby(["create_month", "status_value"], as_index=False)
    .size()
    .rename(columns={"size": "cases"})
    .sort_values("create_month")
)
top_statuses = (
    status_month_all.groupby("status_value", as_index=False)["cases"].sum().sort_values("cases", ascending=False).head(12)["status_value"]
    .tolist()
)
status_month_plot = status_month_all[status_month_all["status_value"].isin(top_statuses)]
fig_status_all = px.bar(
    status_month_plot,
    x="create_month",
    y="cases",
    color="status_value",
    title="Month-wise StatusDescription Count (All Cases)",
)
fig_status_all.update_layout(xaxis_title="Create Month", yaxis_title="Case Count")
st.plotly_chart(fig_status_all, use_container_width=True)

status_month_completed = (
    completed.groupby(["create_month", "status_value"], as_index=False)
    .size()
    .rename(columns={"size": "cases"})
    .sort_values("create_month")
)
if not status_month_completed.empty:
    top_statuses_c = (
        status_month_completed.groupby("status_value", as_index=False)["cases"]
        .sum()
        .sort_values("cases", ascending=False)
        .head(12)["status_value"]
        .tolist()
    )
    fig_status_completed = px.bar(
        status_month_completed[status_month_completed["status_value"].isin(top_statuses_c)],
        x="create_month",
        y="cases",
        color="status_value",
        title="Month-wise StatusDescription Count (Completed Cases)",
    )
    fig_status_completed.update_layout(xaxis_title="Create Month", yaxis_title="Completed Case Count")
    st.plotly_chart(fig_status_completed, use_container_width=True)

# 2) Month-wise completed cases, TAT, average holding time
st.subheader("2. Completed Cases by Month - TAT and Average Holding Time")
if completed.empty:
    st.info("No completed cases in selected filters.")
else:
    monthly_completed = (
        completed.groupby("create_month", as_index=False)
        .agg(
            completed_cases=("is_completed", "size"),
            avg_net_tat_days=("net_tat_days", "mean"),
            p50_net_tat_days=("net_tat_days", lambda s: s.quantile(0.5)),
            p90_net_tat_days=("net_tat_days", lambda s: s.quantile(0.9)),
            avg_hold_days=("total_hold_days", "mean"),
            median_hold_days=("total_hold_days", "median"),
            p50_hold_days=("total_hold_days", lambda s: s.quantile(0.5)),
            p90_hold_days=("total_hold_days", lambda s: s.quantile(0.9)),
        )
        .sort_values("create_month")
    )
    monthly_completed["tat_risk"] = monthly_completed["avg_net_tat_days"].apply(tat_bucket_label)

    c_tat, c_hold = st.columns(2)
    fig_tat_risk = px.bar(
        monthly_completed,
        x="create_month",
        y="avg_net_tat_days",
        color="tat_risk",
        color_discrete_map=TAT_BUCKET_COLORS,
        category_orders={"tat_risk": TAT_BUCKET_ORDER + ["Unknown"]},
        title="Completed Cases: Monthly Avg Net TAT (Risk Colored)",
    )
    fig_tat_risk.add_hline(y=5, line_dash="dot", line_color=TAT_BUCKET_COLORS["5-7 days"])
    fig_tat_risk.add_hline(y=7, line_dash="dot", line_color=TAT_BUCKET_COLORS["7+ days"])
    fig_tat_risk.update_layout(xaxis_title="Create Month", yaxis_title="Avg Net TAT (days)")
    c_tat.plotly_chart(fig_tat_risk, use_container_width=True)

    hold_trend = monthly_completed.melt(
        id_vars=["create_month"],
        value_vars=["avg_hold_days", "p50_hold_days", "p90_hold_days"],
        var_name="metric",
        value_name="days",
    )
    hold_trend["metric"] = hold_trend["metric"].map(
        {
            "avg_hold_days": "Avg Hold",
            "p50_hold_days": "P50 Hold",
            "p90_hold_days": "P90 Hold",
        }
    )
    fig_hold_trend = px.line(
        hold_trend,
        x="create_month",
        y="days",
        color="metric",
        markers=True,
        title="Completed Cases: Monthly Holding Time Trend",
    )
    fig_hold_trend.update_layout(xaxis_title="Create Month", yaxis_title="Days")
    c_hold.plotly_chart(fig_hold_trend, use_container_width=True)

    st.dataframe(monthly_completed, use_container_width=True)

# 3) Non-completed: holding reasons, straight-through and multi-touch
st.subheader("3. Non-Completed Cases - Holding Reasons and StraightThrough vs MultiTouch")
n1, n2, n3 = st.columns(3)
n1.metric("Non-Completed Cases", f"{len(open_cases):,}")
n2.metric("StraightThrough (Open)", f"{int(open_cases['straight_through'].sum()):,}")
n3.metric("MultiTouch (Open)", f"{int(open_cases['multi_touch'].sum()):,}")

if not open_cases.empty:
    open_mix_month = (
        open_cases.assign(case_group=np.where(open_cases["straight_through"], "StraightThrough", "MultiTouch"))
        .groupby(["create_month", "case_group"], as_index=False)
        .size()
        .rename(columns={"size": "cases"})
        .sort_values("create_month")
    )
    fig_open_mix = px.bar(
        open_mix_month,
        x="create_month",
        y="cases",
        color="case_group",
        barmode="group",
        title="Open Cases: StraightThrough vs MultiTouch by Month",
    )
    fig_open_mix.update_layout(xaxis_title="Create Month", yaxis_title="Open Case Count")
    st.plotly_chart(fig_open_mix, use_container_width=True)

    if not hold_segments_filtered.empty:
        open_hold_reason = (
            hold_segments_filtered[~hold_segments_filtered["is_completed"]]
            .groupby("hold_reason_short", as_index=False)
            .agg(total_hold_days=("hold_days", "sum"), hold_events=("hold_days", "size"))
            .sort_values("total_hold_days", ascending=False)
            .head(12)
        )
        if not open_hold_reason.empty:
            fig_open_reason = px.bar(
                open_hold_reason,
                x="hold_reason_short",
                y="total_hold_days",
                color="hold_events",
                title="Open Cases: Top Holding Reasons by Hold Days",
            )
            fig_open_reason.update_layout(xaxis_title="Holding Reason", yaxis_title="Total Hold Days")
            st.plotly_chart(fig_open_reason, use_container_width=True)


def render_segment_analysis(
    segment_df: pd.DataFrame,
    segment_hold_segments_source: pd.DataFrame,
    title: str,
    key_prefix: str,
) -> None:
    st.subheader(title)
    if segment_df.empty:
        st.info("No cases in this segment with current filters.")
        return

    segment_request_ids = set(segment_df["request_id"].astype(str))
    if not segment_hold_segments_source.empty:
        segment_hold_segments = segment_hold_segments_source[
            segment_hold_segments_source["request_id"].astype(str).isin(segment_request_ids)
        ].copy()
    else:
        segment_hold_segments = segment_hold_segments_source.copy()

    r1, r2, r3 = st.columns(3)
    r1.metric("Cases", f"{len(segment_df):,}")
    r2.metric("Completed with TAT", f"{int(segment_df['net_tat_days'].notna().sum()):,}")
    r3.metric("Avg Holding Time", f"{segment_df['total_hold_days'].mean():.2f} days")

    req_top = (
        segment_df.groupby("request_desc_value", as_index=False)
        .size()
        .rename(columns={"size": "cases"})
        .sort_values("cases", ascending=False)
        .head(12)
    )
    fig_req = px.bar(req_top, x="request_desc_value", y="cases", title=f"{title}: Top requestDescription")
    fig_req.update_layout(xaxis_title="requestDescription", yaxis_title="Case Count")
    st.plotly_chart(fig_req, use_container_width=True)

    uw_summary = owner_summary(segment_df, "underwriter_value")
    aa_summary = owner_summary(segment_df, "analyst_value")

    ucol, acol = st.columns(2)
    ucol.markdown("**Top Underwriter Analysis**")
    ucol.dataframe(uw_summary.head(15), use_container_width=True)
    acol.markdown("**Top Analyst Analysis**")
    acol.dataframe(aa_summary.head(15), use_container_width=True)

    st.markdown("**Most handled onHoldReasonDescription and writeOutReasonDescription**")
    render_owner_reason_graphs(
        segment_df,
        segment_hold_segments,
        "underwriter_value",
        "Underwriter",
        f"{key_prefix}_uw",
    )
    render_owner_reason_graphs(
        segment_df,
        segment_hold_segments,
        "analyst_value",
        "Analyst",
        f"{key_prefix}_aa",
    )

    render_bucket_for_owner(segment_df, "underwriter_value", f"{title} - Underwriter Buckets", f"{key_prefix}_uw")
    render_bucket_for_owner(segment_df, "analyst_value", f"{title} - Analyst Buckets", f"{key_prefix}_aa")


# 4) Multi-touch + Multi-hold
render_segment_analysis(multi_touch_df, hold_segments_filtered, "4. MultiTouch Cases", "mt")
render_segment_analysis(multi_hold_df, hold_segments_filtered, "4b. MultiHold Cases", "mh")

# Location and number of vehicle analysis
st.subheader("Location and Number of Vehicle Analysis")
locveh = filtered.dropna(subset=["number_of_locations", "number_of_vehicles"]).copy()
if locveh.empty:
    st.info("Location/Vehicle columns are missing or empty for selected filters.")
else:
    location_order = ["0", "1", "2", "3-5", "6-10", "10+", "Unknown"]
    vehicle_order = ["0", "1", "2", "3-5", "6-10", "11-20", "20+", "Unknown"]

    locveh["location_bucket"] = (
        pd.cut(
            locveh["number_of_locations"],
            bins=[-0.001, 0, 1, 2, 5, 10, np.inf],
            labels=["0", "1", "2", "3-5", "6-10", "10+"],
            include_lowest=True,
        )
        .astype("string")
        .fillna("Unknown")
    )
    locveh["vehicle_bucket"] = (
        pd.cut(
            locveh["number_of_vehicles"],
            bins=[-0.001, 0, 1, 2, 5, 10, 20, np.inf],
            labels=["0", "1", "2", "3-5", "6-10", "11-20", "20+"],
            include_lowest=True,
        )
        .astype("string")
        .fillna("Unknown")
    )

    locveh_completed = locveh[locveh["net_tat_days"].notna()].copy()
    if not locveh_completed.empty:
        heat = (
            locveh_completed.groupby(["location_bucket", "vehicle_bucket"], observed=True, as_index=False)
            .agg(avg_tat_days=("net_tat_days", "mean"), avg_hold_days=("total_hold_days", "mean"), cases=("net_tat_days", "size"))
        )
        heat["location_bucket"] = pd.Categorical(heat["location_bucket"], categories=location_order, ordered=True)
        heat["vehicle_bucket"] = pd.Categorical(heat["vehicle_bucket"], categories=vehicle_order, ordered=True)
        heat = heat.sort_values(["location_bucket", "vehicle_bucket"])

        max_tat_for_scale = float(np.nanmax(heat["avg_tat_days"].values)) if not heat.empty else 7.0
        cmax = max(7.0, max_tat_for_scale)
        pos_5 = min(1.0, 5.0 / cmax)
        pos_7 = min(1.0, 7.0 / cmax)
        tat_scale = [
            (0.0, TAT_BUCKET_COLORS["1-4 days"]),
            (pos_5, TAT_BUCKET_COLORS["1-4 days"]),
            (pos_5, TAT_BUCKET_COLORS["5-7 days"]),
            (pos_7, TAT_BUCKET_COLORS["5-7 days"]),
            (pos_7, TAT_BUCKET_COLORS["7+ days"]),
            (1.0, TAT_BUCKET_COLORS["7+ days"]),
        ]

        fig_heat = px.density_heatmap(
            heat,
            x="location_bucket",
            y="vehicle_bucket",
            z="avg_tat_days",
            title="Avg Net TAT by Location Bucket and Vehicle Bucket",
            color_continuous_scale=tat_scale,
            zmin=0,
            zmax=cmax,
        )
        fig_heat.update_layout(xaxis_title="Number of Locations", yaxis_title="Number of Vehicles")
        st.plotly_chart(fig_heat, use_container_width=True)
        st.dataframe(heat.sort_values("cases", ascending=False), use_container_width=True)
    else:
        st.info("No completed cases with valid TAT for location/vehicle TAT analysis.")

# First hold day from created
st.subheader("First Hold Day from Created Date")
first_hold = filtered[filtered["first_hold_delay_days"].notna()].copy()
if first_hold.empty:
    st.info("No valid first-hold records available.")
else:
    fh_median, fh_p50, fh_p90 = pct_stats(first_hold["first_hold_delay_days"])
    f1, f2, f3 = st.columns(3)
    f1.metric("Median (days)", f"{fh_median:.2f}")
    f2.metric("P50 (days)", f"{fh_p50:.2f}")
    f3.metric("P90 (days)", f"{fh_p90:.2f}")

    fig_first_hold_hist = px.histogram(
        first_hold,
        x="first_hold_delay_days",
        nbins=40,
        title="Distribution: Days from Created to First Hold",
    )
    fig_first_hold_hist.update_layout(xaxis_title="Days to 1st Hold", yaxis_title="Case Count")
    st.plotly_chart(fig_first_hold_hist, use_container_width=True)

    first_hold_month = (
        first_hold.groupby("create_month", as_index=False)
        .agg(
            avg_first_hold_day=("first_hold_delay_days", "mean"),
            p50_first_hold_day=("first_hold_delay_days", lambda s: s.quantile(0.5)),
            p90_first_hold_day=("first_hold_delay_days", lambda s: s.quantile(0.9)),
        )
        .sort_values("create_month")
    )
    first_hold_trend = first_hold_month.melt(
        id_vars=["create_month"],
        value_vars=["avg_first_hold_day", "p50_first_hold_day", "p90_first_hold_day"],
        var_name="metric",
        value_name="days",
    )
    first_hold_trend["metric"] = first_hold_trend["metric"].map(
        {
            "avg_first_hold_day": "Avg",
            "p50_first_hold_day": "P50",
            "p90_first_hold_day": "P90",
        }
    )
    fig_first_hold_trend = px.line(
        first_hold_trend,
        x="create_month",
        y="days",
        color="metric",
        markers=True,
        title="Month-wise: Day of First Hold from Created",
    )
    fig_first_hold_trend.update_layout(xaxis_title="Create Month", yaxis_title="Days")
    st.plotly_chart(fig_first_hold_trend, use_container_width=True)

# Holding time percentiles
st.subheader("Holding Time Percentiles (Median, P50, P90)")
segments = {
    "All Cases": filtered,
    "Completed Cases": completed,
    "Non-Completed Cases": open_cases,
    "MultiTouch Cases": multi_touch_df,
    "MultiHold Cases": multi_hold_df,
}
rows = []
for seg_name, seg_df in segments.items():
    med, p50, p90 = pct_stats(seg_df["total_hold_days"])
    rows.append(
        {
            "segment": seg_name,
            "cases": len(seg_df),
            "median_hold_days": med,
            "p50_hold_days": p50,
            "p90_hold_days": p90,
        }
    )
hold_pct_table = pd.DataFrame(rows)
st.dataframe(hold_pct_table, use_container_width=True)

st.subheader("Prominent Drivers for Longer Hold Time and Longer TAT")

long_hold_threshold = float(filtered["total_hold_days"].quantile(0.9)) if not filtered.empty else np.nan
long_tat_threshold = float(completed["net_tat_days"].quantile(0.9)) if not completed.empty else np.nan

d1, d2 = st.columns(2)
d1.metric("Long Hold Threshold (P90)", f"{long_hold_threshold:.2f} days" if np.isfinite(long_hold_threshold) else "NA")
d2.metric("Long TAT Threshold (P90)", f"{long_tat_threshold:.2f} days" if np.isfinite(long_tat_threshold) else "NA")

driver_base = filtered.copy()
driver_base["is_long_hold"] = False if not np.isfinite(long_hold_threshold) else (driver_base["total_hold_days"] >= long_hold_threshold)
driver_base["is_long_tat"] = False if not np.isfinite(long_tat_threshold) else (
    driver_base["net_tat_days"].notna() & (driver_base["net_tat_days"] >= long_tat_threshold)
)

req_driver = (
    driver_base.groupby("request_desc_value", as_index=False)
    .agg(
        total_cases=("request_id", "size"),
        long_hold_cases=("is_long_hold", "sum"),
        long_tat_cases=("is_long_tat", "sum"),
        avg_hold_days=("total_hold_days", "mean"),
        p90_hold_days=("total_hold_days", lambda s: s.quantile(0.9)),
        avg_tat_days=("net_tat_days", "mean"),
        p90_tat_days=("net_tat_days", lambda s: s.quantile(0.9)),
    )
    .sort_values(["long_tat_cases", "long_hold_cases", "avg_hold_days"], ascending=[False, False, False])
)
st.markdown("**Most Prominent requestTypeDescription**")
st.dataframe(req_driver.head(20), use_container_width=True)

req_plot = req_driver.head(12)
if not req_plot.empty:
    fig_req_driver = px.bar(
        req_plot,
        x="request_desc_value",
        y="long_tat_cases",
        color="long_hold_cases",
        title="requestTypeDescription prominence for longer TAT and hold",
    )
    fig_req_driver.update_layout(xaxis_title="requestTypeDescription", yaxis_title="Long TAT Case Count")
    st.plotly_chart(fig_req_driver, use_container_width=True)

st.markdown("**Most Prominent onHoldReasonDescriptionsHistory**")
if hold_segments_filtered.empty:
    st.info("No hold-segment data available for reason-level prominence.")
else:
    flag_map = driver_base[["request_id", "is_long_hold", "is_long_tat"]].drop_duplicates("request_id")
    reason_driver = hold_segments_filtered.merge(flag_map, on="request_id", how="left")
    reason_driver["is_long_hold"] = reason_driver["is_long_hold"].fillna(False)
    reason_driver["is_long_tat"] = reason_driver["is_long_tat"].fillna(False)

    reason_summary = (
        reason_driver.groupby("hold_reason_short", as_index=False)
        .agg(
            hold_events=("hold_days", "size"),
            total_hold_days=("hold_days", "sum"),
            avg_hold_event_days=("hold_days", "mean"),
            long_hold_events=("is_long_hold", "sum"),
            long_tat_events=("is_long_tat", "sum"),
        )
        .sort_values(["long_tat_events", "long_hold_events", "total_hold_days"], ascending=[False, False, False])
    )
    st.dataframe(reason_summary.head(20), use_container_width=True)

    reason_plot = reason_summary.head(12)
    fig_reason_driver = px.bar(
        reason_plot,
        x="hold_reason_short",
        y="total_hold_days",
        color="long_tat_events",
        title="onHoldReasonDescriptionsHistory prominence for longer hold and TAT",
    )
    fig_reason_driver.update_layout(xaxis_title="onHoldReasonDescription", yaxis_title="Total Hold Days")
    st.plotly_chart(fig_reason_driver, use_container_width=True)

st.subheader("Month-wise: Underwriter and Analyst Taking Most TAT")

owner_min_cases = st.number_input(
    "Minimum completed cases per month for owner ranking",
    min_value=1,
    max_value=50,
    value=3,
    step=1,
)

monthly_owner_base = completed[completed["create_month"].notna() & (completed["create_month"] != "NaT")].copy()
if monthly_owner_base.empty:
    st.info("No completed cases for month-wise owner TAT ranking.")
else:
    uw_month = (
        monthly_owner_base.groupby(["create_month", "underwriter_value"], as_index=False)
        .agg(cases=("net_tat_days", "size"), avg_tat_days=("net_tat_days", "mean"), p90_tat_days=("net_tat_days", lambda s: s.quantile(0.9)))
    )
    uw_month = uw_month[uw_month["cases"] >= owner_min_cases]
    uw_month_top = uw_month.sort_values(["create_month", "avg_tat_days"], ascending=[True, False]).drop_duplicates("create_month")
    uw_month_top["tat_risk"] = uw_month_top["avg_tat_days"].apply(tat_bucket_label)

    aa_month = (
        monthly_owner_base.groupby(["create_month", "analyst_value"], as_index=False)
        .agg(cases=("net_tat_days", "size"), avg_tat_days=("net_tat_days", "mean"), p90_tat_days=("net_tat_days", lambda s: s.quantile(0.9)))
    )
    aa_month = aa_month[aa_month["cases"] >= owner_min_cases]
    aa_month_top = aa_month.sort_values(["create_month", "avg_tat_days"], ascending=[True, False]).drop_duplicates("create_month")
    aa_month_top["tat_risk"] = aa_month_top["avg_tat_days"].apply(tat_bucket_label)

    c_uw, c_aa = st.columns(2)
    c_uw.markdown("**Month-wise Highest TAT Underwriter**")
    c_uw.dataframe(uw_month_top.sort_values("create_month"), use_container_width=True)
    c_aa.markdown("**Month-wise Highest TAT Analyst**")
    c_aa.dataframe(aa_month_top.sort_values("create_month"), use_container_width=True)

    if not uw_month_top.empty:
        fig_uw_month = px.bar(
            uw_month_top.sort_values("create_month"),
            x="create_month",
            y="avg_tat_days",
            color="tat_risk",
            color_discrete_map=TAT_BUCKET_COLORS,
            category_orders={"tat_risk": TAT_BUCKET_ORDER + ["Unknown"]},
            hover_data=["underwriter_value", "cases", "p90_tat_days"],
            title="Monthly Highest Underwriter Avg TAT",
        )
        fig_uw_month.add_hline(y=5, line_dash="dot", line_color=TAT_BUCKET_COLORS["5-7 days"])
        fig_uw_month.add_hline(y=7, line_dash="dot", line_color=TAT_BUCKET_COLORS["7+ days"])
        fig_uw_month.update_layout(xaxis_title="Create Month", yaxis_title="Avg TAT Days")
        st.plotly_chart(fig_uw_month, use_container_width=True)

    if not aa_month_top.empty:
        fig_aa_month = px.bar(
            aa_month_top.sort_values("create_month"),
            x="create_month",
            y="avg_tat_days",
            color="tat_risk",
            color_discrete_map=TAT_BUCKET_COLORS,
            category_orders={"tat_risk": TAT_BUCKET_ORDER + ["Unknown"]},
            hover_data=["analyst_value", "cases", "p90_tat_days"],
            title="Monthly Highest Analyst Avg TAT",
        )
        fig_aa_month.add_hline(y=5, line_dash="dot", line_color=TAT_BUCKET_COLORS["5-7 days"])
        fig_aa_month.add_hline(y=7, line_dash="dot", line_color=TAT_BUCKET_COLORS["7+ days"])
        fig_aa_month.update_layout(xaxis_title="Create Month", yaxis_title="Avg TAT Days")
        st.plotly_chart(fig_aa_month, use_container_width=True)

st.caption(
    "Definitions: Completed = completedDateTime present; Net TAT = completedDateTime - createDateTime - total holding time; "
    "MultiTouch = touches_count > 1; MultiHold = hold_reason_count >= 2."
)
