import io
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(page_title="Auto Issuance Feedback Dashboard", layout="wide")
st.markdown(
    """
    <style>
    div[data-baseweb="tab-list"] {
        position: sticky;
        top: 0;
        z-index: 1000;
        background: white;
        border-bottom: 1px solid rgba(49, 51, 63, 0.2);
        padding-top: 0.2rem;
        padding-bottom: 0.2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DEFAULT_FILE = Path("/Users/rituparnapaldas/Downloads/auto_issuance_synthetic_1year_10000rows.csv")
MISSING_TOKENS = {"", "nan", "none", "null", "na", "n/a", "-", "[]"}
DATE_TOKEN_PATTERN = re.compile(r"\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}")
TAT_BUCKET_ORDER = ["1-4 days", "5-7 days", "7+ days"]
TAT_BUCKET_COLORS = {
    "1-4 days": "#2ca02c",
    "5-7 days": "#FFBF00",
    "7+ days": "#d62728",
}
OPEN_BUCKET_ORDER = ["0-4 days", "5-7 days", "7+ days"]
OPEN_BUCKET_COLORS = {
    "0-4 days": "#2ca02c",
    "5-7 days": "#FFBF00",
    "7+ days": "#d62728",
}
HOLD_BUCKET_ORDER = ["0-4 days", "5-7 days", "7+ days"]
HOLD_BUCKET_COLORS = {
    "0-4 days": "#2ca02c",
    "5-7 days": "#FFBF00",
    "7+ days": "#d62728",
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


def pct_value(numerator: float, denominator: float) -> float:
    if denominator is None or pd.isna(denominator) or float(denominator) <= 0:
        return np.nan
    if numerator is None or pd.isna(numerator):
        return np.nan
    return (float(numerator) / float(denominator)) * 100.0


def pct_text(numerator: float, denominator: float) -> str:
    value = pct_value(numerator, denominator)
    if pd.isna(value):
        return "NA"
    return f"{value:.2f}%"


def calc_hold_metrics(on_hold_text: object, off_hold_text: object, reason_text: object) -> Tuple[float, int, bool]:
    on_values = parse_date_history_cell(on_hold_text)
    off_values = parse_date_history_cell(off_hold_text)
    reason_values = parse_reason_history_cell(reason_text, expected_count=len(on_values) if on_values else None)

    total_hold_days = 0.0
    for idx, on_value in enumerate(on_values):
        on_dt = parse_dt(on_value)
        if pd.isna(on_dt):
            continue

        # Missing off-hold means no valid hold interval to count.
        off_dt = parse_dt(off_values[idx]) if idx < len(off_values) else pd.NaT
        if pd.isna(off_dt):
            continue

        hold_days = (off_dt - on_dt).total_seconds() / 86400
        if hold_days < 0:
            continue
        total_hold_days += hold_days

    hold_reason_count = len(reason_values)
    straight_through = hold_reason_count == 0
    return total_hold_days, hold_reason_count, straight_through


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
def prepare_data(raw_df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    request_col = find_column(df, ["requestId", "request_id"])
    create_col = find_column(df, ["createDateTime", "create_date_time"])
    completed_col = find_column(df, ["completedDateTime", "completed_date_time"])
    status_col = find_column(df, ["statusDescription", "status_description"])
    request_type_col = find_column(df, ["requestTypeDescription", "requestTypeCode", "requestType"])
    bgi_desc_col = find_column(df, ["bgiDescription", "bgi_description"])
    lob_desc_col = find_column(df, ["lineOfBusinessDescription", "line_of_business_description"])
    underwriting_segment_col = find_column(
        df,
        ["underwritingSegmentDescription", "underwriting_segment_description"],
    )
    underwriter_col = find_column(df, ["underwriterName", "underwriter", "underwriter_name"])
    agent_broker_col = find_column(df, ["AgentBrokerName", "agentBrokerName", "AgentBrokerName__2"])
    account_analyst_col = find_column(df, ["accountAnalystName", "accountAnalyst", "account_analyst_name"])
    on_hold_col = find_column(df, ["onHoldDatesHistory"])
    off_hold_col = find_column(df, ["offHoldDatesHistory"])
    hold_reason_col = find_column(df, ["onHoldReasonDescriptionsHistory"])

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

    if status_col is not None:
        df["status_value"] = df[status_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    else:
        df["status_value"] = "Unknown"

    if request_type_col is not None:
        df["request_type_value"] = df[request_type_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    else:
        df["request_type_value"] = "Unknown"
    if bgi_desc_col is not None:
        df["bgi_desc_value"] = df[bgi_desc_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    else:
        df["bgi_desc_value"] = "Unknown"
    if lob_desc_col is not None:
        df["lob_desc_value"] = df[lob_desc_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    else:
        df["lob_desc_value"] = "Unknown"
    if underwriting_segment_col is not None:
        df["underwriting_segment_value"] = (
            df[underwriting_segment_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
        )
    else:
        df["underwriting_segment_value"] = "Unknown"
    if underwriter_col is not None:
        df["underwriter_value"] = df[underwriter_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    else:
        df["underwriter_value"] = "Unknown"
    if agent_broker_col is not None:
        df["agent_broker_value"] = df[agent_broker_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    else:
        df["agent_broker_value"] = "Unknown"
    if account_analyst_col is not None:
        df["account_analyst_value"] = df[account_analyst_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    else:
        df["account_analyst_value"] = "Unknown"

    if create_col is not None:
        df["create_month_dt"] = df["create_dt"].dt.to_period("M").dt.to_timestamp()
        df["create_month"] = df["create_month_dt"].dt.strftime("%Y-%m")
    else:
        df["create_month_dt"] = pd.NaT
        df["create_month"] = "Unknown"

    if all(col is not None for col in [on_hold_col, off_hold_col, hold_reason_col]):
        hold_metrics = df.apply(
            lambda row: calc_hold_metrics(row[on_hold_col], row[off_hold_col], row[hold_reason_col]),
            axis=1,
            result_type="expand",
        )
        hold_metrics.columns = ["total_hold_days", "hold_reason_count", "straight_through"]
        df = pd.concat([df, hold_metrics], axis=1)
    else:
        df["total_hold_days"] = 0.0
        df["hold_reason_count"] = 0
        df["straight_through"] = True

    df["total_hold_days"] = pd.to_numeric(df["total_hold_days"], errors="coerce").fillna(0.0).clip(lower=0)
    df["hold_reason_count"] = pd.to_numeric(df["hold_reason_count"], errors="coerce").fillna(0).astype(int)
    df["straight_through"] = df["straight_through"].fillna(False).astype(bool)

    df["is_completed"] = df["completed_dt"].notna()
    gross_tat = (df["completed_dt"] - df["create_dt"]).dt.total_seconds() / 86400
    gross_tat = gross_tat.where(gross_tat >= 0, np.nan)
    df["gross_tat_days"] = gross_tat

    net_tat = gross_tat - df["total_hold_days"]
    net_tat = net_tat.where(net_tat >= 0, np.nan)
    df["net_tat_days"] = net_tat

    today = pd.Timestamp.today().normalize()
    df["open_days"] = np.nan
    open_mask = ~df["is_completed"]
    df.loc[open_mask, "open_days"] = (today - df.loc[open_mask, "create_dt"]).dt.total_seconds() / 86400
    df.loc[open_mask, "open_days"] = df.loc[open_mask, "open_days"].where(df.loc[open_mask, "open_days"] >= 0, np.nan)

    df["tat_bucket"] = pd.cut(
        df["net_tat_days"],
        bins=[0, 4, 7, np.inf],
        labels=TAT_BUCKET_ORDER,
        include_lowest=True,
    )
    df["open_days_bucket"] = pd.cut(
        df["open_days"],
        bins=[-0.001, 4, 7, np.inf],
        labels=OPEN_BUCKET_ORDER,
        include_lowest=True,
    )
    df["hold_days_bucket"] = pd.cut(
        df["total_hold_days"],
        bins=[-0.001, 4, 7, np.inf],
        labels=HOLD_BUCKET_ORDER,
        include_lowest=True,
    )

    metadata = {
        "request_col": request_col,
        "create_col": create_col,
        "completed_col": completed_col,
        "status_col": status_col,
        "request_type_col": request_type_col,
        "bgi_desc_col": bgi_desc_col,
        "lob_desc_col": lob_desc_col,
        "underwriting_segment_col": underwriting_segment_col,
        "underwriter_col": underwriter_col,
        "agent_broker_col": agent_broker_col,
        "account_analyst_col": account_analyst_col,
        "on_hold_col": on_hold_col,
        "off_hold_col": off_hold_col,
        "hold_reason_col": hold_reason_col,
    }
    return df, metadata


def month_rate(numerator_df: pd.DataFrame, denominator_df: pd.DataFrame, label: str) -> pd.DataFrame:
    den = (
        denominator_df[denominator_df["create_month"].notna() & (denominator_df["create_month"] != "NaT")]
        .groupby("create_month", as_index=False)
        .agg(total_cases=("request_id", "size"))
    )
    num = (
        numerator_df[numerator_df["create_month"].notna() & (numerator_df["create_month"] != "NaT")]
        .groupby("create_month", as_index=False)
        .agg(cases=("request_id", "size"))
    )
    out = den.merge(num, on="create_month", how="left")
    out["cases"] = out["cases"].fillna(0)
    out["pct"] = out.apply(lambda r: pct_value(r["cases"], r["total_cases"]), axis=1)
    out = out.sort_values("create_month")
    out["label"] = label
    return out


def explode_hold_reasons(source_df: pd.DataFrame, hold_reason_col: Optional[str]) -> pd.DataFrame:
    if hold_reason_col is None or hold_reason_col not in source_df.columns or source_df.empty:
        return pd.DataFrame(columns=["request_id", "create_month", "hold_reason_short"])

    rows = []
    work_df = source_df[["request_id", "create_month", hold_reason_col]].copy()
    for _, row in work_df.iterrows():
        reasons = parse_reason_history_cell(row[hold_reason_col], expected_count=None)
        if not reasons:
            continue
        for reason in reasons:
            rows.append(
                {
                    "request_id": row["request_id"],
                    "create_month": row["create_month"],
                    "hold_reason_short": short_reason(reason),
                }
            )

    if not rows:
        return pd.DataFrame(columns=["request_id", "create_month", "hold_reason_short"])
    return pd.DataFrame(rows)


def add_bar_labels(
    fig: object,
    orientation: str = "v",
    value_type: str = "percent",
    use_text_field: bool = False,
    text_as_percent: bool = False,
) -> None:
    if use_text_field:
        text_template = "%{text:.1f}%" if text_as_percent else "%{text:,.0f}"
    elif orientation == "h":
        if value_type == "percent":
            text_template = "%{x:.1f}%"
        elif value_type == "days":
            text_template = "%{x:.2f}"
        else:
            text_template = "%{x:,.0f}"
    else:
        if value_type == "percent":
            text_template = "%{y:.1f}%"
        elif value_type == "days":
            text_template = "%{y:.2f}"
        else:
            text_template = "%{y:,.0f}"

    fig.update_traces(
        texttemplate=text_template,
        textposition="inside",
        insidetextanchor="middle",
        cliponaxis=False,
    )
    fig.update_layout(uniformtext_minsize=8, uniformtext_mode="hide")


def make_bucket_bar(
    counts_df: pd.DataFrame,
    bucket_col: str,
    count_col: str,
    color_map: Dict[str, str],
    title: str,
    category_order: Optional[List[str]] = None,
) -> None:
    if counts_df.empty:
        st.info("No data available.")
        return

    total = counts_df[count_col].sum()
    counts_df = counts_df.copy()
    counts_df["share_pct"] = counts_df[count_col].apply(lambda x: pct_value(x, total))
    fig = px.bar(
        counts_df,
        x=bucket_col,
        y="share_pct",
        text=count_col,
        color=bucket_col,
        color_discrete_map=color_map,
        category_orders={bucket_col: category_order} if category_order else None,
        title=title,
    )
    fig.update_traces(
        customdata=np.stack([counts_df[count_col]], axis=-1),
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Count: %{customdata[0]:,.0f}<br>"
            "Share: %{y:.2f}%<extra></extra>"
        ),
    )
    add_bar_labels(fig, orientation="v", value_type="count", use_text_field=True)
    fig.update_layout(xaxis_title="Bucket", yaxis_title="Share (%)", showlegend=False)
    st.plotly_chart(fig, use_container_width=True)


def make_bucket_month_bar(
    base_df: pd.DataFrame,
    bucket_col: str,
    title: str,
    color_map: Dict[str, str],
    category_order: Optional[List[str]] = None,
) -> None:
    if base_df.empty:
        st.info("No data available for month-wise bucket chart.")
        return

    month_df = base_df[base_df["create_month"].notna() & (base_df["create_month"] != "NaT")].copy()
    if month_df.empty:
        st.info("No valid month values available for month-wise bucket chart.")
        return

    month_df["bucket_value"] = month_df[bucket_col].astype("string").fillna("Unknown")
    month_counts = (
        month_df.groupby(["create_month", "bucket_value"], observed=True, as_index=False)
        .agg(cases=("request_id", "size"))
        .sort_values("create_month")
    )
    month_counts["month_total"] = month_counts.groupby("create_month")["cases"].transform("sum")
    month_counts["share_pct"] = month_counts.apply(lambda r: pct_value(r["cases"], r["month_total"]), axis=1)

    order = category_order[:] if category_order else []
    if "Unknown" in month_counts["bucket_value"].values and "Unknown" not in order:
        order.append("Unknown")

    fig = px.bar(
        month_counts,
        x="create_month",
        y="share_pct",
        text="share_pct",
        color="bucket_value",
        barmode="stack",
        color_discrete_map=color_map,
        category_orders={"bucket_value": order} if order else None,
        title=title,
        hover_data={"cases": ":,.0f", "month_total": ":,.0f", "share_pct": ":.2f"},
    )
    add_bar_labels(fig, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
    fig.update_layout(xaxis_title="Create Month", yaxis_title="Share of Month (%)")
    st.plotly_chart(fig, use_container_width=True)


st.title("Auto Issuance - leakage Dashboard")
st.caption("Fresh version focused on overall completion, open, and straight-through views.")

st.sidebar.header("Input")
uploaded = st.sidebar.file_uploader("Upload file (.csv/.xlsx/.xls)", type=["csv", "xlsx", "xls"])
delimiter_map = {"Auto detect": "auto", "Comma": ",", "Tab": "tab", "Pipe": "|", "Semicolon": ";"}
delimiter_label = st.sidebar.selectbox("Delimiter (text files)", list(delimiter_map.keys()))
delimiter = delimiter_map[delimiter_label]

if uploaded is not None:
    try:
        raw_df = load_data(uploaded.getvalue(), uploaded.name, delimiter)
        source = uploaded.name
    except Exception as exc:
        st.error(f"Could not read uploaded file. Please check file format/delimiter. Details: {exc}")
        st.stop()
elif DEFAULT_FILE.exists():
    try:
        raw_df = pd.read_csv(DEFAULT_FILE, dtype=str, low_memory=False)
        source = str(DEFAULT_FILE)
    except Exception as exc:
        st.warning(
            f"Default file exists but could not be read: {exc}. "
            "Please upload a CSV/XLSX/XLS file from the sidebar."
        )
        st.stop()
else:
    st.info("No input data found. Please upload a CSV/XLSX/XLS file from the sidebar to start analysis.")
    st.stop()

if raw_df.empty:
    st.warning("Loaded file has no rows. Please upload a file with data.")
    st.stop()

df, metadata = prepare_data(raw_df)

with st.expander("Detected columns and source"):
    st.write({"source": source, **metadata})
    st.dataframe(df.head(10), use_container_width=True)

st.sidebar.header("Global Filter")
filtered = df.copy()
if filtered["create_month_dt"].notna().any():
    month_starts = sorted(pd.to_datetime(filtered["create_month_dt"].dropna().unique()))
    month_labels = [m.strftime("%Y-%m") for m in month_starts]
    label_to_month = {lab: m for lab, m in zip(month_labels, month_starts)}
    if len(month_labels) == 1:
        st.sidebar.caption(f"Create month: {month_labels[0]}")
    else:
        selected = st.sidebar.select_slider(
            "Create month range",
            options=month_labels,
            value=(month_labels[0], month_labels[-1]),
        )
        start_month = label_to_month[selected[0]]
        end_month = label_to_month[selected[1]] + pd.offsets.MonthEnd(1) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        filtered = filtered[filtered["create_dt"].between(start_month, end_month, inclusive="both")]

if filtered.empty:
    st.warning("No rows available after filters.")
    st.stop()

total_cases = len(filtered)
completed_df = filtered[filtered["is_completed"]].copy()
open_df = filtered[~filtered["is_completed"]].copy()
straight_df = filtered[filtered["straight_through"]].copy()
completed_straight_df = completed_df[completed_df["straight_through"]].copy()
multi_hold_df = filtered[filtered["hold_reason_count"] >= 1].copy()
multi_hold_completed_df = multi_hold_df[multi_hold_df["is_completed"]].copy()
multi_hold_open_df = multi_hold_df[~multi_hold_df["is_completed"]].copy()

tab_data, tab_cycle, tab_multi, tab_straight, tab_agent, tab_market, tab_reson = st.tabs(
    [
        "Data Explorer",
        "Cycle Time Summary",
        "Multi Hold Cases",
        "Straight Through Cases",
        "Agent Broker Summary",
        "Market Analysis",
        "reson",
    ]
)

with tab_cycle:
    st.subheader("1) Overall Snapshot")
    left, right = st.columns([1.2, 1.0])
    with left:
        a1, a2 = st.columns(2)
        a1.metric("Total Number of Cases", f"{total_cases:,}")
        a2.metric("Completed % Cases", pct_text(len(completed_df), total_cases))
        a3, a4 = st.columns(2)
        a3.metric("StraightThrough % Cases", pct_text(len(straight_df), total_cases))
        a4.metric("Open % Cases", pct_text(len(open_df), total_cases))

    with right:
        mix_df = pd.DataFrame(
            {
                "segment": [
                    "Completed + StraightThrough",
                    "Completed + Non-StraightThrough",
                    "Open + StraightThrough",
                    "Open + Non-StraightThrough",
                ],
                "cases": [
                    int(((filtered["is_completed"]) & (filtered["straight_through"])).sum()),
                    int(((filtered["is_completed"]) & (~filtered["straight_through"])).sum()),
                    int(((~filtered["is_completed"]) & (filtered["straight_through"])).sum()),
                    int(((~filtered["is_completed"]) & (~filtered["straight_through"])).sum()),
                ],
            }
        )
        mix_df = mix_df[mix_df["cases"] > 0]
        fig_mix = px.pie(mix_df, names="segment", values="cases", title="Overall Case Mix")
        fig_mix.update_traces(hovertemplate="<b>%{label}</b><br>Count: %{value:,}<br>Share: %{percent}<extra></extra>")
        st.plotly_chart(fig_mix, use_container_width=True)

    st.markdown("---")
    st.subheader("2) Completed Cases")
    completed_tat = completed_df[completed_df["net_tat_days"].notna()].copy()
    c2_left, c2_right = st.columns(2)
    with c2_left:
        completion_month = month_rate(completed_df, filtered, "completion")
        tat_gt7_month = (
            completed_df[completed_df["create_month"].notna() & (completed_df["create_month"] != "NaT")]
            .groupby("create_month", as_index=False)
            .agg(
                completed_with_valid_tat=("net_tat_days", lambda s: s.notna().sum()),
                tat_over_7_cases=("net_tat_days", lambda s: (pd.to_numeric(s, errors="coerce") > 7).sum()),
            )
        )
        tat_gt7_month["tat_over_7_pct"] = tat_gt7_month.apply(
            lambda r: pct_value(r["tat_over_7_cases"], r["completed_with_valid_tat"]),
            axis=1,
        )
        completion_month = completion_month.merge(
            tat_gt7_month[["create_month", "tat_over_7_pct"]],
            on="create_month",
            how="left",
        )
        completion_plot = completion_month.melt(
            id_vars=["create_month"],
            value_vars=["pct", "tat_over_7_pct"],
            var_name="metric",
            value_name="percent_value",
        )
        completion_plot["metric"] = completion_plot["metric"].map(
            {
                "pct": "Completed %",
                "tat_over_7_pct": "TAT > 7 days % (Completed)",
            }
        )
        fig_completion = px.line(
            completion_plot,
            x="create_month",
            y="percent_value",
            color="metric",
            markers=True,
            title="Completed % by Create Month + % TAT > 7 Days",
        )
        fig_completion.update_layout(xaxis_title="Create Month", yaxis_title="Percent (%)")
        st.plotly_chart(fig_completion, use_container_width=True)
    with c2_right:
        make_bucket_month_bar(
            completed_tat,
            bucket_col="tat_bucket",
            title="Completed Cases - TAT Bucket by Month (%)",
            color_map=TAT_BUCKET_COLORS,
            category_order=TAT_BUCKET_ORDER,
        )

    st.markdown("---")
    st.subheader("3) Open Cases")
    o1, o2 = st.columns(2)
    with o1:
        st.metric("Open % Cases", pct_text(len(open_df), total_cases))
        open_month = month_rate(open_df, filtered, "open")
        fig_open = px.line(
            open_month,
            x="create_month",
            y="pct",
            markers=True,
            hover_data={"cases": ":,.0f", "total_cases": ":,.0f", "pct": ":.2f"},
            title="Open % by Create Month",
        )
        fig_open.update_layout(xaxis_title="Create Month", yaxis_title="Open %")
        st.plotly_chart(fig_open, use_container_width=True)

    with o2:
        st.metric(
            "Average Open Days",
            f"{open_df['open_days'].mean():.2f} days" if not open_df.empty else "NA",
        )

    st.markdown("---")
    st.subheader("4) StraightThrough Cases")
    s1, s2 = st.columns(2)
    straight_tat_source = completed_straight_df[completed_straight_df["net_tat_days"].notna()].copy()
    with s1:
        st.metric("Straight % within Completed", pct_text(len(completed_straight_df), len(completed_df)))
        straight_completed_month = month_rate(completed_straight_df, completed_df, "straight_in_completed")
        fig_straight = px.line(
            straight_completed_month,
            x="create_month",
            y="pct",
            markers=True,
            hover_data={"cases": ":,.0f", "total_cases": ":,.0f", "pct": ":.2f"},
            title="StraightThrough % within Completed by Month",
        )
        fig_straight.update_layout(xaxis_title="Create Month", yaxis_title="StraightThrough % in Completed")
        st.plotly_chart(fig_straight, use_container_width=True)

    with s2:
        st.metric(
            "Average TAT (StraightThrough Completed)",
            f"{completed_straight_df['net_tat_days'].mean():.2f} days" if not completed_straight_df.empty else "NA",
        )
        straight_tat_counts = (
            completed_straight_df["tat_bucket"]
            .astype("string")
            .value_counts()
            .reindex(TAT_BUCKET_ORDER, fill_value=0)
            .rename_axis("bucket")
            .reset_index(name="count")
        )
        make_bucket_bar(
            straight_tat_counts,
            bucket_col="bucket",
            count_col="count",
            color_map=TAT_BUCKET_COLORS,
            title="StraightThrough Completed - TAT Bucket (%)",
            category_order=TAT_BUCKET_ORDER,
        )

    make_bucket_month_bar(
        straight_tat_source,
        bucket_col="tat_bucket",
        title="StraightThrough Completed - TAT Bucket by Month (%)",
        color_map=TAT_BUCKET_COLORS,
        category_order=TAT_BUCKET_ORDER,
    )

with tab_multi:
    st.subheader("Multi Hold Cases")
    st.caption("Definition: Multi Hold = cases where hold_reason_count >= 1")

    if multi_hold_df.empty:
        st.info("No multi-hold cases found in the current filter range.")
    else:
        mh1, mh2, mh3, mh4 = st.columns(4)
        mh1.metric("Multi Hold Cases", f"{len(multi_hold_df):,}")
        mh2.metric("Multi Hold % of Total", pct_text(len(multi_hold_df), total_cases))
        mh3.metric("Completed % in Multi Hold", pct_text(len(multi_hold_completed_df), len(multi_hold_df)))
        mh4.metric("Incomplete % in Multi Hold", pct_text(len(multi_hold_open_df), len(multi_hold_df)))
        multi_hold_work = multi_hold_df.copy()
        multi_hold_work["touches"] = multi_hold_work["hold_reason_count"].fillna(0).astype(float) + 1.0
        multi_hold_reasons_all = explode_hold_reasons(multi_hold_df, metadata.get("hold_reason_col"))

        def top5_from_col(source_df: pd.DataFrame, col_name: str, label_name: str) -> pd.DataFrame:
            out = (
                source_df[col_name]
                .astype("string")
                .fillna("Unknown")
                .replace("", "Unknown")
                .value_counts()
                .head(5)
                .rename_axis(label_name)
                .reset_index(name="cases")
            )
            out["share_pct"] = out["cases"].apply(lambda x: pct_value(x, len(source_df)))
            return out

        def build_handler_summary(source_df: pd.DataFrame, col_name: str, role_label: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
            work = source_df[["request_id", "create_month", col_name]].copy()
            work[col_name] = work[col_name].astype("string").fillna("Unknown").replace("", "Unknown")
            work = work[work[col_name] != "Unknown"]
            if work.empty:
                return pd.DataFrame(), pd.DataFrame()

            overall = (
                work[col_name]
                .value_counts()
                .head(10)
                .rename_axis("handler")
                .reset_index(name="cases")
            )
            overall["role"] = role_label
            overall["share_pct"] = overall["cases"].apply(lambda x: pct_value(x, len(source_df)))

            month_mix = (
                work[work[col_name].isin(overall["handler"])]
                .groupby(["create_month", col_name], as_index=False)
                .agg(cases=("request_id", "size"))
                .rename(columns={col_name: "handler"})
                .sort_values("create_month")
            )
            month_mix["month_total"] = month_mix.groupby("create_month")["cases"].transform("sum")
            month_mix["share_pct"] = month_mix.apply(lambda r: pct_value(r["cases"], r["month_total"]), axis=1)
            return overall, month_mix

        st.markdown("---")
        st.markdown("### 1) Month-wise % Multi Hold and Month-wise TAT Bucket")
        sec1_left, sec1_right = st.columns(2)
        with sec1_left:
            mh_pct_month = month_rate(multi_hold_df, filtered, "multi_hold_rate")
            if mh_pct_month.empty:
                st.info("No valid month values available for multi-hold % trend.")
            else:
                fig_mh_pct = px.line(
                    mh_pct_month,
                    x="create_month",
                    y="pct",
                    markers=True,
                    title="Month-wise % of Cases that are Multi Hold",
                    hover_data={"cases": ":,.0f", "total_cases": ":,.0f", "pct": ":.2f"},
                )
                fig_mh_pct.update_layout(xaxis_title="Create Month", yaxis_title="Multi Hold %")
                st.plotly_chart(fig_mh_pct, use_container_width=True)

        with sec1_right:
            make_bucket_month_bar(
                multi_hold_completed_df[multi_hold_completed_df["net_tat_days"].notna()],
                bucket_col="tat_bucket",
                title="Completed Multi Hold - TAT Bucket by Month (%)",
                color_map=TAT_BUCKET_COLORS,
                category_order=TAT_BUCKET_ORDER,
            )

        st.markdown("---")
        st.markdown("### 2) Touches, Hold Reason Distribution, and Who Handles Multi Hold")
        tsec1, tsec2 = st.columns(2)
        with tsec1:
            touches_month = (
                multi_hold_work[multi_hold_work["create_month"].notna() & (multi_hold_work["create_month"] != "NaT")]
                .groupby("create_month", as_index=False)
                .agg(
                    cases=("request_id", "size"),
                    avg_touches=("touches", "mean"),
                    p90_touches=("touches", lambda s: s.quantile(0.9) if s.notna().any() else np.nan),
                )
                .sort_values("create_month")
            )
            if touches_month.empty:
                st.info("No month-wise touches data available for multi-hold cases.")
            else:
                touches_long = touches_month.melt(
                    id_vars=["create_month", "cases"],
                    value_vars=["avg_touches", "p90_touches"],
                    var_name="metric",
                    value_name="touch_value",
                )
                touches_long["metric"] = touches_long["metric"].map(
                    {"avg_touches": "Average Touches", "p90_touches": "P90 Touches"}
                )
                fig_touches = px.line(
                    touches_long,
                    x="create_month",
                    y="touch_value",
                    color="metric",
                    markers=True,
                    hover_data={"cases": ":,.0f", "touch_value": ":.2f"},
                    title="Multi Hold - Average Number of Touches Over Time",
                )
                fig_touches.update_layout(xaxis_title="Create Month", yaxis_title="Touches")
                st.plotly_chart(fig_touches, use_container_width=True)

        with tsec2:
            if multi_hold_reasons_all.empty:
                st.info("No hold reason history available for multi-hold month-wise distribution.")
            else:
                top_reason_list = multi_hold_reasons_all["hold_reason_short"].value_counts().head(8).index.tolist()
                reason_month_dist = multi_hold_reasons_all.copy()
                reason_month_dist["reason_plot"] = reason_month_dist["hold_reason_short"].where(
                    reason_month_dist["hold_reason_short"].isin(top_reason_list),
                    "Other",
                )
                reason_month_dist = (
                    reason_month_dist.groupby(["create_month", "reason_plot"], as_index=False)
                    .agg(events=("request_id", "size"))
                    .sort_values("create_month")
                )
                reason_month_dist["month_total"] = reason_month_dist.groupby("create_month")["events"].transform("sum")
                reason_month_dist["share_pct"] = reason_month_dist.apply(
                    lambda r: pct_value(r["events"], r["month_total"]),
                    axis=1,
                )
                fig_reason_dist = px.bar(
                    reason_month_dist,
                    x="create_month",
                    y="share_pct",
                    text="share_pct",
                    color="reason_plot",
                    barmode="stack",
                    title="Top Hold Reasons - Month-wise Distribution (Multi Hold)",
                    hover_data={"events": ":,.0f", "month_total": ":,.0f", "share_pct": ":.2f"},
                )
                add_bar_labels(fig_reason_dist, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
                fig_reason_dist.update_layout(xaxis_title="Create Month", yaxis_title="Share of Hold Events (%)", legend_title="Hold Reason")
                st.plotly_chart(fig_reason_dist, use_container_width=True)

        role_map = {
            "Account Analyst": "account_analyst_value",
            "Agent Broker": "agent_broker_value",
            "Underwriter": "underwriter_value",
        }
        role_choice = st.selectbox("Who is handling multi-hold cases? (role view)", list(role_map.keys()), key="mh_handler_role")
        handler_summary, handler_month_mix = build_handler_summary(multi_hold_df, role_map[role_choice], role_choice)
        h1, h2 = st.columns([1.0, 1.4])
        with h1:
            if handler_summary.empty:
                st.info(f"No {role_choice} values available for multi-hold handler summary.")
            else:
                show_handler = handler_summary[["handler", "cases", "share_pct"]].copy()
                st.dataframe(
                    show_handler.style.format({"cases": "{:,.0f}", "share_pct": "{:.2f}%"}),
                    use_container_width=True,
                )
        with h2:
            if handler_month_mix.empty:
                st.info(f"No month-wise {role_choice} distribution available.")
            else:
                fig_handlers = px.bar(
                    handler_month_mix,
                    x="create_month",
                    y="share_pct",
                    text="share_pct",
                    color="handler",
                    barmode="stack",
                    title=f"{role_choice} - Month-wise Share in Multi Hold Cases",
                    hover_data={"cases": ":,.0f", "month_total": ":,.0f", "share_pct": ":.2f"},
                )
                add_bar_labels(fig_handlers, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
                fig_handlers.update_layout(xaxis_title="Create Month", yaxis_title="Share of Month (%)", legend_title=role_choice)
                st.plotly_chart(fig_handlers, use_container_width=True)

        st.markdown("---")
        st.markdown("### 3) Multi Hold Cases with TAT 5-7 or 7+ Days (Top 5 Drivers)")
        long_tat_multi = multi_hold_completed_df[
            multi_hold_completed_df["tat_bucket"].astype("string").isin(["5-7 days", "7+ days"])
        ].copy()
        st.metric("Multi Hold Cases in 5-7 or 7+ TAT", f"{len(long_tat_multi):,}")

        if long_tat_multi.empty:
            st.info("No multi-hold completed cases found in TAT buckets 5-7 or 7+ days.")
        else:
            def plot_monthwise_top5_mix(
                source_df: pd.DataFrame,
                value_col: str,
                value_label: str,
                title: str,
                color_seq: List[str],
            ) -> None:
                mix_df = source_df[source_df["create_month"].notna() & (source_df["create_month"] != "NaT")].copy()
                if mix_df.empty:
                    st.info(f"No month-wise data for {value_label}.")
                    return

                mix_df[value_col] = mix_df[value_col].astype("string").fillna("Unknown").replace("", "Unknown")
                top_vals = mix_df[value_col].value_counts().head(5).index.tolist()
                if not top_vals:
                    st.info(f"No values available for {value_label}.")
                    return

                mix_df["plot_value"] = mix_df[value_col].where(mix_df[value_col].isin(top_vals), "Other")
                month_mix = (
                    mix_df.groupby(["create_month", "plot_value"], as_index=False)
                    .agg(cases=("request_id", "size"))
                    .sort_values("create_month")
                )
                month_mix["month_total"] = month_mix.groupby("create_month")["cases"].transform("sum")
                month_mix["share_pct"] = month_mix.apply(lambda r: pct_value(r["cases"], r["month_total"]), axis=1)

                category_order = top_vals + (["Other"] if "Other" in month_mix["plot_value"].values else [])
                fig = px.bar(
                    month_mix,
                    x="create_month",
                    y="share_pct",
                    text="share_pct",
                    color="plot_value",
                    barmode="stack",
                    category_orders={"plot_value": category_order},
                    title=title,
                    hover_data={"cases": ":,.0f", "month_total": ":,.0f", "share_pct": ":.2f"},
                    color_discrete_sequence=color_seq,
                )
                add_bar_labels(fig, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
                fig.update_layout(xaxis_title="Create Month", yaxis_title="Share of Month (%)", legend_title=value_label)
                st.plotly_chart(fig, use_container_width=True)

            c1, c2 = st.columns(2)
            with c1:
                plot_monthwise_top5_mix(
                    long_tat_multi,
                    value_col="request_type_value",
                    value_label="Request Type",
                    title="Month-wise Distribution: Top 5 Request Type (Multi Hold, TAT 5-7/7+)",
                    color_seq=px.colors.qualitative.Set2,
                )
            with c2:
                hold_reason_long = explode_hold_reasons(long_tat_multi, metadata.get("hold_reason_col"))
                if hold_reason_long.empty:
                    st.info("No hold reason history available for this subset.")
                else:
                    plot_monthwise_top5_mix(
                        hold_reason_long.rename(columns={"hold_reason_short": "hold_reason_value"}),
                        value_col="hold_reason_value",
                        value_label="Hold Reason",
                        title="Month-wise Distribution: Top 5 Hold Reason (Multi Hold, TAT 5-7/7+)",
                        color_seq=px.colors.qualitative.Pastel,
                    )

            c3, c4 = st.columns(2)
            with c3:
                plot_monthwise_top5_mix(
                    long_tat_multi,
                    value_col="bgi_desc_value",
                    value_label="BGI Description",
                    title="Month-wise Distribution: Top 5 BGI Description (Multi Hold, TAT 5-7/7+)",
                    color_seq=px.colors.qualitative.Bold,
                )
            with c4:
                plot_monthwise_top5_mix(
                    long_tat_multi,
                    value_col="lob_desc_value",
                    value_label="Line of Business",
                    title="Month-wise Distribution: Top 5 Line of Business (Multi Hold, TAT 5-7/7+)",
                    color_seq=px.colors.qualitative.Safe,
                )

        st.markdown("---")
        st.markdown("### 4) Month-wise Hold Reason + Completed/Open Buckets")
        sec3_left, sec3_right = st.columns([1.4, 1.0])
        multi_hold_reasons = multi_hold_reasons_all.copy()
        with sec3_left:
            if multi_hold_reasons.empty:
                st.info("No hold reason history available for multi-hold cases.")
            else:
                reason_counts = multi_hold_reasons["hold_reason_short"].value_counts()
                top_reasons = reason_counts.head(8).index.tolist()
                reason_month = multi_hold_reasons.copy()
                reason_month["reason_plot"] = reason_month["hold_reason_short"].where(
                    reason_month["hold_reason_short"].isin(top_reasons),
                    "Other",
                )
                reason_mix = (
                    reason_month.groupby(["create_month", "reason_plot"], as_index=False)
                    .agg(events=("request_id", "size"))
                    .sort_values("create_month")
                )
                reason_mix["month_total_events"] = reason_mix.groupby("create_month")["events"].transform("sum")
                reason_mix["share_pct"] = reason_mix.apply(
                    lambda r: pct_value(r["events"], r["month_total_events"]),
                    axis=1,
                )
                fig_reason = px.bar(
                    reason_mix,
                    x="create_month",
                    y="share_pct",
                    text="share_pct",
                    color="reason_plot",
                    barmode="stack",
                    title="Month-wise Hold Reason Mix (Multi Hold)",
                    hover_data={"events": ":,.0f", "month_total_events": ":,.0f", "share_pct": ":.2f"},
                )
                add_bar_labels(fig_reason, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
                fig_reason.update_layout(xaxis_title="Create Month", yaxis_title="Share of Hold Reason Events (%)", legend_title="Hold Reason")
                st.plotly_chart(fig_reason, use_container_width=True)

        with sec3_right:
            reason_options = ["All"]
            if not multi_hold_reasons.empty:
                reason_options += sorted(multi_hold_reasons["hold_reason_short"].dropna().unique().tolist())
            selected_reason = st.selectbox("Hold Reason focus for buckets", reason_options, key="mh_reason_focus")
            if selected_reason == "All" or multi_hold_reasons.empty:
                mh_completed_reason = multi_hold_completed_df.copy()
                mh_open_reason = multi_hold_open_df.copy()
            else:
                focus_ids = set(
                    multi_hold_reasons.loc[
                        multi_hold_reasons["hold_reason_short"] == selected_reason, "request_id"
                    ].astype(str)
                )
                mh_completed_reason = multi_hold_completed_df[
                    multi_hold_completed_df["request_id"].astype(str).isin(focus_ids)
                ].copy()
                mh_open_reason = multi_hold_open_df[
                    multi_hold_open_df["request_id"].astype(str).isin(focus_ids)
                ].copy()

            make_bucket_month_bar(
                mh_completed_reason[mh_completed_reason["net_tat_days"].notna()],
                bucket_col="tat_bucket",
                title=f"Completed Multi Hold TAT Bucket by Month (%) - {selected_reason}",
                color_map=TAT_BUCKET_COLORS,
                category_order=TAT_BUCKET_ORDER,
            )
            make_bucket_month_bar(
                mh_open_reason[mh_open_reason["total_hold_days"].notna()],
                bucket_col="hold_days_bucket",
                title=f"Incomplete Multi Hold Hold Bucket by Month (%) - {selected_reason}",
                color_map=HOLD_BUCKET_COLORS,
                category_order=HOLD_BUCKET_ORDER,
            )

with tab_agent:
    st.subheader("People Analysis")
    st.caption(
        "Focus cohort defaults to Net TAT 4-7 days and includes only people handling at least 10 cases in that cohort."
    )

    completed_people = completed_df[completed_df["net_tat_days"].notna()].copy()
    if completed_people.empty:
        st.info("No completed cases with valid Net TAT available for people analysis.")
    else:
        tat_min = float(np.floor(completed_people["net_tat_days"].min()))
        tat_max = float(np.ceil(completed_people["net_tat_days"].max()))
        if tat_max < tat_min:
            tat_min, tat_max = 0.0, 7.0

        default_low = max(tat_min, 4.0)
        default_high = min(tat_max, 7.0)
        if default_high < default_low:
            default_low, default_high = tat_min, tat_max

        p1, p2, p3 = st.columns(3)
        with p1:
            if tat_min == tat_max:
                tat_range = (tat_min, tat_max)
                st.caption(f"Net TAT range fixed at {tat_min:.1f} days (single-value dataset).")
            else:
                tat_range = st.slider(
                    "Net TAT focus range (days)",
                    min_value=tat_min,
                    max_value=tat_max,
                    value=(default_low, default_high),
                    step=0.5,
                )
        with p2:
            min_cases_people = st.number_input(
                "Minimum handled cases",
                min_value=1,
                max_value=1000,
                value=10,
                step=1,
            )
        with p3:
            st.metric("Completed Cases (Valid TAT)", f"{len(completed_people):,}")

        people_focus = completed_people[
            completed_people["net_tat_days"].between(tat_range[0], tat_range[1], inclusive="both")
        ].copy()
        st.metric("Cases in selected TAT range", f"{len(people_focus):,}")

        def eligible_names(source_df: pd.DataFrame, col_name: str, min_cases: int) -> List[str]:
            work = source_df.copy()
            work[col_name] = work[col_name].astype("string").fillna("Unknown").replace("", "Unknown")
            counts = work[work[col_name] != "Unknown"][col_name].value_counts()
            return counts[counts >= min_cases].index.tolist()

        def plot_top_col(
            source_df: pd.DataFrame,
            value_col: str,
            y_label: str,
            title: str,
            color_code: str,
            denom_cases: int,
        ) -> None:
            if source_df.empty:
                st.info(f"No data for {y_label}.")
                return
            top_df = (
                source_df[value_col]
                .astype("string")
                .fillna("Unknown")
                .replace("", "Unknown")
                .value_counts()
                .head(5)
                .rename_axis(y_label)
                .reset_index(name="cases")
            )
            if top_df.empty:
                st.info(f"No data for {y_label}.")
                return
            top_df["share_pct"] = top_df["cases"].apply(lambda x: pct_value(x, denom_cases))
            fig = px.bar(
                top_df.sort_values("share_pct", ascending=True),
                x="share_pct",
                y=y_label,
                orientation="h",
                title=title,
                hover_data={"cases": ":,.0f", "share_pct": ":.2f"},
                color_discrete_sequence=[color_code],
            )
            add_bar_labels(fig, orientation="h", value_type="percent")
            fig.update_layout(xaxis_title="Share (%)", yaxis_title=y_label.replace("_", " ").title())
            st.plotly_chart(fig, use_container_width=True)

        def monthly_person_tat(source_df: pd.DataFrame, person_col: str, min_cases: int) -> pd.DataFrame:
            valid = source_df[source_df["create_month"].notna() & (source_df["create_month"] != "NaT")].copy()
            if valid.empty:
                return pd.DataFrame(columns=["create_month", "person", "cases", "avg_tat_days"])
            valid[person_col] = valid[person_col].astype("string").fillna("Unknown").replace("", "Unknown")
            valid = valid[valid[person_col] != "Unknown"]
            if valid.empty:
                return pd.DataFrame(columns=["create_month", "person", "cases", "avg_tat_days"])

            eligible = valid[person_col].value_counts()
            eligible_people = eligible[eligible >= min_cases].index.tolist()
            if not eligible_people:
                return pd.DataFrame(columns=["create_month", "person", "cases", "avg_tat_days"])

            valid = valid[valid[person_col].isin(eligible_people)].copy()
            out = (
                valid.groupby(["create_month", person_col], as_index=False)
                .agg(cases=("request_id", "size"), avg_tat_days=("net_tat_days", "mean"))
                .rename(columns={person_col: "person"})
                .sort_values("create_month")
            )
            return out

        def monthly_high_tat_leaders(
            source_df: pd.DataFrame,
            monthly_df: pd.DataFrame,
            person_col: str,
            person_type: str,
        ) -> pd.DataFrame:
            if monthly_df.empty:
                return pd.DataFrame(
                    columns=[
                        "create_month",
                        "person_type",
                        "person",
                        "cases",
                        "avg_tat_days",
                        "top_request_type",
                        "top_hold_reason",
                    ]
                )

            work = source_df.copy()
            work[person_col] = work[person_col].astype("string").fillna("Unknown").replace("", "Unknown")
            leader_rows = monthly_df.loc[monthly_df.groupby("create_month")["avg_tat_days"].idxmax()].sort_values("create_month")

            records: List[Dict[str, object]] = []
            for _, row in leader_rows.iterrows():
                month = str(row["create_month"])
                person = str(row["person"])
                subset = work[(work["create_month"] == month) & (work[person_col] == person)].copy()
                req_series = subset["request_type_value"].astype("string").fillna("Unknown").replace("", "Unknown")
                top_request_type = req_series.value_counts().idxmax() if not req_series.empty else "Unknown"
                hold_events = explode_hold_reasons(subset, metadata.get("hold_reason_col"))
                top_hold_reason = (
                    hold_events["hold_reason_short"].astype("string").value_counts().idxmax()
                    if not hold_events.empty
                    else "Unspecified"
                )
                records.append(
                    {
                        "create_month": month,
                        "person_type": person_type,
                        "person": person,
                        "cases": int(row["cases"]),
                        "avg_tat_days": float(row["avg_tat_days"]),
                        "top_request_type": top_request_type,
                        "top_hold_reason": top_hold_reason,
                    }
                )
            return pd.DataFrame(records)

        def monthly_people_kpi_table(
            base_df: pd.DataFrame,
            person_col: str,
            eligible_people: List[str],
            selected_person: str,
        ) -> pd.DataFrame:
            work = base_df[base_df["create_month"].notna() & (base_df["create_month"] != "NaT")].copy()
            if work.empty:
                return pd.DataFrame()

            work[person_col] = work[person_col].astype("string").fillna("Unknown").replace("", "Unknown")
            work = work[work[person_col] != "Unknown"]
            if work.empty:
                return pd.DataFrame()

            if selected_person == "All":
                target_people = eligible_people
            else:
                target_people = [selected_person]

            work = work[work[person_col].isin(target_people)].copy()
            if work.empty:
                return pd.DataFrame()

            month_totals = (
                base_df[base_df["create_month"].notna() & (base_df["create_month"] != "NaT")]
                .groupby("create_month", as_index=False)
                .agg(month_total_cases=("request_id", "size"))
            )

            kpi = (
                work.groupby(["create_month", person_col], as_index=False)
                .agg(
                    total_cases=("request_id", "size"),
                    completed_cases=("is_completed", "sum"),
                    valid_tat_cases=("net_tat_days", lambda s: s.notna().sum()),
                    avg_tat_days=("net_tat_days", "mean"),
                    p90_tat_days=("net_tat_days", lambda s: s.quantile(0.9) if s.notna().any() else np.nan),
                )
                .rename(columns={person_col: "person"})
            )

            kpi = kpi.merge(month_totals, on="create_month", how="left")
            kpi["case_share_rate_pct"] = kpi.apply(lambda r: pct_value(r["total_cases"], r["month_total_cases"]), axis=1)
            kpi["completed_rate_pct"] = kpi.apply(lambda r: pct_value(r["completed_cases"], r["total_cases"]), axis=1)
            kpi["tat_coverage_pct"] = kpi.apply(lambda r: pct_value(r["valid_tat_cases"], r["total_cases"]), axis=1)
            return kpi.sort_values(["create_month", "avg_tat_days"], ascending=[True, False])

        def render_tat_kpi_table(kpi_df: pd.DataFrame, title: str) -> None:
            st.markdown(title)
            if kpi_df.empty:
                st.info("No month-wise KPI table data available for current selection.")
                return

            table_cols = [
                "create_month",
                "person",
                "total_cases",
                "completed_cases",
                "case_share_rate_pct",
                "completed_rate_pct",
                "tat_coverage_pct",
                "avg_tat_days",
                "p90_tat_days",
            ]
            table = kpi_df[table_cols].copy()

            def style_tat(val: object) -> str:
                try:
                    if pd.notna(val) and float(val) > 7:
                        return "background-color: #d62728; color: white; font-weight: 600;"
                except Exception:
                    return ""
                return ""

            styled = (
                table.style.format(
                    {
                        "total_cases": "{:,.0f}",
                        "completed_cases": "{:,.0f}",
                        "case_share_rate_pct": "{:.2f}%",
                        "completed_rate_pct": "{:.2f}%",
                        "tat_coverage_pct": "{:.2f}%",
                        "avg_tat_days": "{:.2f}",
                        "p90_tat_days": "{:.2f}",
                    },
                    na_rep="NA",
                )
                .map(style_tat, subset=["avg_tat_days", "p90_tat_days"])
            )
            st.dataframe(styled, use_container_width=True)

        st.markdown("---")
        st.markdown("### 0) Month-wise TAT Scatter and High-TAT Leaders")
        analyst_monthly = monthly_person_tat(people_focus, "account_analyst_value", int(min_cases_people))
        broker_monthly = monthly_person_tat(people_focus, "agent_broker_value", int(min_cases_people))

        sc1, sc2 = st.columns(2)
        with sc1:
            if analyst_monthly.empty:
                st.info("No eligible analyst month-wise points for scatter.")
            else:
                fig_analyst_scatter = px.scatter(
                    analyst_monthly,
                    x="create_month",
                    y="avg_tat_days",
                    size="cases",
                    color="avg_tat_days",
                    color_continuous_scale="YlOrRd",
                    hover_data={"person": True, "cases": ":,.0f", "avg_tat_days": ":.2f"},
                    title="Analyst - Month-wise Avg TAT Scatter",
                )
                fig_analyst_scatter.update_layout(xaxis_title="Create Month", yaxis_title="Avg Net TAT (days)")
                st.plotly_chart(fig_analyst_scatter, use_container_width=True)
        with sc2:
            if broker_monthly.empty:
                st.info("No eligible broker month-wise points for scatter.")
            else:
                fig_broker_scatter = px.scatter(
                    broker_monthly,
                    x="create_month",
                    y="avg_tat_days",
                    size="cases",
                    color="avg_tat_days",
                    color_continuous_scale="YlOrRd",
                    hover_data={"person": True, "cases": ":,.0f", "avg_tat_days": ":.2f"},
                    title="Broker - Month-wise Avg TAT Scatter",
                )
                fig_broker_scatter.update_layout(xaxis_title="Create Month", yaxis_title="Avg Net TAT (days)")
                st.plotly_chart(fig_broker_scatter, use_container_width=True)

        analyst_leaders = monthly_high_tat_leaders(
            people_focus,
            analyst_monthly,
            person_col="account_analyst_value",
            person_type="Account Analyst",
        )
        broker_leaders = monthly_high_tat_leaders(
            people_focus,
            broker_monthly,
            person_col="agent_broker_value",
            person_type="Agent Broker",
        )
        leaders_all = pd.concat([analyst_leaders, broker_leaders], ignore_index=True)

        if leaders_all.empty:
            st.info("No month-wise high-TAT leader list available for analyst/broker with current filters.")
        else:
            fig_leaders = px.bar(
                leaders_all,
                x="create_month",
                y="avg_tat_days",
                color="person_type",
                barmode="group",
                hover_data={
                    "person": True,
                    "cases": ":,.0f",
                    "top_request_type": True,
                    "top_hold_reason": True,
                    "avg_tat_days": ":.2f",
                },
                title="Month-wise Highest Avg TAT (Analyst vs Broker)",
            )
            add_bar_labels(fig_leaders, orientation="v", value_type="days")
            fig_leaders.update_layout(xaxis_title="Create Month", yaxis_title="Highest Avg Net TAT (days)")
            st.plotly_chart(fig_leaders, use_container_width=True)

            display_cols = [
                "create_month",
                "person_type",
                "person",
                "avg_tat_days",
                "cases",
                "top_request_type",
                "top_hold_reason",
            ]
            st.dataframe(
                leaders_all[display_cols].sort_values(["create_month", "person_type"]),
                use_container_width=True,
            )

        st.markdown("---")
        st.markdown("### 1) Account Analyst")
        analysts = eligible_names(people_focus, "account_analyst_value", int(min_cases_people))
        if not analysts:
            st.info("No account analyst meets the minimum case threshold in the selected TAT range.")
        else:
            analyst_choice = st.selectbox(
                "Select Account Analyst (eligible group)",
                ["All"] + analysts,
                key="people_analyst",
            )
            analyst_scope = people_focus[people_focus["account_analyst_value"].isin(analysts)].copy()
            if analyst_choice != "All":
                analyst_scope = analyst_scope[analyst_scope["account_analyst_value"] == analyst_choice].copy()

            make_bucket_month_bar(
                analyst_scope,
                bucket_col="tat_bucket",
                title=f"Month-wise TAT Bucket (%) - Account Analyst: {analyst_choice}",
                color_map=TAT_BUCKET_COLORS,
                category_order=TAT_BUCKET_ORDER,
            )
            analyst_kpi = monthly_people_kpi_table(
                filtered,
                person_col="account_analyst_value",
                eligible_people=analysts,
                selected_person=analyst_choice,
            )
            render_tat_kpi_table(analyst_kpi, "#### Month-wise KPI Table (Account Analyst)")

            st.markdown("#### TAT > 7 days drivers (Account Analyst)")
            analyst_over7 = analyst_scope[analyst_scope["net_tat_days"] > 7].copy()
            a1, a2, a3 = st.columns(3)
            with a1:
                plot_top_col(
                    analyst_over7,
                    value_col="request_type_value",
                    y_label="request_type",
                    title="Top 5 Request Type (>7 days TAT)",
                    color_code="#1f77b4",
                    denom_cases=max(len(analyst_over7), 1),
                )
            with a2:
                hold_reason_analyst = explode_hold_reasons(analyst_over7, metadata.get("hold_reason_col"))
                if hold_reason_analyst.empty:
                    st.info("No hold reason history for analyst cases with TAT > 7 days.")
                else:
                    hold_top_analyst = (
                        hold_reason_analyst.groupby("hold_reason_short", as_index=False)
                        .agg(cases=("request_id", "nunique"))
                        .sort_values("cases", ascending=False)
                        .head(5)
                    )
                    hold_top_analyst["share_pct"] = hold_top_analyst["cases"].apply(
                        lambda x: pct_value(x, max(len(analyst_over7), 1))
                    )
                    fig_hold_analyst = px.bar(
                        hold_top_analyst.sort_values("share_pct", ascending=True),
                        x="share_pct",
                        y="hold_reason_short",
                        orientation="h",
                        title="Top 5 Hold Reason (>7 days TAT)",
                        hover_data={"cases": ":,.0f", "share_pct": ":.2f"},
                        color_discrete_sequence=["#9467bd"],
                    )
                    add_bar_labels(fig_hold_analyst, orientation="h", value_type="percent")
                    fig_hold_analyst.update_layout(xaxis_title="Share (%)", yaxis_title="Hold Reason")
                    st.plotly_chart(fig_hold_analyst, use_container_width=True)
            with a3:
                plot_top_col(
                    analyst_over7,
                    value_col="bgi_desc_value",
                    y_label="bgi_description",
                    title="Top 5 BGI Description (>7 days TAT)",
                    color_code="#2ca02c",
                    denom_cases=max(len(analyst_over7), 1),
                )

        st.markdown("---")
        st.markdown("### 2) Agent Broker")
        brokers = eligible_names(people_focus, "agent_broker_value", int(min_cases_people))
        if not brokers:
            st.info("No agent broker meets the minimum case threshold in the selected TAT range.")
        else:
            broker_choice = st.selectbox(
                "Select Agent Broker (eligible group)",
                ["All"] + brokers,
                key="people_broker",
            )
            broker_scope = people_focus[people_focus["agent_broker_value"].isin(brokers)].copy()
            if broker_choice != "All":
                broker_scope = broker_scope[broker_scope["agent_broker_value"] == broker_choice].copy()

            make_bucket_month_bar(
                broker_scope,
                bucket_col="tat_bucket",
                title=f"Month-wise TAT Bucket (%) - Agent Broker: {broker_choice}",
                color_map=TAT_BUCKET_COLORS,
                category_order=TAT_BUCKET_ORDER,
            )
            broker_kpi = monthly_people_kpi_table(
                filtered,
                person_col="agent_broker_value",
                eligible_people=brokers,
                selected_person=broker_choice,
            )
            render_tat_kpi_table(broker_kpi, "#### Month-wise KPI Table (Agent Broker)")

            st.markdown("#### TAT > 7 days drivers (Agent Broker)")
            broker_over7 = broker_scope[broker_scope["net_tat_days"] > 7].copy()
            b1, b2, b3 = st.columns(3)
            with b1:
                plot_top_col(
                    broker_over7,
                    value_col="request_type_value",
                    y_label="request_type",
                    title="Top 5 Request Type (>7 days TAT)",
                    color_code="#2ca02c",
                    denom_cases=max(len(broker_over7), 1),
                )
            with b2:
                hold_reason_broker = explode_hold_reasons(broker_over7, metadata.get("hold_reason_col"))
                if hold_reason_broker.empty:
                    st.info("No hold reason history for broker cases with TAT > 7 days.")
                else:
                    hold_top_broker = (
                        hold_reason_broker.groupby("hold_reason_short", as_index=False)
                        .agg(cases=("request_id", "nunique"))
                        .sort_values("cases", ascending=False)
                        .head(5)
                    )
                    hold_top_broker["share_pct"] = hold_top_broker["cases"].apply(
                        lambda x: pct_value(x, max(len(broker_over7), 1))
                    )
                    fig_hold_broker = px.bar(
                        hold_top_broker.sort_values("share_pct", ascending=True),
                        x="share_pct",
                        y="hold_reason_short",
                        orientation="h",
                        title="Top 5 Hold Reason (>7 days TAT)",
                        hover_data={"cases": ":,.0f", "share_pct": ":.2f"},
                        color_discrete_sequence=["#ff7f0e"],
                    )
                    add_bar_labels(fig_hold_broker, orientation="h", value_type="percent")
                    fig_hold_broker.update_layout(xaxis_title="Share (%)", yaxis_title="Hold Reason")
                    st.plotly_chart(fig_hold_broker, use_container_width=True)
            with b3:
                plot_top_col(
                    broker_over7,
                    value_col="bgi_desc_value",
                    y_label="bgi_description",
                    title="Top 5 BGI Description (>7 days TAT)",
                    color_code="#1f77b4",
                    denom_cases=max(len(broker_over7), 1),
                )

with tab_straight:
    st.subheader("Straight Through Cases")
    st.caption("TAT Cases with TAT 5-7 or 7+ Days (Top 5 Drivers) over month.")

    straight_long_tat = completed_straight_df[
        completed_straight_df["tat_bucket"].astype("string").isin(["5-7 days", "7+ days"])
    ].copy()

    s1, s2 = st.columns(2)
    with s1:
        st.metric("StraightThrough Completed Cases", f"{len(completed_straight_df):,}")
    with s2:
        st.metric(
            "StraightThrough Cases in TAT 5-7/7+",
            f"{len(straight_long_tat):,}",
        )

    if straight_long_tat.empty:
        st.info("No straight-through completed cases found in TAT buckets 5-7 or 7+ days.")
    else:
        def plot_monthwise_top5_mix_st(
            source_df: pd.DataFrame,
            value_col: str,
            value_label: str,
            title: str,
            color_seq: List[str],
        ) -> None:
            mix_df = source_df[source_df["create_month"].notna() & (source_df["create_month"] != "NaT")].copy()
            if mix_df.empty:
                st.info(f"No month-wise data for {value_label}.")
                return

            mix_df[value_col] = mix_df[value_col].astype("string").fillna("Unknown").replace("", "Unknown")
            top_vals = mix_df[value_col].value_counts().head(5).index.tolist()
            if not top_vals:
                st.info(f"No values available for {value_label}.")
                return

            mix_df["plot_value"] = mix_df[value_col].where(mix_df[value_col].isin(top_vals), "Other")
            month_mix = (
                mix_df.groupby(["create_month", "plot_value"], as_index=False)
                .agg(cases=("request_id", "size"))
                .sort_values("create_month")
            )
            month_mix["month_total"] = month_mix.groupby("create_month")["cases"].transform("sum")
            month_mix["share_pct"] = month_mix.apply(lambda r: pct_value(r["cases"], r["month_total"]), axis=1)

            category_order = top_vals + (["Other"] if "Other" in month_mix["plot_value"].values else [])
            fig = px.bar(
                month_mix,
                x="create_month",
                y="share_pct",
                text="share_pct",
                color="plot_value",
                barmode="stack",
                category_orders={"plot_value": category_order},
                title=title,
                hover_data={"cases": ":,.0f", "month_total": ":,.0f", "share_pct": ":.2f"},
                color_discrete_sequence=color_seq,
            )
            add_bar_labels(fig, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
            fig.update_layout(xaxis_title="Create Month", yaxis_title="Share of Month (%)", legend_title=value_label)
            st.plotly_chart(fig, use_container_width=True)

        st1, st2 = st.columns(2)
        with st1:
            plot_monthwise_top5_mix_st(
                straight_long_tat,
                value_col="request_type_value",
                value_label="Request Type",
                title="Month-wise Distribution: Top 5 Request Type (StraightThrough, TAT 5-7/7+)",
                color_seq=px.colors.qualitative.Set2,
            )
        with st2:
            straight_hold_reason = straight_long_tat[["request_id", "create_month"]].copy()
            straight_hold_reason["hold_reason_value"] = "No Hold (StraightThrough)"
            plot_monthwise_top5_mix_st(
                straight_hold_reason,
                value_col="hold_reason_value",
                value_label="Hold Reason",
                title="Month-wise Distribution: Top Hold Reason (StraightThrough, TAT 5-7/7+)",
                color_seq=px.colors.qualitative.Pastel,
            )

        st3, st4 = st.columns(2)
        with st3:
            plot_monthwise_top5_mix_st(
                straight_long_tat,
                value_col="bgi_desc_value",
                value_label="BGI Description",
                title="Month-wise Distribution: Top 5 BGI Description (StraightThrough, TAT 5-7/7+)",
                color_seq=px.colors.qualitative.Bold,
            )
        with st4:
            plot_monthwise_top5_mix_st(
                straight_long_tat,
                value_col="lob_desc_value",
                value_label="Line of Business",
                title="Month-wise Distribution: Top 5 Line of Business (StraightThrough, TAT 5-7/7+)",
                color_seq=px.colors.qualitative.Safe,
            )

with tab_market:
    st.subheader("Market Analysis")
    st.caption("Month-wise TAT bucket and % of overall by BGI Description, with Request Type filter.")

    market_base = completed_df[completed_df["net_tat_days"].notna()].copy()
    market_base["request_type_value"] = market_base["request_type_value"].astype("string").fillna("Unknown").replace("", "Unknown")
    market_base["bgi_desc_value"] = market_base["bgi_desc_value"].astype("string").fillna("Unknown").replace("", "Unknown")

    request_type_options = ["All"] + sorted(market_base["request_type_value"].unique().tolist())
    selected_request_type = st.selectbox(
        "Request Type filter (requestTypeDescription)",
        request_type_options,
        key="market_request_type_filter",
    )
    if selected_request_type != "All":
        market_base = market_base[market_base["request_type_value"] == selected_request_type].copy()

    if market_base.empty:
        st.info("No completed cases with valid TAT for the selected request type.")
    else:
        m1, m2 = st.columns(2)
        with m1:
            bgi_options = ["All"] + sorted(market_base["bgi_desc_value"].unique().tolist())
            selected_bgi = st.selectbox("BGI Description focus for month-wise TAT bucket", bgi_options, key="market_bgi_focus")
            if selected_bgi == "All":
                tat_scope = market_base.copy()
            else:
                tat_scope = market_base[market_base["bgi_desc_value"] == selected_bgi].copy()

            make_bucket_month_bar(
                tat_scope,
                bucket_col="tat_bucket",
                title=f"Month-wise TAT Bucket (%) - BGI: {selected_bgi}",
                color_map=TAT_BUCKET_COLORS,
                category_order=TAT_BUCKET_ORDER,
            )

        with m2:
            bgi_month = market_base[
                market_base["create_month"].notna() & (market_base["create_month"] != "NaT")
            ].copy()
            if bgi_month.empty:
                st.info("No valid month values available for BGI % overall chart.")
            else:
                top_bgi = bgi_month["bgi_desc_value"].value_counts().head(8).index.tolist()
                bgi_month["bgi_plot"] = bgi_month["bgi_desc_value"].where(
                    bgi_month["bgi_desc_value"].isin(top_bgi),
                    "Other",
                )
                bgi_month_mix = (
                    bgi_month.groupby(["create_month", "bgi_plot"], as_index=False)
                    .agg(cases=("request_id", "size"))
                    .sort_values("create_month")
                )
                overall_total = len(market_base)
                bgi_month_mix["pct_overall"] = bgi_month_mix["cases"].apply(lambda x: pct_value(x, overall_total))
                fig_bgi_overall = px.bar(
                    bgi_month_mix,
                    x="create_month",
                    y="pct_overall",
                    text="pct_overall",
                    color="bgi_plot",
                    barmode="stack",
                    title="BGI Description - Month-wise % of Overall Cases",
                    hover_data={"cases": ":,.0f", "pct_overall": ":.2f"},
                )
                add_bar_labels(fig_bgi_overall, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
                fig_bgi_overall.update_layout(xaxis_title="Create Month", yaxis_title="% of Overall Cases", legend_title="BGI Description")
                st.plotly_chart(fig_bgi_overall, use_container_width=True)

        st.markdown("### Average TAT by BGI Description")
        bgi_avg = (
            market_base.groupby("bgi_desc_value", as_index=False)
            .agg(
                cases=("request_id", "size"),
                avg_tat_days=("net_tat_days", "mean"),
                p90_tat_days=("net_tat_days", lambda s: s.quantile(0.9) if s.notna().any() else np.nan),
            )
            .sort_values(["avg_tat_days", "cases"], ascending=[False, False])
        )
        bgi_avg["share_pct"] = bgi_avg["cases"].apply(lambda x: pct_value(x, len(market_base)))

        top_bgi_avg = bgi_avg.head(15).sort_values("avg_tat_days", ascending=True)
        fig_bgi_avg = px.bar(
            top_bgi_avg,
            x="avg_tat_days",
            y="bgi_desc_value",
            orientation="h",
            color="avg_tat_days",
            color_continuous_scale="YlOrRd",
            title="Average Net TAT by BGI Description (Top 15 by Avg TAT)",
            hover_data={"cases": ":,.0f", "share_pct": ":.2f", "p90_tat_days": ":.2f", "avg_tat_days": ":.2f"},
        )
        add_bar_labels(fig_bgi_avg, orientation="h", value_type="days")
        fig_bgi_avg.update_layout(xaxis_title="Average Net TAT (days)", yaxis_title="BGI Description")
        st.plotly_chart(fig_bgi_avg, use_container_width=True)

        st.dataframe(
            bgi_avg.style.format(
                {
                    "cases": "{:,.0f}",
                    "share_pct": "{:.2f}%",
                    "avg_tat_days": "{:.2f}",
                    "p90_tat_days": "{:.2f}",
                },
                na_rep="NA",
            ),
            use_container_width=True,
        )

with tab_reson:
    st.subheader("reson")
    st.caption("High TAT view (Net TAT > 4 days): average TAT and percent share by requested dimensions.")

    reson_base = completed_df[completed_df["net_tat_days"] > 4].copy()
    reson_base["request_type_value"] = reson_base["request_type_value"].astype("string").fillna("Unknown").replace("", "Unknown")
    reson_base["bgi_desc_value"] = reson_base["bgi_desc_value"].astype("string").fillna("Unknown").replace("", "Unknown")
    reson_base["underwriting_segment_value"] = (
        reson_base["underwriting_segment_value"].astype("string").fillna("Unknown").replace("", "Unknown")
    )
    reson_base["agent_broker_value"] = reson_base["agent_broker_value"].astype("string").fillna("Unknown").replace("", "Unknown")

    if reson_base.empty:
        st.info("No completed cases with Net TAT > 4 days available for reson analysis.")
    else:
        def build_reson_summary(source_df: pd.DataFrame, dim_col: str) -> pd.DataFrame:
            out = (
                source_df.groupby(dim_col, as_index=False)
                .agg(
                    cases=("request_id", "size"),
                    avg_tat_days=("net_tat_days", "mean"),
                )
                .rename(columns={dim_col: "data_point"})
                .sort_values(["avg_tat_days", "cases"], ascending=[False, False])
            )
            out["share_pct"] = out["cases"].apply(lambda x: pct_value(x, len(source_df)))
            return out

        def draw_reson_chart(title: str, dim_col: str, y_label: str) -> None:
            st.markdown(title)
            summary = build_reson_summary(reson_base, dim_col)
            if summary.empty:
                st.info("No data points available.")
                return

            c1, c2 = st.columns([1.15, 1.0])
            with c1:
                fig = px.bar(
                    summary.sort_values("avg_tat_days", ascending=True),
                    x="avg_tat_days",
                    y="data_point",
                    orientation="h",
                    text="share_pct",
                    color="share_pct",
                    color_continuous_scale="YlOrRd",
                    title=f"{y_label}: Avg TAT and % share",
                    hover_data={"cases": ":,.0f", "share_pct": ":.2f", "avg_tat_days": ":.2f"},
                )
                fig.update_traces(texttemplate="%{text:.1f}%", textposition="inside")
                fig.update_layout(xaxis_title="Average Net TAT (days)", yaxis_title=y_label)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                st.dataframe(
                    summary[["data_point", "cases", "share_pct", "avg_tat_days"]].style.format(
                        {"cases": "{:,.0f}", "share_pct": "{:.2f}%", "avg_tat_days": "{:.2f}"},
                        na_rep="NA",
                    ),
                    use_container_width=True,
                )

        draw_reson_chart(
            "### 1) Based on requestTypeDescription and average_TAT",
            dim_col="request_type_value",
            y_label="Request Type",
        )
        draw_reson_chart(
            "### 2) Based on bgiDescription and average_TAT",
            dim_col="bgi_desc_value",
            y_label="BGI Description",
        )
        draw_reson_chart(
            "### 3) Based on underwritingSegmentDescription and average_TAT",
            dim_col="underwriting_segment_value",
            y_label="Underwriting Segment",
        )
        draw_reson_chart(
            "### 4) Based on agent broker and average_TAT",
            dim_col="agent_broker_value",
            y_label="Agent Broker",
        )

with tab_data:
    st.subheader("Data Explorer")
    st.write(f"Rows in current filtered view: **{len(filtered):,}**")
    st.markdown("### Variables Used and Their Ranges")

    variable_order = [
        ("request_id", "categorical"),
        ("create_dt", "datetime"),
        ("completed_dt", "datetime"),
        ("create_month", "categorical"),
        ("status_value", "categorical"),
        ("request_type_value", "categorical"),
        ("bgi_desc_value", "categorical"),
        ("lob_desc_value", "categorical"),
        ("underwriting_segment_value", "categorical"),
        ("underwriter_value", "categorical"),
        ("account_analyst_value", "categorical"),
        ("agent_broker_value", "categorical"),
        ("is_completed", "boolean"),
        ("straight_through", "boolean"),
        ("hold_reason_count", "numeric"),
        ("total_hold_days", "numeric"),
        ("gross_tat_days", "numeric"),
        ("net_tat_days", "numeric"),
        ("open_days", "numeric"),
        ("tat_bucket", "categorical"),
        ("open_days_bucket", "categorical"),
        ("hold_days_bucket", "categorical"),
    ]

    profile_rows: List[Dict[str, object]] = []
    for col_name, col_type in variable_order:
        if col_name not in filtered.columns:
            continue

        col_series = filtered[col_name]
        missing_pct = float(col_series.isna().mean() * 100.0)
        range_text = "NA"
        distinct_values = np.nan

        if col_type == "numeric":
            s_num = pd.to_numeric(col_series, errors="coerce")
            if s_num.notna().any():
                min_v = float(s_num.min())
                max_v = float(s_num.max())
                med_v = float(s_num.median())
                range_text = f"{min_v:.2f} to {max_v:.2f} (median {med_v:.2f})"
            distinct_values = int(s_num.nunique(dropna=True))
        elif col_type == "datetime":
            s_dt = pd.to_datetime(col_series, errors="coerce")
            if s_dt.notna().any():
                min_dt = s_dt.min()
                max_dt = s_dt.max()
                range_text = f"{min_dt.strftime('%Y-%m-%d %H:%M')} to {max_dt.strftime('%Y-%m-%d %H:%M')}"
            distinct_values = int(s_dt.nunique(dropna=True))
        elif col_type == "boolean":
            s_bool = col_series.astype("string").fillna("Unknown").str.lower()
            true_count = int((s_bool == "true").sum())
            false_count = int((s_bool == "false").sum())
            unknown_count = int(((s_bool != "true") & (s_bool != "false")).sum())
            range_text = f"True: {true_count:,}, False: {false_count:,}, Unknown: {unknown_count:,}"
            distinct_values = int(s_bool.nunique(dropna=True))
        else:
            s_cat = col_series.astype("string").fillna("Unknown").replace("", "Unknown")
            distinct_values = int(s_cat.nunique(dropna=True))
            top_vals = s_cat.value_counts().head(3)
            if not top_vals.empty:
                range_text = ", ".join([f"{idx} ({int(val):,})" for idx, val in top_vals.items()])
            else:
                range_text = "No values"

        profile_rows.append(
            {
                "variable": col_name,
                "type": col_type,
                "range_or_top_values": range_text,
                "distinct_values": distinct_values,
                "missing_pct": missing_pct,
            }
        )

    if profile_rows:
        profile_df = pd.DataFrame(profile_rows)
        st.dataframe(
            profile_df.style.format({"distinct_values": "{:,.0f}", "missing_pct": "{:.2f}%"}),
            use_container_width=True,
        )
    else:
        st.info("No variable profile available in current filtered view.")

    st.markdown("### Sample Rows")
    show_cols = [
        "request_id",
        "create_dt",
        "completed_dt",
        "status_value",
        "request_type_value",
        "bgi_desc_value",
        "lob_desc_value",
        "underwriting_segment_value",
        "underwriter_value",
        "account_analyst_value",
        "agent_broker_value",
        "is_completed",
        "straight_through",
        "hold_reason_count",
        "total_hold_days",
        "gross_tat_days",
        "net_tat_days",
        "open_days",
        "tat_bucket",
        "open_days_bucket",
        "hold_days_bucket",
        "create_month",
    ]
    available_cols = [c for c in show_cols if c in filtered.columns]
    st.dataframe(filtered[available_cols].head(500), use_container_width=True)

st.caption(
    "Definitions: Completed = completedDateTime present; StraightThrough = onHoldReasonDescriptionsHistory empty; "
    "Net TAT = completedDateTime - createDateTime - valid holding time."
)
