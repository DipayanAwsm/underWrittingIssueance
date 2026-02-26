import io
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(page_title="Auto Issuance Feedback Dashboard", layout="wide")
st.markdown(
    """
    <style>
    div[data-testid="stTabs"] div[data-baseweb="tab-list"],
    div[data-testid="stTabs"] [role="tablist"] {
        position: sticky !important;
        position: -webkit-sticky !important;
        top: 0 !important;
        z-index: 1002 !important;
        background: rgba(255, 255, 255, 0.98) !important;
        border-bottom: 1px solid rgba(49, 51, 63, 0.2) !important;
        padding-top: 0.2rem !important;
        padding-bottom: 0.2rem !important;
        backdrop-filter: blur(2px);
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
SHOW_CALC_DETAILS = True


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
    write_out_reason_col = find_column(
        df,
        [
            "writeOutReasonDescriptionsHistory",
            "writeOutReasonDescription",
            "writeOutReasonDescriptions",
            "writeOutDescriptions",
        ],
    )
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
    if write_out_reason_col is not None:
        df["write_out_reason_value"] = (
            df[write_out_reason_col].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
        )
    else:
        df["write_out_reason_value"] = "Unknown"

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
        "write_out_reason_col": write_out_reason_col,
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


_plotly_chart_counter = 0
_calc_note_counter = 0
_table_note_counter = 0


def get_chart_calc_lines(title: str) -> List[str]:
    t = (title or "").strip().lower()
    lines: List[str] = [
        "Base data: current global filters are applied first (create month range + request type filter).",
        "Percent values use: (numerator / denominator) * 100.",
    ]

    if "tat" in t:
        lines += [
            "Completed case definition: completedDateTime is present/valid.",
            "Gross TAT (days) = (completedDateTime - createDateTime) in days.",
            "Net TAT (days) = Gross TAT - total_hold_days.",
        ]
    if "tat bucket" in t:
        lines += [
            "TAT bucket rules: 1-4 days, 5-7 days, 7+ days.",
            "Month-wise bucket chart: bucket share in a month = bucket cases / total cases in that month.",
        ]
    if "p50" in t or "median" in t:
        lines.append("P50/Median = 50th percentile (`quantile(0.5)`).")
    if "p90" in t:
        lines.append("P90 = 90th percentile (`quantile(0.9)`), used as outlier-excluded high-end benchmark.")
    if "open cases - straight through vs multi hold" in t:
        lines += [
            "Straight Through = hold_reason_count == 0.",
            "Multi Hold = hold_reason_count >= 1.",
            "Slice share = case_type_count / total_open_cases.",
            "Median TAT in slice label is from completed cases of that type with valid net_tat_days.",
        ]
    if "touch" in t:
        lines.append("Touches = hold_reason_count + 1.")
    if "month-wise" in t:
        lines.append("Month-wise grouping key: create_month derived from createDateTime.")
    if "top 5" in t or "top" in t:
        lines.append("Top-N lists are ranked by highest count in the currently filtered dataset.")
    if "hold reason" in t:
        lines.append("Hold reasons are exploded from onHoldReasonDescriptionsHistory; each parsed reason is counted.")
    if "reason description" in t:
        lines.append("Reason Description uses writeOutReasonDescriptionsHistory (fallback: requestTypeDescription when missing).")

    # De-duplicate while preserving order
    deduped: List[str] = []
    seen = set()
    for line in lines:
        if line not in seen:
            deduped.append(line)
            seen.add(line)
    return deduped


def render_calc_note(lines: List[str], label_prefix: str) -> None:
    global _calc_note_counter
    if not SHOW_CALC_DETAILS:
        return
    _calc_note_counter += 1
    st.caption(f"{label_prefix} ({_calc_note_counter})")
    st.markdown("\n".join([f"- {line}" for line in lines]))


def _format_label_value(val: object) -> str:
    try:
        if val is None or pd.isna(val):
            return ""
    except Exception:
        pass
    try:
        fval = float(val)
        if np.isfinite(fval):
            if abs(fval - round(fval)) < 1e-9:
                return f"{fval:,.0f}"
            return f"{fval:,.2f}"
    except Exception:
        return str(val)
    return str(val)


def apply_auto_data_labels(fig: object) -> None:
    try:
        for trace in fig.data:
            t = str(getattr(trace, "type", "")).lower()

            if t == "bar":
                texttemplate = getattr(trace, "texttemplate", None)
                text = getattr(trace, "text", None)
                if texttemplate:
                    continue
                has_text = text is not None and len(text) > 0
                if has_text:
                    continue
                orientation = str(getattr(trace, "orientation", "v")).lower()
                values = list(getattr(trace, "x", [])) if orientation == "h" else list(getattr(trace, "y", []))
                labels = [_format_label_value(v) for v in values]
                trace.update(text=labels, textposition="inside", cliponaxis=False)
                continue

            if t in {"scatter", "scattergl"}:
                x_vals = list(getattr(trace, "x", []))
                y_vals = list(getattr(trace, "y", []))
                if not x_vals and not y_vals:
                    continue
                if len(x_vals) == 1 and len(y_vals) == 1 and x_vals[0] is None and y_vals[0] is None:
                    # Skip dummy scatter traces used only for legend text.
                    continue
                mode = str(getattr(trace, "mode", "") or "")
                if "text" not in mode:
                    mode = f"{mode}+text" if mode else "markers+text"
                text = getattr(trace, "text", None)
                if text is None or len(text) == 0:
                    values = y_vals if y_vals else x_vals
                    text = [_format_label_value(v) for v in values]
                trace.update(mode=mode, text=text, textposition="top center")
                continue

            if t == "pie":
                texttemplate = getattr(trace, "texttemplate", None)
                textinfo = str(getattr(trace, "textinfo", "") or "")
                if not texttemplate and (not textinfo or textinfo == "none"):
                    trace.update(textinfo="label+percent+value")
                continue

            if t == "box":
                trace.update(boxmean=True)
                continue
    except Exception:
        # Keep dashboard rendering even if a trace type does not support label updates.
        return


def render_plotly_chart(fig: object, **kwargs) -> None:
    global _plotly_chart_counter
    if "key" not in kwargs or kwargs["key"] is None:
        _plotly_chart_counter += 1
        kwargs["key"] = f"plotly_{_plotly_chart_counter}"
    apply_auto_data_labels(fig)
    st.plotly_chart(fig, **kwargs)
    title_text = ""
    try:
        title_text = str(getattr(getattr(fig, "layout", None), "title", None).text or "")
    except Exception:
        title_text = ""
    render_calc_note(get_chart_calc_lines(title_text), "How this chart is calculated")


def render_dataframe(data: object, **kwargs) -> None:
    global _table_note_counter
    st.dataframe(data, **kwargs)
    if SHOW_CALC_DETAILS:
        _table_note_counter += 1
        render_calc_note(
            [
                "This table is computed from the currently filtered dataset.",
                "Count columns use grouped record count (`size`) or distinct request count (`nunique`) based on table logic.",
                "Share % columns use: part / relevant total * 100.",
                "Average columns use arithmetic mean over non-null values.",
                "P50/P90 columns use quantiles (`quantile(0.5)` / `quantile(0.9)`).",
            ],
            "How this table is calculated",
        )


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
    render_plotly_chart(fig, use_container_width=True)


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
    render_plotly_chart(fig, use_container_width=True)


st.title("Auto Issuance Speed to Market – Business Intelligence Prescriptive")

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

st.sidebar.header("Global Filter")
SHOW_CALC_DETAILS = st.sidebar.checkbox(
    "Show calculation details under every chart/table",
    value=True,
    key="show_calc_details_global",
)
filtered = df.copy()
if filtered["create_month_dt"].notna().any():
    month_starts = sorted(pd.to_datetime(filtered["create_month_dt"].dropna().unique()))
    month_labels = [m.strftime("%Y-%m") for m in month_starts]
    st.sidebar.markdown("**Create Month Filter**")
    all_months = st.sidebar.checkbox("All months", value=True, key="global_month_all")
    selected_months: List[str] = []
    for idx, month_label in enumerate(month_labels):
        checked = st.sidebar.checkbox(
            month_label,
            value=all_months,
            key=f"global_month_option_{idx}_{normalize_name(month_label)}",
        )
        if checked:
            selected_months.append(month_label)

    if not all_months:
        if selected_months:
            filtered = filtered[filtered["create_month"].astype(str).isin(selected_months)]
        else:
            filtered = filtered.iloc[0:0]

st.sidebar.markdown("---")
st.sidebar.markdown("**Request Type Filter (requestTypeDescription)**")
request_type_values = sorted(
    filtered["request_type_value"].astype("string").fillna("Unknown").replace("", "Unknown").unique().tolist()
)
issuance_like_options = [
    v
    for v in request_type_values
    if any(token in str(v).lower() for token in ["issuance", "issueance", "reissuance", "re-issuance"])
]
request_type_options = issuance_like_options if issuance_like_options else request_type_values
if issuance_like_options:
    st.sidebar.caption("Showing all Issuance/Issueance variations.")
else:
    st.sidebar.caption("No Issuance variation found. Showing all request types.")

all_request_type = st.sidebar.checkbox("All", value=True, key="global_request_type_all")
selected_request_types: List[str] = []
for idx, option in enumerate(request_type_options):
    checked = st.sidebar.checkbox(
        str(option),
        value=all_request_type,
        key=f"global_request_type_option_{idx}_{normalize_name(str(option))}",
    )
    if checked:
        selected_request_types.append(str(option))

if not all_request_type:
    if selected_request_types:
        filtered = filtered[filtered["request_type_value"].astype(str).isin(selected_request_types)]
    else:
        filtered = filtered.iloc[0:0]

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
        "Reason",
    ]
)

with tab_cycle:
    st.subheader("1) Overall Snapshot")
    a1, a2, a3, a4, a5 = st.columns(5)
    a1.metric("Total Number of Cases", f"{total_cases:,}")
    a2.metric("Completed % Cases", pct_text(len(completed_df), total_cases))
    a3.metric("StraightThrough % Cases", pct_text(len(straight_df), total_cases))
    a4.metric("Multi Hold % Cases", pct_text(len(multi_hold_df), total_cases))
    a5.metric("Open % Cases", pct_text(len(open_df), total_cases))

    overall_tat = completed_df[completed_df["net_tat_days"].notna()].copy()
    if overall_tat.empty:
        st.info("No completed cases with valid Net TAT available for overall TAT box plot.")
    else:
        tat_series = overall_tat["net_tat_days"].dropna()
        tat_stats = {
            "Min": float(tat_series.min()),
            "Q1 (P25)": float(tat_series.quantile(0.25)),
            "Median (P50)": float(tat_series.quantile(0.5)),
            "Q3 (P75)": float(tat_series.quantile(0.75)),
            "P90": float(tat_series.quantile(0.9)),
            "Mean": float(tat_series.mean()),
            "Max": float(tat_series.max()),
        }
        fig_overall_tat_box = px.box(
            overall_tat,
            x="net_tat_days",
            points="outliers",
            title="Overall Snapshot - Net TAT Distribution (Box Plot)",
            color_discrete_sequence=["#1f77b4"],
        )
        fig_overall_tat_box.update_traces(name="Net TAT", showlegend=True)
        for label, value in tat_stats.items():
            fig_overall_tat_box.add_trace(
                go.Scatter(
                    x=[None],
                    y=[None],
                    mode="markers",
                    marker=dict(size=0, opacity=0),
                    showlegend=True,
                    name=f"{label}: {value:.2f} days",
                    hoverinfo="skip",
                )
            )
        fig_overall_tat_box.update_layout(
            xaxis_title="Net TAT (days)",
            yaxis_title="",
            legend_title="TAT Stats",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="left",
                x=0,
            ),
            showlegend=True,
        )
        render_plotly_chart(fig_overall_tat_box, use_container_width=True)

    st.markdown("---")
    st.subheader("2) Completed Cases")
    completed_tat = completed_df[completed_df["net_tat_days"].notna()].copy()
    if not completed_tat.empty:
        avg_all = completed_tat["net_tat_days"].mean()
        p90_all = completed_tat["net_tat_days"].quantile(0.9)
    c2_left, c2_right = st.columns(2)
    with c2_left:
        if completed_tat.empty:
            st.info("No completed cases with valid TAT for bucket chart.")
        else:
            st.metric("Average Days of Issuance", f"{avg_all:.2f} days")
            make_bucket_month_bar(
                completed_tat,
                bucket_col="tat_bucket",
                color_map=TAT_BUCKET_COLORS,
                title="Completed Cases - TAT Bucket by Month (%)",
                category_order=TAT_BUCKET_ORDER,
            )
    with c2_right:
        if completed_tat.empty:
            st.info("No completed cases with valid issuance days for percentile trend.")
        else:
            st.metric("Days of Issuance Excluding Outliers (P90)", f"{p90_all:.2f} days")

            p_month = (
                completed_tat[completed_tat["create_month"].notna() & (completed_tat["create_month"] != "NaT")]
                .groupby("create_month", as_index=False)
                .agg(
                    p50_days=("net_tat_days", lambda s: s.quantile(0.5)),
                    p90_days=("net_tat_days", lambda s: s.quantile(0.9)),
                    cases=("request_id", "size"),
                )
                .sort_values("create_month")
            )
            p_month_long = p_month.melt(
                id_vars=["create_month", "cases"],
                value_vars=["p50_days", "p90_days"],
                var_name="metric",
                value_name="days",
            )
            p_month_long["metric"] = p_month_long["metric"].map(
                {"p50_days": "Median Days (P50)", "p90_days": "P90 Days"}
            )
            fig_p = px.line(
                p_month_long,
                x="create_month",
                y="days",
                color="metric",
                markers=True,
                title="Completed Cases - Month-wise P50 and P90",
                hover_data={"cases": ":,.0f", "days": ":.2f"},
            )
            fig_p.update_layout(xaxis_title="Create Month", yaxis_title="Issuance Days")
            render_plotly_chart(fig_p, use_container_width=True)

    st.markdown("---")
    st.subheader("3) StraightThrough Cases")
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
        render_plotly_chart(fig_straight, use_container_width=True)

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
                render_plotly_chart(fig_mh_pct, use_container_width=True)

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
                render_plotly_chart(fig_touches, use_container_width=True)

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
                render_plotly_chart(fig_reason_dist, use_container_width=True)

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
                render_dataframe(
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
                render_plotly_chart(fig_handlers, use_container_width=True)

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
                render_plotly_chart(fig, use_container_width=True)

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
                render_plotly_chart(fig_reason, use_container_width=True)

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
    st.caption("Analysis is based on all completed cases with valid Net TAT.")

    completed_people = completed_df[completed_df["net_tat_days"].notna()].copy()
    if completed_people.empty:
        st.info("No completed cases with valid Net TAT available for people analysis.")
    else:
        min_cases_people = 10
        people_focus = completed_people.copy()

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
            render_plotly_chart(fig, use_container_width=True)

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
            render_dataframe(styled, use_container_width=True)

        st.markdown("---")
        st.markdown("### 0) Average TAT by Broker and Analyst (Top 10)")
        st.caption("Uses completed cases with valid Net TAT and a minimum handled-case threshold of 10.")

        def top_people_high_tat(source_df: pd.DataFrame, person_col: str, min_cases: int, top_n: int) -> pd.DataFrame:
            work = source_df.copy()
            work[person_col] = work[person_col].astype("string").fillna("Unknown").replace("", "Unknown")
            work = work[work[person_col] != "Unknown"]
            if work.empty:
                return pd.DataFrame(columns=["person", "cases", "avg_tat_days", "p90_tat_days"])

            out = (
                work.groupby(person_col, as_index=False)
                .agg(
                    cases=("request_id", "size"),
                    avg_tat_days=("net_tat_days", "mean"),
                    p90_tat_days=("net_tat_days", lambda s: s.quantile(0.9) if s.notna().any() else np.nan),
                )
                .rename(columns={person_col: "person"})
            )
            out = out[out["cases"] >= min_cases].copy()
            if out.empty:
                return pd.DataFrame(columns=["person", "cases", "avg_tat_days", "p90_tat_days"])
            return out.sort_values(["avg_tat_days", "p90_tat_days", "cases"], ascending=[False, False, False]).head(top_n)

        def top_reason_sets_for_people(source_df: pd.DataFrame, person_col: str, top_people: List[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
            if not top_people:
                return pd.DataFrame(), pd.DataFrame()
            subset = source_df[source_df[person_col].astype("string").isin(top_people)].copy()
            if subset.empty:
                return pd.DataFrame(), pd.DataFrame()

            hold_events = explode_hold_reasons(subset, metadata.get("hold_reason_col"))
            if hold_events.empty:
                hold_top = pd.DataFrame(columns=["reason", "cases", "share_pct"])
            else:
                hold_top = (
                    hold_events.groupby("hold_reason_short", as_index=False)
                    .agg(cases=("request_id", "nunique"))
                    .rename(columns={"hold_reason_short": "reason"})
                    .sort_values("cases", ascending=False)
                    .head(5)
                )
                hold_top["share_pct"] = hold_top["cases"].apply(lambda x: pct_value(x, max(len(subset), 1)))

            reason_series = subset["write_out_reason_value"].astype("string").fillna("Unknown").replace("", "Unknown")
            if reason_series.replace("Unknown", pd.NA).dropna().empty:
                reason_series = subset["request_type_value"].astype("string").fillna("Unknown").replace("", "Unknown")
            reason_top = (
                reason_series.value_counts()
                .head(5)
                .rename_axis("reason")
                .reset_index(name="cases")
            )
            reason_top["share_pct"] = reason_top["cases"].apply(lambda x: pct_value(x, max(len(subset), 1)))
            return hold_top, reason_top

        top_brokers = top_people_high_tat(people_focus, "agent_broker_value", int(min_cases_people), top_n=10)
        top_analysts = top_people_high_tat(people_focus, "account_analyst_value", int(min_cases_people), top_n=10)

        t1, t2 = st.columns(2)
        with t1:
            if top_brokers.empty:
                st.info("No broker meets the minimum handled case filter for Top 10 average TAT.")
            else:
                fig_top_broker = px.bar(
                    top_brokers.sort_values("avg_tat_days", ascending=True),
                    x="avg_tat_days",
                    y="person",
                    orientation="h",
                    text="avg_tat_days",
                    title="Average TAT by Broker (Top 10)",
                    hover_data={"cases": ":,.0f", "avg_tat_days": ":.2f", "p90_tat_days": ":.2f"},
                    color_discrete_sequence=["#d62728"],
                )
                add_bar_labels(fig_top_broker, orientation="h", value_type="days", use_text_field=True)
                fig_top_broker.update_layout(xaxis_title="Average Net TAT (days)", yaxis_title="Broker")
                render_plotly_chart(fig_top_broker, use_container_width=True)
        with t2:
            if top_analysts.empty:
                st.info("No analyst meets the minimum handled case filter for Top 10 average TAT.")
            else:
                fig_top_analyst = px.bar(
                    top_analysts.sort_values("avg_tat_days", ascending=True),
                    x="avg_tat_days",
                    y="person",
                    orientation="h",
                    text="avg_tat_days",
                    title="Average TAT by Analyst (Top 10)",
                    hover_data={"cases": ":,.0f", "avg_tat_days": ":.2f", "p90_tat_days": ":.2f"},
                    color_discrete_sequence=["#ff7f0e"],
                )
                add_bar_labels(fig_top_analyst, orientation="h", value_type="days", use_text_field=True)
                fig_top_analyst.update_layout(xaxis_title="Average Net TAT (days)", yaxis_title="Analyst")
                render_plotly_chart(fig_top_analyst, use_container_width=True)

        r1, r2 = st.columns(2)
        with r1:
            hold_top_broker, reason_top_broker = top_reason_sets_for_people(
                people_focus,
                "agent_broker_value",
                top_brokers["person"].astype(str).tolist() if not top_brokers.empty else [],
            )
            if hold_top_broker.empty:
                st.info("No hold reason data for Top 5 high-TAT brokers.")
            else:
                fig_hold_broker_top5 = px.bar(
                    hold_top_broker.sort_values("share_pct", ascending=True),
                    x="share_pct",
                    y="reason",
                    orientation="h",
                    title="Top Hold Reasons for Top 5 High-TAT Brokers",
                    hover_data={"cases": ":,.0f", "share_pct": ":.2f"},
                    color_discrete_sequence=["#9467bd"],
                )
                add_bar_labels(fig_hold_broker_top5, orientation="h", value_type="percent")
                fig_hold_broker_top5.update_layout(xaxis_title="Share (%)", yaxis_title="Hold Reason")
                render_plotly_chart(fig_hold_broker_top5, use_container_width=True)
            if reason_top_broker.empty:
                st.info("No reason description data for Top 5 high-TAT brokers.")
            else:
                fig_reason_broker_top5 = px.bar(
                    reason_top_broker.sort_values("share_pct", ascending=True),
                    x="share_pct",
                    y="reason",
                    orientation="h",
                    title="Top Reason Descriptions for Top 5 High-TAT Brokers",
                    hover_data={"cases": ":,.0f", "share_pct": ":.2f"},
                    color_discrete_sequence=["#2ca02c"],
                )
                add_bar_labels(fig_reason_broker_top5, orientation="h", value_type="percent")
                fig_reason_broker_top5.update_layout(xaxis_title="Share (%)", yaxis_title="Reason Description")
                render_plotly_chart(fig_reason_broker_top5, use_container_width=True)

        with r2:
            hold_top_analyst, reason_top_analyst = top_reason_sets_for_people(
                people_focus,
                "account_analyst_value",
                top_analysts["person"].astype(str).tolist() if not top_analysts.empty else [],
            )
            if hold_top_analyst.empty:
                st.info("No hold reason data for Top 5 high-TAT analysts.")
            else:
                fig_hold_analyst_top5 = px.bar(
                    hold_top_analyst.sort_values("share_pct", ascending=True),
                    x="share_pct",
                    y="reason",
                    orientation="h",
                    title="Top Hold Reasons for Top 5 High-TAT Analysts",
                    hover_data={"cases": ":,.0f", "share_pct": ":.2f"},
                    color_discrete_sequence=["#17becf"],
                )
                add_bar_labels(fig_hold_analyst_top5, orientation="h", value_type="percent")
                fig_hold_analyst_top5.update_layout(xaxis_title="Share (%)", yaxis_title="Hold Reason")
                render_plotly_chart(fig_hold_analyst_top5, use_container_width=True)
            if reason_top_analyst.empty:
                st.info("No reason description data for Top 5 high-TAT analysts.")
            else:
                fig_reason_analyst_top5 = px.bar(
                    reason_top_analyst.sort_values("share_pct", ascending=True),
                    x="share_pct",
                    y="reason",
                    orientation="h",
                    title="Top Reason Descriptions for Top 5 High-TAT Analysts",
                    hover_data={"cases": ":,.0f", "share_pct": ":.2f"},
                    color_discrete_sequence=["#1f77b4"],
                )
                add_bar_labels(fig_reason_analyst_top5, orientation="h", value_type="percent")
                fig_reason_analyst_top5.update_layout(xaxis_title="Share (%)", yaxis_title="Reason Description")
                render_plotly_chart(fig_reason_analyst_top5, use_container_width=True)

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

            st.markdown("#### TAT > 7 days drivers (Account Analyst - Overall Cases)")
            analyst_scope_overall = completed_people[completed_people["account_analyst_value"].isin(analysts)].copy()
            if analyst_choice != "All":
                analyst_scope_overall = analyst_scope_overall[
                    analyst_scope_overall["account_analyst_value"] == analyst_choice
                ].copy()
            analyst_over7 = analyst_scope_overall[analyst_scope_overall["net_tat_days"] > 7].copy()
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
                    render_plotly_chart(fig_hold_analyst, use_container_width=True)
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

            st.markdown("#### TAT > 7 days drivers (Agent Broker - Overall Cases)")
            broker_scope_overall = completed_people[completed_people["agent_broker_value"].isin(brokers)].copy()
            if broker_choice != "All":
                broker_scope_overall = broker_scope_overall[
                    broker_scope_overall["agent_broker_value"] == broker_choice
                ].copy()
            broker_over7 = broker_scope_overall[broker_scope_overall["net_tat_days"] > 7].copy()
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
                    render_plotly_chart(fig_hold_broker, use_container_width=True)
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
    straight_completed_tat = completed_straight_df[completed_straight_df["net_tat_days"].notna()].copy()
    straight_tat_counts = (
        straight_completed_tat["tat_bucket"]
        .astype("string")
        .value_counts()
        .reindex(TAT_BUCKET_ORDER, fill_value=0)
        .rename_axis("bucket")
        .reset_index(name="count")
    )
    straight_tat_counts["share_pct"] = straight_tat_counts["count"].apply(
        lambda x: pct_value(x, len(straight_completed_tat))
    )

    st.markdown("### 1) StraightThrough Counts and TAT Bucket Counts")
    s1, s2, s3, s4, s5 = st.columns(5)
    s1.metric("StraightThrough Cases (All)", f"{len(straight_df):,}")
    s2.metric("StraightThrough Completed Cases", f"{len(completed_straight_df):,}")
    s3.metric("1-4 Days", f"{int(straight_tat_counts.loc[straight_tat_counts['bucket'] == '1-4 days', 'count'].iloc[0]):,}")
    s4.metric("5-7 Days", f"{int(straight_tat_counts.loc[straight_tat_counts['bucket'] == '5-7 days', 'count'].iloc[0]):,}")
    s5.metric("7+ Days", f"{int(straight_tat_counts.loc[straight_tat_counts['bucket'] == '7+ days', 'count'].iloc[0]):,}")

    b1, b2 = st.columns(2)
    with b1:
        make_bucket_bar(
            straight_tat_counts,
            bucket_col="bucket",
            count_col="count",
            color_map=TAT_BUCKET_COLORS,
            title="StraightThrough Completed - TAT Bucket (%)",
            category_order=TAT_BUCKET_ORDER,
        )
    with b2:
        bucket_table = straight_tat_counts.copy()
        bucket_table.columns = ["TAT Bucket", "Count", "Share %"]
        render_dataframe(
            bucket_table.style.format({"Count": "{:,.0f}", "Share %": "{:.2f}%"}),
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("### 2) Bucket-wise StraightThrough Cases Over Months")
    make_bucket_month_bar(
        straight_completed_tat,
        bucket_col="tat_bucket",
        title="StraightThrough Completed - TAT Bucket by Month (%)",
        color_map=TAT_BUCKET_COLORS,
        category_order=TAT_BUCKET_ORDER,
    )

    st.markdown("---")
    st.markdown("### 3) StraightThrough Cases with TAT 5-7 or 7+ Days (Top 5 Drivers)")

    straight_long_tat = straight_completed_tat[
        straight_completed_tat["tat_bucket"].astype("string").isin(["5-7 days", "7+ days"])
    ].copy()
    st.metric("StraightThrough Cases in TAT 5-7/7+", f"{len(straight_long_tat):,}")

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
            render_plotly_chart(fig, use_container_width=True)

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
    st.caption("Month-wise TAT bucket and % of overall by BGI Description.")

    market_base = completed_df[completed_df["net_tat_days"].notna()].copy()
    market_base["request_type_value"] = market_base["request_type_value"].astype("string").fillna("Unknown").replace("", "Unknown")
    market_base["bgi_desc_value"] = market_base["bgi_desc_value"].astype("string").fillna("Unknown").replace("", "Unknown")

    if market_base.empty:
        st.info("No completed cases with valid TAT for market analysis.")
    else:
        st.markdown("### Average Net TAT by BGI Description (Top 15 by Avg TAT)")
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
        render_plotly_chart(fig_bgi_avg, use_container_width=True)

        render_dataframe(
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

        st.markdown("---")
        st.markdown("### BGI and TAT Bucket Views")
        m1, m2, m3 = st.columns(3)
        with m1:
            bgi_bucket_scope = market_base.copy()
            bgi_bucket_scope["tat_bucket_str"] = bgi_bucket_scope["tat_bucket"].astype("string")
            bgi_bucket_scope = bgi_bucket_scope[bgi_bucket_scope["tat_bucket_str"].isin(TAT_BUCKET_ORDER)]
            if bgi_bucket_scope.empty:
                st.info("No valid TAT bucket values available for BGI by TAT bucket graph.")
            else:
                top_bgi_bucket = bgi_bucket_scope["bgi_desc_value"].value_counts().head(10).index.tolist()
                bgi_bucket_scope["bgi_plot"] = bgi_bucket_scope["bgi_desc_value"].where(
                    bgi_bucket_scope["bgi_desc_value"].isin(top_bgi_bucket),
                    "Other",
                )
                bgi_bucket_mix = (
                    bgi_bucket_scope.groupby(["bgi_plot", "tat_bucket_str"], as_index=False)
                    .agg(cases=("request_id", "size"))
                )
                bgi_bucket_mix["bgi_total"] = bgi_bucket_mix.groupby("bgi_plot")["cases"].transform("sum")
                bgi_bucket_mix["share_pct"] = bgi_bucket_mix.apply(
                    lambda r: pct_value(r["cases"], r["bgi_total"]),
                    axis=1,
                )
                bgi_order = top_bgi_bucket + (["Other"] if "Other" in bgi_bucket_mix["bgi_plot"].unique() else [])
                fig_bgi_bucket = px.bar(
                    bgi_bucket_mix,
                    x="bgi_plot",
                    y="share_pct",
                    color="tat_bucket_str",
                    text="share_pct",
                    barmode="stack",
                    category_orders={"tat_bucket_str": TAT_BUCKET_ORDER, "bgi_plot": bgi_order},
                    color_discrete_map=TAT_BUCKET_COLORS,
                    title="BGI by TAT Bucket (%)",
                    hover_data={"cases": ":,.0f", "bgi_total": ":,.0f", "share_pct": ":.2f"},
                )
                add_bar_labels(fig_bgi_bucket, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
                fig_bgi_bucket.update_layout(
                    xaxis_title="BGI Description",
                    yaxis_title="Share within BGI (%)",
                    legend_title="TAT Bucket",
                )
                render_plotly_chart(fig_bgi_bucket, use_container_width=True)

        with m2:
            make_bucket_month_bar(
                market_base,
                bucket_col="tat_bucket",
                title="Month-wise TAT Bucket (%) - All BGI",
                color_map=TAT_BUCKET_COLORS,
                category_order=TAT_BUCKET_ORDER,
            )

        with m3:
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
                render_plotly_chart(fig_bgi_overall, use_container_width=True)

        st.markdown("---")
        st.markdown("### Month-wise BGI Taking Most Time (Flagged for 5-7 and 5+ TAT)")
        bgi_month_perf = market_base[
            market_base["create_month"].notna() & (market_base["create_month"] != "NaT")
        ].copy()
        if bgi_month_perf.empty:
            st.info("No valid month values available for BGI time-taking analysis.")
        else:
            bgi_month_perf["tat_bucket_str"] = bgi_month_perf["tat_bucket"].astype("string")
            bgi_month_summary = (
                bgi_month_perf.groupby(["create_month", "bgi_desc_value"], as_index=False)
                .agg(
                    cases=("request_id", "size"),
                    avg_tat_days=("net_tat_days", "mean"),
                    p90_tat_days=("net_tat_days", lambda s: s.quantile(0.9) if s.notna().any() else np.nan),
                    tat_5_7_cases=("tat_bucket_str", lambda s: (s == "5-7 days").sum()),
                    tat_7_plus_cases=("tat_bucket_str", lambda s: (s == "7+ days").sum()),
                )
                .sort_values("create_month")
            )
            bgi_month_summary["tat_5_plus_cases"] = bgi_month_summary["tat_5_7_cases"] + bgi_month_summary["tat_7_plus_cases"]
            bgi_month_summary["tat_5_7_pct"] = bgi_month_summary.apply(
                lambda r: pct_value(r["tat_5_7_cases"], r["cases"]), axis=1
            )
            bgi_month_summary["tat_7_plus_pct"] = bgi_month_summary.apply(
                lambda r: pct_value(r["tat_7_plus_cases"], r["cases"]), axis=1
            )
            bgi_month_summary["tat_5_plus_pct"] = bgi_month_summary.apply(
                lambda r: pct_value(r["tat_5_plus_cases"], r["cases"]), axis=1
            )
            bgi_month_summary["flag"] = np.where(
                bgi_month_summary["tat_7_plus_pct"] > 0,
                "Red Flag (7+ present)",
                np.where(
                    bgi_month_summary["tat_5_7_pct"] > 0,
                    "Amber Flag (5-7 present)",
                    "No Flag",
                ),
            )

            min_cases_market = st.slider(
                "Minimum cases per Month-BGI for flag view",
                min_value=1,
                max_value=100,
                value=10,
                step=1,
                key="market_min_cases",
            )
            bgi_month_eligible = bgi_month_summary[bgi_month_summary["cases"] >= min_cases_market].copy()

            if bgi_month_eligible.empty:
                st.info("No Month-BGI groups meet the minimum case threshold.")
            else:
                bgi_top_month = (
                    bgi_month_eligible.sort_values(
                        ["create_month", "avg_tat_days", "tat_5_plus_pct", "cases"],
                        ascending=[True, False, False, False],
                    )
                    .groupby("create_month", as_index=False)
                    .head(1)
                    .sort_values("create_month")
                )
                bgi_top_month["bgi_label"] = bgi_top_month["bgi_desc_value"].astype(str).str.slice(0, 28)

                mm1, mm2 = st.columns(2)
                with mm1:
                    fig_bgi_most_time = px.bar(
                        bgi_top_month,
                        x="create_month",
                        y="avg_tat_days",
                        color="flag",
                        text="bgi_label",
                        title="Month-wise BGI Taking Most Time (Highest Avg TAT)",
                        color_discrete_map={
                            "Red Flag (7+ present)": "#d62728",
                            "Amber Flag (5-7 present)": "#FFBF00",
                            "No Flag": "#2ca02c",
                        },
                        hover_data={
                            "bgi_desc_value": True,
                            "cases": ":,.0f",
                            "avg_tat_days": ":.2f",
                            "p90_tat_days": ":.2f",
                            "tat_5_7_pct": ":.2f",
                            "tat_7_plus_pct": ":.2f",
                            "tat_5_plus_pct": ":.2f",
                        },
                    )
                    fig_bgi_most_time.update_traces(textposition="inside")
                    fig_bgi_most_time.update_layout(
                        xaxis_title="Create Month",
                        yaxis_title="Average Net TAT (days)",
                        legend_title="Flag",
                    )
                    render_plotly_chart(fig_bgi_most_time, use_container_width=True)

                with mm2:
                    top_flag_long = bgi_top_month.melt(
                        id_vars=["create_month", "bgi_desc_value", "cases"],
                        value_vars=["tat_5_7_pct", "tat_7_plus_pct"],
                        var_name="tat_flag_bucket",
                        value_name="pct",
                    )
                    top_flag_long["tat_flag_bucket"] = top_flag_long["tat_flag_bucket"].map(
                        {"tat_5_7_pct": "TAT 5-7 days", "tat_7_plus_pct": "TAT 7+ days"}
                    )
                    fig_flag_bucket = px.bar(
                        top_flag_long,
                        x="create_month",
                        y="pct",
                        text="pct",
                        color="tat_flag_bucket",
                        barmode="stack",
                        title="Flagged Bucket Share for Top BGI Each Month",
                        color_discrete_map={"TAT 5-7 days": "#FFBF00", "TAT 7+ days": "#d62728"},
                        hover_data={"bgi_desc_value": True, "cases": ":,.0f", "pct": ":.2f"},
                    )
                    add_bar_labels(fig_flag_bucket, orientation="v", value_type="percent", use_text_field=True, text_as_percent=True)
                    fig_flag_bucket.update_layout(
                        xaxis_title="Create Month",
                        yaxis_title="Share within that Month-BGI (%)",
                        legend_title="TAT Flag Bucket",
                    )
                    render_plotly_chart(fig_flag_bucket, use_container_width=True)

                show_cols = [
                    "create_month",
                    "bgi_desc_value",
                    "cases",
                    "avg_tat_days",
                    "p90_tat_days",
                    "tat_5_7_pct",
                    "tat_7_plus_pct",
                    "tat_5_plus_pct",
                    "flag",
                ]
                render_dataframe(
                    bgi_top_month[show_cols].style.format(
                        {
                            "cases": "{:,.0f}",
                            "avg_tat_days": "{:.2f}",
                            "p90_tat_days": "{:.2f}",
                            "tat_5_7_pct": "{:.2f}%",
                            "tat_7_plus_pct": "{:.2f}%",
                            "tat_5_plus_pct": "{:.2f}%",
                        },
                        na_rep="NA",
                    ),
                    use_container_width=True,
                )

with tab_reson:
    st.subheader("Reason")
    st.caption("High TAT view (Net TAT > 7 days): Top 5 by dimension with average TAT and percent share.")

    reson_base = completed_df[completed_df["net_tat_days"] > 7].copy()
    reson_base["request_type_value"] = reson_base["request_type_value"].astype("string").fillna("Unknown").replace("", "Unknown")
    reson_base["bgi_desc_value"] = reson_base["bgi_desc_value"].astype("string").fillna("Unknown").replace("", "Unknown")
    reson_base["underwriting_segment_value"] = (
        reson_base["underwriting_segment_value"].astype("string").fillna("Unknown").replace("", "Unknown")
    )
    reson_base["agent_broker_value"] = reson_base["agent_broker_value"].astype("string").fillna("Unknown").replace("", "Unknown")

    if reson_base.empty:
        st.info("No completed cases with Net TAT > 7 days available for Reason analysis.")
    else:
        def build_reson_summary(source_df: pd.DataFrame, dim_col: str) -> pd.DataFrame:
            out = (
                source_df.groupby(dim_col, as_index=False)
                .agg(
                    cases=("request_id", "size"),
                    avg_tat_days=("net_tat_days", "mean"),
                )
                .rename(columns={dim_col: "data_point"})
                .sort_values(["cases", "avg_tat_days"], ascending=[False, False])
            )
            out["share_pct"] = out["cases"].apply(lambda x: pct_value(x, len(source_df)))
            return out.head(5)

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
                    title=f"{y_label}: Top 5 (TAT > 7 days) Avg TAT and % share",
                    hover_data={"cases": ":,.0f", "share_pct": ":.2f", "avg_tat_days": ":.2f"},
                )
                fig.update_traces(texttemplate="%{text:.1f}%", textposition="inside")
                fig.update_layout(xaxis_title="Average Net TAT (days)", yaxis_title=y_label)
                render_plotly_chart(fig, use_container_width=True)
            with c2:
                render_dataframe(
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
        render_dataframe(
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
    render_dataframe(filtered[available_cols].head(500), use_container_width=True)

st.caption(
    "Definitions: Completed = completedDateTime present; StraightThrough = onHoldReasonDescriptionsHistory empty; "
    "Net TAT = completedDateTime - createDateTime - valid holding time."
)
