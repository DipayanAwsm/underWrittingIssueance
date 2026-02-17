import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
from pathlib import Path
import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from io import BytesIO
import base64

# Page configuration
st.set_page_config(
    page_title="Underwriting Issuance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    </style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data(file_path_or_buffer):
    """Load data from CSV or Excel file"""
    try:
        if isinstance(file_path_or_buffer, str):
            # File path
            if file_path_or_buffer.endswith('.csv'):
                df = pd.read_csv(file_path_or_buffer, low_memory=False)
            elif file_path_or_buffer.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path_or_buffer, engine='openpyxl')
            else:
                return None
        else:
            # File buffer (uploaded file)
            if file_path_or_buffer.name.endswith('.csv'):
                df = pd.read_csv(file_path_or_buffer, low_memory=False)
            elif file_path_or_buffer.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path_or_buffer, engine='openpyxl')
            else:
                return None
        return df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

def parse_separated_values(value, separator=None):
    """Parse values that may be separated by ',' (with or without space) or '|'"""
    if pd.isna(value) or value == '':
        return []

    value = str(value).strip()
    if not value:
        return []

    if separator:
        separators = [separator]
    else:
        # Try to detect separator: pipe, comma-space, then plain comma
        if '|' in value:
            separators = ['|']
        elif ', ' in value:
            separators = [', ']
        elif ',' in value:
            separators = [',']
        else:
            return [value]

    result = []
    for sep in separators:
        if sep in value:
            result = [v.strip() for v in value.split(sep) if v.strip()]
            break

    return result if result else [value]

def calculate_aging(df):
    """Calculate aging in days for non-completed cases"""
    df = df.copy()
    df['createDateTime'] = pd.to_datetime(df['createDateTime'], errors='coerce')
    df['completedDateTime'] = pd.to_datetime(df['completedDateTime'], errors='coerce')
    
    # For non-completed cases
    mask = df['completedDateTime'].isna()
    df.loc[mask, 'Aging_Days'] = (datetime.now() - df.loc[mask, 'createDateTime']).dt.days
    df.loc[~mask, 'Aging_Days'] = np.nan
    
    return df

def classify_case_type(on_hold_reasons):
    """Classify case as straight-through, one-touch, or multi-hold"""
    reasons = parse_separated_values(on_hold_reasons)
    
    if len(reasons) == 0:
        return 'Straight Through'
    elif len(reasons) == 1:
        return 'One Touch'
    else:
        return f'Multi Hold ({len(reasons)} touches)'

def calculate_holding_times(df):
    """Calculate holding time for each onHoldReasonDescriptionsHistory entry"""
    df = df.copy()
    
    holding_times_list = []
    
    for idx, row in df.iterrows():
        on_hold_dates = parse_separated_values(row.get('onHoldDatesHistory', ''))
        off_hold_dates = parse_separated_values(row.get('offHoldDatesHistory', ''))
        on_hold_reasons = parse_separated_values(row.get('onHoldReasonDescriptionsHistory', ''))
        
        # Parse dates
        on_hold_dates_parsed = []
        for date_str in on_hold_dates:
            try:
                date_obj = pd.to_datetime(date_str, errors='coerce')
                if pd.notna(date_obj):
                    on_hold_dates_parsed.append(date_obj)
            except:
                pass
        
        off_hold_dates_parsed = []
        for date_str in off_hold_dates:
            try:
                date_obj = pd.to_datetime(date_str, errors='coerce')
                if pd.notna(date_obj):
                    off_hold_dates_parsed.append(date_obj)
            except:
                pass
        
        # Calculate holding times
        hold_times = []
        for i, on_date in enumerate(on_hold_dates_parsed):
            if i < len(off_hold_dates_parsed):
                off_date = off_hold_dates_parsed[i]
                if pd.notna(on_date) and pd.notna(off_date):
                    hold_time = (off_date - on_date).days
                    hold_times.append(hold_time)
            else:
                # If no off date, calculate from on date to now
                if pd.notna(on_date):
                    hold_time = (datetime.now() - on_date).days
                    hold_times.append(hold_time)
        
        holding_times_list.append(hold_times)
    
    df['HoldingTimes'] = holding_times_list
    df['TotalHoldingTime'] = df['HoldingTimes'].apply(lambda x: sum(x) if x else 0)
    df['NumberOfTouches'] = df['HoldingTimes'].apply(lambda x: len(x) if x else 0)
    
    return df

def calculate_tat(df):
    """Calculate TurnAroundTime (TAT) in days for completed cases"""
    df = df.copy()
    df['createDateTime'] = pd.to_datetime(df['createDateTime'], errors='coerce')
    df['completedDateTime'] = pd.to_datetime(df['completedDateTime'], errors='coerce')
    
    # Calculate TAT for completed cases
    mask = df['completedDateTime'].notna()
    df.loc[mask, 'TAT_Days'] = (df.loc[mask, 'completedDateTime'] - df.loc[mask, 'createDateTime']).dt.days
    df.loc[~mask, 'TAT_Days'] = np.nan
    
    return df

def create_tat_buckets(tat_days):
    """Create TAT buckets: 0-5, 5-7, 7+"""
    if pd.isna(tat_days):
        return None
    if tat_days <= 5:
        return '0-5 days'
    elif tat_days <= 7:
        return '5-7 days'
    else:
        return '7+ days'

def main():
    st.markdown('<h1 class="main-header">📊 Underwriting Issuance Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar for data upload
    st.sidebar.header("📁 Data Source")
    
    data_source = st.sidebar.radio(
        "Choose data source:",
        ["Upload File", "Data Folder"]
    )
    
    df = None
    
    if data_source == "Upload File":
        uploaded_file = st.sidebar.file_uploader(
            "Upload CSV or Excel file",
            type=['csv', 'xlsx', 'xls']
        )
        if uploaded_file is not None:
            df = load_data(uploaded_file)
    else:
        # Load from data folder
        data_folder = Path("data")
        if data_folder.exists():
            csv_files = list(data_folder.glob("*.csv"))
            xlsx_files = list(data_folder.glob("*.xlsx")) + list(data_folder.glob("*.xls"))
            all_files = csv_files + xlsx_files
            
            if all_files:
                selected_file = st.sidebar.selectbox(
                    "Select file from data folder:",
                    [f.name for f in all_files]
                )
                if selected_file:
                    file_path = data_folder / selected_file
                    df = load_data(str(file_path))
            else:
                st.sidebar.warning("No CSV or Excel files found in data folder")
    
    if df is None or df.empty:
        st.info("👆 Please upload a file or select from data folder to begin analysis")
        return
    
    st.sidebar.success(f"✅ Loaded {len(df)} rows")
    
    # Save original data before any processing/filters (for Excel export)
    original_df = df.copy()
    
    # Data processing
    with st.spinner("Processing data..."):
        # Calculate aging
        df = calculate_aging(df)
        
        # Classify case types
        df['CaseType'] = df['onHoldReasonDescriptionsHistory'].apply(classify_case_type)
        
        # Calculate holding times
        df = calculate_holding_times(df)
        
        # Calculate TAT
        df = calculate_tat(df)
        
        # Create TAT buckets
        df['TAT_Bucket'] = df['TAT_Days'].apply(create_tat_buckets)
        
        # Extract month from createDateTime for seasonality
        df['createDateTime'] = pd.to_datetime(df['createDateTime'], errors='coerce')
        df['Month'] = df['createDateTime'].dt.to_period('M')
        df['Month_Str'] = df['Month'].astype(str)
    
    # High-level filters
    st.sidebar.header("🔍 Filters")
    if 'statusDescription' in df.columns:
        available_statuses = sorted(df['statusDescription'].dropna().unique().tolist())
        selected_statuses = st.sidebar.multiselect(
            "Filter by Status Description",
            options=available_statuses,
            default=available_statuses,
            help="Select one or more statuses to filter all metrics and charts."
        )
        if selected_statuses:
            df = df[df['statusDescription'].isin(selected_statuses)]
    else:
        st.sidebar.warning("Column 'statusDescription' not found in data for filtering.")
    
    # Main dashboard
    st.header("📈 Key Metrics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_cases = len(df)
        st.metric("Total Cases", f"{total_cases:,}")
    
    with col2:
        completed_cases = df['completedDateTime'].notna().sum()
        st.metric("Completed Cases", f"{completed_cases:,}")
    
    with col3:
        median_tat = df['TAT_Days'].median()
        st.metric("Median TAT (days)", f"{median_tat:.1f}" if not pd.isna(median_tat) else "N/A")
    
    with col4:
        avg_aging = df['Aging_Days'].median()
        st.metric("Median Aging (days)", f"{avg_aging:.1f}" if not pd.isna(avg_aging) else "N/A")
    
    # 1. Status Count
    st.header("1. Status Count Analysis")
    status_counts = df['statusDescription'].value_counts()
    
    col1, col2 = st.columns(2)
    with col1:
        fig_status = px.bar(
            x=status_counts.index,
            y=status_counts.values,
            title="Status Count",
            labels={'x': 'Status', 'y': 'Count'}
        )
        fig_status.update_xaxes(tickangle=45)
        st.plotly_chart(fig_status, use_container_width=True)
    
    with col2:
        st.dataframe(status_counts.reset_index().rename(columns={'index': 'Status', 'statusDescription': 'Count'}))
    
    # 2. Aging Analysis
    st.header("2. Aging Analysis (Non-Completed Cases)")
    aging_data = df[df['Aging_Days'].notna()].copy()
    
    if not aging_data.empty:
        col1, col2 = st.columns(2)
        with col1:
            fig_aging = px.histogram(
                aging_data,
                x='Aging_Days',
                nbins=30,
                title="Distribution of Aging Days",
                labels={'Aging_Days': 'Aging (Days)', 'count': 'Frequency'}
            )
            st.plotly_chart(fig_aging, use_container_width=True)
        
        with col2:
            st.metric("Median Aging", f"{aging_data['Aging_Days'].median():.1f} days")
            st.metric("Mean Aging", f"{aging_data['Aging_Days'].mean():.1f} days")
            st.metric("Max Aging", f"{aging_data['Aging_Days'].max():.0f} days")
    else:
        st.info("No non-completed cases found")
    
    # 3. Case Type Classification
    st.header("3. Case Type Classification")
    case_type_counts = df['CaseType'].value_counts()
    
    col1, col2 = st.columns(2)
    with col1:
        fig_case_type = px.pie(
            values=case_type_counts.values,
            names=case_type_counts.index,
            title="Case Type Distribution"
        )
        st.plotly_chart(fig_case_type, use_container_width=True)
    
    with col2:
        st.dataframe(case_type_counts.reset_index().rename(columns={'index': 'Case Type', 'CaseType': 'Count'}))
    
    # Extract straight-through and multi-hold
    straight_through = df[df['CaseType'] == 'Straight Through']
    multi_hold = df[df['CaseType'].str.contains('Multi Hold', na=False)]
    one_touch = df[df['CaseType'] == 'One Touch']
    
    st.subheader("Case Type Summary")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Straight Through", f"{len(straight_through):,}")
    with col2:
        st.metric("One Touch", f"{len(one_touch):,}")
    with col3:
        st.metric("Multi Hold", f"{len(multi_hold):,}")
    
    # 4. Holding Time Analysis
    st.header("4. Holding Time Analysis")
    holding_data = df[df['TotalHoldingTime'] > 0].copy()
    
    if not holding_data.empty:
        col1, col2 = st.columns(2)
        with col1:
            fig_holding = px.histogram(
                holding_data,
                x='TotalHoldingTime',
                nbins=30,
                title="Distribution of Total Holding Time (Days)",
                labels={'TotalHoldingTime': 'Holding Time (Days)', 'count': 'Frequency'}
            )
            st.plotly_chart(fig_holding, use_container_width=True)
        
        with col2:
            st.metric("Median Holding Time", f"{holding_data['TotalHoldingTime'].median():.1f} days")
            st.metric("Mean Holding Time", f"{holding_data['TotalHoldingTime'].mean():.1f} days")
            st.metric("Max Holding Time", f"{holding_data['TotalHoldingTime'].max():.0f} days")
            st.metric("Average Number of Touches", f"{holding_data['NumberOfTouches'].mean():.2f}")
    else:
        st.info("No holding time data available")
    
    # 5. TAT Analysis
    st.header("5. Turnaround Time (TAT) Analysis")
    tat_data = df[df['TAT_Days'].notna()].copy()
    
    if not tat_data.empty:
        col1, col2 = st.columns(2)
        with col1:
            fig_tat = px.histogram(
                tat_data,
                x='TAT_Days',
                nbins=30,
                title="Distribution of TAT (Days)",
                labels={'TAT_Days': 'TAT (Days)', 'count': 'Frequency'}
            )
            st.plotly_chart(fig_tat, use_container_width=True)
        
        with col2:
            st.metric("Median TAT", f"{tat_data['TAT_Days'].median():.1f} days")
            st.metric("Mean TAT", f"{tat_data['TAT_Days'].mean():.1f} days")
            st.metric("Min TAT", f"{tat_data['TAT_Days'].min():.0f} days")
            st.metric("Max TAT", f"{tat_data['TAT_Days'].max():.0f} days")
    
    # 6. TAT Buckets
    st.header("6. TAT Buckets")
    tat_bucket_counts = df['TAT_Bucket'].value_counts()
    
    col1, col2 = st.columns(2)
    with col1:
        fig_buckets = px.bar(
            x=tat_bucket_counts.index,
            y=tat_bucket_counts.values,
            title="TAT Bucket Distribution",
            labels={'x': 'TAT Bucket', 'y': 'Count'}
        )
        st.plotly_chart(fig_buckets, use_container_width=True)
    
    with col2:
        st.dataframe(tat_bucket_counts.reset_index().rename(columns={'index': 'TAT Bucket', 'TAT_Bucket': 'Count'}))
    
    # 7. Median TAT by Case Type
    st.header("7. Median TAT by Case Type")
    tat_by_case_type = df.groupby('CaseType')['TAT_Days'].median().sort_values(ascending=False)
    
    col1, col2 = st.columns(2)
    with col1:
        fig_tat_case = px.bar(
            x=tat_by_case_type.index,
            y=tat_by_case_type.values,
            title="Median TAT by Case Type",
            labels={'x': 'Case Type', 'y': 'Median TAT (Days)'}
        )
        fig_tat_case.update_xaxes(tickangle=45)
        st.plotly_chart(fig_tat_case, use_container_width=True)
    
    with col2:
        st.dataframe(tat_by_case_type.reset_index().rename(columns={'index': 'Case Type', 'TAT_Days': 'Median TAT (Days)'}))
    
    # 7b. Median TAT by Single Hold vs Multi-Hold (by TAT Buckets)
    st.header("7b. Median TAT by Single Hold vs Multi-Hold (by TAT Buckets)")
    
    # Create simplified case type for single vs multi
    df['HoldCategory'] = df['CaseType'].apply(lambda x: 'Straight Through' if x == 'Straight Through' 
                                               else 'Single Hold' if x == 'One Touch' 
                                               else 'Multi Hold')
    
    # Calculate median TAT by HoldCategory and TAT_Bucket
    tat_by_hold_bucket_raw = df.groupby(['HoldCategory', 'TAT_Bucket'])['TAT_Days'].median().reset_index()
    if not tat_by_hold_bucket_raw.empty:
        tat_by_hold_bucket = tat_by_hold_bucket_raw.pivot(index='TAT_Bucket', columns='HoldCategory', values='TAT_Days')
    else:
        tat_by_hold_bucket = pd.DataFrame()
    
    # Also show median TAT by HoldCategory overall
    tat_by_hold_category = df.groupby('HoldCategory')['TAT_Days'].median().sort_values(ascending=False)
    
    col1, col2 = st.columns(2)
    with col1:
        if not tat_by_hold_bucket.empty:
            fig_tat_hold_bucket = px.bar(
                tat_by_hold_bucket.reset_index(),
                x='TAT_Bucket',
                y=[col for col in tat_by_hold_bucket.columns if col != 'TAT_Bucket'],
                title="Median TAT by Hold Category and TAT Bucket",
                labels={'value': 'Median TAT (Days)', 'TAT_Bucket': 'TAT Bucket'},
                barmode='group'
            )
            st.plotly_chart(fig_tat_hold_bucket, use_container_width=True)
    
    with col2:
        st.dataframe(tat_by_hold_category.reset_index().rename(columns={'index': 'Hold Category', 'TAT_Days': 'Median TAT (Days)'}))
        
        if not tat_by_hold_bucket.empty:
            st.dataframe(tat_by_hold_bucket)
    
    # 8. Drill-down Analysis
    st.header("8. Drill-Down Analysis")
    
    drill_down_category = st.selectbox(
        "Select category for drill-down:",
        ["Straight Through", "Multi Hold", "By TAT Bucket", "By Number of Touches"]
    )
    
    if drill_down_category == "Straight Through":
        drill_data = straight_through.copy()
        st.subheader(f"Drill-down: Straight Through Cases ({len(drill_data)} cases)")
    elif drill_down_category == "Multi Hold":
        drill_data = multi_hold.copy()
        st.subheader(f"Drill-down: Multi Hold Cases ({len(drill_data)} cases)")
    elif drill_down_category == "By TAT Bucket":
        selected_bucket = st.selectbox("Select TAT Bucket:", df['TAT_Bucket'].dropna().unique())
        drill_data = df[df['TAT_Bucket'] == selected_bucket].copy()
        st.subheader(f"Drill-down: TAT Bucket = {selected_bucket} ({len(drill_data)} cases)")
    else:  # By Number of Touches
        selected_touches = st.selectbox("Select Number of Touches:", sorted(df['NumberOfTouches'].unique()))
        drill_data = df[df['NumberOfTouches'] == selected_touches].copy()
        st.subheader(f"Drill-down: Number of Touches = {selected_touches} ({len(drill_data)} cases)")
    
    if not drill_data.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Top 10 Request Types")
            top_request_types = drill_data['requestTypeDescription'].value_counts().head(10)
            if not top_request_types.empty:
                fig_request = px.bar(
                    x=top_request_types.values,
                    y=top_request_types.index,
                    orientation='h',
                    title="Top 10 Request Types",
                    labels={'x': 'Count', 'y': 'Request Type'}
                )
                st.plotly_chart(fig_request, use_container_width=True)
            else:
                st.info("No request type data available")
        
        with col2:
            st.subheader("Top 10 Hold Reasons")
            # Flatten hold reasons
            all_hold_reasons = []
            for reasons in drill_data['onHoldReasonDescriptionsHistory'].apply(parse_separated_values):
                all_hold_reasons.extend(reasons)
            
            if all_hold_reasons:
                hold_reason_counts = pd.Series(all_hold_reasons).value_counts().head(10)
                fig_hold = px.bar(
                    x=hold_reason_counts.values,
                    y=hold_reason_counts.index,
                    orientation='h',
                    title="Top 10 Hold Reasons",
                    labels={'x': 'Count', 'y': 'Hold Reason'}
                )
                st.plotly_chart(fig_hold, use_container_width=True)
            else:
                st.info("No hold reasons found")
    else:
        st.info("No data available for selected drill-down")
    
    # 9. Holding Time Statistics
    st.header("9. Holding Time Statistics")
    if not holding_data.empty:
        all_holding_times = []
        for times in holding_data['HoldingTimes']:
            all_holding_times.extend(times)
        
        if all_holding_times:
            holding_times_series = pd.Series(all_holding_times)
            
            # Calculate median of highest and lowest holding times per case
            highest_per_case = holding_data['HoldingTimes'].apply(lambda x: max(x) if x else 0)
            lowest_per_case = holding_data['HoldingTimes'].apply(lambda x: min(x) if x else 0)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Median Holding Time (All)", f"{holding_times_series.median():.1f} days")
            with col2:
                st.metric("Median of Highest Holding Time", f"{highest_per_case.median():.1f} days")
            with col3:
                st.metric("Median of Lowest Holding Time", f"{lowest_per_case.median():.1f} days")
            with col4:
                st.metric("Average Number of Touches", f"{holding_data['NumberOfTouches'].mean():.2f}")
            
            # Additional metrics
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Highest Holding Time (Overall)", f"{holding_times_series.max():.0f} days")
            with col2:
                st.metric("Lowest Holding Time (Overall)", f"{holding_times_series.min():.0f} days")
            
            # Visualization
            col1, col2 = st.columns(2)
            with col1:
                fig_holding_dist = px.histogram(
                    holding_times_series,
                    nbins=30,
                    title="Distribution of Individual Holding Times",
                    labels={'value': 'Holding Time (Days)', 'count': 'Frequency'}
                )
                st.plotly_chart(fig_holding_dist, use_container_width=True)
            
            with col2:
                fig_high_low = go.Figure()
                fig_high_low.add_trace(go.Box(y=highest_per_case, name='Highest per Case', boxmean='sd'))
                fig_high_low.add_trace(go.Box(y=lowest_per_case, name='Lowest per Case', boxmean='sd'))
                fig_high_low.update_layout(
                    title="Highest vs Lowest Holding Times per Case",
                    yaxis_title="Holding Time (Days)"
                )
                st.plotly_chart(fig_high_low, use_container_width=True)
    
    # 10. Seasonality Analysis
    st.header("10. Seasonality Analysis by Month")
    
    monthly_stats = df.groupby('Month_Str').agg({
        'requestId': 'count',
        'completedDateTime': lambda x: x.notna().sum(),
        'TAT_Days': 'median',
        'TotalHoldingTime': 'median',
        'NumberOfTouches': 'mean'
    }).rename(columns={
        'requestId': 'Total_Cases',
        'completedDateTime': 'Completed_Cases',
        'TAT_Days': 'Median_TAT',
        'TotalHoldingTime': 'Median_Holding_Time',
        'NumberOfTouches': 'Avg_Touches'
    })
    
    monthly_stats['Completion_Rate'] = (monthly_stats['Completed_Cases'] / monthly_stats['Total_Cases'] * 100).round(2)
    monthly_stats['Holding_Rate'] = (df.groupby('Month_Str')['TotalHoldingTime'].apply(lambda x: (x > 0).sum() / len(x) * 100)).round(2)
    
    col1, col2 = st.columns(2)
    with col1:
        fig_completion = px.line(
            monthly_stats.reset_index(),
            x='Month_Str',
            y='Completion_Rate',
            title="Monthly Completion Rate (%)",
            labels={'Month_Str': 'Month', 'Completion_Rate': 'Completion Rate (%)'}
        )
        fig_completion.update_xaxes(tickangle=45)
        st.plotly_chart(fig_completion, use_container_width=True)
    
    with col2:
        fig_holding_rate = px.line(
            monthly_stats.reset_index(),
            x='Month_Str',
            y='Holding_Rate',
            title="Monthly Holding Rate (%)",
            labels={'Month_Str': 'Month', 'Holding_Rate': 'Holding Rate (%)'}
        )
        fig_holding_rate.update_xaxes(tickangle=45)
        st.plotly_chart(fig_holding_rate, use_container_width=True)
    
    st.dataframe(monthly_stats)
    
    # 11. Analysis by Location and Vehicles
    st.header("11. Analysis by Location and Number of Vehicles")
    
    if 'numberOfLocations' in df.columns and 'NumberOfVehicles' in df.columns:
        # Filter out null values
        location_vehicle_data = df[
            (df['numberOfLocations'].notna()) & 
            (df['NumberOfVehicles'].notna())
        ].copy()
        
        if not location_vehicle_data.empty:
            # Group by number of locations
            location_analysis = location_vehicle_data.groupby('numberOfLocations').agg({
                'TAT_Days': 'median',
                'NumberOfTouches': 'mean',
                'TotalHoldingTime': 'median',
                'requestId': 'count'
            }).rename(columns={
                'TAT_Days': 'Median_TAT',
                'NumberOfTouches': 'Avg_Touches',
                'TotalHoldingTime': 'Median_Holding_Time',
                'requestId': 'Count'
            })
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("By Number of Locations")
                fig_location = px.bar(
                    location_analysis.reset_index(),
                    x='numberOfLocations',
                    y='Median_TAT',
                    title="Median TAT by Number of Locations",
                    labels={'numberOfLocations': 'Number of Locations', 'Median_TAT': 'Median TAT (Days)'}
                )
                st.plotly_chart(fig_location, use_container_width=True)
                st.dataframe(location_analysis)
            
            # Group by number of vehicles
            vehicle_analysis = location_vehicle_data.groupby('NumberOfVehicles').agg({
                'TAT_Days': 'median',
                'NumberOfTouches': 'mean',
                'TotalHoldingTime': 'median',
                'requestId': 'count'
            }).rename(columns={
                'TAT_Days': 'Median_TAT',
                'NumberOfTouches': 'Avg_Touches',
                'TotalHoldingTime': 'Median_Holding_Time',
                'requestId': 'Count'
            })
            
            with col2:
                st.subheader("By Number of Vehicles")
                fig_vehicle = px.bar(
                    vehicle_analysis.reset_index(),
                    x='NumberOfVehicles',
                    y='Median_TAT',
                    title="Median TAT by Number of Vehicles",
                    labels={'NumberOfVehicles': 'Number of Vehicles', 'Median_TAT': 'Median TAT (Days)'}
                )
                st.plotly_chart(fig_vehicle, use_container_width=True)
                st.dataframe(vehicle_analysis)
            
            # Correlation analysis
            st.subheader("Correlation: Number of Vehicles vs Touches/Holding Time")
            col1, col2 = st.columns(2)
            with col1:
                fig_corr_touches = px.scatter(
                    location_vehicle_data,
                    x='NumberOfVehicles',
                    y='NumberOfTouches',
                    title="Number of Vehicles vs Number of Touches",
                    trendline="ols"
                )
                st.plotly_chart(fig_corr_touches, use_container_width=True)
            
            with col2:
                fig_corr_holding = px.scatter(
                    location_vehicle_data,
                    x='NumberOfVehicles',
                    y='TotalHoldingTime',
                    title="Number of Vehicles vs Holding Time",
                    trendline="ols"
                )
                st.plotly_chart(fig_corr_holding, use_container_width=True)
        else:
            st.info("No data available for location/vehicle analysis")
    else:
        st.warning("Required columns (numberOfLocations, NumberOfVehicles) not found in data")
    
    # Export to Excel
    st.header("📥 Export to Excel")
    
    if st.button("Generate Excel Report"):
        with st.spinner("Generating Excel report..."):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Summary sheet
                all_holding_times_flat = []
                if not holding_data.empty:
                    for times in holding_data['HoldingTimes']:
                        all_holding_times_flat.extend(times)
                
                highest_per_case = holding_data['HoldingTimes'].apply(lambda x: max(x) if x else 0) if not holding_data.empty else pd.Series()
                lowest_per_case = holding_data['HoldingTimes'].apply(lambda x: min(x) if x else 0) if not holding_data.empty else pd.Series()
                
                summary_data = {
                    'Metric': [
                        'Total Cases', 
                        'Completed Cases', 
                        'Median TAT (days)', 
                        'Median Aging (days)',
                        'Straight Through Cases', 
                        'One Touch Cases', 
                        'Multi Hold Cases',
                        'Median Holding Time (All)',
                        'Median of Highest Holding Time',
                        'Median of Lowest Holding Time',
                        'Average Number of Touches',
                        'Highest Holding Time (Overall)',
                        'Lowest Holding Time (Overall)'
                    ],
                    'Value': [
                        len(df),
                        df['completedDateTime'].notna().sum(),
                        df['TAT_Days'].median() if not df['TAT_Days'].isna().all() else 0,
                        df['Aging_Days'].median() if not df['Aging_Days'].isna().all() else 0,
                        len(straight_through),
                        len(one_touch),
                        len(multi_hold),
                        pd.Series(all_holding_times_flat).median() if all_holding_times_flat else 0,
                        highest_per_case.median() if not highest_per_case.empty else 0,
                        lowest_per_case.median() if not lowest_per_case.empty else 0,
                        holding_data['NumberOfTouches'].mean() if not holding_data.empty else 0,
                        max(all_holding_times_flat) if all_holding_times_flat else 0,
                        min(all_holding_times_flat) if all_holding_times_flat else 0
                    ]
                }
                pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
                
                # Status counts
                status_counts_df = status_counts.reset_index()
                status_counts_df.columns = ['Status', 'Count']
                status_counts_df.to_excel(writer, sheet_name='Status_Counts', index=False)
                
                # Case type classification
                case_type_counts_df = case_type_counts.reset_index()
                case_type_counts_df.columns = ['Case_Type', 'Count']
                case_type_counts_df.to_excel(writer, sheet_name='Case_Type_Classification', index=False)
                
                # TAT buckets
                tat_bucket_counts_df = tat_bucket_counts.reset_index()
                tat_bucket_counts_df.columns = ['TAT_Bucket', 'Count']
                tat_bucket_counts_df.to_excel(writer, sheet_name='TAT_Buckets', index=False)
                
                # Median TAT by Case Type
                tat_by_case_type_df = tat_by_case_type.reset_index()
                tat_by_case_type_df.columns = ['Case_Type', 'Median_TAT_Days']
                tat_by_case_type_df.to_excel(writer, sheet_name='Median_TAT_by_CaseType', index=False)
                
                # Median TAT by Hold Category and TAT Bucket
                if not tat_by_hold_bucket.empty:
                    tat_by_hold_bucket.to_excel(writer, sheet_name='Median_TAT_by_Hold_Bucket', index=True)
                
                # Median TAT by Hold Category
                tat_by_hold_category_df = tat_by_hold_category.reset_index()
                tat_by_hold_category_df.columns = ['Hold_Category', 'Median_TAT_Days']
                tat_by_hold_category_df.to_excel(writer, sheet_name='Median_TAT_by_HoldCategory', index=False)
                
                # Monthly statistics
                monthly_stats.to_excel(writer, sheet_name='Monthly_Statistics', index=True)
                
                # Location analysis
                if 'numberOfLocations' in df.columns:
                    try:
                        if 'location_vehicle_data' in locals() and not location_vehicle_data.empty and 'location_analysis' in locals():
                            location_analysis.to_excel(writer, sheet_name='Location_Analysis', index=True)
                    except:
                        pass
                
                # Vehicle analysis
                if 'NumberOfVehicles' in df.columns:
                    try:
                        if 'location_vehicle_data' in locals() and not location_vehicle_data.empty and 'vehicle_analysis' in locals():
                            vehicle_analysis.to_excel(writer, sheet_name='Vehicle_Analysis', index=True)
                    except:
                        pass
                
                # Holding Time Statistics
                if not holding_data.empty and all_holding_times_flat:
                    holding_stats = {
                        'Metric': [
                            'Median Holding Time (All)',
                            'Median of Highest Holding Time',
                            'Median of Lowest Holding Time',
                            'Highest Holding Time (Overall)',
                            'Lowest Holding Time (Overall)',
                            'Average Number of Touches'
                        ],
                        'Value': [
                            pd.Series(all_holding_times_flat).median(),
                            highest_per_case.median(),
                            lowest_per_case.median(),
                            max(all_holding_times_flat),
                            min(all_holding_times_flat),
                            holding_data['NumberOfTouches'].mean()
                        ]
                    }
                    pd.DataFrame(holding_stats).to_excel(writer, sheet_name='Holding_Time_Stats', index=False)
                
                # Original data (all original columns before processing)
                original_df.to_excel(writer, sheet_name='Original_Data', index=False)
                
                # Processed data (limit to essential columns to avoid Excel size issues)
                essential_cols = [
                    'requestId', 'statusDescription', 'createDateTime', 'completedDateTime',
                    'onHoldReasonDescriptionsHistory', 'onHoldDatesHistory', 'offHoldDatesHistory',
                    'requestTypeDescription', 'onHoldReasonDescription', 'numberOfLocations', 
                    'NumberOfVehicles', 'CaseType', 'HoldCategory', 'Aging_Days', 'TAT_Days', 
                    'TAT_Bucket', 'TotalHoldingTime', 'NumberOfTouches', 'Month_Str'
                ]
                available_cols = [col for col in essential_cols if col in df.columns]
                df[available_cols].to_excel(writer, sheet_name='Processed_Data', index=False)
            
            output.seek(0)
            st.download_button(
                label="Download Excel Report",
                data=output,
                file_name=f"underwriting_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Excel report generated successfully!")

if __name__ == "__main__":
    main()

