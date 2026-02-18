"""
Advanced Underwriting Analytics Dashboard

This Streamlit app provides comprehensive analytics focusing on:
- Holding reasons
- Holding time
- TAT (Turnaround Time)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from pathlib import Path

# Page configuration
st.set_page_config(
    page_title="Underwriting Analytics Dashboard",
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
    </style>
""", unsafe_allow_html=True)


@st.cache_data
def load_data(file_path_or_buffer):
    """Load data from CSV or Excel file"""
    try:
        if isinstance(file_path_or_buffer, str):
            if file_path_or_buffer.endswith('.csv'):
                df = pd.read_csv(file_path_or_buffer, low_memory=False)
            elif file_path_or_buffer.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path_or_buffer, engine='openpyxl')
            else:
                return None
        else:
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


def classify_case_type_by_reasons(on_hold_reasons_history):
    """Classify case type based on number of touches from onHoldReasonDescriptionsHistory"""
    reasons = parse_separated_values(on_hold_reasons_history)
    num_touches = len(reasons)
    
    if num_touches == 0:
        return 'Straight Through', 0
    elif num_touches == 1:
        return 'One Touch', 1
    else:
        return f'Multi Hold ({num_touches} touches)', num_touches


def calculate_tat(df):
    """Calculate TurnaroundTime (TAT) in days for completed cases"""
    df = df.copy()
    df['createDateTime'] = pd.to_datetime(df.get('createDateTime'), errors='coerce')
    df['completedDateTime'] = pd.to_datetime(df.get('completedDateTime'), errors='coerce')
    
    mask = df['completedDateTime'].notna()
    df.loc[mask, 'TAT_Days'] = (
        df.loc[mask, 'completedDateTime'] - df.loc[mask, 'createDateTime']
    ).dt.days
    df.loc[~mask, 'TAT_Days'] = np.nan
    
    return df


def calculate_holding_times(df):
    """Calculate holding time for each onHoldReasonDescriptionsHistory entry"""
    df = df.copy()
    holding_times_list = []
    
    for _, row in df.iterrows():
        on_hold_dates = parse_separated_values(row.get('onHoldDatesHistory', ''))
        off_hold_dates = parse_separated_values(row.get('offHoldDatesHistory', ''))
        
        on_parsed = [pd.to_datetime(d, errors='coerce') for d in on_hold_dates]
        on_parsed = [d for d in on_parsed if pd.notna(d)]
        
        off_parsed = [pd.to_datetime(d, errors='coerce') for d in off_hold_dates]
        off_parsed = [d for d in off_parsed if pd.notna(d)]
        
        hold_times = []
        for i, on_date in enumerate(on_parsed):
            if i < len(off_parsed):
                off_date = off_parsed[i]
                if pd.notna(on_date) and pd.notna(off_date):
                    hold_times.append((off_date - on_date).days)
            else:
                if pd.notna(on_date):
                    hold_times.append((datetime.now() - on_date).days)
        
        holding_times_list.append(hold_times)
    
    df['HoldingTimes'] = holding_times_list
    df['TotalHoldingTime'] = df['HoldingTimes'].apply(lambda x: sum(x) if x else 0)
    df['NumberOfTouches'] = df['HoldingTimes'].apply(lambda x: len(x) if x else 0)
    return df


def main():
    st.markdown('<h1 class="main-header">📊 Underwriting Analytics Dashboard</h1>', unsafe_allow_html=True)
    
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
    
    # Data processing
    with st.spinner("Processing data..."):
        # Calculate TAT
        df = calculate_tat(df)
        
        # Case Type Classification based on onHoldReasonDescriptionsHistory
        if 'onHoldReasonDescriptionsHistory' in df.columns:
            case_type_results = df['onHoldReasonDescriptionsHistory'].apply(classify_case_type_by_reasons)
            df['CaseType'] = [r[0] for r in case_type_results]
            # Number of touches from reasons (will keep separately to avoid overwriting holding-based touches)
            df['Touches_Reasons'] = [r[1] for r in case_type_results]
        else:
            st.warning("⚠️ Column 'onHoldReasonDescriptionsHistory' not found.")
            df['CaseType'] = 'Unknown'
            df['Touches_Reasons'] = 0
        
        # Calculate holding times
        df = calculate_holding_times(df)
        
        # Extract month for seasonality
        if 'createDateTime' in df.columns:
            df['createDateTime'] = pd.to_datetime(df['createDateTime'], errors='coerce')
            df['Month'] = df['createDateTime'].dt.to_period('M')
            df['Month_Str'] = df['Month'].astype(str)
        else:
            df['Month_Str'] = None
    
    # ===================================================================
    # Filters
    # ===================================================================
    st.sidebar.header("🔍 Filters")
    
    # Filter 1: Created Date Range (Monthly Selector)
    if 'createDateTime' in df.columns and 'Month_Str' in df.columns:
        # Get unique months
        unique_months = sorted(df['Month_Str'].dropna().unique())
        
        # Create month options
        month_options = ["All"] + unique_months
        
        selected_month = st.sidebar.selectbox(
            "Created Date Range (Monthly):",
            options=month_options,
            index=0  # Default to "All"
        )
        
        if selected_month != "All":
            df = df[df['Month_Str'] == selected_month].copy()
            st.sidebar.info(f"📅 Filtered by month: {selected_month}")
        else:
            st.sidebar.info(f"📅 Showing all months")
    
    # Filter 2: Case Type
    if 'CaseType' in df.columns:
        filter_option = st.sidebar.radio(
            "Filter by Case Type:",
            ["All Cases", "Straight Through", "Multi Hold Only"],
            index=0  # Default to "All Cases"
        )
        
        original_count = len(df)
        
        if filter_option == "Straight Through":
            df = df[df['CaseType'] == 'Straight Through'].copy()
            filtered_count = len(df)
            st.sidebar.success(f"✅ Showing: {filtered_count:,} Straight Through cases")
            st.sidebar.info(f"📊 Excluded: {original_count - filtered_count:,} cases")
        elif filter_option == "Multi Hold Only":
            df = df[df['CaseType'].str.contains('Multi Hold', na=False)].copy()
            filtered_count = len(df)
            st.sidebar.success(f"✅ Showing: {filtered_count:,} Multi Hold cases")
            st.sidebar.info(f"📊 Excluded: {original_count - filtered_count:,} cases")
        else:
            st.sidebar.info(f"📊 Showing all {original_count:,} cases")
    
    # ===================================================================
    # 1. Total Number of Cases by Month
    # ===================================================================
    st.header("1. Total Number of Cases by Month")
    
    if 'Month_Str' in df.columns:
        monthly_cases = df.groupby('Month_Str').size().reset_index(name='Total_Cases')
        monthly_cases = monthly_cases.sort_values('Month_Str')
        
        fig_monthly = px.bar(
            monthly_cases,
            x='Month_Str',
            y='Total_Cases',
            title="Total Number of Cases by Month",
            labels={'Month_Str': 'Month', 'Total_Cases': 'Number of Cases'}
        )
        fig_monthly.update_xaxes(tickangle=45)
        st.plotly_chart(fig_monthly, use_container_width=True)
        
        st.dataframe(monthly_cases)
    
    # ===================================================================
    # 2. Status Description (Month-wise)
    # ===================================================================
    st.header("2. Status Description (Month-wise)")
    
    if 'statusDescription' in df.columns and 'Month_Str' in df.columns:
        status_monthly = df.groupby(['Month_Str', 'statusDescription']).size().reset_index(name='Count')
        status_monthly = status_monthly.sort_values(['Month_Str', 'Count'], ascending=[True, False])
        
        fig_status_monthly = px.bar(
            status_monthly,
            x='Month_Str',
            y='Count',
            color='statusDescription',
            title="Status Description Distribution by Month",
            labels={'Month_Str': 'Month', 'Count': 'Count'},
            barmode='stack'
        )
        fig_status_monthly.update_xaxes(tickangle=45)
        st.plotly_chart(fig_status_monthly, use_container_width=True)
        
        st.dataframe(status_monthly)
    
    # ===================================================================
    # 3. Side-by-side Chart: Straight Through vs Multi Hold (Month-wise)
    # ===================================================================
    st.header("3. Straight Through vs Multi Hold Cases (Month-wise Seasonality)")
    
    if 'CaseType' in df.columns and 'Month_Str' in df.columns:
        straight_through = df[df['CaseType'] == 'Straight Through']
        multi_hold = df[df['CaseType'].str.contains('Multi Hold', na=False)]
        
        st_monthly = straight_through.groupby('Month_Str').size().reset_index(name='Count')
        st_monthly['CaseType'] = 'Straight Through'
        
        mh_monthly = multi_hold.groupby('Month_Str').size().reset_index(name='Count')
        mh_monthly['CaseType'] = 'Multi Hold'
        
        combined = pd.concat([st_monthly, mh_monthly], ignore_index=True)
        combined = combined.sort_values('Month_Str')
        
        col1, col2 = st.columns(2)
        with col1:
            fig_st = px.line(
                st_monthly,
                x='Month_Str',
                y='Count',
                title="Straight Through Cases (Month-wise)",
                labels={'Month_Str': 'Month', 'Count': 'Number of Cases'}
            )
            fig_st.update_xaxes(tickangle=45)
            st.plotly_chart(fig_st, use_container_width=True)
        
        with col2:
            fig_mh = px.line(
                mh_monthly,
                x='Month_Str',
                y='Count',
                title="Multi Hold Cases (Month-wise)",
                labels={'Month_Str': 'Month', 'Count': 'Number of Cases'}
            )
            fig_mh.update_xaxes(tickangle=45)
            st.plotly_chart(fig_mh, use_container_width=True)
        
        # Combined chart
        fig_combined = px.line(
            combined,
            x='Month_Str',
            y='Count',
            color='CaseType',
            title="Straight Through vs Multi Hold Cases (Month-wise)",
            labels={'Month_Str': 'Month', 'Count': 'Number of Cases'}
        )
        fig_combined.update_xaxes(tickangle=45)
        st.plotly_chart(fig_combined, use_container_width=True)
        
        st.dataframe(combined.pivot(index='Month_Str', columns='CaseType', values='Count').fillna(0))
    
    # ===================================================================
    # 4. TAT Analysis for Completed Cases
    # ===================================================================
    st.header("4. Turnaround Time (TAT) Analysis - Completed Cases")
    
    if 'TAT_Days' in df.columns and 'Month_Str' in df.columns:
        completed_cases = df[df['TAT_Days'].notna()].copy()
        
        if not completed_cases.empty:
            tat_monthly = completed_cases.groupby('Month_Str').agg({
                'TAT_Days': ['median', 'mean']
            }).reset_index()
            tat_monthly.columns = ['Month', 'Median_TAT', 'Average_TAT']
            tat_monthly = tat_monthly.sort_values('Month')
            
            col1, col2 = st.columns(2)
            with col1:
                fig_median = px.line(
                    tat_monthly,
                    x='Month',
                    y='Median_TAT',
                    title="Median TAT by Month (Completed Cases)",
                    labels={'Month': 'Month', 'Median_TAT': 'Median TAT (Days)'},
                    markers=True
                )
                fig_median.update_xaxes(tickangle=45)
                st.plotly_chart(fig_median, use_container_width=True)
            
            with col2:
                fig_avg = px.line(
                    tat_monthly,
                    x='Month',
                    y='Average_TAT',
                    title="Average TAT by Month (Completed Cases)",
                    labels={'Month': 'Month', 'Average_TAT': 'Average TAT (Days)'},
                    markers=True
                )
                fig_avg.update_xaxes(tickangle=45)
                st.plotly_chart(fig_avg, use_container_width=True)
            
            st.dataframe(tat_monthly)
            
            # Top 5 Request Types by TAT
            if 'requestTypeDescription' in completed_cases.columns:
                st.subheader("4a. Top 5 Request Types by TAT")
                
                top5_request_tat = completed_cases.groupby('requestTypeDescription').agg({
                    'TAT_Days': ['median', 'mean', 'count']
                }).reset_index()
                top5_request_tat.columns = ['RequestType', 'Median_TAT', 'Average_TAT', 'Count']
                top5_request_tat = top5_request_tat.sort_values('Average_TAT', ascending=False).head(5)
                
                fig_top5 = px.bar(
                    top5_request_tat,
                    x='RequestType',
                    y=['Median_TAT', 'Average_TAT'],
                    title="Top 5 Request Types by TAT",
                    labels={'value': 'TAT (Days)', 'RequestType': 'Request Type'},
                    barmode='group'
                )
                fig_top5.update_xaxes(tickangle=45)
                st.plotly_chart(fig_top5, use_container_width=True)
                
                st.dataframe(top5_request_tat)
        else:
            st.info("No completed cases found for TAT analysis")
    
    # ===================================================================
    # 5. Multi-Hold Completed Cases Analysis
    # ===================================================================
    st.header("5. Multi-Hold Completed Cases Analysis")
    
    if 'CaseType' in df.columns:
        multi_hold_cases = df[df['CaseType'].str.contains('Multi Hold', na=False)].copy()
        multi_hold_completed = multi_hold_cases[multi_hold_cases['TAT_Days'].notna()].copy()
        
        if not multi_hold_completed.empty and 'Month_Str' in multi_hold_completed.columns:
            # 5a. Month-wise number of multi-hold cases
            mh_monthly_count = multi_hold_completed.groupby('Month_Str').size().reset_index(name='Multi_Hold_Count')
            mh_monthly_count = mh_monthly_count.sort_values('Month_Str')
            
            # 5b. % of multi-hold cases
            total_monthly = df.groupby('Month_Str').size().reset_index(name='Total_Cases')
            mh_pct = pd.merge(total_monthly, mh_monthly_count, on='Month_Str', how='left')
            mh_pct['Multi_Hold_Percentage'] = (mh_pct['Multi_Hold_Count'] / mh_pct['Total_Cases'] * 100).round(2)
            mh_pct = mh_pct.fillna(0)
            
            col1, col2 = st.columns(2)
            with col1:
                fig_mh_count = px.bar(
                    mh_monthly_count,
                    x='Month_Str',
                    y='Multi_Hold_Count',
                    title="5a. Month-wise Number of Multi-Hold Completed Cases",
                    labels={'Month_Str': 'Month', 'Multi_Hold_Count': 'Number of Cases'}
                )
                fig_mh_count.update_xaxes(tickangle=45)
                st.plotly_chart(fig_mh_count, use_container_width=True)
            
            with col2:
                fig_mh_pct = px.bar(
                    mh_pct,
                    x='Month_Str',
                    y='Multi_Hold_Percentage',
                    title="5b. % of Multi-Hold Completed Cases by Month",
                    labels={'Month_Str': 'Month', 'Multi_Hold_Percentage': 'Percentage (%)'}
                )
                fig_mh_pct.update_xaxes(tickangle=45)
                st.plotly_chart(fig_mh_pct, use_container_width=True)
            
            st.dataframe(mh_pct[['Month_Str', 'Total_Cases', 'Multi_Hold_Count', 'Multi_Hold_Percentage']])
            
            # 5c. Request Type TAT (Median and Average) for Multi-Hold
            if 'requestTypeDescription' in multi_hold_completed.columns:
                st.subheader("5c. Request Type TAT (Multi-Hold Completed Cases)")
                
                mh_request_tat = multi_hold_completed.groupby('requestTypeDescription').agg({
                    'TAT_Days': ['median', 'mean']
                }).reset_index()
                mh_request_tat.columns = ['RequestType', 'Median_TAT', 'Average_TAT']
                mh_request_tat = mh_request_tat.sort_values('Average_TAT', ascending=False)
                
                col1, col2 = st.columns(2)
                with col1:
                    fig_mh_req_median = px.bar(
                        mh_request_tat,
                        x='RequestType',
                        y='Median_TAT',
                        title="Median TAT by Request Type (Multi-Hold)",
                        labels={'RequestType': 'Request Type', 'Median_TAT': 'Median TAT (Days)'}
                    )
                    fig_mh_req_median.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_mh_req_median, use_container_width=True)
                
                with col2:
                    fig_mh_req_avg = px.bar(
                        mh_request_tat,
                        x='RequestType',
                        y='Average_TAT',
                        title="Average TAT by Request Type (Multi-Hold)",
                        labels={'RequestType': 'Request Type', 'Average_TAT': 'Average TAT (Days)'}
                    )
                    fig_mh_req_avg.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_mh_req_avg, use_container_width=True)
                
                st.dataframe(mh_request_tat)
            
            # 5d. OnHoldReasonDescriptionsHistory TAT (Median and Average) for Multi-Hold
            if 'onHoldReasonDescriptionsHistory' in multi_hold_completed.columns:
                st.subheader("5d. On-Hold Reason TAT (Multi-Hold Completed Cases)")
                
                records = []
                for _, row in multi_hold_completed.iterrows():
                    reasons = parse_separated_values(row.get('onHoldReasonDescriptionsHistory', ''))
                    tat_val = row.get('TAT_Days')
                    if reasons and pd.notna(tat_val):
                        for reason in reasons:
                            records.append({
                                'HoldReason': reason,
                                'TAT_Days': tat_val
                            })
                
                if records:
                    mh_reason_tat_df = pd.DataFrame(records)
                    mh_reason_tat = mh_reason_tat_df.groupby('HoldReason').agg({
                        'TAT_Days': ['median', 'mean', 'count']
                    }).reset_index()
                    mh_reason_tat.columns = ['HoldReason', 'Median_TAT', 'Average_TAT', 'Count']
                    mh_reason_tat = mh_reason_tat.sort_values('Average_TAT', ascending=False)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        fig_mh_reason_median = px.bar(
                            mh_reason_tat.head(10),
                            x='HoldReason',
                            y='Median_TAT',
                            title="Median TAT by Hold Reason (Top 10, Multi-Hold)",
                            labels={'HoldReason': 'Hold Reason', 'Median_TAT': 'Median TAT (Days)'}
                        )
                        fig_mh_reason_median.update_xaxes(tickangle=45)
                        st.plotly_chart(fig_mh_reason_median, use_container_width=True)
                    
                    with col2:
                        fig_mh_reason_avg = px.bar(
                            mh_reason_tat.head(10),
                            x='HoldReason',
                            y='Average_TAT',
                            title="Average TAT by Hold Reason (Top 10, Multi-Hold)",
                            labels={'HoldReason': 'Hold Reason', 'Average_TAT': 'Average TAT (Days)'}
                        )
                        fig_mh_reason_avg.update_xaxes(tickangle=45)
                        st.plotly_chart(fig_mh_reason_avg, use_container_width=True)
                    
                    st.dataframe(mh_reason_tat)
            
            # 5e. Median and Average Holding Time for each Holding Reason Description
            if 'onHoldReasonDescriptionsHistory' in multi_hold_cases.columns and 'HoldingTimes' in multi_hold_cases.columns:
                st.subheader("5e. Holding Time by Hold Reason (Multi-Hold Cases)")
                
                records_ht = []
                for _, row in multi_hold_cases.iterrows():
                    reasons = parse_separated_values(row.get('onHoldReasonDescriptionsHistory', ''))
                    holding_times = row.get('HoldingTimes', [])
                    
                    if reasons and holding_times:
                        for i, reason in enumerate(reasons):
                            if i < len(holding_times):
                                records_ht.append({
                                    'HoldReason': reason,
                                    'HoldingTime': holding_times[i]
                                })
                
                if records_ht:
                    mh_hold_time_df = pd.DataFrame(records_ht)
                    mh_hold_time_stats = mh_hold_time_df.groupby('HoldReason').agg({
                        'HoldingTime': ['median', 'mean', 'count']
                    }).reset_index()
                    mh_hold_time_stats.columns = ['HoldReason', 'Median_HoldingTime', 'Average_HoldingTime', 'Count']
                    mh_hold_time_stats = mh_hold_time_stats.sort_values('Average_HoldingTime', ascending=False)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        fig_ht_median = px.bar(
                            mh_hold_time_stats.head(10),
                            x='HoldReason',
                            y='Median_HoldingTime',
                            title="Median Holding Time by Hold Reason (Top 10)",
                            labels={'HoldReason': 'Hold Reason', 'Median_HoldingTime': 'Median Holding Time (Days)'}
                        )
                        fig_ht_median.update_xaxes(tickangle=45)
                        st.plotly_chart(fig_ht_median, use_container_width=True)
                    
                    with col2:
                        fig_ht_avg = px.bar(
                            mh_hold_time_stats.head(10),
                            x='HoldReason',
                            y='Average_HoldingTime',
                            title="Average Holding Time by Hold Reason (Top 10)",
                            labels={'HoldReason': 'Hold Reason', 'Average_HoldingTime': 'Average Holding Time (Days)'}
                        )
                        fig_ht_avg.update_xaxes(tickangle=45)
                        st.plotly_chart(fig_ht_avg, use_container_width=True)
                    
                    st.dataframe(mh_hold_time_stats)
            
        else:
            st.info("No multi-hold completed cases found")
    
    # ===================================================================
    # 6. Holding Time for Straight Through Cases (Month-wise)
    # ===================================================================
    st.header("6. Holding Time for Straight Through Cases (Month-wise)")

    if 'CaseType' in df.columns and 'Month_Str' in df.columns and 'TotalHoldingTime' in df.columns:
        st_cases = df[df['CaseType'] == 'Straight Through'].copy()

        if not st_cases.empty:
            st_ht_monthly = st_cases.groupby('Month_Str').agg({
                'TotalHoldingTime': ['median', 'mean', 'count']
            }).reset_index()
            st_ht_monthly.columns = ['Month', 'Median_HoldingTime_Days', 'Average_HoldingTime_Days', 'Count']
            st_ht_monthly = st_ht_monthly.sort_values('Month')

            col1, col2 = st.columns(2)
            with col1:
                fig_st_ht_median = px.line(
                    st_ht_monthly,
                    x='Month',
                    y='Median_HoldingTime_Days',
                    title="Median Holding Time (Days) - Straight Through Cases by Month",
                    labels={'Month': 'Month', 'Median_HoldingTime_Days': 'Median Holding Time (Days)'},
                    markers=True
                )
                fig_st_ht_median.update_xaxes(tickangle=45)
                st.plotly_chart(fig_st_ht_median, use_container_width=True)

            with col2:
                fig_st_ht_avg = px.line(
                    st_ht_monthly,
                    x='Month',
                    y='Average_HoldingTime_Days',
                    title="Average Holding Time (Days) - Straight Through Cases by Month",
                    labels={'Month': 'Month', 'Average_HoldingTime_Days': 'Average Holding Time (Days)'},
                    markers=True
                )
                fig_st_ht_avg.update_xaxes(tickangle=45)
                st.plotly_chart(fig_st_ht_avg, use_container_width=True)

            st.markdown("**Note:** Holding time is measured in **days**, based on the difference between "
                        "`onHoldDatesHistory` and `offHoldDatesHistory` for each hold period.")
            st.dataframe(st_ht_monthly)
        else:
            st.info("No Straight Through cases found for holding time analysis.")

    # ===================================================================
    # 7. Number of Locations and Vehicles with Request Type (Month-wise)
    # ===================================================================
    st.header("7. Number of Locations and Vehicles with Request Type (Month-wise)")
    
    if 'numberOfLocations' in df.columns and 'NumberOfVehicles' in df.columns and 'requestTypeDescription' in df.columns and 'Month_Str' in df.columns:
        lv_data = df[
            (df['numberOfLocations'].notna()) & 
            (df['NumberOfVehicles'].notna()) &
            (df['requestTypeDescription'].notna())
        ].copy()
        
        if not lv_data.empty:
            # Locations with Request Type by Month
            loc_request_monthly = lv_data.groupby(['Month_Str', 'numberOfLocations', 'requestTypeDescription']).size().reset_index(name='Count')
            loc_request_monthly = loc_request_monthly.sort_values(['Month_Str', 'Count'], ascending=[True, False])
            
            fig_loc_req = px.bar(
                loc_request_monthly,
                x='Month_Str',
                y='Count',
                color='requestTypeDescription',
                facet_col='numberOfLocations',
                title="Request Type Count by Number of Locations (Month-wise)",
                labels={'Month_Str': 'Month', 'Count': 'Count'},
                barmode='stack'
            )
            fig_loc_req.update_xaxes(tickangle=45)
            st.plotly_chart(fig_loc_req, use_container_width=True)
            
            # Vehicles banding (groups of 4) and count of cases with holding time
            # Create vehicle bands: 1-4, 5-8, 9-12, ...
            lv_data = lv_data.copy()
            lv_data['NumberOfVehicles'] = lv_data['NumberOfVehicles'].astype(int)
            band_start = ((lv_data['NumberOfVehicles'] - 1) // 4) * 4 + 1
            band_end = band_start + 3
            lv_data['VehicleBand'] = band_start.astype(str) + '-' + band_end.astype(str)

            # Only cases with holding time > 0
            hv_band = lv_data[lv_data['TotalHoldingTime'] > 0].copy()
            if not hv_band.empty:
                vehicle_band_counts = hv_band.groupby('VehicleBand').size().reset_index(name='Cases_With_HoldingTime')
                vehicle_band_counts = vehicle_band_counts.sort_values('VehicleBand')

                fig_veh_band = px.bar(
                    vehicle_band_counts,
                    x='VehicleBand',
                    y='Cases_With_HoldingTime',
                    title="Cases with Holding Time by Vehicle Band (4 vehicles per band)",
                    labels={'VehicleBand': 'Number of Vehicles (Band)', 'Cases_With_HoldingTime': 'Number of Cases'}
                )
                st.plotly_chart(fig_veh_band, use_container_width=True)

                st.dataframe(vehicle_band_counts)
            else:
                st.info("No cases with holding time found for vehicle band analysis.")

            st.dataframe(loc_request_monthly)
    
    # ===================================================================
    # 8. On-Hold Reason Seasonality (Month-wise Counts)
    # ===================================================================
    st.header("8. On-Hold Reason Seasonality (Month-wise Counts)")

    if 'onHoldReasonDescriptionsHistory' in df.columns and 'Month_Str' in df.columns:
        records_hr = []
        for _, row in df.iterrows():
            month = row.get('Month_Str')
            if pd.isna(month) or month is None:
                continue
            reasons = parse_separated_values(row.get('onHoldReasonDescriptionsHistory', ''))
            for r in reasons:
                records_hr.append({'Month_Str': month, 'HoldReason': r})

        if records_hr:
            hr_df = pd.DataFrame(records_hr)

            # Overall top 10 hold reasons to keep chart readable
            top_reasons = (
                hr_df['HoldReason']
                .value_counts()
                .head(10)
                .index
                .tolist()
            )
            hr_top = hr_df[hr_df['HoldReason'].isin(top_reasons)]

            # Month-wise counts
            hr_monthly = (
                hr_top.groupby(['Month_Str', 'HoldReason'])
                .size()
                .reset_index(name='Count')
                .sort_values(['Month_Str', 'Count'], ascending=[True, False])
            )

            col1, col2 = st.columns(2)
            with col1:
                fig_hr_counts = px.bar(
                    hr_monthly,
                    x='Month_Str',
                    y='Count',
                    color='HoldReason',
                    title="On-Hold Reasons by Month (Top 10 Overall)",
                    labels={'Month_Str': 'Month', 'Count': 'Count', 'HoldReason': 'Hold Reason'},
                    barmode='stack'
                )
                fig_hr_counts.update_xaxes(tickangle=45)
                st.plotly_chart(fig_hr_counts, use_container_width=True)

            with col2:
                # Percentage within each month
                month_totals = (
                    hr_monthly.groupby('Month_Str')['Count']
                    .sum()
                    .reset_index(name='Month_Total')
                )
                pct_df = hr_monthly.merge(month_totals, on='Month_Str', how='left')
                pct_df['Percentage'] = (pct_df['Count'] / pct_df['Month_Total'] * 100).round(2)

                fig_hr_pct = px.bar(
                    pct_df,
                    x='Month_Str',
                    y='Percentage',
                    color='HoldReason',
                    title="On-Hold Reasons % by Month (Top 10 Overall)",
                    labels={'Month_Str': 'Month', 'Percentage': 'Percentage (%)', 'HoldReason': 'Hold Reason'},
                    barmode='stack'
                )
                fig_hr_pct.update_xaxes(tickangle=45)
                st.plotly_chart(fig_hr_pct, use_container_width=True)

            st.subheader("On-Hold Reasons by Month - Detail")
            st.dataframe(hr_monthly)
        else:
            st.info("No on-hold reasons found to analyze seasonality.")
    else:
        st.warning("⚠️ Required columns not found: onHoldReasonDescriptionsHistory or Month_Str")

    # ===================================================================
    # 9. Number of Touches (from onHoldReasonDescriptionsHistory) by Month
    # ===================================================================
    st.header("9. Number of Touches (from onHoldReasonDescriptionsHistory) by Month")

    if 'Touches_Reasons' in df.columns and 'Month_Str' in df.columns:
        touches_month = (
            df.groupby('Month_Str')['Touches_Reasons']
            .agg(['median', 'mean', 'sum', 'count'])
            .reset_index()
            .rename(columns={
                'Month_Str': 'Month',
                'median': 'Median_Touches',
                'mean': 'Average_Touches',
                'sum': 'Total_Touches',
                'count': 'Case_Count'
            })
            .sort_values('Month')
        )

        col1, col2 = st.columns(2)
        with col1:
            fig_touch_med = px.line(
                touches_month,
                x='Month',
                y='Median_Touches',
                title="Median Number of Touches per Case by Month",
                labels={'Month': 'Month', 'Median_Touches': 'Median Touches'},
                markers=True
            )
            fig_touch_med.update_xaxes(tickangle=45)
            st.plotly_chart(fig_touch_med, use_container_width=True)

        with col2:
            fig_touch_avg = px.line(
                touches_month,
                x='Month',
                y='Average_Touches',
                title="Average Number of Touches per Case by Month",
                labels={'Month': 'Month', 'Average_Touches': 'Average Touches'},
                markers=True
            )
            fig_touch_avg.update_xaxes(tickangle=45)
            st.plotly_chart(fig_touch_avg, use_container_width=True)

        # Distribution of touches per month (0,1,2,3,...)
        touches_dist = (
            df.groupby(['Month_Str', 'Touches_Reasons'])
            .size()
            .reset_index(name='Count')
            .sort_values(['Month_Str', 'Touches_Reasons'])
        )
        fig_touch_dist = px.bar(
            touches_dist,
            x='Month_Str',
            y='Count',
            color='Touches_Reasons',
            title="Distribution of Number of Touches per Month",
            labels={'Month_Str': 'Month', 'Count': 'Case Count', 'Touches_Reasons': 'Number of Touches'},
            barmode='stack'
        )
        fig_touch_dist.update_xaxes(tickangle=45)
        st.plotly_chart(fig_touch_dist, use_container_width=True)

        st.subheader("Number of Touches per Month - Detail")
        st.dataframe(touches_month)
    else:
        st.warning("⚠️ Required columns not found: Touches_Reasons or Month_Str")

    # ===================================================================
    # 10. Prescriptive Analysis: Reasons for Hold Delay
    # ===================================================================
    st.header("10. Prescriptive Analysis: Reasons for Hold Delay")
    
    if 'onHoldReasonDescriptionsHistory' in df.columns and 'HoldingTimes' in df.columns:
        delay_records = []
        for _, row in df.iterrows():
            reasons = parse_separated_values(row.get('onHoldReasonDescriptionsHistory', ''))
            holding_times = row.get('HoldingTimes', [])
            
            if reasons and holding_times:
                for i, reason in enumerate(reasons):
                    if i < len(holding_times):
                        delay_records.append({
                            'HoldReason': reason,
                            'HoldingTime': holding_times[i],
                            'CaseType': row.get('CaseType', 'Unknown')
                        })
        
        if delay_records:
            delay_df = pd.DataFrame(delay_records)
            
            # Calculate statistics by hold reason
            delay_stats = delay_df.groupby('HoldReason').agg({
                'HoldingTime': ['median', 'mean', 'count', 'max']
            }).reset_index()
            delay_stats.columns = ['HoldReason', 'Median_HoldingTime', 'Average_HoldingTime', 'Count', 'Max_HoldingTime']
            delay_stats = delay_stats.sort_values('Average_HoldingTime', ascending=False)
            
            st.subheader("10a. Hold Reasons Causing Longest Delays")
            st.markdown("**Analysis:** Identifying which hold reasons result in the longest holding times")
            
            col1, col2 = st.columns(2)
            with col1:
                fig_delay_median = px.bar(
                    delay_stats.head(15),
                    x='HoldReason',
                    y='Median_HoldingTime',
                    title="Median Holding Time by Hold Reason (Top 15)",
                    labels={'HoldReason': 'Hold Reason', 'Median_HoldingTime': 'Median Holding Time (Days)'},
                    color='Median_HoldingTime',
                    color_continuous_scale='Reds'
                )
                fig_delay_median.update_xaxes(tickangle=45)
                st.plotly_chart(fig_delay_median, use_container_width=True)
            
            with col2:
                fig_delay_avg = px.bar(
                    delay_stats.head(15),
                    x='HoldReason',
                    y='Average_HoldingTime',
                    title="Average Holding Time by Hold Reason (Top 15)",
                    labels={'HoldReason': 'Hold Reason', 'Average_HoldingTime': 'Average Holding Time (Days)'},
                    color='Average_HoldingTime',
                    color_continuous_scale='Oranges'
                )
                fig_delay_avg.update_xaxes(tickangle=45)
                st.plotly_chart(fig_delay_avg, use_container_width=True)
            
            # Show top delay reasons
            st.subheader("10b. Top Hold Reasons by Delay Impact")
            top_delay_reasons = delay_stats.head(10).copy()
            top_delay_reasons['Total_Delay_Days'] = top_delay_reasons['Average_HoldingTime'] * top_delay_reasons['Count']
            top_delay_reasons = top_delay_reasons.sort_values('Total_Delay_Days', ascending=False)
            
            st.dataframe(top_delay_reasons[['HoldReason', 'Count', 'Median_HoldingTime', 'Average_HoldingTime', 'Max_HoldingTime', 'Total_Delay_Days']])
            
            # Recommendations
            st.subheader("10c. Recommendations")
            if not top_delay_reasons.empty:
                top_reason = top_delay_reasons.iloc[0]
                st.info(f"🔴 **Highest Impact:** '{top_reason['HoldReason']}' causes an average delay of {top_reason['Average_HoldingTime']:.1f} days across {top_reason['Count']:.0f} cases, totaling {top_reason['Total_Delay_Days']:.1f} delay days.")
                
                if len(top_delay_reasons) > 1:
                    st.warning(f"⚠️ **Top 3 Delay Reasons:** {', '.join(top_delay_reasons.head(3)['HoldReason'].tolist())}")
        else:
            st.info("No hold delay data available for analysis")
    else:
        st.warning("⚠️ Required columns not found: onHoldReasonDescriptionsHistory or HoldingTimes")
    
    # ===================================================================
    # 11. Hold Seasonality Analysis
    # ===================================================================
    st.header("11. Hold Seasonality Analysis")
    
    if 'Month_Str' in df.columns and 'TotalHoldingTime' in df.columns:
        # Cases with holds by month
        hold_seasonality = df[df['TotalHoldingTime'] > 0].copy()
        
        if not hold_seasonality.empty:
            monthly_holds = hold_seasonality.groupby('Month_Str').agg({
                'TotalHoldingTime': ['median', 'mean', 'sum', 'count']
            }).reset_index()
            monthly_holds.columns = ['Month', 'Median_HoldingTime', 'Average_HoldingTime', 'Total_HoldingTime', 'Cases_With_Holds']
            
            total_cases_monthly = df.groupby('Month_Str').size().reset_index(name='Total_Cases')
            monthly_holds = pd.merge(monthly_holds, total_cases_monthly, left_on='Month', right_on='Month_Str', how='left')
            monthly_holds['Hold_Percentage'] = (monthly_holds['Cases_With_Holds'] / monthly_holds['Total_Cases'] * 100).round(2)
            monthly_holds = monthly_holds.sort_values('Month')
            
            st.subheader("11a. Hold Frequency by Month")
            col1, col2 = st.columns(2)
            with col1:
                fig_hold_freq = px.bar(
                    monthly_holds,
                    x='Month',
                    y='Cases_With_Holds',
                    title="Number of Cases with Holds by Month",
                    labels={'Month': 'Month', 'Cases_With_Holds': 'Cases with Holds'},
                    color='Cases_With_Holds',
                    color_continuous_scale='Blues'
                )
                fig_hold_freq.update_xaxes(tickangle=45)
                st.plotly_chart(fig_hold_freq, use_container_width=True)
            
            with col2:
                fig_hold_pct = px.line(
                    monthly_holds,
                    x='Month',
                    y='Hold_Percentage',
                    title="Percentage of Cases with Holds by Month",
                    labels={'Month': 'Month', 'Hold_Percentage': 'Hold Percentage (%)'},
                    markers=True
                )
                fig_hold_pct.update_xaxes(tickangle=45)
                st.plotly_chart(fig_hold_pct, use_container_width=True)
            
            st.subheader("11b. Hold Duration Seasonality")
            col1, col2 = st.columns(2)
            with col1:
                fig_hold_median = px.line(
                    monthly_holds,
                    x='Month',
                    y='Median_HoldingTime',
                    title="Median Holding Time by Month",
                    labels={'Month': 'Month', 'Median_HoldingTime': 'Median Holding Time (Days)'},
                    markers=True
                )
                fig_hold_median.update_xaxes(tickangle=45)
                st.plotly_chart(fig_hold_median, use_container_width=True)
            
            with col2:
                fig_hold_avg = px.line(
                    monthly_holds,
                    x='Month',
                    y='Average_HoldingTime',
                    title="Average Holding Time by Month",
                    labels={'Month': 'Month', 'Average_HoldingTime': 'Average Holding Time (Days)'},
                    markers=True
                )
                fig_hold_avg.update_xaxes(tickangle=45)
                st.plotly_chart(fig_hold_avg, use_container_width=True)
            
            st.dataframe(monthly_holds[['Month', 'Total_Cases', 'Cases_With_Holds', 'Hold_Percentage', 'Median_HoldingTime', 'Average_HoldingTime']])
        else:
            st.info("No cases with holds found for seasonality analysis")
    else:
        st.warning("⚠️ Required columns not found: Month_Str or TotalHoldingTime")
    
    # ===================================================================
    # 12. Hold Buckets Analysis
    # ===================================================================
    st.header("12. Hold Buckets Analysis")
    
    if 'TotalHoldingTime' in df.columns:
        # Create hold buckets similar to TAT buckets
        def categorize_hold_time(hold_time):
            if pd.isna(hold_time) or hold_time == 0:
                return "No Hold"
            elif hold_time <= 5:
                return "0-5 days"
            elif hold_time <= 7:
                return "5-7 days"
            else:
                return "7+ days"
        
        df['HoldBucket'] = df['TotalHoldingTime'].apply(categorize_hold_time)
        
        hold_bucket_counts = df['HoldBucket'].value_counts().reset_index()
        hold_bucket_counts.columns = ['HoldBucket', 'Count']
        hold_bucket_counts['Percentage'] = (hold_bucket_counts['Count'] / len(df) * 100).round(2)
        
        st.subheader("12a. Distribution of Cases by Hold Buckets")
        col1, col2 = st.columns(2)
        with col1:
            fig_hold_bucket_bar = px.bar(
                hold_bucket_counts,
                x='HoldBucket',
                y='Count',
                title="Case Count by Hold Bucket",
                labels={'HoldBucket': 'Hold Bucket', 'Count': 'Number of Cases'},
                color='Count',
                color_continuous_scale='Purples'
            )
            st.plotly_chart(fig_hold_bucket_bar, use_container_width=True)
        
        with col2:
            fig_hold_bucket_pie = px.pie(
                hold_bucket_counts,
                values='Count',
                names='HoldBucket',
                title="Hold Bucket Distribution (%)"
            )
            st.plotly_chart(fig_hold_bucket_pie, use_container_width=True)
        
        st.subheader("12b. Hold Buckets by Month")
        if 'Month_Str' in df.columns:
            hold_bucket_monthly = df.groupby(['Month_Str', 'HoldBucket']).size().reset_index(name='Count')
            hold_bucket_monthly = hold_bucket_monthly.sort_values(['Month_Str', 'HoldBucket'])
            
            fig_hold_bucket_monthly = px.bar(
                hold_bucket_monthly,
                x='Month_Str',
                y='Count',
                color='HoldBucket',
                title="Hold Bucket Distribution by Month",
                labels={'Month_Str': 'Month', 'Count': 'Case Count'},
                barmode='stack',
                color_discrete_map={
                    "No Hold": "#90EE90",
                    "0-5 days": "#FFD700",
                    "5-7 days": "#FFA500",
                    "7+ days": "#FF6347"
                }
            )
            fig_hold_bucket_monthly.update_xaxes(tickangle=45)
            st.plotly_chart(fig_hold_bucket_monthly, use_container_width=True)
            
            st.dataframe(hold_bucket_monthly.pivot(index='Month_Str', columns='HoldBucket', values='Count').fillna(0))
        
        st.subheader("12c. Hold Bucket Statistics")
        hold_bucket_stats = df.groupby('HoldBucket').agg({
            'TotalHoldingTime': ['median', 'mean', 'min', 'max'],
            'TAT_Days': ['median', 'mean'],
            'CaseType': 'count'
        }).reset_index()
        hold_bucket_stats.columns = ['HoldBucket', 'Median_HoldTime', 'Avg_HoldTime', 'Min_HoldTime', 'Max_HoldTime', 
                                     'Median_TAT', 'Avg_TAT', 'Case_Count']
        st.dataframe(hold_bucket_stats)
    else:
        st.warning("⚠️ Required column not found: TotalHoldingTime")
    
    # ===================================================================
    # 13. TAT Buckets Analysis (0-5, 5-7, 7+ days)
    # ===================================================================
    st.header("13. TAT Buckets Analysis")
    
    if 'TAT_Days' in df.columns:
        # Create TAT buckets
        def categorize_tat(tat):
            if pd.isna(tat):
                return "Not Completed"
            elif tat <= 5:
                return "0-5 days"
            elif tat <= 7:
                return "5-7 days"
            else:
                return "7+ days"
        
        df['TAT_Bucket'] = df['TAT_Days'].apply(categorize_tat)
        
        tat_bucket_counts = df['TAT_Bucket'].value_counts().reset_index()
        tat_bucket_counts.columns = ['TAT_Bucket', 'Count']
        tat_bucket_counts['Percentage'] = (tat_bucket_counts['Count'] / len(df) * 100).round(2)
        
        st.subheader("13a. Distribution of Cases by TAT Buckets")
        col1, col2 = st.columns(2)
        with col1:
            fig_tat_bucket_bar = px.bar(
                tat_bucket_counts,
                x='TAT_Bucket',
                y='Count',
                title="Case Count by TAT Bucket",
                labels={'TAT_Bucket': 'TAT Bucket', 'Count': 'Number of Cases'},
                color='Count',
                color_continuous_scale='Greens'
            )
            st.plotly_chart(fig_tat_bucket_bar, use_container_width=True)
        
        with col2:
            fig_tat_bucket_pie = px.pie(
                tat_bucket_counts,
                values='Count',
                names='TAT_Bucket',
                title="TAT Bucket Distribution (%)"
            )
            st.plotly_chart(fig_tat_bucket_pie, use_container_width=True)
        
        st.subheader("13b. TAT Buckets by Month")
        if 'Month_Str' in df.columns:
            tat_bucket_monthly = df.groupby(['Month_Str', 'TAT_Bucket']).size().reset_index(name='Count')
            tat_bucket_monthly = tat_bucket_monthly.sort_values(['Month_Str', 'TAT_Bucket'])
            
            fig_tat_bucket_monthly = px.bar(
                tat_bucket_monthly,
                x='Month_Str',
                y='Count',
                color='TAT_Bucket',
                title="TAT Bucket Distribution by Month",
                labels={'Month_Str': 'Month', 'Count': 'Case Count'},
                barmode='stack',
                color_discrete_map={
                    "Not Completed": "#D3D3D3",
                    "0-5 days": "#90EE90",
                    "5-7 days": "#FFD700",
                    "7+ days": "#FF6347"
                }
            )
            fig_tat_bucket_monthly.update_xaxes(tickangle=45)
            st.plotly_chart(fig_tat_bucket_monthly, use_container_width=True)
            
            st.dataframe(tat_bucket_monthly.pivot(index='Month_Str', columns='TAT_Bucket', values='Count').fillna(0))
        
        st.subheader("13c. TAT Bucket Statistics")
        tat_bucket_stats = df.groupby('TAT_Bucket').agg({
            'TAT_Days': ['median', 'mean', 'min', 'max'],
            'TotalHoldingTime': ['median', 'mean'],
            'CaseType': 'count'
        }).reset_index()
        tat_bucket_stats.columns = ['TAT_Bucket', 'Median_TAT', 'Avg_TAT', 'Min_TAT', 'Max_TAT', 
                                    'Median_HoldTime', 'Avg_HoldTime', 'Case_Count']
        st.dataframe(tat_bucket_stats)
        
        st.subheader("13d. TAT Buckets vs Hold Buckets")
        if 'HoldBucket' in df.columns:
            tat_hold_cross = pd.crosstab(df['TAT_Bucket'], df['HoldBucket'], margins=True)
            st.dataframe(tat_hold_cross)
            
            fig_tat_hold_heatmap = px.density_heatmap(
                df[df['TAT_Bucket'] != 'Not Completed'],
                x='HoldBucket',
                y='TAT_Bucket',
                title="TAT Buckets vs Hold Buckets Heatmap",
                labels={'HoldBucket': 'Hold Bucket', 'TAT_Bucket': 'TAT Bucket'}
            )
            st.plotly_chart(fig_tat_hold_heatmap, use_container_width=True)
    else:
        st.warning("⚠️ Required column not found: TAT_Days")
    
    # ===================================================================
    # 14. Write-Out Reason Description Seasonality (Month-wise)
    # ===================================================================
    st.header("14. Write-Out Reason Description Seasonality (Month-wise)")
    
    writeout_col = None
    if 'writeOutReasonDescriptionsHistory' in df.columns:
        writeout_col = 'writeOutReasonDescriptionsHistory'
    elif 'writeOutReasonDescription' in df.columns:
        writeout_col = 'writeOutReasonDescription'
    
    if writeout_col and 'Month_Str' in df.columns:
        records_wo = []
        for _, row in df.iterrows():
            reasons = parse_separated_values(row.get(writeout_col, ''))
            month = row.get('Month_Str')
            if reasons and pd.notna(month):
                for reason in reasons:
                    records_wo.append({
                        'Month': month,
                        'WriteOutReason': reason
                    })
        
        if records_wo:
            wo_df = pd.DataFrame(records_wo)
            wo_monthly = wo_df.groupby(['Month', 'WriteOutReason']).size().reset_index(name='Count')
            wo_monthly = wo_monthly.sort_values(['Month', 'Count'], ascending=[True, False])
            
            fig_wo = px.bar(
                wo_monthly,
                x='Month',
                y='Count',
                color='WriteOutReason',
                title="Write-Out Reason Distribution by Month",
                labels={'Month': 'Month', 'Count': 'Count'},
                barmode='stack'
            )
            fig_wo.update_xaxes(tickangle=45)
            st.plotly_chart(fig_wo, use_container_width=True)
            
            st.dataframe(wo_monthly)
        else:
            st.info("No write-out reasons found")
    else:
        st.warning("⚠️ Required columns not found: writeOutReasonDescriptionsHistory/writeOutReasonDescription or Month_Str")


if __name__ == "__main__":
    main()

