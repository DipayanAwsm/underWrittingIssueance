"""
Advanced Underwriting Visualization Dashboard

This Streamlit app provides advanced visualizations:
1. Case Type Classification based on touches from onHoldDatesHistory
2. Request Type TAT seasonality (median & average) with status
3. Multi-hold cases seasonality analysis
4. Vehicle count vs TAT analysis
5. Location count vs TAT analysis
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
    page_title="Advanced Underwriting Analytics",
    page_icon="📈",
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
    """Parse values that may be separated by ', ' or '|'"""
    if pd.isna(value) or value == '':
        return []
    
    value = str(value).strip()
    if separator:
        separators = [separator]
    else:
        if '|' in value:
            separators = ['|']
        elif ', ' in value:
            separators = [', ']
        else:
            return [value] if value else []
    
    result = []
    for sep in separators:
        if sep in value:
            result = [v.strip() for v in value.split(sep) if v.strip()]
            break
    
    return result if result else ([value] if value else [])


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


def main():
    st.markdown('<h1 class="main-header">📈 Advanced Underwriting Analytics Dashboard</h1>', unsafe_allow_html=True)
    
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
        
        # 1. Case Type Classification based on onHoldReasonDescriptionsHistory
        if 'onHoldReasonDescriptionsHistory' in df.columns:
            case_type_results = df['onHoldReasonDescriptionsHistory'].apply(classify_case_type_by_reasons)
            df['CaseType'] = [r[0] for r in case_type_results]
            df['NumberOfTouches'] = [r[1] for r in case_type_results]
        else:
            st.warning("⚠️ Column 'onHoldReasonDescriptionsHistory' not found. Case type classification skipped.")
            df['CaseType'] = 'Unknown'
            df['NumberOfTouches'] = 0
        
        # Extract month for seasonality
        if 'createDateTime' in df.columns:
            df['createDateTime'] = pd.to_datetime(df['createDateTime'], errors='coerce')
            df['Month'] = df['createDateTime'].dt.to_period('M')
            df['Month_Str'] = df['Month'].astype(str)
        else:
            df['Month_Str'] = None
    
    # ===================================================================
    # Case Type Filter
    # ===================================================================
    st.sidebar.header("🔍 Filter Options")
    
    if 'CaseType' in df.columns:
        original_count = len(df)
        
        # Filter options
        filter_option = st.sidebar.radio(
            "Filter by Case Type:",
            ["All", "Straight Through", "Multi Hold"],
            index=0  # Default to "All"
        )
        
        if filter_option == "Straight Through":
            # Filter to only Straight Through cases
            df_filtered = df[df['CaseType'] == 'Straight Through'].copy()
            filtered_count = len(df_filtered)
            excluded_count = original_count - filtered_count
            st.sidebar.success(f"✅ Showing: {filtered_count:,} Straight Through cases")
            st.sidebar.info(f"📊 Excluded: {excluded_count:,} cases")
            df = df_filtered
        elif filter_option == "Multi Hold":
            # Filter to Multi Hold cases (includes One Touch and Multi Hold)
            df_filtered = df[
                (df['CaseType'] == 'One Touch') | 
                (df['CaseType'].str.contains('Multi Hold', na=False))
            ].copy()
            filtered_count = len(df_filtered)
            excluded_count = original_count - filtered_count
            st.sidebar.success(f"✅ Showing: {filtered_count:,} cases (One Touch + Multi Hold)")
            st.sidebar.info(f"📊 Excluded: {excluded_count:,} cases")
            df = df_filtered
        else:
            # "All" option - no filtering
            st.sidebar.info(f"📊 Showing all {original_count:,} cases")
    else:
        st.sidebar.warning("⚠️ CaseType column not available for filtering")
    
    # ===================================================================
    # 0. Overall Summary - Total Cases and Percent-wise Analysis
    # ===================================================================
    st.header("0. Overall Summary - Total Cases and Percent-wise Analysis")
    
    total_cases = len(df)
    completed_cases = df['completedDateTime'].notna().sum() if 'completedDateTime' in df.columns else 0
    straight_through_count = len(df[df['CaseType'] == 'Straight Through']) if 'CaseType' in df.columns else 0
    one_touch_count = len(df[df['CaseType'] == 'One Touch']) if 'CaseType' in df.columns else 0
    multi_hold_count = len(df[df['CaseType'].str.contains('Multi Hold', na=False)]) if 'CaseType' in df.columns else 0
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Cases", f"{total_cases:,}")
    with col2:
        completion_pct = (completed_cases / total_cases * 100) if total_cases > 0 else 0
        st.metric("Completed Cases", f"{completed_cases:,}", f"{completion_pct:.1f}%")
    with col3:
        st_pct = (straight_through_count / total_cases * 100) if total_cases > 0 else 0
        st.metric("Straight Through", f"{straight_through_count:,}", f"{st_pct:.1f}%")
    with col4:
        ot_pct = (one_touch_count / total_cases * 100) if total_cases > 0 else 0
        st.metric("One Touch", f"{one_touch_count:,}", f"{ot_pct:.1f}%")
    with col5:
        mh_pct = (multi_hold_count / total_cases * 100) if total_cases > 0 else 0
        st.metric("Multi Hold", f"{multi_hold_count:,}", f"{mh_pct:.1f}%")
    
    # Percent-wise breakdown chart
    if 'CaseType' in df.columns:
        case_type_counts = df['CaseType'].value_counts()
        case_type_pct = (case_type_counts / total_cases * 100).round(2)
        
        summary_df = pd.DataFrame({
            'Case Type': case_type_counts.index,
            'Count': case_type_counts.values,
            'Percentage': case_type_pct.values
        })
        
        col1, col2 = st.columns(2)
        with col1:
            fig_summary_bar = px.bar(
                summary_df,
                x='Case Type',
                y='Percentage',
                title="Case Type Distribution (%)",
                labels={'Percentage': 'Percentage (%)', 'Case Type': 'Case Type'},
                text='Percentage'
            )
            fig_summary_bar.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
            fig_summary_bar.update_xaxes(tickangle=45)
            st.plotly_chart(fig_summary_bar, use_container_width=True)
        
        with col2:
            fig_summary_pie = px.pie(
                summary_df,
                values='Count',
                names='Case Type',
                title="Case Type Distribution (Count & %)",
                hole=0.4
            )
            fig_summary_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_summary_pie, use_container_width=True)
        
        st.dataframe(summary_df)
    
    # ===================================================================
    # 1. Case Type Classification (based on onHoldReasonDescriptionsHistory)
    # ===================================================================
    st.header("1. Case Type Classification (Based on onHoldReasonDescriptionsHistory)")
    
    if 'CaseType' in df.columns:
        case_type_counts = df['CaseType'].value_counts()
        
        col1, col2 = st.columns(2)
        with col1:
            fig_case_type = px.pie(
                values=case_type_counts.values,
                names=case_type_counts.index,
                title="Case Type Distribution (by Touches from onHoldDatesHistory)"
            )
            st.plotly_chart(fig_case_type, use_container_width=True)
        
        with col2:
            fig_case_type_bar = px.bar(
                x=case_type_counts.index,
                y=case_type_counts.values,
                title="Case Type Count (Bar Chart)",
                labels={'x': 'Case Type', 'y': 'Count'}
            )
            fig_case_type_bar.update_xaxes(tickangle=45)
            st.plotly_chart(fig_case_type_bar, use_container_width=True)
        
        st.dataframe(case_type_counts.reset_index().rename(columns={'index': 'Case Type', 'CaseType': 'Count'}))
    
    # ===================================================================
    # 2. Request Type TAT Seasonality (Median & Average) with Status
    # ===================================================================
    st.header("2. Request Type TAT Seasonality (Median & Average) with Status")
    
    if 'requestTypeDescription' in df.columns and 'Month_Str' in df.columns and 'TAT_Days' in df.columns:
        # Filter completed cases
        tat_data = df[df['TAT_Days'].notna()].copy()
        
        if not tat_data.empty:
            # Group by requestTypeDescription and Month_Str
            request_month_tat = tat_data.groupby(['requestTypeDescription', 'Month_Str']).agg({
                'TAT_Days': ['median', 'mean', 'count']
            }).reset_index()
            request_month_tat.columns = ['RequestType', 'Month', 'Median_TAT', 'Average_TAT', 'Count']
            
            # Median TAT
            fig_median = px.line(
                request_month_tat,
                x='Month',
                y='Median_TAT',
                color='RequestType',
                title="Median TAT by Request Type (Month-wise)",
                labels={'Median_TAT': 'Median TAT (Days)', 'Month': 'Month'}
            )
            fig_median.update_xaxes(tickangle=45)
            st.plotly_chart(fig_median, use_container_width=True)
            
            # Average TAT
            fig_avg = px.line(
                request_month_tat,
                x='Month',
                y='Average_TAT',
                color='RequestType',
                title="Average TAT by Request Type (Month-wise)",
                labels={'Average_TAT': 'Average TAT (Days)', 'Month': 'Month'}
            )
            fig_avg.update_xaxes(tickangle=45)
            st.plotly_chart(fig_avg, use_container_width=True)
            
            # Status Description distribution by Request Type
            if 'statusDescription' in df.columns:
                status_request = df.groupby(['requestTypeDescription', 'statusDescription']).size().reset_index(name='Count')
                
                fig_status_req = px.bar(
                    status_request,
                    x='requestTypeDescription',
                    y='Count',
                    color='statusDescription',
                    title="Status Description Distribution by Request Type",
                    labels={'requestTypeDescription': 'Request Type', 'Count': 'Count'},
                    barmode='stack'
                )
                fig_status_req.update_xaxes(tickangle=45)
                st.plotly_chart(fig_status_req, use_container_width=True)
            
            st.dataframe(request_month_tat)
        else:
            st.info("No completed cases found for TAT analysis")
    else:
        st.warning("⚠️ Required columns not found: requestTypeDescription, Month_Str, or TAT_Days")
    
    # ===================================================================
    # 3. Multi-Hold Cases Seasonality Analysis
    # ===================================================================
    st.header("3. Multi-Hold Cases Seasonality Analysis")
    
    if 'CaseType' in df.columns and 'Month_Str' in df.columns:
        multi_hold_cases = df[df['CaseType'].str.contains('Multi Hold', na=False)].copy()
        
        if not multi_hold_cases.empty:
            # 3a. Month-wise seasonality % of cases
            total_by_month = df.groupby('Month_Str').size().reset_index(name='Total_Cases')
            multi_by_month = multi_hold_cases.groupby('Month_Str').size().reset_index(name='Multi_Hold_Cases')
            
            monthly_pct = pd.merge(total_by_month, multi_by_month, on='Month_Str', how='left')
            monthly_pct['Multi_Hold_Percentage'] = (monthly_pct['Multi_Hold_Cases'] / monthly_pct['Total_Cases'] * 100).round(2)
            monthly_pct = monthly_pct.fillna(0)
            
            col1, col2 = st.columns(2)
            with col1:
                fig_multi_pct = px.bar(
                    monthly_pct,
                    x='Month_Str',
                    y='Multi_Hold_Percentage',
                    title="Multi-Hold Cases % by Month",
                    labels={'Month_Str': 'Month', 'Multi_Hold_Percentage': 'Percentage (%)'}
                )
                fig_multi_pct.update_xaxes(tickangle=45)
                st.plotly_chart(fig_multi_pct, use_container_width=True)
            
            with col2:
                fig_multi_count = px.line(
                    monthly_pct,
                    x='Month_Str',
                    y='Multi_Hold_Cases',
                    title="Multi-Hold Cases Count by Month",
                    labels={'Month_Str': 'Month', 'Multi_Hold_Cases': 'Number of Cases'}
                )
                fig_multi_count.update_xaxes(tickangle=45)
                st.plotly_chart(fig_multi_count, use_container_width=True)
            
            st.dataframe(monthly_pct)
            
            # 3b. Month-wise Request Type TAT (Average & Median) for Multi-Hold cases
            if 'requestTypeDescription' in multi_hold_cases.columns and 'TAT_Days' in multi_hold_cases.columns:
                multi_tat_data = multi_hold_cases[multi_hold_cases['TAT_Days'].notna()].copy()
                
                if not multi_tat_data.empty:
                    multi_request_month = multi_tat_data.groupby(['requestTypeDescription', 'Month_Str']).agg({
                        'TAT_Days': ['median', 'mean']
                    }).reset_index()
                    multi_request_month.columns = ['RequestType', 'Month', 'Median_TAT', 'Average_TAT']
                    
                    st.subheader("3b. Request Type TAT Seasonality (Multi-Hold Cases Only)")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        fig_multi_median = px.line(
                            multi_request_month,
                            x='Month',
                            y='Median_TAT',
                            color='RequestType',
                            title="Median TAT by Request Type - Multi-Hold Cases",
                            labels={'Median_TAT': 'Median TAT (Days)', 'Month': 'Month'}
                        )
                        fig_multi_median.update_xaxes(tickangle=45)
                        st.plotly_chart(fig_multi_median, use_container_width=True)
                    
                    with col2:
                        fig_multi_avg = px.line(
                            multi_request_month,
                            x='Month',
                            y='Average_TAT',
                            color='RequestType',
                            title="Average TAT by Request Type - Multi-Hold Cases",
                            labels={'Average_TAT': 'Average TAT (Days)', 'Month': 'Month'}
                        )
                        fig_multi_avg.update_xaxes(tickangle=45)
                        st.plotly_chart(fig_multi_avg, use_container_width=True)
                    
                    st.dataframe(multi_request_month)
                else:
                    st.info("No completed multi-hold cases found for TAT analysis")
            
            # ===================================================================
            # 3c. Top 5 Analysis for Multi-Hold Cases
            # ===================================================================
            st.subheader("3c. Top 5 Analysis for Multi-Hold Cases")
            
            # 1. Top 5 Request Types
            if 'requestTypeDescription' in multi_hold_cases.columns:
                top5_request_types = multi_hold_cases['requestTypeDescription'].value_counts().head(5)
                top5_request_pct = (top5_request_types / len(multi_hold_cases) * 100).round(2)
                
                top5_request_df = pd.DataFrame({
                    'Request Type': top5_request_types.index,
                    'Count': top5_request_types.values,
                    'Percentage': top5_request_pct.values
                })
                
                col1, col2 = st.columns(2)
                with col1:
                    fig_top5_req = px.bar(
                        top5_request_df,
                        x='Request Type',
                        y='Count',
                        title="Top 5 Request Types (Multi-Hold Cases)",
                        labels={'Count': 'Count', 'Request Type': 'Request Type'},
                        text='Count'
                    )
                    fig_top5_req.update_traces(textposition='outside')
                    fig_top5_req.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_top5_req, use_container_width=True)
                
                with col2:
                    fig_top5_req_pct = px.bar(
                        top5_request_df,
                        x='Request Type',
                        y='Percentage',
                        title="Top 5 Request Types % (Multi-Hold Cases)",
                        labels={'Percentage': 'Percentage (%)', 'Request Type': 'Request Type'},
                        text='Percentage'
                    )
                    fig_top5_req_pct.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                    fig_top5_req_pct.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_top5_req_pct, use_container_width=True)
                
                st.dataframe(top5_request_df)
            
            # 2. Top 5 Hold Reasons (from onHoldReasonDescriptionsHistory)
            if 'onHoldReasonDescriptionsHistory' in multi_hold_cases.columns:
                all_hold_reasons = []
                for reasons in multi_hold_cases['onHoldReasonDescriptionsHistory'].apply(parse_separated_values):
                    all_hold_reasons.extend(reasons)
                
                if all_hold_reasons:
                    hold_reason_counts = pd.Series(all_hold_reasons).value_counts().head(5)
                    hold_reason_pct = (hold_reason_counts / len(all_hold_reasons) * 100).round(2)
                    
                    top5_hold_df = pd.DataFrame({
                        'Hold Reason': hold_reason_counts.index,
                        'Count': hold_reason_counts.values,
                        'Percentage': hold_reason_pct.values
                    })
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        fig_top5_hold = px.bar(
                            top5_hold_df,
                            x='Hold Reason',
                            y='Count',
                            title="Top 5 Hold Reasons (Multi-Hold Cases)",
                            labels={'Count': 'Count', 'Hold Reason': 'Hold Reason'},
                            text='Count'
                        )
                        fig_top5_hold.update_traces(textposition='outside')
                        fig_top5_hold.update_xaxes(tickangle=45)
                        st.plotly_chart(fig_top5_hold, use_container_width=True)
                    
                    with col2:
                        fig_top5_hold_pct = px.bar(
                            top5_hold_df,
                            x='Hold Reason',
                            y='Percentage',
                            title="Top 5 Hold Reasons % (Multi-Hold Cases)",
                            labels={'Percentage': 'Percentage (%)', 'Hold Reason': 'Hold Reason'},
                            text='Percentage'
                        )
                        fig_top5_hold_pct.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                        fig_top5_hold_pct.update_xaxes(tickangle=45)
                        st.plotly_chart(fig_top5_hold_pct, use_container_width=True)
                    
                    st.dataframe(top5_hold_df)
                else:
                    st.info("No hold reasons found in multi-hold cases")
            elif 'onHoldReasonDescription' in multi_hold_cases.columns:
                # Fallback to single column if history not available
                top5_hold_reasons = multi_hold_cases['onHoldReasonDescription'].value_counts().head(5)
                top5_hold_pct = (top5_hold_reasons / len(multi_hold_cases) * 100).round(2)
                
                top5_hold_df = pd.DataFrame({
                    'Hold Reason': top5_hold_reasons.index,
                    'Count': top5_hold_reasons.values,
                    'Percentage': top5_hold_pct.values
                })
                
                col1, col2 = st.columns(2)
                with col1:
                    fig_top5_hold = px.bar(
                        top5_hold_df,
                        x='Hold Reason',
                        y='Count',
                        title="Top 5 Hold Reasons (Multi-Hold Cases)",
                        labels={'Count': 'Count', 'Hold Reason': 'Hold Reason'},
                        text='Count'
                    )
                    fig_top5_hold.update_traces(textposition='outside')
                    fig_top5_hold.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_top5_hold, use_container_width=True)
                
                with col2:
                    fig_top5_hold_pct = px.bar(
                        top5_hold_df,
                        x='Hold Reason',
                        y='Percentage',
                        title="Top 5 Hold Reasons % (Multi-Hold Cases)",
                        labels={'Percentage': 'Percentage (%)', 'Hold Reason': 'Hold Reason'},
                        text='Percentage'
                    )
                    fig_top5_hold_pct.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                    fig_top5_hold_pct.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_top5_hold_pct, use_container_width=True)
                
                st.dataframe(top5_hold_df)
            
            # 3. Top 5 BGI Descriptions
            if 'bgiDescription' in multi_hold_cases.columns:
                top5_bgi = multi_hold_cases['bgiDescription'].value_counts().head(5)
                top5_bgi_pct = (top5_bgi / len(multi_hold_cases) * 100).round(2)
                
                top5_bgi_df = pd.DataFrame({
                    'BGI Description': top5_bgi.index,
                    'Count': top5_bgi.values,
                    'Percentage': top5_bgi_pct.values
                })
                
                col1, col2 = st.columns(2)
                with col1:
                    fig_top5_bgi = px.bar(
                        top5_bgi_df,
                        x='BGI Description',
                        y='Count',
                        title="Top 5 BGI Descriptions (Multi-Hold Cases)",
                        labels={'Count': 'Count', 'BGI Description': 'BGI Description'},
                        text='Count'
                    )
                    fig_top5_bgi.update_traces(textposition='outside')
                    fig_top5_bgi.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_top5_bgi, use_container_width=True)
                
                with col2:
                    fig_top5_bgi_pct = px.bar(
                        top5_bgi_df,
                        x='BGI Description',
                        y='Percentage',
                        title="Top 5 BGI Descriptions % (Multi-Hold Cases)",
                        labels={'Percentage': 'Percentage (%)', 'BGI Description': 'BGI Description'},
                        text='Percentage'
                    )
                    fig_top5_bgi_pct.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                    fig_top5_bgi_pct.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_top5_bgi_pct, use_container_width=True)
                
                st.dataframe(top5_bgi_df)
            else:
                st.warning("⚠️ Column 'bgiDescription' not found in the dataset")
        else:
            st.info("No multi-hold cases found in the dataset")
    else:
        st.warning("⚠️ Required columns not found: CaseType or Month_Str")
    
    # ===================================================================
    # 4. On-Hold Reason Seasonality (Month-wise)
    # ===================================================================
    st.header("4. On-Hold Reason Seasonality (Month-wise)")
    
    if 'onHoldReasonDescriptionsHistory' in df.columns and 'Month_Str' in df.columns:
        # Flatten month + hold reasons
        records = []
        for _, row in df.iterrows():
            month = row.get('Month_Str')
            if pd.isna(month) or month is None:
                continue
            reasons = parse_separated_values(row.get('onHoldReasonDescriptionsHistory', ''))
            for r in reasons:
                records.append({'Month_Str': month, 'HoldReason': r})
        
        if records:
            reasons_df = pd.DataFrame(records)
            
            # Overall top 10 hold reasons to keep charts readable
            top_overall = (
                reasons_df['HoldReason']
                .value_counts()
                .head(10)
                .index
                .tolist()
            )
            reasons_top = reasons_df[reasons_df['HoldReason'].isin(top_overall)]
            
            # Month-wise counts
            reason_month_counts = (
                reasons_top
                .groupby(['Month_Str', 'HoldReason'])
                .size()
                .reset_index(name='Count')
            )
            
            col1, col2 = st.columns(2)
            with col1:
                fig_hold_month = px.bar(
                    reason_month_counts,
                    x='Month_Str',
                    y='Count',
                    color='HoldReason',
                    title="On-Hold Reasons by Month (Top 10 Overall)",
                    labels={'Month_Str': 'Month', 'Count': 'Number of Holds', 'HoldReason': 'Hold Reason'},
                    barmode='stack'
                )
                fig_hold_month.update_xaxes(tickangle=45)
                st.plotly_chart(fig_hold_month, use_container_width=True)
            
            with col2:
                # Percentage within each month
                month_totals = (
                    reason_month_counts.groupby('Month_Str')['Count'].sum().reset_index(name='Month_Total')
                )
                pct_df = reason_month_counts.merge(month_totals, on='Month_Str', how='left')
                pct_df['Percentage'] = (pct_df['Count'] / pct_df['Month_Total'] * 100).round(2)
                
                fig_hold_pct = px.bar(
                    pct_df,
                    x='Month_Str',
                    y='Percentage',
                    color='HoldReason',
                    title="On-Hold Reasons % by Month (Top 10 Overall)",
                    labels={'Month_Str': 'Month', 'Percentage': 'Percentage (%)', 'HoldReason': 'Hold Reason'},
                    barmode='stack'
                )
                fig_hold_pct.update_xaxes(tickangle=45)
                st.plotly_chart(fig_hold_pct, use_container_width=True)
            
            st.subheader("On-Hold Reasons by Month - Detail")
            st.dataframe(reason_month_counts.sort_values(['Month_Str', 'Count'], ascending=[True, False]))
        else:
            st.info("No hold reasons found to analyze seasonality.")
    else:
        st.warning("⚠️ Required columns not found: onHoldReasonDescriptionsHistory or Month_Str")
    
    # ===================================================================
    # 5. Number of Vehicles vs TAT (Average & Median)
    # ===================================================================
    st.header("5. Number of Vehicles vs TAT (Average & Median)")
    
    if 'NumberOfVehicles' in df.columns and 'TAT_Days' in df.columns:
        vehicle_tat_data = df[df['TAT_Days'].notna() & df['NumberOfVehicles'].notna()].copy()
        
        if not vehicle_tat_data.empty:
            vehicle_tat_stats = vehicle_tat_data.groupby('NumberOfVehicles').agg({
                'TAT_Days': ['median', 'mean', 'count']
            }).reset_index()
            vehicle_tat_stats.columns = ['NumberOfVehicles', 'Median_TAT', 'Average_TAT', 'Count']
            
            col1, col2 = st.columns(2)
            with col1:
                fig_veh_median = px.bar(
                    vehicle_tat_stats,
                    x='NumberOfVehicles',
                    y='Median_TAT',
                    title="Median TAT by Number of Vehicles",
                    labels={'NumberOfVehicles': 'Number of Vehicles', 'Median_TAT': 'Median TAT (Days)'}
                )
                st.plotly_chart(fig_veh_median, use_container_width=True)
            
            with col2:
                fig_veh_avg = px.bar(
                    vehicle_tat_stats,
                    x='NumberOfVehicles',
                    y='Average_TAT',
                    title="Average TAT by Number of Vehicles",
                    labels={'NumberOfVehicles': 'Number of Vehicles', 'Average_TAT': 'Average TAT (Days)'}
                )
                st.plotly_chart(fig_veh_avg, use_container_width=True)
            
            # Combined chart
            fig_veh_combined = go.Figure()
            fig_veh_combined.add_trace(go.Scatter(
                x=vehicle_tat_stats['NumberOfVehicles'],
                y=vehicle_tat_stats['Median_TAT'],
                mode='lines+markers',
                name='Median TAT',
                line=dict(color='blue', width=2)
            ))
            fig_veh_combined.add_trace(go.Scatter(
                x=vehicle_tat_stats['NumberOfVehicles'],
                y=vehicle_tat_stats['Average_TAT'],
                mode='lines+markers',
                name='Average TAT',
                line=dict(color='red', width=2)
            ))
            fig_veh_combined.update_layout(
                title="Median vs Average TAT by Number of Vehicles",
                xaxis_title="Number of Vehicles",
                yaxis_title="TAT (Days)",
                hovermode='x unified'
            )
            st.plotly_chart(fig_veh_combined, use_container_width=True)
            
            st.dataframe(vehicle_tat_stats)
        else:
            st.info("No data available for vehicle TAT analysis")
    else:
        st.warning("⚠️ Required columns not found: NumberOfVehicles or TAT_Days")
    
    # ===================================================================
    # 6. Number of Locations vs TAT (Median & Average)
    # ===================================================================
    st.header("6. Number of Locations vs TAT (Median & Average)")
    
    if 'numberOfLocations' in df.columns and 'TAT_Days' in df.columns:
        location_tat_data = df[df['TAT_Days'].notna() & df['numberOfLocations'].notna()].copy()
        
        if not location_tat_data.empty:
            location_tat_stats = location_tat_data.groupby('numberOfLocations').agg({
                'TAT_Days': ['median', 'mean', 'count']
            }).reset_index()
            location_tat_stats.columns = ['numberOfLocations', 'Median_TAT', 'Average_TAT', 'Count']
            
            col1, col2 = st.columns(2)
            with col1:
                fig_loc_median = px.bar(
                    location_tat_stats,
                    x='numberOfLocations',
                    y='Median_TAT',
                    title="Median TAT by Number of Locations",
                    labels={'numberOfLocations': 'Number of Locations', 'Median_TAT': 'Median TAT (Days)'}
                )
                st.plotly_chart(fig_loc_median, use_container_width=True)
            
            with col2:
                fig_loc_avg = px.bar(
                    location_tat_stats,
                    x='numberOfLocations',
                    y='Average_TAT',
                    title="Average TAT by Number of Locations",
                    labels={'numberOfLocations': 'Number of Locations', 'Average_TAT': 'Average TAT (Days)'}
                )
                st.plotly_chart(fig_loc_avg, use_container_width=True)
            
            # Combined chart
            fig_loc_combined = go.Figure()
            fig_loc_combined.add_trace(go.Scatter(
                x=location_tat_stats['numberOfLocations'],
                y=location_tat_stats['Median_TAT'],
                mode='lines+markers',
                name='Median TAT',
                line=dict(color='green', width=2)
            ))
            fig_loc_combined.add_trace(go.Scatter(
                x=location_tat_stats['numberOfLocations'],
                y=location_tat_stats['Average_TAT'],
                mode='lines+markers',
                name='Average TAT',
                line=dict(color='orange', width=2)
            ))
            fig_loc_combined.update_layout(
                title="Median vs Average TAT by Number of Locations",
                xaxis_title="Number of Locations",
                yaxis_title="TAT (Days)",
                hovermode='x unified'
            )
            st.plotly_chart(fig_loc_combined, use_container_width=True)
            
            st.dataframe(location_tat_stats)
        else:
            st.info("No data available for location TAT analysis")
    else:
        st.warning("⚠️ Required columns not found: numberOfLocations or TAT_Days")

    # ===================================================================
    # 7. Average TAT by Request Type and Write-Out Reason
    # ===================================================================
    st.header("7. Average TAT by Request Type and Write-Out Reason")

    # Prefer history column if available, else fall back to single description
    writeout_col = None
    if 'writeOutReasonDescriptionsHistory' in df.columns:
        writeout_col = 'writeOutReasonDescriptionsHistory'
    elif 'writeOutReasonDescription' in df.columns:
        writeout_col = 'writeOutReasonDescription'

    if writeout_col and 'requestTypeDescription' in df.columns and 'TAT_Days' in df.columns:
        tat_wo_data = df[df['TAT_Days'].notna()].copy()

        if not tat_wo_data.empty:
            records = []
            for _, row in tat_wo_data.iterrows():
                req_type = row.get('requestTypeDescription')
                tat_val = row.get('TAT_Days')
                if pd.isna(tat_val) or pd.isna(req_type):
                    continue

                reasons = parse_separated_values(row.get(writeout_col, ''))
                # If no reasons parsed but a single value exists, we still treat it
                if not reasons and pd.notna(row.get(writeout_col)):
                    reasons = [str(row.get(writeout_col))]

                for r in reasons:
                    records.append({
                        'RequestType': req_type,
                        'WriteOutReason': r,
                        'TAT_Days': tat_val
                    })

            if records:
                tat_wo_df = pd.DataFrame(records)

                # Compute average and median TAT by (RequestType, WriteOutReason)
                tat_wo_stats = tat_wo_df.groupby(['RequestType', 'WriteOutReason']).agg({
                    'TAT_Days': ['mean', 'median', 'count']
                }).reset_index()
                tat_wo_stats.columns = ['RequestType', 'WriteOutReason', 'Average_TAT', 'Median_TAT', 'Count']

                st.subheader("7a. Detailed Table")
                st.dataframe(tat_wo_stats.sort_values(['RequestType', 'Average_TAT'], ascending=[True, False]))

                # Limit to top 5 write-out reasons overall for clearer charts
                top_reasons_overall = (
                    tat_wo_df['WriteOutReason'].value_counts().head(5).index.tolist()
                )
                tat_wo_top = tat_wo_stats[tat_wo_stats['WriteOutReason'].isin(top_reasons_overall)]

                st.subheader("7b. Top 5 Write-Out Reasons Analysis")
                
                # Show top 5 write-out reasons summary
                top5_summary = tat_wo_df[tat_wo_df['WriteOutReason'].isin(top_reasons_overall)].groupby('WriteOutReason').agg({
                    'TAT_Days': ['mean', 'median', 'count']
                }).reset_index()
                top5_summary.columns = ['WriteOutReason', 'Average_TAT', 'Median_TAT', 'Count']
                top5_summary = top5_summary.sort_values('Count', ascending=False)
                
                st.dataframe(top5_summary)

                col1, col2 = st.columns(2)
                with col1:
                    fig_avg_tat = px.bar(
                        tat_wo_top,
                        x='RequestType',
                        y='Average_TAT',
                        color='WriteOutReason',
                        title="Average TAT by Request Type and Write-Out Reason (Top 5 Reasons)",
                        labels={
                            'RequestType': 'Request Type',
                            'Average_TAT': 'Average TAT (Days)',
                            'WriteOutReason': 'Write-Out Reason'
                        },
                        barmode='group'
                    )
                    fig_avg_tat.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_avg_tat, use_container_width=True)

                with col2:
                    fig_median_tat = px.bar(
                        tat_wo_top,
                        x='RequestType',
                        y='Median_TAT',
                        color='WriteOutReason',
                        title="Median TAT by Request Type and Write-Out Reason (Top 5 Reasons)",
                        labels={
                            'RequestType': 'Request Type',
                            'Median_TAT': 'Median TAT (Days)',
                            'WriteOutReason': 'Write-Out Reason'
                        },
                        barmode='group'
                    )
                    fig_median_tat.update_xaxes(tickangle=45)
                    st.plotly_chart(fig_median_tat, use_container_width=True)
                
                # Also show filtered detailed table for top 5
                st.subheader("7c. Detailed Table (Top 5 Write-Out Reasons)")
                st.dataframe(tat_wo_top.sort_values(['RequestType', 'Average_TAT'], ascending=[True, False]))
            else:
                st.info("No write-out reasons found to analyze TAT by request type.")
        else:
            st.info("No completed cases (with TAT) found for write-out reason analysis.")
    else:
        st.warning("⚠️ Required columns not found: writeOutReasonDescriptionsHistory/writeOutReasonDescription, requestTypeDescription, or TAT_Days")


if __name__ == "__main__":
    main()

