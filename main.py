import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import calendar
import os

# Set page configuration
st.set_page_config(
    page_title="Utilization Dashboard",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS to improve aesthetics
st.markdown("""
<style>
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 60px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4e8df5;
        color: white;
    }
    div[data-testid="stMetricValue"] {
        font-size: 28px;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 16px;
    }
    .css-1v0mbdj {
        margin-top: -60px;
    }
    div[data-testid="stSidebarNav"] li div a {
        margin-left: 1rem;
        padding: 1rem;
        width: 300px;
        border-radius: 0.5rem;
    }
    div[data-testid="stSidebarNav"] li div::focus-within {
        background-color: rgba(151, 166, 195, 0.15);
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    """Load and prepare the data"""
    # Load the TIME ENTRIES sheet
    time_entries_df = pd.read_excel('Utilization.xlsx', sheet_name='TIME ENTRIES')
    
    # Load the ATTORNEYS sheet
    attorneys_df = pd.read_excel('Utilization.xlsx', sheet_name='ATTORNEYS')
    
    # Load the CLIENTS sheet
    clients_df = pd.read_excel('Utilization.xlsx', sheet_name='CLIENTS')
    
    # Clean and prepare the data
    
    # Convert date columns to datetime
    time_entries_df['Date'] = pd.to_datetime(time_entries_df['Date'])
    
    # Extract month, year components for filtering
    time_entries_df['Month'] = time_entries_df['Date'].dt.month
    time_entries_df['Year'] = time_entries_df['Date'].dt.year
    time_entries_df['MonthName'] = time_entries_df['Date'].dt.strftime('%b')
    time_entries_df['MonthYear'] = time_entries_df['Date'].dt.strftime('%b %Y')
    
    # Remove "$" and convert to numeric
    if 'Billable ($)' in time_entries_df.columns:
        time_entries_df['Billable ($)'] = pd.to_numeric(time_entries_df['Billable ($)'].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce')
    
    if 'Rate ($)' in time_entries_df.columns:
        time_entries_df['Rate ($)'] = pd.to_numeric(time_entries_df['Rate ($)'].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce')
    
    # Create fee type column
    time_entries_df['FeeType'] = time_entries_df['Type'].apply(
        lambda x: 'Fixed Fee' if 'FixedFee' in str(x) else ('Time' if 'TimeEntry' in str(x) else 'Other')
    )
    
    # Clean attorneys data
    if '🎚️ Target Hours / Month' in attorneys_df.columns:
        attorneys_df['Target Hours'] = attorneys_df['🎚️ Target Hours / Month']
    
    return time_entries_df, attorneys_df, clients_df

def filter_data(df, year_filter, month_filter, rev_band_filter, attorney_filter, pg_filter, fee_type_filter):
    """Apply filters to the dataframe"""
    filtered_df = df.copy()
    
    # Apply year filter
    if year_filter != "All":
        filtered_df = filtered_df[filtered_df['Year'] == int(year_filter)]
    
    # Apply month filter
    if month_filter != "All":
        month_num = list(calendar.month_abbr).index(month_filter)
        filtered_df = filtered_df[filtered_df['Month'] == month_num]
    
    # Apply revenue band filter
    if rev_band_filter != "All":
        filtered_df = filtered_df[filtered_df['CLIENT ANNUAL REV'] == rev_band_filter]
    
    # Apply attorney filter
    if attorney_filter != "All":
        filtered_df = filtered_df[filtered_df['Associated Attorney'] == attorney_filter]
    
    # Apply practice group filter
    if pg_filter != "All":
        filtered_df = filtered_df[filtered_df['PG1'] == pg_filter]
    
    # Apply fee type filter
    if fee_type_filter != "All":
        filtered_df = filtered_df[filtered_df['FeeType'] == fee_type_filter]
    
    return filtered_df

def format_number(num, prefix=""):
    """Format numbers with comma separators and optional prefix"""
    if isinstance(num, (int, float)):
        return f"{prefix}{num:,.0f}"
    return "N/A"

def format_currency(num):
    """Format numbers as currency"""
    if isinstance(num, (int, float)):
        return f"${num:,.2f}"
    return "N/A"

def calculate_metrics(filtered_df):
    """Calculate key metrics from filtered data"""
    # Calculate total billable hours
    total_billable_hours = filtered_df['Quantity / Hours'].sum()
    
    # Calculate total fees
    total_fees = filtered_df['Billable ($)'].sum()
    
    # Calculate average rate
    avg_rate = total_fees / total_billable_hours if total_billable_hours > 0 else 0
    
    # Calculate monthly bills generated
    monthly_bills_generated = filtered_df[filtered_df['Bill Number'].notna()]['Bill Number'].nunique()
    
    return {
        'total_billable_hours': total_billable_hours,
        'total_fees': total_fees,
        'avg_rate': avg_rate,
        'monthly_bills_generated': monthly_bills_generated
    }

def create_overview_section(filtered_df, time_entries_df, attorneys_df):
    """Create the overview section with key metrics and visualizations"""
    metrics = calculate_metrics(filtered_df)
    
    # Create metrics row
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Billable Hours", format_number(metrics['total_billable_hours']))
    with col2:
        st.metric("Monthly Bills Generated", format_number(metrics['monthly_bills_generated']))
    with col3:
        st.metric("Average Rate", format_currency(metrics['avg_rate']))
    with col4:
        st.metric("Total Fees", format_currency(metrics['total_fees']))
    
    # Create two column layout for charts
    col1, col2 = st.columns(2)
    
    with col1:
        # Monthly billable hours trend
        monthly_hours = filtered_df.groupby('MonthYear')['Quantity / Hours'].sum().reset_index()
        monthly_hours['MonthYear'] = pd.Categorical(monthly_hours['MonthYear'], 
                                                   categories=sorted(filtered_df['MonthYear'].unique(), 
                                                                    key=lambda x: datetime.strptime(x, '%b %Y')),
                                                   ordered=True)
        monthly_hours = monthly_hours.sort_values('MonthYear')
        
        fig = px.bar(monthly_hours, x='MonthYear', y='Quantity / Hours',
                     title='Monthly Billable Hours',
                     labels={'MonthYear': 'Month', 'Quantity / Hours': 'Hours'},
                     color_discrete_sequence=['#4e8df5'])
        fig.update_layout(xaxis_title="Month", yaxis_title="Hours", height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Utilization vs Target
        if not attorneys_df.empty:
            # Get attorneys in filtered data
            active_attorneys = filtered_df['Associated Attorney'].unique()
            
            # Filter attorneys dataframe
            attorney_hours = filtered_df.groupby('Associated Attorney')['Quantity / Hours'].sum().reset_index()
            
            # Merge with attorney targets
            attorney_targets = attorneys_df[attorneys_df['Attorney Name'].isin(active_attorneys)]
            attorney_util = pd.merge(attorney_hours, attorney_targets, 
                                     left_on='Associated Attorney', 
                                     right_on='Attorney Name',
                                     how='left')
            
            # Calculate utilization percentage
            attorney_util['Utilization %'] = attorney_util['Quantity / Hours'] / attorney_util['Target Hours'] * 100
            attorney_util = attorney_util.sort_values('Utilization %', ascending=False).head(10)
            
            fig = px.bar(attorney_util, x='Associated Attorney', y='Utilization %',
                         title='Attorney Utilization vs Target (Top 10)',
                         labels={'Associated Attorney': 'Attorney', 'Utilization %': 'Utilization %'},
                         color_discrete_sequence=['#4CAF50'])
            
            # Add reference line at 100%
            fig.add_shape(
                type="line",
                x0=-0.5,
                y0=100,
                x1=len(attorney_util)-0.5,
                y1=100,
                line=dict(color="red", width=2, dash="dash"),
            )
            
            fig.update_layout(xaxis_title="Attorney", yaxis_title="Utilization %", height=350,
                             xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)

def create_client_analysis(filtered_df):
    """Create client analysis section"""
    st.subheader("Client Analysis")
    
    # Fees by client
    client_fees = filtered_df.groupby('Client')['Billable ($)'].sum().reset_index()
    client_fees = client_fees.sort_values('Billable ($)', ascending=False).head(10)
    
    fig = px.bar(client_fees, x='Client', y='Billable ($)',
                 title='Top 10 Clients by Fees',
                 labels={'Client': 'Client', 'Billable ($)': 'Fees ($)'},
                 color_discrete_sequence=['#4e8df5'])
    fig.update_layout(xaxis_title="Client", yaxis_title="Fees ($)", height=400,
                     xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        # Hours by client
        client_hours = filtered_df.groupby('Client')['Quantity / Hours'].sum().reset_index()
        client_hours = client_hours.sort_values('Quantity / Hours', ascending=False).head(10)
        
        fig = px.bar(client_hours, x='Client', y='Quantity / Hours',
                     title='Top 10 Clients by Hours',
                     labels={'Client': 'Client', 'Quantity / Hours': 'Hours'},
                     color_discrete_sequence=['#4CAF50'])
        fig.update_layout(xaxis_title="Client", yaxis_title="Hours", height=350,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Fees by industry
        industry_fees = filtered_df.groupby('CLIENT INDUSTRY')['Billable ($)'].sum().reset_index()
        industry_fees = industry_fees.sort_values('Billable ($)', ascending=False)
        
        fig = px.pie(industry_fees, values='Billable ($)', names='CLIENT INDUSTRY',
                    title='Fees by Industry',
                    color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_layout(height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    # Detailed client metrics
    st.subheader("Detailed Client Metrics")
    client_metrics = filtered_df.groupby('Client').agg({
        'Billable ($)': 'sum',
        'Quantity / Hours': 'sum',
        'Bill Number': pd.Series.nunique
    }).reset_index()
    
    client_metrics.columns = ['Client', 'Total Fees', 'Total Hours', 'Number of Bills']
    client_metrics['Average Rate'] = client_metrics['Total Fees'] / client_metrics['Total Hours']
    client_metrics['Average Bill Amount'] = client_metrics['Total Fees'] / client_metrics['Number of Bills'].replace(0, np.nan)
    
    client_metrics = client_metrics.sort_values('Total Fees', ascending=False)
    
    # Format columns
    client_metrics['Total Fees'] = client_metrics['Total Fees'].apply(lambda x: f"${x:,.2f}")
    client_metrics['Average Rate'] = client_metrics['Average Rate'].apply(lambda x: f"${x:,.2f}" if not pd.isna(x) else "N/A")
    client_metrics['Average Bill Amount'] = client_metrics['Average Bill Amount'].apply(lambda x: f"${x:,.2f}" if not pd.isna(x) else "N/A")
    
    # Hide index
    st.dataframe(client_metrics, hide_index=True, use_container_width=True)

def create_revenue_bands(filtered_df):
    """Create revenue bands analysis section"""
    st.subheader("Fee Bands Analysis")
    
    # Fees by revenue band
    rev_band_fees = filtered_df.groupby('CLIENT ANNUAL REV')['Billable ($)'].sum().reset_index()
    rev_band_fees = rev_band_fees.sort_values('Billable ($)', ascending=False)
    
    # Define a sorting order for revenue bands
    sorting_order = {
        '< $10M': 0,
        '$10M - $30M': 1,
        '$30M – $100M': 2, 
        '$100M – $500M': 3,
        '$500M – $1B': 4,
        '$1B – $3B': 5,
        '$3B – $5B': 6,
        '$5B – $10B': 7,
        '> $10 billion': 8,
        'Confidential': 9
    }
    
    rev_band_fees['sort_order'] = rev_band_fees['CLIENT ANNUAL REV'].map(sorting_order)
    rev_band_fees = rev_band_fees.sort_values('sort_order').drop('sort_order', axis=1)
    
    fig = px.bar(rev_band_fees, x='CLIENT ANNUAL REV', y='Billable ($)',
                 title='Fees by Client Annual Revenue Band',
                 labels={'CLIENT ANNUAL REV': 'Annual Revenue Band', 'Billable ($)': 'Fees ($)'},
                 color_discrete_sequence=['#4e8df5'])
    fig.update_layout(xaxis_title="Annual Revenue Band", yaxis_title="Fees ($)", height=400,
                     xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        # Hours by revenue band
        rev_band_hours = filtered_df.groupby('CLIENT ANNUAL REV')['Quantity / Hours'].sum().reset_index()
        rev_band_hours['sort_order'] = rev_band_hours['CLIENT ANNUAL REV'].map(sorting_order)
        rev_band_hours = rev_band_hours.sort_values('sort_order').drop('sort_order', axis=1)
        
        fig = px.bar(rev_band_hours, x='CLIENT ANNUAL REV', y='Quantity / Hours',
                     title='Hours by Client Annual Revenue Band',
                     labels={'CLIENT ANNUAL REV': 'Annual Revenue Band', 'Quantity / Hours': 'Hours'},
                     color_discrete_sequence=['#4CAF50'])
        fig.update_layout(xaxis_title="Annual Revenue Band", yaxis_title="Hours", height=350,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Client count by revenue band
        rev_band_clients = filtered_df.groupby('CLIENT ANNUAL REV')['Client'].nunique().reset_index()
        rev_band_clients['sort_order'] = rev_band_clients['CLIENT ANNUAL REV'].map(sorting_order)
        rev_band_clients = rev_band_clients.sort_values('sort_order').drop('sort_order', axis=1)
        
        fig = px.bar(rev_band_clients, x='CLIENT ANNUAL REV', y='Client',
                     title='Number of Clients by Annual Revenue Band',
                     labels={'CLIENT ANNUAL REV': 'Annual Revenue Band', 'Client': 'Number of Clients'},
                     color_discrete_sequence=['#FF9800'])
        fig.update_layout(xaxis_title="Annual Revenue Band", yaxis_title="Number of Clients", height=350,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

def create_client_segmentation(filtered_df):
    """Create client segmentation section"""
    st.subheader("Client Segmentation")
    
    # Fees by client stage
    stage_fees = filtered_df.groupby('CLIENT STAGE')['Billable ($)'].sum().reset_index()
    
    # Sort by fees
    stage_fees = stage_fees.sort_values('Billable ($)', ascending=False)
    
    fig = px.bar(stage_fees, x='CLIENT STAGE', y='Billable ($)',
                 title='Fees by Client Stage',
                 labels={'CLIENT STAGE': 'Client Stage', 'Billable ($)': 'Fees ($)'},
                 color_discrete_sequence=['#4e8df5'])
    fig.update_layout(xaxis_title="Client Stage", yaxis_title="Fees ($)", height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        # Fees by practice area
        pa_fees = filtered_df.groupby('PG1')['Billable ($)'].sum().reset_index()
        pa_fees = pa_fees.sort_values('Billable ($)', ascending=False)
        
        fig = px.bar(pa_fees, x='PG1', y='Billable ($)',
                     title='Fees by Practice Area',
                     labels={'PG1': 'Practice Area', 'Billable ($)': 'Fees ($)'},
                     color_discrete_sequence=['#4CAF50'])
        fig.update_layout(xaxis_title="Practice Area", yaxis_title="Fees ($)", height=350,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Average fees per client by stage
        stage_avg_fees = filtered_df.groupby(['CLIENT STAGE', 'Client'])['Billable ($)'].sum().reset_index()
        stage_avg_fees = stage_avg_fees.groupby('CLIENT STAGE').agg({
            'Billable ($)': 'mean',
            'Client': 'count'
        }).reset_index()
        stage_avg_fees.columns = ['CLIENT STAGE', 'Avg Fees per Client', 'Number of Clients']
        
        fig = px.scatter(stage_avg_fees, x='Avg Fees per Client', y='Number of Clients', 
                          size='Avg Fees per Client', color='CLIENT STAGE',
                          title='Average Fees per Client vs Number of Clients by Stage',
                          labels={'Avg Fees per Client': 'Average Fees per Client ($)', 
                                 'Number of Clients': 'Number of Clients',
                                 'CLIENT STAGE': 'Client Stage'})
        fig.update_layout(height=350)
        st.plotly_chart(fig, use_container_width=True)

def create_attorney_analysis(filtered_df, attorneys_df):
    """Create attorney analysis section"""
    st.subheader("Attorney Analysis")
    
    # Hours vs target by attorney
    if not attorneys_df.empty:
        # Get active attorneys
        active_attorneys = filtered_df['Associated Attorney'].unique()
        
        # Get hours by attorney
        attorney_hours = filtered_df.groupby('Associated Attorney')['Quantity / Hours'].sum().reset_index()
        
        # Merge with attorney targets
        attorney_targets = attorneys_df[attorneys_df['Attorney Name'].isin(active_attorneys)]
        attorney_util = pd.merge(attorney_hours, attorney_targets, 
                                 left_on='Associated Attorney', 
                                 right_on='Attorney Name',
                                 how='left')
        
        # Calculate utilization percentage
        attorney_util['Utilization %'] = attorney_util['Quantity / Hours'] / attorney_util['Target Hours'] * 100
        attorney_util = attorney_util.sort_values('Utilization %', ascending=False)
        
        # Filter out attorneys with 0 or 1 hour
        attorney_util = attorney_util[attorney_util['Quantity / Hours'] > 1]
        
        fig = px.bar(attorney_util, x='Associated Attorney', y='Utilization %',
                     title='Attorney Utilization vs Target',
                     labels={'Associated Attorney': 'Attorney', 'Utilization %': 'Utilization %'},
                     color_discrete_sequence=['#4CAF50'])
        
        # Add reference line at 100%
        fig.add_shape(
            type="line",
            x0=-0.5,
            y0=100,
            x1=len(attorney_util)-0.5,
            y1=100,
            line=dict(color="red", width=2, dash="dash"),
        )
        
        fig.update_layout(xaxis_title="Attorney", yaxis_title="Utilization %", height=400,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        # Top attorneys by hours
        attorney_hours = filtered_df.groupby('Associated Attorney')['Quantity / Hours'].sum().reset_index()
        attorney_hours = attorney_hours.sort_values('Quantity / Hours', ascending=False).head(10)
        
        fig = px.bar(attorney_hours, x='Associated Attorney', y='Quantity / Hours',
                     title='Top 10 Attorneys by Hours',
                     labels={'Associated Attorney': 'Attorney', 'Quantity / Hours': 'Hours'},
                     color_discrete_sequence=['#4e8df5'])
        fig.update_layout(xaxis_title="Attorney", yaxis_title="Hours", height=350,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Top attorneys by fees
        attorney_fees = filtered_df.groupby('Associated Attorney')['Billable ($)'].sum().reset_index()
        attorney_fees = attorney_fees.sort_values('Billable ($)', ascending=False).head(10)
        
        fig = px.bar(attorney_fees, x='Associated Attorney', y='Billable ($)',
                     title='Top 10 Attorneys by Fees',
                     labels={'Associated Attorney': 'Attorney', 'Billable ($)': 'Fees ($)'},
                     color_discrete_sequence=['#FF9800'])
        fig.update_layout(xaxis_title="Attorney", yaxis_title="Fees ($)", height=350,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    # Attorney utilization table
    st.subheader("Attorney Utilization Details")
    if not attorneys_df.empty:
        # Prepare attorney utilization table
        attorney_detail = attorney_util.copy()
        attorney_detail = attorney_detail[['Associated Attorney', 'Quantity / Hours', 'Target Hours', 'Utilization %', 'Practice Area (Primary)']]
        attorney_detail.columns = ['Attorney', 'Hours', 'Target Hours', 'Utilization %', 'Primary Practice Area']
        
        # Format columns
        attorney_detail['Utilization %'] = attorney_detail['Utilization %'].apply(lambda x: f"{x:.1f}%")
        
        # Sort by utilization
        attorney_detail = attorney_detail.sort_values('Hours', ascending=False)
        
        # Hide index
        st.dataframe(attorney_detail, hide_index=True, use_container_width=True)
    
    # Practice area distribution
    st.subheader("Practice Area Analysis")
    
    # Hours by practice area
    practice_hours = filtered_df.groupby('PG1')['Quantity / Hours'].sum().reset_index()
    practice_hours = practice_hours.sort_values('Quantity / Hours', ascending=False)
    
    fig = px.bar(practice_hours, x='PG1', y='Quantity / Hours',
                 title='Hours by Practice Area',
                 labels={'PG1': 'Practice Area', 'Quantity / Hours': 'Hours'},
                 color_discrete_sequence=['#4e8df5'])
    fig.update_layout(xaxis_title="Practice Area", yaxis_title="Hours", height=400,
                     xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)
    
    # Fees distribution across practice areas and revenue bands
    practice_fees_band = filtered_df.groupby(['PG1', 'CLIENT ANNUAL REV'])['Billable ($)'].sum().reset_index()
    
    fig = px.sunburst(practice_fees_band, 
                     path=['PG1', 'CLIENT ANNUAL REV'], 
                     values='Billable ($)',
                     title='Fees Distribution by Practice Area and Revenue Band',
                     color_discrete_sequence=px.colors.qualitative.Pastel)
    fig.update_layout(height=600)
    st.plotly_chart(fig, use_container_width=True)

def create_fee_trends(filtered_df):
    """Create fee trends section"""
    st.subheader("Fee Trends Analysis")
    
    # Monthly fees trend
    monthly_fees = filtered_df.groupby(['Year', 'MonthName', 'MonthYear'])['Billable ($)'].sum().reset_index()
    
    # Sort by date
    monthly_fees['MonthYear'] = pd.Categorical(monthly_fees['MonthYear'], 
                                               categories=sorted(filtered_df['MonthYear'].unique(), 
                                                                key=lambda x: datetime.strptime(x, '%b %Y')),
                                               ordered=True)
    monthly_fees = monthly_fees.sort_values('MonthYear')
    
    fig = px.line(monthly_fees, x='MonthYear', y='Billable ($)',
                 title='Monthly Fees Trend',
                 labels={'MonthYear': 'Month', 'Billable ($)': 'Fees ($)'},
                 markers=True,
                 color_discrete_sequence=['#4e8df5'])
    fig.update_layout(xaxis_title="Month", yaxis_title="Fees ($)", height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        # Monthly hours by fee type
        monthly_hours_type = filtered_df.groupby(['MonthYear', 'FeeType'])['Quantity / Hours'].sum().reset_index()
        
        # Sort by date
        monthly_hours_type['MonthYear'] = pd.Categorical(monthly_hours_type['MonthYear'], 
                                                   categories=sorted(filtered_df['MonthYear'].unique(), 
                                                                    key=lambda x: datetime.strptime(x, '%b %Y')),
                                                   ordered=True)
        monthly_hours_type = monthly_hours_type.sort_values('MonthYear')
        
        fig = px.bar(monthly_hours_type, x='MonthYear', y='Quantity / Hours', color='FeeType',
                     title='Monthly Hours by Fee Type',
                     labels={'MonthYear': 'Month', 'Quantity / Hours': 'Hours', 'FeeType': 'Fee Type'},
                     barmode='stack')
        fig.update_layout(xaxis_title="Month", yaxis_title="Hours", height=350,
                         legend_title="Fee Type", xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Average rate trend
        monthly_rate = filtered_df.groupby('MonthYear').agg({
            'Billable ($)': 'sum',
            'Quantity / Hours': 'sum'
        }).reset_index()
        
        monthly_rate['Average Rate'] = monthly_rate['Billable ($)'] / monthly_rate['Quantity / Hours']
        
        # Sort by date
        monthly_rate['MonthYear'] = pd.Categorical(monthly_rate['MonthYear'], 
                                               categories=sorted(filtered_df['MonthYear'].unique(), 
                                                                key=lambda x: datetime.strptime(x, '%b %Y')),
                                               ordered=True)
        monthly_rate = monthly_rate.sort_values('MonthYear')
        
        fig = px.line(monthly_rate, x='MonthYear', y='Average Rate',
                     title='Monthly Average Rate Trend',
                     labels={'MonthYear': 'Month', 'Average Rate': 'Average Rate ($)'},
                     markers=True,
                     color_discrete_sequence=['#FF9800'])
        fig.update_layout(xaxis_title="Month", yaxis_title="Average Rate ($)", height=350,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    # Monthly hours and fees by revenue band
    st.subheader("Monthly Trends by Revenue Band")
    
    # Monthly hours by revenue band
    monthly_hours_band = filtered_df.groupby(['MonthYear', 'CLIENT ANNUAL REV'])['Quantity / Hours'].sum().reset_index()
    
    # Sort by date
    monthly_hours_band['MonthYear'] = pd.Categorical(monthly_hours_band['MonthYear'], 
                                               categories=sorted(filtered_df['MonthYear'].unique(), 
                                                                key=lambda x: datetime.strptime(x, '%b %Y')),
                                               ordered=True)
    monthly_hours_band = monthly_hours_band.sort_values('MonthYear')
    
    fig = px.bar(monthly_hours_band, x='MonthYear', y='Quantity / Hours', color='CLIENT ANNUAL REV',
                 title='Monthly Hours by Revenue Band',
                 labels={'MonthYear': 'Month', 'Quantity / Hours': 'Hours', 'CLIENT ANNUAL REV': 'Revenue Band'},
                 barmode='stack')
    fig.update_layout(xaxis_title="Month", yaxis_title="Hours", height=400,
                     legend_title="Revenue Band", xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)
    
    # Monthly fees by revenue band
    monthly_fees_band = filtered_df.groupby(['MonthYear', 'CLIENT ANNUAL REV'])['Billable ($)'].sum().reset_index()
    
    # Sort by date
    monthly_fees_band['MonthYear'] = pd.Categorical(monthly_fees_band['MonthYear'], 
                                                  categories=sorted(filtered_df['MonthYear'].unique(), 
                                                                  key=lambda x: datetime.strptime(x, '%b %Y')),
                                                  ordered=True)
    monthly_fees_band = monthly_fees_band.sort_values('MonthYear')
    
    fig = px.bar(monthly_fees_band, x='MonthYear', y='Billable ($)', color='CLIENT ANNUAL REV',
                 title='Monthly Fees by Revenue Band',
                 labels={'MonthYear': 'Month', 'Billable ($)': 'Fees ($)', 'CLIENT ANNUAL REV': 'Revenue Band'},
                 barmode='stack')
    fig.update_layout(xaxis_title="Month", yaxis_title="Fees ($)", height=400,
                     legend_title="Revenue Band", xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)

def main():
    # App title
    st.title("Utilization Dashboard")
    
    # Load the data
    time_entries_df, attorneys_df, clients_df = load_data()
    
    # Create sidebar for filters
    st.sidebar.header("Filters")
    
    # Year filter (dropdown)
    years = ["All"] + sorted(time_entries_df['Year'].unique().tolist(), reverse=True)
    year_filter = st.sidebar.selectbox("Year", years)
    
    # Month filter (dropdown)
    months = ["All"] + [calendar.month_abbr[i] for i in sorted(time_entries_df['Month'].unique().tolist())]
    month_filter = st.sidebar.selectbox("Month", months)
    
    # Revenue band filter
    revenue_bands = ["All"] + sorted(time_entries_df['CLIENT ANNUAL REV'].unique().tolist())
    rev_band_filter = st.sidebar.selectbox("Revenue Band", revenue_bands)
    
    # Attorney filter
    attorneys = ["All"] + sorted(time_entries_df['Associated Attorney'].unique().tolist())
    attorney_filter = st.sidebar.selectbox("Attorney", attorneys)
    
    # Practice group filter
    practice_groups = ["All"] + sorted(time_entries_df['PG1'].unique().tolist())
    pg_filter = st.sidebar.selectbox("Practice Group", practice_groups)
    
    # Fee type filter
    fee_types = ["All", "Time", "Fixed Fee"]
    fee_type_filter = st.sidebar.selectbox("Fee Type", fee_types)
    
    # Clear filters button
    if st.sidebar.button("Clear Filters"):
        # This will trigger a rerun with default values
        st.experimental_rerun()
    
    # Apply filters
    filtered_df = filter_data(time_entries_df, year_filter, month_filter, rev_band_filter, attorney_filter, pg_filter, fee_type_filter)
    
    # Create tabs for different sections
    tabs = st.tabs(["Overview", "Client Analysis", "Fee Bands", "Client Segmentation", "Attorney Analysis", "Fee Trends"])
    
    with tabs[0]:
        create_overview_section(filtered_df, time_entries_df, attorneys_df)
    
    with tabs[1]:
        create_client_analysis(filtered_df)
    
    with tabs[2]:
        create_revenue_bands(filtered_df)
    
    with tabs[3]:
        create_client_segmentation(filtered_df)
    
    with tabs[4]:
        create_attorney_analysis(filtered_df, attorneys_df)
    
    with tabs[5]:
        create_fee_trends(filtered_df)

if __name__ == "__main__":
    main()
