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
    /* Make charts stand out in expanders */
    .st-emotion-cache-1r4qj8v {
        border: 1px solid #ddd;
        border-radius: 5px;
        box-shadow: 0 0 5px rgba(0, 0, 0, 0.1);
        margin-bottom: 1rem;
    }
    /* Make expanders a bit more prominent */
    .st-emotion-cache-eqpbcq {
        border: 1px solid #e6e9ef;
        border-radius: 8px;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Define standardized color scheme to use across all visualizations
def get_color_scheme():
    """Return standardized color schemes for consistency"""
    return {
        'primary': '#4e8df5',       # Main color for primary metrics (blue)
        'secondary': '#4CAF50',     # Secondary color (green)
        'accent': '#FF9800',        # Accent color (orange)
        'neutral': '#607D8B',       # Neutral color (blue-grey)
        'sequence': px.colors.sequential.Blues,  # Sequential color scale
        'categorical': px.colors.qualitative.Safe,  # Categorical colors
        'diverging': px.colors.diverging.RdBu,  # Diverging color scale
    }

def create_scrollable_bar_chart(df, x, y, title, x_label=None, y_label=None, color=None, orientation='v', height=600, show_top_n=10):
    """
    Create a bar chart that can be scrolled to show all data while defaulting to show top N items
    
    Parameters:
    - df: DataFrame containing the data
    - x, y: Column names for x and y axes
    - title: Chart title
    - x_label, y_label: Axis labels (defaults to column names if None)
    - color: Color to use (from standardized palette)
    - orientation: 'v' for vertical bars, 'h' for horizontal
    - height: Chart height
    - show_top_n: Number of items to show by default (set to None to show all)
    """
    colors = get_color_scheme()
    
    # If color is specified as a key in our scheme, use it
    if color and color in colors:
        chart_color = colors[color]
    else:
        # Default to primary color
        chart_color = colors['primary']
    
    # Sort data appropriately based on orientation
    if orientation == 'v':
        # For vertical bars, sort by y value
        df_sorted = df.sort_values(y, ascending=False)
        x_title = x_label if x_label else x
        y_title = y_label if y_label else y
    else:
        # For horizontal bars, sort by x value
        df_sorted = df.sort_values(x, ascending=False)
        # Swap x and y labels for horizontal orientation
        x_title = y_label if y_label else y
        y_title = x_label if x_label else x
    
    # Create a copy for display to avoid modifying original
    display_df = df_sorted.copy()
    
    # Limit to top N for initial view if specified
    if show_top_n and len(display_df) > show_top_n:
        display_df = display_df.head(show_top_n)
    
    # Create figure
    if orientation == 'v':
        fig = px.bar(
            display_df, 
            x=x, 
            y=y,
            title=title,
            labels={x: x_title, y: y_title},
            color_discrete_sequence=[chart_color] if isinstance(chart_color, str) else chart_color
        )
    else:
        fig = px.bar(
            display_df, 
            y=x, 
            x=y,
            title=title,
            labels={y: x_title, x: y_title},
            color_discrete_sequence=[chart_color] if isinstance(chart_color, str) else chart_color,
            orientation='h'
        )
    
    # Add layout settings for scrolling
    fig.update_layout(
        height=height,
        xaxis_tickangle=-45 if orientation == 'v' else 0,
        margin=dict(l=50, r=50, b=100, t=100, pad=4),
        autosize=True
    )
    
    # Create expander with information about scrolling
    with st.expander(f"{title} - Showing top {show_top_n if show_top_n else 'all'} of {len(df)} items", expanded=True):
        # Add note about scrolling if there's more data
        if show_top_n and len(df) > show_top_n:
            st.info(f"Showing top {show_top_n} items. Use the interactive chart controls to explore all {len(df)} items.")
        
        # Display chart
        st.plotly_chart(fig, use_container_width=True)
        
        # Option to show full data
        if st.checkbox(f"Show all {len(df)} items in table format"):
            st.dataframe(df_sorted, use_container_width=True, height=min(400, len(df) * 35))

@st.cache_data
def load_data():
    """Load and prepare the data with error handling"""
    try:
        # Load the TIME ENTRIES sheet
        time_entries_df = pd.read_excel('Utilization.xlsx', sheet_name='TIME ENTRIES')
        
        # Try to load the ATTORNEYS sheet, but handle gracefully if it doesn't exist
        try:
            attorneys_df = pd.read_excel('Utilization.xlsx', sheet_name='ATTORNEYS')
        except Exception:
            st.warning("Could not load the ATTORNEYS sheet. Attorney-specific features will be limited.")
            attorneys_df = pd.DataFrame()
        
        # Try to load the CLIENTS sheet, but handle gracefully if it doesn't exist
        try:
            clients_df = pd.read_excel('Utilization.xlsx', sheet_name='CLIENTS')
        except Exception:
            st.warning("Could not load the CLIENTS sheet. Client-specific features will be limited.")
            clients_df = pd.DataFrame()
        
        # Clean and prepare the data
        
        # Convert date columns to datetime
        if 'Date' in time_entries_df.columns:
            time_entries_df['Date'] = pd.to_datetime(time_entries_df['Date'], errors='coerce')
            
            # Extract month, year components for filtering
            time_entries_df['Month'] = time_entries_df['Date'].dt.month
            time_entries_df['Year'] = time_entries_df['Date'].dt.year
            time_entries_df['MonthName'] = time_entries_df['Date'].dt.strftime('%b')
            time_entries_df['MonthYear'] = time_entries_df['Date'].dt.strftime('%b %Y')
        
        # Remove "$" and convert to numeric
        if 'Billable ($)' in time_entries_df.columns:
            time_entries_df['Billable ($)'] = pd.to_numeric(
                time_entries_df['Billable ($)'].astype(str).str.replace('$', '').str.replace(',', ''), 
                errors='coerce'
            )
        
        if 'Rate ($)' in time_entries_df.columns:
            time_entries_df['Rate ($)'] = pd.to_numeric(
                time_entries_df['Rate ($)'].astype(str).str.replace('$', '').str.replace(',', ''), 
                errors='coerce'
            )
        
        # Create fee type column
        if 'Type' in time_entries_df.columns:
            time_entries_df['FeeType'] = time_entries_df['Type'].apply(
                lambda x: 'Fixed Fee' if 'FixedFee' in str(x) else ('Time' if 'TimeEntry' in str(x) else 'Other')
            )
        
        # Clean attorneys data
        if not attorneys_df.empty and '🎚️ Target Hours / Month' in attorneys_df.columns:
            attorneys_df['Target Hours'] = attorneys_df['🎚️ Target Hours / Month']
        
        # FILTER OUT ATTORNEYS NO LONGER WITH THE FIRM (with "x-" in their name)
        if 'Associated Attorney' in time_entries_df.columns:
            active_attorneys = ~time_entries_df['Associated Attorney'].str.contains('x-', case=False, na=False)
            time_entries_df = time_entries_df[active_attorneys]
            
        if not attorneys_df.empty and 'Attorney Name' in attorneys_df.columns:
            active_attorneys = ~attorneys_df['Attorney Name'].str.contains('x-', case=False, na=False)
            attorneys_df = attorneys_df[active_attorneys]
        
        return time_entries_df, attorneys_df, clients_df
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        # Return empty dataframes
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def filter_data(df, year_filter, month_filter, rev_band_filter, attorney_filter, pg_filter, fee_type_filter, client_filter, fx_filter):
    """Apply filters to the dataframe, handling missing columns gracefully"""
    filtered_df = df.copy()
    
    # Apply year filter
    if year_filter != "All" and 'Year' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Year'] == int(year_filter)]
    
    # Apply month filter
    if month_filter != "All" and 'Month' in filtered_df.columns:
        month_num = list(calendar.month_abbr).index(month_filter)
        filtered_df = filtered_df[filtered_df['Month'] == month_num]
    
    # Apply revenue band filter
    if rev_band_filter != "All":
        # First try with the expected column name
        if 'CLIENT ANNUAL REV' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['CLIENT ANNUAL REV'] == rev_band_filter]
        else:
            # Try alternative column names
            possible_rev_columns = [col for col in filtered_df.columns if 'REV' in col.upper() or 'REVENUE' in col.upper()]
            if possible_rev_columns:
                filtered_df = filtered_df[filtered_df[possible_rev_columns[0]] == rev_band_filter]
    
    # Apply client filter (NEW)
    if client_filter != "All" and 'Client' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Client'] == client_filter]
    
    # Apply attorney filter
    if attorney_filter != "All" and 'Associated Attorney' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Associated Attorney'] == attorney_filter]
    
    # Apply practice group filter
    if pg_filter != "All" and 'PG1' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['PG1'] == pg_filter]
    
    # Apply FX filter (NEW)
    if fx_filter != "All":
        # Try both common column names for FX
        if 'xF' in filtered_df.columns:
            # Convert filter to appropriate type based on dataframe column
            try:
                if filtered_df['xF'].dtype in (np.int64, np.float64):
                    filtered_df = filtered_df[filtered_df['xF'] == float(fx_filter)]
                else:
                    filtered_df = filtered_df[filtered_df['xF'] == fx_filter]
            except:
                # Fall back to string comparison if conversion fails
                filtered_df = filtered_df[filtered_df['xF'].astype(str) == fx_filter]
        elif 'FX' in filtered_df.columns:
            try:
                if filtered_df['FX'].dtype in (np.int64, np.float64):
                    filtered_df = filtered_df[filtered_df['FX'] == float(fx_filter)]
                else:
                    filtered_df = filtered_df[filtered_df['FX'] == fx_filter]
            except:
                filtered_df = filtered_df[filtered_df['FX'].astype(str) == fx_filter]
    
    # Apply fee type filter
    if fee_type_filter != "All" and 'FeeType' in filtered_df.columns:
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
    """Calculate key metrics from filtered data, handling missing columns"""
    metrics = {
        'total_billable_hours': 0,
        'total_fees': 0,
        'avg_rate': 0,
        'monthly_bills_generated': 0
    }
    
    if filtered_df.empty:
        return metrics
    
    # Calculate total billable hours
    if 'Quantity / Hours' in filtered_df.columns:
        metrics['total_billable_hours'] = filtered_df['Quantity / Hours'].sum()
    
    # Calculate total fees
    if 'Billable ($)' in filtered_df.columns:
        metrics['total_fees'] = filtered_df['Billable ($)'].sum()
    
    # Calculate average rate
    if metrics['total_billable_hours'] > 0 and metrics['total_fees'] > 0:
        metrics['avg_rate'] = metrics['total_fees'] / metrics['total_billable_hours']
    
    # Calculate monthly bills generated
    if 'Bill Number' in filtered_df.columns:
        metrics['monthly_bills_generated'] = filtered_df[filtered_df['Bill Number'].notna()]['Bill Number'].nunique()
    
    return metrics

def create_overview_section(filtered_df, time_entries_df, attorneys_df):
    """Create the overview section with key metrics and visualizations"""
    # Get color scheme
    colors = get_color_scheme()
    
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
    
    # Check if we have the necessary columns for visualizations
    required_cols = ['MonthYear', 'Quantity / Hours']
    if not all(col in filtered_df.columns for col in required_cols) or filtered_df.empty:
        st.warning("Insufficient data for visualizations. Please check your Excel file structure.")
        return
    
    # Create two column layout for charts
    col1, col2 = st.columns(2)
    
    with col1:
        with st.expander("Monthly Billable Hours", expanded=True):
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
                         color_discrete_sequence=[colors['primary']])
            fig.update_layout(xaxis_title="Month", yaxis_title="Hours", height=350)
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Utilization vs Target
        if not attorneys_df.empty and 'Target Hours' in attorneys_df.columns and 'Attorney Name' in attorneys_df.columns:
            with st.expander("Attorney Utilization", expanded=True):
                # Get attorneys in filtered data
                if 'Associated Attorney' in filtered_df.columns:
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
                                 color_discrete_sequence=[colors['secondary']])
                    
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
                else:
                    st.info("Attorney data not available. Cannot show utilization chart.")
        else:
            st.info("Attorney target data not available. Cannot show utilization chart.")

def create_client_analysis(filtered_df):
    """Create client analysis section"""
    # Get color scheme
    colors = get_color_scheme()
    
    st.subheader("Client Analysis")
    
    # Check for required columns
    if 'Client' not in filtered_df.columns or 'Billable ($)' not in filtered_df.columns or filtered_df.empty:
        st.warning("Required data for client analysis is missing. Please check your Excel file structure.")
        return
    
    # Fees by client
    client_fees = filtered_df.groupby('Client')['Billable ($)'].sum().reset_index()
    client_fees = client_fees.sort_values('Billable ($)', ascending=False)
    
    create_scrollable_bar_chart(
        client_fees,
        'Client',
        'Billable ($)',
        'Top Clients by Fees',
        'Client',
        'Fees ($)',
        color='primary',
        height=500,
        show_top_n=10
    )
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        if 'Quantity / Hours' in filtered_df.columns:
            with st.expander("Top Clients by Hours", expanded=True):
                # Hours by client
                client_hours = filtered_df.groupby('Client')['Quantity / Hours'].sum().reset_index()
                client_hours = client_hours.sort_values('Quantity / Hours', ascending=False)
                
                fig = px.bar(client_hours.head(10), x='Client', y='Quantity / Hours',
                             title='Top 10 Clients by Hours',
                             labels={'Client': 'Client', 'Quantity / Hours': 'Hours'},
                             color_discrete_sequence=[colors['secondary']])
                fig.update_layout(xaxis_title="Client", yaxis_title="Hours", height=350,
                                 xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
                
                if st.checkbox("Show all clients by hours"):
                    st.dataframe(client_hours, hide_index=True, use_container_width=True, height=400)
        else:
            st.info("Hours data not available for this visualization.")
    
    with col2:
        if 'CLIENT INDUSTRY' in filtered_df.columns:
            with st.expander("Fees by Industry", expanded=True):
                # Fees by industry
                industry_fees = filtered_df.groupby('CLIENT INDUSTRY')['Billable ($)'].sum().reset_index()
                industry_fees = industry_fees.sort_values('Billable ($)', ascending=False)
                
                fig = px.pie(industry_fees, values='Billable ($)', names='CLIENT INDUSTRY',
                            title='Fees by Industry',
                            color_discrete_sequence=colors['categorical'])
                fig.update_layout(height=350)
                st.plotly_chart(fig, use_container_width=True)
        else:
            # Try to find alternative industry column
            industry_cols = [col for col in filtered_df.columns if 'INDUSTRY' in col.upper()]
            if industry_cols:
                with st.expander("Fees by Industry", expanded=True):
                    industry_fees = filtered_df.groupby(industry_cols[0])['Billable ($)'].sum().reset_index()
                    industry_fees = industry_fees.sort_values('Billable ($)', ascending=False)
                    
                    fig = px.pie(industry_fees, values='Billable ($)', names=industry_cols[0],
                                title='Fees by Industry',
                                color_discrete_sequence=colors['categorical'])
                    fig.update_layout(height=350)
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Industry data not available for this visualization.")
    
    # Detailed client metrics
    with st.expander("Detailed Client Metrics", expanded=True):
        st.subheader("Detailed Client Metrics")
        client_metrics_columns = {
            'Billable ($)': 'Total Fees',
            'Quantity / Hours': 'Total Hours',
            'Bill Number': 'Number of Bills'
        }
        
        # Check which columns are available
        available_columns = [col for col in client_metrics_columns.keys() if col in filtered_df.columns]
        
        if available_columns:
            client_metrics = filtered_df.groupby('Client').agg({
                col: 'sum' if col != 'Bill Number' else pd.Series.nunique 
                for col in available_columns
            }).reset_index()
            
            # Rename columns
            column_mapping = {'Client': 'Client'}
            column_mapping.update({col: client_metrics_columns[col] for col in available_columns})
            client_metrics.columns = [column_mapping.get(col, col) for col in client_metrics.columns]
            
            # Calculate derived metrics if possible
            if 'Total Fees' in client_metrics.columns and 'Total Hours' in client_metrics.columns:
                client_metrics['Average Rate'] = client_metrics['Total Fees'] / client_metrics['Total Hours']
            
            if 'Total Fees' in client_metrics.columns and 'Number of Bills' in client_metrics.columns:
                client_metrics['Average Bill Amount'] = client_metrics['Total Fees'] / client_metrics['Number of Bills'].replace(0, np.nan)
            
            # Sort by total fees if available
            if 'Total Fees' in client_metrics.columns:
                client_metrics = client_metrics.sort_values('Total Fees', ascending=False)
                
                # Format currency columns
                client_metrics['Total Fees'] = client_metrics['Total Fees'].apply(lambda x: f"${x:,.2f}")
                
                if 'Average Rate' in client_metrics.columns:
                    client_metrics['Average Rate'] = client_metrics['Average Rate'].apply(
                        lambda x: f"${x:,.2f}" if not pd.isna(x) else "N/A"
                    )
                
                if 'Average Bill Amount' in client_metrics.columns:
                    client_metrics['Average Bill Amount'] = client_metrics['Average Bill Amount'].apply(
                        lambda x: f"${x:,.2f}" if not pd.isna(x) else "N/A"
                    )
            
            # Hide index and allow scrolling
            st.dataframe(client_metrics, hide_index=True, use_container_width=True, height=400)
        else:
            st.info("Required data for detailed client metrics is not available.")

def create_revenue_bands(filtered_df):
    """Create revenue bands analysis section"""
    # Get color scheme
    colors = get_color_scheme()
    
    st.subheader("Fee Bands Analysis")
    
    # Check if we have the necessary columns
    revenue_band_col = None
    
    if 'CLIENT ANNUAL REV' in filtered_df.columns:
        revenue_band_col = 'CLIENT ANNUAL REV'
    else:
        # Try to find alternative column
        possible_cols = [col for col in filtered_df.columns if 'REV' in col.upper() or 'REVENUE' in col.upper()]
        if possible_cols:
            revenue_band_col = possible_cols[0]
    
    if not revenue_band_col or 'Billable ($)' not in filtered_df.columns or filtered_df.empty:
        st.warning("Revenue band data is not available. Please check your Excel file structure.")
        return
    
    # Fees by revenue band
    rev_band_fees = filtered_df.groupby(revenue_band_col)['Billable ($)'].sum().reset_index()
    rev_band_fees = rev_band_fees.sort_values('Billable ($)', ascending=False)
    
    # Try to define a sorting order for revenue bands if they follow a pattern
    try:
        # This is a common pattern for revenue bands, adjust as needed
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
        
        rev_band_fees['sort_order'] = rev_band_fees[revenue_band_col].map(
            lambda x: sorting_order.get(x, 999)  # Default to a high number for unknown values
        )
        rev_band_fees = rev_band_fees.sort_values('sort_order').drop('sort_order', axis=1)
    except Exception:
        # If sorting fails, just use the default sort
        pass
    
    with st.expander("Fees by Revenue Band", expanded=True):
        fig = px.bar(rev_band_fees, x=revenue_band_col, y='Billable ($)',
                     title='Fees by Client Annual Revenue Band',
                     labels={revenue_band_col: 'Annual Revenue Band', 'Billable ($)': 'Fees ($)'},
                     color_discrete_sequence=[colors['primary']])
        fig.update_layout(xaxis_title="Annual Revenue Band", yaxis_title="Fees ($)", height=400,
                         xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        if 'Quantity / Hours' in filtered_df.columns:
            with st.expander("Hours by Revenue Band", expanded=True):
                # Hours by revenue band
                rev_band_hours = filtered_df.groupby(revenue_band_col)['Quantity / Hours'].sum().reset_index()
                
                # Apply same sorting if available
                try:
                    rev_band_hours['sort_order'] = rev_band_hours[revenue_band_col].map(
                        lambda x: sorting_order.get(x, 999)
                    )
                    rev_band_hours = rev_band_hours.sort_values('sort_order').drop('sort_order', axis=1)
                except Exception:
                    # If sorting fails, just use the default sort
                    pass
                
                fig = px.bar(rev_band_hours, x=revenue_band_col, y='Quantity / Hours',
                             title='Hours by Client Annual Revenue Band',
                             labels={revenue_band_col: 'Annual Revenue Band', 'Quantity / Hours': 'Hours'},
                             color_discrete_sequence=[colors['secondary']])
                fig.update_layout(xaxis_title="Annual Revenue Band", yaxis_title="Hours", height=350,
                                 xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Hours data not available for this visualization.")
    
    with col2:
        if 'Client' in filtered_df.columns:
            with st.expander("Clients by Revenue Band", expanded=True):
                # Client count by revenue band
                rev_band_clients = filtered_df.groupby(revenue_band_col)['Client'].nunique().reset_index()
                
                # Apply same sorting if available
                try:
                    rev_band_clients['sort_order'] = rev_band_clients[revenue_band_col].map(
                        lambda x: sorting_order.get(x, 999)
                    )
                    rev_band_clients = rev_band_clients.sort_values('sort_order').drop('sort_order', axis=1)
                except Exception:
                    # If sorting fails, just use the default sort
                    pass
                
                fig = px.bar(rev_band_clients, x=revenue_band_col, y='Client',
                             title='Number of Clients by Annual Revenue Band',
                             labels={revenue_band_col: 'Annual Revenue Band', 'Client': 'Number of Clients'},
                             color_discrete_sequence=[colors['accent']])
                fig.update_layout(xaxis_title="Annual Revenue Band", yaxis_title="Number of Clients", height=350,
                                 xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Client data not available for this visualization.")

def create_client_segmentation(filtered_df):
    """Create client segmentation section with enhanced value analysis"""
    # Get color scheme
    colors = get_color_scheme()
    
    st.subheader("Client Segmentation")
    
    # Check for required columns
    stage_col = None
    
    if 'CLIENT STAGE' in filtered_df.columns:
        stage_col = 'CLIENT STAGE'
    else:
        # Try to find alternative column
        possible_cols = [col for col in filtered_df.columns if 'STAGE' in col.upper()]
        if possible_cols:
            stage_col = possible_cols[0]
    
    if not stage_col or 'Billable ($)' not in filtered_df.columns or filtered_df.empty:
        st.warning("Client stage data is not available. Please check your Excel file structure.")
        return
    
    # Fees by client stage
    stage_fees = filtered_df.groupby(stage_col)['Billable ($)'].sum().reset_index()
    
    # Sort by fees
    stage_fees = stage_fees.sort_values('Billable ($)', ascending=False)
    
    # Create container for expandable visualization
    with st.expander("Fees by Client Stage", expanded=True):
        fig = px.bar(stage_fees, x=stage_col, y='Billable ($)',
                    title='Fees by Client Stage',
                    labels={stage_col: 'Client Stage', 'Billable ($)': 'Fees ($)'},
                    color_discrete_sequence=[colors['primary']])
        fig.update_layout(xaxis_title="Client Stage", yaxis_title="Fees ($)", height=500)
        st.plotly_chart(fig, use_container_width=True)
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        if 'PG1' in filtered_df.columns:
            # Fees by practice area
            pa_fees = filtered_df.groupby('PG1')['Billable ($)'].sum().reset_index()
            pa_fees = pa_fees.sort_values('Billable ($)', ascending=False)
            
            # Create container for expandable visualization
            with st.expander("Fees by Practice Area", expanded=True):
                fig = px.bar(pa_fees, x='PG1', y='Billable ($)',
                            title='Fees by Practice Area',
                            labels={'PG1': 'Practice Area', 'Billable ($)': 'Fees ($)'},
                            color_discrete_sequence=[colors['secondary']])
                fig.update_layout(xaxis_title="Practice Area", yaxis_title="Fees ($)", height=600)
                st.plotly_chart(fig, use_container_width=True)
        else:
            # Try to find alternative practice area column
            pa_cols = [col for col in filtered_df.columns if 'PRACTICE' in col.upper() or 'PG' in col.upper()]
            if pa_cols:
                pa_fees = filtered_df.groupby(pa_cols[0])['Billable ($)'].sum().reset_index()
                pa_fees = pa_fees.sort_values('Billable ($)', ascending=False)
                
                with st.expander("Fees by Practice Area", expanded=True):
                    fig = px.bar(pa_fees, x=pa_cols[0], y='Billable ($)',
                                title='Fees by Practice Area',
                                labels={pa_cols[0]: 'Practice Area', 'Billable ($)': 'Fees ($)'},
                                color_discrete_sequence=[colors['secondary']])
                    fig.update_layout(xaxis_title="Practice Area", yaxis_title="Fees ($)", height=600)
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Practice area data not available for this visualization.")
    
    with col2:
        if 'Client' in filtered_df.columns:
            try:
                # Average fees per client by stage
                stage_avg_fees = filtered_df.groupby([stage_col, 'Client'])['Billable ($)'].sum().reset_index()
                stage_avg_fees = stage_avg_fees.groupby(stage_col).agg({
                    'Billable ($)': 'mean',
                    'Client': 'count'
                }).reset_index()
                stage_avg_fees.columns = [stage_col, 'Avg Fees per Client', 'Number of Clients']
                
                with st.expander("Client Distribution by Stage", expanded=True):
                    fig = px.scatter(stage_avg_fees, x='Avg Fees per Client', y='Number of Clients', 
                                    size='Avg Fees per Client', color=stage_col,
                                    title='Average Fees per Client vs Number of Clients by Stage',
                                    labels={'Avg Fees per Client': 'Average Fees per Client ($)', 
                                            'Number of Clients': 'Number of Clients',
                                            stage_col: 'Client Stage'})
                    fig.update_layout(height=600)
                    st.plotly_chart(fig, use_container_width=True)
            except Exception:
                st.info("Could not calculate average fees per client by stage.")
        else:
            st.info("Client data not available for this visualization.")
    
    # NEW SECTION: Client Value Analysis
    st.subheader("Client Value Analysis")
    
    # Check for required columns
    if 'Client' in filtered_df.columns and 'Billable ($)' in filtered_df.columns:
        # Calculate client value
        client_value = filtered_df.groupby('Client').agg({
            'Billable ($)': 'sum',
            'Quantity / Hours': 'sum' if 'Quantity / Hours' in filtered_df.columns else 'count',
            'Bill Number': pd.Series.nunique if 'Bill Number' in filtered_df.columns else 'count'
        }).reset_index()
        
        # Calculate average bill value
        if 'Bill Number' in filtered_df.columns:
            client_value['Avg Bill Value'] = client_value['Billable ($)'] / client_value['Bill Number']
        
        # Sort by total fees
        client_value = client_value.sort_values('Billable ($)', ascending=False)
        
        # Top clients by fees
        top_clients = client_value.head(10)
        
        with st.expander("Top Clients by Value", expanded=True):
            # Create a horizontal bar chart for better readability with many clients
            fig = px.bar(client_value, y='Client', x='Billable ($)',
                        title='Clients by Total Fees',
                        labels={'Client': 'Client', 'Billable ($)': 'Total Fees ($)'},
                        color_discrete_sequence=[colors['primary']],
                        orientation='h')  # Horizontal orientation
            
            fig.update_layout(yaxis_title="Client", xaxis_title="Total Fees ($)", height=800)
            # Enable scrolling for many clients
            fig.update_layout(
                autosize=True,
                margin=dict(l=50, r=50, b=100, t=100, pad=4),
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Client value data not available for this visualization.")
    
    # NEW SECTION: Lifetime Value Analysis
    st.subheader("Lifetime Value Analysis")
    
    # Check for required columns for LTV analysis
    revenue_band_col = None
    if 'CLIENT ANNUAL REV' in filtered_df.columns:
        revenue_band_col = 'CLIENT ANNUAL REV'
    else:
        # Try to find alternative column
        possible_cols = [col for col in filtered_df.columns if 'REV' in col.upper() or 'REVENUE' in col.upper()]
        if possible_cols:
            revenue_band_col = possible_cols[0]
    
    if revenue_band_col and 'Client' in filtered_df.columns and 'Billable ($)' in filtered_df.columns:
        # Calculate lifetime value by revenue band
        ltv_by_band = filtered_df.groupby([revenue_band_col, 'Client'])['Billable ($)'].sum().reset_index()
        ltv_band_summary = ltv_by_band.groupby(revenue_band_col).agg({
            'Billable ($)': ['mean', 'median', 'sum'],
            'Client': 'nunique'
        })
        
        # Flatten the multi-level columns
        ltv_band_summary.columns = ['_'.join(col).strip() for col in ltv_band_summary.columns.values]
        ltv_band_summary = ltv_band_summary.reset_index()
        
        # Rename columns for clarity
        ltv_band_summary = ltv_band_summary.rename(columns={
            'Billable ($)_mean': 'Avg Client LTV',
            'Billable ($)_median': 'Median Client LTV',
            'Billable ($)_sum': 'Total Fees',
            'Client_nunique': 'Number of Clients'
        })
        
        # Calculate average fees per client
        ltv_band_summary['Avg Fees per Client'] = ltv_band_summary['Total Fees'] / ltv_band_summary['Number of Clients']
        
        # Sort by revenue band if possible
        try:
            # Common pattern for revenue bands
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
            
            ltv_band_summary['sort_order'] = ltv_band_summary[revenue_band_col].map(
                lambda x: sorting_order.get(x, 999)
            )
            ltv_band_summary = ltv_band_summary.sort_values('sort_order')
            ltv_band_summary = ltv_band_summary.drop('sort_order', axis=1)
        except Exception:
            # If sorting fails, sort by average client LTV
            ltv_band_summary = ltv_band_summary.sort_values('Avg Client LTV', ascending=False)
        
        # Show the LTV analysis
        with st.expander("Lifetime Value by Revenue Band", expanded=True):
            # First create a bar chart of average LTV by revenue band
            fig = px.bar(ltv_band_summary, x=revenue_band_col, y='Avg Client LTV',
                         title='Average Client Lifetime Value by Revenue Band',
                         labels={revenue_band_col: 'Revenue Band', 'Avg Client LTV': 'Average Client Lifetime Value ($)'},
                         color_discrete_sequence=[colors['primary']])
            fig.update_layout(xaxis_title="Revenue Band", yaxis_title="Average LTV ($)", height=500)
            st.plotly_chart(fig, use_container_width=True)
            
            # Then create a scatter plot of client count vs average fees
            fig = px.scatter(ltv_band_summary, x='Number of Clients', y='Avg Fees per Client',
                            size='Total Fees', color=revenue_band_col,
                            title='Client Distribution and Value by Revenue Band',
                            labels={'Number of Clients': 'Number of Clients', 
                                    'Avg Fees per Client': 'Average Fees per Client ($)',
                                    'Total Fees': 'Total Fees',
                                    revenue_band_col: 'Revenue Band'})
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
        
        # Show Top Clients by Lifetime Value
        with st.expander("Top Clients by Lifetime Value", expanded=True):
            top_ltv_clients = ltv_by_band.sort_values('Billable ($)', ascending=False).head(50)
            
            fig = px.bar(top_ltv_clients, y='Client', x='Billable ($)',
                       title='Top Clients by Lifetime Value',
                       labels={'Client': 'Client', 'Billable ($)': 'Lifetime Value ($)'},
                       color=revenue_band_col,
                       orientation='h')  # Horizontal for better readability
            
            fig.update_layout(yaxis_title="Client", xaxis_title="Lifetime Value ($)", height=800)
            st.plotly_chart(fig, use_container_width=True)
            
            # Also show as a table
            st.subheader("Top Clients by Lifetime Value (Table)")
            top_ltv_table = top_ltv_clients.head(20).copy()
            top_ltv_table['Billable ($)'] = top_ltv_table['Billable ($)'].apply(lambda x: f"${x:,.2f}")
            st.dataframe(top_ltv_table, hide_index=True, use_container_width=True, height=600)
    else:
        st.info("Required data for lifetime value analysis not available.")

def create_attorney_analysis(filtered_df, attorneys_df):
    """Create attorney analysis section with enhanced metrics and visualizations"""
    # Get color scheme
    colors = get_color_scheme()
    
    st.subheader("Attorney Analysis")
    
    # Check for required columns
    if 'Associated Attorney' not in filtered_df.columns or 'Quantity / Hours' not in filtered_df.columns or filtered_df.empty:
        st.warning("Attorney data is not available. Please check your Excel file structure.")
        return
    
    # Hours vs target by attorney
    if not attorneys_df.empty and 'Target Hours' in attorneys_df.columns and 'Attorney Name' in attorneys_df.columns:
        try:
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
            
            with st.expander("Attorney Utilization vs Target", expanded=True):
                fig = px.bar(attorney_util, x='Associated Attorney', y='Utilization %',
                            title='Attorney Utilization vs Target',
                            labels={'Associated Attorney': 'Attorney', 'Utilization %': 'Utilization %'},
                            color_discrete_sequence=[colors['secondary']])
                
                # Add reference line at 100%
                fig.add_shape(
                    type="line",
                    x0=-0.5,
                    y0=100,
                    x1=len(attorney_util)-0.5,
                    y1=100,
                    line=dict(color="red", width=2, dash="dash"),
                )
                
                fig.update_layout(xaxis_title="Attorney", yaxis_title="Utilization %", height=600,
                                xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Could not create utilization chart: {str(e)}")
    else:
        st.info("Attorney target data not available. Cannot show utilization chart.")
    
    # Top attorneys by hours
    try:
        attorney_hours = filtered_df.groupby('Associated Attorney')['Quantity / Hours'].sum().reset_index()
        attorney_hours = attorney_hours.sort_values('Quantity / Hours', ascending=False)
        
        create_scrollable_bar_chart(
            attorney_hours,
            'Associated Attorney',
            'Quantity / Hours',
            'Attorneys by Hours',
            'Attorney',
            'Hours',
            color='primary',
            height=600,
            show_top_n=15
        )
    except Exception as e:
        st.warning(f"Could not create hours chart: {str(e)}")
    
    # Top attorneys by fees
    if 'Billable ($)' in filtered_df.columns:
        try:
            attorney_fees = filtered_df.groupby('Associated Attorney')['Billable ($)'].sum().reset_index()
            attorney_fees = attorney_fees.sort_values('Billable ($)', ascending=False)
            
            create_scrollable_bar_chart(
                attorney_fees,
                'Associated Attorney',
                'Billable ($)',
                'Attorneys by Fees',
                'Attorney',
                'Fees ($)',
                color='accent',
                height=600,
                show_top_n=15
            )
        except Exception as e:
            st.warning(f"Could not create fees chart: {str(e)}")
    else:
        st.info("Fee data not available for this visualization.")
    
    # NEW: Average Hours per Attorney
    st.subheader("Attorney Metrics")
    
    try:
        with st.expander("Average Hours per Attorney", expanded=True):
            # Calculate average hours per attorney per month
            if 'MonthYear' in filtered_df.columns:
                # Group by attorney and month, then calculate average hours per month
                attorney_month_hours = filtered_df.groupby(['Associated Attorney', 'MonthYear'])['Quantity / Hours'].sum().reset_index()
                avg_hours_per_month = attorney_month_hours.groupby('Associated Attorney')['Quantity / Hours'].mean().reset_index()
                avg_hours_per_month.columns = ['Attorney', 'Average Hours per Month']
                avg_hours_per_month = avg_hours_per_month.sort_values('Average Hours per Month', ascending=False)
                
                fig = px.bar(avg_hours_per_month, x='Attorney', y='Average Hours per Month',
                            title='Average Hours per Month by Attorney',
                            labels={'Attorney': 'Attorney', 'Average Hours per Month': 'Average Hours'},
                            color_discrete_sequence=[colors['primary']])
                fig.update_layout(xaxis_title="Attorney", yaxis_title="Average Hours per Month", height=600,
                                xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
            else:
                # If month data is not available, just show total hours
                attorney_hours = filtered_df.groupby('Associated Attorney')['Quantity / Hours'].sum().reset_index()
                attorney_hours.columns = ['Attorney', 'Total Hours']
                attorney_hours = attorney_hours.sort_values('Total Hours', ascending=False)
                
                fig = px.bar(attorney_hours, x='Attorney', y='Total Hours',
                            title='Total Hours by Attorney',
                            labels={'Attorney': 'Attorney', 'Total Hours': 'Hours'},
                            color_discrete_sequence=[colors['primary']])
                fig.update_layout(xaxis_title="Attorney", yaxis_title="Hours", height=600,
                                xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.warning(f"Could not create average hours chart: {str(e)}")
    
    # NEW: Average Clients per Attorney
    try:
        with st.expander("Average Clients per Attorney", expanded=True):
            if 'Client' in filtered_df.columns:
                # Count unique clients per attorney
                clients_per_attorney = filtered_df.groupby('Associated Attorney')['Client'].nunique().reset_index()
                clients_per_attorney.columns = ['Attorney', 'Number of Clients']
                clients_per_attorney = clients_per_attorney.sort_values('Number of Clients', ascending=False)
                
                fig = px.bar(clients_per_attorney, x='Attorney', y='Number of Clients',
                            title='Number of Clients per Attorney',
                            labels={'Attorney': 'Attorney', 'Number of Clients': 'Clients'},
                            color_discrete_sequence=[colors['accent']])
                fig.update_layout(xaxis_title="Attorney", yaxis_title="Number of Clients", height=600,
                                xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
                
                # Also calculate average fees per client for each attorney
                if 'Billable ($)' in filtered_df.columns:
                    # Group by attorney and client to get fees per client
                    atty_client_fees = filtered_df.groupby(['Associated Attorney', 'Client'])['Billable ($)'].sum().reset_index()
                    
                    # Then calculate average fees per client for each attorney
                    avg_fees_per_client = atty_client_fees.groupby('Associated Attorney').agg({
                        'Billable ($)': 'mean',
                        'Client': 'count'
                    }).reset_index()
                    
                    avg_fees_per_client.columns = ['Attorney', 'Avg Fees per Client', 'Client Count']
                    avg_fees_per_client = avg_fees_per_client.sort_values('Avg Fees per Client', ascending=False)
                    
                    fig = px.bar(avg_fees_per_client, x='Attorney', y='Avg Fees per Client',
                                title='Average Fees per Client by Attorney',
                                labels={'Attorney': 'Attorney', 'Avg Fees per Client': 'Average Fees per Client ($)'},
                                color_discrete_sequence=[colors['secondary']])
                    fig.update_layout(xaxis_title="Attorney", yaxis_title="Average Fees per Client ($)", height=600,
                                    xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Client data not available for this visualization.")
    except Exception as e:
        st.warning(f"Could not create clients per attorney chart: {str(e)}")
    
    # NEW: FX Score Visualization
    try:
        with st.expander("Attorney FX Scores", expanded=True):
            # Check for FX column
            fx_col = None
            if 'xF' in filtered_df.columns:
                fx_col = 'xF'
            elif 'FX' in filtered_df.columns:
                fx_col = 'FX'
            
            if fx_col:
                # Try first approach - using FX from filtered data
                fx_by_attorney = filtered_df.groupby('Associated Attorney')[fx_col].mean().reset_index()
                fx_by_attorney.columns = ['Attorney', 'FX Score']
                fx_by_attorney = fx_by_attorney.sort_values('FX Score', ascending=False)
                
                fig = px.bar(fx_by_attorney, x='Attorney', y='FX Score',
                            title='Average FX Score by Attorney',
                            labels={'Attorney': 'Attorney', 'FX Score': 'FX Score'},
                            color='FX Score',
                            color_continuous_scale=colors['sequence'])
                fig.update_layout(xaxis_title="Attorney", yaxis_title="FX Score", height=600,
                                xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
                
                # If we also have the attorney FX data directly
                if not attorneys_df.empty and fx_col in attorneys_df.columns:
                    # Alternative visualization using attorney data directly
                    fx_from_attorney_data = attorneys_df[['Attorney Name', fx_col]].copy()
                    fx_from_attorney_data.columns = ['Attorney', 'FX Score']
                    fx_from_attorney_data = fx_from_attorney_data.sort_values('FX Score', ascending=False)
                    
                    # Create a heatmap-style visualization
                    st.subheader("Attorney FX Score Heatmap")
                    fig = px.imshow([fx_from_attorney_data['FX Score'].values],
                                    y=['FX Score'],
                                    x=fx_from_attorney_data['Attorney'].values,
                                    labels=dict(x="Attorney", y="Metric", color="Score"),
                                    color_continuous_scale=colors['sequence'])
                    fig.update_layout(height=300)
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("FX score data not available for this visualization.")
    except Exception as e:
        st.warning(f"Could not create FX score visualization: {str(e)}")
    
    # Attorney utilization table
    with st.expander("Attorney Utilization Details", expanded=True):
        if not attorneys_df.empty and 'Target Hours' in attorneys_df.columns and 'Attorney Name' in attorneys_df.columns:
            try:
                # Prepare attorney utilization table
                attorney_detail = attorney_util.copy()
                
                # Check if 'Practice Area (Primary)' exists
                practice_area_col = None
                if 'Practice Area (Primary)' in attorney_detail.columns:
                    practice_area_col = 'Practice Area (Primary)'
                else:
                    # Look for alternative practice area column
                    pa_cols = [col for col in attorney_detail.columns if 'PRACTICE' in col.upper()]
                    if pa_cols:
                        practice_area_col = pa_cols[0]
                
                # Select columns based on availability
                columns_to_select = ['Associated Attorney', 'Quantity / Hours', 'Target Hours', 'Utilization %']
                if practice_area_col:
                    columns_to_select.append(practice_area_col)
                
                attorney_detail = attorney_detail[columns_to_select]
                
                # Rename columns
                column_mapping = {
                    'Associated Attorney': 'Attorney', 
                    'Quantity / Hours': 'Hours', 
                    'Target Hours': 'Target Hours', 
                    'Utilization %': 'Utilization %'
                }
                if practice_area_col:
                    column_mapping[practice_area_col] = 'Primary Practice Area'
                
                attorney_detail.columns = [column_mapping.get(col, col) for col in attorney_detail.columns]
                
                # Format columns
                attorney_detail['Utilization %'] = attorney_detail['Utilization %'].apply(lambda x: f"{x:.1f}%" if not pd.isna(x) else "N/A")
                
                # Sort by utilization
                attorney_detail = attorney_detail.sort_values('Hours', ascending=False)
                
                # Hide index and allow scrolling
                st.dataframe(attorney_detail, hide_index=True, use_container_width=True, height=600)
            except Exception as e:
                st.warning(f"Could not create attorney utilization table: {str(e)}")
        else:
            st.info("Attorney target data not available. Cannot show utilization table.")
    
    # Practice area distribution
    st.subheader("Practice Area Analysis")
    
    # Look for practice area column
    practice_area_col = None
    if 'PG1' in filtered_df.columns:
        practice_area_col = 'PG1'
    else:
        # Look for alternative practice area column
        pa_cols = [col for col in filtered_df.columns if 'PRACTICE' in col.upper() or 'PG' in col.upper()]
        if pa_cols:
            practice_area_col = pa_cols[0]
    
    if practice_area_col and 'Quantity / Hours' in filtered_df.columns:
        try:
            with st.expander("Practice Area Hours Distribution", expanded=True):
                # Hours by practice area
                practice_hours = filtered_df.groupby(practice_area_col)['Quantity / Hours'].sum().reset_index()
                practice_hours = practice_hours.sort_values('Quantity / Hours', ascending=False)
                
                fig = px.bar(practice_hours, x=practice_area_col, y='Quantity / Hours',
                            title='Hours by Practice Area',
                            labels={practice_area_col: 'Practice Area', 'Quantity / Hours': 'Hours'},
                            color_discrete_sequence=[colors['primary']])
                fig.update_layout(xaxis_title="Practice Area", yaxis_title="Hours", height=500,
                                xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
                
                # Also show as a pie chart
                fig = px.pie(practice_hours, values='Quantity / Hours', names=practice_area_col,
                            title='Hours Distribution by Practice Area',
                            color_discrete_sequence=colors['categorical'])
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Could not create practice area hour distribution: {str(e)}")
    else:
        st.info("Practice area data not available for this visualization.")
    
    # ENHANCED: Practice Area Metrics
    if practice_area_col and 'Billable ($)' in filtered_df.columns:
        try:
            with st.expander("Practice Area Fee Metrics", expanded=True):
                # Fees by practice area
                practice_fees = filtered_df.groupby(practice_area_col)['Billable ($)'].sum().reset_index()
                practice_fees = practice_fees.sort_values('Billable ($)', ascending=False)
                
                fig = px.bar(practice_fees, x=practice_area_col, y='Billable ($)',
                            title='Fees by Practice Area',
                            labels={practice_area_col: 'Practice Area', 'Billable ($)': 'Fees ($)'},
                            color_discrete_sequence=[colors['accent']])
                fig.update_layout(xaxis_title="Practice Area", yaxis_title="Fees ($)", height=500,
                                xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
                
                # Also show as a pie chart for percentage distribution
                fig = px.pie(practice_fees, values='Billable ($)', names=practice_area_col,
                            title='Fees Distribution by Practice Area',
                            color_discrete_sequence=colors['categorical'])
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
                
                # Calculate and show average rate by practice area
                if 'Quantity / Hours' in filtered_df.columns:
                    practice_rates = filtered_df.groupby(practice_area_col).agg({
                        'Billable ($)': 'sum',
                        'Quantity / Hours': 'sum'
                    }).reset_index()
                    
                    practice_rates['Average Rate'] = practice_rates['Billable ($)'] / practice_rates['Quantity / Hours']
                    practice_rates = practice_rates.sort_values('Average Rate', ascending=False)
                    
                    fig = px.bar(practice_rates, x=practice_area_col, y='Average Rate',
                                title='Average Rate by Practice Area',
                                labels={practice_area_col: 'Practice Area', 'Average Rate': 'Average Rate ($)'},
                                color_discrete_sequence=[colors['secondary']])
                    fig.update_layout(xaxis_title="Practice Area", yaxis_title="Average Rate ($)", height=500,
                                    xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Could not create practice area fee metrics: {str(e)}")
    else:
        st.info("Practice area fee data not available for this visualization.")
    
    # Revenue distribution across practice areas and revenue bands
    revenue_band_col = None
    if 'CLIENT ANNUAL REV' in filtered_df.columns:
        revenue_band_col = 'CLIENT ANNUAL REV'
    else:
        # Try to find alternative column
        possible_cols = [col for col in filtered_df.columns if 'REV' in col.upper() or 'REVENUE' in col.upper()]
        if possible_cols:
            revenue_band_col = possible_cols[0]
    
    if practice_area_col and revenue_band_col and 'Billable ($)' in filtered_df.columns:
        try:
            with st.expander("Practice Area and Revenue Band Distribution", expanded=True):
                # Fees distribution across practice areas and revenue bands
                practice_fees_band = filtered_df.groupby([practice_area_col, revenue_band_col])['Billable ($)'].sum().reset_index()
                
                fig = px.sunburst(practice_fees_band, 
                                 path=[practice_area_col, revenue_band_col], 
                                 values='Billable ($)',
                                 title='Fees Distribution by Practice Area and Revenue Band',
                                 color_discrete_sequence=colors['categorical'])
                fig.update_layout(height=700)
                st.plotly_chart(fig, use_container_width=True)
                
                # Also show as a treemap for alternative visualization
                fig = px.treemap(practice_fees_band,
                                path=[practice_area_col, revenue_band_col],
                                values='Billable ($)',
                                title='Fees Distribution by Practice Area and Revenue Band (Treemap)',
                                color_discrete_sequence=colors['categorical'])
                fig.update_layout(height=700)
                st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.warning(f"Could not create practice area and revenue band distribution: {str(e)}")
    else:
        st.info("Practice area or revenue band data not available for this visualization.")

def create_fee_trends(filtered_df):
    """Create fee trends section"""
    # Get color scheme
    colors = get_color_scheme()
    
    st.subheader("Fee Trends Analysis")
    
    # Check for required columns
    if 'MonthYear' not in filtered_df.columns or 'Billable ($)' not in filtered_df.columns or filtered_df.empty:
        st.warning("Required data for fee trends is missing. Please check your Excel file structure.")
        return
    
    try:
        with st.expander("Monthly Fees Trend", expanded=True):
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
                         color_discrete_sequence=[colors['primary']])
            fig.update_layout(xaxis_title="Month", yaxis_title="Fees ($)", height=400)
            st.plotly_chart(fig, use_container_width=True)
    except Exception:
        st.info("Could not create monthly fees trend visualization.")
    
    # Create two column layout for additional charts
    col1, col2 = st.columns(2)
    
    with col1:
        if 'FeeType' in filtered_df.columns and 'Quantity / Hours' in filtered_df.columns:
            try:
                with st.expander("Monthly Hours by Fee Type", expanded=True):
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
                                 barmode='stack',
                                 color_discrete_sequence=[colors['primary'], colors['accent']])
                    fig.update_layout(xaxis_title="Month", yaxis_title="Hours", height=350,
                                     legend_title="Fee Type", xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
            except Exception:
                st.info("Could not create monthly hours by fee type visualization.")
        else:
            st.info("Fee type data not available for this visualization.")
    
    with col2:
        if 'Quantity / Hours' in filtered_df.columns:
            try:
                with st.expander("Monthly Average Rate Trend", expanded=True):
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
                                 color_discrete_sequence=[colors['accent']])
                    fig.update_layout(xaxis_title="Month", yaxis_title="Average Rate ($)", height=350,
                                     xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
            except Exception:
                st.info("Could not create monthly average rate trend visualization.")
        else:
            st.info("Hours data not available for this visualization.")
    
    # Monthly hours and fees by revenue band
    revenue_band_col = None
    if 'CLIENT ANNUAL REV' in filtered_df.columns:
        revenue_band_col = 'CLIENT ANNUAL REV'
    else:
        # Try to find alternative column
        possible_cols = [col for col in filtered_df.columns if 'REV' in col.upper() or 'REVENUE' in col.upper()]
        if possible_cols:
            revenue_band_col = possible_cols[0]
    
    if revenue_band_col and 'Quantity / Hours' in filtered_df.columns:
        st.subheader("Monthly Trends by Revenue Band")
        
        try:
            with st.expander("Monthly Hours by Revenue Band", expanded=True):
                # Monthly hours by revenue band
                monthly_hours_band = filtered_df.groupby(['MonthYear', revenue_band_col])['Quantity / Hours'].sum().reset_index()
                
                # Sort by date
                monthly_hours_band['MonthYear'] = pd.Categorical(monthly_hours_band['MonthYear'], 
                                                           categories=sorted(filtered_df['MonthYear'].unique(), 
                                                                            key=lambda x: datetime.strptime(x, '%b %Y')),
                                                           ordered=True)
                monthly_hours_band = monthly_hours_band.sort_values('MonthYear')
                
                fig = px.bar(monthly_hours_band, x='MonthYear', y='Quantity / Hours', color=revenue_band_col,
                             title='Monthly Hours by Revenue Band',
                             labels={'MonthYear': 'Month', 'Quantity / Hours': 'Hours', revenue_band_col: 'Revenue Band'},
                             barmode='stack',
                             color_discrete_sequence=colors['categorical'])
                fig.update_layout(xaxis_title="Month", yaxis_title="Hours", height=400,
                                 legend_title="Revenue Band", xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        except Exception:
            st.info("Could not create monthly hours by revenue band visualization.")
        
        try:
            with st.expander("Monthly Fees by Revenue Band", expanded=True):
                # Monthly fees by revenue band
                monthly_fees_band = filtered_df.groupby(['MonthYear', revenue_band_col])['Billable ($)'].sum().reset_index()
                
                # Sort by date
                monthly_fees_band['MonthYear'] = pd.Categorical(monthly_fees_band['MonthYear'], 
                                                              categories=sorted(filtered_df['MonthYear'].unique(), 
                                                                              key=lambda x: datetime.strptime(x, '%b %Y')),
                                                              ordered=True)
                monthly_fees_band = monthly_fees_band.sort_values('MonthYear')
                
                fig = px.bar(monthly_fees_band, x='MonthYear', y='Billable ($)', color=revenue_band_col,
                             title='Monthly Fees by Revenue Band',
                             labels={'MonthYear': 'Month', 'Billable ($)': 'Fees ($)', revenue_band_col: 'Revenue Band'},
                             barmode='stack',
                             color_discrete_sequence=colors['categorical'])
                fig.update_layout(xaxis_title="Month", yaxis_title="Fees ($)", height=400,
                                 legend_title="Revenue Band", xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
        except Exception:
            st.info("Could not create monthly fees by revenue band visualization.")
    else:
        st.info("Revenue band data not available for these visualizations.")

def main():
    # App title
    st.title("Utilization Dashboard")
    
    try:
        # Load the data
        time_entries_df, attorneys_df, clients_df = load_data()
        
        if time_entries_df.empty:
            st.error("No data loaded. Please check your Excel file.")
            return
        
        # Display the Excel structure if requested
        if st.sidebar.checkbox("Show Excel Structure"):
            st.sidebar.write("Available columns in TIME ENTRIES sheet:")
            st.sidebar.write(sorted(time_entries_df.columns.tolist()))
            
            if not attorneys_df.empty:
                st.sidebar.write("Available columns in ATTORNEYS sheet:")
                st.sidebar.write(sorted(attorneys_df.columns.tolist()))
            
            if not clients_df.empty:
                st.sidebar.write("Available columns in CLIENTS sheet:")
                st.sidebar.write(sorted(clients_df.columns.tolist()))
        
        # Create sidebar for filters
        st.sidebar.header("Filters")
        
        # Year filter (dropdown)
        years = ["All"]
        if 'Year' in time_entries_df.columns and not time_entries_df['Year'].empty:
            years += sorted(time_entries_df['Year'].dropna().unique().tolist(), reverse=True)
        year_filter = st.sidebar.selectbox("Year", years)
        
        # Month filter (dropdown)
        months = ["All"]
        if 'Month' in time_entries_df.columns and not time_entries_df['Month'].empty:
            months += [calendar.month_abbr[i] for i in sorted(time_entries_df['Month'].dropna().unique().tolist())]
        month_filter = st.sidebar.selectbox("Month", months)
        
        # Revenue band filter
        revenue_bands = ["All"]
        # Check if the column exists and handle potential errors
        if 'CLIENT ANNUAL REV' in time_entries_df.columns:
            try:
                rev_bands = time_entries_df['CLIENT ANNUAL REV'].dropna().unique().tolist()
                if rev_bands:
                    revenue_bands += sorted(rev_bands)
            except Exception as e:
                st.sidebar.warning(f"Note: Error processing revenue bands: {str(e)}")
        else:
            # Try alternative column names in case there's a mismatch
            possible_rev_columns = [col for col in time_entries_df.columns if 'REV' in col.upper() or 'REVENUE' in col.upper()]
            if possible_rev_columns:
                try:
                    rev_bands = time_entries_df[possible_rev_columns[0]].dropna().unique().tolist()
                    if rev_bands:
                        revenue_bands += sorted(rev_bands)
                        # Update the column name for filtering later
                        st.sidebar.info(f"Using '{possible_rev_columns[0]}' as the revenue band column")
                except Exception:
                    pass
        
        rev_band_filter = st.sidebar.selectbox("Revenue Band", revenue_bands)
        
        # Client filter (NEW)
        clients = ["All"]
        if 'Client' in time_entries_df.columns:
            client_list = time_entries_df['Client'].dropna().unique().tolist()
            if client_list:
                clients += sorted(client_list)
        client_filter = st.sidebar.selectbox("Client", clients)
        
        # Attorney filter
        attorneys = ["All"]
        if 'Associated Attorney' in time_entries_df.columns:
            attorneys += sorted(time_entries_df['Associated Attorney'].dropna().unique().tolist())
        attorney_filter = st.sidebar.selectbox("Attorney", attorneys)
        
        # Practice group filter
        practice_groups = ["All"]
        if 'PG1' in time_entries_df.columns:
            practice_groups += sorted(time_entries_df['PG1'].dropna().unique().tolist())
        pg_filter = st.sidebar.selectbox("Practice Group", practice_groups)
        
        # FX filter (NEW)
        fx_values = ["All"]
        if 'xF' in time_entries_df.columns:
            fx_list = time_entries_df['xF'].dropna().unique().tolist()
            if fx_list:
                fx_values += [str(int(x)) if isinstance(x, (int, float)) else str(x) for x in sorted(fx_list)]
        elif 'FX' in time_entries_df.columns:
            fx_list = time_entries_df['FX'].dropna().unique().tolist()
            if fx_list:
                fx_values += [str(int(x)) if isinstance(x, (int, float)) else str(x) for x in sorted(fx_list)]
        fx_filter = st.sidebar.selectbox("FX", fx_values)
        
        # Fee type filter
        fee_types = ["All", "Time", "Fixed Fee"]
        fee_type_filter = st.sidebar.selectbox("Fee Type", fee_types)
        
        # Clear filters button
        if st.sidebar.button("Clear Filters"):
            # This will trigger a rerun with default values
            st.experimental_rerun()
        
        # Apply filters
        filtered_df = filter_data(time_entries_df, year_filter, month_filter, rev_band_filter, 
                                  attorney_filter, pg_filter, fee_type_filter, client_filter, fx_filter)
        
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
            
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        st.info("Please check your Excel file structure and column names. The dashboard expects specific column names like 'Date', 'Billable ($)', 'CLIENT ANNUAL REV', etc.")
        
        # Display available columns to help troubleshoot
        try:
            df = pd.read_excel('Utilization.xlsx', sheet_name='TIME ENTRIES')
            st.write("Available columns in your Excel file:")
            st.write(sorted(df.columns.tolist()))
        except Exception:
            st.warning("Could not read the Excel file. Please ensure 'Utilization.xlsx' is in the correct location and has a sheet named 'TIME ENTRIES'.")

if __name__ == "__main__":
    main()
    
