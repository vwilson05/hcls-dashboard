import streamlit as st
import pandas as pd
import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from google.oauth2.service_account import Credentials
import os
from dotenv import load_dotenv
import openai
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import time
from google.api_core import retry
import re
import indicators
import strategic_targets

# Load environment variables
load_dotenv()

# Configure page
st.set_page_config(
    page_title="Healthcare Delivery Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if 'data' not in st.session_state:
    st.session_state.data = {}

def setup_google_sheets():
    """Initialize Google Sheets connection with retry logic"""
    max_retries = 3
    retry_delay = 2  # seconds
    
    for attempt in range(max_retries):
        try:
            # Check if credentials file exists
            credentials_file = os.getenv('GOOGLE_SHEETS_CREDENTIALS_FILE')
            if not credentials_file or not os.path.exists(credentials_file):
                st.error(f"Credentials file not found: {credentials_file}")
                return None

            # Check if sheet name is configured
            sheet_name = os.getenv('GOOGLE_SHEET_NAME')
            if not sheet_name:
                st.error("GOOGLE_SHEET_NAME not configured in .env file")
                return None

            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            credentials = Credentials.from_service_account_file(
                credentials_file,
                scopes=scopes
            )
            
            gc = gspread.authorize(credentials)
            
            # Try to open the sheet
            try:
                sheet = gc.open(sheet_name)
                # Test the connection by getting the first worksheet
                sheet.get_worksheet(0)
                return sheet
            except gspread.exceptions.SpreadsheetNotFound:
                st.error(f"Spreadsheet '{sheet_name}' not found. Please check the name and sharing permissions.")
                return None
            except Exception as e:
                if attempt < max_retries - 1:
                    st.warning(f"Attempt {attempt + 1} failed. Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                    continue
                else:
                    st.error(f"Error accessing spreadsheet after {max_retries} attempts: {str(e)}")
                    return None
                
        except Exception as e:
            if attempt < max_retries - 1:
                st.warning(f"Attempt {attempt + 1} failed. Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
                continue
            else:
                st.error(f"Unexpected error in setup_google_sheets: {str(e)}")
                return None
    
    return None

def load_sheet_data(sheet, worksheet_name):
    """Load data from a specific worksheet with retry logic"""
    max_retries = 3
    retry_delay = 2  # seconds
    
    for attempt in range(max_retries):
        try:
            worksheet = sheet.worksheet(worksheet_name)
            all_values = worksheet.get_all_values()
            
            if not all_values:
                return pd.DataFrame()
                
            # Convert to DataFrame
            df = pd.DataFrame(all_values[1:], columns=all_values[0])
            
            # Clean column names
            df.columns = df.columns.str.strip()
            
            # Remove empty rows
            df = df.replace('', pd.NA).dropna(how='all')
            
            return df
            
        except Exception as e:
            if attempt < max_retries - 1:
                st.warning(f"Error loading {worksheet_name}, attempt {attempt + 1}. Retrying...")
                time.sleep(retry_delay)
                continue
            else:
                st.error(f"Error loading {worksheet_name} after {max_retries} attempts: {str(e)}")
                return pd.DataFrame()

def get_fiscal_year():
    """Get current fiscal year (July 1 - June 30)"""
    today = datetime.now()
    if today.month >= 7:  # July or later
        return today.year
    return today.year - 1

def calculate_dashboard_metrics(data):
    """Calculate all dashboard metrics"""
    metrics = {}
    
    # Revenue Metrics
    if 'Project Inventory' in data and not data['Project Inventory'].empty:
        try:
            project_df = data['Project Inventory'].copy()
            # Clean and convert Revenue column properly
            project_df['Revenue'] = project_df['Revenue'].astype(str).str.replace('$', '').str.replace(',', '').str.strip()
            project_df['Revenue'] = pd.to_numeric(project_df['Revenue'], errors='coerce').fillna(0)
            
            # Get current fiscal year
            current_fy = get_fiscal_year()
            
            # Calculate revenue metrics
            metrics['total_revenue'] = float(project_df['Revenue'].sum())
            red_projects = project_df[project_df['Status (R/Y/G)'].str.strip().str.lower() == 'red']
            metrics['red_projects'] = len(red_projects)
            metrics['total_projects'] = len(project_df)
            metrics['red_project_revenue'] = float(red_projects['Revenue'].sum())
        except Exception as e:
            st.error(f"Error calculating revenue metrics: {str(e)}")
    
    # Pipeline Metrics
    if 'Pipeline' in data and not data['Pipeline'].empty:
        try:
            pipeline_df = data['Pipeline'].copy()
            pipeline_df['Percieved Annual AMO'] = pipeline_df['Percieved Annual AMO'].astype(str).str.replace('$', '').str.replace(',', '').str.strip()
            pipeline_df['Percieved Annual AMO'] = pd.to_numeric(pipeline_df['Percieved Annual AMO'], errors='coerce').fillna(0)
            
            metrics['total_pipeline'] = float(pipeline_df['Percieved Annual AMO'].sum())
            metrics['total_potential'] = float(pipeline_df['Percieved Annual AMO'].sum())
        except Exception as e:
            st.error(f"Error calculating pipeline metrics: {str(e)}")
    
    # Risk Metrics
    if 'Project Risks' in data and not data['Project Risks'].empty:
        try:
            risk_df = data['Project Risks'].copy()
            
            # Convert Impact to numeric
            risk_df['Impact ($)'] = risk_df['Impact ($)'].astype(str).str.replace('$', '').str.replace(',', '')
            risk_df['Impact ($)'] = pd.to_numeric(risk_df['Impact ($)'], errors='coerce').fillna(0)
            # Fill missing values for Severity and Impact ($)
            risk_df['Severity'] = risk_df['Severity'].fillna('')
            risk_df['Impact ($)'] = risk_df['Impact ($)'].fillna(0)
            
            # Add risk metrics
            high_risk_impact = risk_df[risk_df['Severity'].str.lower() == 'high']['Impact ($)'].sum()
            total_risk_impact = risk_df['Impact ($)'].sum()
            
            metrics['high_risk_impact'] = float(high_risk_impact)
            metrics['total_risk_impact'] = float(total_risk_impact)
            metrics['high_risk_percent'] = float((high_risk_impact/total_risk_impact*100 if total_risk_impact > 0 else 0))
            
        except Exception as e:
            st.error(f"Error calculating risk metrics: {str(e)}")
    
    # Utilization Metrics
    if 'Team Utilization' in data and not data['Team Utilization'].empty:
        try:
            util_df = data['Team Utilization'].copy()
            util_df['Utilization (%)'] = util_df['Utilization (%)'].astype(str).str.replace('%', '')
            util_df['Utilization (%)'] = pd.to_numeric(util_df['Utilization (%)'], errors='coerce').fillna(0)
            
            # Split executive and delivery
            exec_df = util_df[util_df['Role'].str.contains('Executive', case=False, na=False)]
            delivery_df = util_df[~util_df['Role'].str.contains('Executive', case=False, na=False)]
            
            metrics['exec_utilization'] = float(exec_df['Utilization (%)'].mean()) if not exec_df.empty else 0.0
            metrics['delivery_utilization'] = float(delivery_df['Utilization (%)'].mean()) if not delivery_df.empty else 0.0
            metrics['over_utilized_execs'] = len(exec_df[exec_df['Utilization (%)'] > 70]) if not exec_df.empty else 0
            metrics['under_utilized_delivery'] = len(delivery_df[delivery_df['Utilization (%)'] < 70]) if not delivery_df.empty else 0
        except Exception as e:
            st.error(f"Error calculating utilization metrics: {str(e)}")
    
    # Strategic Metrics
    if 'Executive Activity' in data and not data['Executive Activity'].empty:
        try:
            exec_df = data['Executive Activity'].copy()
            exec_df['Strategic Cost ($)'] = exec_df['Strategic Cost ($)'].astype(str).str.replace('$', '').str.replace(',', '')
            exec_df['Strategic Cost ($)'] = pd.to_numeric(exec_df['Strategic Cost ($)'], errors='coerce').fillna(0)
            
            metrics['strategic_cost'] = float(exec_df['Strategic Cost ($)'].sum())
            metrics['strategic_activities'] = len(exec_df)
        except Exception as e:
            st.error(f"Error calculating strategic metrics: {str(e)}")
    
    return metrics

def query_openai(question, data_context):
    """Query OpenAI with the user's question and data context"""
    try:
        client = openai.OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
        
        # Prepare the prompt with data context
        prompt = f"""Based on the following healthcare delivery data:
        {data_context}
        
        Question: {question}
        
        Please provide a clear, concise answer focusing on actionable insights."""
        
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a healthcare delivery analytics assistant. Provide clear, data-driven insights."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=500
        )
        
        return response.choices[0].message.content
    except Exception as e:
        return f"Error querying OpenAI: {str(e)}"

def create_analytics_visualizations(data):
    """Create analytics visualizations for the dashboard"""
    visualizations = []
    
    # 1. Pipeline Analysis
    if 'Pipeline' in data and not data['Pipeline'].empty:
        try:
            # Create pipeline analysis
            pipeline_df = data['Pipeline'].copy()
            
            # Convert columns to numeric, handling string values
            pipeline_df['Percieved Annual AMO'] = pipeline_df['Percieved Annual AMO'].astype(str).str.replace('$', '').str.replace(',', '')
            pipeline_df['Percieved Annual AMO'] = pd.to_numeric(pipeline_df['Percieved Annual AMO'], errors='coerce').fillna(0)
            
            # Sort by annual amount
            pipeline_df = pipeline_df.sort_values('Percieved Annual AMO', ascending=False)
            
            # Create pipeline chart
            fig_pipeline = px.bar(
                pipeline_df.head(10),
                x='Account',
                y='Percieved Annual AMO',
                title='Top 10 Pipeline Opportunities by Annual Amount',
                labels={'Percieved Annual AMO': 'Annual Amount ($)', 'Account': 'Account'},
                color='Percieved Annual AMO',
                color_continuous_scale='RdYlGn'
            )
            fig_pipeline.update_layout(xaxis_tickangle=-45)
            visualizations.append(('Pipeline Analysis', fig_pipeline))
            
            # Add pipeline metrics
            total_potential = pipeline_df['Percieved Annual AMO'].sum()
            
            # Calculate pipeline health metrics
            deal_registered = len(pipeline_df[pipeline_df['Deal Registered YN'].str.lower() == 'y'])
            has_roadmap = len(pipeline_df[pipeline_df['Agreed Upon Roadmap YN'].str.lower() == 'y'])
            has_business_case = len(pipeline_df[pipeline_df['Business Case_ROI YN'].str.lower() == 'y'])
            has_sponsor = len(pipeline_df[pipeline_df['Business Sponsor YN'].str.lower() == 'y'])
            
            pipeline_metrics = {
                'Total Pipeline Value': f"${total_potential:,.2f}",
                'Deals Registered': str(deal_registered),
                'Deals with Roadmap': str(has_roadmap),
                'Deals with Business Case': str(has_business_case),
                'Deals with Sponsor': str(has_sponsor)
            }
            visualizations.append(('Pipeline Metrics', pipeline_metrics))
            
        except Exception as e:
            st.error(f"Error creating pipeline visualization: {str(e)}")
    
    # 2. Risk Analysis
    if 'Project Risks' in data and not data['Project Risks'].empty:
        try:
            risk_df = data['Project Risks'].copy()
            
            # Convert Impact to numeric
            risk_df['Impact ($)'] = risk_df['Impact ($)'].astype(str).str.replace('$', '').str.replace(',', '')
            risk_df['Impact ($)'] = pd.to_numeric(risk_df['Impact ($)'], errors='coerce').fillna(0)
            # Fill missing values for Severity and Impact ($)
            risk_df['Severity'] = risk_df['Severity'].fillna('')
            risk_df['Impact ($)'] = risk_df['Impact ($)'].fillna(0)
            
            # Create risk distribution chart
            fig_risk = px.pie(
                risk_df,
                names='Severity',
                values='Impact ($)',
                title='Risk Distribution by Severity',
                color='Severity',
                color_discrete_map={
                    'High': '#ffcdd2',
                    'Medium': '#fff9c4',
                    'Low': '#c8e6c9'
                }
            )
            visualizations.append(('Risk Distribution', fig_risk))
            
            # Add risk metrics
            high_risk_impact = risk_df[risk_df['Severity'].str.lower() == 'high']['Impact ($)'].sum()
            total_risk_impact = risk_df['Impact ($)'].sum()
            
            risk_metrics = {
                'High Risk Impact': f"${high_risk_impact:,.2f}",
                'Total Risk Impact': f"${total_risk_impact:,.2f}",
                'High Risk %': f"{(high_risk_impact/total_risk_impact*100 if total_risk_impact > 0 else 0):.1f}%"
            }
            visualizations.append(('Risk Metrics', risk_metrics))
            
        except Exception as e:
            st.error(f"Error creating risk visualization: {str(e)}")
    
    # 3. Team Utilization Analysis
    if 'Team Utilization' in data and not data['Team Utilization'].empty:
        try:
            util_df = data['Team Utilization'].copy()
            
            # Convert Utilization to numeric
            util_df['Utilization (%)'] = util_df['Utilization (%)'].astype(str).str.replace('%', '')
            util_df['Utilization (%)'] = pd.to_numeric(util_df['Utilization (%)'], errors='coerce').fillna(0)
            
            # Split into executive and delivery teams
            exec_df = util_df[util_df['Role'].str.contains('Executive', case=False, na=False)]
            delivery_df = util_df[~util_df['Role'].str.contains('Executive', case=False, na=False)]
            
            # Create executive utilization chart (red is bad, green is good)
            if not exec_df.empty:
                fig_exec = px.bar(
                    exec_df.sort_values('Utilization (%)', ascending=False),
                    x='Employee Name',
                    y='Utilization (%)',
                    title='Executive Team Utilization',
                    color='Utilization (%)',
                    color_continuous_scale='RdYlGn_r'  # Reversed scale: red is high (bad), green is low (good)
                )
                fig_exec.update_layout(xaxis_tickangle=-45)
                visualizations.append(('Executive Utilization', fig_exec))
                
                # Add executive metrics
                avg_exec_util = exec_df['Utilization (%)'].mean()
                high_util_execs = len(exec_df[exec_df['Utilization (%)'] > 70])  # High utilization is concerning for execs
                
                exec_metrics = {
                    'Average Executive Utilization': f"{avg_exec_util:.1f}%",
                    'Executives Over 70% Utilized': str(high_util_execs),
                    'Opportunity Cost Risk': 'High' if high_util_execs > 0 else 'Low'
                }
                visualizations.append(('Executive Metrics', exec_metrics))
            
            # Create delivery team utilization chart (green is good, red is bad)
            if not delivery_df.empty:
                fig_delivery = px.bar(
                    delivery_df.sort_values('Utilization (%)', ascending=False),
                    x='Employee Name',
                    y='Utilization (%)',
                    title='Delivery Team Utilization',
                    color='Utilization (%)',
                    color_continuous_scale='RdYlGn'  # Normal scale: green is high (good), red is low (bad)
                )
                fig_delivery.update_layout(xaxis_tickangle=-45)
                visualizations.append(('Delivery Team Utilization', fig_delivery))
                
                # Add delivery metrics
                avg_delivery_util = delivery_df['Utilization (%)'].mean()
                under_utilized = len(delivery_df[delivery_df['Utilization (%)'] < 70])
                over_utilized = len(delivery_df[delivery_df['Utilization (%)'] > 100])
                
                delivery_metrics = {
                    'Average Delivery Utilization': f"{avg_delivery_util:.1f}%",
                    'Under-Utilized Team Members': str(under_utilized),
                    'Over-Utilized Team Members': str(over_utilized)
                }
                visualizations.append(('Delivery Metrics', delivery_metrics))
            
        except Exception as e:
            st.error(f"Error creating utilization visualization: {str(e)}")
    
    # 4. Project Status Analysis
    if 'Project Inventory' in data and not data['Project Inventory'].empty:
        try:
            project_df = data['Project Inventory'].copy()
            
            # Convert Revenue to numeric
            project_df['Revenue'] = project_df['Revenue'].astype(str).str.replace('$', '').str.replace(',', '')
            project_df['Revenue'] = pd.to_numeric(project_df['Revenue'], errors='coerce').fillna(0)
            
            # Create project status chart
            status_counts = project_df['Status (R/Y/G)'].value_counts()
            fig_status = px.pie(
                values=status_counts.values,
                names=status_counts.index,
                title='Project Status Distribution',
                color=status_counts.index,
                color_discrete_map={
                    'R': '#ffcdd2',
                    'Y': '#fff9c4',
                    'G': '#c8e6c9'
                }
            )
            visualizations.append(('Project Status', fig_status))
            
            # Add project metrics
            total_revenue = project_df['Revenue'].sum()
            at_risk_revenue = project_df[project_df['Status (R/Y/G)'] == 'R']['Revenue'].sum()
            
            project_metrics = {
                'Total Project Revenue': f"${total_revenue:,.2f}",
                'At-Risk Project Revenue': f"${at_risk_revenue:,.2f}",
                'At-Risk %': f"{(at_risk_revenue/total_revenue*100 if total_revenue > 0 else 0):.1f}%"
            }
            visualizations.append(('Project Metrics', project_metrics))
            
        except Exception as e:
            st.error(f"Error creating project visualization: {str(e)}")
    
    return visualizations

def render_homepage(data, openai_client):
    st.title("\U0001F3E5 Healthcare Delivery Dashboard")
    
    # Add CSS for metric cards
    st.markdown("""
        <style>
        .metric-card {
            background-color: #f0f2f6;
            border: 1px solid #e0e0e0;
            border-radius: 5px;
            padding: 15px;
            margin: 5px 0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        .metric-card.above-target {
            border-left: 4px solid #28a745;
        }
        .metric-card.below-target {
            border-left: 4px solid #dc3545;
        }
        .metric-title {
            color: #666;
            font-size: 0.9em;
            margin-bottom: 5px;
        }
        .metric-value {
            font-size: 1.5em;
            font-weight: bold;
            color: #262730;
        }
        .metric-delta {
            font-size: 0.9em;
            color: #666;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.subheader("Leading & Lagging Indicators")
    lagging = indicators.get_lagging_indicators(data)
    leading = indicators.get_leading_indicators(data)
    
    def format_metric_value(key, value):
        """Format metric values based on their type and key."""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return "N/A"
            
        if isinstance(value, (int, float)):
            # Format based on metric type
            if 'Revenue' in key and 'vs_' not in key:  # Only format actual revenue values as currency
                return f"${value:,.0f}"
            elif 'Satisfaction' in key and 'vs_' not in key:  # Raw satisfaction score
                return f"{value:.1f}"
            elif 'NPS' in key and 'vs_' not in key:  # Raw NPS score
                return f"{value:.1f}"
            elif any(x in key for x in ['Ratio', 'vs_Target', 'vs_Stretch']):
                return f"{value:.1f}%"
            elif 'Time' in key:
                return f"{value:.0f} days"
            elif 'Count' in key or 'Number' in key:
                return f"{value:,.0f}"
            else:
                return f"{value:,.2f}"
        elif isinstance(value, dict):
            if not value:
                return "N/A"
            # For dictionary values, show a summary
            if all(k in value for k in ['mean', 'median']):
                try:
                    mean_val = float(value['mean'])
                    median_val = float(value['median'])
                    return f"Avg: {mean_val:.0f}, Med: {median_val:.0f}"
                except (ValueError, TypeError):
                    return "N/A"
            return str(next(iter(value.values())))
        elif isinstance(value, list):
            if not value:
                return "0"
            return f"{len(value)} items"
        else:
            return str(value)
    
    def format_metric_delta(key, value):
        """Format metric delta values based on their type and key."""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
            
        if isinstance(value, (int, float)):
            if 'Revenue_vs_Target' in key:
                return f"{value:.1f}% of ${strategic_targets.REVENUE_TARGET:,.0f} target"
            elif 'Revenue_vs_Stretch' in key:
                return f"{value:.1f}% of ${strategic_targets.REVENUE_STRETCH_GOAL:,.0f} stretch goal"
            elif 'Customer NPS_vs_Target' in key:
                return f"{value:.1f}% of {strategic_targets.CUSTOMER_NPS_TARGET} target"
            elif 'Employee Satisfaction_vs_Target' in key:
                return f"{value:.1f}% of {strategic_targets.EMPLOYEE_PULSE_TARGET} target"
            elif 'Ratio' in key:
                return f"{value:.1f}× target"
            elif 'Count' in key or 'Number' in key:
                return f"{value:,.0f} total"
        return None
    
    def get_metric_class(key, value):
        """Determine if metric is above or below target."""
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return ""
            
        if isinstance(value, (int, float)):
            if 'Revenue_vs_Target' in key or 'Customer NPS_vs_Target' in key or 'Employee Satisfaction_vs_Target' in key:
                return "above-target" if value >= 100 else "below-target"
        return ""
    
    # Display Lagging Indicators
    st.markdown("#### Lagging Indicators")
    
    # Create three columns for lagging indicators
    lag_cols = st.columns(3)
    
    # Revenue Card
    with lag_cols[0]:
        revenue_value = format_metric_value('Revenue', lagging['Revenue'])
        revenue_delta = format_metric_delta('Revenue_vs_Target', lagging['Revenue_vs_Target'])
        revenue_class = get_metric_class('Revenue_vs_Target', lagging['Revenue_vs_Target'])
        st.markdown(f"""
            <div class="metric-card {revenue_class}">
                <div class="metric-title">Revenue</div>
                <div class="metric-value">{revenue_value}</div>
                <div class="metric-delta">{revenue_delta or ''}</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Customer NPS Card
    with lag_cols[1]:
        nps_value = format_metric_value('Customer NPS', lagging['Customer NPS'])
        nps_delta = format_metric_delta('Customer NPS_vs_Target', lagging['Customer NPS_vs_Target'])
        nps_class = get_metric_class('Customer NPS_vs_Target', lagging['Customer NPS_vs_Target'])
        st.markdown(f"""
            <div class="metric-card {nps_class}">
                <div class="metric-title">Customer NPS</div>
                <div class="metric-value">{nps_value}</div>
                <div class="metric-delta">{nps_delta or ''}</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Employee Satisfaction Card
    with lag_cols[2]:
        satisfaction_value = format_metric_value('Employee Satisfaction', lagging['Employee Satisfaction'])
        satisfaction_delta = format_metric_delta('Employee Satisfaction_vs_Target', lagging['Employee Satisfaction_vs_Target'])
        satisfaction_class = get_metric_class('Employee Satisfaction_vs_Target', lagging['Employee Satisfaction_vs_Target'])
        st.markdown(f"""
            <div class="metric-card {satisfaction_class}">
                <div class="metric-title">Employee Satisfaction</div>
                <div class="metric-value">{satisfaction_value}</div>
                <div class="metric-delta">{satisfaction_delta or ''}</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Display Leading Indicators
    st.markdown("#### Leading Indicators")
    
    # Create columns for leading indicators
    lead_cols = st.columns(3)
    
    # Pipeline Coverage Card
    with lead_cols[0]:
        pipeline_value = format_metric_value('Pipeline Coverage', leading['Pipeline Coverage'])
        pipeline_ratio = leading['Pipeline Coverage Ratio']
        pipeline_delta = f"{pipeline_ratio:.1f}× of ${strategic_targets.REVENUE_TARGET:,.0f} target"
        pipeline_class = get_metric_class('Pipeline Coverage Ratio', pipeline_ratio * 100)
        st.markdown(f"""
            <div class="metric-card {pipeline_class}">
                <div class="metric-title">Pipeline Coverage</div>
                <div class="metric-value">${leading['Pipeline Coverage']:,.0f}</div>
                <div class="metric-delta">{pipeline_delta}</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Deal Cycle Time Card
    with lead_cols[1]:
        cycle_time = leading.get('Avg Deal Cycle Time')
        cycle_time_value = format_metric_value('Avg Deal Cycle Time', cycle_time)
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Deal Cycle Time (Avg)</div>
                <div class="metric-value">{cycle_time_value}</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Next Deal Gap Card
    with lead_cols[2]:
        next_deal_gap = leading.get('Avg Next Deal Gap')
        next_deal_value = format_metric_value('Avg Next Deal Gap', next_deal_gap)
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Time Between Project End and Next Deal Discussion (Avg)</div>
                <div class="metric-value">{next_deal_value}</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Second row of leading indicators
    lead_cols2 = st.columns(2)
    
    # Sponsor Check-ins Card
    with lead_cols2[0]:
        checkins = leading.get('Recent Meaningful Check-ins', 0)
        checkins_pct = leading.get('Recent Meaningful Check-ins %', 0)
        checkins_value = f"{checkins} of {len(data['Project Inventory'])}"
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">Meaningful Sponsor Check-ins (Last 30 Days)</div>
                <div class="metric-value">{checkins_value}</div>
                <div class="metric-delta">{checkins_pct:.1f}%</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Green Project Ratio Card
    with lead_cols2[1]:
        green_ratio = leading.get('Green Project Ratio', 0)
        green_ratio_delta = f"{green_ratio * 100:.1f}% of {strategic_targets.GREEN_PROJECT_TARGET * 100}% target"
        green_class = get_metric_class('Green Project Ratio', green_ratio * 100)
        st.markdown(f"""
            <div class="metric-card {green_class}">
                <div class="metric-title">Green Project Ratio</div>
                <div class="metric-value">{green_ratio * 100:.1f}%</div>
                <div class="metric-delta">{green_ratio_delta}</div>
            </div>
        """, unsafe_allow_html=True)
    
    # Top 3 Action Items (AI)
    st.markdown("---")
    st.subheader("Top 3 Action Items for Today (AI-Powered)")
    # Only generate Top 3 items if they haven't been generated yet
    if 'top3_items' not in st.session_state:
        # Prepare data context (full data for OpenAI)
        data_context = "\n".join([
            f"{name}:\n" + df.to_string(index=False) for name, df in data.items() if not df.empty
        ])
        if openai_client:
            with st.spinner("Analyzing data and generating action items..."):
                st.session_state.top3_items = indicators.get_top3_action_items(data, openai_client, data_context)
        else:
            st.warning("OpenAI API key not configured.")
    # Display the stored Top 3 items
    if 'top3_items' in st.session_state:
        st.markdown(st.session_state.top3_items)

    st.subheader("Project & Pipeline Scoring Insights")
    lagging = indicators.get_lagging_indicators(data)
    leading = indicators.get_leading_indicators(data)
    # --- Project Health Score Metrics ---
    st.markdown("#### Project Health Score Metrics")
    phs_avg = lagging.get('Avg Project Health Score')
    phs_median = lagging.get('Median Project Health Score')
    phs_bands = lagging.get('Project Health Score Bands %')
    top_project = lagging.get('Top Project by Health Score')
    bottom_project = lagging.get('Bottom Project by Health Score')
    cols = st.columns(4)
    with cols[0]:
        st.metric("Avg Health Score", f"{phs_avg:.1f}" if phs_avg is not None else "N/A")
    with cols[1]:
        st.metric("Median Health Score", f"{phs_median:.1f}" if phs_median is not None else "N/A")
    with cols[2]:
        if phs_bands:
            st.write("Score Bands:")
            for band, pct in phs_bands.items():
                st.write(f"{band}: {pct:.1f}%")
    with cols[3]:
        st.write(f"Top: {top_project if top_project else 'N/A'}")
        st.write(f"Bottom: {bottom_project if bottom_project else 'N/A'}")
    # --- Total Project Score Metrics ---
    st.markdown("#### Total Project Score Metrics")
    tps_avg = lagging.get('Avg Total Project Score')
    tps_median = lagging.get('Median Total Project Score')
    tps_bands = lagging.get('Total Project Score Bands %')
    top_tps = lagging.get('Top Project by Total Score')
    bottom_tps = lagging.get('Bottom Project by Total Score')
    cols_tps = st.columns(4)
    with cols_tps[0]:
        st.metric("Avg Total Project Score", f"{tps_avg:.1f}" if tps_avg is not None else "N/A")
    with cols_tps[1]:
        st.metric("Median Total Project Score", f"{tps_median:.1f}" if tps_median is not None else "N/A")
    with cols_tps[2]:
        if tps_bands:
            st.write("Score Bands:")
            for band, pct in tps_bands.items():
                st.write(f"{band}: {pct:.1f}%")
    with cols_tps[3]:
        st.write(f"Top: {top_tps if top_tps else 'N/A'}")
        st.write(f"Bottom: {bottom_tps if bottom_tps else 'N/A'}")
    # --- Pipeline Score Metrics ---
    st.markdown("#### Pipeline Score Metrics")
    pls_avg = leading.get('Avg Pipeline Score')
    pls_median = leading.get('Median Pipeline Score')
    pls_bands = leading.get('Pipeline Score Bands %')
    top_pipeline = leading.get('Top Pipeline by Score')
    bottom_pipeline = leading.get('Bottom Pipeline by Score')
    cols2 = st.columns(4)
    with cols2[0]:
        st.metric("Avg Pipeline Score", f"{pls_avg:.1f}" if pls_avg is not None else "N/A")
    with cols2[1]:
        st.metric("Median Pipeline Score", f"{pls_median:.1f}" if pls_median is not None else "N/A")
    with cols2[2]:
        if pls_bands:
            st.write("Score Bands:")
            for band, pct in pls_bands.items():
                st.write(f"{band}: {pct:.1f}%")
    with cols2[3]:
        st.write(f"Top: {top_pipeline if top_pipeline else 'N/A'}")
        st.write(f"Bottom: {bottom_pipeline if bottom_pipeline else 'N/A'}")
    # --- Total Deal Score Metrics ---
    st.markdown("#### Total Deal Score Metrics")
    tds_avg = leading.get('Avg Total Deal Score')
    tds_median = leading.get('Median Total Deal Score')
    tds_bands = leading.get('Total Deal Score Bands %')
    top_tds = leading.get('Top Pipeline by Total Score')
    bottom_tds = leading.get('Bottom Pipeline by Total Score')
    cols_tds = st.columns(4)
    with cols_tds[0]:
        st.metric("Avg Total Deal Score", f"{tds_avg:.1f}" if tds_avg is not None else "N/A")
    with cols_tds[1]:
        st.metric("Median Total Deal Score", f"{tds_median:.1f}" if tds_median is not None else "N/A")
    with cols_tds[2]:
        if tds_bands:
            st.write("Score Bands:")
            for band, pct in tds_bands.items():
                st.write(f"{band}: {pct:.1f}%")
    with cols_tds[3]:
        st.write(f"Top: {top_tds if top_tds else 'N/A'}")
        st.write(f"Bottom: {bottom_tds if bottom_tds else 'N/A'}")
    # --- Score Distribution Visualizations ---
    st.markdown("#### Score Distributions")
    project_df = data.get('Project Inventory')
    pipeline_df = data.get('Pipeline')
    if project_df is not None and not project_df.empty and 'Project Health Score' in project_df.columns:
        scores = pd.to_numeric(project_df['Project Health Score'], errors='coerce').dropna()
        fig = px.histogram(scores, nbins=10, title="Project Health Score Distribution", labels={'value': 'Health Score'})
        st.plotly_chart(fig, use_container_width=True)
    if project_df is not None and not project_df.empty and 'Total Project Score' in project_df.columns:
        scores = pd.to_numeric(project_df['Total Project Score'], errors='coerce').dropna()
        fig = px.histogram(scores, nbins=10, title="Total Project Score Distribution", labels={'value': 'Total Project Score'})
        st.plotly_chart(fig, use_container_width=True)
    if pipeline_df is not None and not pipeline_df.empty and 'Pipeline Score' in pipeline_df.columns:
        scores = pd.to_numeric(pipeline_df['Pipeline Score'], errors='coerce').dropna()
        fig = px.histogram(scores, nbins=10, title="Pipeline Score Distribution", labels={'value': 'Pipeline Score'})
        st.plotly_chart(fig, use_container_width=True)
    if pipeline_df is not None and not pipeline_df.empty and 'Total Deal Score' in pipeline_df.columns:
        scores = pd.to_numeric(pipeline_df['Total Deal Score'], errors='coerce').dropna()
        fig = px.histogram(scores, nbins=10, title="Total Deal Score Distribution", labels={'value': 'Total Deal Score'})
        st.plotly_chart(fig, use_container_width=True)

def render_strategic_targets_section(data):
    """Render the strategic targets and leading indicators section."""
    st.header("🎯 Strategic Targets & Leading Indicators")
    
    # Get indicators
    lagging_indicators = indicators.get_lagging_indicators(data)
    leading_indicators = indicators.get_leading_indicators(data)
    
    # Create three columns for the summary tiles
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Revenue Progress
        if lagging_indicators['Revenue'] is not None:
            st.metric(
                "Revenue Progress",
                f"${lagging_indicators['Revenue']:,.0f}",
                f"{lagging_indicators['Revenue_vs_Target']:.1f}% of Target"
            )
        
        # Pipeline Coverage
        if leading_indicators['Pipeline Coverage'] is not None:
            st.metric(
                "Pipeline Coverage",
                f"${leading_indicators['Pipeline Coverage']:,.0f}",
                f"{leading_indicators['Pipeline Coverage Ratio']:.1f}× Target"
            )
    
    with col2:
        # Green Project Ratio
        if leading_indicators['Green Project Ratio'] is not None:
            st.metric(
                "Green Projects",
                f"{leading_indicators['Green Project Ratio']*100:.1f}%",
                f"{leading_indicators['Green Project Ratio_vs_Target']:.1f}% of Target"
            )
        
        # Deal Cycle Time
        if leading_indicators['Avg Deal Cycle Time'] is not None:
            st.metric(
                "Avg Deal Cycle Time",
                f"{leading_indicators['Avg Deal Cycle Time']:.0f} days",
                f"Median: {leading_indicators['Median Deal Cycle Time']:.0f} days"
            )
    
    with col3:
        # eNPS Score
        if lagging_indicators['Customer NPS'] is not None:
            st.metric(
                "Customer eNPS",
                f"{lagging_indicators['Customer NPS']:.1f}",
                f"{lagging_indicators['Customer NPS_vs_Target']:.1f}% of Target"
            )
        
        # Sponsor Check-ins
        if leading_indicators['Recent Meaningful Check-ins %'] is not None:
            st.metric(
                "Recent Sponsor Check-ins",
                f"{leading_indicators['Recent Meaningful Check-ins']} of {len(data['Project Inventory'])}",
                f"{leading_indicators['Recent Meaningful Check-ins %']:.1f}%"
            )
    
    # Detailed Metrics Section
    st.subheader("📊 Detailed Metrics")
    
    # Create tabs for different metric categories
    tab1, tab2, tab3 = st.tabs(["Pipeline & Deals", "Project Health", "Action Items"])
    
    with tab1:
        # Pipeline and Deal Metrics
        st.write("#### Pipeline Coverage")
        if leading_indicators['Pipeline Coverage'] is not None:
            st.write(f"Total Pipeline: ${leading_indicators['Pipeline Coverage']:,.0f}")
            st.write(f"Coverage Ratio: {leading_indicators['Pipeline Coverage Ratio']:.1f}× Target")
        
        st.write("#### Deal Cycle Analysis")
        if leading_indicators['Deal Cycle Time by Tier'] is not None:
            st.write("Deal Cycle Time by Pursuit Tier:")
            for tier, times in leading_indicators['Deal Cycle Time by Tier'].items():
                st.write(f"- {tier}: Avg {times.get('mean', 0):.0f} days, Median {times.get('median', 0):.0f} days")
        
        st.write("#### Next Deal Discussions")
        if leading_indicators['Overdue Next Deal Projects'] is not None:
            st.write("Projects Overdue for Next Deal Discussion:")
            for project in leading_indicators['Overdue Next Deal Projects']:
                st.write(f"- {project['Project Name']}: {project['Next Deal Gap']} days since project end")
    
    with tab2:
        # Project Health Metrics
        st.write("#### Project Status")
        if leading_indicators['Non-Green Projects'] is not None:
            st.write("Projects Not Meeting Green Status:")
            for project in leading_indicators['Non-Green Projects']:
                st.write(f"- {project['Project Name']} ({project['Status (R/Y/G)']}): {project['Key Issues']}")
        
        st.write("#### Sponsor Check-ins")
        if leading_indicators['Overdue Check-in Projects'] is not None:
            st.write("Projects Overdue for Sponsor Check-in:")
            for project in leading_indicators['Overdue Check-in Projects']:
                st.write(f"- {project['Project Name']}: Last check-in {project['Last Sponsor Checkin Date']}")
    
    with tab3:
        # Action Items
        st.write("#### Top Action Items")
        action_items = indicators.get_top3_action_items(data, st.session_state.openai_client, str(data))
        st.write(action_items)

def render_dashboard():
    """Render the main dashboard."""
    # Add CSS for metric boxes
    st.markdown("""
        <style>
        .metric-box {
            background-color: #f0f2f6;
            padding: 20px;
            border-radius: 5px;
            margin: 10px 0;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Load data
    data = st.session_state.data
    if not data:
        st.error("Failed to load data. Please check your connection and try again.")
        return
    
    # Render strategic targets section
    render_strategic_targets_section(data)
    
    # Render other dashboard sections
    render_homepage(data, st.session_state.openai_client)

def escape_markdown(text):
    # Escape underscores and asterisks
    text = re.sub(r'([_*])', r'\\\1', text)
    return text

def answer_critical_question(user_question, data):
    """Answer critical questions directly without using OpenAI"""
    q = user_question.lower().strip()
    project_df = data.get('Project Inventory')
    pipeline_df = data.get('Pipeline')
    risk_df = data.get('Project Risks')
    util_df = data.get('Team Utilization')
    exec_df = data.get('Executive Activity')
    
    # Clean up dataframes as needed
    if project_df is not None and not project_df.empty:
        project_df = project_df.copy()
        # Clean and convert Revenue column properly
        project_df['Revenue'] = project_df['Revenue'].astype(str).str.replace('$', '').str.replace(',', '').str.strip()
        project_df['Revenue'] = pd.to_numeric(project_df['Revenue'], errors='coerce').fillna(0)
        project_df['Margin'] = project_df['Margin'].astype(str).str.replace('%', '').str.strip()
        project_df['Margin'] = pd.to_numeric(project_df['Margin'], errors='coerce').fillna(0)
    
    if pipeline_df is not None and not pipeline_df.empty:
        pipeline_df = pipeline_df.copy()
        pipeline_df['Percieved Annual AMO'] = pipeline_df['Percieved Annual AMO'].astype(str).str.replace('$', '').str.replace(',', '').str.strip()
        pipeline_df['Percieved Annual AMO'] = pd.to_numeric(pipeline_df['Percieved Annual AMO'], errors='coerce').fillna(0)
    
    if risk_df is not None and not risk_df.empty:
        risk_df = risk_df.copy()
        risk_df['Impact ($)'] = risk_df['Impact ($)'].astype(str).str.replace('$', '').str.replace(',', '')
        risk_df['Impact ($)'] = pd.to_numeric(risk_df['Impact ($)'], errors='coerce').fillna(0)
        risk_df['Severity'] = risk_df['Severity'].fillna('')
    
    if util_df is not None and not util_df.empty:
        util_df = util_df.copy()
        util_df['Utilization (%)'] = util_df['Utilization (%)'].astype(str).str.replace('%', '')
        util_df['Utilization (%)'] = pd.to_numeric(util_df['Utilization (%)'], errors='coerce').fillna(0)
    
    if exec_df is not None and not exec_df.empty:
        exec_df = exec_df.copy()
        exec_df['Strategic Cost ($)'] = exec_df['Strategic Cost ($)'].astype(str).str.replace('$', '').str.replace(',', '')
        exec_df['Strategic Cost ($)'] = pd.to_numeric(exec_df['Strategic Cost ($)'], errors='coerce').fillna(0)

    # 1. Most valuable project
    if re.search(r"most (valuable|revenue|expensive) project", q):
        if project_df is not None and not project_df.empty:
            # Sort by Revenue in descending order to verify
            sorted_df = project_df.sort_values('Revenue', ascending=False)
            top_project = sorted_df.iloc[0]  # Get the first row after sorting
            return f'The project with the most revenue is "{top_project["Project Name"]}" with a revenue of ${top_project["Revenue"]:,.2f}.'
    
    # 2. Least valuable project
    if re.search(r"least (valuable|revenue|expensive) project", q):
        if project_df is not None and not project_df.empty:
            low_project = project_df.loc[project_df['Revenue'].idxmin()]
            return f'The project with the least revenue is "{low_project["Project Name"]}" with a revenue of ${low_project["Revenue"]:,.2f}.'
    
    # 3. Total revenue
    if re.search(r"total revenue", q):
        if project_df is not None and not project_df.empty:
            return f'The total revenue is ${project_df["Revenue"].sum():,.2f}.'
    
    # 4. Average project revenue
    if re.search(r"average (project )?revenue", q):
        if project_df is not None and not project_df.empty:
            return f'The average project revenue is ${project_df["Revenue"].mean():,.2f}.'
    
    # 5. How many projects
    if re.search(r"how many projects", q):
        if project_df is not None and not project_df.empty:
            return f'There are {len(project_df)} projects.'
    
    # 6. How many red projects
    if re.search(r"how many red projects", q):
        if project_df is not None and not project_df.empty:
            red_count = (project_df['Status (R/Y/G)'].str.strip().str.lower() == 'red').sum()
            return f'There are {red_count} red projects.'
    
    # 7. Total revenue of red projects
    if re.search(r"red project revenue|revenue of red projects", q):
        if project_df is not None and not project_df.empty:
            red_revenue = project_df.loc[project_df['Status (R/Y/G)'].str.strip().str.lower() == 'red', 'Revenue'].sum()
            return f'The total revenue of red projects is ${red_revenue:,.2f}.'
    
    # 8. Average margin
    if re.search(r"average margin", q):
        if project_df is not None and not project_df.empty:
            return f'The average project margin is {project_df["Margin"].mean():.2f}%.'
    
    # 9. Highest margin project
    if re.search(r"highest margin project", q):
        if project_df is not None and not project_df.empty:
            top_margin = project_df.loc[project_df['Margin'].idxmax()]
            return f'The project with the highest margin is "{top_margin["Project Name"]}" with a margin of {top_margin["Margin"]:.2f}%.'
    
    # 10. Lowest margin project
    if re.search(r"lowest margin project", q):
        if project_df is not None and not project_df.empty:
            low_margin = project_df.loc[project_df['Margin'].idxmin()]
            return f'The project with the lowest margin is "{low_margin["Project Name"]}" with a margin of {low_margin["Margin"]:.2f}%.'
    
    # 11. Total pipeline value
    if re.search(r"total pipeline", q):
        if pipeline_df is not None and not pipeline_df.empty:
            return f'The total pipeline value is ${pipeline_df["Percieved Annual AMO"].sum():,.2f}.'
    
    # 12. Pipeline health metrics
    if re.search(r"pipeline health|deal registration|roadmap|business case|sponsor", q):
        if pipeline_df is not None and not pipeline_df.empty:
            deal_registered = len(pipeline_df[pipeline_df['Deal Registered YN'].str.lower() == 'y'])
            has_roadmap = len(pipeline_df[pipeline_df['Agreed Upon Roadmap YN'].str.lower() == 'y'])
            has_business_case = len(pipeline_df[pipeline_df['Business Case_ROI YN'].str.lower() == 'y'])
            has_sponsor = len(pipeline_df[pipeline_df['Business Sponsor YN'].str.lower() == 'y'])
            return f'Pipeline Health:\n- {deal_registered} deals registered\n- {has_roadmap} deals with roadmap\n- {has_business_case} deals with business case\n- {has_sponsor} deals with sponsor'
    
    # 13. Total at-risk revenue
    if re.search(r"at[- ]?risk revenue", q):
        if risk_df is not None and not risk_df.empty:
            at_risk = risk_df.loc[risk_df['Severity'].str.lower() == 'high', 'Impact ($)'].sum()
            return f'The total at-risk revenue is ${at_risk:,.2f}.'
    
    # 14. How many high-risk items
    if re.search(r"how many high[- ]?risk", q):
        if risk_df is not None and not risk_df.empty:
            high_risk_count = (risk_df['Severity'].str.lower() == 'high').sum()
            return f'There are {high_risk_count} high-risk items.'
    
    # 15. Total strategic cost
    if re.search(r"strategic cost", q):
        if exec_df is not None and not exec_df.empty:
            return f'The total strategic cost is ${exec_df["Strategic Cost ($)"].sum():,.2f}.'
    
    # 16. Average executive utilization
    if re.search(r"average executive utilization", q):
        if util_df is not None and not util_df.empty:
            exec_mask = util_df['Role'].str.contains('Executive', case=False, na=False)
            avg_exec = util_df.loc[exec_mask, 'Utilization (%)'].mean()
            return f'The average executive utilization is {avg_exec:.1f}%.'
    
    # 17. Average delivery utilization
    if re.search(r"average delivery utilization", q):
        if util_df is not None and not util_df.empty:
            deliv_mask = ~util_df['Role'].str.contains('Executive', case=False, na=False)
            avg_deliv = util_df.loc[deliv_mask, 'Utilization (%)'].mean()
            return f'The average delivery team utilization is {avg_deliv:.1f}%.'
    
    # 18. Over-utilized execs
    if re.search(r"over[- ]?utilized exec", q):
        if util_df is not None and not util_df.empty:
            exec_mask = util_df['Role'].str.contains('Executive', case=False, na=False)
            over_util = (util_df.loc[exec_mask, 'Utilization (%)'] > 70).sum()
            return f'There are {over_util} over-utilized executives (over 70% utilized).'
    
    # 19. Under-utilized delivery
    if re.search(r"under[- ]?utilized delivery", q):
        if util_df is not None and not util_df.empty:
            deliv_mask = ~util_df['Role'].str.contains('Executive', case=False, na=False)
            under_util = (util_df.loc[deliv_mask, 'Utilization (%)'] < 70).sum()
            return f'There are {under_util} under-utilized delivery team members (under 70% utilized).'
    
    # 20. Total risk impact
    if re.search(r"total risk impact", q):
        if risk_df is not None and not risk_df.empty:
            return f'The total risk impact is ${risk_df["Impact ($)"].sum():,.2f}.'
    
    return None

def render_ai_assistant(data):
    """Render the AI assistant in the sidebar"""
    st.sidebar.title("🤖 AI Assistant")
    st.sidebar.write("Ask questions about the data in natural language")
    # Initialize session state for the question if it doesn't exist
    if 'ai_question' not in st.session_state:
        st.session_state.ai_question = ""
    # Create a form for the question input and send button
    with st.sidebar.form(key="ai_assistant_form"):
        user_question = st.text_input("Your question:", value=st.session_state.ai_question)
        send_button = st.form_submit_button("Send")
    # Only process the question if the send button was clicked
    if send_button and user_question:
        # Store the question in session state
        st.session_state.ai_question = user_question
        # Prepare data context (full data for OpenAI)
        data_context = "\n".join([
            f"{name}:\n" + df.to_string(index=False) for name, df in data.items() if not df.empty
        ])
        # Intercept and answer critical questions
        py_answer = answer_critical_question(user_question, data)
        if py_answer:
            st.sidebar.success(py_answer)
        else:
            with st.spinner("Thinking..."):
                response = query_openai(user_question, data_context)
                st.sidebar.markdown(escape_markdown(response), unsafe_allow_html=False)

def main():
    # Initialize session state
    if 'data' not in st.session_state:
        st.session_state.data = {}
    if 'metrics' not in st.session_state:
        st.session_state.metrics = {}
    # Initialize Google Sheets connection
    sheet = setup_google_sheets()
    if not sheet:
        st.error("Failed to connect to Google Sheets. Please check your credentials.")
        return
    # Load data for all worksheets
    worksheet_names = [
        'Project Inventory', 'Project Risks', 'Pipeline', 'Team Utilization',
        'Talent Gaps', 'Operational Gaps', 'Executive Activity',
        'Scenario Model Inputs', 'Do Nothing Scenario', 'Proposed Scenario',
        'Scenario Comparison'
    ]
    if not st.session_state.data or all(df.empty for df in st.session_state.data.values()):
        for name in worksheet_names:
            st.session_state.data[name] = load_sheet_data(sheet, name)
        if any(not df.empty for df in st.session_state.data.values()):
            st.session_state.metrics = calculate_dashboard_metrics(st.session_state.data)
    
    # --- NAVIGATION BAR ---
    nav_cols = st.columns([8, 1, 1, 1, 1])
    with nav_cols[1]:
        if st.button("🏠 Home"):
            st.session_state.current_page = "homepage"
    with nav_cols[2]:
        if st.button("📊 Data Views"):
            st.session_state.current_page = "data_views"
    with nav_cols[3]:
        if st.button("📈 Analytics"):
            st.session_state.current_page = "analytics"
    with nav_cols[4]:
        if st.button("📋 Metrics"):
            st.session_state.current_page = "metrics"
    
    # Default to homepage
    if 'current_page' not in st.session_state:
        st.session_state.current_page = "homepage"
    
    # --- PAGE RENDERING ---
    if st.session_state.current_page == "homepage":
        openai_api_key = os.getenv('OPENAI_API_KEY')
        openai_client = None
        if openai_api_key:
            import openai
            openai_client = openai.OpenAI(api_key=openai_api_key)
        render_homepage(st.session_state.data, openai_client)
    elif st.session_state.current_page == "data_views":
        st.title("Data Views")
        selected_view = st.selectbox(
            "Select View",
            worksheet_names
        )
        if selected_view in st.session_state.data:
            df = st.session_state.data[selected_view]
            if selected_view == 'Project Inventory' and 'Status (R/Y/G)' in df.columns:
                # Convert Revenue to numeric for sorting, then format as currency for display
                df = df.copy()
                if 'Revenue' in df.columns:
                    df['Revenue'] = df['Revenue'].astype(str).str.replace('$', '').str.replace(',', '').str.strip()
                    df['Revenue'] = pd.to_numeric(df['Revenue'], errors='coerce').fillna(0)
                    df['Revenue'] = df['Revenue'].map(lambda x: f"${x:,.2f}")
                def color_status(val):
                    colors = {
                        'Red': 'background-color: #ffcdd2',
                        'Yellow': 'background-color: #fff9c4',
                        'Green': 'background-color: #c8e6c9'
                    }
                    return colors.get(val, '')
                styled_df = df.style.map(color_status, subset=['Status (R/Y/G)'])
                st.dataframe(styled_df, use_container_width=True)
            else:
                st.dataframe(df, use_container_width=True)
    elif st.session_state.current_page == "analytics":
        st.title("Analytics")
        visualizations = create_analytics_visualizations(st.session_state.data)
        for i in range(0, len(visualizations), 2):
            cols = st.columns(2)
            for j in range(2):
                if i + j < len(visualizations):
                    title, content = visualizations[i + j]
                    with cols[j]:
                        st.subheader(title)
                        if isinstance(content, dict):
                            for metric_name, metric_value in content.items():
                                st.metric(metric_name, metric_value)
                        else:
                            st.plotly_chart(content, use_container_width=True)
    elif st.session_state.current_page == "metrics":
        st.title("Detailed Metrics")
        render_dashboard()
    
    # Render AI Assistant in sidebar
    render_ai_assistant(st.session_state.data)

if __name__ == "__main__":
    main() 