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
            project_df['Revenue'] = project_df['Revenue'].astype(str).str.replace('$', '').str.replace(',', '')
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
            pipeline_df['Revenue Potential'] = pipeline_df['Revenue Potential'].astype(str).str.replace('$', '').str.replace(',', '')
            pipeline_df['Revenue Potential'] = pd.to_numeric(pipeline_df['Revenue Potential'], errors='coerce').fillna(0)
            pipeline_df['Probability (%)'] = pipeline_df['Probability (%)'].astype(str).str.replace('%', '')
            pipeline_df['Probability (%)'] = pd.to_numeric(pipeline_df['Probability (%)'], errors='coerce').fillna(0)
            
            pipeline_df['Weighted Revenue'] = pipeline_df['Revenue Potential'] * (pipeline_df['Probability (%)'] / 100)
            metrics['total_pipeline'] = float(pipeline_df['Weighted Revenue'].sum())
            metrics['total_potential'] = float(pipeline_df['Revenue Potential'].sum())
            metrics['avg_probability'] = float(pipeline_df['Probability (%)'].mean())
        except Exception as e:
            st.error(f"Error calculating pipeline metrics: {str(e)}")
    
    # Risk Metrics
    if 'Project Risks' in data and not data['Project Risks'].empty:
        try:
            risk_df = data['Project Risks'].copy()
            risk_df['Impact ($)'] = risk_df['Impact ($)'].astype(str).str.replace('$', '').str.replace(',', '')
            risk_df['Impact ($)'] = pd.to_numeric(risk_df['Impact ($)'], errors='coerce').fillna(0)
            
            metrics['at_risk_revenue'] = float(risk_df[risk_df['Severity'].str.lower() == 'high']['Impact ($)'].sum())
            metrics['total_risk_impact'] = float(risk_df['Impact ($)'].sum())
            metrics['high_risk_count'] = len(risk_df[risk_df['Severity'].str.lower() == 'high'])
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
            model="gpt-4-turbo-preview",
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
            # Create weighted revenue by probability
            pipeline_df = data['Pipeline'].copy()
            
            # Convert columns to numeric, handling string values
            pipeline_df['Revenue Potential'] = pipeline_df['Revenue Potential'].astype(str).str.replace('$', '').str.replace(',', '')
            pipeline_df['Revenue Potential'] = pd.to_numeric(pipeline_df['Revenue Potential'], errors='coerce').fillna(0)
            
            pipeline_df['Probability (%)'] = pipeline_df['Probability (%)'].astype(str).str.replace('%', '')
            pipeline_df['Probability (%)'] = pd.to_numeric(pipeline_df['Probability (%)'], errors='coerce').fillna(0)
            
            pipeline_df['Weighted Revenue'] = pipeline_df['Revenue Potential'] * (pipeline_df['Probability (%)'] / 100)
            
            # Sort by weighted revenue
            pipeline_df = pipeline_df.sort_values('Weighted Revenue', ascending=False)
            
            # Create pipeline chart
            fig_pipeline = px.bar(
                pipeline_df.head(10),
                x='Opportunity Name',
                y='Weighted Revenue',
                title='Top 10 Pipeline Opportunities by Weighted Revenue',
                labels={'Weighted Revenue': 'Weighted Revenue ($)', 'Opportunity Name': 'Opportunity'},
                color='Probability (%)',
                color_continuous_scale='RdYlGn'
            )
            fig_pipeline.update_layout(xaxis_tickangle=-45)
            visualizations.append(('Pipeline Analysis', fig_pipeline))
            
            # Add pipeline metrics
            total_potential = pipeline_df['Revenue Potential'].sum()
            total_weighted = pipeline_df['Weighted Revenue'].sum()
            avg_probability = pipeline_df['Probability (%)'].mean()
            
            pipeline_metrics = {
                'Total Potential Revenue': f"${total_potential:,.2f}",
                'Total Weighted Revenue': f"${total_weighted:,.2f}",
                'Average Probability': f"{avg_probability:.1f}%"
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
            
            # Create risk distribution chart
            fig_risk = px.pie(
                risk_df,
                names='Severity',  # Changed from 'Severity (High/Medium/Low)' to 'Severity'
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

def render_dashboard():
    # Add CSS for metric boxes
    st.markdown("""
        <style>
        .metric-box {
            background: #23272f;
            border-radius: 12px;
            padding: 1.5em 1em 1em 1em;
            margin-bottom: 1em;
            box-shadow: 0 2px 8px rgba(0,0,0,0.07);
            border: 1px solid #33353a;
            min-height: 110px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: flex-start;
        }
        .metric-title {
            font-size: 1.1em;
            color: #b0b8c1;
            margin-bottom: 0.2em;
            font-weight: 600;
        }
        .metric-value {
            font-size: 2em;
            color: #fff;
            font-weight: 700;
        }
        .metric-sub {
            font-size: 0.95em;
            color: #b0b8c1;
            margin-top: 0.2em;
        }
        </style>
    """, unsafe_allow_html=True)
    current_fy = get_fiscal_year()
    st.title("🏥 Healthcare Delivery Dashboard")
    st.subheader(f"Fiscal Year {current_fy}-{current_fy+1} Dashboard")
    # Revenue Section
    st.markdown("### 📊 Revenue Overview")
    rev_col1, rev_col2, rev_col3, rev_col4 = st.columns(4)
    with rev_col1:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Total Revenue</div>
                <div class="metric-value">${st.session_state.metrics.get('total_revenue', 0.0):,.2f}</div>
                <div class="metric-sub">Total revenue from all projects</div>
            </div>
        ''', unsafe_allow_html=True)
    with rev_col2:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Total Pipeline</div>
                <div class="metric-value">${st.session_state.metrics.get('total_pipeline', 0.0):,.2f}</div>
                <div class="metric-sub">Weighted pipeline revenue</div>
            </div>
        ''', unsafe_allow_html=True)
    with rev_col3:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">At-Risk Revenue</div>
                <div class="metric-value">${st.session_state.metrics.get('at_risk_revenue', 0.0):,.2f}</div>
                <div class="metric-sub">Revenue at risk from high-severity risks</div>
            </div>
        ''', unsafe_allow_html=True)
    with rev_col4:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Strategic Cost</div>
                <div class="metric-value">${st.session_state.metrics.get('strategic_cost', 0.0):,.2f}</div>
                <div class="metric-sub">Total strategic opportunity cost</div>
            </div>
        ''', unsafe_allow_html=True)
    # Project Health Section
    st.markdown("### 🏥 Project Health")
    health_col1, health_col2, health_col3, health_col4 = st.columns(4)
    with health_col1:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Total Projects</div>
                <div class="metric-value">{st.session_state.metrics.get('total_projects', 0)}</div>
                <div class="metric-sub">Total number of active projects</div>
            </div>
        ''', unsafe_allow_html=True)
    with health_col2:
        red_projects = st.session_state.metrics.get('red_projects', 0)
        percent_red = red_projects / st.session_state.metrics.get('total_projects', 1) * 100
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Red Projects</div>
                <div class="metric-value">{red_projects}</div>
                <div class="metric-sub">{percent_red:.1f}% of total</div>
            </div>
        ''', unsafe_allow_html=True)
    with health_col3:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Red Project Revenue</div>
                <div class="metric-value">${st.session_state.metrics.get('red_project_revenue', 0.0):,.2f}</div>
                <div class="metric-sub">Revenue from projects with red status</div>
            </div>
        ''', unsafe_allow_html=True)
    with health_col4:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">High Risk Items</div>
                <div class="metric-value">{st.session_state.metrics.get('high_risk_count', 0)}</div>
                <div class="metric-sub">Number of high-severity risks</div>
            </div>
        ''', unsafe_allow_html=True)
    # Team Utilization Section
    st.markdown("### 👥 Team Utilization")
    util_col1, util_col2, util_col3, util_col4 = st.columns(4)
    with util_col1:
        exec_util = st.session_state.metrics.get('exec_utilization', 0.0)
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Executive Utilization</div>
                <div class="metric-value">{exec_util:.1f}%</div>
                <div class="metric-sub">Average executive utilization (lower is better)</div>
            </div>
        ''', unsafe_allow_html=True)
    with util_col2:
        delivery_util = st.session_state.metrics.get('delivery_utilization', 0.0)
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Delivery Utilization</div>
                <div class="metric-value">{delivery_util:.1f}%</div>
                <div class="metric-sub">Average delivery team utilization (higher is better)</div>
            </div>
        ''', unsafe_allow_html=True)
    with util_col3:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Over-Utilized Execs</div>
                <div class="metric-value">{st.session_state.metrics.get('over_utilized_execs', 0)}</div>
                <div class="metric-sub">Number of executives over 70% utilized</div>
            </div>
        ''', unsafe_allow_html=True)
    with util_col4:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Under-Utilized Delivery</div>
                <div class="metric-value">{st.session_state.metrics.get('under_utilized_delivery', 0)}</div>
                <div class="metric-sub">Number of delivery team members under 70% utilized</div>
            </div>
        ''', unsafe_allow_html=True)
    # Strategic Overview Section
    st.markdown("### 🎯 Strategic Overview")
    strat_col1, strat_col2, strat_col3, strat_col4 = st.columns(4)
    with strat_col1:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Strategic Activities</div>
                <div class="metric-value">{st.session_state.metrics.get('strategic_activities', 0)}</div>
                <div class="metric-sub">Number of strategic activities</div>
            </div>
        ''', unsafe_allow_html=True)
    with strat_col2:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Strategic Cost</div>
                <div class="metric-value">${st.session_state.metrics.get('strategic_cost', 0.0):,.2f}</div>
                <div class="metric-sub">Total strategic opportunity cost</div>
            </div>
        ''', unsafe_allow_html=True)
    with strat_col3:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Pipeline Probability</div>
                <div class="metric-value">{st.session_state.metrics.get('avg_probability', 0.0):.1f}%</div>
                <div class="metric-sub">Average pipeline probability</div>
            </div>
        ''', unsafe_allow_html=True)
    with strat_col4:
        st.markdown(f'''
            <div class="metric-box">
                <div class="metric-title">Total Risk Impact</div>
                <div class="metric-value">${st.session_state.metrics.get('total_risk_impact', 0.0):,.2f}</div>
                <div class="metric-sub">Total impact of all risks</div>
            </div>
        ''', unsafe_allow_html=True)

def escape_markdown(text):
    # Escape underscores and asterisks
    text = re.sub(r'([_*])', r'\\\1', text)
    return text

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
        if st.button("🏠 Dashboard"):
            st.session_state.current_page = "dashboard"
    with nav_cols[2]:
        if st.button("📊 Data Views"):
            st.session_state.current_page = "data_views"
    with nav_cols[3]:
        if st.button("📈 Analytics"):
            st.session_state.current_page = "analytics"
    with nav_cols[4]:
        if st.button("🤖 AI Assistant"):
            st.session_state.current_page = "ai_assistant"
    # Default to dashboard
    if 'current_page' not in st.session_state:
        st.session_state.current_page = "dashboard"
    # --- PAGE RENDERING ---
    if st.session_state.current_page == "dashboard":
        render_dashboard()
    elif st.session_state.current_page == "data_views":
        st.title("Data Views")
        selected_view = st.selectbox(
            "Select View",
            worksheet_names
        )
        if selected_view in st.session_state.data:
            df = st.session_state.data[selected_view]
            if selected_view == 'Project Inventory' and 'Status (R/Y/G)' in df.columns:
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
    elif st.session_state.current_page == "ai_assistant":
        st.title("AI Assistant")
        st.write("Ask questions about the data in natural language")
        data_context = "\n".join([
            f"{name}:\n{df.head().to_string()}"
            for name, df in st.session_state.data.items()
            if not df.empty
        ])
        user_question = st.text_input("Your question:")
        if user_question:
            with st.spinner("Thinking..."):
                response = query_openai(user_question, data_context)
                st.markdown(escape_markdown(response), unsafe_allow_html=False)

if __name__ == "__main__":
    main() 