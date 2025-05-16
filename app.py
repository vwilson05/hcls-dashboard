import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os
from dotenv import load_dotenv
import openai
from datetime import datetime, date # Import date for st.date_input
import plotly.express as px
import plotly.graph_objects as go
import time
import re
import indicators # Our new indicators module
import strategic_targets # For referencing targets in display

# Load environment variables
load_dotenv()

# --- Page Configuration ---
st.set_page_config(
    page_title="Healthcare Delivery OS",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Global Styling ---
st.markdown("""
    <style>
        /* General improvements */
        .stApp {
            /* background-color: #f0f2f5; */ 
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 24px; 
        }
        .stTabs [data-baseweb="tab"] { 
            height: 44px;
            white-space: pre-wrap;
            background-color: #f0f2f6; 
            border-radius: 4px 4px 0px 0px;
            padding: 10px 15px;
            color: #4A5568; 
            font-weight: 500; 
            border-bottom: 1px solid #e0e0e0; 
        }
        .stTabs [aria-selected="true"] { 
            background-color: #FFFFFF; 
            color: #2D3748; 
            font-weight: 600; 
            box-shadow: 0 -2px 4px rgba(0,0,0,0.03); 
            border-bottom: 1px solid #FFFFFF; 
        }
        
        /* Metric card styling */
        .metric-card {
            background-color: #FFFFFF;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 10px; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.04);
            transition: all 0.3s ease-in-out;
        }
        .metric-card:hover {
            box-shadow: 0 6px 12px rgba(0,0,0,0.08);
        }
        .metric-card-title {
            font-size: 0.95em;
            color: #4A5568; 
            margin-bottom: 8px;
            font-weight: 500;
        }
        .metric-card-value {
            font-size: 1.8em;
            font-weight: 700;
            color: #1A202C; 
        }
        .metric-card-delta {
            font-size: 0.85em;
            color: #718096; 
        }
        .metric-card-delta .positive { color: #38A169; }
        .metric-card-delta .negative { color: #E53E3E; }
        .metric-card.good { border-left: 5px solid #48BB78; } 
        .metric-card.warning { border-left: 5px solid #ECC94B; } 
        .metric-card.danger { border-left: 5px solid #F56565; } 

        /* Section headers */
        .section-header {
            font-size: 1.5em;
            font-weight: 600;
            color: #2D3748; 
            margin-top: 20px;
            margin-bottom: 15px;
            border-bottom: 2px solid #CBD5E0; 
            padding-bottom: 5px;
        }
        
        /* --- Mobile Responsiveness --- */
        @media (max-width: 768px) { /* Applies to tablets and smaller phones */
            .stTabs [data-baseweb="tab-list"] {
                gap: 8px; /* Smaller gap between tabs */
                /* Consider making tabs scrollable if too many for mobile: */
                /* overflow-x: auto; white-space: nowrap; */
            }
            .stTabs [data-baseweb="tab"] { 
                padding: 8px 10px; 
                font-size: 0.9em; 
            }

            .metric-card {
                padding: 12px; /* Smaller padding for metric cards */
                margin-bottom: 8px;
            }
            .metric-card-title {
                font-size: 0.85em; /* Adjusted for slightly better readability */
                margin-bottom: 4px;
            }
            .metric-card-value {
                font-size: 1.5em; /* Smaller font for metric values */
            }
            .metric-card-delta {
                font-size: 0.75em;
            }

            .section-header {
                font-size: 1.25em; /* Smaller section headers */
                margin-top: 15px;
                margin-bottom: 10px;
            }

            /* --- Streamlit Column Stacking --- */
            /* This targets the container for columns. */
            /* Streamlit's internal structure can change, so these selectors might need future adjustments. */
            /* Use browser dev tools "Inspect" to verify selectors if stacking doesn't work. */
            
            /* General approach for blocks used by st.columns */
            div[data-testid="stHorizontalBlock"] {
                flex-direction: column !important; /* Stack the main column blocks */
            }

            /* Ensure that individual columns within the now-stacked block take full width */
            div[data-testid="stHorizontalBlock"] > div[data-baseweb="block"] > div[data-testid="stVerticalBlock"],
            div[data-testid="stHorizontalBlock"] > div.stButton > button, /* Target buttons directly in columns */
            div[data-testid="stHorizontalBlock"] > div.stSelectbox { /* Target selectboxes directly in columns */
                width: 100% !important;
                margin-bottom: 10px; /* Add space between stacked items */
            }
            /* If columns are nested, you might need more specific selectors or a more general one: */
            /*
            .main > div > div > div > div[data-testid="stVerticalBlock"] {
                 width: 100% !important;
                 margin-bottom: 10px;
            }
            */


            /* Adjust font size for general text if needed */
            body, .stApp, div[data-testid="stMarkdownContainer"] p { /* Target paragraphs within markdown */
                font-size: 15px; /* Slightly smaller base font on mobile */
            }
            .stDataFrame { /* Make dataframes take less vertical space by default on mobile */
                /* Consider reducing height or showing fewer rows/columns by default on mobile */
                /* max-height: 350px; */ /* Adjust as needed */
            }
            /* Make selectboxes and text inputs more touch-friendly / readable */
            .stSelectbox > div, .stTextInput > div > div > input, .stTextArea > div > textarea, .stDateInput > div > div > input {
                font-size: 0.95em !important; /* Slightly larger touch target / font */
            }
        }

        @media (max-width: 480px) { /* Specific overrides for very small phone screens */
            .metric-card-title {
                font-size: 0.8em;
            }
            .metric-card-value {
                font-size: 1.3em; 
            }
            .stButton > button { /* Ensure buttons are full width and have decent padding */
                width: 100%;
                padding: 0.6em 0.5em; /* Adjust padding for better touch area */
                font-size: 0.95em; /* Ensure text is readable */
            }
             .stSelectbox div[data-baseweb="select"] { 
                font-size: 0.95em; /* Ensure readability */
             }
             .stDateInput input { /* Ensure readability */
                 font-size: 0.95em;
             }
             .stTextArea textarea { /* Ensure readability */
                 font-size: 0.95em;
             }
             .stTabs [data-baseweb="tab"] { 
                font-size: 0.85em; /* Even smaller tabs if needed */
                padding: 6px 8px;
            }
        }
    </style>
""", unsafe_allow_html=True)

# --- Session State Initialization ---
def init_session_state():
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    if 'all_data' not in st.session_state:
        st.session_state.all_data = {}
    if 'indicators' not in st.session_state:
        st.session_state.indicators = {}
    if 'openai_client' not in st.session_state:
        try:
            st.session_state.openai_client = openai.OpenAI(api_key=get_env_var('OPENAI_API_KEY'))
        except Exception as e:
            st.session_state.openai_client = None
    if 'current_page' not in st.session_state:
        st.session_state.current_page = "🏠 Home" 
    if 'ai_question' not in st.session_state:
        st.session_state.ai_question = ""
    if 'ai_chat_history' not in st.session_state:
        st.session_state.ai_chat_history = []
    
    # For contextual editing flow
    if 'selected_project_to_edit' not in st.session_state: 
        st.session_state.selected_project_to_edit = None
    if 'selected_pipeline_to_edit' not in st.session_state: 
        st.session_state.selected_pipeline_to_edit = None
    if 'manage_data_entity_type' not in st.session_state: 
        st.session_state.manage_data_entity_type = "Project" 
    if 'initial_load_complete' not in st.session_state:
        st.session_state.initial_load_complete = False

# --- Secret/Env Helper Functions ---
def get_env_var(key, default=None):
    if hasattr(st, "secrets") and key in st.secrets:
        return st.secrets[key]
    return os.getenv(key, default)

def get_google_credentials_file():
    if hasattr(st, "secrets") and "google_service_account" in st.secrets:
        import tempfile, json
        creds_dict = dict(st.secrets["google_service_account"])
        with tempfile.NamedTemporaryFile(delete=False, mode='w', suffix='.json') as f:
            json.dump(creds_dict, f)
            return f.name
    return os.getenv("GOOGLE_SHEETS_CREDENTIALS_FILE", "credentials.json")

# --- Data Loading and Processing ---
@st.cache_resource(ttl=300) 
def setup_google_sheets_cached():
    credentials_file = get_google_credentials_file()
    sheet_name = get_env_var('GOOGLE_SHEET_NAME')
    if not credentials_file or not os.path.exists(credentials_file):
        st.error(f"Credentials file not found: {credentials_file}")
        return None
    if not sheet_name:
        st.error("GOOGLE_SHEET_NAME not configured in .env or secrets")
        return None
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_file(credentials_file, scopes=scopes)
    
    max_retries = 3
    retry_delay = 5
    for attempt in range(max_retries):
        try:
            gc = gspread.authorize(credentials)
            sheet = gc.open(sheet_name)
            sheet.get_worksheet(0) 
            return sheet
        except Exception as e:
            if attempt < max_retries - 1:
                st.warning(f"GSheets connection attempt {attempt + 1} failed. Retrying in {retry_delay}s... Error: {e}")
                time.sleep(retry_delay)
            else:
                st.error(f"Error accessing spreadsheet after {max_retries} attempts: {e}")
                return None
    return None

@st.cache_data(ttl=300) 
def load_sheet_data_cached(_sheet_resource, worksheet_name): 
    if _sheet_resource is None: return pd.DataFrame()
    try:
        worksheet = _sheet_resource.worksheet(worksheet_name)
        all_values = worksheet.get_all_values()
        if not all_values: return pd.DataFrame()
        
        df = pd.DataFrame(all_values[1:], columns=all_values[0])
        df.columns = df.columns.str.strip()
        df = df.replace('', pd.NA).dropna(how='all')
        return df
    except Exception as e:
        st.warning(f"Error loading worksheet '{worksheet_name}': {e}")
        return pd.DataFrame()

def load_all_data():
    if not st.session_state.data_loaded:
        sheet = setup_google_sheets_cached()
        if sheet:
            worksheet_names = [
                'Project Inventory', 'Project Risks', 'Pipeline', 'Team Utilization',
                'Talent Gaps', 'Operational Gaps', 'Executive Activity',
                'Scenario Model Inputs', 'Do Nothing Scenario', 'Proposed Scenario',
                'Scenario Comparison', 'MappingTable', 'Project Observations'
            ]
            progress_bar = st.progress(0, text="Loading data...")
            for i, name in enumerate(worksheet_names):
                st.session_state.all_data[name] = load_sheet_data_cached(sheet, name)
                progress_bar.progress((i + 1) / len(worksheet_names), text=f"Loading {name}...")
            
            st.session_state.indicators = indicators.get_all_indicators(st.session_state.all_data)
            st.session_state.data_loaded = True
            progress_bar.empty()
            
            data_context_parts = []
            key_sheets_max_rows = {
                'Project Inventory': 50, 'Pipeline': 50, 'Project Risks': 30
            }
            default_max_rows = 15

            for name, df_sheet in st.session_state.all_data.items():
                if not df_sheet.empty:
                    current_max_rows = key_sheets_max_rows.get(name, default_max_rows)
                    col_info = f"Sheet: {name}\nColumns: {', '.join(df_sheet.columns)}\n"
                    if len(df_sheet) <= current_max_rows:
                        data_context_parts.append(col_info + df_sheet.to_string(index=False) + "\n")
                    else:
                        data_context_parts.append(col_info + df_sheet.head(current_max_rows).to_string(index=False) + f"\n... (showing top {current_max_rows} of {len(df_sheet)} rows)\n")
            st.session_state.data_context_string = "\n".join(data_context_parts)
            
            if not st.session_state.get('initial_load_complete', False): 
                st.sidebar.success("Data loaded successfully!")
                st.session_state.initial_load_complete = True
        else:
            st.error("Failed to connect to Google Sheets. Dashboard may not function correctly.")
            st.session_state.data_loaded = False

# --- Helper Functions for Display ---
def format_currency(value, default_na="N/A"):
    if pd.isna(value) or value is None: return default_na
    return f"${value:,.0f}"

def format_percentage(value, decimals=1, default_na="N/A"):
    if pd.isna(value) or value is None: return default_na
    return f"{value:.{decimals}f}%"

def format_number(value, decimals=0, default_na="N/A"):
    if pd.isna(value) or value is None: return default_na
    return f"{value:,.{decimals}f}"
    
def get_delta_color_class(value, is_higher_better=True):
    if pd.isna(value) or value is None or value == 0: return ""
    if is_higher_better: return "positive" if value > 0 else "negative"
    else: return "negative" if value > 0 else "positive"

def render_metric_card(title, value, delta=None, delta_label=None, card_class=""):
    delta_html = ""
    if delta is not None and delta_label:
        delta_color_class = "" 
        delta_html = f"<div class='metric-card-delta {delta_color_class}'>{delta} {delta_label}</div>"
    st.markdown(f"""
        <div class="metric-card {card_class}">
            <div class="metric-card-title">{title}</div>
            <div class="metric-card-value">{value}</div>
            {delta_html}
        </div>
    """, unsafe_allow_html=True)

# --- AI Assistant Functions ---
def escape_markdown_for_st(text):
    if not isinstance(text, str): return text
    escape_chars = r"([\\`*__{}\[\]()#+-.!])"
    return re.sub(escape_chars, r"\\\1", text)

def answer_critical_question_custom(user_question, data_kpis):
    q = user_question.lower().strip()
    if "total revenue" in q:
        return f"Total revenue is {format_currency(data_kpis.get('total_revenue', 0))}."
    if "red project" in q and ("count" in q or "how many" in q) :
        return f"There are {format_number(data_kpis.get('red_projects_count',0))} red projects."
    return None

def render_ai_assistant():
    st.sidebar.subheader("🤖 AI Assistant")
    if not st.session_state.openai_client:
        st.sidebar.warning("OpenAI client not initialized. AI Assistant may not function.")
        return

    for chat in st.session_state.ai_chat_history:
        with st.sidebar.chat_message(chat["role"]):
            st.markdown(escape_markdown_for_st(chat["content"]))

    prompt = st.sidebar.chat_input("Ask about your data...", key="ai_chat_input")

    if prompt:
        st.session_state.ai_chat_history.append({"role": "user", "content": prompt})
        with st.sidebar.chat_message("user"):
            st.markdown(escape_markdown_for_st(prompt))

        with st.sidebar.chat_message("assistant"):
            message_placeholder = st.empty()
            with st.spinner("Thinking..."):
                direct_answer = answer_critical_question_custom(prompt, st.session_state.indicators)
                
                if direct_answer:
                    response_text = direct_answer
                elif st.session_state.openai_client:
                    full_prompt = f"""Based on the following healthcare delivery data snapshot:
                    {st.session_state.get("data_context_string", "No data context available.")}
                    Question: {prompt}
                    Please provide a clear, concise answer. If the data is unavailable, say so.
                    """
                    try:
                        completion = st.session_state.openai_client.chat.completions.create(
                            model="gpt-4o",
                            messages=[
                                {"role": "system", "content": "You are a helpful healthcare delivery analytics assistant."},
                                {"role": "user", "content": full_prompt}
                            ]
                        )
                        response_text = completion.choices[0].message.content
                    except Exception as e:
                        response_text = f"Error querying OpenAI: {str(e)}"
                else:
                    response_text = "OpenAI client not available. Cannot process this question."
            
            message_placeholder.markdown(escape_markdown_for_st(response_text))
            st.session_state.ai_chat_history.append({"role": "assistant", "content": response_text})

# --- GSheet Update Helper ---
def update_gsheet_row(worksheet_name, identifier_col_name, identifier_value, update_data_dict):
    try:
        sheet = setup_google_sheets_cached()
        if not sheet:
            st.error("Failed to connect to Google Sheets for update.")
            return False
        
        worksheet = sheet.worksheet(worksheet_name)
        header = worksheet.row_values(1) 
        
        if identifier_col_name not in header:
            st.error(f"Identifier column '{identifier_col_name}' not found in sheet '{worksheet_name}'. Update failed.")
            return False
            
        cell = worksheet.find(identifier_value, in_column=header.index(identifier_col_name) + 1)
        if not cell:
            st.error(f"Could not find '{identifier_value}' in column '{identifier_col_name}' of sheet '{worksheet_name}'.")
            return False
        
        row_index = cell.row
        
        cells_to_update = []
        for col_name, new_value in update_data_dict.items():
            if col_name in header:
                col_index = header.index(col_name) + 1
                cells_to_update.append(gspread.Cell(row_index, col_index, str(new_value))) 
            else:
                st.warning(f"Column '{col_name}' not found in sheet '{worksheet_name}'. Skipping update for this field.")
        
        if cells_to_update:
            worksheet.update_cells(cells_to_update, value_input_option='USER_ENTERED')
            return True
        st.info("No valid fields to update were provided.") 
        return False 

    except gspread.exceptions.APIError as e:
        st.error(f"Google Sheets API Error updating '{worksheet_name}': {e}")
        return False
    except Exception as e:
        st.error(f"An unexpected error occurred updating Google Sheet '{worksheet_name}': {e}")
        return False

# --- Page Rendering Functions ---
def render_home_dashboard():
    st.title("🏠 Executive Dashboard")
    kpis = st.session_state.indicators

    st.markdown("<div class='section-header'>Overall Health Snapshot</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    with cols[0]:
        render_metric_card("Total Revenue", format_currency(kpis.get('total_revenue')), 
                           delta=format_percentage(kpis.get('revenue_vs_target_pct')), 
                           delta_label="of target", 
                           card_class="good" if kpis.get('revenue_vs_target_pct', 0) >= 100 else "warning")
    with cols[1]:
        render_metric_card("Pipeline Coverage", format_number(kpis.get('pipeline_coverage_ratio'), 1) + "x",
                           delta=format_percentage(kpis.get('pipeline_coverage_vs_target_pct')), 
                           delta_label="of target",
                           card_class="good" if kpis.get('pipeline_coverage_vs_target_pct',0) >=100 else "warning")
    with cols[2]:
        render_metric_card("Green Project Ratio", format_percentage(kpis.get('green_project_ratio', 0) * 100), 
                           delta=format_percentage(kpis.get('green_project_ratio_vs_target_pct')), 
                           delta_label="of target",
                           card_class="good" if kpis.get('green_project_ratio_vs_target_pct',0) >=100 else "danger")
    with cols[3]:
        render_metric_card("Avg. Customer NPS", format_number(kpis.get('avg_customer_nps'), 1),
                           delta=format_percentage(kpis.get('customer_nps_vs_target_pct')),
                           delta_label="of target",
                           card_class="good" if kpis.get('customer_nps_vs_target_pct',0) >=100 else "warning")

    st.markdown("<div class='section-header'>📄 Daily Executive Digest</div>", unsafe_allow_html=True)
    if 'daily_digest_content' not in st.session_state or st.button("🔄 Regenerate Daily Digest"):
        with st.spinner("Generating Daily Executive Digest..."):
            st.session_state.daily_digest_content = indicators.get_daily_digest_content(
                st.session_state.all_data,
                st.session_state.openai_client,
                st.session_state.get("data_context_string", "No data context available.")
            )
    if 'daily_digest_content' in st.session_state:
        digest_content = st.session_state.daily_digest_content
        if not isinstance(digest_content, str): digest_content = str(digest_content)
        st.markdown(digest_content, unsafe_allow_html=True) # Using st.markdown for better rendering of the digest

    st.markdown("<div class='section-header'>📉 Lagging Indicators</div>", unsafe_allow_html=True)
    lag_cols = st.columns(3)
    with lag_cols[0]:
        render_metric_card("FY Revenue", format_currency(kpis.get('total_revenue')), 
                           f"{format_currency(strategic_targets.REVENUE_TARGET)} target",
                           card_class="good" if kpis.get('revenue_vs_target_pct', 0) >= 100 else "warning")
    with lag_cols[1]:
        render_metric_card("Customer NPS (Avg)", format_number(kpis.get('avg_customer_nps'),1),
                           f"{strategic_targets.CUSTOMER_NPS_TARGET} target",
                           card_class="good" if kpis.get('customer_nps_vs_target_pct',0) >=100 else "warning")
    with lag_cols[2]:
        render_metric_card("Employee Pulse (Avg)", format_number(kpis.get('avg_employee_pulse_score'),1),
                           f"{strategic_targets.EMPLOYEE_PULSE_TARGET} target",
                           card_class="good" if kpis.get('employee_pulse_vs_target_pct',0) >=100 else "warning")

    st.markdown("<div class='section-header'>📈 Leading Indicators</div>", unsafe_allow_html=True)
    lead_cols1 = st.columns(3)
    with lead_cols1[0]:
        render_metric_card("Active Pipeline Value", format_currency(kpis.get('active_pipeline_value')),
                           f"{format_number(kpis.get('pipeline_coverage_ratio'),1)}x coverage",
                           card_class="good" if kpis.get('pipeline_coverage_vs_target_pct',0) >=100 else "warning")
    with lead_cols1[1]:
        render_metric_card("Avg. Deal Cycle", f"{format_number(kpis.get('avg_deal_cycle_time_days'))} days",
                           f"Median: {format_number(kpis.get('median_deal_cycle_time_days'))} days")
    with lead_cols1[2]:
         render_metric_card("Green Project Ratio", format_percentage(kpis.get('green_project_ratio', 0) * 100),
                           f"{format_percentage(strategic_targets.GREEN_PROJECT_TARGET*100)} target",
                           card_class="good" if kpis.get('green_project_ratio_vs_target_pct',0) >=100 else "danger")
    
    lead_cols2 = st.columns(3)
    with lead_cols2[0]:
        render_metric_card("Recent Sponsor Check-ins", f"{format_number(kpis.get('recent_meaningful_checkins_count'))} ({format_percentage(kpis.get('recent_meaningful_checkins_pct'))})",
                            f"{kpis.get('overdue_sponsor_checkin_count')} overdue")
    with lead_cols2[1]:
        render_metric_card("High Severity Risks", f"{format_number(kpis.get('high_severity_risk_count'))}",
                           f"{format_currency(kpis.get('high_severity_risk_impact'))} impact")
    with lead_cols2[2]:
        render_metric_card("Avg. Delivery Utilization", format_percentage(kpis.get('avg_delivery_utilization_pct')),
                           f"{kpis.get('under_utilized_delivery_count')} under-utilized")

def render_projects_page():
    st.title("📊 Projects Deep Dive")
    kpis = st.session_state.indicators
    project_df_original = st.session_state.all_data.get('Project Inventory', pd.DataFrame())

    st.markdown("<div class='section-header'>Project Health Overview</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    with cols[0]: st.metric("Total Projects", format_number(kpis.get('total_projects')))
    with cols[1]: st.metric("Red Projects", format_number(kpis.get('red_projects_count')), delta=f"{format_currency(kpis.get('red_project_revenue'))} at risk", delta_color="inverse")
    with cols[2]: st.metric("Avg. Health Score", format_number(kpis.get('avg_project_health_score'), 1), f"Median: {format_number(kpis.get('median_project_health_score'),1)}")
    with cols[3]: st.metric("Avg. Total Score", format_number(kpis.get('avg_total_project_score'), 1), f"Median: {format_number(kpis.get('median_total_project_score'),1)}")

    tab1, tab2, tab3, tab4 = st.tabs(["📋 All Projects", "🚨 At-Risk Projects", "🏆 Score Analysis", "📅 Sponsor Check-ins"])

    with tab1:
        st.subheader("All Projects List")
        if not project_df_original.empty and 'Project Name' in project_df_original.columns:
            display_df = project_df_original.copy()
            if 'Revenue' in display_df.columns: display_df['Revenue'] = indicators.safe_to_numeric(display_df['Revenue']).apply(lambda x: f"${x:,.0f}")
            if 'Project Health Score' in display_df.columns: display_df['Project Health Score'] = indicators.safe_to_numeric(display_df['Project Health Score']).round(1)
            
            display_cols = ['Project Name', 'Client', 'Status (R/Y/G)', 'Revenue', 'Project Health Score', 'Project End Date', 'Key Issues']
            final_cols = [col for col in display_cols if col in display_df.columns]
            st.dataframe(display_df[final_cols], use_container_width=True, height=400) # Reduced height
            
            st.markdown("---")
            st.subheader("Edit Project Details")
            project_names_for_edit = [""] + sorted(project_df_original['Project Name'].astype(str).unique().tolist())
            selected_project_name_for_action = st.selectbox(
                "Select a project to edit:", project_names_for_edit, index=0, key="project_action_selector"
            )
            if selected_project_name_for_action:
                if st.button(f"✏️ Edit '{selected_project_name_for_action}'", key=f"edit_proj_{selected_project_name_for_action.replace(' ','_')}"): # Unique key for button
                    st.session_state.selected_project_to_edit = selected_project_name_for_action
                    st.session_state.manage_data_entity_type = "Project" # Set type for Manage Data page
                    st.session_state.current_page = "📝 Manage Data"
                    st.rerun()
        else: st.info("No project data or 'Project Name' column available.")
            
    with tab2:
        st.subheader("At-Risk & Non-Green Projects")
        if not project_df_original.empty and 'Status (R/Y/G)' in project_df_original.columns:
            project_df_original_status_copy = project_df_original.copy() # Work on a copy for modification
            project_df_original_status_copy['Status (R/Y/G)'] = project_df_original_status_copy['Status (R/Y/G)'].astype(str).str.upper()
            at_risk_df = project_df_original_status_copy[project_df_original_status_copy['Status (R/Y/G)'].isin(['R', 'Y'])].copy()
            if not at_risk_df.empty:
                display_at_risk_cols = ['Project Name', 'Status (R/Y/G)', 'Revenue', 'Key Issues', 'Next Steps', 'Executive Support Required']
                final_at_risk_cols = [col for col in display_at_risk_cols if col in at_risk_df.columns]
                if 'Revenue' in final_at_risk_cols:
                     at_risk_df['RevenueNum'] = indicators.safe_to_numeric(at_risk_df['Revenue'])
                     at_risk_df['Revenue'] = at_risk_df['RevenueNum'].apply(lambda x: f"${x:,.0f}")
                st.dataframe(at_risk_df[final_at_risk_cols], use_container_width=True)
            else: st.success("🎉 No projects currently marked Red or Yellow!")
        else: st.info("No project data or 'Status (R/Y/G)' column available to determine at-risk projects.")
        # st.subheader("Non-Green Projects (from indicators module)")
        # non_green_list = kpis.get('non_green_projects_list', [])
        # if non_green_list: st.table(pd.DataFrame(non_green_list))
        # else: st.success("🎉 All projects are Green according to ratio calculation (or data unavailable)!")

    with tab3:
        st.subheader("Project Score Analysis")
        col_score1, col_score2 = st.columns(2)
        project_df_for_scores = project_df_original.copy() 
        with col_score1:
            if 'Project Health Score' in project_df_for_scores.columns and not project_df_for_scores['Project Health Score'].dropna().empty:
                project_df_for_scores['Project Health Score Num'] = indicators.safe_to_numeric(project_df_for_scores['Project Health Score'])
                fig_health = px.histogram(project_df_for_scores.dropna(subset=['Project Health Score Num']), x='Project Health Score Num', title='Project Health Score Distribution', nbins=10, text_auto=True)
                fig_health.update_layout(bargap=0.1); st.plotly_chart(fig_health, use_container_width=True)
            else: st.info("Project Health Score data not available for distribution chart.")
            st.write(f"Top Project (Health Score): **{kpis.get('top_project_by_health_score', 'N/A')}**")
            st.write(f"Bottom Project (Health Score): **{kpis.get('bottom_project_by_health_score', 'N/A')}**")
            st.write("Health Score Bands (%):"); st.json(kpis.get('project_health_score_bands_pct', {}))
        with col_score2:
            if 'Total Project Score' in project_df_for_scores.columns and not project_df_for_scores['Total Project Score'].dropna().empty:
                project_df_for_scores['Total Project Score Num'] = indicators.safe_to_numeric(project_df_for_scores['Total Project Score'])
                fig_total = px.histogram(project_df_for_scores.dropna(subset=['Total Project Score Num']), x='Total Project Score Num', title='Total Project Score Distribution', nbins=10, text_auto=True)
                fig_total.update_layout(bargap=0.1); st.plotly_chart(fig_total, use_container_width=True)
            else: st.info("Total Project Score data not available for distribution.")
            st.write(f"Top Project (Total Score): **{kpis.get('top_project_by_total_score', 'N/A')}**")
            st.write(f"Bottom Project (Total Score): **{kpis.get('bottom_project_by_total_score', 'N/A')}**")
            st.write("Total Score Bands (%):"); st.json(kpis.get('total_project_score_bands_pct', {}))
            
    with tab4:
        st.subheader("Sponsor Check-in Status")
        st.metric("Recent Meaningful Check-ins (last 30 days)", f"{format_number(kpis.get('recent_meaningful_checkins_count'))} ({format_percentage(kpis.get('recent_meaningful_checkins_pct'))})")
        st.metric("Projects Overdue for Sponsor Check-in", kpis.get('overdue_sponsor_checkin_count'))
        overdue_list = kpis.get('overdue_checkin_projects_list', [])
        if overdue_list: st.write("Projects Overdue for Check-in:"); st.table(pd.DataFrame(overdue_list))
        else: st.success("👍 All active projects have recent sponsor check-ins logged (or data unavailable).")

def render_pipeline_page():
    st.title("📈 Pipeline Deep Dive")
    kpis = st.session_state.indicators
    pipeline_df_original = st.session_state.all_data.get('Pipeline', pd.DataFrame())

    st.markdown("<div class='section-header'>Pipeline Health Overview</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    with cols[0]: st.metric("Active Pipeline Value", format_currency(kpis.get('active_pipeline_value')))
    with cols[1]: st.metric("Total Potential Value", format_currency(kpis.get('total_potential_pipeline_value'))) 
    with cols[2]: st.metric("Pipeline Coverage Ratio", f"{format_number(kpis.get('pipeline_coverage_ratio'),1)}x")
    with cols[3]: st.metric("Avg. Pipeline Score", format_number(kpis.get('avg_pipeline_score'),1), f"Median: {format_number(kpis.get('median_pipeline_score'),1)}")

    tab1, tab2, tab3 = st.tabs(["📋 All Opportunities", "🏆 Score Analysis", "⏱️ Deal Cycle Analysis"])

    with tab1:
        st.subheader("All Pipeline Opportunities")
        if not pipeline_df_original.empty and 'Account' in pipeline_df_original.columns: # Identifier for pipeline
            display_df = pipeline_df_original.copy()
            if 'Open Pipeline_Active Work' in display_df.columns: display_df['Open Pipeline_Active Work'] = indicators.safe_to_numeric(display_df['Open Pipeline_Active Work']).apply(lambda x: f"${x:,.0f}")
            if 'Percieved Annual AMO' in display_df.columns: display_df['Percieved Annual AMO'] = indicators.safe_to_numeric(display_df['Percieved Annual AMO']).apply(lambda x: f"${x:,.0f}")
            if 'Pipeline Score' in display_df.columns: display_df['Pipeline Score'] = indicators.safe_to_numeric(display_df['Pipeline Score']).round(1)
            display_cols = ['Account', 'Open Pipeline_Active Work', 'Percieved Annual AMO', 'Pipeline Score', 'Pursuit Tier', 'Horizon', 'Opportunity Created Date', 'Closed Won Date']
            final_cols = [col for col in display_cols if col in display_df.columns]
            st.dataframe(display_df[final_cols], use_container_width=True, height=400) # Reduced height
            
            st.markdown("---")
            st.subheader("Edit Opportunity Details")
            pipeline_account_names_for_edit = [""] + sorted(pipeline_df_original['Account'].astype(str).unique().tolist())
            selected_pipeline_account_for_action = st.selectbox(
                "Select an Account/Opportunity to edit:", pipeline_account_names_for_edit, index=0, key="pipeline_action_selector"
            )
            if selected_pipeline_account_for_action:
                if st.button(f"✏️ Edit '{selected_pipeline_account_for_action}'", key=f"edit_pipe_{selected_pipeline_account_for_action.replace(' ','_')}"): # Unique key for button
                    st.session_state.selected_pipeline_to_edit = selected_pipeline_account_for_action
                    st.session_state.manage_data_entity_type = "Pipeline Opportunity" # Set type for Manage Data page
                    st.session_state.current_page = "📝 Manage Data"
                    st.rerun()
        else: st.info("No pipeline data or 'Account' column available.")
            
    with tab2:
        st.subheader("Pipeline Score Analysis")
        pipeline_df_for_scores = pipeline_df_original.copy()
        col_score1, col_score2 = st.columns(2)
        with col_score1:
            if 'Pipeline Score' in pipeline_df_for_scores.columns and not pipeline_df_for_scores['Pipeline Score'].dropna().empty:
                pipeline_df_for_scores['Pipeline Score Num'] = indicators.safe_to_numeric(pipeline_df_for_scores['Pipeline Score'])
                fig_health = px.histogram(pipeline_df_for_scores.dropna(subset=['Pipeline Score Num']), x='Pipeline Score Num', title='Pipeline Score Distribution', nbins=10, text_auto=True)
                fig_health.update_layout(bargap=0.1); st.plotly_chart(fig_health, use_container_width=True)
            else: st.info("Pipeline Score data not available for distribution.")
            st.write(f"Top Opportunity (Pipeline Score): **{kpis.get('top_pipeline_by_score', 'N/A')}**"); st.write(f"Bottom Opportunity (Pipeline Score): **{kpis.get('bottom_pipeline_by_score', 'N/A')}**")
            st.write("Pipeline Score Bands (%):"); st.json(kpis.get('pipeline_score_bands_pct', {}))
        with col_score2:
            if 'Total Deal Score' in pipeline_df_for_scores.columns and not pipeline_df_for_scores['Total Deal Score'].dropna().empty:
                pipeline_df_for_scores['Total Deal Score Num'] = indicators.safe_to_numeric(pipeline_df_for_scores['Total Deal Score'])
                fig_total = px.histogram(pipeline_df_for_scores.dropna(subset=['Total Deal Score Num']), x='Total Deal Score Num', title='Total Deal Score Distribution', nbins=10, text_auto=True)
                fig_total.update_layout(bargap=0.1); st.plotly_chart(fig_total, use_container_width=True)
            else: st.info("Total Deal Score data not available for distribution.")
            st.write(f"Top Opportunity (Total Score): **{kpis.get('top_pipeline_by_total_score', 'N/A')}**"); st.write(f"Bottom Opportunity (Total Score): **{kpis.get('bottom_pipeline_by_total_score', 'N/A')}**")
            st.write("Total Deal Score Bands (%):"); st.json(kpis.get('total_deal_score_bands_pct', {}))
            
    with tab3:
        st.subheader("Deal Cycle Analysis")
        st.metric("Avg. Deal Cycle Time", f"{format_number(kpis.get('avg_deal_cycle_time_days'))} days")
        st.metric("Median Deal Cycle Time", f"{format_number(kpis.get('median_deal_cycle_time_days'))} days")
        cycle_by_tier = kpis.get('deal_cycle_time_by_tier', {})
        if cycle_by_tier:
            st.write("Deal Cycle Time by Pursuit Tier:")
            tier_data = [{"Pursuit Tier": tier, "Avg Days": format_number(times.get('mean',0)), "Median Days": format_number(times.get('median',0))} for tier, times in cycle_by_tier.items()]
            st.table(pd.DataFrame(tier_data))
        st.metric("Avg. Next Deal Gap (after project end)", f"{format_number(kpis.get('avg_next_deal_gap_days'))} days")
        overdue_next_deal_count = kpis.get('overdue_next_deal_discussion_count', 0)
        st.metric("Projects Overdue for Next Deal Discussion", overdue_next_deal_count)
        if overdue_next_deal_count > 0: st.write("Projects Overdue:"); st.table(pd.DataFrame(kpis.get('overdue_next_deal_projects_list', [])))

def render_risks_page():
    st.title("⚠️ Risks Deep Dive")
    kpis = st.session_state.indicators
    risk_df_original = st.session_state.all_data.get('Project Risks', pd.DataFrame())
    st.markdown("<div class='section-header'>Risk Overview</div>", unsafe_allow_html=True)
    cols = st.columns(3)
    with cols[0]: st.metric("Total Risk Impact", format_currency(kpis.get('total_risk_impact')))
    with cols[1]: st.metric("High Severity Risk Items", format_number(kpis.get('high_severity_risk_count')))
    with cols[2]: st.metric("High Severity Risk Impact", format_currency(kpis.get('high_severity_risk_impact')), delta=f"{format_percentage(kpis.get('high_risk_impact_as_pct_of_total'))} of total impact")
    if not risk_df_original.empty:
        if 'Impact ($)' in risk_df_original.columns and 'Severity' in risk_df_original.columns:
            risk_df_cleaned = risk_df_original.copy()
            risk_df_cleaned['Impact ($)'] = indicators.safe_to_numeric(risk_df_cleaned['Impact ($)'])
            risk_df_cleaned['Severity'] = risk_df_cleaned['Severity'].astype(str).fillna('Unknown').str.capitalize()
            severity_impact = risk_df_cleaned.groupby('Severity')['Impact ($)'].sum().reset_index()
            fig_sev_val = px.pie(severity_impact, names='Severity', values='Impact ($)', title='Risk Impact by Severity', hole=0.3, color_discrete_map={'High': '#F56565', 'Medium': '#ECC94B', 'Low': '#48BB78', 'Unknown': '#A0AEC0'})
            st.plotly_chart(fig_sev_val, use_container_width=True)
        else: st.info("Required columns ('Impact ($)' or 'Severity') not found for risk distribution chart.")
        st.subheader("Full Risk Register"); st.dataframe(risk_df_original, use_container_width=True, height=600)
    else: st.info("No risk data available.")

def render_team_ops_page():
    st.title("👥 Team & Operations")
    kpis = st.session_state.indicators
    tab_util, tab_gaps, tab_exec = st.tabs(["📊 Team Utilization & Pulse", "🛠️ Talent & Operational Gaps", "💼 Executive Activity"])
    with tab_util:
        st.markdown("<div class='section-header'>Team Utilization</div>", unsafe_allow_html=True)
        cols_util = st.columns(2)
        with cols_util[0]: st.metric("Avg. Executive Utilization", format_percentage(kpis.get('avg_exec_utilization_pct')), f"{kpis.get('over_utilized_execs_count')} execs >70%")
        with cols_util[1]: st.metric("Avg. Delivery Utilization", format_percentage(kpis.get('avg_delivery_utilization_pct')), f"{kpis.get('under_utilized_delivery_count')} under (<70%), {kpis.get('over_utilized_delivery_count')} over (>100%)")
        util_df_original = st.session_state.all_data.get('Team Utilization', pd.DataFrame())
        if not util_df_original.empty and 'Employee Name' in util_df_original.columns and 'Utilization (%)' in util_df_original.columns and 'Role' in util_df_original.columns:
            util_df_c = util_df_original.copy()
            util_df_c['Utilization (%)'] = indicators.safe_to_numeric(util_df_c['Utilization (%)'])
            fig_util = px.bar(util_df_c.sort_values('Utilization (%)', ascending=False), x='Employee Name', y='Utilization (%)', color='Role', title='Team Member Utilization', text_auto=True)
            fig_util.update_layout(xaxis_tickangle=-45, height=500); st.plotly_chart(fig_util, use_container_width=True)
        else: st.info("Team utilization data or required columns not available for chart.")
        st.markdown("<div class='section-header'>Employee Pulse</div>", unsafe_allow_html=True)
        st.metric("Average Employee Pulse Score", format_number(kpis.get('avg_employee_pulse_score'),1), f"{format_percentage(kpis.get('employee_pulse_vs_target_pct'))} of target")
    with tab_gaps:
        st.markdown("<div class='section-header'>Talent Gaps</div>", unsafe_allow_html=True)
        talent_gaps_df = st.session_state.all_data.get('Talent Gaps', pd.DataFrame())
        if not talent_gaps_df.empty: st.dataframe(talent_gaps_df, use_container_width=True)
        else: st.info("No talent gap data available.")
        st.markdown("<div class='section-header'>Operational Gaps</div>", unsafe_allow_html=True)
        operational_gaps_df = st.session_state.all_data.get('Operational Gaps', pd.DataFrame())
        if not operational_gaps_df.empty: st.dataframe(operational_gaps_df, use_container_width=True)
        else: st.info("No operational gap data available.")
    with tab_exec:
        st.markdown("<div class='section-header'>Executive Activity & Strategic Cost</div>", unsafe_allow_html=True)
        st.metric("Total Strategic Cost (from Exec Activity)", format_currency(kpis.get('total_strategic_cost')))
        st.metric("Total Strategic Activities Logged", format_number(kpis.get('total_strategic_activities_count')))
        exec_activity_df_original = st.session_state.all_data.get('Executive Activity', pd.DataFrame())
        if not exec_activity_df_original.empty: st.dataframe(exec_activity_df_original, use_container_width=True)
        else: st.info("No executive activity data available.")

def render_scenario_modeling_page():
    st.title("⚙️ Scenario Modeling")
    scenario_tabs_map = {"Inputs": 'Scenario Model Inputs', "Do Nothing": 'Do Nothing Scenario', "Proposed": 'Proposed Scenario', "Comparison": 'Scenario Comparison'}
    selected_tab_name = st.selectbox("Select Scenario View", list(scenario_tabs_map.keys()))
    df_name = scenario_tabs_map[selected_tab_name]
    df = st.session_state.all_data.get(df_name, pd.DataFrame())
    st.subheader(f"{selected_tab_name} Data")
    if not df.empty: st.dataframe(df, use_container_width=True)
    else: st.info(f"No data available for '{df_name}'.")

def render_data_explorer_page():
    st.title("🔍 Data Explorer")
    worksheet_names = list(st.session_state.all_data.keys())
    if not worksheet_names: st.warning("No data loaded yet. Please wait or try refreshing."); return
    selected_view = st.selectbox("Select Worksheet to View", worksheet_names)
    if selected_view in st.session_state.all_data:
        df = st.session_state.all_data[selected_view]
        if not df.empty:
            st.subheader(f"Raw Data: {selected_view}")
            if selected_view == 'Project Inventory' and 'Status (R/Y/G)' in df.columns:
                def color_status(val):
                    color = ''; s_val = str(val).upper()
                    if pd.isna(val): return color
                    if s_val == 'R': color = 'background-color: #ffcdd2' 
                    elif s_val == 'Y': color = 'background-color: #fff9c4' 
                    elif s_val == 'G': color = 'background-color: #c8e6c9' 
                    return color
                try: st.dataframe(df.style.applymap(color_status, subset=['Status (R/Y/G)']), use_container_width=True)
                except Exception as e: st.warning(f"Could not apply color styling for Status: {e}"); st.dataframe(df, use_container_width=True)
            else: st.dataframe(df, use_container_width=True)
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(label=f"Download {selected_view} as CSV", data=csv, file_name=f'{selected_view.lower().replace(" ", "_")}.csv', mime='text/csv')
        else: st.info(f"Worksheet '{selected_view}' is empty or failed to load.")

# --- Page for Data Management ---
def render_manage_data_page():
    st.title("📝 Manage Data")

    entity_type_options = ["Project", "Pipeline Opportunity"]
    # Determine default index for radio based on session state, ensure it's valid
    try:
        entity_type_index = entity_type_options.index(st.session_state.manage_data_entity_type)
    except ValueError:
        entity_type_index = 0 # Default to "Project" if current state is invalid
        st.session_state.manage_data_entity_type = "Project"


    entity_type = st.radio(
        "What do you want to manage?",
        entity_type_options,
        index=entity_type_index, 
        horizontal=True,
        key="entity_type_selector_main" # Unique key
    )
    st.session_state.manage_data_entity_type = entity_type 

    st.markdown("---")

    if entity_type == "Project":
        render_manage_project_form()
    elif entity_type == "Pipeline Opportunity":
        render_manage_pipeline_form()

def render_manage_project_form():
    st.subheader("Update Project Details")
    project_df = st.session_state.all_data.get('Project Inventory', pd.DataFrame())

    if project_df.empty or 'Project Name' not in project_df.columns:
        st.warning("Project Inventory data or 'Project Name' column not loaded. Cannot manage project data.")
        st.session_state.selected_project_to_edit = None 
        return

    project_names_list = [""] + sorted(project_df['Project Name'].astype(str).unique().tolist())
    
    default_project_index = 0
    if st.session_state.selected_project_to_edit and st.session_state.selected_project_to_edit in project_names_list:
        default_project_index = project_names_list.index(st.session_state.selected_project_to_edit)

    selected_project_name = st.selectbox(
        "Select Project to Update", 
        project_names_list, 
        index=default_project_index,
        key="project_update_selector_on_manage_page_projectform" 
    )
    
    if st.session_state.selected_project_to_edit: 
         st.session_state.selected_project_to_edit = None

    if selected_project_name:
        project_data_series = project_df[project_df['Project Name'] == selected_project_name]
        if project_data_series.empty:
            st.error(f"Project '{selected_project_name}' not found. Please refresh data.")
            return
        project_data = project_data_series.iloc[0]
        
        try:
            current_status_val = str(project_data.get('Status (R/Y/G)', 'G')).upper()
            status_options = ['R', 'Y', 'G']
            default_status_index = status_options.index(current_status_val) if current_status_val in status_options else status_options.index('G')
        except ValueError: default_status_index = 2 

        default_key_issues = str(project_data.get('Key Issues', ''))
        default_next_steps = str(project_data.get('Next Steps', ''))
        
        default_checkin_date = None
        raw_checkin_date = project_data.get('Last Sponsor Checkin Date')
        if pd.notna(raw_checkin_date):
            if isinstance(raw_checkin_date, (datetime, date)):
                default_checkin_date = raw_checkin_date if isinstance(raw_checkin_date, date) else raw_checkin_date.date()
            else: 
                try: default_checkin_date = pd.to_datetime(raw_checkin_date, errors='raise').date()
                except (ValueError, TypeError): st.warning(f"Could not parse existing check-in date '{raw_checkin_date}' for {selected_project_name}. Please re-enter.")
        
        default_checkin_notes = str(project_data.get('Sponsor Checkin Notes', ''))

        with st.form(key=f"update_project_form_{selected_project_name.replace(' ','_')}"):
            st.write(f"#### Updating Project: {selected_project_name}")
            new_status = st.selectbox("Status (R/Y/G)", ['R', 'Y', 'G'], index=default_status_index)
            new_key_issues = st.text_area("Key Issues", value=default_key_issues, height=100)
            new_next_steps = st.text_area("Next Steps", value=default_next_steps, height=100)
            new_checkin_date = st.date_input("Last Sponsor Check-in Date", value=default_checkin_date)
            new_checkin_notes = st.text_area("Sponsor Checkin Notes", value=default_checkin_notes, height=150)
            
            submitted = st.form_submit_button("💾 Update Project in Google Sheet")

            if submitted:
                with st.spinner(f"Updating {selected_project_name}..."):
                    update_payload = {
                        "Status (R/Y/G)": new_status, "Key Issues": new_key_issues, "Next Steps": new_next_steps,
                        "Last Sponsor Checkin Date": new_checkin_date.strftime('%Y-%m-%d') if new_checkin_date else "",
                        "Sponsor Checkin Notes": new_checkin_notes
                    }
                    success = update_gsheet_row('Project Inventory', 'Project Name', selected_project_name, update_payload)
                    if success:
                        st.success(f"Project '{selected_project_name}' updated successfully!")
                        st.session_state.data_loaded = False; st.cache_data.clear(); st.cache_resource.clear()
                        st.session_state.initial_load_complete = False
                        st.rerun()
    else: st.info("Select a project to update its details.")

def render_manage_pipeline_form():
    st.subheader("Update Pipeline Opportunity Details")
    pipeline_df = st.session_state.all_data.get('Pipeline', pd.DataFrame())

    if pipeline_df.empty or 'Account' not in pipeline_df.columns:
        st.warning("Pipeline data or 'Account' column not loaded. Cannot manage pipeline data.")
        st.session_state.selected_pipeline_to_edit = None
        return

    pipeline_account_names_list = [""] + sorted(pipeline_df['Account'].astype(str).unique().tolist())
    
    default_pipeline_index = 0
    if st.session_state.selected_pipeline_to_edit and st.session_state.selected_pipeline_to_edit in pipeline_account_names_list:
        default_pipeline_index = pipeline_account_names_list.index(st.session_state.selected_pipeline_to_edit)

    selected_account_name = st.selectbox(
        "Select Pipeline Opportunity (by Account) to Update",
        pipeline_account_names_list, index=default_pipeline_index, key="pipeline_update_selector_on_manage_page"
    )

    if st.session_state.selected_pipeline_to_edit: 
        st.session_state.selected_pipeline_to_edit = None

    if selected_account_name:
        opportunity_data_series = pipeline_df[pipeline_df['Account'] == selected_account_name]
        if opportunity_data_series.empty:
            st.error(f"Opportunity for Account '{selected_account_name}' not found. Please refresh data.")
            return
        opportunity_data = opportunity_data_series.iloc[0] 
        
        # --- Pre-fill form fields for Pipeline (ADD MORE AS NEEDED) ---
        def get_options_and_index(df, col_name, current_val_str):
            options = [""] # Start with a blank option
            if col_name in df.columns:
                options.extend(sorted(df[col_name].astype(str).unique().tolist()))
            
            default_idx = 0
            if current_val_str in options:
                default_idx = options.index(current_val_str)
            return options, default_idx

        horizon_options, default_horizon_idx = get_options_and_index(pipeline_df, 'Horizon', str(opportunity_data.get('Horizon', '')))
        pursuit_tier_options, default_pursuit_tier_idx = get_options_and_index(pipeline_df, 'Pursuit Tier', str(opportunity_data.get('Pursuit Tier', '')))
        
        default_notes = str(opportunity_data.get('Notes', ''))
        default_help_needed = str(opportunity_data.get('Help Needed', ''))
        default_actions = str(opportunity_data.get('Actions', ''))

        yn_options = ["", "Y", "N"] 
        default_deal_reg_val = str(opportunity_data.get('Deal Registered YN', '')).upper()
        default_deal_reg_idx = yn_options.index(default_deal_reg_val) if default_deal_reg_val in yn_options else 0
        
        # Example: Percieved Annual AMO (numeric)
        default_percieved_amo = 0.0
        if 'Percieved Annual AMO' in opportunity_data and pd.notna(opportunity_data['Percieved Annual AMO']):
            try:
                default_percieved_amo = float(str(opportunity_data['Percieved Annual AMO']).replace('$','').replace(',',''))
            except ValueError:
                default_percieved_amo = 0.0


        with st.form(key=f"update_pipeline_form_{selected_account_name.replace(' ','_')}"):
            st.write(f"#### Updating Pipeline Opportunity: {selected_account_name}")
            
            cols1, cols2 = st.columns(2)
            with cols1:
                new_horizon = st.selectbox("Horizon", horizon_options, index=default_horizon_idx)
                new_deal_reg = st.selectbox("Deal Registered (Y/N)", yn_options, index=default_deal_reg_idx)
            with cols2:
                new_pursuit_tier = st.selectbox("Pursuit Tier", pursuit_tier_options, index=default_pursuit_tier_idx)
                new_percieved_amo = st.number_input("Percieved Annual AMO", value=default_percieved_amo, step=1000.0, format="%d")


            new_notes = st.text_area("Notes", value=default_notes, height=100)
            new_help_needed = st.text_area("Help Needed", value=default_help_needed, height=75)
            new_actions = st.text_area("Actions", value=default_actions, height=75)
            
            submitted = st.form_submit_button("💾 Update Pipeline Opportunity in Google Sheet")

            if submitted:
                with st.spinner(f"Updating {selected_account_name}..."):
                    update_payload = {
                        "Horizon": new_horizon, "Pursuit Tier": new_pursuit_tier, "Notes": new_notes,
                        "Help Needed": new_help_needed, "Actions": new_actions, "Deal Registered YN": new_deal_reg,
                        "Percieved Annual AMO": int(new_percieved_amo) # Ensure it's a number for GSheet if it expects one
                    }
                    update_payload_cleaned = {k: v for k, v in update_payload.items() if isinstance(v, (int, float)) or (isinstance(v, str) and v != "")}


                    success = update_gsheet_row('Pipeline', 'Account', selected_account_name, update_payload_cleaned)
                    if success:
                        st.success(f"Pipeline opportunity for '{selected_account_name}' updated successfully!")
                        st.session_state.data_loaded = False; st.cache_data.clear(); st.cache_resource.clear()
                        st.session_state.initial_load_complete = False
                        st.rerun()
    else: st.info("Select a pipeline opportunity (by Account) to update its details.")

def render_scenario_playground_page():
    st.title("🧪 Scenario Playground")
    st.markdown("Interactively adjust scenario assumptions and see the impact in real time.")

    # --- Load scenario inputs ---
    inputs_df = st.session_state.all_data.get('Scenario Model Inputs', pd.DataFrame())
    if inputs_df.empty or not {'Assumption', 'Value'}.issubset(inputs_df.columns):
        st.warning("No scenario inputs found or missing required columns.")
        return

    # --- Build baseline_inputs from the original sheet values (never changes) ---
    baseline_inputs = {}
    for _, row in inputs_df.iterrows():
        label = row['Assumption']
        value = row['Value']
        if isinstance(value, (int, float)):
            baseline_inputs[label] = float(value)
        elif isinstance(value, str) and value.strip().endswith('%'):
            try:
                baseline_inputs[label] = float(value.strip().replace('%',''))
            except Exception:
                baseline_inputs[label] = 0.0
        elif isinstance(value, str) and value.strip().startswith('$'):
            try:
                baseline_inputs[label] = float(value.strip().replace('$','').replace(',',''))
            except Exception:
                baseline_inputs[label] = 0.0
        else:
            baseline_inputs[label] = value

    # --- Render input widgets dynamically, using baseline as default ---
    st.subheader("Adjust Scenario Assumptions")
    proposed_inputs = {}
    for label, default in baseline_inputs.items():
        # Use slider for percentages, number_input for numbers, etc.
        if isinstance(default, float) or isinstance(default, int):
            # If the label or value suggests percent, use slider
            if '%' in label or (0 <= default <= 100 and not isinstance(default, bool)):
                proposed_inputs[label] = st.slider(label, 0.0, 100.0, float(default), step=1.0)
            else:
                proposed_inputs[label] = st.number_input(label, value=float(default))
        else:
            proposed_inputs[label] = st.text_input(label, value=str(default))

    # --- Calculation logic for scenarios ---
    def calculate_scenarios(inputs):
        try:
            vp_hourly = float(inputs.get('VP Hourly Selling Value', 0))
            vp_hours_weekly = float(inputs.get('VP Hours Weekly on Delivery', 0))
            head_hourly = float(inputs.get('Head of Delivery Hourly Strategic Delivery Value', 0))
            head_hours_weekly = float(inputs.get('Head of Delivery Weekly Tactical Delivery Hours', 0))
            avg_project_size = float(inputs.get('Avg. Project Size', 0))
            sales_conv_rate = float(inputs.get('Sales Conversion Rate (%)', 0)) / 100.0
            cost_turnover = float(inputs.get('Avg. Cost of Turnover per Senior Employee', 0))
            pct_projects_at_risk = float(inputs.get('% Projects at Risk', 0)) / 100.0
            revenue_at_risk = float(inputs.get('Revenue at Risk due to troubled projects', 0))
            work_weeks = float(inputs.get('Work Weeks in a year', 50))
            chief_salary = float(inputs.get('Cost of Chief of Staff Salary', 100000)) if 'Cost of Chief of Staff Salary' in inputs else 100000
        except Exception as e:
            st.error(f"Error parsing inputs: {e}")
            return {}
        results = {}
        # Do Nothing scenario calculations
        results['Lost Sales (VP involvement)'] = -1 * vp_hours_weekly * vp_hourly * work_weeks
        results['Strategic Loss (Head of Delivery Role)'] = -1 * head_hours_weekly * head_hourly * work_weeks
        results['Project Recovery Costs'] = -1 * revenue_at_risk * pct_projects_at_risk
        results['Employee Turnover Impact'] = -1 * 3 * cost_turnover
        results['Regained VP Selling Time'] = 12 * vp_hourly * work_weeks
        results['Strategic Delivery Capacity Recovered'] = 15 * head_hourly * work_weeks
        results['Project Health Improvement'] = 0.5 * revenue_at_risk * pct_projects_at_risk
        results['Improved Retention'] = 2 * cost_turnover
        results['Cost of Chief of Staff Salary'] = -1 * chief_salary
        # Totals
        results['Total Negative Impact'] = (
            results['Lost Sales (VP involvement)'] +
            results['Strategic Loss (Head of Delivery Role)'] +
            results['Project Recovery Costs'] +
            results['Employee Turnover Impact']
        )
        results['Total Positive Impact'] = (
            results['Regained VP Selling Time'] +
            results['Strategic Delivery Capacity Recovered'] +
            results['Project Health Improvement'] +
            results['Improved Retention'] +
            results['Cost of Chief of Staff Salary']
        )
        return results

    # --- Calculate both scenarios ---
    baseline_results = calculate_scenarios(baseline_inputs)
    proposed_results = calculate_scenarios(proposed_inputs)

    # --- Display results ---
    st.subheader("Scenario Results")
    categories = [
        'Lost Sales (VP involvement)',
        'Strategic Loss (Head of Delivery Role)',
        'Project Recovery Costs',
        'Employee Turnover Impact',
        'Regained VP Selling Time',
        'Strategic Delivery Capacity Recovered',
        'Project Health Improvement',
        'Improved Retention',
        'Cost of Chief of Staff Salary',
        'Total Negative Impact',
        'Total Positive Impact',
    ]
    data = []
    for cat in categories:
        base_val = baseline_results.get(cat, 0)
        prop_val = proposed_results.get(cat, 0)
        diff = prop_val - base_val
        data.append({
            "Category": cat,
            "Do Nothing": base_val,
            "Proposed": prop_val,
            "Difference": diff
        })
    # Ensure totals are at the bottom
    results_df = pd.DataFrame(data)
    results_df['sort_order'] = results_df['Category'].apply(lambda x: 2 if 'Total' in x else (1 if 'Recovered' in x or 'Improved' in x or 'Regained' in x else 0))
    results_df = results_df.sort_values(by=['sort_order', 'Category']).drop('sort_order', axis=1)
    def fmt(val):
        if isinstance(val, (int, float)):
            return f"${val:,.0f}"
        return val
    st.dataframe(results_df.style.format({"Do Nothing": fmt, "Proposed": fmt, "Difference": fmt}), use_container_width=True)

    st.subheader("🤖 Scenario AI Assistant")
    scenario_question = st.text_input("Ask about this scenario (tradeoffs, opportunity cost, etc.)...")
    if scenario_question and st.session_state.openai_client:
        with st.spinner("Thinking..."):
            context = f"Scenario Inputs:\n{proposed_inputs}\n\nScenario Results:\n{results_df.to_string(index=False)}"
            try:
                completion = st.session_state.openai_client.chat.completions.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": "You are a strategic scenario modeling assistant for healthcare delivery. Help the user analyze tradeoffs, opportunity cost, and scenario impacts."},
                        {"role": "user", "content": f"{context}\n\nQuestion: {scenario_question}"}
                    ]
                )
                response_text = completion.choices[0].message.content
            except Exception as e:
                response_text = f"Error querying OpenAI: {str(e)}"
            st.markdown(response_text)
    elif scenario_question:
        st.warning("OpenAI client not initialized. Cannot process scenario questions.")

def render_whale_hunting_page():
    st.title("🐳 Whale Hunting: Top Strategic Pursuits")
    pipeline_df = st.session_state.all_data.get('Pipeline', pd.DataFrame())

    if pipeline_df.empty:
        st.warning("Pipeline data not loaded. Cannot display whale hunting information.")
        return

    # Ensure correct data types
    pipeline_df['Percieved Annual AMO'] = indicators.safe_to_numeric(pipeline_df['Percieved Annual AMO'])
    pipeline_df['Open Pipeline_Active Work'] = indicators.safe_to_numeric(pipeline_df.get('Open Pipeline_Active Work', pd.Series(dtype=float)))
    pipeline_df['Pipeline Score'] = indicators.safe_to_numeric(pipeline_df.get('Pipeline Score', pd.Series(dtype=float)))
    pipeline_df['Pursuit Tier'] = pipeline_df['Pursuit Tier'].astype(str)

    # Filter for Tier 1 whales
    tier_1_whales = pipeline_df[pipeline_df['Pursuit Tier'].str.upper() == 'TIER 1']
    
    if tier_1_whales.empty:
        st.info("No Tier 1 opportunities found in the pipeline.")
        return

    # Sort by Percieved Annual AMO
    whales_sorted = tier_1_whales.sort_values(by='Percieved Annual AMO', ascending=False)

    num_whales_to_show = st.number_input("Number of Top Whales to Display:", min_value=1, max_value=len(whales_sorted), value=min(5, len(whales_sorted)), step=1)

    if 'Account' not in whales_sorted.columns:
        st.error("'Account' column is missing from the Pipeline data. Cannot display battle cards.")
        return

    top_n_whales = whales_sorted.head(num_whales_to_show)

    for index, row in top_n_whales.iterrows():
        account_name = row['Account']
        amo = row['Percieved Annual AMO']
        
        expander_title = f"{account_name} - AMO: {format_currency(amo)}"
        with st.expander(expander_title):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(f"**Account:** {account_name}")
                st.markdown(f"**Percieved Annual AMO:** {format_currency(amo)}")
                st.markdown(f"**Current Opportunity Value:** {format_currency(row.get('Open Pipeline_Active Work', 0))}")
                st.markdown(f"**Pursuit Tier:** {row.get('Pursuit Tier', 'N/A')}")
                st.markdown(f"**Horizon:** {row.get('Horizon', 'N/A')}")
                st.markdown(f"**Pipeline Score:** {format_number(row.get('Pipeline Score', 0), 1)}")
                st.markdown(f"**Deal Registered (Y/N):** {row.get('Deal Registered YN', 'N/A')}") # Corrected from Deal Registered

            with col2:
                st.markdown(f"**Opportunity Created:** {row.get('Opportunity Created Date', 'N/A')}")
                st.markdown(f"**Last Touchpoint:** {row.get('Last Touchpoint Date', 'N/A')}")
                st.markdown(f"**Key Client Contacts:** {row.get('Key Client Contacts', 'N/A')}")
                st.markdown(f"**Internal Pursuit Team:** {row.get('Internal Pursuit Team', 'N/A')}")
                st.markdown(f"**Win Themes:** {row.get('Win Themes', 'N/A')}")
                st.markdown(f"**Known Competitors:** {row.get('Known Competitors', 'N/A')}")

            st.markdown("**Notes:**")
            st.text_area("Notes:", value=str(row.get('Notes', '')), height=100, key=f"notes_{account_name}_{index}", disabled=True)
            st.markdown("**Actions:**")
            st.text_area("Actions:", value=str(row.get('Actions', '')), height=100, key=f"actions_{account_name}_{index}", disabled=True)
            st.markdown("**Help Needed:**")
            st.text_area("Help Needed:", value=str(row.get('Help Needed', '')), height=75, key=f"help_{account_name}_{index}", disabled=True)
            
            edit_button_key = f"edit_whale_{account_name.replace(' ','_')}_{index}"
            if st.button(f"✏️ Edit {account_name}", key=edit_button_key):
                st.session_state.selected_pipeline_to_edit = account_name
                st.session_state.manage_data_entity_type = "Pipeline Opportunity"
                st.session_state.current_page = "📝 Manage Data"
                st.rerun()

def render_staffing_health_page():
    st.title("🧑‍💻 Project Staffing Health")
    
    project_df = st.session_state.all_data.get('Project Inventory', pd.DataFrame())
    util_df = st.session_state.all_data.get('Team Utilization', pd.DataFrame())

    if project_df.empty:
        st.warning("Project Inventory data not loaded. Cannot display staffing health.")
        return
    
    if util_df.empty:
        st.info("Team Utilization data not available. Staffing details may be incomplete.")
        # We can still show projects, but with missing team data

    # --- 1. Filter for Active Projects ---
    project_df_c = project_df.copy()
    if 'Project End Date' not in project_df_c.columns:
        st.error("'Project End Date' column missing from Project Inventory. Cannot determine active projects.")
        return
    project_df_c['Project End Date'] = pd.to_datetime(project_df_c['Project End Date'], errors='coerce')
    current_date = pd.to_datetime(date.today())
    active_projects_df = project_df_c[
        (project_df_c['Project End Date'].isna()) | (project_df_c['Project End Date'] >= current_date)
    ]

    if active_projects_df.empty:
        st.info("No active projects found.")
        return

    # --- 2. Prepare Utilization Data ---
    util_df_c = util_df.copy()
    if not util_df_c.empty:
        if 'Utilization (%)' in util_df_c.columns:
            util_df_c['Utilization (%)'] = indicators.safe_to_numeric(util_df_c['Utilization (%)'])
        else:
            util_df_c['Utilization (%)'] = 0 # Default if missing
            st.warning("'Utilization (%)' column missing in Team Utilization.")

        if 'Latest Pulse Score' in util_df_c.columns:
            util_df_c['Latest Pulse Score'] = indicators.safe_to_numeric(util_df_c['Latest Pulse Score'])
        else:
            util_df_c['Latest Pulse Score'] = 0 # Default if missing
            st.warning("'Latest Pulse Score' column missing in Team Utilization.")
        
        if 'Project Assignments' not in util_df_c.columns:
            st.warning("'Project Assignments' column missing in Team Utilization. Cannot link staff to projects.")
            util_df_c = pd.DataFrame() # Effectively disable staff lookups
        else:
             util_df_c['Project Assignments'] = util_df_c['Project Assignments'].astype(str).fillna('')

    staffing_health_data = []
    required_project_cols = ['Project Name', 'Client', 'Status (R/Y/G)', 'Team Resourcing', 'Project End Date']
    if not all(col in active_projects_df.columns for col in required_project_cols):
        st.error(f"One or more required columns ({required_project_cols}) are missing from Project Inventory.")
        return

    # --- 3. Combine Data for Each Active Project ---
    for index, project_row in active_projects_df.iterrows():
        project_name = project_row['Project Name']
        assigned_team_members = []
        team_utilizations = []
        team_pulse_scores = []

        if not util_df_c.empty and 'Employee Name' in util_df_c.columns:
            for _, staff_row in util_df_c.iterrows():
                # Split by comma, strip whitespace from each project name
                assigned_projects = [p.strip() for p in staff_row.get('Project Assignments', '').split(',')]
                if project_name in assigned_projects:
                    assigned_team_members.append(staff_row['Employee Name'])
                    if 'Utilization (%)' in staff_row and pd.notna(staff_row['Utilization (%)']):
                        team_utilizations.append(staff_row['Utilization (%)'])
                    if 'Latest Pulse Score' in staff_row and pd.notna(staff_row['Latest Pulse Score']):
                        team_pulse_scores.append(staff_row['Latest Pulse Score'])
        
        avg_utilization = sum(team_utilizations) / len(team_utilizations) if team_utilizations else 0
        avg_pulse = sum(team_pulse_scores) / len(team_pulse_scores) if team_pulse_scores else 0

        staffing_health_data.append({
            'Project Name': project_name,
            'Client': project_row['Client'],
            'Status (R/Y/G)': project_row['Status (R/Y/G)'],
            'Team Resourcing': project_row['Team Resourcing'],
            'Assigned Team Size': len(assigned_team_members),
            'Assigned Team Members': ", ".join(assigned_team_members) if assigned_team_members else 'N/A',
            'Avg. Team Utilization %': avg_utilization,
            'Avg. Team Pulse Score': avg_pulse,
            'Project End Date': project_row['Project End Date'].strftime('%Y-%m-%d') if pd.notna(project_row['Project End Date']) else 'N/A'
        })

    if not staffing_health_data:
        st.info("No staffing health data to display for active projects.")
        return

    display_df = pd.DataFrame(staffing_health_data)
    
    # --- 4. Display Data with Styling and Filters ---
    st.subheader("Active Projects Staffing Overview")

    # Filters
    client_list = ["All"] + sorted(display_df['Client'].unique().tolist())
    selected_client = st.selectbox("Filter by Client:", client_list)

    # Handle potential NA values in Team Resourcing for sorting
    unique_resourcing_statuses = display_df['Team Resourcing'].unique()
    # Convert NA to a string like 'N/A' or filter them out, then convert all to string before sorting
    resourcing_status_list = ["All"] + sorted([str(status) if pd.notna(status) else "N/A" for status in unique_resourcing_statuses])
    # Remove duplicates that might arise if 'N/A' was already a string value
    resourcing_status_list = sorted(list(set(resourcing_status_list)))
    # Ensure "All" is first if it got re-sorted by set operation
    if "All" in resourcing_status_list: 
        resourcing_status_list.remove("All")
        resourcing_status_list.insert(0, "All")
        
    selected_resourcing = st.selectbox("Filter by Team Resourcing Status:", resourcing_status_list)

    filtered_df = display_df.copy()
    if selected_client != "All":
        filtered_df = filtered_df[filtered_df['Client'] == selected_client]
    if selected_resourcing != "All":
        filtered_df = filtered_df[filtered_df['Team Resourcing'] == selected_resourcing]

    # Styling function for Team Resourcing
    def style_team_resourcing(val):
        val_lower = str(val).lower()
        color = 'white' # Default
        if val_lower == 'yes':
            color = 'lightgreen'
        elif val_lower == 'some gaps':
            color = 'lightyellow'
        elif val_lower in ['understaffed', 'misaligned', 'no core team']:
            color = '#ffcccb' # Light red
        return f'background-color: {color}'

    # Format numbers in the DataFrame before styling
    if not filtered_df.empty:
        formatted_df = filtered_df.copy()
        formatted_df['Avg. Team Utilization %'] = formatted_df['Avg. Team Utilization %'].apply(lambda x: format_percentage(x, 1, default_na='N/A'))
        formatted_df['Avg. Team Pulse Score'] = formatted_df['Avg. Team Pulse Score'].apply(lambda x: format_number(x, 1, default_na='N/A'))

        st.dataframe(
            formatted_df.style.applymap(style_team_resourcing, subset=['Team Resourcing']),
            use_container_width=True
        )
    else:
        st.info("No projects match the selected filter criteria.")

# --- Main Application ---
PAGES = {
    "🏠 Home": render_home_dashboard,
    "📊 Projects": render_projects_page,
    "📈 Pipeline": render_pipeline_page,
    "⚠️ Risks": render_risks_page,
    "👥 Team & Ops": render_team_ops_page,
    "📝 Manage Data": render_manage_data_page, 
    "⚙️ Scenario Modeling": render_scenario_modeling_page,
    "🔍 Data Explorer": render_data_explorer_page,
    "🧪 Scenario Playground": render_scenario_playground_page,
    "🐳 Whale Hunting": render_whale_hunting_page,
    "🧑‍💻 Project Staffing Health": render_staffing_health_page,
}

def main():
    init_session_state()
    
    st.sidebar.image("https://mma.prnewswire.com/media/1677414/Hakkoda_Logo.jpg?p=facebook", width=200)
    st.sidebar.title("Healthcare Delivery OS")
    st.sidebar.markdown("---")
    
    st.sidebar.subheader("Navigation")
    
    current_page_key_index = list(PAGES.keys()).index(st.session_state.current_page) if st.session_state.current_page in PAGES else 0
    
    st.session_state.current_page = st.sidebar.radio(
        "Go to", list(PAGES.keys()), index=current_page_key_index, key="navigation_radio" 
    )
    st.sidebar.markdown("---")

    if not st.session_state.data_loaded :
        load_all_data() 
    
    if st.sidebar.button("🔄 Refresh Data"):
        st.session_state.data_loaded = False 
        st.cache_data.clear() 
        st.cache_resource.clear() 
        st.session_state.initial_load_complete = False 
        load_all_data() 
        st.rerun()

    render_ai_assistant()
    
    if st.session_state.data_loaded and st.session_state.indicators:
        page_function = PAGES.get(st.session_state.current_page) 
        if page_function: page_function()
        else: st.error("Selected page not found. Defaulting to Home."); render_home_dashboard() 
    elif not st.session_state.data_loaded: st.warning("Data is loading or connection failed. Please wait or check console for errors.")
    else: st.warning("Indicators are not yet calculated. Please wait or refresh data.")

if __name__ == "__main__":
    main()