import pandas as pd
from datetime import datetime, timedelta
import re
from strategic_targets import (
    REVENUE_TARGET, REVENUE_STRETCH_GOAL,
    GREEN_PROJECT_TARGET, EMPLOYEE_PULSE_TARGET, PIPELINE_COVERAGE_TARGET,
    SPONSOR_CHECKIN_WINDOW_DAYS, NEXT_DEAL_DISCUSSION_THRESHOLD_DAYS, CUSTOMER_NPS_TARGET,
    PROJECT_SCORE_BANDS, PIPELINE_SCORE_BANDS
)

# --- Helper Functions ---
def safe_to_numeric(series, Ttype=float, remove_chars=r'[$,%()A-Za-z]'):
    if series is None:
        return pd.Series(dtype=Ttype)
    if not isinstance(series, pd.Series):
        series = pd.Series(series)
    
    if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_object_dtype(series):
        return series.astype(Ttype)

    cleaned_series = series.astype(str).str.replace(remove_chars, '', regex=True).str.strip()
    cleaned_series = cleaned_series.replace('', '0') 
    return pd.to_numeric(cleaned_series, errors='coerce').fillna(0).astype(Ttype)


def band_score(score, bands):
    if pd.isna(score):
        return 'N/A'
    for band, (low, high) in bands.items():
        if low <= score <= high:
            return band
    return 'N/A' 


def score_band_distribution(scores, bands):
    band_counts = {band: 0 for band in bands}
    band_counts['N/A'] = 0 
    valid_scores = 0
    for s in scores:
        band = band_score(s, bands)
        if band in band_counts:
            band_counts[band] += 1
        if band != 'N/A':
            valid_scores +=1
            
    total_valid = valid_scores
    band_pct = {band: (count / total_valid * 100 if total_valid > 0 else 0) for band, count in band_counts.items() if band != 'N/A'}
    return band_counts, band_pct

# --- Main Indicator Functions ---

def get_general_and_project_kpis(data, kpis=None):
    if kpis is None:
        kpis = {}
    project_df = data.get('Project Inventory')
    
    project_name_col = 'Project Name' 

    if project_df is not None and not project_df.empty:
        project_df = project_df.copy()
        
        # Ensure critical columns exist or handle gracefully
        if 'Revenue' in project_df.columns:
            project_df['Revenue'] = safe_to_numeric(project_df['Revenue'])
        else:
            project_df['Revenue'] = 0 # Default if column missing
            print("Warning: 'Revenue' column missing in Project Inventory.")

        if 'Status (R/Y/G)' in project_df.columns:
            project_df['Status (R/Y/G)'] = project_df['Status (R/Y/G)'].astype(str).str.strip().str.upper()
        else:
            project_df['Status (R/Y/G)'] = 'UNKNOWN' # Default if column missing
            print("Warning: 'Status (R/Y/G)' column missing in Project Inventory.")

        kpis['total_projects'] = len(project_df)
        
        red_projects = project_df[project_df['Status (R/Y/G)'] == 'R']
        kpis['red_projects_count'] = len(red_projects)
        kpis['red_project_revenue'] = red_projects['Revenue'].sum()
        
        yellow_projects = project_df[project_df['Status (R/Y/G)'] == 'Y']
        kpis['yellow_projects_count'] = len(yellow_projects)
        
        green_projects = project_df[project_df['Status (R/Y/G)'] == 'G']
        kpis['green_projects_count'] = len(green_projects)
        
        total_revenue = project_df['Revenue'].sum()
        kpis['total_revenue'] = total_revenue
        kpis['revenue_vs_target_pct'] = (total_revenue / REVENUE_TARGET * 100) if REVENUE_TARGET else 0
        kpis['revenue_vs_stretch_pct'] = (total_revenue / REVENUE_STRETCH_GOAL * 100) if REVENUE_STRETCH_GOAL else 0
        
        has_project_name_col = project_name_col in project_df.columns

        if 'Project Health Score' in project_df.columns:
            scores = safe_to_numeric(project_df['Project Health Score']).dropna()
            kpis['avg_project_health_score'] = scores.mean() if not scores.empty else 0
            kpis['median_project_health_score'] = scores.median() if not scores.empty else 0
            _, band_pct = score_band_distribution(scores, PROJECT_SCORE_BANDS)
            kpis['project_health_score_bands_pct'] = band_pct
            if not scores.empty and has_project_name_col:
                kpis['top_project_by_health_score'] = project_df.loc[scores.idxmax(), project_name_col]
                kpis['bottom_project_by_health_score'] = project_df.loc[scores.idxmin(), project_name_col]
            else:
                kpis['top_project_by_health_score'] = "N/A"
                kpis['bottom_project_by_health_score'] = "N/A"
                if not has_project_name_col and not scores.empty:
                    print(f"Warning: Column '{project_name_col}' not found in Project Inventory for top/bottom health score.")
        else:
            kpis.update({'avg_project_health_score': 0, 'median_project_health_score': 0, 'project_health_score_bands_pct': {}, 'top_project_by_health_score': "N/A", 'bottom_project_by_health_score': "N/A"})
            print("Warning: 'Project Health Score' column missing in Project Inventory.")
            
        if 'Total Project Score' in project_df.columns:
            total_scores = safe_to_numeric(project_df['Total Project Score']).dropna()
            kpis['avg_total_project_score'] = total_scores.mean() if not total_scores.empty else 0
            kpis['median_total_project_score'] = total_scores.median() if not total_scores.empty else 0
            _, band_pct_total = score_band_distribution(total_scores, PROJECT_SCORE_BANDS)
            kpis['total_project_score_bands_pct'] = band_pct_total
            if not total_scores.empty and has_project_name_col:
                kpis['top_project_by_total_score'] = project_df.loc[total_scores.idxmax(), project_name_col]
                kpis['bottom_project_by_total_score'] = project_df.loc[total_scores.idxmin(), project_name_col]
            else:
                kpis['top_project_by_total_score'] = "N/A"
                kpis['bottom_project_by_total_score'] = "N/A"
                if not has_project_name_col and not total_scores.empty:
                     print(f"Warning: Column '{project_name_col}' not found in Project Inventory for top/bottom total score.")
        else:
            kpis.update({'avg_total_project_score': 0, 'median_total_project_score': 0, 'total_project_score_bands_pct': {}, 'top_project_by_total_score': "N/A", 'bottom_project_by_total_score': "N/A"})
            print("Warning: 'Total Project Score' column missing in Project Inventory.")
    else: # project_df is None or empty
        kpis.update({
            'total_projects': 0, 'red_projects_count': 0, 'red_project_revenue': 0,
            'yellow_projects_count': 0, 'green_projects_count': 0, 'total_revenue': 0,
            'revenue_vs_target_pct': 0, 'revenue_vs_stretch_pct': 0,
            'avg_project_health_score': 0, 'median_project_health_score': 0,
            'project_health_score_bands_pct': {}, 'top_project_by_health_score': "N/A", 'bottom_project_by_health_score': "N/A",
            'avg_total_project_score': 0, 'median_total_project_score': 0,
            'total_project_score_bands_pct': {}, 'top_project_by_total_score': "N/A", 'bottom_project_by_total_score': "N/A"
        })
    return kpis

def get_pipeline_and_risk_kpis(data, kpis=None):
    if kpis is None:
        kpis = {}
    pipeline_df = data.get('Pipeline')
    risk_df = data.get('Project Risks')

    pipeline_account_col = 'Account' # Adjusted to 'Account'

    if pipeline_df is not None and not pipeline_df.empty:
        pipeline_df = pipeline_df.copy()

        if 'Open Pipeline_Active Work' in pipeline_df.columns:
            pipeline_df['Open Pipeline_Active Work'] = safe_to_numeric(pipeline_df['Open Pipeline_Active Work'])
        else:
            pipeline_df['Open Pipeline_Active Work'] = 0
            print("Warning: 'Open Pipeline_Active Work' column missing in Pipeline.")
        
        if 'Percieved Annual AMO' in pipeline_df.columns:
            pipeline_df['Percieved Annual AMO'] = safe_to_numeric(pipeline_df['Percieved Annual AMO'])
        else:
            pipeline_df['Percieved Annual AMO'] = 0
            print("Warning: 'Percieved Annual AMO' column missing in Pipeline.")


        kpis['active_pipeline_value'] = pipeline_df['Open Pipeline_Active Work'].sum()
        kpis['total_potential_pipeline_value'] = pipeline_df['Percieved Annual AMO'].sum()
        
        kpis['pipeline_coverage_ratio'] = (kpis['active_pipeline_value'] / REVENUE_TARGET) if REVENUE_TARGET else 0
        kpis['pipeline_coverage_vs_target_pct'] = (kpis['pipeline_coverage_ratio'] / PIPELINE_COVERAGE_TARGET * 100) if PIPELINE_COVERAGE_TARGET else 0
        
        has_pipeline_account_col = pipeline_account_col in pipeline_df.columns

        if 'Pipeline Score' in pipeline_df.columns:
            scores = safe_to_numeric(pipeline_df['Pipeline Score']).dropna()
            kpis['avg_pipeline_score'] = scores.mean() if not scores.empty else 0
            kpis['median_pipeline_score'] = scores.median() if not scores.empty else 0
            _, band_pct = score_band_distribution(scores, PIPELINE_SCORE_BANDS)
            kpis['pipeline_score_bands_pct'] = band_pct
            if not scores.empty and has_pipeline_account_col:
                 kpis['top_pipeline_by_score'] = pipeline_df.loc[scores.idxmax(), pipeline_account_col]
                 kpis['bottom_pipeline_by_score'] = pipeline_df.loc[scores.idxmin(), pipeline_account_col]
            else:
                 kpis['top_pipeline_by_score'] = "N/A"
                 kpis['bottom_pipeline_by_score'] = "N/A"
                 if not has_pipeline_account_col and not scores.empty:
                     print(f"Warning: Column '{pipeline_account_col}' not found in Pipeline data for top/bottom score calculation.")
        else:
            kpis.update({'avg_pipeline_score': 0, 'median_pipeline_score': 0, 'pipeline_score_bands_pct': {}, 'top_pipeline_by_score': "N/A", 'bottom_pipeline_by_score': "N/A"})
            print("Warning: 'Pipeline Score' column missing in Pipeline.")

        if 'Total Deal Score' in pipeline_df.columns:
            total_scores = safe_to_numeric(pipeline_df['Total Deal Score']).dropna()
            kpis['avg_total_deal_score'] = total_scores.mean() if not total_scores.empty else 0
            kpis['median_total_deal_score'] = total_scores.median() if not total_scores.empty else 0
            _, band_pct_total = score_band_distribution(total_scores, PIPELINE_SCORE_BANDS)
            kpis['total_deal_score_bands_pct'] = band_pct_total
            if not total_scores.empty and has_pipeline_account_col:
                 kpis['top_pipeline_by_total_score'] = pipeline_df.loc[total_scores.idxmax(), pipeline_account_col]
                 kpis['bottom_pipeline_by_total_score'] = pipeline_df.loc[total_scores.idxmin(), pipeline_account_col]
            else:
                 kpis['top_pipeline_by_total_score'] = "N/A"
                 kpis['bottom_pipeline_by_total_score'] = "N/A"
                 if not has_pipeline_account_col and not total_scores.empty: 
                     print(f"Warning: Column '{pipeline_account_col}' not found in Pipeline data for top/bottom total score calculation.")
        else:
            kpis.update({'avg_total_deal_score': 0, 'median_total_deal_score': 0, 'total_deal_score_bands_pct': {}, 'top_pipeline_by_total_score': "N/A", 'bottom_pipeline_by_total_score': "N/A"})
            print("Warning: 'Total Deal Score' column missing in Pipeline.")
    else: # pipeline_df is None or empty
        kpis.update({
            'active_pipeline_value': 0, 'total_potential_pipeline_value': 0,
            'pipeline_coverage_ratio': 0, 'pipeline_coverage_vs_target_pct': 0,
            'avg_pipeline_score': 0, 'median_pipeline_score': 0, 'pipeline_score_bands_pct': {},
            'top_pipeline_by_score': "N/A", 'bottom_pipeline_by_score': "N/A",
            'avg_total_deal_score': 0, 'median_total_deal_score': 0, 'total_deal_score_bands_pct': {},
            'top_pipeline_by_total_score': "N/A", 'bottom_pipeline_by_total_score': "N/A"
        })

    # Risk Metrics
    if risk_df is not None and not risk_df.empty:
        risk_df = risk_df.copy()
        if 'Impact ($)' in risk_df.columns:
            risk_df['Impact ($)'] = safe_to_numeric(risk_df['Impact ($)'])
        else:
            risk_df['Impact ($)'] = 0
            print("Warning: 'Impact ($)' column missing in Project Risks.")
        
        if 'Severity' in risk_df.columns:
            risk_df['Severity'] = risk_df['Severity'].astype(str).fillna('').str.lower()
        else:
            risk_df['Severity'] = 'unknown'
            print("Warning: 'Severity' column missing in Project Risks.")


        high_risks = risk_df[risk_df['Severity'] == 'high']
        kpis['high_severity_risk_count'] = len(high_risks)
        kpis['high_severity_risk_impact'] = high_risks['Impact ($)'].sum()
        kpis['total_risk_impact'] = risk_df['Impact ($)'].sum()
        kpis['high_risk_impact_as_pct_of_total'] = (kpis['high_severity_risk_impact'] / kpis['total_risk_impact'] * 100) if kpis['total_risk_impact'] else 0
    else: # risk_df is None or empty
        kpis.update({
            'high_severity_risk_count': 0, 'high_severity_risk_impact': 0,
            'total_risk_impact': 0, 'high_risk_impact_as_pct_of_total': 0
        })
    return kpis

def get_satisfaction_and_efficiency_kpis(data, kpis=None):
    if kpis is None:
        kpis = {}
    project_df = data.get('Project Inventory')
    pipeline_df = data.get('Pipeline')
    util_df = data.get('Team Utilization')
    exec_activity_df = data.get('Executive Activity')
    
    project_name_col = 'Project Name'

    # Customer NPS (eNPS from Project Inventory)
    if project_df is not None and not project_df.empty and 'eNPS' in project_df.columns:
        enps_scores = safe_to_numeric(project_df['eNPS']).dropna()
        kpis['avg_customer_nps'] = enps_scores.mean() if not enps_scores.empty else 0
        kpis['customer_nps_vs_target_pct'] = (kpis['avg_customer_nps'] / CUSTOMER_NPS_TARGET * 100) if CUSTOMER_NPS_TARGET and kpis['avg_customer_nps'] is not None and kpis['avg_customer_nps'] != 0 else 0
    else:
        kpis['avg_customer_nps'] = 0
        kpis['customer_nps_vs_target_pct'] = 0
        if project_df is None or project_df.empty: print("Info: Project Inventory data not available for eNPS.")
        elif 'eNPS' not in project_df.columns: print("Warning: 'eNPS' column missing in Project Inventory for eNPS.")

    # Employee Satisfaction (Pulse Score from Team Utilization)
    if util_df is not None and not util_df.empty and 'Latest Pulse Score' in util_df.columns:
        pulse_scores = safe_to_numeric(util_df['Latest Pulse Score']).dropna()
        kpis['avg_employee_pulse_score'] = pulse_scores.mean() if not pulse_scores.empty else 0
        kpis['employee_pulse_vs_target_pct'] = (kpis['avg_employee_pulse_score'] / EMPLOYEE_PULSE_TARGET * 100) if EMPLOYEE_PULSE_TARGET and kpis['avg_employee_pulse_score'] is not None and kpis['avg_employee_pulse_score'] != 0 else 0
    else:
        kpis['avg_employee_pulse_score'] = 0
        kpis['employee_pulse_vs_target_pct'] = 0
        if util_df is None or util_df.empty: print("Info: Team Utilization data not available for Pulse Score.")
        elif 'Latest Pulse Score' not in util_df.columns: print("Warning: 'Latest Pulse Score' column missing in Team Utilization.")
        
    # Deal Cycle Time
    if pipeline_df is not None and not pipeline_df.empty and 'Opportunity Created Date' in pipeline_df.columns and 'Closed Won Date' in pipeline_df.columns:
        pipeline_df_copy = pipeline_df.copy()
        pipeline_df_copy['Opportunity Created Date'] = pd.to_datetime(pipeline_df_copy['Opportunity Created Date'], errors='coerce')
        pipeline_df_copy['Closed Won Date'] = pd.to_datetime(pipeline_df_copy['Closed Won Date'], errors='coerce')
        won_deals = pipeline_df_copy.dropna(subset=['Opportunity Created Date', 'Closed Won Date'])
        if not won_deals.empty:
            won_deals['Cycle Time'] = (won_deals['Closed Won Date'] - won_deals['Opportunity Created Date']).dt.days
            cycle_times = won_deals['Cycle Time'][won_deals['Cycle Time'] >= 0] 
            kpis['avg_deal_cycle_time_days'] = cycle_times.mean() if not cycle_times.empty else 0
            kpis['median_deal_cycle_time_days'] = cycle_times.median() if not cycle_times.empty else 0
            if 'Pursuit Tier' in won_deals.columns:
                kpis['deal_cycle_time_by_tier'] = won_deals.groupby('Pursuit Tier')['Cycle Time'].agg(['mean', 'median']).to_dict('index')
            else:
                kpis['deal_cycle_time_by_tier'] = {}
        else: # No won deals with both dates
            kpis['avg_deal_cycle_time_days'] = 0
            kpis['median_deal_cycle_time_days'] = 0
            kpis['deal_cycle_time_by_tier'] = {}
    else: # Pipeline df or required date columns missing
        kpis['avg_deal_cycle_time_days'] = 0
        kpis['median_deal_cycle_time_days'] = 0
        kpis['deal_cycle_time_by_tier'] = {}
        if pipeline_df is None or pipeline_df.empty: print("Info: Pipeline data not available for Deal Cycle Time.")
        else:
            if 'Opportunity Created Date' not in pipeline_df.columns: print("Warning: 'Opportunity Created Date' missing in Pipeline.")
            if 'Closed Won Date' not in pipeline_df.columns: print("Warning: 'Closed Won Date' missing in Pipeline.")

    # Time Between Project End and Next Deal Discussion
    required_cols_next_deal = [project_name_col, 'Project End Date', 'Next Opp First Discussion Date']
    if project_df is not None and not project_df.empty and all(col in project_df.columns for col in required_cols_next_deal):
        project_df_copy = project_df.copy()

        project_df_copy['Project End Date'] = pd.to_datetime(project_df_copy['Project End Date'], errors='coerce')
        project_df_copy['Next Opp First Discussion Date'] = pd.to_datetime(project_df_copy['Next Opp First Discussion Date'], errors='coerce')
        
        completed_projects_with_next_opp = project_df_copy.dropna(subset=['Project End Date', 'Next Opp First Discussion Date'])
        if not completed_projects_with_next_opp.empty:
            valid_end_dates_mask = completed_projects_with_next_opp['Project End Date'].notna()
            if valid_end_dates_mask.any():
                df_calc = completed_projects_with_next_opp[valid_end_dates_mask].copy() # Use .copy() to avoid SettingWithCopyWarning
                df_calc['Next Deal Gap'] = (df_calc['Next Opp First Discussion Date'] - df_calc['Project End Date']).dt.days
                next_deal_gaps = df_calc['Next Deal Gap'][df_calc['Next Deal Gap'] >=0]
                kpis['avg_next_deal_gap_days'] = next_deal_gaps.mean() if not next_deal_gaps.empty else 0
            else:
                kpis['avg_next_deal_gap_days'] = 0
        else:
            kpis['avg_next_deal_gap_days'] = 0
        
        ended_projects = project_df_copy[project_df_copy['Project End Date'].notna()] 
        if not ended_projects.empty:
            current_date_naive = pd.to_datetime('today').normalize() 

            overdue_next_deal_df = ended_projects[
                ((current_date_naive - ended_projects['Project End Date']).dt.days > NEXT_DEAL_DISCUSSION_THRESHOLD_DAYS) & 
                ended_projects['Next Opp First Discussion Date'].isna() & 
                (ended_projects['Project End Date'] < current_date_naive) 
            ]
            kpis['overdue_next_deal_discussion_count'] = len(overdue_next_deal_df)
            # Ensure 'Project Name' column exists for the list
            if project_name_col in overdue_next_deal_df.columns:
                 kpis['overdue_next_deal_projects_list'] = overdue_next_deal_df[[project_name_col, 'Project End Date']].to_dict('records')
            else:
                 kpis['overdue_next_deal_projects_list'] = overdue_next_deal_df[['Project End Date']].to_dict('records') # Fallback if 'Project Name' is missing
                 print(f"Warning: '{project_name_col}' column missing for overdue_next_deal_projects_list.")
        else:
            kpis['overdue_next_deal_discussion_count'] = 0
            kpis['overdue_next_deal_projects_list'] = []
    else: 
        kpis['avg_next_deal_gap_days'] = 0
        kpis['overdue_next_deal_discussion_count'] = 0
        kpis['overdue_next_deal_projects_list'] = []
        if project_df is None or project_df.empty: print("Info: Project Inventory data not available for Next Deal Gap.")
        else: 
            for col in required_cols_next_deal:
                if col not in project_df.columns: print(f"Warning: Column '{col}' missing in Project Inventory for Next Deal Gap.")


    # Meaningful Sponsor Check-ins
    required_cols_checkin = [project_name_col, 'Last Sponsor Checkin Date', 'Sponsor Checkin Notes', 'Project End Date']
    if project_df is not None and not project_df.empty and all(col in project_df.columns for col in required_cols_checkin):
        project_df_copy = project_df.copy()
        project_df_copy['Last Sponsor Checkin Date'] = pd.to_datetime(project_df_copy['Last Sponsor Checkin Date'], errors='coerce')
        project_df_copy['Project End Date'] = pd.to_datetime(project_df_copy['Project End Date'], errors='coerce')
        current_date_naive = pd.to_datetime('today').normalize()
        
        active_projects_df = project_df_copy[
            project_df_copy['Project End Date'].isna() | 
            (project_df_copy['Project End Date'] >= current_date_naive)
        ].copy() # Use .copy() here
        total_active_projects = len(active_projects_df)

        if total_active_projects > 0:
            active_projects_df['Sponsor Checkin Notes'] = active_projects_df['Sponsor Checkin Notes'].astype(str) # Ensure string for strip
            recent_checkins_df = active_projects_df[
                (active_projects_df['Last Sponsor Checkin Date'].notna()) &
                (active_projects_df['Last Sponsor Checkin Date'] >= current_date_naive - pd.Timedelta(days=SPONSOR_CHECKIN_WINDOW_DAYS)) &
                (active_projects_df['Sponsor Checkin Notes'].str.strip() != '')
            ]
            kpis['recent_meaningful_checkins_count'] = len(recent_checkins_df)
            kpis['recent_meaningful_checkins_pct'] = (kpis['recent_meaningful_checkins_count'] / total_active_projects * 100)
            
            overdue_checkin_df = active_projects_df[
                ((active_projects_df['Last Sponsor Checkin Date'].isna()) |
                (active_projects_df['Last Sponsor Checkin Date'] < current_date_naive - pd.Timedelta(days=SPONSOR_CHECKIN_WINDOW_DAYS)))
            ]
            kpis['overdue_sponsor_checkin_count'] = len(overdue_checkin_df)
            if project_name_col in overdue_checkin_df.columns:
                kpis['overdue_checkin_projects_list'] = overdue_checkin_df[[project_name_col, 'Last Sponsor Checkin Date']].to_dict('records')
            else:
                kpis['overdue_checkin_projects_list'] = overdue_checkin_df[['Last Sponsor Checkin Date']].to_dict('records')
                print(f"Warning: '{project_name_col}' column missing for overdue_checkin_projects_list.")

        else: # No active projects
            kpis['recent_meaningful_checkins_count'] = 0
            kpis['recent_meaningful_checkins_pct'] = 0
            kpis['overdue_sponsor_checkin_count'] = 0
            kpis['overdue_checkin_projects_list'] = []
    else:
        kpis['recent_meaningful_checkins_count'] = 0
        kpis['recent_meaningful_checkins_pct'] = 0
        kpis['overdue_sponsor_checkin_count'] = 0
        kpis['overdue_checkin_projects_list'] = []
        if project_df is None or project_df.empty: print("Info: Project Inventory data not available for Sponsor Check-ins.")
        else:
            for col in required_cols_checkin:
                if col not in project_df.columns: print(f"Warning: Column '{col}' missing in Project Inventory for Sponsor Check-ins.")


    # Green Project Ratio
    required_cols_green_ratio = [project_name_col, 'Status (R/Y/G)', 'Key Issues']
    if project_df is not None and not project_df.empty and 'Status (R/Y/G)' in project_df.columns: # Key Issues is optional for list
        # Status (R/Y/G) already processed in get_general_and_project_kpis
        total_projects = kpis.get('total_projects', 0) 
        green_projects_count = kpis.get('green_projects_count', 0)
        
        kpis['green_project_ratio'] = (green_projects_count / total_projects) if total_projects > 0 else 0
        kpis['green_project_ratio_vs_target_pct'] = (kpis['green_project_ratio'] / GREEN_PROJECT_TARGET * 100) if GREEN_PROJECT_TARGET else 0
        
        non_green_df = project_df[project_df['Status (R/Y/G)'].isin(['R', 'Y'])]
        
        list_cols = [project_name_col, 'Status (R/Y/G)']
        if 'Key Issues' in non_green_df.columns:
            list_cols.append('Key Issues')
        elif 'Key Issues' not in project_df.columns : # Check original df
            print("Info: 'Key Issues' column missing in Project Inventory for non-green projects list.")
        
        if project_name_col in non_green_df.columns:
            kpis['non_green_projects_list'] = non_green_df[list_cols].to_dict('records')
        else: # If project name is missing, this list is less useful but won't error
            kpis['non_green_projects_list'] = non_green_df[[col for col in list_cols if col != project_name_col]].to_dict('records')
            print(f"Warning: '{project_name_col}' column missing for non_green_projects_list.")
    else:
        kpis['green_project_ratio'] = 0
        kpis['green_project_ratio_vs_target_pct'] = 0
        kpis['non_green_projects_list'] = []
        if project_df is None or project_df.empty: print("Info: Project Inventory data not available for Green Project Ratio.")
        elif 'Status (R/Y/G)' not in project_df.columns: print("Warning: 'Status (R/Y/G)' column missing in Project Inventory.")


    # Team Utilization
    required_cols_util = ['Role', 'Utilization (%)']
    if util_df is not None and not util_df.empty and all(col in util_df.columns for col in required_cols_util):
        util_df_copy = util_df.copy()
        util_df_copy['Utilization (%)'] = safe_to_numeric(util_df_copy['Utilization (%)'])
        
        # Ensure 'Role' is string for .str.contains
        util_df_copy['Role'] = util_df_copy['Role'].astype(str)
        exec_df = util_df_copy[util_df_copy['Role'].str.contains('Executive', case=False, na=False)]
        delivery_df = util_df_copy[~util_df_copy['Role'].str.contains('Executive', case=False, na=False)]
        
        kpis['avg_exec_utilization_pct'] = exec_df['Utilization (%)'].mean() if not exec_df.empty else 0
        kpis['avg_delivery_utilization_pct'] = delivery_df['Utilization (%)'].mean() if not delivery_df.empty else 0
        kpis['over_utilized_execs_count'] = len(exec_df[exec_df['Utilization (%)'] > 70]) if not exec_df.empty else 0 
        kpis['under_utilized_delivery_count'] = len(delivery_df[delivery_df['Utilization (%)'] < 70]) if not delivery_df.empty else 0 
        kpis['over_utilized_delivery_count'] = len(delivery_df[delivery_df['Utilization (%)'] > 100]) if not delivery_df.empty else 0 
    else:
        kpis['avg_exec_utilization_pct'] = 0
        kpis['avg_delivery_utilization_pct'] = 0
        kpis['over_utilized_execs_count'] = 0
        kpis['under_utilized_delivery_count'] = 0
        kpis['over_utilized_delivery_count'] = 0
        if util_df is None or util_df.empty: print("Info: Team Utilization data not available.")
        else:
            for col in required_cols_util:
                if col not in util_df.columns: print(f"Warning: Column '{col}' missing in Team Utilization.")
        
    # Strategic Costs
    if exec_activity_df is not None and not exec_activity_df.empty and 'Strategic Cost ($)' in exec_activity_df.columns:
        exec_activity_df_copy = exec_activity_df.copy()
        exec_activity_df_copy['Strategic Cost ($)'] = safe_to_numeric(exec_activity_df_copy['Strategic Cost ($)'])
        kpis['total_strategic_cost'] = exec_activity_df_copy['Strategic Cost ($)'].sum()
        kpis['total_strategic_activities_count'] = len(exec_activity_df_copy)
    else:
        kpis['total_strategic_cost'] = 0
        kpis['total_strategic_activities_count'] = 0
        if exec_activity_df is None or exec_activity_df.empty: print("Info: Executive Activity data not available.")
        elif 'Strategic Cost ($)' not in exec_activity_df.columns : print("Warning: 'Strategic Cost ($)' column missing in Executive Activity.")
        
    return kpis


def get_all_indicators(data):
    kpis = {}
    kpis = get_general_and_project_kpis(data, kpis)
    kpis = get_pipeline_and_risk_kpis(data, kpis)
    kpis = get_satisfaction_and_efficiency_kpis(data, kpis)
    return kpis

def get_top3_action_items(data, openai_client, data_context_string):
    if not openai_client:
        return "OpenAI client not configured. Cannot generate action items."
        
    max_context_len = 100000 
    if len(data_context_string) > max_context_len:
        data_context_string = data_context_string[:max_context_len] + "\n... (data truncated)"

    prompt = f"""
You are a business operations assistant for a healthcare delivery organization. Based on the following data snapshot, identify the top 3 most urgent and actionable items for the leadership today. 
Focus on risks, underperformance against targets, critical project issues, or urgent pipeline needs.
Be specific: mention project names, people, or exact metrics that need attention. Frame each item as a clear action.

Data:
{data_context_string}

Return your answer as a markdown numbered list. Example:
1. **Address Red Project 'X':** Key issue is Y, revenue at risk is $Z. Action: Schedule emergency meeting with PM.
2. **Boost Pipeline Coverage:** Currently at A.Bc ratio, target is 3.0x. Action: Focus sales team on deals in TIER.
3. **Support Underutilized Staff:** N team members below 70% utilization. Action: Review upcoming project needs with resource manager.
"""
    try:
        response = openai_client.chat.completions.create(
            model="gpt-4o", 
            messages=[
                {"role": "system", "content": "You are a highly experienced operations director for a professional services firm. Provide concise, actionable, and data-driven recommendations."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3, 
            max_tokens=400
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating AI action items: {str(e)}"

def get_upcoming_key_dates(data, days_ahead=7):
    """
    Scans project and pipeline data for key dates occurring within the specified number of days.
    """
    upcoming_events = []
    today = datetime.now().date()
    future_date_limit = today + timedelta(days=days_ahead)

    # Check Project Inventory for Project End Dates
    project_df = data.get('Project Inventory')
    if project_df is not None and not project_df.empty:
        if 'Project Name' in project_df.columns and 'Project End Date' in project_df.columns:
            project_df_copy = project_df.copy()
            project_df_copy['Project End Date'] = pd.to_datetime(project_df_copy['Project End Date'], errors='coerce').dt.date
            upcoming_ends = project_df_copy[
                (project_df_copy['Project End Date'] >= today) &
                (project_df_copy['Project End Date'] <= future_date_limit)
            ]
            for _, row in upcoming_ends.iterrows():
                upcoming_events.append(f"- Project '{row['Project Name']}' is scheduled to end on {row['Project End Date']:%Y-%m-%d}.")
        else:
            if 'Project Name' not in project_df.columns: print("Warning: 'Project Name' column missing in Project Inventory for upcoming dates.")
            if 'Project End Date' not in project_df.columns: print("Warning: 'Project End Date' column missing in Project Inventory for upcoming dates.")


    # Check Pipeline for Closed Won Dates and Next Touchpoint Dates
    pipeline_df = data.get('Pipeline')
    if pipeline_df is not None and not pipeline_df.empty:
        pipeline_df_copy = pipeline_df.copy()
        # Closed Won Date
        if 'Account' in pipeline_df_copy.columns and 'Closed Won Date' in pipeline_df_copy.columns:
            pipeline_df_copy['Closed Won Date'] = pd.to_datetime(pipeline_df_copy['Closed Won Date'], errors='coerce').dt.date
            upcoming_closes = pipeline_df_copy[
                (pipeline_df_copy['Closed Won Date'] >= today) &
                (pipeline_df_copy['Closed Won Date'] <= future_date_limit)
            ]
            for _, row in upcoming_closes.iterrows():
                upcoming_events.append(f"- Deal for '{row['Account']}' is scheduled to close (won) on {row['Closed Won Date']:%Y-%m-%d}.")
        else:
            if 'Account' not in pipeline_df_copy.columns: print("Warning: 'Account' column missing in Pipeline for upcoming dates.")
            if 'Closed Won Date' not in pipeline_df_copy.columns: print("Warning: 'Closed Won Date' column missing in Pipeline for upcoming dates.")

        # Next Touchpoint Date
        if 'Account' in pipeline_df_copy.columns and 'Next Touchpoint Date' in pipeline_df_copy.columns:
            pipeline_df_copy['Next Touchpoint Date'] = pd.to_datetime(pipeline_df_copy['Next Touchpoint Date'], errors='coerce').dt.date
            upcoming_touchpoints = pipeline_df_copy[
                (pipeline_df_copy['Next Touchpoint Date'] >= today) &
                (pipeline_df_copy['Next Touchpoint Date'] <= future_date_limit)
            ]
            for _, row in upcoming_touchpoints.iterrows():
                upcoming_events.append(f"- Next touchpoint for '{row['Account']}' is scheduled on {row['Next Touchpoint Date']:%Y-%m-%d}.")
        else:
            # 'Account' missing check already done above for Closed Won Date
            if 'Next Touchpoint Date' not in pipeline_df_copy.columns: print("Warning: 'Next Touchpoint Date' column missing in Pipeline for upcoming dates.")

    if not upcoming_events:
        return "No key dates identified in the next 7 days."
    return "\n".join(upcoming_events) # Using \n for markdown newlines in the prompt

def get_daily_digest_content(data, openai_client, data_context_string):
    if not openai_client:
        return "OpenAI client not configured. Cannot generate the Daily Digest."

    max_context_len = 100000  # Adjust as needed, keeping OpenAI token limits in mind
    if len(data_context_string) > max_context_len:
        data_context_string = data_context_string[:max_context_len] + "\n... (data truncated for brevity)"

    upcoming_key_dates_str = get_upcoming_key_dates(data) # Get upcoming dates

    prompt = f"""
You are an AI executive assistant for a healthcare delivery organization. Your task is to generate a "Daily Executive Digest".
This digest should be concise, data-driven, and highlight key information for leadership.
The output must be in well-formatted markdown.

Based on the following data snapshot:
{data_context_string}

Please structure the digest as follows:

**Overall Summary:**
Provide a brief (2-3 sentences) overview of the current business health.

**✅ On Track:**
Highlight 2-3 key areas or metrics that are performing well or meeting targets.
Be specific (e.g., "Revenue is X% of target", "Green Project Ratio at Y%").

**⚠️ Needs Attention:**
Identify 2-3 critical areas or metrics that are underperforming, at risk, or require immediate focus.
Be specific (e.g., "Pipeline Coverage is A.Bx, below target of 3.0x", "N Red Projects with $Y total revenue at risk").

**🗓️ Key Dates This Week:**
{upcoming_key_dates_str}

**🎯 Top 3 Action Items:**
List the three most urgent and actionable items for leadership today. These should be distinct from the "Needs Attention" section but can be derived from it.
Focus on risks, underperformance, critical project issues, or urgent pipeline needs.
Frame each item as a clear action, mentioning specific projects, people, or metrics if relevant.
Example:
1.  **Address Red Project 'Alpha':** Key issue is 'XYZ', revenue at risk is $ABC. Action: Convene an immediate review with the Project Manager.
2.  **Boost Pipeline Coverage:** Currently at A.Bc (target 3.0x). Action: Direct sales team to prioritize 'Tier 1' opportunities and accelerate 'Tier 2' deal closures.
3.  **Mitigate High Severity Risks:** N high-severity risks identified with a total impact of $Y. Action: Review mitigation plans for the top 2 risks with the risk owners.

**💡 Focus for Today:**
A brief concluding thought or strategic focus point for the day (1-2 sentences).

Ensure the language is professional, direct, and suitable for an executive audience.
The entire digest should be easily readable and scannable.
"""
    try:
        response = openai_client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are an AI executive assistant tasked with generating a concise and actionable daily digest for healthcare delivery leadership. Output in markdown."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=800  # Increased token limit for a more comprehensive digest
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating Daily Digest: {str(e)}"