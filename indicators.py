import pandas as pd
from datetime import datetime, timedelta
from strategic_targets import (
    REVENUE_TARGET, REVENUE_STRETCH_GOAL, FY27_TARGET, FY27_BOOKED_IN_FY26_TARGET,
    GREEN_PROJECT_TARGET, EMPLOYEE_PULSE_TARGET, PIPELINE_COVERAGE_TARGET,
    SPONSOR_CHECKIN_WINDOW_DAYS, NEXT_DEAL_DISCUSSION_THRESHOLD_DAYS, CUSTOMER_NPS_TARGET,
    PROJECT_SCORE_BANDS, PIPELINE_SCORE_BANDS, PROJECT_HEALTH_SCORE_TARGET, PIPELINE_SCORE_TARGET
)

# --- Score Banding Helpers ---
def band_score(score, bands):
    for band, (low, high) in bands.items():
        if low <= score <= high:
            return band
    return None

def score_band_distribution(scores, bands):
    band_counts = {band: 0 for band in bands}
    for s in scores:
        band = band_score(s, bands)
        if band:
            band_counts[band] += 1
    total = sum(band_counts.values())
    band_pct = {band: (count / total * 100 if total > 0 else 0) for band, count in band_counts.items()}
    return band_counts, band_pct

def get_lagging_indicators(data):
    """Extract lagging indicators from the data dict."""
    indicators = {}
    
    # Revenue
    project_df = data.get('Project Inventory')
    if project_df is not None and not project_df.empty:
        project_df = project_df.copy()
        project_df['Revenue'] = project_df['Revenue'].astype(str).str.replace('$', '').str.replace(',', '')
        project_df['Revenue'] = pd.to_numeric(project_df['Revenue'], errors='coerce').fillna(0)
        total_revenue = project_df['Revenue'].sum()
        indicators['Revenue'] = total_revenue
        indicators['Revenue_vs_Target'] = (total_revenue / REVENUE_TARGET) * 100
        indicators['Revenue_vs_Stretch'] = (total_revenue / REVENUE_STRETCH_GOAL) * 100
    else:
        indicators['Revenue'] = None
        indicators['Revenue_vs_Target'] = None
        indicators['Revenue_vs_Stretch'] = None

    # Customer NPS (eNPS)
    if project_df is not None and not project_df.empty and 'eNPS' in project_df.columns:
        try:
            enps = pd.to_numeric(project_df['eNPS'], errors='coerce').dropna()
            avg_enps = enps.mean() if not enps.empty else None
            indicators['Customer NPS'] = avg_enps
            indicators['Customer NPS_vs_Target'] = (avg_enps / CUSTOMER_NPS_TARGET) * 100 if avg_enps is not None else None
        except Exception:
            indicators['Customer NPS'] = None
            indicators['Customer NPS_vs_Target'] = None
    else:
        indicators['Customer NPS'] = None
        indicators['Customer NPS_vs_Target'] = None

    # Employee Satisfaction (Pulse Score)
    util_df = data.get('Team Utilization')
    if util_df is not None and not util_df.empty and 'Latest Pulse Score' in util_df.columns:
        try:
            pulse = pd.to_numeric(util_df['Latest Pulse Score'], errors='coerce').dropna()
            avg_pulse = pulse.mean() if not pulse.empty else None
            indicators['Employee Satisfaction'] = avg_pulse
            indicators['Employee Satisfaction_vs_Target'] = (avg_pulse / EMPLOYEE_PULSE_TARGET) * 100 if avg_pulse is not None else None
        except Exception:
            indicators['Employee Satisfaction'] = None
            indicators['Employee Satisfaction_vs_Target'] = None
    else:
        indicators['Employee Satisfaction'] = None
        indicators['Employee Satisfaction_vs_Target'] = None

    # Project Health Score Metrics
    if project_df is not None and not project_df.empty and 'Project Health Score' in project_df.columns:
        try:
            scores = pd.to_numeric(project_df['Project Health Score'], errors='coerce').dropna()
            indicators['Avg Project Health Score'] = scores.mean() if not scores.empty else None
            indicators['Median Project Health Score'] = scores.median() if not scores.empty else None
            band_counts, band_pct = score_band_distribution(scores, PROJECT_SCORE_BANDS)
            indicators['Project Health Score Bands'] = band_counts
            indicators['Project Health Score Bands %'] = band_pct
            # Top/Bottom
            if not scores.empty:
                top_idx = scores.idxmax()
                bot_idx = scores.idxmin()
                indicators['Top Project by Health Score'] = project_df.loc[top_idx, 'Project Name']
                indicators['Bottom Project by Health Score'] = project_df.loc[bot_idx, 'Project Name']
            # Total Project Score
            if 'Delivery Efficiency Score' in project_df.columns:
                efficiency_scores = pd.to_numeric(project_df['Delivery Efficiency Score'], errors='coerce').dropna()
                total_project_scores = scores * 0.7 + efficiency_scores * 0.3
                indicators['Avg Total Project Score'] = total_project_scores.mean() if not total_project_scores.empty else None
                indicators['Median Total Project Score'] = total_project_scores.median() if not total_project_scores.empty else None
                band_counts, band_pct = score_band_distribution(total_project_scores, PROJECT_SCORE_BANDS)
                indicators['Total Project Score Bands'] = band_counts
                indicators['Total Project Score Bands %'] = band_pct
                # Top/Bottom
                if not total_project_scores.empty:
                    top_idx = total_project_scores.idxmax()
                    bot_idx = total_project_scores.idxmin()
                    indicators['Top Project by Total Score'] = project_df.loc[top_idx, 'Project Name']
                    indicators['Bottom Project by Total Score'] = project_df.loc[bot_idx, 'Project Name']
        except Exception:
            indicators['Avg Project Health Score'] = None
            indicators['Median Project Health Score'] = None
            indicators['Project Health Score Bands'] = None
            indicators['Project Health Score Bands %'] = None
            indicators['Top Project by Health Score'] = None
            indicators['Bottom Project by Health Score'] = None
            indicators['Avg Total Project Score'] = None
            indicators['Median Total Project Score'] = None
            indicators['Total Project Score Bands'] = None
            indicators['Total Project Score Bands %'] = None
            indicators['Top Project by Total Score'] = None
            indicators['Bottom Project by Total Score'] = None
    else:
        indicators['Avg Project Health Score'] = None
        indicators['Median Project Health Score'] = None
        indicators['Project Health Score Bands'] = None
        indicators['Project Health Score Bands %'] = None
        indicators['Top Project by Health Score'] = None
        indicators['Bottom Project by Health Score'] = None
        indicators['Avg Total Project Score'] = None
        indicators['Median Total Project Score'] = None
        indicators['Total Project Score Bands'] = None
        indicators['Total Project Score Bands %'] = None
        indicators['Top Project by Total Score'] = None
        indicators['Bottom Project by Total Score'] = None

    return indicators

def get_leading_indicators(data):
    """Extract leading indicators from the data dict."""
    indicators = {}
    
    # Pipeline Coverage
    pipeline_df = data.get('Pipeline')
    if pipeline_df is not None and not pipeline_df.empty:
        pipeline_df = pipeline_df.copy()
        pipeline_df['Open Pipeline_Active Work'] = pipeline_df['Open Pipeline_Active Work'].astype(str).str.replace('$', '').str.replace(',', '')
        pipeline_df['Open Pipeline_Active Work'] = pd.to_numeric(pipeline_df['Open Pipeline_Active Work'], errors='coerce').fillna(0)
        total_pipeline = pipeline_df['Open Pipeline_Active Work'].sum()
        pipeline_coverage = total_pipeline / REVENUE_TARGET
        indicators['Pipeline Coverage'] = total_pipeline
        indicators['Pipeline Coverage Ratio'] = pipeline_coverage
        indicators['Pipeline Coverage_vs_Target'] = (pipeline_coverage / PIPELINE_COVERAGE_TARGET) * 100
    else:
        indicators['Pipeline Coverage'] = None
        indicators['Pipeline Coverage Ratio'] = None
        indicators['Pipeline Coverage_vs_Target'] = None

    # Deal Cycle Time
    if pipeline_df is not None and not pipeline_df.empty:
        try:
            pipeline_df['Opportunity Created Date'] = pd.to_datetime(pipeline_df['Opportunity Created Date'], errors='coerce')
            pipeline_df['Closed Won Date'] = pd.to_datetime(pipeline_df['Closed Won Date'], errors='coerce')
            won_deals = pipeline_df.dropna(subset=['Closed Won Date'])
            won_deals['Cycle Time'] = (won_deals['Closed Won Date'] - won_deals['Opportunity Created Date']).dt.days
            indicators['Avg Deal Cycle Time'] = won_deals['Cycle Time'].mean()
            indicators['Median Deal Cycle Time'] = won_deals['Cycle Time'].median()
            
            # Group by Pursuit Tier if available
            if 'Pursuit Tier' in won_deals.columns:
                tier_cycle_times = won_deals.groupby('Pursuit Tier')['Cycle Time'].agg(['mean', 'median']).to_dict()
                indicators['Deal Cycle Time by Tier'] = tier_cycle_times
        except Exception:
            indicators['Avg Deal Cycle Time'] = None
            indicators['Median Deal Cycle Time'] = None
            indicators['Deal Cycle Time by Tier'] = None
    else:
        indicators['Avg Deal Cycle Time'] = None
        indicators['Median Deal Cycle Time'] = None
        indicators['Deal Cycle Time by Tier'] = None

    # Time Between Project End and Next Deal Discussion
    project_df = data.get('Project Inventory')
    if project_df is not None and not project_df.empty:
        try:
            project_df['Project End Date'] = pd.to_datetime(project_df['Project End Date'], errors='coerce')
            project_df['Next Opp First Discussion Date'] = pd.to_datetime(project_df['Next Opp First Discussion Date'], errors='coerce')
            completed_projects = project_df.dropna(subset=['Project End Date'])
            completed_projects['Next Deal Gap'] = (completed_projects['Next Opp First Discussion Date'] - completed_projects['Project End Date']).dt.days
            indicators['Avg Next Deal Gap'] = completed_projects['Next Deal Gap'].mean()
            
            # Flag projects overdue for next deal discussion
            overdue_projects = completed_projects[
                (completed_projects['Next Deal Gap'] > NEXT_DEAL_DISCUSSION_THRESHOLD_DAYS) |
                (completed_projects['Next Opp First Discussion Date'].isna())
            ]
            indicators['Overdue Next Deal Projects'] = overdue_projects[['Project Name', 'Project End Date', 'Next Opp First Discussion Date', 'Next Deal Gap']].to_dict('records')
        except Exception:
            indicators['Avg Next Deal Gap'] = None
            indicators['Overdue Next Deal Projects'] = None
    else:
        indicators['Avg Next Deal Gap'] = None
        indicators['Overdue Next Deal Projects'] = None

    # Meaningful Sponsor Check-ins
    if project_df is not None and not project_df.empty:
        try:
            project_df['Last Sponsor Checkin Date'] = pd.to_datetime(project_df['Last Sponsor Checkin Date'], errors='coerce')
            current_date = pd.Timestamp.now()
            recent_checkins = project_df[
                (project_df['Last Sponsor Checkin Date'] >= current_date - pd.Timedelta(days=SPONSOR_CHECKIN_WINDOW_DAYS)) &
                (project_df['Sponsor Checkin Notes'].notna() & (project_df['Sponsor Checkin Notes'].str.strip() != ''))
            ]
            total_projects = len(project_df)
            indicators['Recent Meaningful Check-ins'] = len(recent_checkins)
            indicators['Recent Meaningful Check-ins %'] = (len(recent_checkins) / total_projects * 100) if total_projects > 0 else 0
            
            # Flag projects overdue for check-in
            overdue_checkins = project_df[
                (project_df['Last Sponsor Checkin Date'] < current_date - pd.Timedelta(days=SPONSOR_CHECKIN_WINDOW_DAYS)) |
                (project_df['Last Sponsor Checkin Date'].isna())
            ]
            indicators['Overdue Check-in Projects'] = overdue_checkins[['Project Name', 'Last Sponsor Checkin Date', 'Sponsor Checkin Notes']].to_dict('records')
        except Exception:
            indicators['Recent Meaningful Check-ins'] = None
            indicators['Recent Meaningful Check-ins %'] = None
            indicators['Overdue Check-in Projects'] = None
    else:
        indicators['Recent Meaningful Check-ins'] = None
        indicators['Recent Meaningful Check-ins %'] = None
        indicators['Overdue Check-in Projects'] = None

    # Green Project Ratio
    if project_df is not None and not project_df.empty:
        try:
            green_projects = project_df[project_df['Status (R/Y/G)'].str.strip().str.upper() == 'G']
            total_projects = len(project_df)
            green_ratio = len(green_projects) / total_projects if total_projects > 0 else 0
            indicators['Green Project Ratio'] = green_ratio
            indicators['Green Project Ratio_vs_Target'] = (green_ratio / GREEN_PROJECT_TARGET) * 100
            
            # Flag non-green projects
            non_green_projects = project_df[project_df['Status (R/Y/G)'].str.strip().str.upper() != 'G']
            indicators['Non-Green Projects'] = non_green_projects[['Project Name', 'Status (R/Y/G)', 'Key Issues']].to_dict('records')
        except Exception:
            indicators['Green Project Ratio'] = None
            indicators['Green Project Ratio_vs_Target'] = None
            indicators['Non-Green Projects'] = None
    else:
        indicators['Green Project Ratio'] = None
        indicators['Green Project Ratio_vs_Target'] = None
        indicators['Non-Green Projects'] = None

    # Pipeline Score Metrics
    if pipeline_df is not None and not pipeline_df.empty and 'Pipeline Score' in pipeline_df.columns:
        try:
            scores = pd.to_numeric(pipeline_df['Pipeline Score'], errors='coerce').dropna()
            indicators['Avg Pipeline Score'] = scores.mean() if not scores.empty else None
            indicators['Median Pipeline Score'] = scores.median() if not scores.empty else None
            band_counts, band_pct = score_band_distribution(scores, PIPELINE_SCORE_BANDS)
            indicators['Pipeline Score Bands'] = band_counts
            indicators['Pipeline Score Bands %'] = band_pct
            # Top/Bottom
            if not scores.empty:
                top_idx = scores.idxmax()
                bot_idx = scores.idxmin()
                indicators['Top Pipeline by Score'] = pipeline_df.loc[top_idx, 'Account']
                indicators['Bottom Pipeline by Score'] = pipeline_df.loc[bot_idx, 'Account']
            # Total Deal Score
            if 'Relational Efficiency Score' in pipeline_df.columns:
                relational_scores = pd.to_numeric(pipeline_df['Relational Efficiency Score'], errors='coerce').dropna()
                total_deal_scores = scores * 0.7 + relational_scores * 0.3
                indicators['Avg Total Deal Score'] = total_deal_scores.mean() if not total_deal_scores.empty else None
                indicators['Median Total Deal Score'] = total_deal_scores.median() if not total_deal_scores.empty else None
                band_counts, band_pct = score_band_distribution(total_deal_scores, PIPELINE_SCORE_BANDS)
                indicators['Total Deal Score Bands'] = band_counts
                indicators['Total Deal Score Bands %'] = band_pct
                # Top/Bottom
                if not total_deal_scores.empty:
                    top_idx = total_deal_scores.idxmax()
                    bot_idx = total_deal_scores.idxmin()
                    indicators['Top Pipeline by Total Score'] = pipeline_df.loc[top_idx, 'Account']
                    indicators['Bottom Pipeline by Total Score'] = pipeline_df.loc[bot_idx, 'Account']
        except Exception:
            indicators['Avg Pipeline Score'] = None
            indicators['Median Pipeline Score'] = None
            indicators['Pipeline Score Bands'] = None
            indicators['Pipeline Score Bands %'] = None
            indicators['Top Pipeline by Score'] = None
            indicators['Bottom Pipeline by Score'] = None
            indicators['Avg Total Deal Score'] = None
            indicators['Median Total Deal Score'] = None
            indicators['Total Deal Score Bands'] = None
            indicators['Total Deal Score Bands %'] = None
            indicators['Top Pipeline by Total Score'] = None
            indicators['Bottom Pipeline by Total Score'] = None
    else:
        indicators['Avg Pipeline Score'] = None
        indicators['Median Pipeline Score'] = None
        indicators['Pipeline Score Bands'] = None
        indicators['Pipeline Score Bands %'] = None
        indicators['Top Pipeline by Score'] = None
        indicators['Bottom Pipeline by Score'] = None
        indicators['Avg Total Deal Score'] = None
        indicators['Median Total Deal Score'] = None
        indicators['Total Deal Score Bands'] = None
        indicators['Total Deal Score Bands %'] = None
        indicators['Top Pipeline by Total Score'] = None
        indicators['Bottom Pipeline by Total Score'] = None

    return indicators

def get_top3_action_items(data, openai_client, data_context):
    """Use OpenAI to generate top 3 actionable items for today."""
    prompt = f"""
You are a business operations assistant. Based on the following data, identify the top 3 most urgent or actionable items for today. Be specific, actionable, and reference the relevant person, project, or metric.

Data:
{data_context}

Return your answer as a numbered list.
"""
    try:
        response = openai_client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a business operations assistant. Provide actionable, specific recommendations."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.5,
            max_tokens=300
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating action items: {str(e)}" 