"""
Strategic targets and constants for FY26 (July 1 2025 - June 30 2026).
These values are used to calculate progress against targets and visualize metrics.
"""

# Revenue Targets
REVENUE_TARGET = 15_000_000  # $15M
REVENUE_STRETCH_GOAL = 20_000_000  # $20M
FY27_TARGET = 20_000_000  # $20M
FY27_BOOKED_IN_FY26_TARGET = 10_000_000  # $10M (50% of FY27 target)

# Project Health Targets
GREEN_PROJECT_TARGET = 0.90  # 90% of projects should be green
EMPLOYEE_PULSE_TARGET = 8.0  # Target pulse score
CUSTOMER_NPS_TARGET = 50.00

# Pipeline Coverage Target
PIPELINE_COVERAGE_TARGET = 3.0  # Pipeline should be 3x revenue target

# Sponsor Check-in Target
SPONSOR_CHECKIN_WINDOW_DAYS = 30  # Check-ins should occur within 30 days

# Next Deal Discussion Target
NEXT_DEAL_DISCUSSION_THRESHOLD_DAYS = 30  # Flag if no discussion within 30 days of project end 

# Score Bands for Projects and Pipeline
PROJECT_SCORE_BANDS = {
    'Excellent': (85, 100),
    'Strong': (70, 84.99),
    'Weak': (0, 69.99)
}
PIPELINE_SCORE_BANDS = {
    'Strong': (70, 100),
    'Medium': (40, 69.99),
    'Weak': (0, 39.99)
}

# Score Targets
PROJECT_HEALTH_SCORE_TARGET = 85  # Target average project health score
PIPELINE_SCORE_TARGET = 70        # Target average pipeline score 