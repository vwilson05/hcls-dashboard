## 📊 Google Sheet Tabs and Columns

Below is the schema definition for each tab in the google sheet.

### 1. Project Inventory

| Column                        | Description                                                                 |
|-------------------------------|-----------------------------------------------------------------------------|
| Project Name                  | The official name of the engagement.                                        |
| Client                        | The customer or business unit sponsoring the project.                       |
| Project Start Date            | YYYY-MM-DD format; when work kicked off.                                     |
| Project End Date              | YYYY-MM-DD format; expected completion date.                                  |
| Revenue                       | Total contracted revenue for the project, in USD.                           |
| Margin                        | Percentage margin (= (Revenue – Cost) / Revenue).                           |
| Status (R/Y/G)                | Traffic-light status: Red=behind schedule, Yellow=at-risk, Green=on-track.   |
| Key Issues                    | Short summary of top 1–3 risks or blockers.                                 |
| Next Steps                    | Immediate actions planned to keep the project on track.                     |
| Executive Support Required    | Yes/No flag for whether C-suite involvement is needed.                     |
| eNPS                          | Employee Net Promoter Score for the project team.                           |
| Next Opp First Discussion Date| When expansion was first discussed with the client.                         |
| Last Sponsor Checkin Date     | Date of most recent steering committee or sponsor meeting.                  |
| Sponsor Checkin Notes         | Key takeaways or action items from that meeting.                            |
| Timeline Health               | Qualitative health indicator for schedule (options are On Track, Minor Delay,s Major Delays, Off Track).   |
| Budget and Scope              | Qualitative health indicator for budget/scope (options are On Track,Some Risk, Over Budget/Scope Creep).                              |
| Client Relationship Strength  | Qualitative measure of trust and satisfaction (options are Strong Champion, Supportive Team, Neutral, Challenging, Deteriorating).                              |
| Feedback Recency              | How recently client feedback was captured (options are <2 weeks, <1 month, 1-2 months, >2 months, Never).             |
| Business Outcome Defined      | Yes/No flag for clear success criteria defined in the SOW (options are Clearly Defined, Partially Defined, Vague, None).                  |
| Team Resourcing               | Qualitative indicator of staffing adequacy (options are Yes, Some Gaps, Understaffed, Misaligned, No Core Team)                                  |
| Strategic Value to Client     | High/Medium/Low rating of overall business impact.                          |
| Issue Resolution Hygiene      | Rating of how well issues are being tracked and closed (options are Yes, Partially, No Process).                     |
| Expansion Discussion          | Indicator for whether expansion conversations are in progress (options are Already in Motion, Planned, Mentioned, No).            |
| Exec Engagement               | Qualitative rating of executive engagement (options are Yes (Executive Sponsor), Occasionally, Rarely, Not at All).                  |
| Project Health Score          | Composite numeric score summarizing overall health (score calulated by each of the scorecard columns and their "score" in the mapping table, using this formula for each scorecard value: =SUM(
  IFERROR(INDEX(FILTER(MappingTable!D:D, MappingTable!A:A="Project Inventory", MappingTable!C:C=O16), 1), 0),).                         |
| Health Band                   | Banding (“A”/“B”/“C”) based on Project Health Score thresholds.             |
| Delivery Relational Effort    | Effort rating (Low/Medium/High) for managing client relationships.          |
| Expansion Yield Tier          | High/Medium/Low tier for expected revenue uplift from expansion.            |
| Delivery Efficiency Score     | Numeric score weighting key pairs of Delivery Relational Effort and Expansion Yield Tier.             |
| Total Project Score           | =(Project Health Score*0.7)+(Delivery Efficiency Score*0.3)) 

---

### 2. Project Risks

| Column                  | Description                                                             |
|-------------------------|-------------------------------------------------------------------------|
| Project Name            | The name of the project to which the risk applies.                     |
| Risk Description        | A brief summary of the potential problem or threat.                    |
| Severity (High/Medium/Low) | Qualitative rating of how critical the risk is.                   |
| Impact ($)              | Estimated financial loss or cost if the risk materializes.             |
| Mitigation Plan         | Actions or controls planned to reduce the risk likelihood or impact.   |
| Owner                   | The person responsible for managing and tracking this risk.            |

---

### 3. Pipeline

| Column                      | Description                                                                 |
|-----------------------------|-----------------------------------------------------------------------------|
| Account                | Name of the customer or account in the sales pipeline.                      |
| Opportunity Created Date    | Date when the opportunity was first logged (YYYY-MM-DD).                     |
| Last Touchpoint Date        | A specific field to track the last touchpoint for this pursuit.             |
| Next Touchpoint Date        | A specific field to track the next touchpoint for this pursuit.             |
| Closed Won Date             | Date when the deal was successfully closed (YYYY-MM-DD).                    |
| Support Focus Type          | Primary type of support required (Tier 1/2/3).              |
| Perceived Annual AMO        | User’s estimated Annual Managed Offerings revenue potential.               |
| Open Pipeline_Active Work   | Current open opportunities and associated work in progress.                |
| Horizon                     | Timeframe category (e.g., Near-term, Mid-term, Long-term).                  |
| Pursuit Tier                | Tier classification based on strategic fit (e.g., Tier 1, 2, 3).            |
| Notes                       | Free-form notes on the opportunity.                                         |
| Help Needed                 | Specific assistance requested (e.g., Executive intro).                      |
| Actions                     | Next actions planned on this pipeline item.                                 |
| Internal Pursuit Team       | Which internal teams are engaged on this opportunity.                       |
| Key Client Contacts	        | The key client contacts for this deal.                                      |
| Win Themes	                | What's our core strategy to win this?                                       |
| Known Competitors           | Who are we up against?                                                      |
| Deal Registered             | Yes/No flag indicating if the deal is officially registered.                |
| MSA/ICA Status              | Status of Master Services Agreement / Intercompany Agreement.                |
| Roadmap Alignment           | How well this opportunity aligns to the clients product roadmap (options are Yes, Strongly Aligned, Early Draft, In Discussion, No).                    |
| Sponsor Type                | Type of sponsor (options are Executive Business Sponsor, Director-level Business Sponsor, IT Sponsor, Influencer Only, None).                     |
| Business Case_ROI           | Strenth of tbe Business Case/ROI definition (options are Quantified + Signed Off, Quantified Only, Draft in Progress, Discussed Conceptually, None).                              |
| HCLS Expertise needed       | Whether Healthcare/ Life Sciences expertise is required (options are Mission Critical, Strongly Beneficial, Helpful, Somewhat Useful, Not Needed).                    |
| Executor Pool Size          | Number of internal resources available to execute (options are 5+ Qualified, 3-4 Qualified, 2 Qualified, 1 Qualified, None).                          |
| Snowflake Investment Level  | Estimated level of Snowflake spend currently (options are Strategic Platform, Multi-Domain Strategic, Key Use Case, Trial Project, Not at All ).                                |
| IBM Growth in Account       | Historical IBM revenue growth trend in this account (options are Growing, Slightly Growing, Neutral, Slightly Shrinking, Shrinking).                        |
| Snowflake Growth in Account | Historical Snowflake spend growth trend (options are Growing, Slightly Growing, Neutral, Slightly Shrinking, Shrinking).                                    |
| Pipeline Score              | omposite numeric score summarizing overall health of the opportunity (score calulated by each of the scorecard columns and their "score" in the mapping table, using this formula for each scorecard value: =SUM(
  IFERROR(INDEX(FILTER(MappingTable!D:D, MappingTable!A:A="Pipeline", MappingTable!C:C=O16), 1), 0),).                         |                 |
| Score Band                  | Banding (e.g., A/B/C) based on Pipeline Score thresholds.                   |
| Pre-Sales Effort Level      | Estimated effort level for pre-sales activities (Low/Med/High).             |
| Revenue Potential Tier      | Tier classification for revenue potential (High/Medium/Low).                |
| Relational Efficiency Score | Numeric score weighting key pairs of Pre-Sales Effort Level and Revenue Potential Tier.               |
| Total Deal Score            | =(Pipeline Score*0.7)+(Relational Efficiency Score*0.3))                   |

---

### 4. Team Utilization

| Column                         | Description                                                                 |
|--------------------------------|-----------------------------------------------------------------------------|
| Employee Name                  | Full name of the team member.                                               |
| Role                           | Job title or function (e.g., Data Engineer, Consultant).                    |
| Project Assignments            | List of active projects the employee is assigned to.                        |
| Utilization (%)                | Percentage of their capacity currently billable.                            |
| Billable Rate ($/hr)           | Their hourly billing rate to the client.                                    |
| Strategic Opportunity Cost ($/week) | Estimated weekly cost of diverting them from strategic work.      |
| Latest Pulse Score             | Most recent engagement or satisfaction score (e.g., from an internal survey).|

---

### 5. Talent Gaps

| Column                | Description                                                                 |
|-----------------------|-----------------------------------------------------------------------------|
| Skill/Role Needed     | The specific skill or role that is lacking.                                 |
| Gap Impact (High/Medium/Low) | How severely this gap affects delivery or strategy.               |
| Urgency               | Time sensitivity for filling this gap (e.g., Immediate, Short-term).        |
| Target Hire Date      | Desired date to have this position filled. (YYYY-MM-DD)                     |
| Hiring Owner          | Person or team responsible for recruiting/hiring.                           |

---

### 6. Operational Gaps

| Column                     | Description                                                                 |
|----------------------------|-----------------------------------------------------------------------------|
| Operational Issue          | Name or summary of the process or system gap.                                |
| Severity (High/Medium/Low) | Qualitative rating of how critical the gap is.                            |
| Frequency                  | How often the issue occurs (e.g., Daily, Weekly, As-needed).                 |
| Recommended Process/Solution | Proposed fix or process change to resolve the gap.                     |
| Owner                      | The person responsible for implementing and monitoring the solution.        |

---

### 7. Executive Activity

| Column                   | Description                                                                 |
|--------------------------|-----------------------------------------------------------------------------|
| Activity                 | Description of the executive task or meeting.                               |
| Type (Delivery/Sales/Ops)| Category of the activity.                                                   |
| Time Spent Weekly (hrs)  | Number of hours per week dedicated to this activity.                        |
| Strategic Cost ($)       | Implied cost of executive time at their hourly rate.                        |
| Delegate to?             | If delegable, the person or team to whom it can be delegated.               |

---

### 8. Scenario Model Inputs

| Column     | Description                         |
|------------|-------------------------------------|
| Assumption | Description of the input assumption.|
| Value      | Numeric or categorical value.       |

---

### 9. Do Nothing Scenario

| Column                | Description                                                               |
|-----------------------|---------------------------------------------------------------------------|
| Category              | Business area or cost center impacted under “do nothing.”                 |
| Calculation Details   | Formula or logic used to compute the impact.                              |
| Annualized Impact ($) | Total estimated yearly cost or lost revenue if no action is taken.        |

---

### 10. Proposed Scenario

| Column                | Description                                                               |
|-----------------------|---------------------------------------------------------------------------|
| Category              | Business area or cost center impacted under the proposed change.          |
| Calculation Details   | Formula or logic used to compute the impact.                              |
| Annualized Impact ($) | Total estimated yearly benefit or cost under the proposed scenario.       |

---

### 11. Scenario Comparison

| Column                     | Description                                                               |
|----------------------------|---------------------------------------------------------------------------|
| Scenario                   | Name of the scenario (e.g., Do Nothing, Proposed).                        |
| Total Annualized Impact ($)| Annualized impact for this scenario.                                      |
| Net Incremental Value ($)  | Difference in impact versus the baseline (Do Nothing).                    |

---

### 12. Sensitivity & Assumption Analysis *(Optional)*

| Column                        | Description                                                                  |
|-------------------------------|------------------------------------------------------------------------------|
| Assumption/Scenario Input     | The input variable or assumption being tested.                               |
| Base Case                     | Default value used in baseline calculation.                                  |
| Sensitivity Range             | Range of values over which the assumption is varied (e.g., ±10%).            |
| Financial Impact (Low End)    | Impact on the metric at the low end of the sensitivity range.                |
| Financial Impact (High End)   | Impact on the metric at the high end of the sensitivity range.               |

---

### 13. Project Observations

| Column       | Description                                                                          |
|--------------|--------------------------------------------------------------------------------------|
| Date         | Date the observation was recorded (YYYY-MM-DD).                                      |
| Project      | Project to which the observation applies.                                            |
| Observation  | Free-form note on anything noteworthy (issues, lessons learned, highlights).         |

---

### 14. MappingTable

| Column   | Description                                                             |
|----------|-------------------------------------------------------------------------|
| Tab      | Name of the sheet/tab where the question applies.                      |
| Question | The specific question or field being mapped (e.g., “Delivery Relationship Efficiency”). |
| Value    | The raw value(s) to look up (e.g., “Low|High”).                        |
| Score    | Numeric score that the mapping yields.                                 |
| Key      | Composite key used for lookup logic (e.g., “Delivery|Low|High”).       |