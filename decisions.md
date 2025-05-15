# 🧠 decisions.md — AI-Powered Delivery Dashboard

A log of all architectural, design, and functionality decisions made during the development of the internal delivery operations app.

## 📅 2025-05-14 — Context for OpenAI
**Decision**: Send all rows in all tabs every time when querying OpenAI.

**Rationale**: Small data and OpenAI needs the full context to reply correctly.

## 📅 2025-05-14 — Pattern Matching for AI Assistant Queries
**Decision**: Use pattern matching (regex) in the AI Assistant to answer certain quantitative or well-defined queries directly in Python, bypassing OpenAI when possible.

**Rationale**: Ensures accuracy, speed, and cost-effectiveness for common questions (e.g., highest revenue project, total pipeline value) by leveraging structured data and avoiding unnecessary LLM calls.

---

## 📅 2025-05-13 — App Framework Chosen
**Decision**: Use Streamlit for the front-end UI.

**Rationale**: Lightweight, fast to prototype, requires no front-end expertise, integrates easily with Python backend and Google Sheets.

---

## 📅 2025-05-13 — Backend Data Source
**Decision**: Use Google Sheets as the live backend database.

**Rationale**: All operational data is already in a structured Google Sheet. `gspread` makes it easy to read/write data without managing a database.

---

## 📅 2025-05-13 — Natural Language Integration
**Decision**: Use OpenAI GPT-4o to allow natural-language querying of project data.

**Rationale**: Users should be able to ask questions like \"Which projects are most at risk?\" or \"What's the opportunity cost of executive delivery time?\"

---

## 📅 2025-05-13 — Data Tabs and Structure
**Decision**: Use 13 structured tabs in Google Sheets to manage Projects, Pipeline, Risks, Utilization, Scenarios, etc.

**Rationale**: Aligns with the strategic objectives (revenue, project health, staffing) and supports future reporting needs.

---

## 📅 2025-05-13 — Initial Deployment Target
**Decision**: Deploy on Streamlit Cloud.

**Rationale**: No infrastructure setup needed. Enables rapid internal access, testing, and iteration without DevOps work.

---

## 📅 2025-05-13 — Prompt Format for OpenAI
**Decision**: Format prompt as:  
\"Based on this data: {data_csv_or_summary}, answer the question: {user_question}\"

**Rationale**: Keeps prompt concise while still passing enough context for accurate responses. Truncate or summarize if token count is too high.