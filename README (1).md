# GovPulse: Automated Action Tracking & Weekly Governance Reporting

Turns a simple Excel action tracker into a ready-to-present weekly management update and
follow-up emails, replacing the manual "chase owners, update status, build slides" cycle.

## What it does
1. **Tracker (Excel):** dropdowns for Status/Priority/Function, plus formulas for *Days to Due*
   and a **RAG status** (Red / Amber / Green / Done) with conditional formatting.
2. **Report generator (Python):** reads the tracker, validates data quality (bad status, missing
   owner/due date), calculates KPIs, and builds a 5-slide PowerPoint:
   headline + KPI cards, status and function charts, overdue/blocked items, next-7-day follow-ups.
3. **Follow-up drafts:** one email draft per owner listing only their overdue, blocked or
   soon-due actions.

RAG rules (identical in Excel and Python):
- Done: Completed
- Red: open and past due date
- Amber: Blocked, or due within 7 days with under 50% complete
- Green: everything else

## Run it
```bash
pip install -r requirements.txt
python create_tracker.py                        # creates data/action_tracker.xlsx (sample data)
python report.py --input data/action_tracker.xlsx
python report.py --as-of 2026-10-08             # re-run for any report date
```
Outputs go to `reports/`. Replace the sample rows with real actions (columns A-L only).

## Project structure
- `create_tracker.py`: builds the Excel template
- `report.py`: metrics, RAG logic, deck and email generation
- `data/`, `reports/`: input and output


