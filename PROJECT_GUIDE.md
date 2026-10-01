# GovPulse: Project Guide

A complete walkthrough of what this project does, how every part works, why it was designed this way, and how to talk about it in an interview.

---

## 1. The project in one paragraph

In a COO Office, dozens of actions (from steering committees, audits, regulatory work and transformation programmes) are tracked in Excel. Every week someone has to chase owners, work out what is overdue, and build a status deck for management. **GovPulse automates that cycle.** You keep one Excel tracker; a Python script reads it, calculates status and RAG (Red/Amber/Green) metrics, checks the data for errors, and produces a ready-to-present PowerPoint plus follow-up email drafts for each action owner.

**Why it matters for the role:** it directly mirrors the job description: *maintain action trackers, coordinate follow-ups, prepare management presentations and governance reports, build structured reports, and support workflow automation.*

---

## 2. End-to-end flow

```
  action_tracker.xlsx            report.py                         reports/
  (people type here)     ─────►  1. load + validate data   ─────►  Weekly_Action_Report_<date>.pptx
                                 2. calculate Days to Due           Follow_Up_Emails_<date>.txt
  Sheet "Tracker"                3. assign RAG status
  Sheet "Instructions"           4. compute KPIs
                                 5. build slides + email drafts
```

| Step | What happens | Where in code |
|---|---|---|
| 1 | Tracker is created with sample data, dropdowns, formulas | `create_tracker.py` |
| 2 | Excel is read, cleaned and validated | `load_tracker()` |
| 3 | Each action gets a RAG status | `_rag()` |
| 4 | Totals, overdue, blocked, due-soon lists are computed | `compute_metrics()` |
| 5 | The 5-slide deck is generated | `build_deck()` |
| 6 | One email draft per owner is written | `write_followups()` |

---

## 3. The data model (tracker columns)

| Column | Typed by a person? | Meaning |
|---|---|---|
| ID | Yes | Unique action ID (A-001...) |
| Initiative | Yes | The programme the action belongs to |
| Action Item | Yes | What needs to be done |
| Owner | Yes | Person accountable |
| Function | Yes | Business, Operations, Finance, Risk, Compliance, Technology, Infrastructure |
| Priority | Yes | High / Medium / Low |
| Status | Yes | Not Started / In Progress / Blocked / Completed |
| Start Date, Due Date | Yes | Planned dates |
| % Complete | Yes | Self-reported progress |
| Last Update, Remarks | Yes | Context for reviewers |
| Days to Due | **Formula** | Due Date minus today |
| RAG | **Formula** | Calculated status colour |

Blue text in Excel = typed input. Black text = formula. This is a standard modelling convention that makes it obvious what to edit.

---

## 4. RAG logic (the core business rule)

| RAG | Rule |
|---|---|
| **Done** | Status is Completed |
| **Red** | Not completed **and** due date is before today |
| **Amber** | Status is Blocked, **or** due within 7 days with under 50% complete |
| **Green** | Everything else |

**Worked example.** Report date is 1 Oct 2026. Action A-012 ("Close access-review findings") is In Progress and was due 17 Sep. Days to Due = -14, so it is not Done, and it is past due, so it is **Red**, shown as "Overdue 14d" on slide 4.

The same rules exist twice: as Excel formulas (so users see RAG live in the sheet) and in Python `_rag()` (so the report is consistent). The Instructions sheet documents the rules.

---

## 5. Code walkthrough

### `create_tracker.py`
Builds the Excel file with `openpyxl`. Key points:
- Sample dates are generated **relative to today**, so the demo always contains overdue, blocked and upcoming items.
- Dropdowns (data validation) on Function, Priority, Status prevent typos like "inprogress".
- Conditional formatting colours the RAG column.
- Freeze panes, filters and an Instructions sheet make it usable by someone who has never seen it.

### `report.py`

**`load_tracker()`**: reads the "Tracker" sheet and:
- checks required columns exist (fails with a clear message if not)
- standardises text (`" in progress "` becomes `"In Progress"`)
- converts dates; fixes percentages typed as `60` instead of `60%`
- **prints data-quality warnings** (unknown status, missing owner, missing/invalid due date) rather than crashing or silently hiding problems
- calculates `Days to Due` and `RAG`

**`compute_metrics()`**: produces the numbers behind the report:
- *Completion rate* = completed / total
- *Overdue* = open and Days to Due < 0
- *Blocked* = Status is Blocked
- *Due soon* = open and due in the next 0-7 days
- *Needs attention* = overdue or blocked, sorted by priority (High first), then most overdue first
- *Average open progress* = mean % Complete of open actions

**`build_deck()`**: builds 5 slides with `python-pptx`:
1. Title
2. Headline (auto-written from the data, e.g. "24% of actions are complete; 5 are overdue"), KPI cards, key messages
3. Native, editable charts: actions by status, and actions by function split by RAG
4. Table of overdue and blocked actions needing escalation
5. Table of actions due in the next 7 days

Slide titles are **conclusions, not labels** ("5 overdue and 2 blocked actions need attention" rather than "Overdue Items"). This is how management decks are written.

**`write_followups()`**: groups flagged actions by owner and writes a polite, specific email draft per person.

**`main()`**: command-line interface (`--input`, `--out`, `--as-of`) so the report can be re-run for any date and later scheduled.

---

## 6. Design decisions and trade-offs

| Decision | Why | Trade-off |
|---|---|---|
| Excel as the input | Business teams already live in Excel; zero training needed | Free-text entry invites errors, hence dropdowns and validation |
| Python generates the deck | Repeatable, consistent, no manual copy-paste | Layout changes need code edits |
| Native PowerPoint charts | Users can edit and restyle them | Slightly less flexible than images |
| Warn on bad data instead of crashing | Real trackers are messy; the report should still run | Warnings must be read by someone |
| Same RAG rules in Excel and Python | Sheet and deck always agree | Logic is duplicated; a change must be made in both places |
| Draft emails, not sent emails | Human stays in control of communication | One manual step remains |

---

## 7. Limitations (know these before an interviewer asks)

- The 7-day window and 50% threshold are hard-coded business assumptions. In practice they should be agreed with stakeholders and moved to a config file.
- % Complete is self-reported, so the report is only as honest as the inputs.
- No history: it shows a snapshot, not a trend over weeks.
- Sample data is fictional; the tool has been validated on realistic structure, not on a live bank tracker.
- Emails are drafts only.

---

## 8. Interview preparation

### 30-second pitch
"I built GovPulse, a tool that turns an Excel action tracker into a weekly governance report. It validates the data, assigns Red/Amber/Green status, builds a management deck with charts and an escalation list, and drafts follow-up emails to action owners. I built it because a COO Office spends a lot of time on exactly this cycle, and it's repeatable work that automation handles well."

### Likely questions and answer outlines

1. **Why Excel and not a database or web app?** Users already work in Excel and need no training; adoption matters more than architecture. A database would be a sensible next step at larger scale.
2. **How do you make sure Excel and Python agree on RAG?** Same documented rules in both places, plus I compared outputs on the sample data. Longer term: one source of truth (the script writes RAG back to the sheet).
3. **What if the data is messy?** Validation flags unknown statuses, missing owners and bad dates, normalises text and percentages, and the run continues with clear warnings.
4. **Why those RAG thresholds?** They are reasonable defaults. In a real team I would agree thresholds with stakeholders, document them, and make them configurable.
5. **How would you add AI?** An LLM-written executive summary, with an automated check that every number it quotes matches the data. This links to my hallucination-validation project.
6. **How would you scale it to hundreds of actions or multiple teams?** Add filters per team, move data to a database, add a dashboard view, and schedule runs.
7. **What was the hardest part?** Pick one honestly, for example making slides readable (long action names, table fit) or keeping logic consistent across Excel and Python.
8. **How would you measure success?** Time saved per week, number of actions closed on time, fewer manual chasers, report accuracy.

### How to describe your contribution honestly
Say what you designed (RAG rules, metrics, slide structure, validation) and what you built. If you used AI tools to help write code, say so openly and be ready to explain every function. The job description explicitly values responsible use of AI tools.

---

## 9. Skills this project demonstrates (mapped to the JD)

| JD expectation | Evidence in this project |
|---|---|
| Track initiatives and maintain action trackers | Tracker design, RAG logic, validation |
| Prepare management presentations and governance reports | Auto-generated 5-slide deck with conclusion-style titles |
| Prepare dashboards and analyse performance | KPIs, charts by status and function |
| Coordinate follow-ups with stakeholders | Per-owner follow-up email drafts |
| Process simplification and automation | End-to-end automation of a weekly manual cycle |
| Attention to detail | Data-quality checks, consistent rules, tested output |
| Responsible use of data and AI | Validation-first approach, human-reviewed drafts |

---

## 10. Glossary

- **COO Office:** team supporting the Chief Operating Officer on strategy execution, performance, governance and transformation.
- **RAG:** Red/Amber/Green status used for quick risk-based reporting.
- **Governance:** the oversight processes (reviews, audits, committees) that keep a business controlled and compliant.
- **KPI:** key performance indicator, a measurable value showing progress.
- **MIS:** management information system / report used for decisions.
- **Data validation (Excel):** dropdown rules that restrict what can be typed in a cell.
