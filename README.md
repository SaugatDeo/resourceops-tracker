ResourceOps Tracker

A personal project that simulates the daily work of a Resource Management support team: scheduling practitioners for client-site stock counts, keeping trackers accurate, catching errors, and chasing mandatory training. Built in Excel with a Python exceptions report on top.


What it does
Tracks practitioners, stock-count requests, schedule, leave and mandatory training
Automatically flags double-bookings, people scheduled during leave, and understaffed requests
Tracks SLA ageing (target: 5 working days from receipt to closure)
Lists employees who need chasing for overdue training
Shows everything on a one-screen dashboard
Produces a weekly exceptions report with a Python script
Files
File	Purpose
data/ResourceOps_Tracker.xlsx	The Excel workbook (6 sheets, 5 tables, dashboard)
data/planted_errors_answer_key.csv	The 7 errors deliberately planted in the data, used to test the checks
data/Monday_Exceptions_Report_<date>.xlsx	Output of the report script
scripts/generate_data.py	Generates the synthetic data and plants 7 known errors
scripts/exceptions_report.py	Re-runs every check in Python and writes the weekly report
Workbook structure

Tables are linked by IDs, never by names: PractID links Practitioners, Schedule, Leave and Training; RequestID links Requests and Schedule.

Check	Where	Logic
Double-booking	Schedule[DoubleBooked]	Same PractID and Date appears more than once
Leave clash	Schedule[LeaveClash]	Date falls between LeaveStart and LeaveEnd for the same person
Staffing gap	Requests[StaffGap]	PeopleNeeded minus people scheduled; Open requests shown as PENDING
SLA ageing	Requests[SLADays]	Working days from DateReceived to DateClosed (or today)
Training chase	Training[ChaseFlag]	Not completed and due date passed

Dropdown validation on Region, Grade, Status, WeekendAvailable and LeaveType prevents typos from breaking the formulas.

How to use (weekly routine)
Add or update rows in the relevant table. New rows join the table automatically.
Use dropdowns for status fields. Use IDs, not names, to link records.
Check the Dashboard: any red number above zero needs action.
Download the workbook into data/.
Run python scripts/exceptions_report.py "data/ResourceOps_Tracker.xlsx" and send the report to the Resource Managers.
Results
7 of 7 planted errors detected (3 double-bookings, 2 leave clashes, 2 understaffed requests), with 0 false positives, in both Excel and Python.
Excel and Python summaries matched on all six metrics.
Test on 3 new errors added to a copy of the workbook: the Python script caught 3 of 3.
11 of 14 closed requests met the 5-working-day SLA (78.6%); 8 training enrolments flagged for chasing.
What I learned
Linking tables by IDs instead of names prevents most lookup errors.
Checking the same logic in two tools (Excel and Python) catches mistakes in either one.
A script that works on clean data can break when someone adds columns to the source file; selecting only the needed columns fixed it.
Setup
pip install pandas faker openpyxl
python scripts/generate_data.py
python scripts/exceptions_report.py
Data protection note

Real versions of these trackers hold personal data. In practice: share only with people who need it, use employee IDs rather than names where possible, avoid personal email or unsecured copies, and follow the organisation's data protection policy.

Assumptions and limitations
The 5-working-day SLA is my own assumption; real targets would come from the team.
Public holidays are not excluded from working-day counts.
Schedule records are single-day; multi-day assignments are not modelled.
Scheduling and forecasting tools used in industry are more complex; this project demonstrates the underlying checks and tracker discipline.
Tools

Microsoft Excel (Tables, COUNTIFS, SUMIFS, FILTER, NETWORKDAYS, data validation, conditional formatting, charts), Python (pandas, openpyxl, Faker), GitHub.
