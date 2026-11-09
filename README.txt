ACT 10 PROPERTY TAX ESTIMATOR
=============================
Start:  double-click "Start Act10 Tool.command" (Mac), or run:  python3 act10_app.py
        Needs Python 3 only (no installs). Your browser opens automatically.

First time: click "Download DPI data". Then pick a school district, enter its staff-cost
inputs (or an added annual cost), and click "Generate report".

Each run saves to reports/<district>_<timestamp>/ :
  report.html, district_summary.csv, home_values.csv, municipalities.csv, parcels.csv

Staff pay data: in the app, Maintenance > "Update staff pay data" downloads DPI's current-year teacher
salary report and All Staff file (names are never saved) and fills the cost table for ~418 districts.
2009-10 baseline = DPI PI-1202 teacher salary + fringe by district (data/teacher_comp_2009_10.csv).
Added cost = (2009-10 comp per teacher x inflation - current comp per teacher) x FTEs, floored at 0.
Default staff basis: all salaried staff (the teacher pay gap applied to every salaried FTE); switch to "Teachers only" in the app.

Cost inputs are kept in data/district_costs.csv (editable in Excel):
  ftes, comp_2009_per_fte (salary+fringe, 2009 $), comp_now_per_fte, or added_cost_override.
Districts with no cost data still get property-value figures and a "per $1M" sensitivity table.

Method: WILL (Feb 2025) "Analysis: Act 10 Repeal - What It Will Cost You", re-implemented.
Files: act10_app.py (web UI), act10_core.py (data + math), wi_assessments.py (parcel downloader CLI).
