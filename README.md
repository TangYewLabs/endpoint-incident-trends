\# Endpoint Incident Trend Analyzer (Python)



A lightweight tool that ingests incident exports (CSV) and produces an Excel report for incident trending, RCA support, and operational review.



\## What it does

\- Trends incidents by category and frequency

\- Summarizes resolution-time patterns

\- Breaks out executive vs non-executive impact

\- Flags basic SLA risk indicators (simple thresholds)

\- Outputs an Excel report with multiple tabs



\## Run locally

```bash

pip install -r requirements.txt

python src/analyze\_incidents.py data/sample\_incidents.csv --out reports/incident\_trends\_report.xlsx



