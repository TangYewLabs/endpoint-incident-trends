## Data Source & Methodology

This project uses **synthetically generated incident data** to demonstrate an end-to-end
incident analysis, trending, and reporting workflow in the absence of access to
confidential production ticketing systems.

### Why Synthetic Data
- Real endpoint and security incident data is typically restricted due to privacy,
  security, and compliance requirements.
- Synthetic data allows safe demonstration of:
  - Incident volume trends
  - Resolution time analysis (MTTR, P95)
  - SLA breach detection
  - Root Cause Analysis (RCA)
  - Executive and endpoint impact reporting

### How the Data Is Generated
A Python data generator (`generate_sample_data.py`) creates a realistic CSV dataset
that simulates:
- 30 days of endpoint and security incidents
- Multiple issue categories (VPN, O365, Teams, AV, EDR, MFA, Phishing, etc.)
- Priority levels (P1–P4) with realistic resolution distributions
- Resolved and unresolved incidents
- Executive vs non-executive users
- Device type, site, network path, and vendor involvement

The output is a structured CSV file (`sample_incidents.csv`) designed to mirror
exports from enterprise ticketing systems such as ServiceNow or Jira Service Management.

### Analysis Pipeline
1. Synthetic incident data is generated as a CSV file
2. The analyzer script ingests the CSV
3. Metrics and trends are calculated (MTTR, SLA breaches, daily volume)
4. RCA-style breakdowns and executive impact views are produced
5. A multi-sheet Excel report with charts is generated automatically

This approach mirrors real-world IT operations workflows where analysts work from
ticket exports rather than live systems.
