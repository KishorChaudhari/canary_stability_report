# Synthetics Canary Stability Report

Fetches CloudWatch Synthetics canary pass rates and failure reasons across CXone production
regions for a date range, writes an Excel report, and builds an HTML email summary from it.

## What it produces

| File | Contents |
|---|---|
| `canary_report_<start>_<end>.xlsx` | Summary, success % matrices, failure reasons, daily pass rate, individual failures, recent runs, info |
| `canary_report_<start>_<end>_success_pct.csv` | Success % matrix: canary (rows) x region (columns) |
| `canary_report_<start>_<end>_email.html` | Email-ready HTML summary built from the .xlsx |

## Requirements

- Python 3.10+
- AWS CLI v2 with access to the IAM Identity Center (SSO) portal
- `pip install -r requirements.txt`

## Sign-in

The script uses AWS IAM Identity Center (SSO). One login covers every region/account:

```
aws configure sso      # one-time setup, use the company SSO start URL
aws sso login
```

For each region, credentials are tried in this order:

1. `--profile REGION=PROFILE` override, if given
2. Default credentials, if they belong to that region's account
3. Profile `wfo-acct-<account>`
4. The cached `aws sso login` token from `~/.aws/sso/cache`, using role `Dept_Dev-Platform-SnC`
   (or the first role available in that account)

If a region fails, the error lists what was tried; usually `aws sso login` again fixes it.

## Usage

```
# Monthly report (UTC) + email
python fetch_canary_report.py --start 2026-10-01 --end 2026-10-31
python build_email_html.py canary_report_2026-10-01_2026-10-31.xlsx

# Selected regions or canaries
python fetch_canary_report.py --start 2026-10-01 --end 2026-10-31 --regions na1 de
python fetch_canary_report.py --start 2026-10-01 --end 2026-10-31 --canaries gla-canary ua-curr-can

# Another time zone for the date range and timestamps
python fetch_canary_report.py --start 2026-10-01 --end 2026-10-31 --tz IST
```

Dates are inclusive: `--start 2026-10-01 --end 2026-10-31` covers 2026-10-01 00:00:00 to
2026-10-31 23:59:59 in the chosen time zone (UTC by default).

## Data sources

- **Pass rate**: CloudWatch metric `CloudWatchSynthetics/SuccessPercent` (one sample per run, kept 15 months)
- **Failure reasons**: "Step failed" and Lambda timeout lines in the canary Lambda logs via Logs Insights
  (log retention is typically 90 days); falls back to `GetCanaryRuns`, which only keeps the last few days

## Verifying numbers in the AWS console

Report in **UTC** when stakeholders verify in the console. In CloudWatch → Synthetics → canary →
Availability, set the timezone to **UTC** and an absolute range of the first day 00:00:00 to the last
day 23:59:59; the Average shown matches the report's Pass %.

Non-UTC reports (e.g. IST, UTC+05:30) can differ slightly from the console, because the console groups
data into hourly buckets aligned to UTC hours, which pulls in a run from just before the range start.
