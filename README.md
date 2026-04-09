# CIS Automation Project

This project automates CIS reporting for a construction business.

It pulls reconciled CIS labour payments from Xero, generates employer and subcontractor reports, uploads them to OneDrive via Microsoft Graph, and distributes reports automatically via email.

The system is designed to run as a scheduled pipeline using GitHub Actions.

## Key Features

* Pulls CIS-related transactions directly from Xero via API
* Filters and aggregates labour payments
* Calculates correct CIS tax periods (6th → 5th)
* Generates employer and subcontractor reports (PDF + CSV)
* Uploads reports to structured OneDrive folders
* Emails employer summaries and individual subcontractor reports
* Sends automated reconciliation reminders on the 6th of each month
* Runs automatically on the 10th via GitHub Actions

## Important

Do not store credentials in source files. This project is designed to use environment variables and GitHub Actions secrets for secure execution.

## Environment Variables

### Xero

* `XERO_CLIENT_ID`
* `XERO_CLIENT_SECRET`

### Microsoft Graph

* `GRAPH_TENANT_ID`
* `GRAPH_CLIENT_ID`
* `GRAPH_CLIENT_SECRET`

### Optional

* `MAILBOX_USER` (email account used to send reports)
* `CIS_ACCOUNT_CODE`
* `REQUIRE_REFERENCE_CONTAINS` (optional transaction filter)

## Installation

```bash
python3 -m pip install -r requirements.txt
```

## Usage

### Run monthly report

```bash
python3 run_cis_reports.py
```

### Send reconciliation reminder

```bash
python3 send_reconcile_reminder.py
```

## Automation

The system is intended to run automatically via GitHub Actions:

* 6th of each month: send reconciliation reminder
* 15th of each month: generate and distribute reports

## Notes

* Ensure all required environment variables are configured before running
* Designed to handle incomplete data and missing contact details
* Built for real-world usage with minimal manual intervention

## Next Steps

* Add pytest-based validation to improve reliability
* Introduce additional monitoring and logging for production use
