# CIS Automation Project

This project automates CIS reporting for a construction business.

It pulls reconciled CIS labour payments from Xero, generates employer and subcontractor reports, uploads them to OneDrive via Microsoft Graph, and distributes reports by email.

The report workflow is designed to run manually from GitHub Actions once Xero has been reconciled.

## Key Features

* Pulls CIS-related transactions directly from Xero via API
* Filters and aggregates labour payments
* Calculates correct CIS tax periods (6th → 5th)
* Generates polished employer and subcontractor reports (PDF + CSV)
* Includes subcontractor contact details from Xero contacts where available
* Generates subcontractor Statements of Payment and Deduction
* Uploads reports to structured OneDrive folders
* Emails employer summaries and individual subcontractor reports
* Sends automated reconciliation reminders on the 6th of each month
* Runs from the GitHub Actions button after Xero reconciliation is confirmed

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
* `MAILBOX_USER` (email account used to send reports)
* `EMPLOYER_EMAIL`
* `CONTRACTOR_NAME`
* `CONTRACTOR_ADDRESS`
* `CONTRACTOR_EMPLOYERS_REFERENCE`
* `CONTRACTOR_TAXPAYER_REFERENCE`

### Optional

* `CIS_ACCOUNT_CODE`
* `REQUIRE_REFERENCE_CONTAINS` (optional transaction filter)
* `EXTRA_REPORT_RECIPIENTS` (comma-separated extra employer report recipients)

For local development, copy `.env.example` to `.env` and fill in the values. Do not commit `.env`.

## Installation

```bash
python3 -m pip install -r requirements.txt
```

## Usage

### Run monthly report

```bash
python3 run_cis_reports.py
```

### Generate reports without uploading or emailing

```bash
python3 run_cis_reports.py --reports-only
```

### Send one test subcontractor email

This sends one employee email to the test recipient only. It skips OneDrive uploads and all live subcontractor recipients.

```bash
python3 run_cis_reports.py --test-employee-email test@example.com --employee-name "CHRISTOPHER SWALLOW"
```

### Backfill or test a specific run date

```bash
python3 run_cis_reports.py --reports-only --run-date 2026-05-17
```

### Send reconciliation reminder

```bash
python3 send_reconcile_reminder.py
```

## Automation

The system is intended to run via GitHub Actions:

* 6th of each month: send reconciliation reminder
* On demand: generate and distribute reports once Xero reconciliation is confirmed

Cron schedules in GitHub Actions run in UTC. The monthly report workflow does not have an automatic schedule.

### GitHub Secrets

Configure these as repository secrets:

* `XERO_CLIENT_ID`
* `XERO_CLIENT_SECRET`
* `GRAPH_TENANT_ID`
* `GRAPH_CLIENT_ID`
* `GRAPH_CLIENT_SECRET`
* `MAILBOX_USER`
* `EMPLOYER_EMAIL`
* `CONTRACTOR_NAME`
* `CONTRACTOR_ADDRESS`
* `CONTRACTOR_EMPLOYERS_REFERENCE`
* `CONTRACTOR_TAXPAYER_REFERENCE`
* `CIS_ACCOUNT_CODE`
* `REQUIRE_REFERENCE_CONTAINS` (optional)
* `EXTRA_REPORT_RECIPIENTS` (optional)

No GitHub repository variables are required. The workflow reads all configuration values from repository secrets so they are not visible as plain text.

### Manual GitHub Runs

The monthly reports workflow can be run manually in three modes:

* `live`: uploads reports and emails employer/subcontractors
* `reports-only`: generates reports without uploads or emails
* `test-employee-email`: sends one subcontractor email to a test recipient only

## Notes

* Ensure all required environment variables are configured before running
* Designed to handle incomplete data and missing contact details
* Built for real-world usage with minimal manual intervention

## Next Steps

* Add pytest-based validation to improve reliability
* Introduce additional monitoring and logging for production use
