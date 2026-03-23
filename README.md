# CIS Automation Project

This project pulls reconciled CIS labour payments from Xero, builds employer and employee CIS reports, uploads them to OneDrive via Microsoft Graph, emails the employer summary to `info@ppmbuilders.com`, emails each subcontractor their own PDF, sends a reconcile reminder on the 6th, and is designed to run in GitHub Actions on the 15th.

## Important
Rotate both your Xero and Microsoft Graph secrets before using this project. Do not keep secrets hardcoded in source files.

## Environment variables

### Xero
- `XERO_CLIENT_ID`
- `XERO_CLIENT_SECRET`

### Microsoft Graph
- `GRAPH_TENANT_ID`
- `GRAPH_CLIENT_ID`
- `GRAPH_CLIENT_SECRET`

### Optional
- `MAILBOX_USER` defaults to `info@ppmbuilders.com`
- `CIS_ACCOUNT_CODE` defaults to `1000`
- `REQUIRE_REFERENCE_CONTAINS` optional extra filter

## Install
```bash
python3 -m pip install -r requirements.txt
```

## Run monthly report
```bash
python3 run_cis_reports.py
```

## Run reminder email
```bash
python3 send_reconcile_reminder.py
```
