from datetime import datetime
from pathlib import Path

from graph_ops import (
    build_remote_paths,
    create_view_link,
    ensure_nested_folder,
    get_graph_access_token,
    get_reporting_period_label_for_run,
    safe_name,
    send_email,
    upload_file_to_onedrive,
)
from xero_reports import (
    build_contact_email_map,
    build_reports,
    get_all_bank_transactions,
    transactions_to_dataframe,
)

import os
from dotenv import load_dotenv
load_dotenv()

def get_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value

EMPLOYER_EMAIL = get_env("EMPLOYER_EMAIL")

def main() -> None:
    run_date = datetime.now()
    remote = build_remote_paths(run_date)
    graph_token = get_graph_access_token()

    ensure_nested_folder(graph_token, "CIS Reports")
    ensure_nested_folder(graph_token, remote["employer_root"])
    ensure_nested_folder(graph_token, remote["employer_tax_year"])
    ensure_nested_folder(graph_token, remote["employer_month"])
    ensure_nested_folder(graph_token, remote["employees_root"])

    transactions = get_all_bank_transactions(max_pages=50)
    df = transactions_to_dataframe(transactions)
    result = build_reports(df)

    target_label = get_reporting_period_label_for_run(run_date)

    if target_label not in result["monthly_artifacts"]:
        raise RuntimeError(
            f"Target CIS period '{target_label}' not found in generated reports. "
            f"Available periods: {list(result['monthly_artifacts'].keys())}"
        )

    month_artifacts = result["monthly_artifacts"][target_label]
    current_month_cis = result["monthly_summary_totals"][target_label]

    upload_file_to_onedrive(
        graph_token,
        Path(month_artifacts["summary_pdf"]),
        f"{remote['employer_month']}/monthly_summary.pdf"
    )
    upload_file_to_onedrive(
        graph_token,
        Path(month_artifacts["summary_csv"]),
        f"{remote['employer_month']}/monthly_summary.csv"
    )
    upload_file_to_onedrive(
        graph_token,
        Path(month_artifacts["detailed_pdf"]),
        f"{remote['employer_month']}/detailed_breakdown.pdf"
    )
    upload_file_to_onedrive(
        graph_token,
        Path(month_artifacts["detailed_csv"]),
        f"{remote['employer_month']}/detailed_breakdown.csv"
    )

    latest_month_link = create_view_link(graph_token, remote["employer_month"])
    employer_root_link = create_view_link(graph_token, remote["employer_tax_year"])

    employer_subject = f"CIS Report - {target_label}"
    employer_body = f"""Hi Matt,

The CIS report for {target_label} is ready.

Current month CIS total: £{int(round(current_month_cis))}
Tax year to date CIS total: £{int(round(result['total_cis_ytd']))}

Latest month folder:
{latest_month_link}

Employer root folder:
{employer_root_link}
"""

    send_email(
        token=graph_token,
        to_addresses=[EMPLOYER_EMAIL],
        subject=employer_subject,
        body_text=employer_body,
        attachments=[
            str(month_artifacts["summary_pdf"]),
            str(month_artifacts["detailed_pdf"]),
        ],
    )

    email_map = build_contact_email_map()
    for employee_name, employee_data in result["employee_artifacts"].items():
        employee_email = email_map.get(employee_name)
        if not employee_email:
            print(f"Skipping {employee_name}: no email found in Xero contacts")
            continue

        employee_folder = (
            f"{remote['employees_root']}/{safe_name(employee_name)}/{remote['tax_year']}"
        )
        ensure_nested_folder(graph_token, employee_folder)

        upload_file_to_onedrive(
            graph_token,
            Path(employee_data["pdf"]),
            f"{employee_folder}/cis_summary.pdf"
        )
        upload_file_to_onedrive(
            graph_token,
            Path(employee_data["csv"]),
            f"{employee_folder}/employee_summary.csv"
        )

        employee_subject = f"Your CIS Summary - {target_label}"
        employee_body = f"""Hi {employee_name.title()},

Please find attached your updated CIS summary for {target_label}.

Current month gross: £{int(round(employee_data['current_month_gross']))}
Current month CIS: £{int(round(employee_data['current_month_cis']))}
Tax year to date gross: £{int(round(employee_data['ytd_gross']))}
Tax year to date CIS: £{int(round(employee_data['ytd_cis']))}
"""

        send_email(
            token=graph_token,
            to_addresses=[employee_email],
            subject=employee_subject,
            body_text=employee_body,
            attachments=[str(employee_data["pdf"])],
        )

    print(f"CIS report run complete for {target_label}.")


if __name__ == "__main__":
    main()