import argparse
import os
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
    build_contact_details_map,
    build_reports,
    get_all_bank_transactions,
    normalize_name,
    transactions_to_dataframe,
)

from dotenv import load_dotenv
load_dotenv()

def get_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value

EMPLOYER_EMAIL = get_env("EMPLOYER_EMAIL")

EXTRA_REPORT_RECIPIENTS = [
    email.strip()
    for email in os.getenv("EXTRA_REPORT_RECIPIENTS", "").split(",")
    if email.strip()
]

report_recipients = [EMPLOYER_EMAIL] + EXTRA_REPORT_RECIPIENTS

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate and optionally send CIS reports.")
    parser.add_argument(
        "--run-date",
        help="Optional run date in YYYY-MM-DD format. Useful for backfills and manual workflow runs.",
    )
    parser.add_argument(
        "--reports-only",
        action="store_true",
        help="Generate local reports only. Skips OneDrive uploads and emails.",
    )
    parser.add_argument(
        "--test-employee-email",
        help="Send one employee email to this test address only. Skips OneDrive uploads and live recipients.",
    )
    parser.add_argument(
        "--employee-name",
        help="Employee/subcontractor name to use with --test-employee-email. Defaults to the first generated employee.",
    )
    args = parser.parse_args()
    if args.reports_only and args.test_employee_email:
        parser.error("--reports-only cannot be combined with --test-employee-email")
    return args

def parse_run_date(value: str | None) -> datetime:
    if not value:
        return datetime.now()
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("--run-date must use YYYY-MM-DD format") from exc

def choose_employee_artifact(result: dict, employee_name: str | None) -> tuple[str, dict]:
    artifacts = result["employee_artifacts"]
    if not artifacts:
        raise RuntimeError("No employee reports were generated.")

    if not employee_name:
        chosen_name = sorted(artifacts.keys())[0]
        return chosen_name, artifacts[chosen_name]

    wanted = normalize_name(employee_name)
    for generated_name, generated_data in artifacts.items():
        if normalize_name(generated_name) == wanted:
            return generated_name, generated_data

    available = ", ".join(sorted(artifacts.keys()))
    raise RuntimeError(f"No generated employee report for '{employee_name}'. Available: {available}")

def build_employee_email(
    employee_name: str,
    employee_data: dict,
    contact_details: dict,
    target_label: str,
    test_mode: bool = False,
) -> tuple[str, str]:
    display_name = contact_details.get("name") or employee_name.title()
    first_name = display_name.split()[0] if display_name.split() else employee_name.title()
    subject_prefix = "[TEST] " if test_mode else ""
    subject = f"{subject_prefix}Your CIS Statement - {target_label}"
    test_note = (
        f"TEST EMAIL: This was generated for {display_name}. "
        "No live subcontractor recipient was used.\n\n"
        if test_mode else ""
    )
    body = f"""Hi {first_name},

{test_note}Please find attached your updated CIS year-to-date summary and your Statement of Payment and Deduction for {target_label}.

{target_label} gross: £{int(round(employee_data['current_month_gross']))}
{target_label} CIS: £{int(round(employee_data['current_month_cis']))}
YTD gross: £{int(round(employee_data['ytd_gross']))}
YTD CIS: £{int(round(employee_data['ytd_cis']))}
"""
    return subject, body

def main() -> None:
    args = parse_args()
    run_date = parse_run_date(args.run_date)
    transactions = get_all_bank_transactions(max_pages=50)
    df = transactions_to_dataframe(transactions)
    contact_details_map = build_contact_details_map()
    result = build_reports(df, run_date=run_date, contact_details_map=contact_details_map)

    target_label = get_reporting_period_label_for_run(run_date)

    if target_label not in result["monthly_artifacts"]:
        raise RuntimeError(
            f"Target CIS period '{target_label}' not found in generated reports. "
            f"Available periods: {list(result['monthly_artifacts'].keys())}"
        )

    month_artifacts = result["monthly_artifacts"][target_label]
    current_month_cis = result["monthly_summary_totals"][target_label]

    if args.reports_only:
        print(f"Generated CIS reports for {target_label}.")
        print(f"Output folder: {Path('output').resolve()}")
        print(f"Employer summary: {month_artifacts['summary_pdf']}")
        print(f"Employer detailed breakdown: {month_artifacts['detailed_pdf']}")
        print(f"Employee reports generated: {len(result['employee_artifacts'])}")
        print("Reports-only mode: skipped OneDrive uploads and emails.")
        return

    graph_token = get_graph_access_token()

    if args.test_employee_email:
        employee_name, employee_data = choose_employee_artifact(result, args.employee_name)
        contact_details = contact_details_map.get(normalize_name(employee_name), {})
        employee_subject, employee_body = build_employee_email(
            employee_name,
            employee_data,
            contact_details,
            target_label,
            test_mode=True,
        )
        send_email(
            token=graph_token,
            to_addresses=[args.test_employee_email],
            subject=employee_subject,
            body_text=employee_body,
            attachments=[
                str(employee_data["pdf"]),
                str(employee_data["statement_pdf"]),
            ],
        )
        print(
            f"Sent test employee email for {employee_name} to {args.test_employee_email}. "
            "Skipped OneDrive uploads and live recipients."
        )
        return

    remote = build_remote_paths(run_date)

    ensure_nested_folder(graph_token, "CIS Reports")
    ensure_nested_folder(graph_token, "CIS Reports/employer")
    ensure_nested_folder(graph_token, f"CIS Reports/employer/{remote['tax_year']}")
    ensure_nested_folder(graph_token, remote["employer_month_folder"])
    ensure_nested_folder(graph_token, remote["employees_root"])

    upload_file_to_onedrive(
        graph_token,
        Path(month_artifacts["summary_pdf"]),
        f"{remote['employer_month_folder']}/monthly_summary.pdf"
    )
    upload_file_to_onedrive(
        graph_token,
        Path(month_artifacts["summary_csv"]),
        f"{remote['employer_month_folder']}/monthly_summary.csv"
    )
    upload_file_to_onedrive(
        graph_token,
        Path(month_artifacts["detailed_pdf"]),
        f"{remote['employer_month_folder']}/detailed_breakdown.pdf"
    )
    upload_file_to_onedrive(
        graph_token,
        Path(month_artifacts["detailed_csv"]),
        f"{remote['employer_month_folder']}/detailed_breakdown.csv"
    )

    latest_month_link = create_view_link(graph_token, remote["employer_month_folder"])
    employer_root_link = create_view_link(graph_token, "CIS Reports")

    employer_subject = f"CIS Report - {target_label}"
    employer_body = f"""Hi Matt,

The CIS report for {target_label} is ready.

{target_label} CIS total: £{int(round(current_month_cis))}
YTD CIS total: £{int(round(result['total_cis_ytd']))}

{target_label} folder:
{latest_month_link}

Root folder:
{employer_root_link}
"""

    send_email(
        token=graph_token,
        to_addresses=report_recipients,
        subject=employer_subject,
        body_text=employer_body,
        attachments=[
            str(month_artifacts["summary_pdf"]),
        ],
    )

    for employee_name, employee_data in result["employee_artifacts"].items():
        employee_folder = (
            f"{remote['employees_root']}/{safe_name(employee_name)}/{remote['tax_year']}"
        )
        ensure_nested_folder(graph_token, employee_folder)
        ensure_nested_folder(graph_token, f"{employee_folder}/statements")

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
        upload_file_to_onedrive(
            graph_token,
            Path(employee_data["statement_pdf"]),
            f"{employee_folder}/statement_of_payment_and_deduction.pdf"
        )
        for period_label, statement_path in employee_data.get("statement_pdfs", {}).items():
            upload_file_to_onedrive(
                graph_token,
                Path(statement_path),
                f"{employee_folder}/statements/{safe_name(period_label)}.pdf"
            )

        contact_details = contact_details_map.get(normalize_name(employee_name), {})
        employee_email = (contact_details.get("email") or "").strip()
        if not employee_email:
            print(f"Uploaded files for {employee_name}, but no email found in Xero contacts")
            continue

        employee_subject, employee_body = build_employee_email(
            employee_name,
            employee_data,
            contact_details,
            target_label,
        )

        send_email(
            token=graph_token,
            to_addresses=[employee_email],
            subject=employee_subject,
            body_text=employee_body,
            attachments=[
                str(employee_data["pdf"]),
                str(employee_data["statement_pdf"]),
            ],
        )

    print(f"CIS report run complete for {target_label}.")


if __name__ == "__main__":
    main()
