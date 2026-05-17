import os
import re
import calendar
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests
from fpdf import FPDF

from dotenv import load_dotenv
load_dotenv()

def get_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value

XERO_CLIENT_ID = get_env("XERO_CLIENT_ID")
XERO_CLIENT_SECRET = get_env("XERO_CLIENT_SECRET")

TOKEN_URL = "https://identity.xero.com/connect/token"
BANK_TRANSACTIONS_URL = "https://api.xero.com/api.xro/2.0/BankTransactions"

CIS_ACCOUNT_CODE = os.environ.get("CIS_ACCOUNT_CODE", "1000")
REQUIRE_REFERENCE_CONTAINS = os.environ.get("REQUIRE_REFERENCE_CONTAINS") or None

CURRENT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = CURRENT_DIR / "output"
LOGO_PATH = CURRENT_DIR / "logo" / "report_logo.png"
FALLBACK_LOGO_PATH = CURRENT_DIR / "logo" / "logo.png"
COMPANY_NAME = get_env("CONTRACTOR_NAME")
COMPANY_ADDRESS = get_env("CONTRACTOR_ADDRESS")
CONTRACTOR_EMPLOYERS_REFERENCE = get_env("CONTRACTOR_EMPLOYERS_REFERENCE")
CONTRACTOR_TAXPAYER_REFERENCE = get_env("CONTRACTOR_TAXPAYER_REFERENCE")
REPORT_NOTE = "Prepared from reconciled Xero CIS labour payment records. Please retain this report for your records."

def normalize_name(name: str) -> str:
    name = str(name).replace(",", " ")
    return re.sub(r"\s+", " ", name.strip()).upper()

def parse_xero_date(xero_date: str) -> datetime:
    match = re.search(r"/Date\((\d+)([+-]\d{4})?\)/", str(xero_date))
    if not match:
        raise ValueError(f"Unexpected Xero date format: {xero_date}")
    millis = int(match.group(1))
    return datetime.fromtimestamp(millis / 1000, tz=timezone.utc)

def get_tax_period_start(dt: pd.Timestamp) -> pd.Timestamp:
    if pd.isnull(dt):
        return pd.NaT
    if dt.day >= 6:
        return pd.Timestamp(year=dt.year, month=dt.month, day=6)
    prev = dt - pd.DateOffset(months=1)
    return pd.Timestamp(year=prev.year, month=prev.month, day=6)

def get_cis_tax_year_start(dt: pd.Timestamp) -> pd.Timestamp:
    if pd.isnull(dt):
        return pd.NaT
    if (dt.month, dt.day) >= (4, 6):
        return pd.Timestamp(year=dt.year, month=4, day=6)
    return pd.Timestamp(year=dt.year - 1, month=4, day=6)

def get_reporting_period_start_for_run(run_date: datetime) -> pd.Timestamp:
    if run_date.month == 1:
        return pd.Timestamp(year=run_date.year - 1, month=12, day=6)
    return pd.Timestamp(year=run_date.year, month=run_date.month - 1, day=6)

def get_access_token() -> str:
    response = requests.post(
        TOKEN_URL,
        data={"grant_type": "client_credentials"},
        auth=(XERO_CLIENT_ID, XERO_CLIENT_SECRET),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    response.raise_for_status()
    payload = response.json()
    token = payload.get("access_token")
    if not token:
        raise RuntimeError(f"No access token returned: {payload}")
    return token

def build_headers(access_token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}

def get_bank_transactions(access_token: str, page: int = 1, if_modified_since: Optional[datetime] = None) -> Dict[str, Any]:
    headers = build_headers(access_token)
    if if_modified_since is not None:
        if if_modified_since.tzinfo is None:
            if_modified_since = if_modified_since.replace(tzinfo=timezone.utc)
        headers["If-Modified-Since"] = if_modified_since.strftime("%a, %d %b %Y %H:%M:%S GMT")
    response = requests.get(BANK_TRANSACTIONS_URL, headers=headers, params={"page": page}, timeout=30)
    response.raise_for_status()
    return response.json()

def get_all_bank_transactions(if_modified_since: Optional[datetime] = None, max_pages: int = 50) -> List[Dict[str, Any]]:
    token = get_access_token()
    all_rows: List[Dict[str, Any]] = []
    for page in range(1, max_pages + 1):
        data = get_bank_transactions(token, page, if_modified_since)
        rows = data.get("BankTransactions", [])
        if not rows:
            break
        all_rows.extend(rows)
    return all_rows

def txn_is_cis(txn: Dict[str, Any]) -> bool:
    if txn.get("Type") != "SPEND":
        return False
    if txn.get("Status") != "AUTHORISED":
        return False
    line_items = txn.get("LineItems", [])
    if not line_items:
        return False
    has_cis_account = any(str(item.get("AccountCode")) == CIS_ACCOUNT_CODE for item in line_items)
    if not has_cis_account:
        return False
    if REQUIRE_REFERENCE_CONTAINS:
        reference = str(txn.get("Reference") or "").lower()
        if REQUIRE_REFERENCE_CONTAINS.lower() not in reference:
            return False
    return True

def transactions_to_dataframe(transactions: List[Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for txn in transactions:
        if not txn_is_cis(txn):
            continue
        contact = txn.get("Contact") or {}
        contact_name = contact.get("Name") or "UNKNOWN CONTACT"
        txn_date = parse_xero_date(txn.get("Date"))
        total = float(txn.get("Total") or 0.0)
        rows.append({
            "Date": pd.Timestamp(txn_date).tz_localize(None),
            "To": normalize_name(contact_name),
            "Paid out": total,
            "Reference": txn.get("Reference"),
            "BankTransactionID": txn.get("BankTransactionID"),
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values(["To", "Date", "BankTransactionID"], na_position="last")

def get_contacts() -> List[Dict[str, Any]]:
    token = get_access_token()
    url = "https://api.xero.com/api.xro/2.0/Contacts"
    contacts: List[Dict[str, Any]] = []
    page = 1
    while True:
        response = requests.get(url, headers=build_headers(token), params={"page": page}, timeout=30)
        response.raise_for_status()
        rows = response.json().get("Contacts", [])
        if not rows:
            break
        contacts.extend(rows)
        page += 1
    return contacts

def build_contact_email_map() -> Dict[str, str]:
    contacts = build_contact_details_map()
    mapping: Dict[str, str] = {}
    for name, contact in contacts.items():
        email = (contact.get("email") or "").strip()
        if email:
            mapping[name] = email
    return mapping

def format_contact_address(contact: Dict[str, Any]) -> str:
    addresses = contact.get("Addresses") or []
    preferred = None
    for address in addresses:
        if address.get("AddressType") == "POBOX":
            preferred = address
            break
    if preferred is None and addresses:
        preferred = addresses[0]
    if not preferred:
        return ""

    parts = [
        preferred.get("AddressLine1"),
        preferred.get("AddressLine2"),
        preferred.get("AddressLine3"),
        preferred.get("AddressLine4"),
        preferred.get("City"),
        preferred.get("Region"),
        preferred.get("PostalCode"),
        preferred.get("Country"),
    ]
    return "\n".join(str(part).strip() for part in parts if str(part or "").strip())

def build_contact_details_map() -> Dict[str, Dict[str, str]]:
    contacts = get_contacts()
    mapping: Dict[str, Dict[str, str]] = {}
    for contact in contacts:
        name = normalize_name(contact.get("Name") or "")
        if not name:
            continue
        mapping[name] = {
            "name": contact.get("Name") or name,
            "email": (contact.get("EmailAddress") or "").strip(),
            "utr": (contact.get("AccountNumber") or "").strip(),
            "address": format_contact_address(contact),
        }
    return mapping

def add_report_header(pdf: FPDF, title: str, subtitle: Optional[str] = None) -> None:
    is_a4 = pdf.w < 250
    margin = 14 if is_a4 else 10
    top = 10 if is_a4 else 8
    logo_width = 34 if is_a4 else 65
    details_width = 118 if is_a4 else 220

    logo_path = LOGO_PATH if LOGO_PATH.exists() else FALLBACK_LOGO_PATH
    if logo_path.exists():
        pdf.image(str(logo_path), x=margin, y=top, w=logo_width)

    pdf.set_xy(pdf.w - margin - details_width, top + 1)
    pdf.set_font("Arial", "B", 9 if is_a4 else 12)
    pdf.cell(details_width, 5, COMPANY_NAME, ln=True, align="R")
    pdf.set_x(pdf.w - margin - details_width)
    pdf.set_font("Arial", "", 7 if is_a4 else 9)
    pdf.multi_cell(details_width, 3.8 if is_a4 else 4, COMPANY_ADDRESS, align="R")

    line_y = 33 if is_a4 else 35
    pdf.set_draw_color(210, 210, 210)
    pdf.line(margin, line_y, pdf.w - margin, line_y)
    pdf.set_draw_color(0, 0, 0)

    pdf.set_y(43 if is_a4 else 41)
    pdf.set_font("Arial", "B", 14 if is_a4 else 15)
    pdf.cell(0, 7, title, ln=True, align="C")
    if subtitle:
        pdf.set_font("Arial", "", 8.5 if is_a4 else 10)
        pdf.cell(0, 5, subtitle, ln=True, align="C")
    pdf.ln(6 if is_a4 else 4)

def add_report_footer(pdf: FPDF) -> None:
    margin = 10
    pdf.set_y(pdf.h - 18)
    pdf.set_draw_color(210, 210, 210)
    pdf.line(margin, pdf.get_y(), pdf.w - margin, pdf.get_y())
    pdf.set_draw_color(0, 0, 0)
    pdf.ln(2)
    pdf.set_font("Arial", "", 7 if pdf.w < 300 else 8)
    pdf.cell(
        0,
        4,
        f"Generated {datetime.now().strftime('%d %B %Y')} | {REPORT_NOTE}",
        ln=True,
        align="C",
    )

def add_generated_note_page(pdf: FPDF) -> None:
    margin = 14
    pdf.add_page()
    add_report_header(pdf, "Report Notes")
    pdf.set_y(78)
    pdf.set_x(margin)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "Generated", ln=True)
    pdf.set_x(margin)
    pdf.set_font("Arial", "", 9)
    pdf.cell(0, 6, datetime.now().strftime("%d %B %Y"), ln=True)
    pdf.ln(6)
    pdf.set_x(margin)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "Basis of preparation", ln=True)
    pdf.set_x(margin)
    pdf.set_font("Arial", "", 9)
    pdf.multi_cell(pdf.w - (margin * 2), 5, REPORT_NOTE)

def money(value: float) -> str:
    return f"£{float(value):,.2f}"

def format_address_for_pdf(address: str) -> str:
    return "\n".join(line for line in str(address or "").splitlines() if line.strip())

def add_info_panel(pdf: FPDF, title: str, fields: List[tuple], x: float, y: float, width: float, height: float) -> None:
    pdf.set_draw_color(205, 205, 205)
    pdf.set_fill_color(247, 247, 247)
    pdf.rect(x, y, width, height, style="DF")
    pdf.set_fill_color(235, 235, 235)
    pdf.rect(x, y, width, 8, style="F")

    pdf.set_xy(x + 4, y + 2)
    pdf.set_font("Arial", "B", 8.5)
    pdf.cell(width - 8, 4, title)

    cursor_y = y + 12
    label_width = 26
    value_width = width - label_width - 10
    for label, value in fields:
        pdf.set_xy(x + 4, cursor_y)
        pdf.set_font("Arial", "B", 7.5)
        pdf.cell(label_width, 4.5, label)
        pdf.set_xy(x + 4 + label_width, cursor_y)
        pdf.set_font("Arial", "", 7.5)
        pdf.multi_cell(value_width, 4.5, value or "N/A")
        cursor_y = max(cursor_y + 5, pdf.get_y() + 1)

    pdf.set_draw_color(0, 0, 0)

def add_contractor_details(pdf: FPDF) -> None:
    margin = 14
    width = pdf.w - (margin * 2)
    y = pdf.get_y()
    fields = [
        ("Name", COMPANY_NAME),
        ("Emp. ref", CONTRACTOR_EMPLOYERS_REFERENCE),
        ("Tax ref", CONTRACTOR_TAXPAYER_REFERENCE),
    ]
    add_info_panel(pdf, "Contractor details", fields, margin, y, width, 28)
    pdf.set_y(y + 36)

def add_table_header(pdf: FPDF, headers: List[str], widths: List[float], margin: float) -> None:
    pdf.set_x(margin)
    pdf.set_font("Arial", "B", 8.5)
    pdf.set_fill_color(235, 235, 235)
    pdf.set_draw_color(190, 190, 190)
    for index, header in enumerate(headers):
        pdf.cell(widths[index], 8, header, border=1, align="C", fill=True)
    pdf.ln(8)

def create_monthly_pdf(
    period_label: str,
    period_df: pd.DataFrame,
    period_start: pd.Timestamp,
    period_end: pd.Timestamp,
    pdf_output_path: Path,
    contact_details_map: Optional[Dict[str, Dict[str, str]]] = None,
) -> None:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=22)

    for employee, group in period_df.sort_values(["To", "Date"]).groupby("To"):
        pdf.add_page()
        add_report_header(
            pdf,
            f"CIS Payment Detail: {employee.title()}",
            f"Employer payment record for {period_label}: {period_start.strftime('%d/%m/%Y')} to {period_end.strftime('%d/%m/%Y')}",
        )
        add_cis_details(
            pdf,
            employee,
            (contact_details_map or {}).get(normalize_name(employee)),
        )

        margin = 14
        table_width = pdf.w - (margin * 2)
        widths = [34, table_width - 34 - 42 - 42 - 42, 42, 42, 42]
        headers = ["Date", "Reference", "Gross", "CIS Deducted", "Net Paid"]
        add_table_header(pdf, headers, widths, margin)

        totals = {"gross": 0.0, "cis": 0.0, "net": 0.0}
        pdf.set_font("Arial", "", 8.5)
        pdf.set_draw_color(210, 210, 210)
        for _, row in group.iterrows():
            net_paid = float(row["Paid out"] or 0.0)
            gross = net_paid / 0.8
            cis = net_paid * 0.25
            totals["gross"] += gross
            totals["cis"] += cis
            totals["net"] += net_paid
            values = [
                row["Date"].strftime("%d/%m/%Y"),
                str(row.get("Reference") or ""),
                money(gross),
                money(cis),
                money(net_paid),
            ]

            pdf.set_x(margin)
            for index, value in enumerate(values):
                align = "L" if index == 1 else "R" if index >= 2 else "C"
                pdf.cell(widths[index], 7, value[:48], border=1, align=align)
            pdf.ln(7)

        pdf.set_x(margin)
        pdf.set_font("Arial", "B", 8.5)
        pdf.set_fill_color(248, 248, 248)
        pdf.cell(widths[0] + widths[1], 8, "Period total", border=1, align="R", fill=True)
        pdf.cell(widths[2], 8, money(totals["gross"]), border=1, align="R", fill=True)
        pdf.cell(widths[3], 8, money(totals["cis"]), border=1, align="R", fill=True)
        pdf.cell(widths[4], 8, money(totals["net"]), border=1, align="R", fill=True)
        pdf.set_draw_color(0, 0, 0)

    add_generated_note_page(pdf)
    pdf.output(str(pdf_output_path))

def create_monthly_summary_pdf(period_label: str, df_summary: pd.DataFrame, pdf_output_path: Path) -> None:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=22)
    pdf.add_page()
    add_report_header(
        pdf,
        f"Monthly CIS Summary: {period_label}",
        "Employer summary of gross payments, CIS deductions and net paid by subcontractor.",
    )
    add_contractor_details(pdf)

    margin = 14
    table_width = pdf.w - (margin * 2)
    widths = [table_width - 42 - 42 - 42, 42, 42, 42]
    add_table_header(pdf, ["Subcontractor", "Gross", "CIS Deducted", "Net Paid"], widths, margin)

    summary_cols = ["Gross", "-20% CIS", "Total"]
    sorted_summary = df_summary.sort_index()
    pdf.set_font("Arial", "", 8.5)
    pdf.set_draw_color(210, 210, 210)
    for employee, row in sorted_summary.iterrows():
        pdf.set_x(margin)
        pdf.cell(widths[0], 7, str(employee).title()[:60], border=1)
        pdf.cell(widths[1], 7, money(row["Gross"]), border=1, align="R")
        pdf.cell(widths[2], 7, money(row["-20% CIS"]), border=1, align="R")
        pdf.cell(widths[3], 7, money(row["Total"]), border=1, ln=True, align="R")

    totals = df_summary[summary_cols].sum()
    pdf.set_x(margin)
    pdf.set_font("Arial", "B", 8.5)
    pdf.set_fill_color(248, 248, 248)
    pdf.cell(widths[0], 8, "Total", border=1, align="R", fill=True)
    pdf.cell(widths[1], 8, money(totals["Gross"]), border=1, align="R", fill=True)
    pdf.cell(widths[2], 8, money(totals["-20% CIS"]), border=1, align="R", fill=True)
    pdf.cell(widths[3], 8, money(totals["Total"]), border=1, ln=True, align="R", fill=True)
    pdf.set_draw_color(0, 0, 0)
    add_report_footer(pdf)
    pdf.output(str(pdf_output_path))

def add_cis_details(pdf: FPDF, employee_name: str, contact_details: Optional[Dict[str, str]]) -> None:
    details = contact_details or {}
    margin = 14
    gap = 6
    panel_width = (pdf.w - (margin * 2) - gap) / 2
    panel_height = 50
    y = pdf.get_y()

    subcontractor_fields = [
        ("Name", details.get("name") or employee_name.title()),
        ("UTR", details.get("utr") or "N/A"),
        ("Address", format_address_for_pdf(details.get("address") or "N/A")),
    ]
    contractor_fields = [
        ("Name", COMPANY_NAME),
        ("Emp. ref", CONTRACTOR_EMPLOYERS_REFERENCE),
        ("Tax ref", CONTRACTOR_TAXPAYER_REFERENCE),
    ]

    add_info_panel(pdf, "Subcontractor", subcontractor_fields, margin, y, panel_width, panel_height)
    add_info_panel(pdf, "Contractor", contractor_fields, margin + panel_width + gap, y, panel_width, panel_height)
    pdf.set_y(y + panel_height + 8)

def create_employee_pdf(
    employee_name: str,
    summary_df: pd.DataFrame,
    pdf_output_path: Path,
    contact_details: Optional[Dict[str, str]] = None,
) -> None:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    margin = 14
    available_width = 210 - (margin * 2)
    col_width = available_width / 4
    add_report_header(
        pdf,
        f"CIS Statement: {employee_name.title()}",
        "Year-to-date subcontractor CIS payment summary.",
    )
    add_cis_details(pdf, employee_name, contact_details)

    pdf.set_x(margin)
    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(235, 235, 235)
    pdf.set_draw_color(190, 190, 190)
    for index, heading in enumerate(["Tax Period", "Gross", "CIS Deducted", "Net Paid"]):
        pdf.cell(col_width, 8, heading, border=1, ln=index == 3, align="C", fill=True)

    pdf.set_font("Arial", "", 9)
    for _, row in summary_df.iterrows():
        is_total = str(row["TaxPeriod"]) == "YEAR TOTAL"
        pdf.set_x(margin)
        pdf.set_font("Arial", "B" if is_total else "", 9)
        pdf.set_fill_color(248, 248, 248) if is_total else pdf.set_fill_color(255, 255, 255)
        pdf.cell(col_width, 8, str(row["TaxPeriod"]), border=1, align="C", fill=is_total)
        pdf.cell(col_width, 8, money(row["Gross"]), border=1, align="R", fill=is_total)
        pdf.cell(col_width, 8, money(row["CIS"]), border=1, align="R", fill=is_total)
        pdf.cell(col_width, 8, money(row["Total"]), border=1, ln=True, align="R", fill=is_total)
    pdf.set_draw_color(0, 0, 0)
    add_report_footer(pdf)
    pdf.output(str(pdf_output_path))

def create_payment_deduction_statement_pdf(
    employee_name: str,
    contact_details: Optional[Dict[str, str]],
    period_start: pd.Timestamp,
    period_end: pd.Timestamp,
    gross_amount: float,
    cis_deducted: float,
    pdf_output_path: Path,
) -> None:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    add_report_header(
        pdf,
        "Statement of Payment and Deduction",
        f"CIS tax period: {period_start.strftime('%d/%m/%Y')} to {period_end.strftime('%d/%m/%Y')}",
    )
    add_cis_details(pdf, employee_name, contact_details)

    material_cost = 0.0
    liable_amount = gross_amount - material_cost

    margin = 14
    table_width = pdf.w - (margin * 2)
    label_width = 118
    value_width = table_width - label_width

    pdf.set_x(margin)
    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(235, 235, 235)
    pdf.set_draw_color(190, 190, 190)
    pdf.cell(table_width, 8, "Payment and deduction breakdown", border=1, ln=True, fill=True)

    rows = [
        ("Gross amount of payment", gross_amount),
        ("Less cost of materials", material_cost),
        ("Amount liable to deduction", liable_amount),
        ("Amount deducted", cis_deducted),
    ]

    for label, value in rows:
        pdf.set_x(margin)
        is_deduction = label == "Amount deducted"
        pdf.set_font("Arial", "B" if is_deduction else "", 9)
        pdf.set_fill_color(248, 248, 248) if is_deduction else pdf.set_fill_color(255, 255, 255)
        pdf.cell(label_width, 8, label, border=1, fill=is_deduction)
        pdf.cell(value_width, 8, money(value), border=1, ln=True, align="R", fill=is_deduction)

    pdf.set_draw_color(0, 0, 0)
    pdf.ln(5)
    pdf.set_x(margin)
    pdf.set_font("Arial", "", 8)
    pdf.multi_cell(
        table_width,
        4,
        "This statement is prepared from reconciled CIS labour payment records in Xero.",
    )
    add_report_footer(pdf)
    pdf.output(str(pdf_output_path))

def build_reports(
    df: pd.DataFrame,
    run_date: Optional[datetime] = None,
    contact_details_map: Optional[Dict[str, Dict[str, str]]] = None,
) -> Dict[str, Any]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if df.empty:
        raise RuntimeError("No CIS transactions found.")
    if run_date is None:
        run_date = datetime.now()

    target_period_start = get_reporting_period_start_for_run(run_date)
    target_label = target_period_start.strftime("%B %Y")
    target_period_end = target_period_start + pd.DateOffset(months=1) - pd.DateOffset(days=1)
    tax_year_start = get_cis_tax_year_start(target_period_start)
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month
    df["Day"] = df["Date"].dt.day
    monthly_pivots = {}
    for (year, month), group in df.groupby(["Year", "Month"]):
        num_days = calendar.monthrange(year, month)[1]
        pivot = group.pivot_table(index="To", columns="Day", values="Paid out", aggfunc="sum", fill_value=0)
        pivot = pivot.reindex(range(1, num_days + 1), axis=1, fill_value=0)
        monthly_pivots[(year, month)] = pivot

    detailed_dir = OUTPUT_DIR / "monthly_cis_returns"
    monthly_summary_dir = OUTPUT_DIR / "monthly_summary"
    employee_output_dir = OUTPUT_DIR / "employee_totals"
    employee_statement_dir = OUTPUT_DIR / "employee_statements"
    detailed_dir.mkdir(parents=True, exist_ok=True)
    monthly_summary_dir.mkdir(parents=True, exist_ok=True)
    employee_output_dir.mkdir(parents=True, exist_ok=True)
    employee_statement_dir.mkdir(parents=True, exist_ok=True)

    monthly_combined_list: List[pd.DataFrame] = []
    monthly_artifacts: Dict[str, Dict[str, Path]] = {}

    for (year, month) in sorted(monthly_pivots.keys()):
        current_pivot = monthly_pivots[(year, month)]
        current_part = current_pivot.loc[:, current_pivot.columns >= 6].copy()
        current_part.columns = [pd.Timestamp(year=year, month=month, day=d) for d in current_part.columns]
        next_key = (year, month + 1) if month < 12 else (year + 1, 1)
        if next_key in monthly_pivots:
            next_pivot = monthly_pivots[next_key]
            next_part = next_pivot.loc[:, next_pivot.columns <= 5].copy()
            next_year, next_month = next_key
            next_part.columns = [pd.Timestamp(year=next_year, month=next_month, day=d) for d in next_part.columns]
        else:
            next_part = pd.DataFrame()

        combined = pd.concat([current_part, next_part], axis=1)
        combined = combined.reindex(sorted(combined.columns), axis=1)
        combined.columns = combined.columns.strftime("%d")
        total_series = combined.sum(axis=1)
        combined["Gross"] = total_series / 0.8
        combined["-20% CIS"] = total_series * 0.25
        combined["Total"] = total_series
        daily_cols = [col for col in combined.columns if col not in ["Gross", "-20% CIS", "Total"]]
        combined = combined[daily_cols + ["Gross", "-20% CIS", "Total"]]
        period_start = pd.Timestamp(year=year, month=month, day=6)
        period_label = period_start.strftime("%B %Y")
        csv_output_path = detailed_dir / f"{period_label}.csv"
        pdf_output_path = detailed_dir / f"{period_label}.pdf"
        period_end = period_start + pd.DateOffset(months=1) - pd.DateOffset(days=1)
        period_df = df[
            (df["Date"] >= period_start) &
            (df["Date"] <= period_end)
        ].copy()
        combined.to_csv(csv_output_path)
        create_monthly_pdf(
            period_label,
            period_df,
            period_start,
            period_end,
            pdf_output_path,
            contact_details_map,
        )
        temp = combined[["Gross", "-20% CIS", "Total"]].copy()
        temp["Period"] = period_label
        temp["Employee"] = temp.index
        monthly_combined_list.append(temp)
        monthly_artifacts[period_label] = {
            "period_start": period_start,
            "detailed_csv": csv_output_path,
            "detailed_pdf": pdf_output_path,
        }

    all_monthly_summary = pd.concat(monthly_combined_list)
    for period_label, group in all_monthly_summary.groupby("Period"):
        summary_df = group.set_index("Employee")[["Gross", "-20% CIS", "Total"]]
        summary_pdf_output_path = monthly_summary_dir / f"{period_label}_summary.pdf"
        summary_csv_output_path = monthly_summary_dir / f"{period_label}_summary.csv"
        summary_df.to_csv(summary_csv_output_path)
        create_monthly_summary_pdf(period_label, summary_df, summary_pdf_output_path)
        monthly_artifacts[period_label]["summary_pdf"] = summary_pdf_output_path
        monthly_artifacts[period_label]["summary_csv"] = summary_csv_output_path

    df["TaxPeriodStart"] = df["Date"].apply(get_tax_period_start)
    df["TaxPeriod"] = df["TaxPeriodStart"].dt.strftime("%B %Y")
    df["CisTaxYearStart"] = df["TaxPeriodStart"].apply(get_cis_tax_year_start)

    employee_summary = df.groupby(
        ["To", "TaxPeriodStart", "TaxPeriod", "CisTaxYearStart"],
        as_index=False
    )["Paid out"].sum()

    employee_summary.rename(columns={"Paid out": "Total"}, inplace=True)
    employee_summary["Gross"] = employee_summary["Total"] / 0.8
    employee_summary["CIS"] = employee_summary["Total"] * 0.25
    employee_summary.sort_values(by=["To", "TaxPeriodStart"], inplace=True)

    employee_artifacts: Dict[str, Dict[str, Any]] = {}

    for employee, group in employee_summary.groupby("To"):
        tax_year_group = group[
            (group["CisTaxYearStart"] == tax_year_start) &
            (group["TaxPeriodStart"] <= target_period_start)
        ].copy()
        if tax_year_group.empty:
            continue

        output_df = tax_year_group[["TaxPeriod", "Gross", "CIS", "Total"]].copy()

        totals_row = {
            "TaxPeriod": "YEAR TOTAL",
            "Gross": output_df["Gross"].sum(),
            "CIS": output_df["CIS"].sum(),
            "Total": output_df["Total"].sum(),
        }

        output_df_with_total = pd.concat(
            [output_df, pd.DataFrame([totals_row])],
            ignore_index=True
        )

        safe_emp_name = "".join(c for c in employee if c.isalnum() or c in " _-").strip()
        emp_csv_path = employee_output_dir / f"{safe_emp_name}.csv"
        emp_pdf_path = employee_output_dir / f"{safe_emp_name}.pdf"
        contact_details = (contact_details_map or {}).get(employee)

        output_df_with_total.to_csv(emp_csv_path, index=False)
        create_employee_pdf(employee, output_df_with_total, emp_pdf_path, contact_details)

        statement_pdfs: Dict[str, Path] = {}
        current_statement_pdf_path = employee_statement_dir / f"{safe_emp_name}_{target_period_start.strftime('%Y-%m')}_statement.pdf"

        for _, statement_row in tax_year_group.iterrows():
            statement_start = statement_row["TaxPeriodStart"]
            statement_end = statement_start + pd.DateOffset(months=1) - pd.DateOffset(days=1)
            statement_path = employee_statement_dir / f"{safe_emp_name}_{statement_start.strftime('%Y-%m')}_statement.pdf"
            create_payment_deduction_statement_pdf(
                employee_name=employee,
                contact_details=contact_details,
                period_start=statement_start,
                period_end=statement_end,
                gross_amount=float(statement_row["Gross"]),
                cis_deducted=float(statement_row["CIS"]),
                pdf_output_path=statement_path,
            )
            statement_pdfs[statement_row["TaxPeriod"]] = statement_path

        current_month_match = tax_year_group[tax_year_group["TaxPeriodStart"] == target_period_start]

        if current_month_match.empty:
            current_month_gross = 0.0
            current_month_cis = 0.0
            create_payment_deduction_statement_pdf(
                employee_name=employee,
                contact_details=contact_details,
                period_start=target_period_start,
                period_end=target_period_end,
                gross_amount=current_month_gross,
                cis_deducted=current_month_cis,
                pdf_output_path=current_statement_pdf_path,
            )
        else:
            current_month_gross = float(current_month_match["Gross"].iloc[0])
            current_month_cis = float(current_month_match["CIS"].iloc[0])
            current_statement_pdf_path = statement_pdfs.get(
                current_month_match["TaxPeriod"].iloc[0],
                current_statement_pdf_path,
            )

        employee_artifacts[employee] = {
            "pdf": emp_pdf_path,
            "csv": emp_csv_path,
            "statement_pdf": current_statement_pdf_path,
            "statement_pdfs": statement_pdfs,
            "current_month_gross": current_month_gross,
            "current_month_cis": current_month_cis,
            "ytd_gross": float(output_df["Gross"].sum()),
            "ytd_cis": float(output_df["CIS"].sum()),
        }

    employee_list_output = OUTPUT_DIR / "employee_list.csv"
    pd.DataFrame(sorted(df["To"].unique()), columns=["Employee"]).to_csv(employee_list_output, index=False)
    target_tax_year_df = df[
        (df["Date"] >= tax_year_start) &
        (df["Date"] <= target_period_end)
    ].copy()

    overall_total = float(target_tax_year_df["Paid out"].sum())
    overall_cis = overall_total * 0.25
    latest_label = max(
        monthly_artifacts.keys(),
        key=lambda label: monthly_artifacts[label]["period_start"]
    )
    latest_period_summary = all_monthly_summary[all_monthly_summary["Period"] == latest_label]
    latest_month_cis = float(latest_period_summary["-20% CIS"].sum())

    monthly_summary_totals = {
        period_label: float(
            all_monthly_summary[
                all_monthly_summary["Period"] == period_label
            ]["-20% CIS"].sum()
        )
        for period_label in monthly_artifacts.keys()
    }

    return {
        "latest_label": latest_label,
        "latest_month_cis": latest_month_cis,
        "total_cis_ytd": float(overall_cis),
        "monthly_summary_totals": monthly_summary_totals,
        "monthly_artifacts": monthly_artifacts,
        "employee_artifacts": employee_artifacts,
        "employee_list_csv": employee_list_output,
    }

if __name__ == "__main__":
    transactions = get_all_bank_transactions(max_pages=50)
    df = transactions_to_dataframe(transactions)
    result = build_reports(df, run_date=datetime.now())
    print(result)
