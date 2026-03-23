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
    return df.groupby(["To", "Date"], as_index=False)["Paid out"].sum()

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
    contacts = get_contacts()
    mapping: Dict[str, str] = {}
    for contact in contacts:
        name = normalize_name(contact.get("Name") or "")
        email = (contact.get("EmailAddress") or "").strip()
        if name and email:
            mapping[name] = email
    return mapping

def create_monthly_pdf(period_label: str, df_month: pd.DataFrame, pdf_output_path: Path) -> None:
    pdf = FPDF(orientation="L", unit="mm", format="A3")
    pdf.set_auto_page_break(auto=False, margin=10)
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f"Detailed Tax-Period Breakdown: {period_label}", ln=True, align="C")
    pdf.ln(4)
    available_width = 420 - 20
    total_cols = 1 + len(df_month.columns)
    cell_width = available_width / total_cols
    pdf.set_font("Arial", "B", 6)
    pdf.cell(cell_width, 6, "Employee", border=1, align="C")
    for col_name in df_month.columns:
        pdf.cell(cell_width, 6, str(col_name), border=1, align="C")
    pdf.ln(6)
    pdf.set_font("Arial", "", 6)
    for employee, row in df_month.iterrows():
        pdf.cell(cell_width, 6, str(employee), border=1, align="C")
        for col in df_month.columns:
            val = row[col]
            cell_text = "" if pd.isna(val) or float(val) == 0 else f"{int(round(val))}"
            pdf.cell(cell_width, 6, cell_text, border=1, align="C")
        pdf.ln(6)
    pdf.output(str(pdf_output_path))

def create_monthly_summary_pdf(period_label: str, df_summary: pd.DataFrame, pdf_output_path: Path) -> None:
    pdf = FPDF(orientation="L", unit="mm", format="A3")
    pdf.set_auto_page_break(auto=False, margin=10)
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f"Monthly Summary: {period_label}", ln=True, align="C")
    pdf.ln(4)
    summary_cols = ["Gross", "-20% CIS", "Total"]
    total_cols = 1 + len(summary_cols)
    available_width = 420 - 20
    cell_width = available_width / total_cols
    pdf.set_font("Arial", "B", 6)
    pdf.cell(cell_width, 6, "Employee", border=1, align="C")
    for col in summary_cols:
        pdf.cell(cell_width, 6, col, border=1, align="C")
    pdf.ln(6)
    pdf.set_font("Arial", "", 6)
    for employee, row in df_summary.iterrows():
        pdf.cell(cell_width, 6, str(employee), border=1, align="C")
        for col in summary_cols:
            val = row[col]
            cell_text = "" if pd.isna(val) or float(val) == 0 else f"{int(round(val))}"
            pdf.cell(cell_width, 6, cell_text, border=1, align="C")
        pdf.ln(6)
    totals = df_summary[summary_cols].sum()
    pdf.set_font("Arial", "B", 6)
    pdf.cell(cell_width, 6, "TOTAL", border=1, align="C")
    for col in summary_cols:
        cell_text = f"{int(round(totals[col]))}"
        if col == "-20% CIS":
            pdf.set_fill_color(255, 255, 0)
            pdf.cell(cell_width, 6, cell_text, border=1, align="C", fill=True)
            pdf.set_fill_color(255, 255, 255)
        else:
            pdf.cell(cell_width, 6, cell_text, border=1, align="C")
    pdf.ln(6)
    pdf.output(str(pdf_output_path))

def create_employee_pdf(employee_name: str, summary_df: pd.DataFrame, pdf_output_path: Path) -> None:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    available_width = 210 - 20
    col_width = available_width / 4
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, f"Employee Summary: {employee_name}", ln=True, align="C")
    pdf.ln(10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(col_width, 10, "Tax Period", border=1, align="C")
    pdf.cell(col_width, 10, "Gross", border=1, align="C")
    pdf.cell(col_width, 10, "CIS", border=1, align="C")
    pdf.cell(col_width, 10, "Total", border=1, ln=True, align="C")
    pdf.set_font("Arial", "", 12)
    for _, row in summary_df.iterrows():
        pdf.cell(col_width, 10, str(row["TaxPeriod"]), border=1, align="C")
        pdf.cell(col_width, 10, f"{int(round(row['Gross']))}", border=1, align="C")
        pdf.cell(col_width, 10, f"{int(round(row['CIS']))}", border=1, align="C")
        pdf.cell(col_width, 10, f"{int(round(row['Total']))}", border=1, ln=True, align="C")
    pdf.output(str(pdf_output_path))

def build_reports(df: pd.DataFrame) -> Dict[str, Any]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if df.empty:
        raise RuntimeError("No CIS transactions found.")
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
    detailed_dir.mkdir(parents=True, exist_ok=True)
    monthly_summary_dir.mkdir(parents=True, exist_ok=True)
    employee_output_dir.mkdir(parents=True, exist_ok=True)

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
        combined.to_csv(csv_output_path)
        create_monthly_pdf(period_label, combined, pdf_output_path)
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
    employee_summary = df.groupby(["To", "TaxPeriodStart", "TaxPeriod"], as_index=False)["Paid out"].sum()
    employee_summary.rename(columns={"Paid out": "Total"}, inplace=True)
    employee_summary["Gross"] = employee_summary["Total"] / 0.8
    employee_summary["CIS"] = employee_summary["Total"] * 0.25
    employee_summary.sort_values(by=["To", "TaxPeriodStart"], inplace=True)

    employee_artifacts: Dict[str, Dict[str, Any]] = {}
    for employee, group in employee_summary.groupby("To"):
        output_df = group[["TaxPeriod", "Gross", "CIS", "Total"]].copy()
        totals_row = {
            "TaxPeriod": "YEAR TOTAL",
            "Gross": output_df["Gross"].sum(),
            "CIS": output_df["CIS"].sum(),
            "Total": output_df["Total"].sum(),
        }
        output_df_with_total = pd.concat([output_df, pd.DataFrame([totals_row])], ignore_index=True)
        safe_emp_name = "".join(c for c in employee if c.isalnum() or c in " _-").strip()
        emp_csv_path = employee_output_dir / f"{safe_emp_name}.csv"
        emp_pdf_path = employee_output_dir / f"{safe_emp_name}.pdf"
        output_df_with_total.to_csv(emp_csv_path, index=False)
        create_employee_pdf(employee, output_df_with_total, emp_pdf_path)
        current_month_row = group.iloc[-1]
        employee_artifacts[employee] = {
            "pdf": emp_pdf_path,
            "csv": emp_csv_path,
            "current_month_gross": float(current_month_row["Gross"]),
            "current_month_cis": float(current_month_row["CIS"]),
            "ytd_gross": float(output_df["Gross"].sum()),
            "ytd_cis": float(output_df["CIS"].sum()),
        }

    employee_list_output = OUTPUT_DIR / "employee_list.csv"
    pd.DataFrame(sorted(df["To"].unique()), columns=["Employee"]).to_csv(employee_list_output, index=False)
    overall_total = float(df["Paid out"].sum())
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
    result = build_reports(df)
    print(result)
