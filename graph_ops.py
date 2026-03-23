import os
import base64
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
import requests
import pandas as pd

from dotenv import load_dotenv
load_dotenv()

def get_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value

GRAPH_TENANT_ID = get_env("GRAPH_TENANT_ID")
GRAPH_CLIENT_ID = get_env("GRAPH_CLIENT_ID")
GRAPH_CLIENT_SECRET = get_env("GRAPH_CLIENT_SECRET")

GRAPH_TOKEN_URL = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
MAILBOX_USER = get_env("MAILBOX_USER")

def get_graph_access_token() -> str:
    response = requests.post(
        GRAPH_TOKEN_URL,
        data={
            "client_id": GRAPH_CLIENT_ID,
            "client_secret": GRAPH_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["access_token"]

def graph_headers(token: str, json_content: bool = True) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    if json_content:
        headers["Content-Type"] = "application/json"
    return headers

def get_cis_period_end_for_run(run_date: datetime) -> datetime:
    reporting_start = get_reporting_period_start_for_run(run_date)
    reporting_end = reporting_start.replace(day=6) + pd.DateOffset(months=1) - pd.DateOffset(days=1)
    return reporting_end.to_pydatetime()

def get_cis_tax_year(dt: datetime) -> str:
    if (dt.month, dt.day) >= (4, 6):
        start_year = dt.year
        end_year = dt.year + 1
    else:
        start_year = dt.year - 1
        end_year = dt.year
    return f"{start_year}-{end_year}"

def get_month_key(period_end: datetime) -> str:
    return period_end.strftime("%Y-%m")

def get_month_label(period_end: datetime) -> str:
    return period_end.strftime("%B %Y")

def safe_name(name: str) -> str:
    return "".join(c for c in name if c.isalnum() or c in " _-").strip()

def build_remote_paths(run_date: datetime) -> Dict[str, str]:
    reporting_start = get_reporting_period_start_for_run(run_date)
    period_end = get_cis_period_end_for_run(run_date)

    tax_year = get_cis_tax_year(period_end)
    month_key = reporting_start.strftime("%Y-%m")

    employer_month_folder = f"CIS Reports/employer/{tax_year}/{month_key}"
    employees_root = "CIS Reports/employees"

    return {
        "tax_year": tax_year,
        "month_key": month_key,
        "employer_month_folder": employer_month_folder,
        "employees_root": employees_root,
    }

def get_drive_item_by_path(token: str, remote_path: str) -> Optional[dict]:
    url = f"{GRAPH_BASE_URL}/users/{MAILBOX_USER}/drive/root:/{remote_path}"
    response = requests.get(url, headers=graph_headers(token, json_content=False), timeout=30)
    if response.status_code == 200:
        return response.json()
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return None

def ensure_folder(token: str, parent_path: str, folder_name: str) -> dict:
    existing = get_drive_item_by_path(token, f"{parent_path}/{folder_name}")
    if existing:
        return existing
    url = f"{GRAPH_BASE_URL}/users/{MAILBOX_USER}/drive/root:/{parent_path}:/children"
    payload = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    response = requests.post(url, headers=graph_headers(token), json=payload, timeout=30)
    if response.status_code in (200, 201):
        return response.json()
    response.raise_for_status()
    return response.json()

def ensure_root_folder(token: str, folder_name: str) -> dict:
    existing = get_drive_item_by_path(token, folder_name)
    if existing:
        return existing
    url = f"{GRAPH_BASE_URL}/users/{MAILBOX_USER}/drive/root/children"
    payload = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    response = requests.post(url, headers=graph_headers(token), json=payload, timeout=30)
    if response.status_code in (200, 201):
        return response.json()
    response.raise_for_status()
    return response.json()

def ensure_nested_folder(token: str, full_path: str) -> dict:
    parts = [p for p in full_path.split("/") if p.strip()]
    if not parts:
        raise ValueError("full_path must not be empty")
    current_path = parts[0]
    ensure_root_folder(token, current_path)
    for part in parts[1:]:
        ensure_folder(token, current_path, part)
        current_path = f"{current_path}/{part}"
    final_item = get_drive_item_by_path(token, current_path)
    if not final_item:
        raise RuntimeError(f"Failed to ensure folder path: {full_path}")
    return final_item

def upload_file_to_onedrive(token: str, local_path: Path, remote_path: str) -> dict:
    url = f"{GRAPH_BASE_URL}/users/{MAILBOX_USER}/drive/root:/{remote_path}:/content"
    with open(local_path, "rb") as f:
        response = requests.put(
            url,
            headers={"Authorization": f"Bearer {token}"},
            data=f.read(),
            timeout=60,
        )
    if response.status_code not in (200, 201):
        response.raise_for_status()
    return response.json()

def create_view_link(token: str, remote_path: str) -> str:
    url = f"{GRAPH_BASE_URL}/users/{MAILBOX_USER}/drive/root:/{remote_path}:/createLink"
    payload = {"type": "view", "scope": "organization"}
    response = requests.post(url, headers=graph_headers(token), json=payload, timeout=30)
    if response.status_code not in (200, 201):
        response.raise_for_status()
    return response.json()["link"]["webUrl"]

def send_email(token: str, to_addresses: List[str], subject: str, body_text: str, attachments: Optional[List[str]] = None) -> None:
    url = f"{GRAPH_BASE_URL}/users/{MAILBOX_USER}/sendMail"
    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body_text},
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_addresses],
    }
    if attachments:
        graph_attachments = []
        for file_path in attachments:
            with open(file_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")
            graph_attachments.append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": os.path.basename(file_path),
                "contentBytes": encoded,
            })
        message["attachments"] = graph_attachments
    payload = {"message": message, "saveToSentItems": True}
    response = requests.post(url, headers=graph_headers(token), json=payload, timeout=60)
    response.raise_for_status()

def get_reporting_period_start_for_run(run_date: datetime) -> datetime:
    """
    If the job runs on the 15th, report the CIS month that started on the 6th
    of the PREVIOUS calendar month.

    Example:
    2026-03-15 -> 2026-02-06
    """
    if run_date.month == 1:
        return datetime(run_date.year - 1, 12, 6)
    return datetime(run_date.year, run_date.month - 1, 6)


def get_reporting_period_label_for_run(run_date: datetime) -> str:
    return get_reporting_period_start_for_run(run_date).strftime("%B %Y")