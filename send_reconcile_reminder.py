from graph_ops import get_graph_access_token, send_email
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
    token = get_graph_access_token()
    send_email(
        token=token,
        to_addresses=[EMPLOYER_EMAIL],
        subject="CIS Reconciliation Reminder",
        body_text="Please reconcile CIS transactions in Xero before the report runs on the 15th.",
        attachments=None,
    )
    print("Reminder sent.")

if __name__ == "__main__":
    main()
