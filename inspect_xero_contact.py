import argparse
import json
from typing import Any, Dict, List, Optional

import requests

from xero_reports import build_headers, get_access_token, normalize_name


CONTACTS_URL = "https://api.xero.com/api.xro/2.0/Contacts"


def get_contact_by_id(contact_id: str) -> Dict[str, Any]:
    token = get_access_token()
    response = requests.get(
        f"{CONTACTS_URL}/{contact_id}",
        headers=build_headers(token),
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def get_contacts_page(access_token: str, page: int = 1) -> List[Dict[str, Any]]:
    response = requests.get(
        CONTACTS_URL,
        headers=build_headers(access_token),
        params={"page": page},
        timeout=30,
    )
    response.raise_for_status()
    return response.json().get("Contacts", [])


def find_contacts_by_name(name: str, limit: int) -> List[Dict[str, Any]]:
    wanted = normalize_name(name)
    matches: List[Dict[str, Any]] = []
    token = get_access_token()
    page = 1

    while True:
        rows = get_contacts_page(token, page)
        if not rows:
            break

        for contact in rows:
            contact_name = normalize_name(contact.get("Name") or "")
            if wanted in contact_name:
                matches.append(contact)
                if len(matches) >= limit:
                    return matches

        page += 1

    return matches


def print_payload(payload: Any, output_path: Optional[str]) -> None:
    text = json.dumps(payload, indent=2, sort_keys=True)
    if output_path:
        with open(output_path, "w", encoding="utf-8") as file:
            file.write(text)
            file.write("\n")
        print(f"Wrote contact payload to {output_path}")
        return

    print(text)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Pull a Xero contact payload for inspection."
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--name", help="Contact name, or part of a contact name, to search for.")
    group.add_argument("--contact-id", help="Exact Xero ContactID to fetch.")
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        help="Maximum name-search matches to return. Defaults to 5.",
    )
    parser.add_argument(
        "--output",
        help="Optional file path to write the JSON payload instead of printing it.",
    )

    args = parser.parse_args()

    if args.contact_id:
        payload = get_contact_by_id(args.contact_id)
    else:
        payload = {"Contacts": find_contacts_by_name(args.name, args.limit)}

    print_payload(payload, args.output)


if __name__ == "__main__":
    main()
