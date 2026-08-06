#!/usr/bin/env python3
"""Emails the weekly scrape results (or a failure notice) via the Composio CLI.

Usage: send_report.py <freshst_outcome> <saveon_outcome> <recipient_email>
  outcomes are GitHub Actions step outcomes: "success" or "failure"
"""
import json
import os
import subprocess
import sys
from datetime import datetime

HOME = os.path.expanduser("~")
FAILURE_RECIPIENT = "sadrabajoghli2777727@gmail.com"

FILES = {
    "Fresh St. Market": os.path.join(HOME, "Desktop", "Fresh_St_Market_Products.xlsx"),
    "Save-On-Foods": os.path.join(HOME, "Desktop", "Save_On_Foods_Products.xlsx"),
}


def send(payload):
    subprocess.run(
        ["composio", "execute", "GMAIL_SEND_EMAIL", "-d", json.dumps(payload)],
        check=True,
    )


def main():
    freshst_ok = sys.argv[1] == "success"
    saveon_ok = sys.argv[2] == "success"
    recipient = sys.argv[3]

    present = {name: path for name, path in FILES.items() if os.path.exists(path)}
    failed = [
        name
        for name, ok in [("Fresh St. Market", freshst_ok), ("Save-On-Foods", saveon_ok)]
        if not ok
    ]
    today = datetime.now().strftime("%B %d, %Y")

    if present:
        body = f"Weekly grocery price scrape — {today}\n\nAttached:\n"
        body += "\n".join(f"- {name}" for name in present)
        if failed:
            body += f"\n\nNote: {', '.join(failed)} did not complete successfully this run."
        send({
            "recipient_email": recipient,
            "subject": f"Grocery Price Data — {today}",
            "body": body,
            "attachment": list(present.values()),
        })

    if failed:
        send({
            "recipient_email": FAILURE_RECIPIENT,
            "subject": "Weekly grocery scrape had a failure",
            "body": (
                f"{', '.join(failed)} failed during this week's scheduled run "
                f"({today}). Check the GitHub Actions log for details."
            ),
        })

    if not present:
        sys.exit(1)


if __name__ == "__main__":
    main()
