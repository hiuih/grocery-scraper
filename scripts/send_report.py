#!/usr/bin/env python3
"""Emails the weekly scrape results (or a failure notice) via Gmail SMTP.

Usage: send_report.py <freshst_outcome> <saveon_outcome> <recipient_email>
  outcomes are GitHub Actions step outcomes: "success" or "failure"
  Requires GMAIL_APP_PASSWORD env var (an app password for SENDER, not the
  account's normal login password).
"""
import os
import smtplib
import sys
from datetime import datetime
from email.message import EmailMessage

HOME = os.path.expanduser("~")
SENDER = "sadrabajoghli2777727@gmail.com"
FAILURE_RECIPIENT = SENDER

FILES = {
    "Fresh St. Market": os.path.join(HOME, "Desktop", "Fresh_St_Market_Products.xlsx"),
    "Save-On-Foods": os.path.join(HOME, "Desktop", "Save_On_Foods_Products.xlsx"),
}


def send(to, subject, body, attachments=()):
    msg = EmailMessage()
    msg["From"] = SENDER
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    for path in attachments:
        with open(path, "rb") as f:
            data = f.read()
        msg.add_attachment(
            data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(path),
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER, os.environ["GMAIL_APP_PASSWORD"])
        smtp.send_message(msg)


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
        send(recipient, f"Grocery Price Data — {today}", body, present.values())

    if failed:
        send(
            FAILURE_RECIPIENT,
            "Weekly grocery scrape had a failure",
            f"{', '.join(failed)} failed during this week's scheduled run "
            f"({today}). Check the GitHub Actions log for details.",
        )

    if not present:
        sys.exit(1)


if __name__ == "__main__":
    main()
