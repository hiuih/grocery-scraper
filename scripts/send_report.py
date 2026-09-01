#!/usr/bin/env python3
"""Emails one site's scrape result (or a failure notice) via Gmail SMTP,
as soon as that site's job finishes -- independent of the other site.

Usage: send_report.py <site_name> <outcome> <file_path> <recipient_email>
  outcome is a GitHub Actions job status ("success", "failure", "cancelled");
  anything other than exactly "success" is treated as a failure.
  Requires GMAIL_APP_PASSWORD env var (an app password for SENDER, not the
  account's normal login password).
"""
import os
import smtplib
import sys
from datetime import datetime
from email.message import EmailMessage

SENDER = "sadrabajoghli2777727@gmail.com"
FAILURE_RECIPIENT = SENDER


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
    site_name, outcome, file_path, recipient = sys.argv[1:5]
    ok = outcome == "success" and os.path.exists(file_path)
    today = datetime.now().strftime("%B %d, %Y")

    if ok:
        send(
            recipient,
            f"{site_name} Price Data — {today}",
            f"{site_name} price scrape — {today}\n\nAttached: {os.path.basename(file_path)}",
            [file_path],
        )
    else:
        send(
            FAILURE_RECIPIENT,
            f"{site_name} scrape failed",
            f"{site_name} did not complete successfully this run ({today}, "
            f"outcome: {outcome}). Check the GitHub Actions log for details.",
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
