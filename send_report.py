"""
Send the generated Excel report through Brevo SMTP.

Configuration is read from environment variables and an optional .env file in
this project folder. Do not commit real SMTP credentials.
"""

from __future__ import annotations

import mimetypes
import os
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from typing import Iterable


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_REPORT = PROJECT_DIR / "HR_Jobs_Last24h.xlsx"


def load_env_file(path: Path = PROJECT_DIR / ".env") -> None:
    """Load simple KEY=VALUE lines into os.environ if they are not already set."""
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def _split_recipients(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def send_report(
    report_path: str | Path = DEFAULT_REPORT,
    *,
    subject: str | None = None,
    body: str | None = None,
    mail_to: str | None = None,
    mail_cc: str | None = None,
    mail_bcc: str | None = None,
    sender_email: str | None = None,
    sender_name: str | None = None,
) -> None:
    load_env_file()

    report = Path(report_path).resolve()
    if not report.exists():
        raise FileNotFoundError(f"Report file not found: {report}")

    smtp_host = os.getenv("BREVO_SMTP_HOST", "smtp-relay.brevo.com")
    smtp_port = int(os.getenv("BREVO_SMTP_PORT", "587"))
    smtp_user = _required_env("BREVO_SMTP_USER")
    smtp_key = _required_env("BREVO_SMTP_KEY")

    sender_email = sender_email or _required_env("MAIL_FROM")
    sender_name = sender_name or os.getenv("MAIL_FROM_NAME", "HR Job Scraper").strip()
    to_recipients = _split_recipients(mail_to or _required_env("MAIL_TO"))
    cc_recipients = _split_recipients(mail_cc if mail_cc is not None else os.getenv("MAIL_CC"))
    bcc_recipients = _split_recipients(mail_bcc if mail_bcc is not None else os.getenv("MAIL_BCC"))

    msg = EmailMessage()
    msg["From"] = formataddr((sender_name, sender_email))
    msg["To"] = ", ".join(to_recipients)
    if cc_recipients:
        msg["Cc"] = ", ".join(cc_recipients)
    msg["Subject"] = subject or os.getenv("MAIL_SUBJECT", "Daily HR Jobs Report")
    msg.set_content(
        body
        or "Hi,\n\nPlease find attached the latest HR jobs Excel report.\n\nRegards,\nHR Job Scraper"
    )

    content_type, _ = mimetypes.guess_type(report.name)
    maintype, subtype = (content_type or "application/octet-stream").split("/", 1)
    msg.add_attachment(
        report.read_bytes(),
        maintype=maintype,
        subtype=subtype,
        filename=report.name,
    )

    all_recipients: Iterable[str] = to_recipients + cc_recipients + bcc_recipients
    with smtplib.SMTP(smtp_host, smtp_port, timeout=60) as server:
        server.starttls()
        server.login(smtp_user, smtp_key)
        server.send_message(msg, to_addrs=list(all_recipients))


if __name__ == "__main__":
    send_report()
