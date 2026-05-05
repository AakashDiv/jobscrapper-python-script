"""
Daily automation entry point.

Runs one or more job profiles, creates one Excel file per profile, and emails
each report to that profile's recipients through Brevo SMTP.
"""

from __future__ import annotations

import logging
import os
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from hr_job_scraper_v4 import HR_KEYWORDS, run_scrape
from send_report import load_env_file, send_report


PROJECT_DIR = Path(__file__).resolve().parent
LOG_FILE = PROJECT_DIR / "daily_scraper.log"

FRONTEND_KEYWORDS = [
    "Frontend Developer",
    "React Developer",
    "React JS Developer",
    "Frontend Engineer",
    "UI Developer",
    "JavaScript Developer",
    "Next.js Developer",
    "Web Developer React",
]


@dataclass
class JobProfile:
    key: str
    name: str
    keywords: list[str]
    output_file: Path
    mail_to: str
    subject: str


def configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)-7s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        force=True,
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def split_csv(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def env_name(profile_key: str, suffix: str) -> str:
    return f"PROFILE_{profile_key.upper()}_{suffix}"


def profile_keywords(profile_key: str, default_keywords: list[str]) -> list[str]:
    configured = split_csv(os.getenv(env_name(profile_key, "KEYWORDS")))
    return configured or default_keywords


def build_profiles() -> list[JobProfile]:
    profile_keys = split_csv(os.getenv("JOB_PROFILES")) or ["hr"]
    profiles: list[JobProfile] = []

    for key in profile_keys:
        normalized = key.strip().lower()
        if not normalized:
            continue

        if normalized == "frontend":
            default_name = "Frontend React Jobs"
            default_keywords = FRONTEND_KEYWORDS
            default_output = "Frontend_Jobs_Last24h.xlsx"
            default_subject = "Daily Frontend React Jobs Report"
        else:
            default_name = "HR Jobs"
            default_keywords = HR_KEYWORDS
            default_output = "HR_Jobs_Last24h.xlsx"
            default_subject = "Daily HR Jobs Report"

        name = os.getenv(env_name(normalized, "NAME"), default_name).strip()
        output = os.getenv(env_name(normalized, "OUTPUT"), default_output).strip()
        mail_to = os.getenv(env_name(normalized, "MAIL_TO"), os.getenv("MAIL_TO", "")).strip()
        subject = os.getenv(env_name(normalized, "SUBJECT"), default_subject).strip()

        if not mail_to:
            raise RuntimeError(f"Missing recipients for profile '{normalized}'. Set {env_name(normalized, 'MAIL_TO')}.")

        profiles.append(
            JobProfile(
                key=normalized,
                name=name,
                keywords=profile_keywords(normalized, default_keywords),
                output_file=PROJECT_DIR / output,
                mail_to=mail_to,
                subject=subject,
            )
        )

    if not profiles:
        raise RuntimeError("No job profiles configured.")
    return profiles


def maybe_send_failure_email(profile: JobProfile, error_text: str) -> None:
    if os.getenv("SEND_FAILURE_EMAIL", "false").strip().lower() not in {"1", "true", "yes"}:
        return

    try:
        # Failure email uses the last available report for this profile, if one exists.
        if profile.output_file.exists():
            send_report(
                profile.output_file,
                mail_to=profile.mail_to,
                subject=f"{profile.name} Scraper Failed",
                body=(
                    "Hi,\n\nThe daily job scraper failed. "
                    "The last available report is attached.\n\n"
                    f"Profile: {profile.name}\n\n"
                    f"Error:\n{error_text}\n"
                ),
            )
    except Exception:
        logging.exception("Failed to send failure email for profile: %s", profile.name)


def run_profile(profile: JobProfile) -> None:
    logging.info("Starting profile: %s", profile.name)
    logging.info("Keywords: %s", ", ".join(profile.keywords))

    run_scrape(
        keywords=profile.keywords,
        output_file=str(profile.output_file),
        profile_name=profile.name,
    )

    if not profile.output_file.exists():
        raise FileNotFoundError(f"Expected report was not generated: {profile.output_file}")

    send_report(
        profile.output_file,
        mail_to=profile.mail_to,
        subject=profile.subject,
        body=(
            "Hi,\n\n"
            f"Please find attached the latest {profile.name} Excel report.\n\n"
            "Regards,\nHR Job Scraper"
        ),
    )
    logging.info("Report emailed successfully for %s: %s", profile.name, profile.output_file)


def main() -> int:
    os.chdir(PROJECT_DIR)
    load_env_file()
    configure_logging()

    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logging.info("Daily scraper started at %s", started_at)

    failures = 0
    try:
        profiles = build_profiles()
        for profile in profiles:
            try:
                run_profile(profile)
            except Exception as exc:
                failures += 1
                error_text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
                logging.error("Profile failed: %s\n%s", profile.name, error_text)
                maybe_send_failure_email(profile, error_text)

        return 1 if failures else 0

    finally:
        logging.info("Daily scraper finished with %d failure(s)", failures)


if __name__ == "__main__":
    raise SystemExit(main())
