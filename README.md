# HR Job Scraper v4

Python scraper for collecting recent HR job postings from major job portals for Delhi NCR, Noida, and Gurugram. The script searches multiple HR-related keywords, filters jobs to roughly the last 24 hours, deduplicates the results, and exports a formatted Excel file.

## What It Scrapes

The main scraper is [`hr_job_scraper_v4.py`](hr_job_scraper_v4.py). It currently supports:

| Portal | Method |
| --- | --- |
| LinkedIn | Selenium browser scraping with LinkedIn's 24-hour filter |
| Naukri | API request, Next.js data extraction, and HTML fallback |
| Indeed | Search pages plus embedded JSON extraction |
| Glassdoor | Direct search URL with JSON-LD and HTML fallback |
| Shine | HTML and JSON-LD extraction |
| Internshala | Internal AJAX endpoint and HTML fallback |

The default output file is:

```text
HR_Jobs_Last24h.xlsx
```

## Output Columns

The Excel report includes:

- Portal and source keyword
- Job title and company name
- Job location and posted date
- Salary, experience, and employment type when available
- Job URL
- Contact email and phone when detected in the job description
- Short job description summary
- UTC fetch timestamp

The Excel file is styled with filters, frozen headers, portal-based row colors, clickable job links, clickable email links, and clickable phone links.

## Requirements

- Python 3.10 or newer recommended
- Google Chrome installed
- Internet connection

Install dependencies:

```bash
pip install selenium webdriver-manager requests beautifulsoup4 pandas openpyxl tqdm fake-useragent undetected-chromedriver
```

`undetected-chromedriver` is strongly recommended because LinkedIn is more likely to block normal Selenium sessions.

## Run

Use the small entry point:

```bash
python main.py
```

Or run the scraper directly:

```bash
python hr_job_scraper_v4.py
```

On Windows, you can also double-click:

```text
run.bat
```

The batch file installs/verifies dependencies, runs `python main.py`, and saves the Excel file in this project folder.

## Daily Email Automation

This project includes a local daily runner:

```text
daily_run.py
```

It runs the scraper, waits for `HR_Jobs_Last24h.xlsx`, and emails the Excel report through Brevo SMTP.

### 1. Configure Brevo Email

Copy `.env.example` to `.env`:

```bash
copy .env.example .env
```

Then edit `.env`:

```env
BREVO_SMTP_HOST=smtp-relay.brevo.com
BREVO_SMTP_PORT=587
BREVO_SMTP_USER=your_brevo_smtp_login
BREVO_SMTP_KEY=your_brevo_smtp_key

MAIL_FROM=your_verified_sender@example.com
MAIL_FROM_NAME=HR Job Scraper
MAIL_TO=receiver@example.com
MAIL_SUBJECT=Daily HR Jobs Report
```

Use a sender email that is verified in Brevo. `MAIL_TO`, `MAIL_CC`, and `MAIL_BCC` support comma-separated email addresses.

### Multiple Job Reports

`daily_run.py` can run multiple profiles in one scheduled run. Each profile gets:

- Its own keyword list
- Its own Excel output file
- Its own email recipient list
- Its own email subject

Example `.env` setup:

```env
JOB_PROFILES=hr,frontend

PROFILE_HR_NAME=HR Jobs
PROFILE_HR_OUTPUT=HR_Jobs_Last24h.xlsx
PROFILE_HR_MAIL_TO=hr@example.com,manager@example.com
PROFILE_HR_SUBJECT=Daily HR Jobs Report
PROFILE_HR_KEYWORDS=HR Manager,HR Executive,HR Generalist,Recruiter,Talent Acquisition Specialist

PROFILE_FRONTEND_NAME=Frontend React Jobs
PROFILE_FRONTEND_OUTPUT=Frontend_Jobs_Last24h.xlsx
PROFILE_FRONTEND_MAIL_TO=frontend@example.com
PROFILE_FRONTEND_SUBJECT=Daily Frontend React Jobs Report
PROFILE_FRONTEND_KEYWORDS=Frontend Developer,React Developer,React JS Developer,Frontend Engineer,UI Developer,JavaScript Developer,Next.js Developer
```

To send only one report, set:

```env
JOB_PROFILES=hr
```

or:

```env
JOB_PROFILES=frontend
```

### 2. Test One Full Run

Run:

```bash
python daily_run.py
```

Or:

```bash
run_daily.bat
```

This will scrape jobs, generate the Excel file, send the email, and write logs to:

```text
daily_scraper.log
```

### 3. Schedule It Daily at 11 AM

Open PowerShell in this project folder and run:

```powershell
.\setup_daily_task.ps1
```

This creates or updates a Windows Task Scheduler task named:

```text
HR Job Scraper v4 Daily Email
```

To use another time:

```powershell
.\setup_daily_task.ps1 -RunTime "10:30"
```

The scheduled task runs `run_daily.bat`, which then runs `daily_run.py`.

### Automation Files

| File | Purpose |
| --- | --- |
| `send_report.py` | Sends the Excel report through Brevo SMTP |
| `daily_run.py` | Runs scraper, checks report, sends email, writes logs |
| `run_daily.bat` | Windows batch entry point for manual or scheduled runs |
| `setup_daily_task.ps1` | Creates the daily Windows scheduled task |
| `.env.example` | Template for Brevo and email settings |

Do not commit `.env`; it contains your SMTP credentials.

## Configuration

Most settings are near the top of [`hr_job_scraper_v4.py`](hr_job_scraper_v4.py):

```python
OUTPUT_FILE        = "HR_Jobs_Last24h.xlsx"
HEADLESS           = True
MAX_RESULTS_PER_KW = 20
MAX_JOB_AGE_HOURS  = 24
SEARCH_PAGES       = 2
DEBUG_MODE         = True
```

Portal switches:

```python
ENABLE_LINKEDIN    = True
ENABLE_NAUKRI      = True
ENABLE_INDEED      = True
ENABLE_GLASSDOOR   = True
ENABLE_SHINE       = True
ENABLE_INTERNSHALA = True
```

Search terms are stored in `HR_KEYWORDS`, and target locations are controlled by `LOCATION_QUERIES` and `TARGET_LOCATION_KEYWORDS`.

## Chrome Driver Note

LinkedIn uses `undetected-chromedriver` when available. In `build_driver()`, it is currently pinned like this:

```python
driver = uc.Chrome(options=opts, use_subprocess=True, version_main=145)
```

If your installed Chrome major version is different, update `version_main` to match your Chrome version. If `undetected-chromedriver` is unavailable, the script falls back to regular Selenium with `webdriver-manager`.

## Debugging

`DEBUG_MODE = True` prints detailed reasons for skipped jobs, such as:

- Posted date is older than 24 hours
- Posted date text could not be parsed
- Location does not match the Delhi NCR target list
- HTTP request failed or was blocked

Set it to `False` for quieter normal runs.

## Important Notes

Job sites change their HTML, APIs, and bot protection often. Some portals may return fewer results, block requests, or require selector/API updates over time.

The "last 24 hours" filter is best-effort. Some portals expose exact posted times, while others only provide text such as "today", "actively hiring", or no date at all. LinkedIn is handled by using its server-side 24-hour search filter and accepting cards with missing posted text.

Generated files such as `HR_Jobs_Last24h.xlsx`, zip archives, and `__pycache__` are runtime artifacts, not source files.

## Render Cron Job

For Render, create a **Cron Job**, not a Web Service.

Recommended Render settings:

```text
Build Command: pip install -r requirements.txt
Command: python daily_run.py
Schedule: 30 5 * * *
```

`30 5 * * *` runs daily at 05:30 UTC, which is 11:00 AM India time.

Set secrets and profile settings in Render Environment Variables. Do not upload `.env`.

Recommended Render-only setting:

```env
ENABLE_LINKEDIN=false
```

LinkedIn uses Selenium/Chrome and may need a Docker setup on Render. Disabling LinkedIn keeps the cron job simpler and lets the requests/API-based portals run first.
