# Lemlist Engagement Reports

A Python tool that turns Lemlist campaign activity into ready-to-use weekly Excel reports, showing which people and which companies are actually engaging with outreach emails.

## The problem

Lemlist shows opens and clicks one event at a time. Our team needed a quick weekly view of **who** was engaging, **how often**, and **which companies** were warming up, without manual exports. Two things made the raw data misleading:

- **Bot activity.** Email security scanners "open" and "click" links automatically, inflating engagement.
- **Missing company names.** Leads uploaded straight into Lemlist often had no company field filled in, so many showed up as "Unknown Company."

## What it does

- Pulls opens and clicks for the last 7 days from the Lemlist API, for every running campaign
- Lets you **run all campaigns or pick specific ones** from a checkbox window
- **Filters out bot opens and clicks** using Lemlist's bot detection flag
- **Fills in missing company names** by checking several field variations, then the lead's full Lemlist record, then the email domain (skipping Gmail, Outlook, etc.), and records which source was used
- Builds one styled Excel workbook per campaign with 4 tabs:
  - **Opens – Individuals:** one row per person, with engagement count (color-coded green / yellow / red)
  - **Opens – Top Companies:** companies ranked by how many people engaged
  - **Clicks – Individuals** and **Clicks – Top Companies:** the same views for clicks
- Handles API pagination, rate limits and time zones, and stops paging early once it passes the 7-day window

<!-- Add a screenshot of a report here (use fake or blurred data):
![Example report](screenshot.png) -->

## Impact

<!-- Replace with your own numbers, e.g.: -->
_Replaced a manual weekly export process about 2 hours per week and gave the sales team a ranked list of engaged leads and accounts to follow up with._

## How to run it

1. Install Python 3, then install the libraries:
   ```
   pip install pandas openpyxl
   ```
2. Copy `config.example.txt` to `config.txt` and paste your Lemlist API key into it (Lemlist → Settings → Integrations → API).
3. Double-click `Run Lemlist Reports.bat` (or run `python lemlist_report.py`).
4. Choose to run all campaigns or tick the ones you want. Reports are saved to a dated folder on your Desktop.

To run without the prompt (e.g. from Windows Task Scheduler), use `python lemlist_report.py --all`.

## Settings

Options at the top of `lemlist_report.py`:

| Setting | What it does |
| --- | --- |
| `DAYS_BACK` | How many days of activity to include (default 7) |
| `ACTIVE_STATUSES` | Which campaign statuses count as active |
| `EXCLUDE_BOTS` | Drop bot-flagged opens and clicks |
| `LOOKUP_LEAD_RECORDS` | Look up a lead's full record when the company is missing |
| `USE_EMAIL_DOMAIN_FALLBACK` | Use the email domain as a last resort for the company |
| `DEBUG_SHOW_FIELDS` | Print the raw fields Lemlist returns, for troubleshooting |

## Tech

Python · pandas · openpyxl · Lemlist REST API · tkinter

`config.txt` and generated reports are excluded via `.gitignore`, so API keys and lead data never end up in the repository.
