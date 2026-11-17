import base64
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------------
CONFIG_PATH = "config.txt"         # put your Lemlist API key in config.txt next to this script
DAYS_BACK = 7
PAGE_SIZE = 100                    # Lemlist returns at most ~100 rows per request
ACTIVE_STATUSES = {"running"}      # campaign statuses treated as ACTIVE
MAX_PAGES = 500                    # safety stop so a bad response can't loop forever

# Ask which campaigns to run each time. Set to False (or run the script with
# --all) to skip the prompt, e.g. when it runs from Task Scheduler.
ASK_WHICH_CAMPAIGNS = True

# Drop opens/clicks that Lemlist flags as bots (link scanners, security filters).
EXCLUDE_BOTS = True

# Company name lookup, tried in this order until one works:
#   1. company fields on the activity itself (Company, companyName, associated company...)
#   2. the lead's full record in Lemlist (one extra API call per person still missing)
#   3. the person's email domain, e.g. jane@evanslangford.com -> evanslangford.com
LOOKUP_LEAD_RECORDS = True
USE_EMAIL_DOMAIN_FALLBACK = True

# Set True once to print which fields Lemlist actually sends, handy for troubleshooting.
DEBUG_SHOW_FIELDS = False

# Field names checked first (compared ignoring case, spaces, _ and -).
COMPANY_FIELD_PRIORITY = [
    "leadcompanyname", "companyname", "company", "associatedcompany",
    "associatedcompanyname", "accountname", "organization", "organisation",
    "organizationname", "employer", "businessname", "business", "firm",
]
# Any other field containing "company"/"organi" counts too, unless it has one of these.
COMPANY_FIELD_EXCLUDE = [
    "domain", "description", "employee", "size", "growth", "url", "website",
    "linkedin", "industry", "id", "location", "city", "country", "revenue",
    "phone", "type", "founded", "logo", "headcount", "address",
]
FREE_EMAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "yahoo.ca", "hotmail.com",
    "hotmail.ca", "outlook.com", "live.com", "live.ca", "msn.com", "aol.com",
    "icloud.com", "me.com", "mac.com", "protonmail.com", "proton.me",
    "gmx.com", "mail.com", "yandex.com", "rogers.com", "bell.net",
    "shaw.ca", "telus.net", "sympatico.ca", "comcast.net", "att.net",
}


# ----------------------------------------------------------------------------
# API HELPERS
# ----------------------------------------------------------------------------
def get_lemlist_api_key():
    """Look for config.txt next to this script first, then at CONFIG_PATH."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [os.path.join(script_dir, "config.txt"), CONFIG_PATH]

    for path in candidates:
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    key = f.read().strip()
                if key:
                    return key
        except Exception as e:
            print(f"Could not read {path}: {e}")

    print("Looked for config.txt in:")
    for path in candidates:
        print(f"   {path}")
    return None


def fetch_lemlist_data(api_key, endpoint_path, retries=3, quiet=False):
    """GET from the Lemlist API. Prints real errors instead of hiding them."""
    url = f"https://api.lemlist.com/api/{endpoint_path}"
    token = base64.b64encode(f":{api_key}".encode("utf-8")).decode("utf-8")

    for attempt in range(1, retries + 1):
        req = urllib.request.Request(url, headers={"Authorization": f"Basic {token}"})
        try:
            with urllib.request.urlopen(req, timeout=30) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")[:200]
            if not quiet or e.code == 429:
                print(f"[API Error] {e.code} on {endpoint_path}: {body}")
            if e.code == 429 and attempt < retries:  # rate limited, wait and retry
                time.sleep(2 * attempt)
                continue
            return []
        except Exception as e:
            if not quiet:
                print(f"[API Error] {endpoint_path}: {e}")
            if attempt < retries:
                time.sleep(1)
                continue
            return []
    return []


def unwrap_list(payload, keys=("campaigns", "activities", "data", "results")):
    """The API sometimes returns a bare list and sometimes {"campaigns": [...]}."""
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for k in keys:
            if isinstance(payload.get(k), list):
                return payload[k]
    return []


def fetch_all_campaigns(api_key):
    campaigns, seen_ids, offset = [], set(), 0
    for _ in range(MAX_PAGES):
        payload = fetch_lemlist_data(
            api_key, f"campaigns?version=v2&limit={PAGE_SIZE}&offset={offset}"
        )
        batch = unwrap_list(payload)
        new_rows = [c for c in batch if c.get("_id") not in seen_ids]
        if not new_rows:
            break
        for c in new_rows:
            seen_ids.add(c.get("_id"))
        campaigns.extend(new_rows)
        if len(batch) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
    return campaigns


def fetch_all_activities(api_key, camp_id, activity_type, stop_before=None):
    """Page through activities of one type for one campaign.

    If stop_before (a UTC Timestamp) is given and the API returns newest-first,
    paging stops as soon as a page reaches events older than that time, so we
    don't download a campaign's entire history just to keep the last 7 days.
    """
    rows, seen_ids, offset = [], set(), 0
    newest_first = None  # worked out from the first page
    for _ in range(MAX_PAGES):
        payload = fetch_lemlist_data(
            api_key,
            f"activities?version=v2&campaignId={camp_id}&type={activity_type}"
            f"&limit={PAGE_SIZE}&offset={offset}",
        )
        batch = unwrap_list(payload)
        new_rows = [a for a in batch if a.get("_id") not in seen_ids]
        if not new_rows:
            break
        for a in new_rows:
            seen_ids.add(a.get("_id"))
        rows.extend(new_rows)
        print(f"   ...{activity_type}: {len(rows)} fetched so far", flush=True)
        if len(batch) < PAGE_SIZE:
            break

        if stop_before is not None:
            times = pd.to_datetime(
                [a.get("createdAt") for a in batch], errors="coerce", utc=True
            )
            times = times.dropna()
            if len(times) > 1:
                if newest_first is None:
                    newest_first = times[0] >= times[-1]
                if newest_first and times.min() < stop_before:
                    break  # everything from here on is older than our window

        offset += PAGE_SIZE
        time.sleep(0.2)  # be gentle with rate limits
    return rows


# ----------------------------------------------------------------------------
# CAMPAIGN SELECTION
# ----------------------------------------------------------------------------
def ask_run_all(campaigns):
    """Ask in the console whether to run every active campaign."""
    while True:
        ans = input(
            f"Run reports for ALL {len(campaigns)} active campaign(s)? (Y/n): "
        ).strip().lower()
        if ans in ("", "y", "yes"):
            return True
        if ans in ("n", "no"):
            return False
        print("  Please type y or n.")


def choose_campaigns_gui(campaigns):
    """Pop up a window with a checkbox per campaign. Returns the checked ones."""
    import tkinter as tk

    result = {"selected": []}
    root = tk.Tk()
    root.title("Lemlist Reports - Choose Campaigns")
    root.attributes("-topmost", True)  # don't hide behind the console window

    tk.Label(
        root, text="Tick the campaigns you want reports for:",
        font=("Segoe UI", 11, "bold"),
    ).pack(padx=12, pady=(12, 6), anchor="w")

    # Scrollable list, so it still works with lots of campaigns
    container = tk.Frame(root)
    container.pack(fill="both", expand=True, padx=12)
    canvas = tk.Canvas(
        container, width=480, height=min(28 * len(campaigns), 420),
        highlightthickness=0,
    )
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    canvas.bind_all(
        "<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units")
    )

    check_vars = []
    for c in campaigns:
        var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            inner, text=c.get("name", "Unknown Campaign"), variable=var,
            anchor="w", font=("Segoe UI", 10),
        ).pack(fill="x", anchor="w")
        check_vars.append(var)

    def set_all(value):
        for v in check_vars:
            v.set(value)

    def run_selected():
        result["selected"] = [c for c, v in zip(campaigns, check_vars) if v.get()]
        root.destroy()

    def cancel():
        result["selected"] = []
        root.destroy()

    buttons = tk.Frame(root)
    buttons.pack(fill="x", padx=12, pady=12)
    tk.Button(buttons, text="Select All", command=lambda: set_all(True)).pack(side="left")
    tk.Button(buttons, text="Clear All", command=lambda: set_all(False)).pack(side="left", padx=6)
    tk.Button(buttons, text="Cancel", command=cancel).pack(side="right")
    tk.Button(
        buttons, text="Run Selected", command=run_selected, font=("Segoe UI", 10, "bold")
    ).pack(side="right", padx=6)

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.after(100, root.focus_force)
    root.mainloop()
    return result["selected"]


def choose_campaigns_console(campaigns):
    """Text fallback: type numbers like 1,3,5 or ranges like 2-4."""
    print("\nActive campaigns:")
    for i, c in enumerate(campaigns, 1):
        print(f"  {i}. {c.get('name', 'Unknown Campaign')}")

    while True:
        raw = input(
            "Enter numbers to run (e.g. 1,3,5 or 2-4), or press Enter to cancel: "
        ).strip()
        if not raw:
            return []
        picked = set()
        try:
            for part in raw.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    start, end = part.split("-", 1)
                    picked.update(range(int(start), int(end) + 1))
                else:
                    picked.add(int(part))
        except ValueError:
            print("  Couldn't read that. Use numbers like 1,3,5 or 2-4.")
            continue
        bad = sorted(n for n in picked if n < 1 or n > len(campaigns))
        if bad:
            print(f"  Out of range: {bad}. Pick between 1 and {len(campaigns)}.")
            continue
        return [campaigns[n - 1] for n in sorted(picked)]


def select_campaigns(campaigns):
    """Run all, or let the user pick. Uses a checkbox window, falls back to text."""
    if not ASK_WHICH_CAMPAIGNS or "--all" in sys.argv:
        return campaigns
    if ask_run_all(campaigns):
        return campaigns
    try:
        return choose_campaigns_gui(campaigns)
    except Exception as e:
        print(f"(Checkbox window unavailable: {e}. Using text prompt instead.)")
        return choose_campaigns_console(campaigns)


# ----------------------------------------------------------------------------
# COMPANY NAME LOOKUP
# ----------------------------------------------------------------------------
def _norm_key(key):
    return "".join(ch for ch in str(key).lower() if ch.isalnum())


def _clean_company_value(value):
    """Return a usable company name from a field value, or None."""
    if isinstance(value, dict):  # e.g. an associated company object
        for k in ("name", "companyName", "label"):
            if k in value:
                return _clean_company_value(value[k])
        return None
    if not isinstance(value, str):
        return None
    value = value.strip()
    if not value or value.lower() in {"unknown", "unknown company", "n/a", "na", "none", "null", "-"}:
        return None
    return value


def resolve_company(record):
    """Find a company name in one dict of fields, trying many name variations."""
    if not isinstance(record, dict):
        return None
    by_norm = {}
    for k, v in record.items():
        by_norm.setdefault(_norm_key(k), v)

    for key in COMPANY_FIELD_PRIORITY:
        if key in by_norm:
            name = _clean_company_value(by_norm[key])
            if name:
                return name

    for key, v in by_norm.items():
        if ("company" in key or "organi" in key) and not any(
            bad in key for bad in COMPANY_FIELD_EXCLUDE
        ):
            name = _clean_company_value(v)
            if name:
                return name
    return None


def _walk_dicts(obj):
    """Yield every dict inside a nested API response (lead variables are often nested)."""
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from _walk_dicts(v)
    elif isinstance(obj, list):
        for item in obj:
            yield from _walk_dicts(item)


def lookup_company_from_lead(api_key, email, lead_cache):
    """Ask Lemlist for the lead's full record and look for a company there."""
    if email in lead_cache:
        return lead_cache[email]
    payload = fetch_lemlist_data(
        api_key, f"leads/{urllib.parse.quote(email)}?version=v2", retries=2, quiet=True
    )
    if DEBUG_SHOW_FIELDS and not lead_cache:
        print(f"   [Debug] lead record for {email}: {json.dumps(payload)[:1500]}")
    name = None
    for d in _walk_dicts(payload):
        name = resolve_company(d)
        if name:
            break
    lead_cache[email] = name
    time.sleep(0.15)  # be gentle with rate limits
    return name


def company_from_email_domain(email):
    domain = str(email).split("@")[-1].strip().lower()
    if not domain or "." not in domain or domain in FREE_EMAIL_DOMAINS:
        return None
    return domain


def add_company_names(df, api_key, lead_cache):
    """Fill leadCompanyName for every person, and note where it came from."""
    if df.empty:
        return df
    companies, sources = {}, {}
    for email, group in df.groupby("leadEmail"):
        name, source = None, None
        for rec in group.to_dict("records"):
            name = resolve_company(rec)
            if name:
                source = "Lemlist activity"
                break
        if not name and LOOKUP_LEAD_RECORDS:
            name = lookup_company_from_lead(api_key, email, lead_cache)
            if name:
                source = "Lemlist lead record"
        if not name and USE_EMAIL_DOMAIN_FALLBACK:
            name = company_from_email_domain(email)
            if name:
                source = "Email domain"
        if not name:
            name, source = "Unknown Company", "Not found"
        companies[email], sources[email] = name, source

    df = df.copy()
    df["leadCompanyName"] = df["leadEmail"].map(companies)
    df["Company Source"] = df["leadEmail"].map(sources)
    counts = pd.Series(sources).value_counts().to_dict()
    print(f"   company names: {counts}")
    return df


# ----------------------------------------------------------------------------
# DATA PREP
# ----------------------------------------------------------------------------
def to_clean_df(rows, camp_name, local_tz, cutoff_date):
    """Raw activity rows -> deduped DataFrame limited to the last N calendar days."""
    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)

    if "_id" in df.columns:
        df = df.drop_duplicates(subset=["_id"])

    if "leadEmail" in df.columns:
        df = df.dropna(subset=["leadEmail"])
    else:
        return pd.DataFrame()

    if EXCLUDE_BOTS and "bot" in df.columns:
        is_bot = df["bot"].astype(str).str.strip().str.lower().isin(["true", "1", "yes"])
        if is_bot.any():
            print(f"   removed {int(is_bot.sum())} bot events")
        df = df[~is_bot]

    df["campaignName"] = camp_name

    df["createdAt"] = pd.to_datetime(df["createdAt"], errors="coerce", utc=True)
    df = df.dropna(subset=["createdAt"])
    df["createdAt"] = df["createdAt"].dt.tz_convert(local_tz)
    df = df[df["createdAt"].dt.date >= cutoff_date].copy()
    return df


def build_tab_data(df, engagement_label):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df = df.copy()
    if "leadCompanyName" in df.columns:
        df["leadCompanyName"] = df["leadCompanyName"].fillna("Unknown Company")
    else:
        df["leadCompanyName"] = "Unknown Company"

    # TAB 1: one row per unique person, with how many times they did this action
    engagement_counts = (
        df.groupby("leadEmail").size().reset_index(name="Total Engagements")
    )
    # keep each person's most recent event for the detail columns
    clean_df = (
        df.sort_values("createdAt", ascending=False)
        .drop_duplicates(subset=["leadEmail"])
        .copy()
    )
    clean_df = pd.merge(clean_df, engagement_counts, on="leadEmail", how="left")

    columns_to_keep = [
        "Total Engagements", "leadFirstName", "leadLastName", "leadCompanyName",
        "Company Source", "leadEmail", "leadPhone", "campaignName", "sequenceStep",
        "subject", "sendUserName", "createdAt",
    ]
    available = [c for c in columns_to_keep if c in clean_df.columns]
    tab1_df = clean_df[available].copy()

    # Excel can't store timezone-aware datetimes
    for col in tab1_df.select_dtypes(include=["datetimetz"]).columns:
        tab1_df[col] = tab1_df[col].dt.tz_localize(None)

    tab1_df = tab1_df.sort_values(
        by=["Total Engagements", "leadCompanyName"], ascending=[False, True]
    )

    # TAB 2: company leaderboard
    company_df = (
        df.groupby("leadCompanyName")
        .agg(
            People_Interacted=("leadEmail", "nunique"),
            Total_Interactions=("leadEmail", "count"),
        )
        .reset_index()
    )
    company_df = company_df.sort_values(
        by=["People_Interacted", "Total_Interactions"], ascending=[False, False]
    )
    company_df = company_df.rename(
        columns={
            "leadCompanyName": "Company Name",
            "People_Interacted": "Total People Interacted",
            "Total_Interactions": f"Total {engagement_label}",
        }
    )
    return tab1_df, company_df


# ----------------------------------------------------------------------------
# EXCEL STYLING
# ----------------------------------------------------------------------------
def style_workbook(writer):
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    side = Side(style="thin", color="D3D3D3")
    thin_border = Border(left=side, right=side, top=side, bottom=side)

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    for sheetname in writer.sheets:
        ws = writer.sheets[sheetname]
        is_individual_tab = "Individuals" in sheetname

        eng_col_idx = None
        if is_individual_tab:
            for idx, cell in enumerate(ws[1], 1):
                if cell.value == "Total Engagements":
                    eng_col_idx = idx
                    break

        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        ws.row_dimensions[1].height = 28

        for row_num in range(2, ws.max_row + 1):
            count_val = 0
            if eng_col_idx:
                try:
                    count_val = int(ws.cell(row=row_num, column=eng_col_idx).value or 0)
                except (TypeError, ValueError):
                    pass

            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_num, column=col_num)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="left", vertical="center")

                if is_individual_tab and col_num == eng_col_idx:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Calibri", size=11, bold=True)
                    if count_val >= 3:
                        cell.fill = green_fill
                    elif count_val == 2:
                        cell.fill = yellow_fill
                    elif count_val == 1:
                        cell.fill = red_fill

        for col_num in range(1, ws.max_column + 1):
            max_len = 0
            for row_num in range(1, ws.max_row + 1):
                val = ws.cell(row=row_num, column=col_num).value
                if val:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[get_column_letter(col_num)].width = max(max_len + 5, 18)


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def generate_multi_campaign_reports():
    print("--- Lemlist Automated Report Generator ---")

    api_key = get_lemlist_api_key()
    if not api_key:
        print("[Error] Lemlist API Key not found in config.txt!")
        return

    print("[API] Fetching campaigns...")
    all_campaigns = fetch_all_campaigns(api_key)
    if not all_campaigns:
        print("[Warning] No campaigns found.")
        return

    # Show every campaign's status so you can confirm what counts as "active"
    print(f"[API] {len(all_campaigns)} campaigns found:")
    for c in all_campaigns:
        print(f"   - {c.get('name', '?')}  [status: {c.get('status', 'unknown')}]")

    campaigns = [
        c for c in all_campaigns
        if str(c.get("status", "")).lower() in ACTIVE_STATUSES
    ]
    if not campaigns:
        print(
            f"[Warning] No campaigns matched ACTIVE_STATUSES={ACTIVE_STATUSES}. "
            "Check the status values printed above and update ACTIVE_STATUSES."
        )
        return
    print(f"[API] {len(campaigns)} active campaign(s) available.\n")

    campaigns = select_campaigns(campaigns)
    if not campaigns:
        print("No campaigns selected. Nothing to do.")
        return
    print(f"\nRunning {len(campaigns)} campaign(s):")
    for c in campaigns:
        print(f"   - {c.get('name', 'Unknown Campaign')}")
    print()

    now_local = datetime.now().astimezone()
    local_tz = now_local.tzinfo
    cutoff_date = (now_local - timedelta(days=DAYS_BACK)).date()
    print(f"Window: {cutoff_date} through {now_local.date()} (your local time)\n")
    # stop paging once we hit events older than this (1 extra day of safety margin)
    stop_before = (
        pd.Timestamp(cutoff_date).tz_localize(local_tz) - pd.Timedelta(days=1)
    ).tz_convert("UTC")

    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_folder = os.path.join(
        desktop_path, f"Lemlist Stats - {now_local.strftime('%Y-%m-%d')}"
    )
    os.makedirs(output_folder, exist_ok=True)
    timestamp_str = now_local.strftime("%H-%M-%S")
    lead_cache = {}  # email -> company, shared across campaigns so each lead is looked up once

    for camp in campaigns:
        camp_id = camp.get("_id")
        camp_name = camp.get("name", "Unknown Campaign")
        if not camp_id:
            print(f"Skipping campaign without ID: {camp_name}")
            continue

        print(f"Fetching activities for: {camp_name}...")
        raw_opens = fetch_all_activities(api_key, camp_id, "emailsOpened", stop_before)
        raw_clicks = fetch_all_activities(api_key, camp_id, "emailsClicked", stop_before)
        print(f"   raw from API: {len(raw_opens)} opens, {len(raw_clicks)} clicks")
        if DEBUG_SHOW_FIELDS and (raw_opens or raw_clicks):
            sample = (raw_opens or raw_clicks)[0]
            print(f"   [Debug] fields on an activity: {sorted(sample.keys())}")

        opens_df = to_clean_df(raw_opens, camp_name, local_tz, cutoff_date)
        clicks_df = to_clean_df(raw_clicks, camp_name, local_tz, cutoff_date)

        opens_df = add_company_names(opens_df, api_key, lead_cache)
        clicks_df = add_company_names(clicks_df, api_key, lead_cache)

        if opens_df.empty and clicks_df.empty:
            print(f"-> No human opens or clicks in the last {DAYS_BACK} days for {camp_name}\n")
            continue

        clean_name = str(camp_name).strip()
        for ch in ["/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
            clean_name = clean_name.replace(ch, "_")

        output_file = os.path.join(
            output_folder, f"{clean_name} 4-Tab Report - {timestamp_str}.xlsx"
        )

        t1_open, t2_open = build_tab_data(opens_df, "Opens")
        t1_click, t2_click = build_tab_data(clicks_df, "Clicks")

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            t1_open.to_excel(writer, sheet_name="Opens - Individuals", index=False)
            t2_open.to_excel(writer, sheet_name="Opens - Top Companies", index=False)
            t1_click.to_excel(writer, sheet_name="Clicks - Individuals", index=False)
            t2_click.to_excel(writer, sheet_name="Clicks - Top Companies", index=False)
            style_workbook(writer)

        print(
            f"[Success] {camp_name} -> "
            f"Opens: {len(opens_df)} total / {len(t1_open)} unique people | "
            f"Clicks: {len(clicks_df)} total / {len(t1_click)} unique people\n"
        )

    print(f"All reports generated in folder: {output_folder}")


if __name__ == "__main__":
    generate_multi_campaign_reports()
