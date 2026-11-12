#!/usr/bin/env python3
"""
PM Surya Ghar Report Maker
==========================
A small offline desktop app. Pick the SBPDCL and NBPDCL consumer-level exports
from the PM Surya Ghar national portal (CSV, any size) and an output folder, press
"Create report". It writes ONE Excel workbook that is a complete report. Every
sheet is a single plain table (a header row, then data: no titles or notes):

    Headline KPIs       the main figures for Bihar, SBPDCL and NBPDCL
    Funnel              application to installation, with conversion percentages
    District KPIs       one row per district, with targets and DISCOM / Bihar totals
    Division KPIs       one row per division
    Sub-division KPIs   one row per sub-division
    Section KPIs        one row per section (made only if the export has a Section column)
    Recent progress     last 7 / 15 / 30 days, 12 weeks, 6 and 12 months, by district
    Recent summary      the same periods for Bihar and each DISCOM, with dates
    Monthly, Weekly and Daily trend
    Pendency            pending cases by district, next step, who must act, and age
    Who has to act      the same for Bihar and each DISCOM

Each district is shown only under its own DISCOM. If an application's District column
names a district of the other DISCOM, it is counted under the main district of its own
supply division ("District not identified" when the division does not say).

Only Regular applications are reported. A row whose Model column says ULA is left out of
every count, table and date, as if it were not in the export. The log says how many were left out.

Percentages, balances, ranks and totals are Excel formulas, so they follow if a
target or figure is edited. The workbook holds counts only: no consumer name,
consumer number, mobile, email, GPS point or bank detail is written. Data checks
on the exports are shown in the app's log, not in the workbook.

Run it
    Double-click, or:   python pmsg_report_app.py
    Without a window:   python pmsg_report_app.py --sbpdcl S.csv --nbpdcl N.csv --out folder
                        add  --dashboard LINK --password PASSWORD  to upload it and open the dashboard

Dashboard (optional, needs internet)
    Paste the dashboard's web app link. With the tick-box on, each new report is sent to the
    dashboard's Google Sheet and the dashboard is opened in the browser. Making the report itself
    works offline.
    The app ASKS for the upload password (the one set in Upload.gs of the Apps Script project)
    when it first needs it. It keeps it only until the app is closed and never writes it to disk.
    To fix the link inside the app so that nobody has to type it, fill in DASHBOARD_LINK just
    below these notes. DASHBOARD_PASSWORD does the same for the password, if wanted.

District targets are built in. To change them, choose a targets CSV in the app:
first column District, second column Target. Districts not listed in it keep their
built-in target. The app can save a blank file with just these two headings.

Needs only Python 3.8+ (no extra packages). To make a single .exe on Windows:
    pip install pyinstaller
    pyinstaller --onefile --windowed --name PMSG_Report_Maker pmsg_report_app.py

Give it the files exactly as downloaded. Do not open and re-save them in Excel
first: Excel rewrites dates and long numbers.
"""
import argparse
import csv
import datetime as dt
import io
import json
import math
import os
import queue
import re
import sys
import threading
import base64
import socket
import urllib.error
import urllib.request
import webbrowser
import zipfile
from collections import Counter, defaultdict
from xml.sax.saxutils import escape

APP_NAME = "PM Surya Ghar Report Maker"
APP_VERSION = "3.0"

# ----------------------------------------------------------------------------
# FIXED DASHBOARD SETTINGS  (fill these in once; then nobody has to type them in the app)
# ----------------------------------------------------------------------------
# The dashboard's web app link, between the quotes. Copy it from the Apps Script project:
# Deploy > Manage deployments > Web app. It starts with https://script.google.com/ and ends with /exec.
# (This is NOT the Google Sheet link.) Leave it empty to type the link in the app instead.
DASHBOARD_LINK = ""

# The upload password, between the quotes: the same one as UPLOAD_PASSWORD in Upload.gs.
# Leave it empty and the app asks for the password when it sends a report. If you fill it in, anyone who gets this
# file, or the .exe made from it, can read the password and upload to the dashboard.
DASHBOARD_PASSWORD = ""

# ----------------------------------------------------------------------------
# Definitions
# ----------------------------------------------------------------------------
# Funnel steps in order. An application is counted at a step when it has reached
# that step OR ANY LATER ONE.
STEPS = ["applications", "feasibility_approved", "vendor_selected", "vendor_accepted",
         "agreement_uploaded", "installed", "inspected", "subsidy_released"]
# Where an application stands today = the last step completed (up to installation).
POSITIONS = ["at_applied", "at_feasibility", "at_vendor_selected", "at_vendor_accepted", "at_agreement"]
EXTRA = ["installed_kw", "subsidy_released_lakh", "pending_consumer", "pending_vendor", "pending_discom",
         "pending_rec", "give_up_subsidy", "rural", "urban"]
MEASURES = STEPS + POSITIONS + EXTRA
M = {name: i for i, name in enumerate(MEASURES)}

# Things that happen on a date
EVENTS = ["applications", "feasibility_approved", "vendor_selected", "agreements",
          "installations", "installed_kw", "subsidy_released"]
E = {name: i for i, name in enumerate(EVENTS)}
EVENTS_LOWER_LEVELS = ["applications", "vendor_selected", "agreements", "installations"]
EVENTS_WITH_PREVIOUS = ["applications", "vendor_selected", "agreements", "installations"]

# Recent periods, all ending on the extract date: (label, bucket size, number of buckets)
WINDOWS = [("Last 7 days", "day", 7), ("Last 15 days", "day", 15), ("Last 30 days", "day", 30),
           ("Last 12 weeks", "week", 12), ("Last 6 months", "month", 6), ("Last 12 months", "month", 12)]
AGE_BUCKETS = ["0-30 days", "31-90 days", "91-180 days", "Over 180 days", "No date"]
LEVELS = ["Bihar", "DISCOM", "District", "Division", "Sub-division", "Section"]
BIHAR = ("Bihar", "Bihar", "", "", "", "")

# Columns read from the export, matched by header name (case and spacing ignored).
# A tuple lists alternative header names.
COLUMNS = {
    "app_no": "application number",
    "app_date": "initial application submission date",
    "cur_stage": "current stage",
    "next_stage": "next stage",
    "action_by": "action to be taken by",
    "last_action": "last action date",
    "district": "district",
    "division": "division name",
    "subdivision": "sub-division name",
    "feas_status": "feasibility status",
    "feas_date": "feasibility status update date",
    "vendor_status": "vendor stage status",
    "vendor_sel_date": "initial date of vendor selection by consumer",
    "vendor_id": "vendor id",
    "agreement": "vendor - customer agreement uploaded",
    "agreement_date": "agreement upload date",
    "inst_status": "installation status",
    "inst_kw": "installed capacity (in kw)",
    "inst_date": "initial date of installation by vendor",
    "insp_status": "inspection status",
    "sub_status": "subsidy stage status",
    "sub_date": "subsidy release date",
    "sub_amount": "subsidy amount released to consumer account",
    # optional: the report still runs if these are missing
    "discom_name": "discom (full name)",
    "circle": "circle name",
    "consumer_no": "consumer number",
    "rural_urban": "rural/urban",
    "lat": "latitude",
    "lon": "longitude",
    "giveup": "give-up subsidy (yes/no)",
    "section": ("section name", "section", "section office", "section office name", "sub-station name"),
    "model": "model",
}
OPTIONAL = {"discom_name", "circle", "consumer_no", "rural_urban", "lat", "lon", "giveup", "section", "model"}
# The Model column says Regular or ULA. Rows of these models are left out of everything.
MODELS_LEFT_OUT = {"ula"}
MODELS_KNOWN = {"regular"} | MODELS_LEFT_OUT
FIELDS = list(COLUMNS)

CHECKS = [
    ("app_date_day1", "Application date is the 1st of a month",
     "If this is most rows, the day was lost before export or by Excel; daily and weekly application counts are then unreliable."),
    ("app_date_missing", "Application date missing or unreadable", "Not counted in any dated table."),
    ("dates_out_of_order", "Stage dates out of order",
     "A later step is dated before an earlier one (application, vendor selection, agreement, installation, subsidy)."),
    ("out_of_sequence", "Later step done, earlier step not marked",
     "Counted as having reached the earlier step too, so the funnel never widens."),
    ("agreement_stage_mismatch", "Agreement uploaded but stage still shows Upload Agreement",
     "Status columns disagree; the agreement flag is used."),
    ("event_after_extract", "Step dated after the extract date", "Left out of the dated tables."),
    ("gps_missing", "GPS location missing or zero", "Cannot be mapped."),
    ("gps_outside", "GPS location outside Bihar", "Latitude 24.0-27.6, longitude 83.0-88.4 expected."),
    ("rural_urban_blank", "Rural / Urban not stated", ""),
    ("dup_app", "Repeat application number", "Counted once."),
    ("dup_consumer", "Consumer number with more than one application", "Each application is still counted."),
    ("no_app_no", "Rows without an application number", "Left out of every count."),
    ("short_rows", "Broken rows (too few columns)", "Left out; usually a line break inside a remarks field."),
    ("discom_mismatch", "Row says the other DISCOM than the file it came in",
     "The row's own DISCOM column is used."),
    ("district_outside", "District column names a district of the other DISCOM",
     "Counted under the main district of the application's own supply division."),
]

# District installation targets ("Tentative Target on a First-Come, First-Served Basis"),
# as read from the review presentation for the week 27 Sep - 3 Oct 2026. They add up to
# 5,00,000. A district_targets.csv beside the app changes them.
DEFAULT_TARGETS = {
    "Patna": 28000, "Muzaffarpur": 23000, "Nalanda": 14000, "Bhagalpur": 15000, "Vaishali": 17000,
    "Gaya": 21000, "Samastipur": 21000, "Saran": 19000, "Purbi Champaran": 24000, "Darbhanga": 19000,
    "Sitamarhi": 16000, "Pashchim Champaran": 19000, "Begusarai": 14000, "Nawada": 11000,
    "Bhojpur": 13000, "Supaul": 11000, "Rohtas": 14000, "Khagaria": 8000, "Purnia": 16000,
    "Munger": 7000, "Kaimur Bhabua": 8000, "Madhubani": 22000, "Madhepura": 10000, "Siwan": 16000,
    "Aurangabad": 12000, "Katihar": 15000, "Jehanabad": 5000, "Buxar": 8000, "Banka": 10000,
    "Gopalganj": 12000, "Araria": 13000, "Saharsa": 9000, "Jamui": 8000, "Sheohar": 3000,
    "Sheikhpura": 3000, "Arwal": 3000, "Kishanganj": 8000, "Lakhisarai": 5000,
}
# The 17 districts SBPDCL serves; the other 21 are NBPDCL's. An application whose District
# column names a district of the other DISCOM is counted under the main district of its own
# supply division instead (the District column is filled in by the applicant and is sometimes wrong).
SOUTH_BIHAR = ["Patna", "Nalanda", "Bhagalpur", "Gaya", "Nawada", "Bhojpur", "Rohtas", "Munger", "Kaimur Bhabua",
               "Aurangabad", "Jehanabad", "Buxar", "Banka", "Jamui", "Sheikhpura", "Arwal", "Lakhisarai"]
UNPLACED = "District not identified"
_STRAY = "\x00"                                   # marks a district still to be worked out from the division

DISTRICT_ALIASES = {
    "eastchamparan": "purbichamparan", "purvichamparan": "purbichamparan", "purbachamparan": "purbichamparan",
    "westchamparan": "pashchimchamparan", "paschimchamparan": "pashchimchamparan",
    "purnea": "purnia", "kaimur": "kaimurbhabua", "bhabua": "kaimurbhabua",
}

DATE_FORMATS = ("%d-%b-%y", "%d-%b-%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
                "%d-%m-%Y %H:%M", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M",
                "%Y-%m-%d %H:%M:%S", "%d %b %Y", "%d-%B-%Y")
# tried on the date part alone when a value has a time after it, or is written another way (day first, as in India)
DATE_ONLY_FORMATS = ("%d-%b-%y", "%d-%b-%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%b/%Y",
                     "%d/%b/%y", "%d-%B-%Y", "%d-%m-%y", "%d/%m/%y")
_date_part = re.compile(r"^\s*(\d{1,4}[-/.][A-Za-z0-9]{1,9}[-/.]\d{1,4})")
DATE_COLUMNS = [("app_date", "Initial Application Submission Date"), ("feas_date", "Feasibility Status Update Date"),
                ("vendor_sel_date", "Initial Date of Vendor Selection by Consumer"), ("agreement_date", "Agreement Upload Date"),
                ("inst_date", "Initial Date of Installation by Vendor"), ("sub_date", "Subsidy Release Date"),
                ("last_action", "Last Action Date")]

# Status wordings seen in the sample export, and how each column is counted. A wording outside
# these lists is reported, so nothing is counted silently in a way nobody has looked at.
KNOWN_VALUES = {
    "Feasibility Status": ({"approved", "pending"}, "Only 'Approved' counts as feasibility approved."),
    "Vendor Stage Status": ({"vendor selected by consumer and accepted by vendor", "vendor selection pending consumer",
                             "vendor selected by consumer but vendor acceptance pending"},
                            "It counts as vendor selected if it starts with 'Vendor Selected', and as vendor accepted if it contains 'Accepted by Vendor'."),
    "Agreement Uploaded": ({"yes", "no"}, "Only 'Yes' counts as agreement uploaded."),
    "Installation Status": ({"installation completed and accepted by consumer", "installation pending"},
                            "It counts as installed if it starts with 'Installation Completed'."),
    "Inspection Status": ({"approved"}, "Only 'Approved' counts as inspection approved."),
    "Subsidy Stage Status": ({"subsidy released", "subsidy pending for verification"}, "Only 'Subsidy Released' counts as subsidy released."),
    "Action to be taken by": ({"consumer", "vendor", "discom", "rec"}, "Only consumer, vendor, DISCOM and REC are counted in the 'Waiting on' columns."),
}


class UserError(Exception):
    """A problem the person can fix; shown as a plain message."""


class Cancelled(UserError):
    """The person pressed Cancel."""


# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------
def indian(n):
    """12345678 -> '1,23,45,678'"""
    s = str(abs(int(n)))
    if len(s) > 3:
        s = re.sub(r"(\d)(?=(\d\d)+$)", r"\1,", s[:-3]) + "," + s[-3:]
    return ("-" if n < 0 else "") + s


def app_folder():
    """Folder of the .exe (when packaged) or of this script."""
    return os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))


def norm(s):
    return re.sub(r"\s+", " ", s.strip().lower())


_date_cache = {}


def parse_date(s):
    if not s:
        return None
    d = _date_cache.get(s, 0)
    if d != 0:
        return d
    d = None
    for f in DATE_FORMATS:
        try:
            d = dt.datetime.strptime(s, f).date()
            break
        except ValueError:
            pass
    if d is None:                                 # a time after the date, or another separator: read the date part alone
        m = _date_part.match(s)
        if m:
            for f in DATE_ONLY_FORMATS:
                try:
                    d = dt.datetime.strptime(m.group(1), f).date()
                    break
                except ValueError:
                    pass
    if d is not None and not 2000 <= d.year <= 2100:
        d = None                                  # not a believable date
    _date_cache[s] = d
    return d


def num(s):
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return 0.0


_name_cache = {}


def tidy_name(s):
    """'KANKARBAGH  II' -> 'Kankarbagh II', 'PESU WEST' -> 'PESU West', '' -> 'Not stated'."""
    t = _name_cache.get(s)
    if t is None:
        t = re.sub(r"\s+", " ", s.strip()).title()
        t = re.sub(r"\b(Ii|Iii|Iv|Vi|Vii|Viii|Ix)\b", lambda m: m.group(0).upper(), t)
        t = re.sub(r"\bPesu\b", "PESU", t) or "Not stated"
        _name_cache[s] = t
    return t


def place_key(s):
    """Loose district match: 'Gaya Ji' = 'GAYA', 'Kaimur (Bhabua)' = 'Kaimur Bhabua', 'Purnea' = 'Purnia'."""
    k = re.sub(r"[^a-z]", "", re.sub(r"\bji\b", "", s.lower()))
    return DISTRICT_ALIASES.get(k, k)


TARGETS_TEMPLATE = "District,Target\r\n"          # the blank file the app can save: headings only


_home = {}


def home_discom(district):
    """'SBPDCL' or 'NBPDCL' for one of Bihar's 38 districts, None for any other name."""
    if not _home:
        south = {place_key(d) for d in SOUTH_BIHAR}
        for name in DEFAULT_TARGETS:
            _home[place_key(name)] = "SBPDCL" if place_key(name) in south else "NBPDCL"
    return _home.get(place_key(district))


def load_targets(path, ignored=None):
    """Reads a targets CSV: first column District, second column Target. A heading row is skipped.
    Districts whose target is not a whole number are added to the list `ignored`."""
    if not os.path.isfile(path):
        raise UserError(f"Targets file not found:\n{path}")
    if os.path.splitext(path)[1].lower() in (".xlsx", ".xls", ".xlsm"):
        raise UserError("The targets file is an Excel file. Save it as CSV (comma separated) and choose that.")
    targets = {}
    try:
        with open(path, "rb") as raw:
            bom = raw.read(2)
        encoding = "utf-16" if bom in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
        with open(path, encoding=encoding, errors="replace", newline="") as fh:
            first = True
            for row in csv.reader(fh):
                if not row or not row[0].strip():
                    continue
                number = re.sub(r"[,\s]", "", row[1]) if len(row) > 1 else ""
                if re.fullmatch(r"\d+(\.0+)?", number):
                    targets[re.sub(r"\s+", " ", row[0].strip())] = int(float(number))
                elif not first and ignored is not None:      # the first row may be the headings
                    ignored.append(row[0].strip())
                first = False
    except PermissionError:
        raise UserError("The targets file cannot be opened. If it is open in Excel, close it and try again.")
    except (OSError, csv.Error) as exc:
        raise UserError(f"The targets file cannot be read: {exc}")
    if not targets:
        raise UserError("No targets were found in the file. It needs District in the first column and Target (a whole number) in the second.")
    return targets


def check_targets(path):
    """State of a chosen targets file for the window: (state, message)."""
    if not path:
        return "empty", "Built-in targets are used"
    ignored = []
    try:
        found = load_targets(path, ignored)
    except UserError as exc:
        message = str(exc).split("\n")[0]
        return "bad", "No District / Target rows found in this file" if message.startswith("No targets were found") else message
    return "ok", (f"{len(found)} target{'s' if len(found) != 1 else ''} read; the rest stay built-in"
                  + (f" ({len(ignored)} row{'s' if len(ignored) != 1 else ''} ignored)" if ignored else ""))


def month_start(d, back):
    k = d.year * 12 + d.month - 1 - back
    return dt.date(k // 12, k % 12 + 1, 1)


def window_buckets(as_of, kind, n):
    """Buckets of a period, oldest first: (start, end, label)."""
    out = []
    for i in range(n - 1, -1, -1):
        if kind == "day":
            d = as_of - dt.timedelta(days=i)
            out.append((d, d, d.strftime("%d %b")))
        elif kind == "week":                      # 7-day blocks ending on the extract date
            end = as_of - dt.timedelta(days=7 * i)
            out.append((end - dt.timedelta(days=6), end, end.strftime("%d %b")))
        else:                                     # calendar months; the latest is month to date
            start = month_start(as_of, i)
            end = min(month_start(as_of, i - 1) - dt.timedelta(days=1), as_of)
            out.append((start, end, start.strftime("%b %Y")))
    return out


def ymd(d):
    return d.year * 10000 + d.month * 100 + d.day


def from_ymd(n):
    return dt.date(n // 10000, n // 100 % 100, n % 100)


# ----------------------------------------------------------------------------
# Reading the exports
# ----------------------------------------------------------------------------
def summarise(inputs, targets_path=None, progress=None, log=None, cancel=None, where=None):
    """inputs: list of (label, path), label being 'SBPDCL' or 'NBPDCL'.
    Returns (rows for the DATA sheet, extract date, facts)."""
    log = log or (lambda msg: None)
    progress = progress or (lambda phase, fraction, msg: None)
    cancel = cancel or (lambda: False)
    where = where if where is not None else {}
    csv.field_size_limit(10 ** 9)
    today = dt.date.today()

    named_targets = dict(DEFAULT_TARGETS)
    imported, ignored_rows = {}, []
    if targets_path:                              # districts listed in the file replace the built-in figure; the rest keep it
        built_in = {place_key(k): k for k in named_targets}
        imported = load_targets(targets_path, ignored_rows)
        for name, value in imported.items():
            named_targets.pop(built_in.get(place_key(name), name), None)
            named_targets[name] = value
        changed = sum(1 for name, value in imported.items() if DEFAULT_TARGETS.get(built_in.get(place_key(name))) != value)
        log(f"Targets: {len(imported)} read from {os.path.basename(targets_path)} ({changed} differ from the built-in figure); other districts keep the built-in target")
    targets = {place_key(k): v for k, v in named_targets.items()}
    target_name = {place_key(k): k for k in named_targets}

    office = defaultdict(lambda: [0.0] * len(MEASURES))   # office key -> measures
    office_id = {}                                         # office key -> small number
    events = {}                                            # (office id, day number) -> counts per event
    circle_votes = defaultdict(Counter)                    # division / sub-division key -> circle
    parent_votes = defaultdict(Counter)                    # sub-division key -> division
    pend = Counter()                                       # (office id, stage, who, day number or 0)
    checks = defaultdict(Counter)                          # discom -> check -> rows
    audit = Counter()                                      # (discom, column, value) -> rows
    latest = {}                                            # discom -> latest action date
    rows_read = Counter()
    files = {}
    seen_app, seen_consumer = set(), set()
    unread_dates, unread_example = Counter(), {}  # date column -> values that could not be read
    stray_names = {}                              # (discom, district of the other DISCOM) -> applications
    combo_cache = {}
    section_seen = False
    model_rows = defaultdict(Counter)             # file -> Model value as written -> rows
    no_model_column = []                          # files without a Model column

    def key_id(key):
        i = office_id.get(key)
        if i is None:
            i = office_id[key] = len(office_id)
        return i

    for slot, path in inputs:
        if not os.path.isfile(path):
            raise UserError(f"{slot} file not found:\n{path}")

    for slot, path in inputs:
        if os.path.splitext(path)[1].lower() in (".xlsx", ".xls", ".xlsm"):
            raise UserError(f"{os.path.basename(path)} is an Excel file. Choose the CSV file as downloaded from the portal.")
        files[slot] = os.path.basename(path)
        size = os.path.getsize(path) or 1
        log(f"Reading {slot}: {os.path.basename(path)} ({size / 1e6:,.1f} MB)")
        progress(slot, 0.0, f"Reading {slot}…")
        try:
            raw = open(path, "rb")
        except PermissionError:
            raise UserError(f"{os.path.basename(path)} cannot be opened. If it is open in Excel, close it and try again.")
        except OSError as exc:
            raise UserError(f"{os.path.basename(path)} cannot be opened: {exc.strerror or exc}")
        with raw:
            bom = raw.read(2)
            raw.seek(0)
            encoding = "utf-16" if bom in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
            fh = io.TextIOWrapper(raw, encoding=encoding, errors="replace", newline="")
            reader = csv.reader(fh)
            where.update(file=os.path.basename(path), reader=reader)
            col = None
            for row in reader:
                names = [norm(c) for c in row]
                if "application number" in names and "district" in names:
                    col = {}
                    for i, n in enumerate(names):
                        col.setdefault(n, i)
                    break
            if col is None:
                raise UserError(f"{os.path.basename(path)} does not look like the portal export: "
                                "no header row with 'Application Number' and 'District' was found.")

            def find(names):
                for name in ((names,) if isinstance(names, str) else names):
                    if name in col:
                        return col[name]
                return -1

            pos = {k: find(v) for k, v in COLUMNS.items()}
            missing = [v for k, v in COLUMNS.items() if pos[k] < 0 and k not in OPTIONAL]
            if missing:
                raise UserError(f"{os.path.basename(path)} is missing these columns:\n  " + "\n  ".join(missing))
            absent = [k.replace("_", " ") for k in OPTIONAL if pos[k] < 0 and k not in ("section", "model")]
            if absent:
                log("  Optional columns not in this file (related checks are skipped): " + ", ".join(sorted(absent)))
            idx = [pos[k] for k in FIELDS]
            width = max(idx) + 1
            has_gps = pos["lat"] >= 0 and pos["lon"] >= 0
            has_ru = pos["rural_urban"] >= 0
            has_section = pos["section"] >= 0
            section_seen = section_seen or has_section
            log("  Section column found: section-wise KPIs will be made" if has_section
                else "  No Section column in this file: section-wise KPIs cannot be made from it")
            has_model = pos["model"] >= 0
            if not has_model:
                no_model_column.append(slot)
                log("  No Model column in this file: ULA rows cannot be told apart, so every row is counted")
            models = model_rows[slot]
            ck_slot = checks[slot]
            n = 0

            for row in reader:
                n += 1
                if n % 10000 == 0:
                    if cancel():
                        raise Cancelled("Cancelled.")
                    progress(slot, min(0.999, raw.tell() / size), f"Reading {slot}: {indian(n)} rows")
                if len(row) < width:
                    if any(c.strip() for c in row):
                        ck_slot["short_rows"] += 1
                    continue
                if row[0].strip().lower() == "logic":        # the explanatory row under the header
                    continue
                (app_no, app_date, cur_stage, next_stage, action_by, last_action, district, division,
                 subdivision, feas_status, feas_date, vendor_status, vendor_sel_date, vendor_id, agreement,
                 agreement_date, inst_status, inst_kw, inst_date, insp_status, sub_status, sub_date,
                 sub_amount, discom_name, circle, consumer_no, rural_urban, lat, lon, giveup, section, model) = \
                    [row[i].strip() if i >= 0 else "" for i in idx]

                if has_model:                                 # only Regular applications are reported
                    models[model] += 1
                    if model.lower() in MODELS_LEFT_OUT:
                        continue

                dn = discom_name.lower()
                discom = "NBPDCL" if "north" in dn else "SBPDCL" if "south" in dn else slot
                ck = checks[discom]
                rows_read[discom] += 1
                if discom != slot:
                    ck["discom_mismatch"] += 1
                if not app_no:
                    ck["no_app_no"] += 1
                    continue
                hashed = hash(app_no)
                if hashed in seen_app:
                    ck["dup_app"] += 1
                    continue                                  # count each application once
                seen_app.add(hashed)
                if consumer_no:
                    hashed = hash(consumer_no.lstrip("'"))
                    if hashed in seen_consumer:
                        ck["dup_consumer"] += 1
                    seen_consumer.add(hashed)

                audit[(discom, "Current Stage", cur_stage or "(blank)")] += 1
                audit[(discom, "Next Stage", next_stage or "(blank)")] += 1
                audit[(discom, "Action to be taken by", action_by or "(blank)")] += 1
                audit[(discom, "Feasibility Status", feas_status or "(blank)")] += 1
                audit[(discom, "Vendor Stage Status", vendor_status or "(blank)")] += 1
                audit[(discom, "Agreement Uploaded", agreement or "(blank)")] += 1
                audit[(discom, "Installation Status", inst_status or "(blank)")] += 1
                audit[(discom, "Inspection Status", insp_status or "(blank)")] += 1
                audit[(discom, "Subsidy Stage Status", sub_status or "(blank)")] += 1

                # ---- the offices this row belongs to
                combo = (discom, district, division, subdivision, section)
                hit = combo_cache.get(combo)
                if hit is None:
                    d_, v_, s_ = tidy_name(district), tidy_name(division), tidy_name(subdivision)
                    home = home_discom(d_)
                    stray = home is not None and home != discom
                    if stray:                     # the district belongs to the other DISCOM: place it by division later
                        stray_names.setdefault((discom, d_), 0)
                        d_ = _STRAY + v_
                    keys = [BIHAR, ("DISCOM", discom, "", "", "", ""), ("District", discom, d_, "", "", ""),
                            ("Division", discom, d_, v_, "", ""), ("Sub-division", discom, d_, "", s_, "")]
                    sec_key = ("Section", discom, d_, "", s_, tidy_name(section)) if has_section else None
                    hit = combo_cache[combo] = ([office[k] for k in keys] + ([office[sec_key]] if sec_key else []),
                                                [key_id(k) for k in keys], keys[3], keys[4], v_,
                                                key_id(sec_key) if sec_key else -1,
                                                (discom, tidy_name(district)) if stray else None)
                lists, ids, key_div, key_sub, div_name, sec_id, stray_key = hit
                if stray_key:
                    ck["district_outside"] += 1
                    stray_names[stray_key] += 1
                parent_votes[key_sub][div_name] += 1
                if circle:
                    circle_votes[key_div][circle] += 1
                    circle_votes[key_sub][circle] += 1

                # ---- step flags, then make them cumulative from the last step backwards
                vs = vendor_status.lower()
                raw_flags = [
                    True,
                    feas_status.lower() == "approved",
                    vs.startswith("vendor selected") or bool(vendor_id),
                    "accepted by vendor" in vs,
                    agreement.lower() == "yes",
                    inst_status.lower().startswith("installation completed"),
                    insp_status.lower() == "approved",
                    sub_status.lower() == "subsidy released",
                ]
                reached = raw_flags[:]
                for i in range(6, -1, -1):
                    reached[i] = reached[i] or reached[i + 1]
                if reached != raw_flags:
                    ck["out_of_sequence"] += 1

                kw = num(inst_kw) if reached[5] else 0.0
                actor = action_by.lower()
                ru = rural_urban.lower()
                adds = [(i, 1.0) for i in range(8) if reached[i]]
                furthest = 5 if reached[5] else 4 if reached[4] else 3 if reached[3] else 2 if reached[2] else 1 if reached[1] else 0
                if furthest < 5:
                    adds.append((M[POSITIONS[furthest]], 1.0))
                if kw:
                    adds.append((M["installed_kw"], kw))
                if reached[7]:
                    amt = num(sub_amount)
                    if amt:
                        adds.append((M["subsidy_released_lakh"], amt / 1e5))
                if actor in ("consumer", "vendor", "discom", "rec"):
                    adds.append((M["pending_" + actor], 1.0))
                if giveup.lower() == "yes":
                    adds.append((M["give_up_subsidy"], 1.0))
                if ru in ("rural", "urban"):
                    adds.append((M[ru], 1.0))
                for lst in lists:
                    for i, a in adds:
                        lst[i] += a

                # ---- dated events
                d_app, d_feas, d_vs = parse_date(app_date), parse_date(feas_date), parse_date(vendor_sel_date)
                d_agr, d_inst, d_sub = parse_date(agreement_date), parse_date(inst_date), parse_date(sub_date)
                d_last = parse_date(last_action)
                if ((app_date and d_app is None) or (feas_date and d_feas is None) or (vendor_sel_date and d_vs is None)
                        or (agreement_date and d_agr is None) or (inst_date and d_inst is None) or (sub_date and d_sub is None)
                        or (last_action and d_last is None)):
                    for label, written, got in (("Initial Application Submission Date", app_date, d_app),
                                                ("Feasibility Status Update Date", feas_date, d_feas),
                                                ("Initial Date of Vendor Selection by Consumer", vendor_sel_date, d_vs),
                                                ("Agreement Upload Date", agreement_date, d_agr),
                                                ("Initial Date of Installation by Vendor", inst_date, d_inst),
                                                ("Subsidy Release Date", sub_date, d_sub), ("Last Action Date", last_action, d_last)):
                        if written and got is None:
                            unread_dates[label] += 1
                            unread_example.setdefault(label, written)
                for e, d, flag in ((0, d_app, True), (1, d_feas, raw_flags[1]), (2, d_vs, raw_flags[2]),
                                   (3, d_agr, raw_flags[4]), (4, d_inst, raw_flags[5]), (6, d_sub, raw_flags[7])):
                    if flag and d is not None:
                        o = d.toordinal()
                        for i in (ids + [sec_id] if e == 4 and sec_id >= 0 else ids):   # sections keep installations only
                            slot_ = events.get((i, o))
                            if slot_ is None:
                                slot_ = events[(i, o)] = [0.0] * len(EVENTS)
                            slot_[e] += 1
                            if e == 4:
                                slot_[5] += kw

                if d_last is not None and d_last <= today:
                    if discom not in latest or d_last > latest[discom]:
                        latest[discom] = d_last

                # ---- pending: who has to act, and since when
                stage = next_stage or (cur_stage if actor else "")
                if stage:
                    who = {"rec": "REC", "discom": "DISCOM"}.get(actor, actor.title() or "Not stated")
                    o = d_last.toordinal() if d_last else 0
                    for i in ids[:3]:
                        pend[(i, stage, who, o)] += 1

                # ---- data checks
                if d_app is None:
                    ck["app_date_missing"] += 1
                elif d_app.day == 1:
                    ck["app_date_day1"] += 1
                seq = [d for d in (d_app, d_vs, d_agr, d_inst, d_sub) if d]
                for a, b in zip(seq, seq[1:]):
                    if b < a:
                        ck["dates_out_of_order"] += 1
                        break
                if has_gps:
                    la_, lo_ = num(lat), num(lon)
                    if la_ == 0 or lo_ == 0:
                        ck["gps_missing"] += 1
                    elif not (24.0 <= la_ <= 27.6 and 83.0 <= lo_ <= 88.4):
                        ck["gps_outside"] += 1
                if has_ru and ru not in ("rural", "urban"):
                    ck["rural_urban_blank"] += 1
                if raw_flags[4] and cur_stage.lower() == "upload agreement":
                    ck["agreement_stage_mismatch"] += 1

            fh.detach()
        progress(slot, 1.0, f"{slot}: {indian(n)} lines read")
        log(f"  {indian(n)} lines read")
        if has_model:
            left = sum(c for v, c in models.items() if v.lower() in MODELS_LEFT_OUT)
            log(f"  Model column: {indian(sum(models.values()) - left)} rows used, {indian(left)} ULA rows left out")

    left_out = {slot: sum(c for v, c in m.items() if v.lower() in MODELS_LEFT_OUT) for slot, m in model_rows.items()}
    apps_total = int(office[BIHAR][0]) if BIHAR in office else 0
    if not apps_total:
        if sum(left_out.values()):
            raise UserError("Every application row in the selected files has Model = ULA, and ULA rows are left out. "
                            "There is nothing to report.")
        raise UserError("No application rows were found in the selected files.")
    progress("build", 0.0, "Building the tables…")

    as_of = max(latest.values()) if latest else today
    as_ord = as_of.toordinal()

    # ---- group the dated events by office
    by_office = defaultdict(dict)
    for (i, o), counts in events.items():
        by_office[i][o] = counts
    events.clear()

    # ---- districts named from the other DISCOM: place them by their supply division
    moved = Counter()                              # (discom, district it was placed under) -> applications
    waiting = [k for k in office if k[2].startswith(_STRAY)]
    if waiting:
        main = {}                                  # (discom, division) -> its main district within the DISCOM's own area
        for k in office:
            if k[0] == "Division" and not k[2].startswith(_STRAY) and k[3] != "Not stated" and home_discom(k[2]) == k[1]:
                best_so_far = main.get((k[1], k[3]))
                if best_so_far is None or office[k][0] > office[best_so_far][0]:
                    main[(k[1], k[3])] = k
        id_map = {}
        for old in waiting:
            division_name = old[2][len(_STRAY):]
            found = main.get((old[1], division_name))
            if found:
                placed_in = found[2]
            elif home_discom(division_name) == old[1]:      # no other row to go by, but the division is named after a district
                placed_in = next(tidy_name(n) for n in DEFAULT_TARGETS if place_key(n) == place_key(division_name))
            else:
                placed_in = UNPLACED
            new = (old[0], old[1], placed_in) + old[3:]
            values = office.pop(old)
            target = office[new]
            for i, v in enumerate(values):
                target[i] += v
            if old[0] == "District":
                moved[(old[1], new[2])] += int(values[0])
            id_map[office_id[old]] = key_id(new)   # the old entry stays in office_id so that new numbers never repeat one
            for votes in (circle_votes, parent_votes):
                if old in votes:
                    votes[new].update(votes.pop(old))
        for old_id, new_id in id_map.items():
            days = by_office.pop(old_id, None)
            if days:
                target = by_office[new_id]
                for o, counts in days.items():
                    if o in target:
                        target[o] = [a + b for a, b in zip(target[o], counts)]
                    else:
                        target[o] = counts
        if any(i in id_map for (i, _, _, _) in pend):
            remapped = Counter()
            for (i, stage, who, o), cnt in pend.items():
                remapped[(id_map.get(i, i), stage, who, o)] += cnt
            pend = remapped

    def total(i, first, last):
        out = [0.0] * len(EVENTS)
        days = by_office.get(i)
        if days:
            if len(days) <= last - first:
                for o, counts in days.items():
                    if first <= o <= last:
                        for k, v in enumerate(counts):
                            out[k] += v
            else:
                for o in range(first, last + 1):
                    counts = days.get(o)
                    if counts:
                        for k, v in enumerate(counts):
                            out[k] += v
        return out

    # ---- the output rows
    out = []

    def add(table, key=("", "", "", "", "", ""), period="", item="", item2="", sort="", measure="", value="", text=""):
        if isinstance(value, float):
            # kW and Rs lakh keep their full value in the cell (the sheet shows 1 or 2 decimals). Rounding each
            # cell to 2 decimals first made the totals of different sheets differ in the last digit.
            value = int(value) if value == int(value) else round(value, 8)
        out.append((table,) + tuple(key) + (period, item, item2, sort, measure, value, text))

    show = lambda d: d.strftime("%d %b %Y")
    add("info", measure="extract_date", value=ymd(as_of), text=show(as_of))
    add("info", measure="generated_on", value=ymd(today), text=show(today))
    add("info", measure="app_version", text=f"{APP_NAME} {APP_VERSION}")
    add("info", measure="targets_source",
        text="Built-in targets (review presentation, 27 Sep - 3 Oct 2026)"
             + (f", with changes from {os.path.basename(targets_path)}" if targets_path else ""))
    add("info", measure="section_column", text="found" if section_seen else "not in the export")
    for d in sorted(rows_read):
        k = ("DISCOM", d, "", "", "", "")
        add("info", k, measure="rows_read", value=rows_read[d])
        if d in files:
            add("info", k, measure="source_file", text=files[d])
        if d in latest:
            add("info", k, measure="latest_action_date", value=ymd(latest[d]), text=show(latest[d]))

    # office names: give each sub-division and section its division
    def final_key(key):
        if key[0] == "Sub-division":
            return key[:3] + (parent_votes[key].most_common(1)[0][0],) + key[4:]
        if key[0] == "Section":
            sub = ("Sub-division",) + key[1:5] + ("",)
            return key[:3] + (parent_votes[sub].most_common(1)[0][0],) + key[4:]
        return key

    order = sorted(office, key=lambda k: (LEVELS.index(k[0]), k[1], -office[k][0], k[2], k[3], k[4], k[5]))

    # targets: one per district; DISCOM and Bihar are the sum of their districts
    target_of, best = Counter(), {}
    for k in order:
        if k[0] == "District" and place_key(k[2]) in targets:
            pk = place_key(k[2])
            if pk not in best or office[k][0] > office[best[pk]][0]:
                best[pk] = k
    for pk, k in best.items():
        target_of[k] = targets[pk]
        target_of[("DISCOM", k[1], "", "", "", "")] += targets[pk]
        target_of[BIHAR] += targets[pk]

    for k in order:
        fk = final_key(k)
        for name, v in zip(MEASURES, office[k]):
            if v:
                add("office", fk, measure=name, value=v)
        if target_of.get(k):
            add("office", fk, measure="target", value=target_of[k])
        if k in circle_votes:
            add("office", fk, measure="circle", text=tidy_name(circle_votes[k].most_common(1)[0][0]))

    for wno, (label, kind, n) in enumerate(WINDOWS, 1):
        if cancel():
            raise Cancelled("Cancelled.")
        progress("build", 0.05 + 0.6 * (wno - 1) / len(WINDOWS), f"Building the tables: {label.lower()}")
        buckets = window_buckets(as_of, kind, n)
        start = buckets[0][0]
        length = (as_of - start).days + 1
        s_ord = start.toordinal()
        add("period", period=label, sort=wno, measure="from", value=ymd(start), text=show(start))
        add("period", period=label, sort=wno, measure="to", value=ymd(as_of), text=show(as_of))
        add("period", period=label, sort=wno, measure="days", value=length)
        for bno, (b0, b1, tick) in enumerate(buckets, 1):
            add("bucket", period=label, item=kind, sort=bno, measure="from", value=ymd(b0), text=tick)
            add("bucket", period=label, item=kind, sort=bno, measure="to", value=ymd(b1), text=tick)
        for k in order:
            level = k[0]
            upper = level in ("Bihar", "DISCOM", "District")
            if level == "Section":
                if label not in ("Last 7 days", "Last 30 days"):
                    continue
                names = ["installations"]
            else:
                names = EVENTS if upper else EVENTS_LOWER_LEVELS
            i, fk = office_id[k], final_key(k)
            now = total(i, s_ord, as_ord)
            for name in names:
                if now[E[name]]:
                    add("recent", fk, period=label, sort=wno, measure=name, value=now[E[name]])
            if upper:
                before = total(i, s_ord - length, s_ord - 1)
                for name in EVENTS_WITH_PREVIOUS:
                    if before[E[name]]:
                        add("recent", fk, period=label, sort=wno, measure="prev_" + name, value=before[E[name]])
                for bno, (b0, b1, _) in enumerate(buckets, 1):
                    got = total(i, b0.toordinal(), b1.toordinal())
                    for name in EVENTS:
                        if got[E[name]]:
                            add("series", fk, period=label, sort=bno, measure=name, value=got[E[name]])

    progress("build", 0.68, "Building the tables: monthly and pending")
    late = 0
    for k in order:
        if k[0] not in ("Bihar", "DISCOM", "District"):
            continue
        months = defaultdict(lambda: [0.0] * len(EVENTS))
        for o, counts in by_office.get(office_id[k], {}).items():
            if o > as_ord:
                if k == BIHAR:
                    late += int(sum(counts) - counts[E["installed_kw"]])
                continue
            d = dt.date.fromordinal(o)
            m = months[d.year * 100 + d.month]
            for j, v in enumerate(counts):
                m[j] += v
        for ym in sorted(months):
            for name in EVENTS:
                if months[ym][E[name]]:
                    add("monthly", k, item=f"{ym // 100}M{ym % 100:02d}", sort=ym, measure=name, value=months[ym][E[name]])
    if late:
        checks["Bihar"]["event_after_extract"] = late

    key_of_id = {i: k for k, i in office_id.items()}
    rank = {k: n for n, k in enumerate(order)}
    aged = Counter()
    for (i, stage, who, o), cnt in pend.items():
        age = as_ord - o
        b = 4 if not o else 0 if age <= 30 else 1 if age <= 90 else 2 if age <= 180 else 3
        aged[(i, stage, who, b)] += cnt
    for (i, stage, who, b), cnt in sorted(aged.items(), key=lambda kv: (rank[key_of_id[kv[0][0]]], kv[0][1], kv[0][2], kv[0][3])):
        add("pending", key_of_id[i], period=AGE_BUCKETS[b], item=stage, item2=who, sort=b + 1, measure="applications", value=cnt)

    discoms = sorted(d for d in checks if d != "Bihar")
    for cno, (code, label, note) in enumerate(CHECKS, 1):
        both = sum(checks[d][code] for d in discoms) + checks["Bihar"][code]
        if both:
            add("check", BIHAR, item=label, sort=cno, measure="applications", value=both, text=note)
        for d in discoms:
            if checks[d][code]:
                add("check", ("DISCOM", d, "", "", "", ""), item=label, sort=cno, measure="applications", value=checks[d][code], text=note)

    for (d, column, value), cnt in sorted(audit.items(), key=lambda kv: (kv[0][0], kv[0][1], -kv[1])):
        add("audit", ("DISCOM", d, "", "", "", ""), item=column, item2=value, measure="applications", value=cnt)

    # ---- what the person should know
    facts = {"applications": apps_total, "as_of": as_of, "rows": len(out), "warnings": [],
             "has_section": section_seen, "discoms": sorted(rows_read), "files": files, "rows_read": dict(rows_read),
             "ula_left_out": left_out}
    for d in sorted(rows_read):
        k = ("DISCOM", d, "", "", "", "")
        log(f"{d}: {indian(office[k][0])} applications, {indian(office[k][5])} installed, latest action {show(latest[d]) if d in latest else 'not found'}")
    warn = facts["warnings"].append
    if sum(left_out.values()):
        log("ULA rows left out of the report: " + ", ".join(f"{slot} {indian(c)}" for slot, c in left_out.items()))
    if no_model_column:
        which = "SBPDCL and NBPDCL files have" if len(no_model_column) > 1 else f"{no_model_column[0]} file has"
        warn(f"The {which} no Model column, so ULA rows could not be left out: every row is counted there. "
             "Download the export again with the Model column.")
    for slot, m in model_rows.items():
        blank = m.get("", 0)
        odd = {v: c for v, c in m.items() if v and v.lower() not in MODELS_KNOWN}
        if odd:
            listed = ", ".join(f"'{v}' ({indian(c)} rows)" for v, c in sorted(odd.items(), key=lambda kv: -kv[1])[:6])
            warn(f"{slot}: the Model column has values other than Regular and ULA: {listed}. Only ULA is left out, "
                 "so these rows are counted. Say if they should be left out too.")
        if blank:
            warn(f"{slot}: the Model column is blank in {indian(blank)} rows. Only ULA is left out, so these rows are counted.")
    if set(latest) == {"SBPDCL", "NBPDCL"} and abs((latest["SBPDCL"] - latest["NBPDCL"]).days) > 1:
        warn(f"The two files are from different dates (SBPDCL {show(latest['SBPDCL'])}, NBPDCL {show(latest['NBPDCL'])}). "
             f"Recent periods end on {show(as_of)} for both.")
    for d in discoms:
        if checks[d]["discom_mismatch"]:
            warn(f"{indian(checks[d]['discom_mismatch'])} rows say {d} but came in the other DISCOM's file. Check the files were chosen in the right boxes.")
        if checks[d]["app_date_day1"] > 0.5 * max(rows_read[d], 1):
            warn(f"{d}: most application dates are the 1st of a month. The file may have been re-saved in Excel; use a fresh download.")
    for _, label in DATE_COLUMNS:
        if unread_dates[label]:
            warn(f"{indian(unread_dates[label])} values in '{label}' could not be read as dates (for example '{unread_example[label]}'). "
                 "Those steps are missing from the recent, monthly, weekly and daily tables. Send this line so the date format can be added.")
    month_first = [v for v in unread_example.values() if re.match(r"^\s*\d{1,2}[-/.](1[3-9]|2\d|3[01])[-/.]\d{2,4}", v)]
    if month_first:
        warn(f"Dates in the export look as if the month is written first (for example '{month_first[0]}'). The app reads the day first, "
             "so other dates may have been read wrongly too. Do not rely on the dated tables of this report.")
    for column, (known, rule) in KNOWN_VALUES.items():
        new = Counter()
        for (d, col_name, value), cnt in audit.items():
            if col_name == column and value != "(blank)" and value.lower() not in known:
                new[value] += cnt
        if new:
            listed = ", ".join(f"'{v}' ({indian(n)} rows)" for v, n in new.most_common(6)) + (" and more" if len(new) > 6 else "")
            warn(f"'{column}' has wording that was not in the sample file: {listed}. {rule} Check this is how these should be counted.")
    if not section_seen:
        warn("The exports have no Section column, so no Section KPIs sheet is made.")
    if stray_names:
        named = Counter()
        for (d, name), cnt in stray_names.items():
            named[f"{name} in the {d} file"] += cnt
        placed = ", ".join(f"{name} {indian(n)}" for (d, name), n in moved.most_common(8))
        warn(f"{indian(sum(stray_names.values()))} applications name a district of the other DISCOM ("
             + ", ".join(f"{k}: {indian(v)}" for k, v in named.most_common(8)) + (" and more" if len(named) > 8 else "")
             + f"). They are counted under the main district of their supply division ({placed}"
             + (" and more" if len(moved) > 8 else "") + ").")
    no_target = sorted({k[2] for k in order if k[0] == "District" and not target_of.get(k) and k[2] != UNPLACED})
    if no_target:
        warn("Districts with no target: " + ", ".join(no_target))
    known = {place_key(k) for k in DEFAULT_TARGETS}
    strangers = sorted(name for name in imported if place_key(name) not in best and place_key(name) not in known)
    if ignored_rows:
        warn("Targets file: these rows were ignored because Target is not a whole number: " + ", ".join(ignored_rows[:20]))
    if strangers:
        warn("Targets file: no district with these names was found, so they were not used: " + ", ".join(strangers))
    unused = sorted(target_name[k] for k in set(targets) - set(best) if k in known)
    if unused and len(rows_read) == 2:
        warn("Targets for districts that are not in the files: " + ", ".join(unused))
    progress("build", 0.75, "Tables built")
    return out, as_of, facts


# ----------------------------------------------------------------------------
# The report workbook
# ----------------------------------------------------------------------------
DATA_HEADER = ["table", "level", "discom", "district", "division", "subdivision", "section", "period",
               "item", "item2", "sort", "measure", "value", "text"]
DATA_WIDTHS = [10, 13, 10, 20, 22, 24, 20, 15, 28, 16, 9, 26, 12, 60]


def xl_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def c(value, style="t"):
    """A plain cell."""
    return ("v", value, style)


def fx(formula, cached, style="n"):
    """A formula cell with the value it works out to (so every viewer shows a number)."""
    return ("f", formula, cached, style)


class Sheet:
    def __init__(self, name, widths=None):
        self.name = name
        self.rows = []            # list of lists of cells (None = empty)
        self.widths = widths or []
        self.freeze = None        # top-left cell of the scrolling area, e.g. "C5"
        self.filter = None        # e.g. "A4:Z42"
        self.merges = []
        self.bars = []            # ranges that get data bars (0..1)
        self.heights = {}         # row number -> height
        self.landscape = True

    def add(self, *cells):
        self.rows.append(list(cells))
        return len(self.rows)


KPI_COLS = [  # key, heading, kind, width
    ("target", "Target", "n", 11),
    ("applications", "Applications", "n", 13),
    ("feasibility_approved", "Feasibility approved", "n", 12),
    ("vendor_selected", "Vendor selected", "n", 11),
    ("vendor_accepted", "Vendor accepted", "n", 11),
    ("agreement_uploaded", "Agreement uploaded", "n", 12),
    ("installed", "Installed", "n", 11),
    ("pct_apps", "Installed, % of applications", "p", 13),
    ("pct_target", "Installed, % of target", "p", 12),
    ("balance", "Balance to target", "n", 11),
    ("rank", "Rank by installed", "r", 10),
    ("installed_kw", "Installed kW", "d", 12),
    ("inst7", "Installed, last 7 days", "n", 11),
    ("inst30", "Installed, last 30 days", "n", 11),
    ("at_applied", "Feasibility pending", "n", 12),
    ("at_feasibility", "Vendor not chosen", "n", 11),
    ("at_vendor_selected", "Vendor acceptance pending", "n", 12),
    ("at_vendor_accepted", "Agreement pending", "n", 11),
    ("at_agreement", "Installation pending", "n", 12),
    ("pending_consumer", "Waiting on consumer", "n", 11),
    ("pending_vendor", "Waiting on vendor", "n", 11),
    ("pending_discom", "Waiting on DISCOM", "n", 11),
    ("pending_rec", "Waiting on REC", "n", 11),
    ("inspected", "Inspection approved", "n", 11),
    ("subsidy_released", "Subsidy released", "n", 11),
    ("subsidy_released_lakh", "Subsidy released, ₹ lakh", "d2", 12),
    ("rural", "Rural", "n", 9),
    ("urban", "Urban", "n", 9),
    ("give_up_subsidy", "Subsidy given up", "n", 10),
]
TARGET_ONLY = {"target", "pct_target", "balance", "rank"}
TREND = [("applications", "Applications"), ("vendor_selected", "Vendor selected"), ("agreements", "Agreements"),
         ("installations", "Installations"), ("installed_kw", "Installed kW"), ("subsidy_released", "Subsidy released")]


def xl_date(d):
    """A date as Excel stores it."""
    return (d - dt.date(1899, 12, 30)).days


def build_report(out, as_of, facts):
    """Turns the DATA rows into the report sheets. Every sheet is one plain table:
    a header in row 1, then data rows. Returns a list of Sheet."""
    office = defaultdict(dict)
    recent = defaultdict(dict)
    series = defaultdict(dict)
    monthly = defaultdict(dict)
    pending = defaultdict(lambda: [0] * 5)
    periods, buckets, checks, audit, info = {}, defaultdict(dict), {}, defaultdict(dict), {}
    for (table, level, discom, district, division, subdivision, section, period, item, item2, sort, measure, value, text) in out:
        key = (level, discom, district, division, subdivision, section)
        if table == "office":
            office[key][measure] = text if measure == "circle" else value
        elif table == "recent":
            recent[(key, period)][measure] = value
        elif table == "series":
            series[(key, period, sort)][measure] = value
        elif table == "monthly":
            monthly[(key, sort)][measure] = value
        elif table == "pending":
            pending[(key, item, item2)][sort - 1] += value
        elif table == "period":
            periods.setdefault(period, {"sort": sort})[measure] = value
        elif table == "bucket":
            buckets[period].setdefault(sort, {})[measure] = value
        elif table == "check":
            checks.setdefault((sort, item, text), {})[discom] = value
        elif table == "audit":
            audit[(item, item2)][discom] = value
        elif table == "info":
            info[(discom, measure)] = text if text != "" else value

    discoms = [d for d in ("SBPDCL", "NBPDCL") if d in facts["discoms"]]
    groups = ["Bihar"] + discoms
    gkey = lambda g: BIHAR if g == "Bihar" else ("DISCOM", g, "", "", "", "")
    by_discom = lambda d: discoms.index(d) if d in discoms else 9
    win_labels = sorted(periods, key=lambda p: periods[p]["sort"])
    kinds = {k: kind for k, _, kind, _ in KPI_COLS}

    def base(key):
        o = office.get(key, {})
        b = {k: o.get(k, 0) or 0 for k, _, kind, _ in KPI_COLS if kind in ("n", "d", "d2")}
        b["inst7"] = recent.get((key, "Last 7 days"), {}).get("installations", 0)
        b["inst30"] = recent.get((key, "Last 30 days"), {}).get("installations", 0)
        return b

    def header(sh, heads, height=44):
        sh.add(*[c(h, "h") for h in heads])
        sh.heights[1] = height

    # ------------------------------------------------------------------ KPI sheets
    def kpi_sheet(name, level, id_heads, id_of, with_target):
        sh = Sheet(name)
        cols = [x for x in KPI_COLS if with_target or x[0] not in TARGET_ONLY]
        n_id = len(id_heads)
        L = {k: xl_col(n_id + i + 1) for i, (k, _, _, _) in enumerate(cols)}
        sh.widths = [11] + [22] * (n_id - 1) + [w for _, _, _, w in cols]
        header(sh, id_heads + [h for _, h, _, _ in cols])
        keys = sorted((k for k in office if k[0] == level),
                      key=lambda k: (by_discom(k[1]),) +
                      ((-office[k].get("installed", 0), k[2]) if level == "District" else tuple(id_of(k)[1:])))
        first, last = 2, 1 + len(keys)
        bases = [base(k) for k in keys]
        installed_all = [b["installed"] for b in bases]

        def row_cells(r, b, bold, sums=None):
            s = lambda kind: kind + ("b" if bold else "")
            cells = []
            for k, _, kind, _ in cols:
                if k == "pct_apps":
                    cells.append(fx(f"IFERROR({L['installed']}{r}/{L['applications']}{r},0)",
                                    b["installed"] / b["applications"] if b["applications"] else 0, s("p")))
                elif k == "pct_target":
                    cells.append(fx(f"IFERROR({L['installed']}{r}/{L['target']}{r},0)",
                                    b["installed"] / b["target"] if b["target"] else 0, s("p")))
                elif k == "balance":
                    cells.append(fx(f"MAX(0,{L['target']}{r}-{L['installed']}{r})", max(0, b["target"] - b["installed"]), s("n")))
                elif k == "rank":
                    cells.append(c("", s("t")) if bold else
                                 fx(f"RANK({L['installed']}{r},${L['installed']}${first}:${L['installed']}${last})",
                                    1 + sum(1 for v in installed_all if v > b["installed"]), "n"))
                elif sums:
                    cells.append(fx(sums(L[k]), b[k], s(kind)))
                else:
                    cells.append(c(b[k], s(kind)))
            return cells

        for i, (k, b) in enumerate(zip(keys, bases)):
            sh.add(*([c(v, "t") for v in id_of(k)] + row_cells(first + i, b, False)))
        totals, total_values = {}, {}
        if keys:
            for g in discoms + ["Bihar"]:
                rows_g = [b for k, b in zip(keys, bases) if g == "Bihar" or k[1] == g]
                tb = {k: sum(b[k] for b in rows_g) for k in bases[0]}
                r = len(sh.rows) + 1
                sums = (lambda col: f"SUM({col}{first}:{col}{last})") if g == "Bihar" else \
                    (lambda col, g=g: f'SUMIF($A${first}:$A${last},"{g}",{col}{first}:{col}{last})')
                sh.add(*([c(f"{g} total", "tb")] + [c("", "tb")] * (n_id - 1) + row_cells(r, tb, True, sums)))
                totals[g] = r
                total_values[g] = tb
            sh.filter = f"A1:{xl_col(n_id + len(cols))}{last}"
            sh.bars = [f"{L[k]}{first}:{L[k]}{last}" for k in ("pct_apps", "pct_target") if k in L]
        sh.freeze = f"{xl_col(n_id + 1)}2"
        return sh, L, totals, total_values

    dist, DL, dist_tot, dist_val = kpi_sheet("District KPIs", "District", ["DISCOM", "District"], lambda k: (k[1], k[2]), True)
    divn = kpi_sheet("Division KPIs", "Division", ["DISCOM", "District", "Circle", "Division"],
                     lambda k: (k[1], k[2], office[k].get("circle", ""), k[3]), False)[0]
    subd = kpi_sheet("Sub-division KPIs", "Sub-division", ["DISCOM", "District", "Division", "Sub-division"],
                     lambda k: (k[1], k[2], k[3], k[4]), False)[0]
    sect = kpi_sheet("Section KPIs", "Section", ["DISCOM", "District", "Division", "Sub-division", "Section"],
                     lambda k: (k[1], k[2], k[3], k[4], k[5]), False)[0]     # left out when the exports have no Section column

    # ------------------------------------------------------------------ Headline KPIs and Funnel
    gb = {g: dict(dist_val[g]) if g in dist_val else base(gkey(g)) for g in groups}   # the District KPIs totals, as shown there
    for g in groups:
        b = gb[g]
        b.setdefault("target", office.get(gkey(g), {}).get("target", 0))
        b["pct_apps"] = b["installed"] / b["applications"] if b["applications"] else 0
        b["pct_target"] = b["installed"] / b["target"] if b["target"] else 0
        b["balance"] = max(0, b["target"] - b["installed"])
    head = Sheet("Headline KPIs", [34] + [15] * len(groups))
    header(head, ["KPI"] + groups, 24)
    head_rows = {}
    for k, label in [("target", "Installation target"), ("applications", "Applications submitted"),
                     ("feasibility_approved", "Feasibility approved"), ("vendor_selected", "Vendor selected"),
                     ("vendor_accepted", "Vendor accepted"), ("agreement_uploaded", "Agreement uploaded"),
                     ("installed", "Installations completed"), ("pct_apps", "Installed, % of applications"),
                     ("pct_target", "Installed, % of target"), ("balance", "Balance to target"),
                     ("installed_kw", "Installed capacity, kW"), ("inst7", "Installed in the last 7 days"),
                     ("inst30", "Installed in the last 30 days"), ("inspected", "Inspection approved"),
                     ("subsidy_released", "Subsidy released, applications"), ("subsidy_released_lakh", "Subsidy released, ₹ lakh"),
                     ("pending_consumer", "Waiting on consumer"), ("pending_vendor", "Waiting on vendor"),
                     ("pending_discom", "Waiting on DISCOM"), ("pending_rec", "Waiting on REC")]:
        head_rows[k] = len(head.rows) + 1
        head.add(c(label, "t"), *[fx(f"'District KPIs'!{DL[k]}{dist_tot[g]}", gb[g][k], kinds[k]) if g in dist_tot
                                  else c(gb[g][k], kinds[k]) for g in groups])
    head.freeze = "B2"

    fun = Sheet("Funnel", [28] + [13, 14, 14] * len(groups))
    heads = ["Step"]
    for g in groups:
        heads += [g, f"{g}, % of applications", f"{g}, % of step before"]
    header(fun, heads)
    steps = [("applications", "Applications submitted"), ("feasibility_approved", "Feasibility approved"),
             ("vendor_selected", "Vendor selected"), ("vendor_accepted", "Vendor accepted"),
             ("agreement_uploaded", "Agreement uploaded"), ("installed", "Installations completed")]
    for si, (k, label) in enumerate(steps):
        r = len(fun.rows) + 1
        cells = [c(label, "t")]
        for gi, g in enumerate(groups):
            a = xl_col(2 + 3 * gi)
            v, v0 = gb[g][k], gb[g]["applications"]
            vp = gb[g][steps[si - 1][0]] if si else v
            cells += [fx(f"'Headline KPIs'!{xl_col(2 + gi)}{head_rows[k]}", v, "n"),
                      fx(f"IFERROR({a}{r}/{a}$2,0)", v / v0 if v0 else 0, "p"),
                      fx(f"IFERROR({a}{r}/{a}{r - 1},0)", v / vp if vp else 0, "p") if si else c("", "t")]
        fun.add(*cells)
    fun.freeze = "B2"

    # ------------------------------------------------------------------ Recent progress
    rec_cols = [("applications", "Applications"), ("vendor_selected", "Vendor selected"), ("agreements", "Agreements"),
                ("installations", "Installations"), ("prev_installations", "Installations, period before")]
    rec = Sheet("Recent progress", [11, 22] + [13] * (len(win_labels) * len(rec_cols)))
    header(rec, ["DISCOM", "District"] + [f"{label}: {h}" for label in win_labels for _, h in rec_cols], 58)
    dkeys = sorted((k for k in office if k[0] == "District"),
                   key=lambda k: (by_discom(k[1]), -office[k].get("installed", 0), k[2]))
    first, last = 2, 1 + len(dkeys)
    rec_tot = {}
    for k in dkeys:
        rec.add(c(k[1], "t"), c(k[2], "t"), *[c(recent.get((k, label), {}).get(m, 0), "n") for label in win_labels for m, _ in rec_cols])
    if dkeys:
        for g in discoms + ["Bihar"]:
            r = len(rec.rows) + 1
            cells = [c(f"{g} total", "tb"), c("", "tb")]
            for w, label in enumerate(win_labels):
                for j, (m, _) in enumerate(rec_cols):
                    col = xl_col(3 + w * len(rec_cols) + j)
                    f = f"SUM({col}{first}:{col}{last})" if g == "Bihar" else f'SUMIF($A${first}:$A${last},"{g}",{col}{first}:{col}{last})'
                    cells.append(fx(f, recent.get((gkey(g), label), {}).get(m, 0), "nb"))
            rec.add(*cells)
            rec_tot[g] = r
        rec.filter = f"A1:{xl_col(2 + len(win_labels) * len(rec_cols))}{last}"
    rec.freeze = "C2"

    rsum = Sheet("Recent summary", [11, 16, 13, 13, 8] + [13] * (len(rec_cols) + 1))
    header(rsum, ["Scope", "Period", "From", "To", "Days"] + [h for _, h in rec_cols] + ["Change in installations"])
    for g in groups:
        for w, label in enumerate(win_labels):
            p, r = periods[label], len(rsum.rows) + 1
            got = recent.get((gkey(g), label), {})
            cells = [c(g, "t"), c(label, "t"), c(xl_date(from_ymd(p["from"])), "date"), c(xl_date(from_ymd(p["to"])), "date"), c(p["days"], "n")]
            for j, (m, _) in enumerate(rec_cols):
                v = got.get(m, 0)
                cells.append(fx(f"'Recent progress'!{xl_col(3 + w * len(rec_cols) + j)}{rec_tot[g]}", v, "n") if g in rec_tot else c(v, "n"))
            cells.append(fx(f"I{r}-J{r}", got.get("installations", 0) - got.get("prev_installations", 0), "n"))
            rsum.add(*cells)
    rsum.filter = f"A1:K{max(1, len(rsum.rows))}"
    rsum.freeze = "C2"

    # ------------------------------------------------------------------ Trends
    def trend_sheet(name, first_head, labels, value_of, cumulative):
        sh = Sheet(name, [26 if isinstance(labels[0] if labels else "", str) else 13] + [13] * (len(groups) * len(TREND)) + ([15] if cumulative else []))
        header(sh, [first_head] + [f"{g}: {h}" for g in groups for _, h in TREND] + (["Bihar: cumulative installations"] if cumulative else []), 58)
        running = 0
        for i, label in enumerate(labels):
            r = 2 + i
            cells = [c(label, "t") if isinstance(label, str) else c(xl_date(label), "date")]
            for gi, g in enumerate(groups):
                for j, (m, _) in enumerate(TREND):
                    style = "d" if m == "installed_kw" else "n"
                    v = round(sum(value_of(d, i, m) for d in discoms), 8) if g == "Bihar" and discoms else value_of(g, i, m)
                    if g == "Bihar" and discoms:
                        parts = "+".join(f"{xl_col(2 + (di + 1) * len(TREND) + j)}{r}" for di in range(len(discoms)))
                        cells.append(fx(parts, v, style))
                    else:
                        cells.append(c(v, style))
            if cumulative:
                running += value_of("Bihar", i, "installations")
                col = xl_col(2 + [m for m, _ in TREND].index("installations"))
                cells.append(fx(f"SUM({col}$2:{col}{r})", running, "n"))
            sh.add(*cells)
        sh.freeze = "B2"
        return sh

    yms = sorted({ym for (k, ym) in monthly if k == BIHAR})
    months = []
    if yms:
        y, m = divmod(yms[0], 100)
        while y * 100 + m <= as_of.year * 100 + as_of.month:
            months.append(y * 100 + m)
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    mon = trend_sheet("Monthly trend", "Month", [dt.date(ym // 100, ym % 100, 1).strftime("%b %Y") for ym in months],
                      lambda g, i, m: monthly.get((gkey(g), months[i]), {}).get(m, 0), True)

    def bucket_sheet(name, label, head_text, as_dates):
        bks = [buckets[label][n] for n in sorted(buckets.get(label, {}))]
        names = [from_ymd(b["from"]) if as_dates else f"{from_ymd(b['from']):%d %b} to {from_ymd(b['to']):%d %b %Y}" for b in bks]
        return trend_sheet(name, head_text, names, lambda g, i, m: series.get((gkey(g), label, i + 1), {}).get(m, 0), False)

    wk = bucket_sheet("Weekly trend", "Last 12 weeks", "Week", False)
    dy = bucket_sheet("Daily trend", "Last 30 days", "Date", True)

    # ------------------------------------------------------------------ Pendency
    pen = Sheet("Pendency", [11, 22, 22, 14, 11, 11, 11, 11, 11, 11])
    header(pen, ["DISCOM", "District", "Next step", "Action by"] + AGE_BUCKETS + ["Total"], 32)
    prow = sorted(((k, st, who, v) for (k, st, who), v in pending.items() if k[0] == "District"),
                  key=lambda x: (by_discom(x[0][1]), x[0][2], -sum(x[3]), x[1], x[2]))
    for k, st, who, v in prow:
        r = len(pen.rows) + 1
        pen.add(c(k[1], "t"), c(k[2], "t"), c(st, "t"), c(who, "t"), *[c(x, "n") for x in v], fx(f"SUM(E{r}:I{r})", sum(v), "nb"))
    pen_first, pen_last = 2, max(2, 1 + len(prow))
    if prow:
        pen.filter = f"A1:J{pen_last}"
    pen.freeze = "C2"

    act = Sheet("Who has to act", [11, 22, 14, 11, 11, 11, 11, 11, 11])
    header(act, ["Scope", "Next step", "Action by"] + AGE_BUCKETS + ["Total"], 32)
    rng = lambda col: f"Pendency!${col}${pen_first}:${col}${pen_last}"
    for g in groups:
        for (k, st, who), v in sorted(((kk, vv) for kk, vv in pending.items() if kk[0] == gkey(g)), key=lambda kv: -sum(kv[1])):
            r = len(act.rows) + 1
            scope = "" if g == "Bihar" else f",{rng('A')},$A{r}"
            act.add(c(g, "t"), c(st, "t"), c(who, "t"),
                    *[fx(f"SUMIFS({rng(xl_col(5 + j))},{rng('C')},$B{r},{rng('D')},$C{r}{scope})", x, "n") for j, x in enumerate(v)],
                    fx(f"SUM(D{r}:H{r})", sum(v), "nb"))
    if len(act.rows) > 1:
        act.filter = f"A1:I{len(act.rows)}"
    act.freeze = "D2"

    sheets = [head, fun, dist, divn, subd] + ([sect] if facts["has_section"] else []) + [rec, rsum, mon, wk, dy, pen, act]
    return sheets


# ----------------------------------------------------------------------------
# Excel writer (no extra packages)
# ----------------------------------------------------------------------------
_INDIAN = r"[&gt;=10000000]##\,##\,##\,##0{0};[&gt;=100000]##\,##\,##0{0};##,##0{0}"
NUMFMTS = {164: _INDIAN.format(""), 165: _INDIAN.format(".0"), 166: "0.0%", 167: _INDIAN.format(".00"), 168: "dd mmm yyyy"}
FONTS = ['<font><sz val="10"/><name val="Arial"/></font>',
         '<font><b/><sz val="10"/><name val="Arial"/></font>',
         '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>',
         '<font><b/><sz val="14"/><color rgb="FF0E2A47"/><name val="Arial"/></font>',
         '<font><sz val="9"/><color rgb="FF5B6B7B"/><name val="Arial"/></font>',
         '<font><b/><sz val="11"/><color rgb="FF0E2A47"/><name val="Arial"/></font>']
FILLS = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>',
         '<fill><patternFill patternType="solid"><fgColor rgb="FF0E2A47"/><bgColor indexed="64"/></patternFill></fill>',
         '<fill><patternFill patternType="solid"><fgColor rgb="FFE3EAF2"/><bgColor indexed="64"/></patternFill></fill>']
_side = '<{0} style="thin"><color rgb="FFC8D2DD"/></{0}>'
BORDERS = ['<border><left/><right/><top/><bottom/><diagonal/></border>',
           "<border>" + "".join(_side.format(s) for s in ("left", "right", "top", "bottom")) + "<diagonal/></border>"]
# name: (font, fill, border, number format, alignment)
STYLES = {
    "plain": (0, 0, 0, 0, ""), "bold": (1, 0, 0, 0, ""),
    "date": (0, 0, 1, 168, ""), "datel": (0, 0, 1, 168, '<alignment horizontal="left"/>'),
    "nl": (0, 0, 1, 164, '<alignment horizontal="left"/>'),
    "h": (2, 2, 1, 0, '<alignment horizontal="center" vertical="center" wrapText="1"/>'),
    "t": (0, 0, 1, 0, ""), "n": (0, 0, 1, 164, ""), "d": (0, 0, 1, 165, ""), "d2": (0, 0, 1, 167, ""), "p": (0, 0, 1, 166, ""),
    "tb": (1, 3, 1, 0, ""), "nb": (1, 3, 1, 164, ""), "db": (1, 3, 1, 165, ""), "d2b": (1, 3, 1, 167, ""), "pb": (1, 3, 1, 166, ""),
}
STYLE_ID = {name: i for i, name in enumerate(STYLES)}
_bad_xml = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _styles_xml():
    xfs = []
    for font, fill, border, fmt, align in STYLES.values():
        attrs = f'numFmtId="{fmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0"'
        attrs += ' applyFont="1"' * bool(font) + ' applyFill="1"' * bool(fill) + ' applyBorder="1"' * bool(border) + ' applyNumberFormat="1"' * bool(fmt)
        xfs.append(f"<xf {attrs} applyAlignment=\"1\">{align}</xf>" if align else f"<xf {attrs}/>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<numFmts count="{len(NUMFMTS)}">' + "".join(f'<numFmt numFmtId="{i}" formatCode="{code}"/>' for i, code in NUMFMTS.items()) + "</numFmts>"
            f'<fonts count="{len(FONTS)}">' + "".join(FONTS) + "</fonts>"
            f'<fills count="{len(FILLS)}">' + "".join(FILLS) + "</fills>"
            f'<borders count="{len(BORDERS)}">' + "".join(BORDERS) + "</borders>"
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{len(xfs)}">' + "".join(xfs) + "</cellXfs>"
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            "</styleSheet>")


def _text(v):
    return escape(_bad_xml.sub("", str(v)))


def _cell_xml(ref, cell):
    if cell is None:
        return ""
    if cell[0] == "f":
        _, formula, cached, style = cell
        s = f' s="{STYLE_ID[style]}"'
        if isinstance(cached, str):
            return f'<c r="{ref}"{s} t="str"><f>{_text(formula)}</f><v>{_text(cached)}</v></c>'
        return f'<c r="{ref}"{s}><f>{_text(formula)}</f><v>{round(cached, 10)}</v></c>'
    _, v, style = cell
    s = f' s="{STYLE_ID[style]}"' if STYLE_ID[style] else ""
    if v == "" or v is None:
        return f'<c r="{ref}"{s}/>' if s else ""
    if isinstance(v, (int, float)):
        return f'<c r="{ref}"{s}><v>{v}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{_text(v)}</t></is></c>'


def _sheet_head(sh, n_rows, n_cols):
    pane = ""
    if sh.freeze:
        m = re.match(r"([A-Z]+)(\d+)", sh.freeze)
        xs, ys = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(m.group(1)))) - 1, int(m.group(2)) - 1
        which = "bottomRight" if xs and ys else "bottomLeft" if ys else "topRight"
        split = (' xSplit="%d"' % xs if xs else "") + (' ySplit="%d"' % ys if ys else "")
        pane = '<pane%s topLeftCell="%s" activePane="%s" state="frozen"/>' % (split, sh.freeze, which)
    widths = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(sh.widths, 1))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            + ('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' if sh.landscape else "")
            + f'<dimension ref="A1:{xl_col(max(n_cols, 1))}{max(n_rows, 1)}"/>'
            f'<sheetViews><sheetView workbookViewId="0">{pane}</sheetView></sheetViews>'
            + (f"<cols>{widths}</cols>" if widths else "") + "<sheetData>")


def _sheet_tail(sh):
    x = "</sheetData>"
    if sh.filter:
        x += f'<autoFilter ref="{sh.filter}"/>'
    if sh.merges:
        x += f'<mergeCells count="{len(sh.merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in sh.merges) + "</mergeCells>"
    for n, ref in enumerate(sh.bars, 1):
        x += (f'<conditionalFormatting sqref="{ref}"><cfRule type="dataBar" priority="{n}"><dataBar>'
              '<cfvo type="num" val="0"/><cfvo type="num" val="1"/><color rgb="FF8CCFA8"/></dataBar></cfRule></conditionalFormatting>')
    if sh.landscape:
        x += ('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
              '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
    return x + "</worksheet>"


def write_workbook(path, sheets, progress=None, cancel=None):
    total_rows = sum(len(sh.rows) + len(getattr(sh, "raw", ())) for sh in sheets) or 1
    if any(len(sh.rows) + len(getattr(sh, "raw", ())) > 1000000 for sh in sheets):
        raise UserError("The report is too large for one Excel sheet.")
    quote = lambda name: "'" + name.replace("'", "''") + "'"
    absolute = lambda ref: re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", ref)
    names = "".join('<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">%s!%s</definedName>'
                    % (i, escape(quote(sh.name)), absolute(sh.filter)) for i, sh in enumerate(sheets) if sh.filter)
    parts = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                      for i in range(1, len(sheets) + 1))
            + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>' + "".join(f'<sheet name="{escape(sh.name)}" sheetId="{i}" r:id="rId{i}"/>' for i, sh in enumerate(sheets, 1)) + '</sheets>'
            + (f"<definedNames>{names}</definedNames>" if names else "")
            + '</workbook>',
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                      for i in range(1, len(sheets) + 1))
            + f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>',
        "xl/styles.xml": _styles_xml(),
    }
    tmp = path + ".tmp"
    done = 0
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, xml in parts.items():
                z.writestr(name, xml)
            for i, sh in enumerate(sheets, 1):
                raw = getattr(sh, "raw", ())
                n_cols = max([len(r) for r in sh.rows] + [len(DATA_HEADER) if raw else 1])
                with z.open(f"xl/worksheets/sheet{i}.xml", "w") as part:
                    part.write(_sheet_head(sh, len(sh.rows) + len(raw), n_cols).encode("utf-8"))
                    chunk = []
                    for r, row in enumerate(sh.rows, 1):
                        ht = f' ht="{sh.heights[r]}" customHeight="1"' if r in sh.heights else ""
                        chunk.append(f'<row r="{r}"{ht}>' + "".join(_cell_xml(f"{xl_col(ci)}{r}", cell) for ci, cell in enumerate(row, 1)) + "</row>")
                    letters = [xl_col(ci) for ci in range(1, len(DATA_HEADER) + 1)]
                    for r, row in enumerate(raw, len(sh.rows) + 1):
                        bits = []
                        for letter, v in zip(letters, row):
                            if v == "" or v is None:
                                continue
                            if isinstance(v, (int, float)):
                                bits.append(f'<c r="{letter}{r}"><v>{v}</v></c>')
                            else:
                                bits.append(f'<c r="{letter}{r}" t="inlineStr"><is><t xml:space="preserve">{_text(v)}</t></is></c>')
                        chunk.append(f'<row r="{r}">' + "".join(bits) + "</row>")
                        if len(chunk) >= 5000:
                            part.write("".join(chunk).encode("utf-8"))
                            chunk = []
                            if cancel and cancel():
                                raise Cancelled("Cancelled.")
                            if progress:
                                progress("write", (done + r) / total_rows, f"Writing Excel: {indian(done + r)} of {indian(total_rows)} rows")
                    part.write(("".join(chunk) + _sheet_tail(sh)).encode("utf-8"))
                done += len(sh.rows) + len(raw)
        os.replace(tmp, path)
    except PermissionError:
        raise UserError(f"Could not write {os.path.basename(path)}. If it is open in Excel, close it and try again.")
    except OSError as exc:
        raise UserError(f"Could not write {os.path.basename(path)}: {exc.strerror or exc}. Check the disk has free space.")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def make_report(sbpdcl, nbpdcl, out_dir, targets_path=None, progress=None, log=None, cancel=None):
    """Runs the whole job. Returns (path of the Excel file, facts)."""
    log = log or (lambda msg: None)
    inputs = [(label, p) for label, p in (("SBPDCL", sbpdcl), ("NBPDCL", nbpdcl)) if p]
    if not inputs:
        raise UserError("Choose at least one export file.")
    if sbpdcl and nbpdcl and os.path.abspath(sbpdcl) == os.path.abspath(nbpdcl):
        raise UserError("The same file is chosen for SBPDCL and NBPDCL.")
    if not out_dir:
        raise UserError("Choose an output folder.")
    try:                                          # fail now, not after minutes of reading
        os.makedirs(out_dir, exist_ok=True)
        probe = os.path.join(out_dir, ".pmsg_write_test")
        with open(probe, "w"):
            pass
        os.remove(probe)
    except OSError as exc:
        raise UserError(f"Cannot write to the output folder:\n{out_dir}\n{exc.strerror or exc}")
    where = {}
    try:
        rows, as_of, facts = summarise(inputs, targets_path or None, progress, log, cancel, where)
        if progress:
            progress("build", 0.8, "Laying out the report sheets")
        sheets = build_report(rows, as_of, facts)
        path = os.path.join(out_dir, f"PMSG_Report_{as_of:%Y-%m-%d}.xlsx")
        if progress:
            progress("build", 1.0, "Report laid out")
            progress("write", 0.0, "Writing Excel…")
        write_workbook(path, sheets, progress, cancel)
    except csv.Error as exc:
        line = getattr(where.get("reader"), "line_num", "?")
        raise UserError(f"{where.get('file', 'The file')} could not be read near line {line}: {exc}.\n"
                        "Download the export again; the file may be damaged or not a CSV.")
    except MemoryError:
        raise UserError("The computer ran out of memory. Close other programs and try again, or run one DISCOM at a time.")
    facts["path"] = path
    facts["sheets"] = [sh.name for sh in sheets]
    for row in rows:                              # data checks are reported here, not in the workbook
        if row[0] == "check" and row[1] == "Bihar":
            log(f"Data check: {row[8]}: {indian(row[12])} applications")
    log(f"Extract date {as_of:%d %b %Y}: {indian(facts['applications'])} applications in all")
    for w in facts["warnings"]:
        log("CHECK: " + w)
    log(f"Saved {os.path.basename(path)}: {len(sheets)} sheets ({', '.join(facts['sheets'])})")
    return path, facts


# ----------------------------------------------------------------------------
# Sending the report to the dashboard (the Apps Script web app)
# ----------------------------------------------------------------------------
UPLOAD_TIMEOUT = 300                              # seconds; loading a new report into the sheet takes about a minute


def check_dashboard_link(link):
    """State of the dashboard link for the window: (state, message)."""
    link = (link or "").strip()
    if not link:
        return "empty", "Not set: the report is only saved on this computer"
    if not link.lower().startswith("https://script.google.com/"):
        return "bad", "This is not the dashboard's web app link. It starts with https://script.google.com/"
    if link.split("?")[0].rstrip("/").endswith("/dev"):
        return "bad", "This is the test link (ends with /dev). Use the web app link that ends with /exec"
    if not link.split("?")[0].rstrip("/").endswith("/exec"):
        return "bad", "The web app link ends with /exec. Copy it again from Deploy > Manage deployments"
    return "ok", "Dashboard link set"


def dashboard_call(link, payload, timeout=None):
    """Sends one request to the dashboard's script and returns its answer (a dict with ok = True).
    Any failure is raised as a UserError that says what to do."""
    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(link.strip(), data=data, headers={"Content-Type": "text/plain;charset=utf-8",
                                                                      "User-Agent": "PMSG-Report-Maker/" + APP_VERSION})
    setup = ("In the Apps Script project: add Upload.gs, then Deploy > Manage deployments > pencil > "
             "Version: New version, Who has access: Anyone > Deploy.")
    try:
        with urllib.request.urlopen(request, timeout=timeout or UPLOAD_TIMEOUT) as reply:   # Google answers with a redirect; it is followed
            final = reply.geturl()
            text = reply.read(2_000_000).decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        if exc.code in (401, 403):
            raise UserError(f"Google refused the request (error {exc.code}). The web app must be open to Anyone. " + setup)
        if exc.code == 404:
            raise UserError("Google does not know this dashboard link (error 404). Copy the web app link again from "
                            "Deploy > Manage deployments.")
        raise UserError(f"Google answered with error {exc.code}. Try again in a few minutes.")
    except (socket.timeout, TimeoutError):
        raise UserError("Google did not answer in time. The report may still have been loaded: open the dashboard and "
                        "check its date before trying again.")
    except urllib.error.URLError as exc:
        if isinstance(exc.reason, (socket.timeout, TimeoutError)):
            raise UserError("Google did not answer in time. The report may still have been loaded: open the dashboard and "
                            "check its date before trying again.")
        raise UserError(f"Could not reach Google ({exc.reason}). Check the internet connection, then try again.")
    except (OSError, ValueError) as exc:
        raise UserError(f"Could not reach Google ({exc}). Check the dashboard link and the internet connection.")
    try:
        answer = json.loads(text)
    except ValueError:
        answer = None
    if not isinstance(answer, dict) or "ok" not in answer:
        if "accounts.google.com" in final or "ServiceLogin" in text or "Sign in" in text[:3000]:
            raise UserError("Google asked for a sign-in, so the web app is not open to Anyone. " + setup)
        if "doPost" in text:
            raise UserError("The dashboard's script has no upload part yet. " + setup)
        raise UserError("The link did not answer like the dashboard's script. Check the link. " + setup)
    if not answer.get("ok"):
        raise UserError(str(answer.get("error") or "The dashboard's script refused the upload."))
    return answer


def test_dashboard(link, password):
    """Checks the link and password without sending a report. Returns a sentence to show."""
    state, text = check_dashboard_link(link)
    if state != "ok":
        raise UserError(text)
    if not password:
        raise UserError("No upload password was given (the one set in Upload.gs).")
    answer = dashboard_call(link, {"password": password, "action": "check"}, timeout=60)
    return f"Connected: folder \"{answer.get('folder', '?')}\", sheet \"{answer.get('sheet', '?')}\""


def upload_report(link, password, path, log=None):
    """Sends the Excel report to the dashboard and has it loaded into the Google Sheet. Returns the script's answer."""
    log = log or (lambda msg: None)
    state, text = check_dashboard_link(link)
    if state != "ok":
        raise UserError(text)
    if not password:
        raise UserError("No upload password was given (the one set in Upload.gs).")
    with open(path, "rb") as fh:
        blob = fh.read()
    log(f"Sending {os.path.basename(path)} ({len(blob) / 1e3:,.0f} KB) to the dashboard…")
    answer = dashboard_call(link, {"password": password, "action": "upload", "name": os.path.basename(path),
                                   "data": base64.b64encode(blob).decode("ascii")})
    updated, left = answer.get("updated") or [], answer.get("left") or []
    log(f"Dashboard updated: {len(updated)} tabs loaded" + (f" ({', '.join(updated)})" if updated else ""))
    if answer.get("added"):
        log("  New tabs added to the sheet: " + ", ".join(answer["added"]))
    if left:
        log("  Tabs not in the report, left as they were: " + ", ".join(left))
    if answer.get("replaced"):
        log("  An earlier report of the same date in the Drive folder was moved to the bin")
    return answer


# ----------------------------------------------------------------------------
# The window
# ----------------------------------------------------------------------------
SETTINGS = os.path.join(os.path.expanduser("~"), ".pmsg_report_maker.json")
STEPS_SHOWN = [("SBPDCL", "SBPDCL"), ("NBPDCL", "NBPDCL"), ("build", "Build"), ("write", "Excel"), ("upload", "Dashboard")]


def check_export(path):
    """State of a chosen export file for the window: (state, message)."""
    if not path:
        return "empty", "Not chosen"
    if not os.path.isfile(path):
        return "bad", "File not found"
    if os.path.splitext(path)[1].lower() in (".xlsx", ".xls", ".xlsm"):
        return "bad", "This is an Excel file. Choose the CSV as downloaded from the portal."
    size = os.path.getsize(path)
    if size == 0:
        return "bad", "This file is empty"
    return "ok", f"{os.path.basename(path)}  ·  {size / 1e6:,.1f} MB"


def run_gui():
    import time
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import filedialog, messagebox, simpledialog, ttk

    BG, CARD, LINE, FIELD = "#EEF2F6", "#FFFFFF", "#D3DCE6", "#F6F8FB"
    INK, MUTED, NAVY = "#12263A", "#5B6B7B", "#0E2A47"
    SUN, SUN_HOT, SUN_DEEP = "#F5A623", "#FFBA47", "#D98B0B"
    GREEN, RED, AMBER = "#1E8E5A", "#C0392B", "#9A6700"
    CHIP = {"SBPDCL": "#0F766E", "NBPDCL": "#4338CA", "TARGETS": "#B45309", "OUTPUT": "#475569",
            "DASHBOARD": "#1D4ED8", "PASSWORD": "#64748B"}

    try:
        with open(SETTINGS, encoding="utf-8") as fh:
            saved = json.load(fh)
        if not isinstance(saved, dict):
            saved = {}
    except Exception:
        saved = {}

    try:                                        # sharp text on Windows; sizes below are scaled to match
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    if os.environ.get("PMSG_TK_SCALING"):       # for testing other screen scalings
        root.tk.call("tk", "scaling", float(os.environ["PMSG_TK_SCALING"]))
    scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
    px = lambda v: int(round(v * scale))
    root.title(APP_NAME)
    root.configure(bg=BG)
    width = min(px(1120), root.winfo_screenwidth() - px(40))
    height = min(px(680), root.winfo_screenheight() - px(90))
    root.geometry(f"{width}x{height}")
    root.minsize(min(px(980), width), min(px(600), height))

    families = set(tkfont.families())
    ui = next((f for f in ("Segoe UI", "Helvetica Neue", "Noto Sans", "DejaVu Sans", "Arial") if f in families), "TkDefaultFont")
    mono = next((f for f in ("Cascadia Mono", "Consolas", "DejaVu Sans Mono", "Courier New") if f in families), "TkFixedFont")
    F = lambda size, weight="normal": (ui, size, weight)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure(".", font=F(10), background=CARD, foreground=INK)
    style.configure("TEntry", fieldbackground=FIELD, foreground=INK, bordercolor=LINE, lightcolor=FIELD, darkcolor=FIELD, padding=px(6))
    style.map("TEntry", bordercolor=[("focus", NAVY)])
    style.configure("Primary.TButton", font=F(11, "bold"), background=SUN, foreground=INK, padding=(px(24), px(10)),
                    borderwidth=0, focuscolor=SUN_DEEP, relief="flat")
    style.map("Primary.TButton", background=[("disabled", "#E1E6EC"), ("pressed", SUN_DEEP), ("active", SUN_HOT)],
              foreground=[("disabled", "#9AA6B2")])
    style.configure("Ghost.TButton", background=CARD, foreground=INK, padding=(px(13), px(8)), bordercolor=LINE,
                    lightcolor=CARD, darkcolor=CARD, focuscolor=NAVY, relief="flat")
    style.map("Ghost.TButton", background=[("disabled", "#F3F5F8"), ("pressed", "#DDE5EE"), ("active", "#EDF2F7")],
              foreground=[("disabled", "#A7B1BC")])
    style.configure("Stop.TButton", background=CARD, foreground=RED, padding=(px(13), px(8)), bordercolor="#E7B8B2",
                    lightcolor=CARD, darkcolor=CARD, focuscolor=RED, relief="flat")
    style.map("Stop.TButton", background=[("disabled", "#F3F5F8"), ("pressed", "#F6D5D0"), ("active", "#FBEAE8")],
              foreground=[("disabled", "#A7B1BC")], bordercolor=[("disabled", LINE)])
    style.configure("Soft.TButton", background="#E7EDF4", foreground=INK, padding=(px(12), px(7)), bordercolor="#E7EDF4",
                    lightcolor="#E7EDF4", darkcolor="#E7EDF4", focuscolor=NAVY, relief="flat")
    style.map("Soft.TButton", background=[("disabled", "#F1F4F8"), ("pressed", "#CBD6E2"), ("active", "#D8E1EB")],
              foreground=[("disabled", "#A7B1BC")])
    style.configure("Small.Soft.TButton", padding=(px(9), px(3)), font=F(9))
    style.configure("Run.Horizontal.TProgressbar", troughcolor="#E2E8EF", background=GREEN, bordercolor="#E2E8EF",
                    lightcolor=GREEN, darkcolor=GREEN, thickness=px(14))
    style.configure("Vertical.TScrollbar", background="#3D556F", troughcolor="#1B3048", bordercolor="#1B3048", arrowcolor="#B7C6D6",
                    lightcolor="#3D556F", darkcolor="#3D556F", gripcount=0, relief="flat", arrowsize=px(12))
    style.map("Vertical.TScrollbar", background=[("active", "#52708F")])

    # ---- header band
    import math
    head = tk.Canvas(root, height=px(84), bg=NAVY, highlightthickness=0)
    head.pack(fill="x")
    cx, cy = px(50), px(42)
    for k in range(12):
        a = math.pi * k / 6
        head.create_line(cx + px(18) * math.cos(a), cy + px(18) * math.sin(a), cx + px(26) * math.cos(a), cy + px(26) * math.sin(a),
                         fill=SUN, width=max(2, px(3)), capstyle="round")
    head.create_oval(cx - px(13), cy - px(13), cx + px(13), cy + px(13), fill=SUN, outline="")
    head.create_text(px(92), px(31), anchor="w", text=APP_NAME, fill="#FFFFFF", font=F(17, "bold"))
    head.create_text(px(93), px(58), anchor="w", fill="#B7C6D6", font=F(10),
                     text="Turns the SBPDCL and NBPDCL portal exports into one Excel report, ready for PPTs and reports")
    tag = head.create_text(0, px(31), anchor="e", text=f"Version {APP_VERSION}", fill="#8FA4BA", font=F(9))
    head.bind("<Configure>", lambda e: head.coords(tag, e.width - px(22), px(31)))
    tk.Frame(root, height=px(4), bg=SUN).pack(fill="x")

    foot = tk.Label(root, bg=BG, fg=MUTED, font=F(9), anchor="w",
                    text="Regular applications only: ULA rows are left out. Counts only: no consumer name, number, mobile, email or bank detail is written to the Excel file.")
    foot.pack(side="bottom", fill="x", padx=px(20), pady=(0, px(8)))
    body = tk.Frame(root, bg=BG)
    body.pack(fill="both", expand=True, padx=px(18), pady=(px(14), px(6)))
    body.columnconfigure(0, weight=5, uniform="cols")
    body.columnconfigure(1, weight=4, uniform="cols")
    body.rowconfigure(0, weight=1)
    left = tk.Frame(body, bg=BG)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, px(14)))
    right = tk.Frame(body, bg=BG)
    right.grid(row=0, column=1, sticky="nsew")
    banner_slot = tk.Frame(right, bg=BG, height=1)      # the result message appears here, above the log
    banner_slot.pack(fill="x")

    def card(parent, title, expand=False):
        outer = tk.Frame(parent, bg=LINE)
        outer.pack(fill="both" if expand else "x", expand=expand, pady=(0, px(12)))
        inner = tk.Frame(outer, bg=CARD, padx=px(16), pady=px(12))
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        top = tk.Frame(inner, bg=CARD)
        top.pack(fill="x")
        tk.Label(top, text=title, bg=CARD, fg=INK, font=F(11, "bold")).pack(side="left")
        return inner, top

    # ---- files
    files_card, _ = card(left, "Files and dashboard")
    grid = tk.Frame(files_card, bg=CARD)
    grid.pack(fill="x", pady=(px(8), 0))
    grid.columnconfigure(1, weight=1)
    var = {k: tk.StringVar(value=saved.get(k, "")) for k in ("sbpdcl", "nbpdcl", "targets", "out", "dashboard")}
    fixed_link, fixed_password = DASHBOARD_LINK.strip(), DASHBOARD_PASSWORD.strip()
    if fixed_link:                                # built in above: it always wins over what was typed before
        var["dashboard"].set(fixed_link)
    secret = {"pw": fixed_password}               # the upload password: kept in memory only, never written to disk
    if "password" in saved:                       # earlier versions stored it in the settings file: remove it from there now
        saved.pop("password")
        try:
            with open(SETTINGS, "w", encoding="utf-8") as fh:
                json.dump(saved, fh)
        except Exception:
            pass
    auto_upload = tk.BooleanVar(value=bool(saved.get("auto_upload", bool(fixed_link))))
    note, entries = {}, {}
    lockable = []
    q = queue.Queue()
    job = {"testing": False, "running": False}

    def refresh(key):
        path = var[key].get().strip()
        if key == "out":
            if not path:
                state, text = "empty", "Not chosen"
            elif os.path.isdir(path):
                state, text = "ok", "The Excel file will be saved here"
            elif os.path.exists(path):
                state, text = "bad", "This is a file, not a folder"
            else:
                state, text = "ok", "This folder will be created"
        elif key == "targets":
            state, text = check_targets(path)
        elif key == "dashboard":
            state, text = check_dashboard_link(path)
            if fixed_link:
                text = "Dashboard link is built in to this app" if state == "ok" else "The built-in link is wrong. " + text
        else:
            state, text = check_export(path)
            if state == "empty":
                text = "Not chosen (leave empty to summarise the other DISCOM only)"
        note[key].configure(text=("✓  " if state == "ok" else "✕  " if state == "bad" else "") + text,
                            fg=GREEN if state == "ok" else RED if state == "bad" else MUTED)
        if key in entries:                        # show the end of a long path: the file name matters most
            entries[key].after_idle(lambda: entries[key].xview_moveto(1))
        return state

    def browse(key):
        if key == "out":
            p = filedialog.askdirectory(title="Output folder", initialdir=var["out"].get() or saved.get("last_dir") or None)
        else:
            start = os.path.dirname(var[key].get()) or saved.get("last_dir") or None
            title = "District targets (CSV: District, Target)" if key == "targets" else f"{key.upper()} export (CSV)"
            p = filedialog.askopenfilename(title=title, initialdir=start, filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not p:
            return
        p = os.path.normpath(p)
        var[key].set(p)
        if key in ("sbpdcl", "nbpdcl"):
            saved["last_dir"] = os.path.dirname(p)
            if not var["out"].get().strip():
                var["out"].set(os.path.dirname(p))

    def save_template():
        p = filedialog.asksaveasfilename(title="Save a blank targets file", defaultextension=".csv", initialfile="district_targets.csv",
                                         filetypes=[("CSV files", "*.csv")], initialdir=var["out"].get() or saved.get("last_dir") or None)
        if not p:
            return
        try:
            with open(p, "w", encoding="utf-8-sig", newline="") as fh:
                fh.write(TARGETS_TEMPLATE)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Could not save the file:\n{exc.strerror or exc}")
            return
        say(f"Blank targets file saved: {os.path.normpath(p)}")
        say("Fill in District and Target, save it as CSV, then choose it in the TARGETS row.")
        detail.configure(text="Blank targets file saved: " + os.path.basename(p))

    rows = [("sbpdcl", "SBPDCL", "South Bihar export (CSV)"), ("nbpdcl", "NBPDCL", "North Bihar export (CSV)"),
            ("targets", "TARGETS", "District targets (CSV, optional)"), ("out", "OUTPUT", "Folder for the Excel file")]
    for r, (key, chip, hint) in enumerate(rows):
        tk.Label(grid, text=chip, bg=CHIP[chip], fg="#FFFFFF", font=F(9, "bold"), width=11, pady=px(5)).grid(
            row=2 * r, column=0, sticky="w", padx=(0, px(10)), pady=(px(6), 0))
        e = entries[key] = ttk.Entry(grid, textvariable=var[key])
        e.grid(row=2 * r, column=1, sticky="ew", pady=(px(6), 0))
        b = ttk.Button(grid, text="Browse…", style="Soft.TButton", command=lambda k=key: browse(k))
        b.grid(row=2 * r, column=2, padx=(px(8), 0), pady=(px(6), 0))
        c = ttk.Button(grid, text="Clear", style="Soft.TButton", command=lambda k=key: var[k].set(""))
        c.grid(row=2 * r, column=3, padx=(px(6), 0), pady=(px(6), 0))
        note[key] = tk.Label(grid, bg=CARD, fg=MUTED, font=F(9), anchor="w", justify="left", text=hint,
                             wraplength=px(330) if key == "targets" else 0)
        note[key].grid(row=2 * r + 1, column=1, columnspan=2 if key == "targets" else 3, sticky="w", pady=(px(2), 0))
        if key == "targets":                      # a blank file with just the two headings, to fill in
            tb = ttk.Button(grid, text="Blank file…", style="Small.Soft.TButton", command=save_template)
            tb.grid(row=2 * r + 1, column=3, sticky="ew", padx=(px(6), 0), pady=(px(3), 0))
            lockable.append(tb)
        var[key].trace_add("write", lambda *_, k=key: refresh(k))
        lockable += [e, b, c]
        refresh(key)

    # ---- dashboard link and upload password (same grid, two more rows)
    def save_settings():
        saved.update({k: v.get().strip() for k, v in var.items() if not (k == "dashboard" and fixed_link)})
        saved.pop("password", None)
        saved["auto_upload"] = bool(auto_upload.get())
        try:
            with open(SETTINGS, "w", encoding="utf-8") as fh:
                json.dump(saved, fh)
        except Exception:
            pass

    def open_dashboard():
        state, text = check_dashboard_link(var["dashboard"].get())
        if state != "ok":
            messagebox.showwarning(APP_NAME, "Paste the dashboard's web app link first." if state == "empty" else text)
            return
        webbrowser.open(var["dashboard"].get().strip())

    def show_password_state(wrong=False):
        if fixed_password:
            return
        if wrong:
            pw_state.configure(text="✕  Wrong, asked again", fg=RED)
        elif secret["pw"]:
            pw_state.configure(text="✓  Entered for now", fg=GREEN)
        else:
            pw_state.configure(text="Asked when needed", fg=MUTED)

    def ask_password():
        """Asks for the upload password in a small window. Returns it, or '' if the person cancelled."""
        typed = simpledialog.askstring(APP_NAME, "Upload password for the dashboard\n(the one set in Upload.gs of the Apps Script project):",
                                       show="•", parent=root)
        typed = (typed or "").strip()
        if typed:
            secret["pw"] = typed
            show_password_state()
        return typed

    def forget_password():
        secret["pw"] = fixed_password
        show_password_state()

    def test_link():
        if job["testing"] or job["running"]:
            return
        state, text = check_dashboard_link(var["dashboard"].get())
        if state != "ok":
            messagebox.showwarning(APP_NAME, "Paste the dashboard's web app link first." if state == "empty" else text)
            return
        link, password = var["dashboard"].get().strip(), secret["pw"] or ask_password()
        if not password:
            return
        job["testing"] = True
        test_btn.configure(state="disabled")
        note["dashboard"].configure(text="Checking the link and password…", fg=MUTED)
        save_settings()

        def run():
            try:
                q.put(("tested", True, test_dashboard(link, password)))
            except UserError as exc:
                q.put(("tested", False, str(exc)))
            except Exception as exc:
                q.put(("tested", False, f"Unexpected problem: {exc}"))
        threading.Thread(target=run, daemon=True).start()

    r = 2 * len(rows)
    tk.Label(grid, text="DASHBOARD", bg=CHIP["DASHBOARD"], fg="#FFFFFF", font=F(9, "bold"), width=11, pady=px(5)).grid(
        row=r, column=0, sticky="w", padx=(0, px(10)), pady=(px(10), 0))
    e = entries["dashboard"] = ttk.Entry(grid, textvariable=var["dashboard"])
    e.grid(row=r, column=1, sticky="ew", pady=(px(10), 0))
    test_btn = ttk.Button(grid, text="Test", style="Soft.TButton", command=test_link)
    test_btn.grid(row=r, column=2, sticky="ew", padx=(px(8), 0), pady=(px(10), 0))
    open_btn = ttk.Button(grid, text="Open", style="Soft.TButton", command=open_dashboard)
    open_btn.grid(row=r, column=3, sticky="ew", padx=(px(6), 0), pady=(px(10), 0))
    note["dashboard"] = tk.Label(grid, bg=CARD, fg=MUTED, font=F(9), anchor="w", justify="left", wraplength=px(430))
    note["dashboard"].grid(row=r + 1, column=1, columnspan=3, sticky="w", pady=(px(2), 0))
    tk.Label(grid, text="PASSWORD", bg=CHIP["PASSWORD"], fg="#FFFFFF", font=F(9, "bold"), width=11, pady=px(5)).grid(
        row=r + 2, column=0, sticky="w", padx=(0, px(10)), pady=(px(6), 0))
    pw_buttons = []
    if fixed_password:                            # built in: it is neither shown nor asked for
        tk.Label(grid, text="Built in to this app", bg=CARD, fg=GREEN, font=F(10), anchor="w").grid(
            row=r + 2, column=1, columnspan=3, sticky="w", pady=(px(6), 0))
    else:
        pw_state = tk.Label(grid, bg=CARD, fg=MUTED, font=F(10), anchor="w")
        pw_state.grid(row=r + 2, column=1, sticky="w", pady=(px(6), 0))
        for col, (label, command) in enumerate((("Enter…", ask_password), ("Forget", forget_password))):
            b = ttk.Button(grid, text=label, style="Soft.TButton", command=command)
            b.grid(row=r + 2, column=2 + col, sticky="ew", padx=(px(8) if col == 0 else px(6), 0), pady=(px(6), 0))
            pw_buttons.append(b)
        show_password_state()
    if fixed_link:
        e.configure(state="readonly")
    style.configure("Card.TCheckbutton", background=CARD, foreground=INK, font=F(10), focuscolor=CARD)
    style.map("Card.TCheckbutton", background=[("active", CARD)], indicatorcolor=[("selected", NAVY), ("!selected", "#FFFFFF")])
    auto_box = ttk.Checkbutton(grid, variable=auto_upload, style="Card.TCheckbutton",
                               text="Send each new report to the dashboard and open it")
    auto_box.grid(row=r + 3, column=1, columnspan=3, sticky="w", pady=(px(4), 0))
    var["dashboard"].trace_add("write", lambda *_: refresh("dashboard"))
    lockable += ([] if fixed_link else [e]) + pw_buttons + [auto_box, test_btn]
    refresh("dashboard")

    root.after(300, lambda: [w.xview_moveto(1) for w in entries.values()])

    # ---- actions
    actions = tk.Frame(left, bg=BG)
    actions.pack(fill="x", pady=(0, px(12)))
    run_btn = ttk.Button(actions, text="Create report", style="Primary.TButton")
    run_btn.pack(side="left")
    stop_btn = ttk.Button(actions, text="Cancel", style="Stop.TButton", state="disabled")
    stop_btn.pack(side="left", padx=(px(10), 0))
    folder_btn = ttk.Button(actions, text="Open folder", style="Ghost.TButton", state="disabled")
    folder_btn.pack(side="right")
    excel_btn = ttk.Button(actions, text="Open report", style="Ghost.TButton", state="disabled")
    excel_btn.pack(side="right", padx=(0, px(8)))

    # ---- progress
    prog_card, pills_row = card(right, "")
    for w in pills_row.winfo_children():
        w.destroy()
    pills_row.pack_configure(pady=(0, px(8)))
    pills = {}
    for i, (sid, label) in enumerate(STEPS_SHOWN):
        pills[sid] = tk.Label(pills_row, text=f"{i + 1}  {label}", bg="#E9EEF3", fg=MUTED, font=F(9, "bold"), padx=px(7), pady=px(4))
        pills[sid].pack(side="left", padx=(0, px(5)))
    bar = ttk.Progressbar(prog_card, maximum=1000, style="Run.Horizontal.TProgressbar")
    bar.pack(fill="x")
    line = tk.Frame(prog_card, bg=CARD)
    line.pack(fill="x", pady=(px(6), 0))
    percent = tk.Label(line, bg=CARD, fg=INK, font=F(13, "bold"), text="0%", width=5, anchor="w")
    percent.pack(side="left")
    clock = tk.Label(line, bg=CARD, fg=MUTED, font=F(9), text="")
    clock.pack(side="right")
    detail = tk.Label(line, bg=CARD, fg=MUTED, font=F(10), text="Ready. Choose the files and press Create report.", anchor="w")
    detail.pack(side="left", fill="x", expand=True)
    banner = tk.Label(banner_slot, font=F(10, "bold"), anchor="w", padx=px(12), pady=px(8), justify="left")

    def show_banner(kind, text):
        colours = {"ok": ("#E1F4E9", "#14613F"), "warn": ("#FFF1CF", "#7A5200"), "bad": ("#FBE3E0", "#9B2C1F")}
        if not kind:
            banner.pack_forget()
            banner_slot.configure(height=1)
            return
        banner.configure(text=text, bg=colours[kind][0], fg=colours[kind][1], wraplength=max(px(300), right.winfo_width() - px(30)))
        banner.pack(fill="x", pady=(0, px(12)))

    def pill(sid, state):
        look = {"wait": ("#E9EEF3", MUTED), "now": (NAVY, "#FFFFFF"), "done": ("#DDF3E7", "#14613F"),
                "skip": ("#F3F5F8", "#A7B1BC"), "bad": ("#FBE3E0", "#9B2C1F")}[state]
        n = [s for s, _ in STEPS_SHOWN].index(sid)
        label = STEPS_SHOWN[n][1]
        mark = "✓" if state == "done" else "–" if state == "skip" else "!" if state == "bad" else str(n + 1)
        pills[sid].configure(bg=look[0], fg=look[1], text=f"{mark}  {label}")

    # ---- log
    log_card, log_top = card(right, "Log", expand=True)
    box_frame = tk.Frame(log_card, bg="#14263A")
    box_frame.pack(fill="both", expand=True, pady=(px(8), 0))
    scroll = ttk.Scrollbar(box_frame, orient="vertical")
    scroll.pack(side="right", fill="y")
    box = tk.Text(box_frame, height=7, wrap="word", state="disabled", font=(mono, 10), bg="#14263A", fg="#DCE6F0",
                  insertbackground="#DCE6F0", relief="flat", padx=px(10), pady=px(8), yscrollcommand=scroll.set,
                  selectbackground="#2E4A68", borderwidth=0, highlightthickness=0)
    box.pack(side="left", fill="both", expand=True)
    scroll.configure(command=box.yview)
    box.tag_configure("time", foreground="#7F93A8")
    box.tag_configure("info", foreground="#DCE6F0")
    box.tag_configure("check", foreground="#FFC857")
    box.tag_configure("bad", foreground="#FF8A7A")
    box.tag_configure("good", foreground="#6FDDA0")

    def say(msg, tag=None):
        if tag is None:
            tag = "check" if msg.startswith("CHECK") else "bad" if msg.startswith(("STOPPED", "ERROR", "Traceback")) else \
                "good" if msg.startswith(("Saved", "DONE")) else "info"
        box.configure(state="normal")
        for i, part in enumerate(msg.rstrip("\n").split("\n")):
            box.insert("end", time.strftime("%H:%M:%S  ") if i == 0 else " " * 10, "time")
            box.insert("end", part + "\n", tag)
        box.see("end")
        box.configure(state="disabled")

    def log_text():
        return box.get("1.0", "end").rstrip() + "\n"

    def copy_log():
        root.clipboard_clear()
        root.clipboard_append(log_text())
        detail.configure(text="Log copied.")

    def save_log():
        p = filedialog.asksaveasfilename(title="Save log", defaultextension=".txt", filetypes=[("Text file", "*.txt")],
                                         initialdir=var["out"].get() or None, initialfile=time.strftime("PMSG_Summary_log_%Y-%m-%d_%H%M.txt"))
        if p:
            try:
                with open(p, "w", encoding="utf-8") as fh:
                    fh.write(log_text())
                detail.configure(text="Log saved: " + os.path.basename(p))
            except OSError as exc:
                messagebox.showerror(APP_NAME, f"Could not save the log:\n{exc}")

    def clear_log():
        box.configure(state="normal")
        box.delete("1.0", "end")
        box.configure(state="disabled")

    for text, command in (("Clear", clear_log), ("Save log…", save_log), ("Copy", copy_log)):
        ttk.Button(log_top, text=text, style="Small.Soft.TButton", command=command).pack(side="right", padx=(px(6), 0))

    # ---- running a job
    job.update({"running": False, "cancel": threading.Event(), "started": 0.0, "path": None, "weights": {}, "starts": {}})

    def open_path(path):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            else:
                import subprocess
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not open:\n{path}\n\n{exc}")

    def set_running(on):
        job["running"] = on
        for w in lockable:
            w.configure(state="disabled" if on or (w is test_btn and job["testing"]) else "normal")
        run_btn.configure(state="disabled" if on else "normal")
        stop_btn.configure(state="normal" if on else "disabled", text="Cancel")

    def set_progress(fraction):
        fraction = max(0.0, min(1.0, fraction))
        bar["value"] = min(999.9, 1000 * fraction)      # exactly the maximum would draw as empty
        percent.configure(text=f"{int(100 * fraction)}%")

    def worker(args, send):
        try:
            path, facts = make_report(*args,
                                       progress=lambda phase, frac, msg: q.put(("progress", phase, frac, msg)),
                                       log=lambda msg: q.put(("log", msg)), cancel=job["cancel"].is_set)
            sent = None                           # None: not asked for; dict: the script's answer; str: why it failed
            if send:
                q.put(("progress", "upload", 0.1, "Sending to the dashboard (about a minute)…"))
                try:
                    sent = upload_report(send[0], send[1], path, lambda msg: q.put(("log", msg)))
                except UserError as exc:
                    sent = str(exc)
                except Exception as exc:
                    sent = f"Unexpected problem while sending: {exc}"
            q.put(("done", path, facts, sent))
        except Cancelled:
            q.put(("cancelled",))
        except UserError as exc:
            q.put(("error", str(exc)))
        except Exception as exc:                 # anything unexpected: show it in full, never vanish silently
            import traceback
            q.put(("log", traceback.format_exc()))
            q.put(("error", f"Something unexpected went wrong: {exc}\n\nUse Save log and share the file so it can be fixed."))

    def start():
        sb, nb, out, tg = (var[k].get().strip() for k in ("sbpdcl", "nbpdcl", "out", "targets"))
        problems = []
        if tg:
            state, text = check_targets(tg)
            if state == "bad":
                problems.append("Targets file: " + text)
        if not (sb or nb):
            problems.append("Choose at least one export file.")
        for key, label in (("sbpdcl", "SBPDCL"), ("nbpdcl", "NBPDCL")):
            state, text = check_export(var[key].get().strip())
            if state == "bad":
                problems.append(f"{label}: {text}")
        if sb and nb and os.path.abspath(sb) == os.path.abspath(nb):
            problems.append("The same file is chosen for SBPDCL and NBPDCL.")
        if not out:
            problems.append("Choose an output folder.")
        elif os.path.exists(out) and not os.path.isdir(out):
            problems.append("The output location is a file, not a folder.")
        link = var["dashboard"].get().strip()
        sending = bool(auto_upload.get())
        if sending:
            state, text = check_dashboard_link(link)
            if state != "ok":
                problems.append("Dashboard: " + ("paste the dashboard's web app link, or untick the box below it." if state == "empty" else text))
        if problems:
            show_banner("bad", "\n".join(problems))
            messagebox.showwarning(APP_NAME, "\n".join(problems))
            return
        send = None
        if sending:                               # ask for the password now, before the minutes of reading
            password = secret["pw"] or ask_password()
            if not password:
                detail.configure(text="Not started: no password was entered.")
                return
            send = (link, password)
        save_settings()

        sizes = {"SBPDCL": os.path.getsize(sb) if sb else 0, "NBPDCL": os.path.getsize(nb) if nb else 0}
        both = sum(sizes.values()) or 1
        reading = 0.84 if send else 0.88
        job["weights"] = {"SBPDCL": reading * sizes["SBPDCL"] / both, "NBPDCL": reading * sizes["NBPDCL"] / both, "build": 0.08,
                          "write": 0.04, "upload": 0.04 if send else 0.0}
        at = 0.0
        for sid, _ in STEPS_SHOWN:
            job["starts"][sid] = at
            at += job["weights"][sid]
            pill(sid, "skip" if (sid in sizes and not sizes[sid]) or (sid == "upload" and not send) else "wait")
        job["cancel"].clear()
        job["path"] = None
        job["started"] = time.time()
        show_banner(None, "")
        clear_log()
        set_progress(0)
        detail.configure(text="Starting…")
        folder_btn.configure(state="disabled")
        excel_btn.configure(state="disabled")
        set_running(True)
        say(f"{APP_NAME} {APP_VERSION}")
        threading.Thread(target=worker, args=((sb, nb, out, tg), send), daemon=True).start()

    def cancel():
        if job["running"] and not job["cancel"].is_set():
            job["cancel"].set()
            stop_btn.configure(state="disabled", text="Cancelling…")
            detail.configure(text="Cancelling…")

    def took():
        s = int(time.time() - job["started"])
        return f"{s // 60}:{s % 60:02d}"

    def finish(kind, text, step_state):
        active = [sid for sid, _ in STEPS_SHOWN if pills[sid].cget("bg") == NAVY]
        for sid in active:
            pill(sid, step_state)
        set_running(False)
        show_banner(kind, text)
        root.after(150, lambda: box.see("end"))    # the banner takes room from the log: keep its last lines in view
        if os.path.isdir(var["out"].get().strip()):
            folder_btn.configure(state="normal")

    def poll():
        try:
            while True:
                msg = q.get_nowait()
                if msg[0] == "progress":
                    _, phase, frac, text = msg
                    seen = False
                    for sid, _ in STEPS_SHOWN:
                        if sid == phase:
                            seen = True
                            pill(sid, "done" if frac >= 1 and phase in ("SBPDCL", "NBPDCL") else "now")
                        elif not seen and job["weights"].get(sid):
                            pill(sid, "done")
                    set_progress(job["starts"][phase] + job["weights"][phase] * frac)
                    detail.configure(text=text)
                elif msg[0] == "log":
                    say(msg[1])
                elif msg[0] == "tested":
                    job["testing"] = False
                    if not job["running"]:
                        test_btn.configure(state="normal")
                    if not msg[1] and "Wrong upload password" in msg[2] and not fixed_password:
                        secret["pw"] = ""
                        show_password_state(wrong=True)
                    short = msg[2] if msg[1] else msg[2].split(". ")[0].rstrip(".") + ". See the log."
                    note["dashboard"].configure(text=("✓  " if msg[1] else "✕  ") + short, fg=GREEN if msg[1] else RED)
                    say(("Dashboard test: " if msg[1] else "CHECK: Dashboard test failed. ") + msg[2])
                elif msg[0] == "done":
                    _, path, facts, sent = msg
                    job["path"] = path
                    for sid, _ in STEPS_SHOWN:
                        if job["weights"].get(sid):
                            pill(sid, "bad" if sid == "upload" and isinstance(sent, str) else "done")
                    set_progress(1)
                    size = os.path.getsize(path)
                    size_text = f"{size / 1e6:,.1f} MB" if size >= 1e6 else f"{size / 1e3:,.0f} KB"
                    text = (f"Report created in {took()}:  {os.path.basename(path)}"
                            f"\n{len(facts['sheets'])} sheets  ·  {indian(facts['applications'])} applications"
                            f"  ·  data up to {facts['as_of']:%d %b %Y}")
                    ula = sum(facts.get("ula_left_out", {}).values())
                    text += f"\nRegular applications only  ·  {indian(ula)} ULA rows left out"
                    checks_found = len(facts["warnings"])
                    if checks_found:
                        text += f"\n{checks_found} point{'s' if checks_found > 1 else ''} to check: see the yellow lines in the log"
                    if isinstance(sent, dict):
                        text += f"\nDashboard updated ({len(sent.get('updated') or [])} tabs) and opened in the browser"
                        detail.configure(text="Done. Dashboard opened.")
                    elif isinstance(sent, str):
                        if "Wrong upload password" in sent and not fixed_password:
                            secret["pw"] = ""             # ask again next time
                            show_password_state(wrong=True)
                            sent += " Press Create report again and type the right one."
                        text += "\nNOT sent to the dashboard: " + sent
                        say("CHECK: The report was saved but not sent to the dashboard. " + sent)
                        detail.configure(text="Saved, but not sent to the dashboard.")
                    else:
                        detail.configure(text="Done. Press Open report to see it.")
                    say(f"DONE in {took()}")
                    finish("warn" if checks_found or isinstance(sent, str) else "ok", text, "done")
                    excel_btn.configure(state="normal")
                    if isinstance(sent, dict):
                        webbrowser.open(var["dashboard"].get().strip())
                elif msg[0] == "cancelled":
                    say("STOPPED: cancelled. No report was written.")
                    detail.configure(text="Cancelled.")
                    finish("bad", "Cancelled. No report was written.", "bad")
                else:
                    say("STOPPED: " + msg[1])
                    detail.configure(text="Stopped. See the message.")
                    finish("bad", msg[1], "bad")
                    messagebox.showerror(APP_NAME, msg[1])
        except queue.Empty:
            pass
        if job["running"]:
            clock.configure(text="Elapsed " + took())
            root.update_idletasks()               # repaint now: the busy worker would otherwise starve redraws
        root.after(100, poll)

    def close():
        if job["running"]:
            if not messagebox.askyesno(APP_NAME, "A report is being created.\nStop it and close?"):
                return
            job["cancel"].set()
        root.destroy()

    run_btn.configure(command=start)
    stop_btn.configure(command=cancel)
    folder_btn.configure(command=lambda: open_path(os.path.dirname(job["path"]) if job["path"] else var["out"].get().strip()))
    excel_btn.configure(command=lambda: job["path"] and open_path(job["path"]))
    root.protocol("WM_DELETE_WINDOW", close)
    root.bind("<Return>", lambda e: None if job["running"] else start())
    root.report_callback_exception = lambda exc, val, tb: say(f"ERROR: {val}")
    poll()
    root.mainloop()


def main():
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("--sbpdcl", default="")
    ap.add_argument("--nbpdcl", default="")
    ap.add_argument("--targets", default="", help="optional CSV (district,target) replacing the built-in targets")
    ap.add_argument("--out", default="")
    ap.add_argument("--dashboard", default=DASHBOARD_LINK.strip(),
                    help="the dashboard's web app link: the report is sent there and the dashboard opened (default: the built-in link)")
    ap.add_argument("--password", default=os.environ.get("PMSG_UPLOAD_PASSWORD", "") or DASHBOARD_PASSWORD.strip(),
                    help="upload password (or set PMSG_UPLOAD_PASSWORD; default: the built-in password)")
    ap.add_argument("--no-upload", action="store_true", help="only save the Excel file, do not send it to the dashboard")
    ap.add_argument("--no-open", action="store_true", help="with --dashboard: do not open the browser afterwards")
    args = ap.parse_args()
    if not (args.sbpdcl or args.nbpdcl):
        run_gui()
        return
    shown = {}

    def progress(phase, frac, msg):
        step = int(4 * frac)
        if shown.get(phase) != step:
            shown[phase] = step
            print(f"  {msg}")

    try:
        path, _ = make_report(args.sbpdcl, args.nbpdcl, args.out or "pmsg_report", args.targets, progress, print)
    except UserError as exc:
        sys.exit("STOPPED: " + str(exc))
    if args.dashboard and not args.no_upload:
        if not args.password and sys.stdin and sys.stdin.isatty():
            import getpass
            args.password = getpass.getpass("Upload password for the dashboard: ").strip()
        try:
            upload_report(args.dashboard, args.password, path, print)
        except UserError as exc:
            sys.exit("The report was saved but NOT sent to the dashboard: " + str(exc))
        if not args.no_open:
            webbrowser.open(args.dashboard.strip())


if __name__ == "__main__":
    main()
