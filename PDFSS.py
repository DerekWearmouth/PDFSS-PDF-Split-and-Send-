#   ,-.       _,---._ __  / \
#  /  )    .-'       `./ /   \
# (  (   ,'            `/    /|
#  \  `-"             \'\   / |
#   `.              ,  \ \ /  |
#    /`.          ,'-`----Y   |
#   (            ;        |   '
#   |  ,-.    ,-'         |  /
#   |  | (   |   Meechie  | /
#   )  |  \  `.___________|/

#================================================================================
# PDFSS (PDF-Split-Send) - v1.0 Date: 2026-11-11
#================================================================================

"""
Comments and documentation created by Claude AI.~~~~ Just more of an FYI.. im terrible at comments. 


PDFSS (PDF-Split-Send)

Desktop tool that replaces PDF-eXPLODE for the weekly FoxPro statement run.

FoxPro writes one large PDF that holds every customer statement. Each page has
a hidden (white text) tag near the top that says who the page belongs to:

    <pdfexplode>UCS105#someone@example.com</pdfexplode>   email this statement
    <pdfexplode>UCS105#PRINT</pdfexplode>                 no email, print it

This app:
    1. Reads every page of the chosen PDF and groups pages by that tag.
       All statements for the same email address become one PDF.
       Each PRINT statement becomes its own PDF ("No Sender").
       Pages with no tag at all (the report totals page) become "Untagged pages".
    2. Shows every group in a table so a person can review it, open any PDF,
       and choose Email, Print or Skip per row.
    3. On Send, emails each row through the logged-in user's classic Outlook
       (COM automation) and prints rows to the default Windows printer.
    4. Saves a copy of everything to a per-run folder and a per-recipient
       folder on the share, and writes a CSV log.

Run modes (command line flag --mode):
    send     real run, emails are sent and printouts are printed (default)
    display  emails open in Outlook instead of sending, printing still happens
    dryrun   nothing touches Outlook or the printer, files and logs still written

Shared files kept in LOG_FOLDER:
    settings.json           default From, file name, subject, body, send delay
    recipient_history.json  every address emailed and its last page count
    sent_files.json         weekly PDFs already sent (drives the re-send warning)
    send_log_<stamp>.csv    one log per run

Runtime requirements: Windows, classic Outlook, packages in requirements.txt.
"""

import argparse
import csv
import getpass
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import tkinter as tk
from collections import OrderedDict
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pymupdf


# ---- Settings ----------------------------------------------------------------

# Name and version shown in the window title and message boxes.
APP_NAME = "PDFSS"
APP_VERSION = "1.0"

# Share locations. These are built into the .exe, so change them before building.
# STATEMENTS_FOLDER must already exist; the other two are created when needed.
STATEMENTS_FOLDER = r"\\fileserver\share\Statements"
OUTPUT_FOLDER = r"\\fileserver\share\Statements\Output"
LOG_FOLDER = r"\\fileserver\share\Statements\Logs"

# Email defaults used until someone saves the Settings dialog (settings.json).
# "{date}" in any of these is replaced with today's date at send time.
DEFAULT_FILE_NAME = "statement.pdf"
DEFAULT_SUBJECT = "Statements {date}"
DEFAULT_BODY = "Please find your statement(s) attached."
DATE_FORMAT = "%m/%d/%Y"          # {date} in subject and body
FILE_DATE_FORMAT = "%m-%d-%Y"     # {date} in file names (no slashes allowed)

# How pages are read.
# HEADER_BAND: only the top 20% of each page is searched, so emails printed in
#   the body or footer of a statement are never mistaken for the recipient.
# TAG_RE: the hidden FoxPro tag. Group 1 is the value after the "#" (an email
#   address or the word PRINT). The "UCS105#" prefix is optional.
# START_RE: "Page 1" marks the first page of a statement. Used to split
#   back-to-back PRINT statements and to count statements.
# TOTALS_RE: statement count on the untagged report totals page, used as a
#   cross-check against the number of statements found.
# CARRY_FORWARD: if True, an untagged page joins the statement before it.
#   Left False because every real statement page carries a tag, so an untagged
#   page should never be silently added to someone else's statement.
HEADER_BAND = 0.20
TAG_RE = re.compile(r"<pdfexplode>(?:[^#<]*#)?([^<]+)</pdfexplode>", re.I)
PRINT_VALUE = "print"
START_RE = re.compile(r"Page\s+1\b")
TOTALS_RE = re.compile(r"Report Totals:\s*(\d+)")
CARRY_FORWARD = False

# Behavior limits.
# MOVE_AFTER_SEND: move the weekly PDF into its run folder after a real send.
# MAX_ATTACHMENT_MB: rows over this size can be printed or skipped, not emailed.
# DELAY_MINUTES: starting send delay until one is saved in Settings.
# MAX_DELAY_MINUTES: upper bound accepted by the Settings dialog (24 hours).
# PRINT_DPI: render resolution used when printing.
# JUMP_FACTOR / JUMP_MIN_PAGES: flag a recipient whose page count is at least
#   2x their last run AND at least 5 pages more (catches merged statements).
# STALE_LOCK_HOURS: a lock file older than this is assumed left by a crash.
MOVE_AFTER_SEND = False
MAX_ATTACHMENT_MB = 20
DELAY_MINUTES = 10
MAX_DELAY_MINUTES = 1440
PRINT_DPI = 300
JUMP_FACTOR = 2.0
JUMP_MIN_PAGES = 5
STALE_LOCK_HOURS = 4

# Display labels and file names used on the share.
NO_SENDER = "No Sender"
NO_FROM = "(choose a sender)"     # placeholder in the From dropdown, means "not set"
LOCK_NAME = ".pdfss.lock"
SETTINGS_NAME = "settings.json"
RECIPIENTS_NAME = "recipient_history.json"
SENT_FILES_NAME = "sent_files.json"

# EMAIL_RE finds an address anywhere in a string (fallback when no tag exists).
# VALID_EMAIL checks that a whole string is one address (used for the To field).
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
VALID_EMAIL = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

# Row actions, and the status words that count as "emailed" or "printed" in
# each mode. Test modes use their own words so logs never claim a real send.
EMAIL, PRINT, SKIP = "Email", "Print", "Skip"
EMAILED = ("Emailed", "Displayed", "Would email")
PRINTED = ("Printed", "Would print")


# ---- Text helpers ------------------------------------------------------------

def fill(text, for_file=False):
    """Replace {date} with today's date. File names use dashes instead of slashes."""
    return text.replace("{date}", datetime.now().strftime(FILE_DATE_FORMAT if for_file else DATE_FORMAT))


def clean_file_name(name):
    """Turn a user-entered attachment name into a safe Windows file name ending in .pdf."""
    # Replace characters Windows does not allow in file names, then trim dots
    # and spaces from the ends (Windows strips those and can misbehave).
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", fill(name, True)).strip().strip(".")
    name = name or DEFAULT_FILE_NAME
    return name if name.lower().endswith(".pdf") else name + ".pdf"


def safe_name(text):
    """Make any text (an email address or customer name) safe to use as a file or folder name."""
    # Removing slashes and trimming dots means text from the PDF can never
    # point outside the intended folder (no "..\" style paths).
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", text).strip().strip(".") or "unnamed"


def unique_path(path):
    """Return path unchanged if free, otherwise add " (2)", " (3)" and so on until it is."""
    path = Path(path)
    n = 2
    while path.exists():
        path = path.with_name(f"{path.stem.rsplit(' (', 1)[0]} ({n}){path.suffix}")
        n += 1
    return path


# ---- Row ---------------------------------------------------------------------

class Row:
    """
    One line in the review table: everything that goes to one recipient.

    kind is one of:
        "email"     pages tagged with an email address (all combined into one PDF)
        "print"     one statement tagged PRINT ("No Sender")
        "untagged"  pages with no tag, such as the report totals page

    found is the address read from the PDF. Anything the user changes in the
    Edit dialog is stored in `over` (overrides), so the original is never lost
    and "Reset to defaults" just clears `over`.
    """

    def __init__(self, found, kind):
        self.found = found              # email from the PDF tag, or None
        self.kind = kind                # "email", "print" or "untagged"
        self.pages = []                 # list of (source pdf path, page index)
        self.names = []                 # customer name from each statement's first page
        self.pdf = None                 # path of the built PDF in the temp work folder
        self.page_count = 0
        self.size_mb = 0.0
        self.build_error = None         # set if the PDF could not be built
        self.over = {}                  # per-row overrides: to, sender, file_name, subject, body
        self.action = None              # EMAIL, PRINT or SKIP
        self.status = ""                # result after Send, written to the log
        self.sent_name = ""             # attachment file name actually used
        self.flags = []                 # notes shown in the "Check this" column
        self.level = ""                 # "bad" (red), "warn" (yellow) or "" (normal)

    @property
    def to(self):
        """The address this row will be emailed to: the edited one if set, else the one from the PDF."""
        return self.over.get("to", self.found or "").strip()

    @property
    def label(self):
        """Text for the Recipient column."""
        if self.to:
            return self.to
        return "Untagged pages" if self.kind == "untagged" else NO_SENDER

    @property
    def file_base(self):
        """Name used for this row's saved PDF and print job. Adds the customer name for No Sender rows."""
        if self.to:
            return self.to
        return f"{self.label} - {self.names[0]}" if self.names else self.label

    def names_text(self):
        """Text for the "Statement for" column: first name plus a count of the rest."""
        if not self.names:
            return ""
        return self.names[0] + (f"  (+{len(self.names) - 1} more)" if len(self.names) > 1 else "")

    def value(self, key, defaults):
        """A message setting for this row: its own override if edited, else the shared default."""
        return self.over.get(key, defaults[key])

    def can_email(self):
        """True if the PDF was built, the To address is valid, and it is under the size limit."""
        return bool(self.pdf) and bool(VALID_EMAIL.match(self.to)) and self.size_mb <= MAX_ATTACHMENT_MB

    def allowed(self):
        """Actions this row may take, in the order clicking the Action column cycles through them."""
        return [a for a, ok in ((EMAIL, self.can_email()), (PRINT, bool(self.pdf)), (SKIP, True)) if ok]

    def fix_action(self):
        """If the current action is no longer allowed (or unset), fall back to the first allowed one."""
        if self.action not in self.allowed():
            self.action = self.allowed()[0]

    def check(self, history):
        """
        Rebuild this row's flags and color from scratch.

        Called after loading, after an edit, and after Settings change, so the
        notes always match the row's current state. `history` is the contents
        of recipient_history.json.
        """
        self.flags, self.level = [], ""

        def add(text, level=None):
            # "bad" always wins; "warn" only applies if nothing is set yet.
            # A level of None adds a note without changing the color.
            self.flags.append(text)
            if level == "bad" or (level and not self.level):
                self.level = level

        if self.build_error:
            add(f"Could not build PDF: {self.build_error}", "bad")
        if not self.to:
            if self.kind == "untagged":
                add("No statement tag (report totals or stray page). Skipped by default.", "bad")
            else:
                add("No email. Print or skip.", "bad")
        elif not VALID_EMAIL.match(self.to):
            add("To address looks wrong", "bad")
        if self.size_mb > MAX_ATTACHMENT_MB:
            add(f"Too large to email ({self.size_mb:.0f} MB)", "bad")
        if self.to and VALID_EMAIL.match(self.to):
            seen = history.get(self.to.lower())
            if not seen:
                add("New recipient", "warn")
            else:
                # Flag a big jump in pages, which usually means two customers'
                # statements ended up under one address.
                last = seen.get("last_pages", 0)
                if last and self.page_count >= last * JUMP_FACTOR and self.page_count - last >= JUMP_MIN_PAGES:
                    add(f"Pages up from {last} last time", "warn")
        if self.over:
            add("Edited")


# ---- Lock --------------------------------------------------------------------
# A lock file in the statements folder stops two people running the app at the
# same time, which could send the same week twice.

def acquire_lock(folder):
    """
    Create the lock file. Returns (lock path, info) on success, or
    (None, who has it) if someone else holds the lock.
    """
    path = Path(folder) / LOCK_NAME
    # A lock this old was almost certainly left behind by a crash.
    if path.exists() and time.time() - path.stat().st_mtime > STALE_LOCK_HOURS * 3600:
        path.unlink(missing_ok=True)
    info = f"{getpass.getuser()} on {socket.gethostname()} since {datetime.now():%I:%M %p}"
    try:
        # O_EXCL makes create-if-missing a single atomic step, so two users
        # starting at the same moment cannot both get the lock.
        fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        try:
            return None, path.read_text(errors="ignore")
        except OSError:
            return None, "another user"
    with os.fdopen(fd, "w") as fh:
        fh.write(info)
    return path, info


def release_lock(path):
    """Delete the lock file if this instance created it."""
    if path:
        Path(path).unlink(missing_ok=True)


# ---- Stored data -------------------------------------------------------------

def load_json(folder, name):
    """Read a JSON file from the logs folder. A missing or unreadable file returns {}."""
    try:
        return json.loads((Path(folder) / name).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_json(folder, name, data):
    """Write a JSON file safely: write a temp file first, then swap it in."""
    Path(folder).mkdir(parents=True, exist_ok=True)
    path = Path(folder) / name
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=1, sort_keys=True), encoding="utf-8")
    # os.replace is atomic, so a crash mid-write never leaves a half-written file.
    os.replace(tmp, path)


def load_defaults(folder):
    """
    The shared message defaults: built-in values overlaid with settings.json.
    Unknown keys in the file are ignored, and the delay is clamped to a valid range.
    """
    d = {"sender": "", "file_name": DEFAULT_FILE_NAME, "subject": DEFAULT_SUBJECT, "body": DEFAULT_BODY,
         "delay": DELAY_MINUTES}
    d.update({k: v for k, v in load_json(folder, SETTINGS_NAME).items() if k in d})
    try:
        d["delay"] = min(max(int(d["delay"]), 0), MAX_DELAY_MINUTES)
    except (TypeError, ValueError):
        d["delay"] = DELAY_MINUTES
    return d


def file_key(path):
    """Identifies a weekly PDF in sent_files.json by name and size."""
    return f"{Path(path).name}|{Path(path).stat().st_size}"


# ---- PDF scanning and building -----------------------------------------------

def newest_pdf(folder):
    """The most recently modified PDF in a folder, preselected in the file picker."""
    files = [p for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def page_info(page):
    """
    Read one page's header band.

    Returns (value, start, name):
        value  the tag value in lower case (an email or "print"), or None
        start  True if this is page 1 of a statement
        name   first line of text on the page (the customer name on page 1)
    """
    r = page.rect
    header = page.get_text("text", clip=pymupdf.Rect(r.x0, r.y0, r.x1, r.y0 + r.height * HEADER_BAND))
    m = TAG_RE.search(header)
    value = m.group(1).strip().lower() if m else None
    if value is None:
        # Fallback for PDFs without the FoxPro tag: any address in the header band.
        e = EMAIL_RE.search(header)
        value = e.group(0).lower() if e else None
    start = bool(START_RE.search(header))
    name = next((ln.strip() for ln in header.splitlines() if ln.strip()), "")
    return value, start, name


def scan_file(path, progress=None):
    """
    Walk every page of the weekly PDF and group pages into Rows.

    Returns (total pages, rows, statements found, statement count from the
    totals page or None). `progress(n, total)` is called every 25 pages.
    """
    # Keys: an email address for email rows (so repeat addresses merge into one
    # row), or a (kind, page index) tuple for print and untagged rows (so each
    # stays separate). OrderedDict keeps the order pages appeared in.
    groups = OrderedDict()
    doc = pymupdf.open(path)
    total = doc.page_count
    current = None          # the row the previous page went into
    statements = 0
    expected = None
    for i, page in enumerate(doc):
        value, start, name = page_info(page)
        if value and value != PRINT_VALUE:
            # Email page: join (or start) that address's row, wherever it appears.
            row = groups.get(value) or groups.setdefault(value, Row(value, "email"))
        elif value == PRINT_VALUE:
            # PRINT page: a new No Sender row at each "Page 1", otherwise this is
            # page 2+ of the PRINT statement before it.
            if start or current is None or current.kind != "print":
                row = groups.setdefault(("print", i), Row(None, "print"))
            else:
                row = current
        elif CARRY_FORWARD and current is not None and not start:
            row = current
        else:
            # Untagged page: consecutive untagged pages share one row.
            if current is not None and current.kind == "untagged":
                row = current
            else:
                row = groups.setdefault(("untagged", i), Row(None, "untagged"))
            # The report totals page lists how many statements FoxPro produced.
            m = TOTALS_RE.search(page.get_text("text"))
            if m:
                expected = int(m.group(1))
        if start and value:
            statements += 1
            row.names.append(name)
        row.pages.append((path, i))
        current = row
        if progress and (i % 25 == 0 or i == total - 1):
            progress(i + 1, total)
    doc.close()
    return total, list(groups.values()), statements, expected


def build_pdf(pages, out_path):
    """Copy the given (source, page index) pages, in order, into a new PDF. Returns its page count."""
    out = pymupdf.open()
    srcs = {}           # each source PDF is opened once and reused
    try:
        for f, i in pages:
            if f not in srcs:
                srcs[f] = pymupdf.open(f)
            out.insert_pdf(srcs[f], from_page=i, to_page=i)
        # garbage=3 and deflate drop unused objects and compress, keeping attachments small.
        out.save(out_path, garbage=3, deflate=True)
        return out.page_count
    finally:
        out.close()
        for d in srcs.values():
            d.close()


# ---- Outlook -----------------------------------------------------------------
# Email goes through classic Outlook's COM interface (pywin32), so messages are
# sent from the user's own mailbox and land in their Sent Items. The new Outlook
# for Windows has no COM interface and will not work.

def get_outlook():
    """Connect to Outlook. Starts it in the background if it is not running."""
    import win32com.client
    return win32com.client.Dispatch("Outlook.Application")


def account_name(acct):
    """An Outlook account's address, falling back to its display name."""
    return (acct.SmtpAddress or acct.DisplayName or "").strip()


def outlook_accounts(outlook):
    """All account addresses in this user's Outlook profile (fills the From dropdown)."""
    accts = outlook.Session.Accounts
    # Outlook collections are 1-based.
    return [n for n in (account_name(accts.Item(i)) for i in range(1, accts.Count + 1)) if n]


def find_account(outlook, name):
    """The Outlook account object matching an address, or None."""
    accts = outlook.Session.Accounts
    for i in range(1, accts.Count + 1):
        if account_name(accts.Item(i)).lower() == name.lower():
            return accts.Item(i)
    return None


def send_mail(outlook, to, sender, subject, body, attachment, mode, delay):
    """
    Create and send (or display) one email.

    A sender is required: there is deliberately no fallback to Outlook's default
    account, so nothing can go out from an unexpected mailbox. Raises on any
    problem so the caller can mark just this row as an error.
    """
    if not sender:
        raise RuntimeError("No From address")
    acct = find_account(outlook, sender)
    if acct is None:
        raise RuntimeError(f"From account {sender} is not set up in this Outlook")
    mail = outlook.CreateItem(0)        # 0 = olMailItem
    # Sets MailItem.SendUsingAccount. 64209 is that property's dispatch id and
    # 8 is DISPATCH_PROPERTYPUTREF. A plain attribute assignment does not work
    # reliably through pywin32's late binding, so the property is set directly.
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, acct))
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(str(attachment))
    if mode == "display":
        mail.Display()
        return
    if delay > 0:
        # Outlook holds the message in the Outbox until this time, giving the
        # user a window to delete it. Passed as text in US date format.
        when = datetime.now() + timedelta(minutes=delay)
        mail.DeferredDeliveryTime = when.strftime("%m/%d/%Y %I:%M:%S %p")
    mail.Send()


# ---- Printing ----------------------------------------------------------------
# Pages are rendered to images with PyMuPDF and drawn straight onto the default
# printer through the Windows GDI. This needs no PDF reader installed and shows
# no print dialog.

def default_printer():
    """Name of the default Windows printer. Raises if none is set."""
    import win32print
    return win32print.GetDefaultPrinter()


def print_pdf(path, title):
    """Print every page of a PDF to the default printer as one print job named `title`."""
    import win32print
    import win32ui
    from PIL import Image, ImageWin

    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(win32print.GetDefaultPrinter())
    # Printable area in printer pixels (8 = HORZRES, 10 = VERTRES).
    width, height = dc.GetDeviceCaps(8), dc.GetDeviceCaps(10)
    doc = pymupdf.open(path)
    try:
        dc.StartDoc(title)
        for page in doc:
            pix = page.get_pixmap(dpi=PRINT_DPI, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            # Rotate landscape pages onto portrait paper (and the reverse).
            if (img.width > img.height) != (width > height):
                img = img.rotate(90, expand=True)
            # Scale to fit the printable area, keep proportions, and center.
            scale = min(width / img.width, height / img.height)
            w, h = int(img.width * scale), int(img.height * scale)
            x, y = (width - w) // 2, (height - h) // 2
            dc.StartPage()
            ImageWin.Dib(img).draw(dc.GetHandleOutput(), (x, y, x + w, y + h))
            dc.EndPage()
        dc.EndDoc()
    finally:
        doc.close()
        dc.DeleteDC()


# ---- Wrap-up -----------------------------------------------------------------

def finish_run(cfg, rows, source, key, mode, defaults):
    """
    Save the results of a Send.

    Writes, for every row with a built PDF:
        Output/Runs/<date> <file>/Emailed|Printed|Not Sent/<name>.pdf
        Output/Recipients/<email>/<date> <file name>   (emailed rows only)
    plus send_log.csv in the run folder and a copy in the logs folder.

    In test modes the Recipients folder is created inside the run folder, and
    history and sent_files are not updated, so tests never affect real runs.
    Returns the run folder path.
    """
    now = datetime.now()
    stamp = now.strftime("%Y-%m-%d_%H%M%S")
    user = getpass.getuser()
    test = mode != "send"

    name = f"{now:%Y-%m-%d} {source.stem}" + (f" ({mode} test)" if test else "")
    run = Path(cfg.output) / "Runs" / name
    if run.exists():
        # Same file run twice in one day: add the time so nothing is overwritten.
        run = Path(cfg.output) / "Runs" / f"{name} {now:%H%M%S}"
    run.mkdir(parents=True)
    people = (run / "Recipients") if test else (Path(cfg.output) / "Recipients")

    # Copy each row's PDF into the right subfolder.
    for r in rows:
        if not r.pdf or not r.pdf.exists():
            continue
        if r.status in EMAILED:
            sub, base = "Emailed", r.to
        elif r.status in PRINTED:
            sub, base = "Printed", r.file_base
        else:
            sub, base = "Not Sent", r.file_base
        (run / sub).mkdir(exist_ok=True)
        shutil.copy2(r.pdf, unique_path(run / sub / f"{safe_name(base)}.pdf"))
        if r.status in EMAILED:
            folder = people / safe_name(r.to.lower())
            folder.mkdir(parents=True, exist_ok=True)
            shutil.copy2(r.pdf, unique_path(folder / f"{now:%Y-%m-%d} {r.sent_name}"))

    # One log line per row, written to the run folder and the logs folder.
    header = ["timestamp", "user", "mode", "source_file", "action", "to", "from", "file_name",
              "subject", "statements", "pages", "notes", "status"]
    log_rows = [[stamp, user, mode, source.name, r.action, r.to,
                 r.value("sender", defaults), r.sent_name,
                 fill(r.value("subject", defaults)) if r.action == EMAIL else "",
                 " | ".join(r.names), r.page_count, "; ".join(r.flags), r.status] for r in rows]
    for log_path in (run / "send_log.csv", Path(cfg.logs) / f"send_log_{stamp}.csv"):
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with open(log_path, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(header)
            w.writerows(log_rows)

    if not test:
        # Remember each address and its page count for the New recipient and
        # page jump checks next time. Only real sends count.
        history = load_json(cfg.logs, RECIPIENTS_NAME)
        for r in rows:
            if r.status == "Emailed":
                history[r.to.lower()] = {"last_pages": r.page_count, "last_sent": stamp}
        save_json(cfg.logs, RECIPIENTS_NAME, history)

        # Record this weekly file so opening it again shows the re-send warning.
        sent_files = load_json(cfg.logs, SENT_FILES_NAME)
        sent_files[key] = {"sent": now.strftime("%m/%d/%Y %I:%M %p"), "user": user, "output": run.name}
        save_json(cfg.logs, SENT_FILES_NAME, sent_files)

        if MOVE_AFTER_SEND:
            shutil.move(str(source), str(run / source.name))

    return run


# ---- Message form ------------------------------------------------------------
# The From / file name / subject / body fields are shared by the Settings dialog
# and the Edit dialog, so they are built by one function.

def message_form(frame, start, accounts, values, include_to):
    """
    Add the message fields to `frame` starting at grid row `start`.

    `values` holds the starting text for each field. `include_to` adds a To
    field (Edit dialog only). Returns (getters, next free grid row), where
    getters maps each field name to a function that returns its current value.
    """
    getters = {}
    r = start

    def label(text):
        ttk.Label(frame, text=text).grid(row=r, column=0, sticky="nw", pady=4, padx=(0, 10))

    if include_to:
        label("To")
        v_to = tk.StringVar(value=values["to"])
        ttk.Entry(frame, textvariable=v_to, width=60).grid(row=r, column=1, sticky="ew", pady=4)
        getters["to"] = lambda: v_to.get().strip()
        r += 1

    # From: read-only dropdown of this user's Outlook accounts. If no accounts
    # were found (dry run, Outlook unavailable) the box allows typing instead.
    # NO_FROM is the "not set" placeholder and is returned as an empty string.
    label("From")
    v_from = tk.StringVar(value=values["sender"] or NO_FROM)
    ttk.Combobox(frame, textvariable=v_from, values=[NO_FROM] + accounts, width=58,
                 state="readonly" if accounts else "normal").grid(row=r, column=1, sticky="ew", pady=4)
    getters["sender"] = lambda: "" if v_from.get().strip() in ("", NO_FROM) else v_from.get().strip()
    r += 1

    label("File name")
    v_file = tk.StringVar(value=values["file_name"])
    ttk.Entry(frame, textvariable=v_file, width=60).grid(row=r, column=1, sticky="ew", pady=4)
    getters["file_name"] = lambda: v_file.get().strip()
    r += 1

    label("Subject")
    v_subj = tk.StringVar(value=values["subject"])
    ttk.Entry(frame, textvariable=v_subj, width=60).grid(row=r, column=1, sticky="ew", pady=4)
    getters["subject"] = lambda: v_subj.get().strip()
    r += 1

    label("Body")
    body = tk.Text(frame, width=60, height=8, wrap="word", font=("Segoe UI", 9))
    body.insert("1.0", values["body"])
    body.grid(row=r, column=1, sticky="nsew", pady=4)
    getters["body"] = lambda: body.get("1.0", "end-1c").strip()
    frame.rowconfigure(r, weight=1)
    r += 1

    ttk.Label(frame, text="{date} is replaced with today's date.", foreground="#666").grid(
        row=r, column=1, sticky="w")
    frame.columnconfigure(1, weight=1)
    # Keep references to the StringVars on the frame. Tkinter clears a
    # variable's value if the Python object is garbage collected.
    frame.form_vars = [v for v in (locals().get("v_to"), v_from, v_file, v_subj) if v is not None]
    return getters, r + 1


class Dialog(tk.Toplevel):
    """Base for the Settings and Edit windows: a body frame, a button row, and modal display."""

    def __init__(self, app, title):
        super().__init__(app.root)
        self.app = app
        self.title(title)
        self.transient(app.root)
        self.minsize(560, 420)
        self.body = ttk.Frame(self, padding=12)
        self.body.pack(fill="both", expand=True)
        self.buttons = ttk.Frame(self, padding=(12, 0, 12, 12))
        self.buttons.pack(fill="x")
        self.bind("<Escape>", lambda e: self.destroy())

    def show(self):
        """Place the window over the main window and block until it closes."""
        self.update_idletasks()
        x = self.app.root.winfo_rootx() + 80
        y = self.app.root.winfo_rooty() + 60
        self.geometry(f"+{x}+{y}")
        self.grab_set()         # modal: the main window ignores clicks meanwhile
        self.wait_window()


# ---- Settings dialog ---------------------------------------------------------

class SettingsDialog(Dialog):
    """Edits the shared defaults (settings.json): From, file name, subject, body, send delay."""

    def __init__(self, app):
        super().__init__(app, "Default email settings")
        ttk.Label(self.body, text="These defaults apply to every recipient unless changed with Edit.",
                  foreground="#555").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self.getters, r = message_form(self.body, 1, app.accounts, app.defaults, include_to=False)

        ttk.Label(self.body, text="Send delay").grid(row=r, column=0, sticky="w", pady=(10, 4), padx=(0, 10))
        df = ttk.Frame(self.body)
        df.grid(row=r, column=1, sticky="w", pady=(10, 4))
        self.delay = tk.StringVar(value=str(app.defaults["delay"]))
        ttk.Spinbox(df, from_=0, to=MAX_DELAY_MINUTES, increment=5, textvariable=self.delay, width=6).pack(side="left")
        ttk.Label(df, text="minutes before emails leave the Outbox  (0 = send right away)",
                  foreground="#666").pack(side="left", padx=8)

        ttk.Button(self.buttons, text="Save", command=self.save).pack(side="right")
        ttk.Button(self.buttons, text="Cancel", command=self.destroy).pack(side="right", padx=8)
        self.show()

    def save(self):
        """Validate, write settings.json, and refresh the main window."""
        values = {k: g() for k, g in self.getters.items()}
        values["file_name"] = values["file_name"] or DEFAULT_FILE_NAME
        try:
            values["delay"] = int(self.delay.get().strip())
            if not 0 <= values["delay"] <= MAX_DELAY_MINUTES:
                raise ValueError
        except ValueError:
            messagebox.showerror("Settings", f"Send delay must be a whole number from 0 to {MAX_DELAY_MINUTES}.",
                                 parent=self)
            return
        try:
            save_json(self.app.cfg.logs, SETTINGS_NAME, values)
        except OSError as e:
            messagebox.showerror("Settings", f"Settings could not be saved.\n\n{e}", parent=self)
            return
        self.app.defaults = values
        self.app.update_from_label()
        self.app.recheck_all()      # rows without overrides pick up the new defaults
        self.destroy()


# ---- Edit dialog -------------------------------------------------------------

class EditDialog(Dialog):
    """Edits one row: action, To, From, file name, subject, body. Changes apply to this run only."""

    def __init__(self, app, row):
        super().__init__(app, f"Edit  {row.file_base}")
        self.row = row
        d = app.defaults

        # Show which customers this row covers (first six, then a count).
        if row.names:
            ttk.Label(self.body, text="Statements:  " + ", ".join(row.names[:6]) +
                      (f"  (+{len(row.names) - 6} more)" if len(row.names) > 6 else ""),
                      foreground="#555", wraplength=520).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(self.body, text="Action").grid(row=1, column=0, sticky="w", pady=(8, 4))
        self.action = tk.StringVar(value=row.action)
        af = ttk.Frame(self.body)
        af.grid(row=1, column=1, sticky="w", pady=(8, 4))
        for a in (EMAIL, PRINT, SKIP):
            ttk.Radiobutton(af, text=a, value=a, variable=self.action).pack(side="left", padx=(0, 14))

        # Start each field from this row's current effective value.
        values = {"to": row.to, "sender": row.value("sender", d), "file_name": row.value("file_name", d),
                  "subject": row.value("subject", d), "body": row.value("body", d)}
        self.getters, _ = message_form(self.body, 2, app.accounts, values, include_to=True)

        ttk.Button(self.buttons, text="Open statement", command=lambda: app.open_pdf(row)).pack(side="left")
        ttk.Button(self.buttons, text="Reset to defaults", command=self.reset).pack(side="left", padx=8)
        ttk.Button(self.buttons, text="Save", command=self.save).pack(side="right")
        ttk.Button(self.buttons, text="Cancel", command=self.destroy).pack(side="right", padx=8)
        self.show()

    def reset(self):
        """Drop all overrides and return the row to its default action."""
        self.row.over = {}
        self.row.action = None
        self.row.check(self.app.history)
        self.row.fix_action()
        if self.row.kind == "untagged":
            self.row.action = SKIP
        self.app.refresh_all()
        self.destroy()

    def save(self):
        """Validate the chosen action, store only the fields that differ, and refresh."""
        values = {k: g() for k, g in self.getters.items()}
        values["file_name"] = values["file_name"] or DEFAULT_FILE_NAME
        action = self.action.get()
        if action == EMAIL:
            if not VALID_EMAIL.match(values["to"]):
                messagebox.showerror("Edit", "Enter a valid To address to email this statement.", parent=self)
                return
            if self.row.size_mb > MAX_ATTACHMENT_MB:
                messagebox.showerror("Edit", "This statement is too large to email. Print or skip it.", parent=self)
                return
        if action == PRINT and not self.row.pdf:
            messagebox.showerror("Edit", "This statement could not be built, so it can't be printed.", parent=self)
            return

        # Only keep values that differ from the PDF's address and the shared
        # defaults. Untouched fields keep following future Settings changes.
        base = {"to": self.row.found or "", **self.app.defaults}
        self.row.over = {k: v for k, v in values.items() if v != base[k]}
        self.row.action = action
        self.row.check(self.app.history)
        self.row.fix_action()
        self.app.refresh_all()
        self.destroy()


# ---- Main window -------------------------------------------------------------

class App:
    """
    The main window: toolbar, review table and status bar.

    Flow: choose() -> load() builds the rows -> the user reviews and edits ->
    send() emails and prints -> finish_run() saves records -> reset().
    """

    COLS = ("action", "to", "names", "pages", "file", "notes")

    def __init__(self, root, cfg):
        self.root = root
        self.cfg = cfg                  # command line options: mode and folder paths
        self.rows = []
        self.history = {}               # recipient_history.json, loaded with each file
        self.defaults = load_defaults(cfg.logs)
        self.source = None              # the weekly PDF currently loaded
        self.key = None                 # its sent_files.json key
        self.mtime = None               # its modified time when loaded (change detection)
        self.work = None                # temp folder holding the built PDFs
        self.busy = False               # True while loading or sending
        self.ready = False              # False if the page/statement counts did not reconcile
        self.outlook = None
        self.accounts = []

        root.title(f"{APP_NAME} (PDF-Split-Send) v{APP_VERSION}" + ("" if cfg.mode == "send" else f"  [{cfg.mode.upper()} MODE]"))
        root.geometry("1200x650")
        root.minsize(900, 420)
        root.protocol("WM_DELETE_WINDOW", self.close)

        # Toolbar
        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")
        self.btn_choose = ttk.Button(top, text="1.  Choose file", command=self.choose, width=16)
        self.btn_choose.pack(side="left")
        self.btn_send = ttk.Button(top, text="2.  Send", command=self.send, state="disabled")
        self.btn_send.pack(side="left", padx=8)
        self.btn_edit = ttk.Button(top, text="Edit selected", command=self.edit, state="disabled")
        self.btn_edit.pack(side="left", padx=(16, 0))
        self.btn_settings = ttk.Button(top, text="Settings", command=lambda: SettingsDialog(self))
        self.btn_settings.pack(side="left", padx=8)
        self.from_label = ttk.Label(top)
        self.from_label.pack(side="left", padx=(4, 0))
        self.file_label = ttk.Label(top, text="No file chosen", foreground="#555")
        self.file_label.pack(side="left", padx=12)

        hint = ttk.Frame(root, padding=(10, 0, 10, 6))
        hint.pack(fill="x")
        ttk.Label(hint, foreground="#555",
                  text="Click the Action column to switch Email / Print / Skip.   "
                       "Double-click a row to open the statement.   "
                       "Select a row and click Edit to change To, From, file name, subject or body.").pack(side="left")

        # Review table. Each tree item id is the row's index in self.rows.
        mid = ttk.Frame(root, padding=(10, 0))
        mid.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(mid, columns=self.COLS, show="headings", selectmode="browse")
        for col, text, width, anchor, stretch in (
            ("action", "Action", 70, "center", False),
            ("to", "Recipient", 260, "w", True),
            ("names", "Statement for", 260, "w", True),
            ("pages", "Pages", 55, "center", False),
            ("file", "File name", 150, "w", False),
            ("notes", "Check this", 300, "w", True),
        ):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=anchor, stretch=stretch)
        self.tree.tag_configure("bad", background="#f8d7da")     # red
        self.tree.tag_configure("warn", background="#fff3cd")    # yellow
        self.tree.tag_configure("off", foreground="#888")        # skipped
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.update_buttons())

        # Status bar
        bottom = ttk.Frame(root, padding=10)
        bottom.pack(fill="x")
        self.status = ttk.Label(bottom, text="Click Choose file and pick this week's statement PDF.")
        self.status.pack(side="left")
        self.bar = ttk.Progressbar(bottom, length=220, mode="determinate")
        self.bar.pack(side="right")

        # Connect to Outlook up front so the From dropdown can list accounts.
        if cfg.mode != "dryrun":
            self.connect_outlook(quiet=False)
        self.update_from_label()

    # ---- helpers ----

    def connect_outlook(self, quiet):
        """Try to connect to Outlook and read its accounts. Returns True on success."""
        try:
            self.outlook = get_outlook()
            self.accounts = outlook_accounts(self.outlook)
            return True
        except Exception as e:
            self.outlook = None
            if not quiet:
                messagebox.showwarning("Outlook", f"Outlook could not be opened. Emailing won't work until it can.\n\n{e}")
            return False

    def update_from_label(self):
        """Show the default sender and delay next to Settings. Red when no sender is set."""
        sender = self.defaults.get("sender")
        delay = self.defaults.get("delay", 0)
        delay_text = f"   Delay: {delay} min" if delay else "   Delay: none"
        self.from_label.config(text=(f"From:  {sender}" if sender else "From:  none set (choose in Settings)") + delay_text,
                               foreground="#555" if sender else "#b00020")

    def set_status(self, text):
        """Update the status bar and repaint immediately (long tasks run on the UI thread)."""
        self.status.config(text=text)
        self.root.update()

    def step(self, n, total, text):
        """Advance the progress bar and status text."""
        self.bar.config(maximum=max(total, 1), value=n)
        self.set_status(text)

    def lock_ui(self, on):
        """Disable buttons while loading or sending so nothing can be clicked mid-task."""
        self.busy = on
        state = "disabled" if on else "normal"
        for b in (self.btn_choose, self.btn_settings):
            b.config(state=state)
        if on:
            self.btn_send.config(state="disabled")
            self.btn_edit.config(state="disabled")

    def clear_work(self):
        """Delete the temp folder of built PDFs."""
        if self.work:
            shutil.rmtree(self.work, ignore_errors=True)
            self.work = None

    def reset(self, status):
        """Empty the table and return to the "choose a file" state."""
        self.tree.delete(*self.tree.get_children())
        self.rows = []
        self.source = None
        self.ready = False
        self.clear_work()
        self.file_label.config(text="No file chosen")
        self.lock_ui(False)
        self.update_buttons()
        self.bar.config(value=0)
        self.set_status(status)

    def refresh_row(self, iid):
        """Redraw one table line from its Row."""
        r = self.rows[int(iid)]
        tags = [r.level] if r.level else []
        if r.action == SKIP:
            tags.append("off")
        file_name = clean_file_name(r.value("file_name", self.defaults)) if r.action == EMAIL else ""
        self.tree.item(iid, values=(r.action, r.label, r.names_text(), r.page_count or "",
                                    file_name, "; ".join(r.flags)), tags=tags)

    def refresh_all(self):
        """Redraw every table line and the button states."""
        for iid in self.tree.get_children():
            self.refresh_row(iid)
        self.update_buttons()

    def recheck_all(self):
        """Re-run the checks on every row (after Settings change), then redraw."""
        for r in self.rows:
            r.check(self.history)
            r.fix_action()
        self.refresh_all()

    def update_buttons(self):
        """Enable Send only when the file reconciled and there is something to do. Show the counts on it."""
        emails = sum(1 for r in self.rows if r.action == EMAIL)
        prints = sum(1 for r in self.rows if r.action == PRINT)
        ok = self.ready and (emails or prints) and not self.busy
        self.btn_send.config(state="normal" if ok else "disabled",
                             text=f"2.  Send  ({emails} email, {prints} print)" if self.rows else "2.  Send")
        self.btn_edit.config(state="normal" if self.tree.selection() and not self.busy else "disabled")

    def open_pdf(self, row):
        """Open a row's built PDF in the default PDF viewer."""
        if row.pdf and row.pdf.exists():
            if sys.platform == "win32":
                os.startfile(row.pdf)
            else:
                subprocess.Popen(["xdg-open", str(row.pdf)])

    # ---- choose file ----

    def choose(self):
        """Show the file picker in the statements folder with the newest PDF preselected."""
        start = Path(self.cfg.folder)
        newest = newest_pdf(start) if start.is_dir() else None
        path = filedialog.askopenfilename(
            title="Choose this week's statement file",
            initialdir=str(start),
            initialfile=newest.name if newest else "",
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self.load(Path(path))

    def load(self, path):
        """
        Read the weekly PDF, build one PDF per row in a temp folder, run the
        checks, fill the table, and reconcile page and statement counts.
        """
        # Warn before reopening a file that was already sent for real.
        key = file_key(path)
        prior = load_json(self.cfg.logs, SENT_FILES_NAME).get(key)
        if prior and not messagebox.askyesno(
                "Already sent",
                f"{path.name} was already sent on {prior['sent']} by {prior['user']}.\n\n"
                "Do you want to open it anyway?", icon="warning"):
            return

        self.reset("Reading file...")
        self.lock_ui(True)
        self.source, self.key, self.mtime = path, key, path.stat().st_mtime
        self.file_label.config(text=f"File:  {path.name}" + ("   (already sent before)" if prior else ""))

        try:
            total, rows, statements, expected = scan_file(
                path, lambda n, t: self.step(n, t, f"Reading pages... {n} of {t}"))
        except Exception as e:
            self.reset("Choose a different file.")
            messagebox.showerror("Can't read file", f"{path.name} could not be opened.\n\n{e}")
            return

        # Build each row's PDF in its own temp subfolder, then check it.
        self.history = load_json(self.cfg.logs, RECIPIENTS_NAME)
        self.work = Path(tempfile.mkdtemp(prefix="statements_"))
        for n, r in enumerate(rows, 1):
            self.step(n, len(rows), f"Building statements... {n} of {len(rows)}")
            d = self.work / f"{n:04d}"
            d.mkdir()
            r.pdf = d / "statement.pdf"
            try:
                r.page_count = build_pdf(r.pages, r.pdf)
                r.size_mb = r.pdf.stat().st_size / 1_048_576
            except Exception as e:
                r.pdf = None
                r.build_error = str(e)
            r.check(self.history)
            # Untagged pages default to Skip; everything else to its first allowed action.
            r.action = SKIP if r.kind == "untagged" else None
            r.fix_action()

        # Sort red rows first, then yellow, then the rest; untagged rows lead
        # their group; alphabetical within each.
        order = {"bad": 0, "warn": 1, "": 2}
        rows.sort(key=lambda r: (order[r.level], r.kind != "untagged", r.label.lower()))
        self.rows = rows
        for i, _ in enumerate(rows):
            self.tree.insert("", "end", iid=str(i))
            self.refresh_row(str(i))

        # Reconcile: every page must land in a built PDF, and the statement
        # count must match the totals page. Any mismatch blocks Send.
        built = sum(r.page_count for r in rows if r.pdf)
        problems = []
        if built != total:
            problems.append(f"The file has {total} pages but only {built} made it into statements.")
        if expected is not None and expected != statements:
            problems.append(f"The report totals page says {expected} statements, but {statements} were found.")
        self.ready = not problems
        self.lock_ui(False)
        self.bar.config(value=0)

        emails = sum(1 for r in rows if r.kind == "email")
        no_sender = sum(1 for r in rows if r.kind == "print")
        flagged = sum(1 for r in rows if r.level)
        msg = f"{statements} statements, {total} pages.   {emails} email recipients, {no_sender} No Sender."
        if flagged:
            msg += f"   {flagged} row(s) highlighted at the top to check."
        self.set_status(msg)
        self.update_buttons()

        if problems:
            messagebox.showerror("Count mismatch", "\n\n".join(problems) +
                                 "\n\nNothing can be sent. Please contact IT.")

    # ---- table clicks ----

    def on_click(self, event):
        """A click in the Action column cycles that row through its allowed actions."""
        if self.busy or self.tree.identify_column(event.x) != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        r = self.rows[int(iid)]
        options = r.allowed()
        r.action = options[(options.index(r.action) + 1) % len(options)] if r.action in options else options[0]
        self.refresh_row(iid)
        self.tree.selection_set(iid)
        self.update_buttons()
        return "break"      # stop Tk's default click handling for this column

    def on_double(self, event):
        """A double-click anywhere except the Action column opens the row's PDF."""
        iid = self.tree.identify_row(event.y)
        if not iid or self.tree.identify_column(event.x) == "#1":
            return
        self.open_pdf(self.rows[int(iid)])

    def edit(self):
        """Open the Edit dialog for the selected row."""
        sel = self.tree.selection()
        if sel and not self.busy:
            EditDialog(self, self.rows[int(sel[0])])

    # ---- send ----

    def send(self):
        """
        Run every safety check, confirm with the user, then email and print
        each row. Failures are recorded per row; one failure never stops the run.
        """
        d = self.defaults
        emails = [r for r in self.rows if r.action == EMAIL]
        prints = [r for r in self.rows if r.action == PRINT]
        skips = [r for r in self.rows if r.action == SKIP]
        if not (emails or prints) or not self.source:
            return

        # The built PDFs came from the file as it was at load time. If it has
        # changed or moved since, make the user reload it.
        if not self.source.exists() or self.source.stat().st_mtime != self.mtime:
            messagebox.showwarning("File changed",
                                   f"{self.source.name} changed or moved after it was opened.\n\n"
                                   "Click Choose file and open it again.")
            self.ready = False
            self.update_buttons()
            return

        # "live" means Outlook and the printer are really used.
        live = self.cfg.mode in ("send", "display")
        if live and emails and not self.outlook and not self.connect_outlook(quiet=False):
            return

        # Every email needs an explicit sender. Checked in every mode.
        no_from = [r for r in emails if not r.value("sender", d)]
        if no_from:
            messagebox.showerror("No sender", f"{len(no_from)} email(s) have no From address.\n\n"
                                 "Pick a sender in Settings (applies to everyone) or with Edit for one row, "
                                 "then click Send again.")
            return

        # Every sender must exist in this user's Outlook (settings are shared,
        # and the other user's Outlook may have different accounts).
        senders = sorted({r.value("sender", d) for r in emails}, key=str.lower)
        if live and emails:
            known = {a.lower() for a in self.accounts}
            missing = [s for s in senders if s.lower() not in known]
            if missing:
                messagebox.showerror("From account", "These From accounts aren't set up in this Outlook:\n\n"
                                     + "\n".join(missing) + "\n\nChange them in Settings or Edit.")
                return

        printer = None
        if live and prints:
            try:
                printer = default_printer()
            except Exception as e:
                messagebox.showerror("Printer", f"No default printer could be found.\n\n{e}")
                return

        # Confirmation with counts, senders, printer and delay.
        lines = []
        if emails:
            lines.append(f"Email: {len(emails)}   (from {', '.join(senders)})")
        if prints:
            lines.append(f"Print: {len(prints)}" + (f"   (to {printer})" if printer else ""))
        if skips:
            lines.append(f"Skip: {len(skips)}")
        delay = d["delay"]
        if self.cfg.mode == "send" and emails:
            lines.append(f"\nEmails go out in {delay} minutes, so you can still stop one from the Outbox."
                         if delay else "\nEmails go out right away.")
        if not messagebox.askyesno(APP_NAME, "Go ahead?\n\n" + "\n".join(lines)):
            return

        self.lock_ui(True)
        for r in skips:
            r.status = "Skipped"
        todo = emails + prints
        for n, r in enumerate(todo, 1):
            self.step(n, len(todo), f"{'Emailing' if r.action == EMAIL else 'Printing'}... {n} of {len(todo)}")
            try:
                if r.action == EMAIL:
                    # Copy the built PDF under the chosen attachment name. Each
                    # email gets its own subfolder so names never collide.
                    r.sent_name = clean_file_name(r.value("file_name", d))
                    attach = self.work / f"send_{n:04d}" / r.sent_name
                    attach.parent.mkdir()
                    shutil.copy2(r.pdf, attach)
                    if live:
                        send_mail(self.outlook, r.to, r.value("sender", d), fill(r.value("subject", d)),
                                  fill(r.value("body", d)), attach, self.cfg.mode, delay)
                    r.status = {"send": "Emailed", "display": "Displayed"}.get(self.cfg.mode, "Would email")
                else:
                    if live:
                        print_pdf(r.pdf, r.file_base)
                    r.status = "Printed" if live else "Would print"
            except Exception as e:
                # Record and keep going so one bad row does not stop the rest.
                r.status = f"Error: {e}"

        self.set_status("Finishing up...")
        try:
            run = finish_run(self.cfg, self.rows, self.source, self.key, self.cfg.mode, d)
        except Exception as e:
            run = None
            messagebox.showerror("Wrap-up problem", f"Statements were handled, but saving the records failed.\n\n{e}")

        emailed = sum(1 for r in self.rows if r.status in EMAILED)
        printed = sum(1 for r in self.rows if r.status in PRINTED)
        errors = [r for r in self.rows if r.status.startswith("Error")]

        self.reset("Done. Click Choose file to start the next one.")

        # Summary
        out = [f"Emailed: {emailed}", f"Printed: {printed}", f"Skipped: {len(skips)}"]
        if errors:
            out.append(f"Errors: {len(errors)}  (saved in the Not Sent folder)")
            out += [f"   {r.file_base}: {r.status[7:]}" for r in errors[:5]]
        if self.cfg.mode == "send" and emailed and delay:
            out.append(f"\nEmails will leave the Outbox at about {datetime.now() + timedelta(minutes=delay):%I:%M %p}.")
        if run:
            out.append(f"\nSaved to:\n{run}")
        (messagebox.showwarning if errors else messagebox.showinfo)("Done", "\n".join(out))

    # ---- close ----

    def close(self):
        """Close the window, unless a load or send is in progress. Deletes the temp PDFs."""
        if self.busy:
            return
        self.clear_work()
        self.root.destroy()


# ---- Main --------------------------------------------------------------------

def main():
    """Parse flags, check the share, take the lock, and run the window."""
    ap = argparse.ArgumentParser(description=f"{APP_NAME} (PDF-Split-Send)")
    ap.add_argument("--mode", choices=["send", "display", "dryrun"], default="send",
                    help="send: real run (default). display: open emails instead of sending. dryrun: no Outlook or printing.")
    ap.add_argument("--folder", default=STATEMENTS_FOLDER, help="Folder the file picker opens in; also holds the lock file.")
    ap.add_argument("--output", default=OUTPUT_FOLDER, help="Where Runs and Recipients folders are created.")
    ap.add_argument("--logs", default=LOG_FOLDER, help="Where settings, history and logs are kept.")
    # parse_known_args ignores anything unexpected (some launchers add their own arguments).
    cfg, _ = ap.parse_known_args()

    # Crisp text on high-DPI screens instead of Windows bitmap scaling.
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    # Hidden root window so message boxes can show before the main window does.
    root = tk.Tk()
    root.withdraw()

    if not Path(cfg.folder).is_dir():
        messagebox.showerror(APP_NAME, f"Can't reach the statements folder:\n\n{cfg.folder}")
        root.destroy()
        return

    lock, info = acquire_lock(cfg.folder)
    if not lock:
        messagebox.showwarning(APP_NAME,
                               f"Statements are already being sent by {info}.\n\nTry again when they are done.")
        root.destroy()
        return

    try:
        root.deiconify()
        App(root, cfg)
        root.mainloop()
    finally:
        # Always release the lock, even if the app crashes.
        release_lock(lock)


if __name__ == "__main__":
    main()
