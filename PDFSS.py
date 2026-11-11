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
#  ____  ____  _____ ____ ____
# |  _ \|  _ \|  ___/ ___/ ___|
# | |_) | | | | |_  \___ \___ \
# |  __/| |_| |  _|  ___) |__) |
# |_|   |____/|_|   |____/____/
#================================================================================
# PDFSS (PDF-Split-Send) - v1.0 Date: 2026-11-11
#================================================================================

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


# ==============================================================================
# Settings
# ==============================================================================

APP_NAME = "PDFSS"
APP_VERSION = "1.0"

STATEMENTS_FOLDER = r"\\fileserver\share\Statements"
OUTPUT_FOLDER = r"\\fileserver\share\Statements\Output"
LOG_FOLDER = r"\\fileserver\share\Statements\Logs"

DEFAULT_FILE_NAME = "statement.pdf"
DEFAULT_SUBJECT = "Statements {date}"
DEFAULT_BODY = "Please find your statement(s) attached."
DATE_FORMAT = "%m/%d/%Y"
FILE_DATE_FORMAT = "%m-%d-%Y"

HEADER_BAND = 0.20
TAG_RE = re.compile(r"<pdfexplode>(?:[^#<]*#)?([^<]+)</pdfexplode>", re.I)
PRINT_VALUE = "print"
START_RE = re.compile(r"Page\s+1\b")
TOTALS_RE = re.compile(r"Report Totals:\s*(\d+)")
CARRY_FORWARD = False

MOVE_AFTER_SEND = False
MAX_ATTACHMENT_MB = 20
DELAY_MINUTES = 10
MAX_DELAY_MINUTES = 1440
PRINT_DPI = 300
JUMP_FACTOR = 2.0
JUMP_MIN_PAGES = 5
STALE_LOCK_HOURS = 4

NO_SENDER = "No Sender"
NO_FROM = "(choose a sender)"
LOCK_NAME = ".pdfss.lock"
SETTINGS_NAME = "settings.json"
RECIPIENTS_NAME = "recipient_history.json"
SENT_FILES_NAME = "sent_files.json"
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
VALID_EMAIL = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

EMAIL, PRINT, SKIP = "Email", "Print", "Skip"
EMAILED = ("Emailed", "Displayed", "Would email")
PRINTED = ("Printed", "Would print")


# ==============================================================================
# Text helpers
# ==============================================================================

def fill(text, for_file=False):
    return text.replace("{date}", datetime.now().strftime(FILE_DATE_FORMAT if for_file else DATE_FORMAT))


def clean_file_name(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", fill(name, True)).strip().strip(".")
    name = name or DEFAULT_FILE_NAME
    return name if name.lower().endswith(".pdf") else name + ".pdf"


def safe_name(text):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", text).strip().strip(".") or "unnamed"


def unique_path(path):
    path = Path(path)
    n = 2
    while path.exists():
        path = path.with_name(f"{path.stem.rsplit(' (', 1)[0]} ({n}){path.suffix}")
        n += 1
    return path


# ==============================================================================
# Row
# ==============================================================================

class Row:
    def __init__(self, found, kind):
        self.found = found
        self.kind = kind
        self.pages = []
        self.names = []
        self.pdf = None
        self.page_count = 0
        self.size_mb = 0.0
        self.build_error = None
        self.over = {}
        self.action = None
        self.status = ""
        self.sent_name = ""
        self.flags = []
        self.level = ""

    @property
    def to(self):
        return self.over.get("to", self.found or "").strip()

    @property
    def label(self):
        if self.to:
            return self.to
        return "Untagged pages" if self.kind == "untagged" else NO_SENDER

    @property
    def file_base(self):
        if self.to:
            return self.to
        return f"{self.label} - {self.names[0]}" if self.names else self.label

    def names_text(self):
        if not self.names:
            return ""
        return self.names[0] + (f"  (+{len(self.names) - 1} more)" if len(self.names) > 1 else "")

    def value(self, key, defaults):
        return self.over.get(key, defaults[key])

    def can_email(self):
        return bool(self.pdf) and bool(VALID_EMAIL.match(self.to)) and self.size_mb <= MAX_ATTACHMENT_MB

    def allowed(self):
        return [a for a, ok in ((EMAIL, self.can_email()), (PRINT, bool(self.pdf)), (SKIP, True)) if ok]

    def fix_action(self):
        if self.action not in self.allowed():
            self.action = self.allowed()[0]

    def check(self, history):
        self.flags, self.level = [], ""

        def add(text, level=None):
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
                last = seen.get("last_pages", 0)
                if last and self.page_count >= last * JUMP_FACTOR and self.page_count - last >= JUMP_MIN_PAGES:
                    add(f"Pages up from {last} last time", "warn")
        if self.over:
            add("Edited")


# ==============================================================================
# Lock
# ==============================================================================

def acquire_lock(folder):
    path = Path(folder) / LOCK_NAME
    if path.exists() and time.time() - path.stat().st_mtime > STALE_LOCK_HOURS * 3600:
        path.unlink(missing_ok=True)
    info = f"{getpass.getuser()} on {socket.gethostname()} since {datetime.now():%I:%M %p}"
    try:
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
    if path:
        Path(path).unlink(missing_ok=True)


# ==============================================================================
# Stored data
# ==============================================================================

def load_json(folder, name):
    try:
        return json.loads((Path(folder) / name).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_json(folder, name, data):
    Path(folder).mkdir(parents=True, exist_ok=True)
    path = Path(folder) / name
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=1, sort_keys=True), encoding="utf-8")
    os.replace(tmp, path)


def load_defaults(folder):
    d = {"sender": "", "file_name": DEFAULT_FILE_NAME, "subject": DEFAULT_SUBJECT, "body": DEFAULT_BODY,
         "delay": DELAY_MINUTES}
    d.update({k: v for k, v in load_json(folder, SETTINGS_NAME).items() if k in d})
    try:
        d["delay"] = min(max(int(d["delay"]), 0), MAX_DELAY_MINUTES)
    except (TypeError, ValueError):
        d["delay"] = DELAY_MINUTES
    return d


def file_key(path):
    return f"{Path(path).name}|{Path(path).stat().st_size}"


# ==============================================================================
# PDF scanning and building
# ==============================================================================

def newest_pdf(folder):
    files = [p for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def page_info(page):
    r = page.rect
    header = page.get_text("text", clip=pymupdf.Rect(r.x0, r.y0, r.x1, r.y0 + r.height * HEADER_BAND))
    m = TAG_RE.search(header)
    value = m.group(1).strip().lower() if m else None
    if value is None:
        e = EMAIL_RE.search(header)
        value = e.group(0).lower() if e else None
    start = bool(START_RE.search(header))
    name = next((ln.strip() for ln in header.splitlines() if ln.strip()), "")
    return value, start, name


def scan_file(path, progress=None):
    groups = OrderedDict()
    doc = pymupdf.open(path)
    total = doc.page_count
    current = None
    statements = 0
    expected = None
    for i, page in enumerate(doc):
        value, start, name = page_info(page)
        if value and value != PRINT_VALUE:
            row = groups.get(value) or groups.setdefault(value, Row(value, "email"))
        elif value == PRINT_VALUE:
            if start or current is None or current.kind != "print":
                row = groups.setdefault(("print", i), Row(None, "print"))
            else:
                row = current
        elif CARRY_FORWARD and current is not None and not start:
            row = current
        else:
            if current is not None and current.kind == "untagged":
                row = current
            else:
                row = groups.setdefault(("untagged", i), Row(None, "untagged"))
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
    out = pymupdf.open()
    srcs = {}
    try:
        for f, i in pages:
            if f not in srcs:
                srcs[f] = pymupdf.open(f)
            out.insert_pdf(srcs[f], from_page=i, to_page=i)
        out.save(out_path, garbage=3, deflate=True)
        return out.page_count
    finally:
        out.close()
        for d in srcs.values():
            d.close()


# ==============================================================================
# Outlook
# ==============================================================================

def get_outlook():
    import win32com.client
    return win32com.client.Dispatch("Outlook.Application")


def account_name(acct):
    return (acct.SmtpAddress or acct.DisplayName or "").strip()


def outlook_accounts(outlook):
    accts = outlook.Session.Accounts
    return [n for n in (account_name(accts.Item(i)) for i in range(1, accts.Count + 1)) if n]


def find_account(outlook, name):
    accts = outlook.Session.Accounts
    for i in range(1, accts.Count + 1):
        if account_name(accts.Item(i)).lower() == name.lower():
            return accts.Item(i)
    return None


def send_mail(outlook, to, sender, subject, body, attachment, mode, delay):
    if not sender:
        raise RuntimeError("No From address")
    acct = find_account(outlook, sender)
    if acct is None:
        raise RuntimeError(f"From account {sender} is not set up in this Outlook")
    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, acct))
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(str(attachment))
    if mode == "display":
        mail.Display()
        return
    if delay > 0:
        when = datetime.now() + timedelta(minutes=delay)
        mail.DeferredDeliveryTime = when.strftime("%m/%d/%Y %I:%M:%S %p")
    mail.Send()


# ==============================================================================
# Printing
# ==============================================================================

def default_printer():
    import win32print
    return win32print.GetDefaultPrinter()


def print_pdf(path, title):
    import win32print
    import win32ui
    from PIL import Image, ImageWin

    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(win32print.GetDefaultPrinter())
    width, height = dc.GetDeviceCaps(8), dc.GetDeviceCaps(10)
    doc = pymupdf.open(path)
    try:
        dc.StartDoc(title)
        for page in doc:
            pix = page.get_pixmap(dpi=PRINT_DPI, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            if (img.width > img.height) != (width > height):
                img = img.rotate(90, expand=True)
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


# ==============================================================================
# Wrap-up
# ==============================================================================

def finish_run(cfg, rows, source, key, mode, defaults):
    now = datetime.now()
    stamp = now.strftime("%Y-%m-%d_%H%M%S")
    user = getpass.getuser()
    test = mode != "send"

    name = f"{now:%Y-%m-%d} {source.stem}" + (f" ({mode} test)" if test else "")
    run = Path(cfg.output) / "Runs" / name
    if run.exists():
        run = Path(cfg.output) / "Runs" / f"{name} {now:%H%M%S}"
    run.mkdir(parents=True)
    people = (run / "Recipients") if test else (Path(cfg.output) / "Recipients")

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
        history = load_json(cfg.logs, RECIPIENTS_NAME)
        for r in rows:
            if r.status == "Emailed":
                history[r.to.lower()] = {"last_pages": r.page_count, "last_sent": stamp}
        save_json(cfg.logs, RECIPIENTS_NAME, history)

        sent_files = load_json(cfg.logs, SENT_FILES_NAME)
        sent_files[key] = {"sent": now.strftime("%m/%d/%Y %I:%M %p"), "user": user, "output": run.name}
        save_json(cfg.logs, SENT_FILES_NAME, sent_files)

        if MOVE_AFTER_SEND:
            shutil.move(str(source), str(run / source.name))

    return run


# ==============================================================================
# Message form
# ==============================================================================

def message_form(frame, start, accounts, values, include_to):
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
    frame.form_vars = [v for v in (locals().get("v_to"), v_from, v_file, v_subj) if v is not None]
    return getters, r + 1


class Dialog(tk.Toplevel):
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
        self.update_idletasks()
        x = self.app.root.winfo_rootx() + 80
        y = self.app.root.winfo_rooty() + 60
        self.geometry(f"+{x}+{y}")
        self.grab_set()
        self.wait_window()


# ==============================================================================
# Settings dialog
# ==============================================================================

class SettingsDialog(Dialog):
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
        self.app.recheck_all()
        self.destroy()


# ==============================================================================
# Edit dialog
# ==============================================================================

class EditDialog(Dialog):
    def __init__(self, app, row):
        super().__init__(app, f"Edit  {row.file_base}")
        self.row = row
        d = app.defaults

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

        values = {"to": row.to, "sender": row.value("sender", d), "file_name": row.value("file_name", d),
                  "subject": row.value("subject", d), "body": row.value("body", d)}
        self.getters, _ = message_form(self.body, 2, app.accounts, values, include_to=True)

        ttk.Button(self.buttons, text="Open statement", command=lambda: app.open_pdf(row)).pack(side="left")
        ttk.Button(self.buttons, text="Reset to defaults", command=self.reset).pack(side="left", padx=8)
        ttk.Button(self.buttons, text="Save", command=self.save).pack(side="right")
        ttk.Button(self.buttons, text="Cancel", command=self.destroy).pack(side="right", padx=8)
        self.show()

    def reset(self):
        self.row.over = {}
        self.row.action = None
        self.row.check(self.app.history)
        self.row.fix_action()
        if self.row.kind == "untagged":
            self.row.action = SKIP
        self.app.refresh_all()
        self.destroy()

    def save(self):
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

        base = {"to": self.row.found or "", **self.app.defaults}
        self.row.over = {k: v for k, v in values.items() if v != base[k]}
        self.row.action = action
        self.row.check(self.app.history)
        self.row.fix_action()
        self.app.refresh_all()
        self.destroy()


# ==============================================================================
# Main window
# ==============================================================================

class App:
    COLS = ("action", "to", "names", "pages", "file", "notes")

    def __init__(self, root, cfg):
        self.root = root
        self.cfg = cfg
        self.rows = []
        self.history = {}
        self.defaults = load_defaults(cfg.logs)
        self.source = None
        self.key = None
        self.mtime = None
        self.work = None
        self.busy = False
        self.ready = False
        self.outlook = None
        self.accounts = []

        root.title(f"{APP_NAME} (PDF-Split-Send) v{APP_VERSION}" + ("" if cfg.mode == "send" else f"  [{cfg.mode.upper()} MODE]"))
        root.geometry("1200x650")
        root.minsize(900, 420)
        root.protocol("WM_DELETE_WINDOW", self.close)

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
        self.tree.tag_configure("bad", background="#f8d7da")
        self.tree.tag_configure("warn", background="#fff3cd")
        self.tree.tag_configure("off", foreground="#888")
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.update_buttons())

        bottom = ttk.Frame(root, padding=10)
        bottom.pack(fill="x")
        self.status = ttk.Label(bottom, text="Click Choose file and pick this week's statement PDF.")
        self.status.pack(side="left")
        self.bar = ttk.Progressbar(bottom, length=220, mode="determinate")
        self.bar.pack(side="right")

        if cfg.mode != "dryrun":
            self.connect_outlook(quiet=False)
        self.update_from_label()

    # ---- helpers ----

    def connect_outlook(self, quiet):
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
        sender = self.defaults.get("sender")
        delay = self.defaults.get("delay", 0)
        delay_text = f"   Delay: {delay} min" if delay else "   Delay: none"
        self.from_label.config(text=(f"From:  {sender}" if sender else "From:  none set (choose in Settings)") + delay_text,
                               foreground="#555" if sender else "#b00020")

    def set_status(self, text):
        self.status.config(text=text)
        self.root.update()

    def step(self, n, total, text):
        self.bar.config(maximum=max(total, 1), value=n)
        self.set_status(text)

    def lock_ui(self, on):
        self.busy = on
        state = "disabled" if on else "normal"
        for b in (self.btn_choose, self.btn_settings):
            b.config(state=state)
        if on:
            self.btn_send.config(state="disabled")
            self.btn_edit.config(state="disabled")

    def clear_work(self):
        if self.work:
            shutil.rmtree(self.work, ignore_errors=True)
            self.work = None

    def reset(self, status):
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
        r = self.rows[int(iid)]
        tags = [r.level] if r.level else []
        if r.action == SKIP:
            tags.append("off")
        file_name = clean_file_name(r.value("file_name", self.defaults)) if r.action == EMAIL else ""
        self.tree.item(iid, values=(r.action, r.label, r.names_text(), r.page_count or "",
                                    file_name, "; ".join(r.flags)), tags=tags)

    def refresh_all(self):
        for iid in self.tree.get_children():
            self.refresh_row(iid)
        self.update_buttons()

    def recheck_all(self):
        for r in self.rows:
            r.check(self.history)
            r.fix_action()
        self.refresh_all()

    def update_buttons(self):
        emails = sum(1 for r in self.rows if r.action == EMAIL)
        prints = sum(1 for r in self.rows if r.action == PRINT)
        ok = self.ready and (emails or prints) and not self.busy
        self.btn_send.config(state="normal" if ok else "disabled",
                             text=f"2.  Send  ({emails} email, {prints} print)" if self.rows else "2.  Send")
        self.btn_edit.config(state="normal" if self.tree.selection() and not self.busy else "disabled")

    def open_pdf(self, row):
        if row.pdf and row.pdf.exists():
            if sys.platform == "win32":
                os.startfile(row.pdf)
            else:
                subprocess.Popen(["xdg-open", str(row.pdf)])

    # ---- choose file ----

    def choose(self):
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
            r.action = SKIP if r.kind == "untagged" else None
            r.fix_action()

        order = {"bad": 0, "warn": 1, "": 2}
        rows.sort(key=lambda r: (order[r.level], r.kind != "untagged", r.label.lower()))
        self.rows = rows
        for i, _ in enumerate(rows):
            self.tree.insert("", "end", iid=str(i))
            self.refresh_row(str(i))

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
        return "break"

    def on_double(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid or self.tree.identify_column(event.x) == "#1":
            return
        self.open_pdf(self.rows[int(iid)])

    def edit(self):
        sel = self.tree.selection()
        if sel and not self.busy:
            EditDialog(self, self.rows[int(sel[0])])

    # ---- send ----

    def send(self):
        d = self.defaults
        emails = [r for r in self.rows if r.action == EMAIL]
        prints = [r for r in self.rows if r.action == PRINT]
        skips = [r for r in self.rows if r.action == SKIP]
        if not (emails or prints) or not self.source:
            return

        if not self.source.exists() or self.source.stat().st_mtime != self.mtime:
            messagebox.showwarning("File changed",
                                   f"{self.source.name} changed or moved after it was opened.\n\n"
                                   "Click Choose file and open it again.")
            self.ready = False
            self.update_buttons()
            return

        live = self.cfg.mode in ("send", "display")
        if live and emails and not self.outlook and not self.connect_outlook(quiet=False):
            return

        no_from = [r for r in emails if not r.value("sender", d)]
        if no_from:
            messagebox.showerror("No sender", f"{len(no_from)} email(s) have no From address.\n\n"
                                 "Pick a sender in Settings (applies to everyone) or with Edit for one row, "
                                 "then click Send again.")
            return

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
        if self.busy:
            return
        self.clear_work()
        self.root.destroy()


# ==============================================================================
# Main
# ==============================================================================

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["send", "display", "dryrun"], default="send")
    ap.add_argument("--folder", default=STATEMENTS_FOLDER)
    ap.add_argument("--output", default=OUTPUT_FOLDER)
    ap.add_argument("--logs", default=LOG_FOLDER)
    cfg, _ = ap.parse_known_args()

    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

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
        release_lock(lock)


if __name__ == "__main__":
    main()
