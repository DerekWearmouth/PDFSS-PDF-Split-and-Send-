# PDFSS (PDF-Split-Send)

Splits, reviews and sends the weekly statement run.

Each week, one large PDF of all statements is saved to the statements share. Every page carries a hidden tag at the top left (the exact format is set by `TAG_RE`), for example:

- `<tag>ID#someone@example.com</tag>` for statements to email
- `<tag>ID#PRINT</tag>` for statements with no email (No Sender)

The app splits the PDF, combines every statement for the same email into one PDF, and lets the user review everything. Then it emails through their Outlook, prints to the default printer, or skips, row by row.

---

## Requirements

- Windows with Outlook (desktop) installed and signed in
- Python 3 (run with `py`)
- Python packages:

```
py -m pip install pymupdf pywin32 pillow
```

`pywin32` and `pillow` are only needed for `send` and `display` modes. A dry run only needs `pymupdf`.

Users of the .exe need none of this (see Building the .exe).

---

## Running it

Normal run (emails and prints for real):

```
py PDFSS.py
```

Test run in a local folder, with no Outlook and no printing:

```
py PDFSS.py --mode dryrun --folder C:\StatementTest --output C:\StatementTest\Output --logs C:\StatementTest\Logs
```

Open each email in Outlook without sending, and print for real (only use this on a small test file, since it opens one window per email):

```
py PDFSS.py --mode display --folder C:\StatementTest --output C:\StatementTest\Output --logs C:\StatementTest\Logs
```

The app won't start if it can't reach `--folder`. On a PC without the share, always pass all three folder flags.

---

## Command line flags

| Flag | Default | What it does |
|---|---|---|
| `--mode` | `send` | `send`, `display` or `dryrun` (see below) |
| `--folder` | `STATEMENTS_FOLDER` | Folder the file picker opens in. The lock file also lives here. |
| `--output` | `OUTPUT_FOLDER` | Where run folders and per-recipient folders are created |
| `--logs` | `LOG_FOLDER` | Shared settings, history, sent-file list and run logs |

### Modes

| Mode | Email rows | Print rows | Output | History and sent-file list |
|---|---|---|---|---|
| `send` | Sent | Printed | `Output\Runs\...` and `Output\Recipients\...` | Updated |
| `display` | Opened in Outlook, not sent | Printed | Run folder ending `(display test)`, with its own Recipients folder inside | Not updated |
| `dryrun` | Not touched | Not printed | Run folder ending `(dryrun test)`, with its own Recipients folder inside | Not updated |

In `display` and `dryrun` modes, the window title shows the mode, for example **PDFSS (PDF-Split-Send) v1.0 [DRYRUN MODE]**.

---

## How a run works

1. Click **1. Choose file**. The picker opens in the statements folder with the newest PDF already selected. They can also browse to any other folder, such as the Desktop.
2. The preview loads with one row per recipient. Rows that need a look are at the top:
   - **Red**: can't be emailed as-is (No Sender, bad address, too large, untagged page)
   - **Yellow**: worth a look (new recipient, page count jump)
   - **Grey text**: set to Skip
3. Click the **Action** cell to cycle through **Email / Print / Skip**. Only the options that make sense for that row are offered.
4. Double-click a row to open its statement.
5. Select a row and click **Edit selected** to change that recipient's Action, To, From, file name, subject or body. **Reset to defaults** undoes the edits.
6. Click **2. Send (N email, N print)**. A confirmation shows the counts, the From accounts and the printer.
7. A summary shows emailed, printed, skipped, any errors, and where everything was saved.

If the file was sent before, a warning shows the date and user and asks whether to open it anyway.

### Settings button

Sets the defaults for every recipient: **From** account, **file name**, **subject**, **body** and **send delay**.

- From lists the accounts in this user's Outlook. It starts as `(choose a sender)`, and **nothing can be sent until a sender is picked**. The app never falls back to Outlook's default account.
- The current default sender and delay show next to the Settings button. The sender is red when none is set.
- **Send delay** is how many minutes emails wait in the Outbox before going out (0 to 1440). `0` sends right away. The confirmation and summary show when they'll leave.
- `{date}` in the subject or body becomes today's date (`11/06/2026`). In the file name it becomes `11-06-2026`.
- `.pdf` is added to the file name automatically if it's missing.
- Settings are saved to `settings.json` in the logs folder, so both users share them.
- Anything changed with Edit on a single row overrides these defaults for that row only.

### Row types

| Row | Default action | Notes |
|---|---|---|
| Email address | Email | All statements for the same email are combined into one PDF |
| No Sender | Print | Statements tagged `PRINT`. One row per statement. A To address can be typed in Edit to email it instead. |
| Untagged pages | Skip | Pages with no tag at all, such as the report totals page at the end |

---

## Checks

| Check | What happens |
|---|---|
| No email | Row is red and defaults to Print |
| To address looks wrong | Row is red. Email isn't allowed until it's fixed in Edit. |
| Too large to email | Over `MAX_ATTACHMENT_MB`. Email isn't allowed. Print or skip. |
| Could not build PDF | Row can only be skipped |
| New recipient | Yellow. This address has never been emailed before. |
| Pages up from X last time | Yellow. Page count jumped well above this person's last run. |
| Edited | Shown when a row has its own To, From, file name, subject or body |
| Page count | If the pages in the file don't match the pages in the rows, Send is blocked |
| Statement count | If the report totals page count doesn't match the statements found, Send is blocked |
| File changed | If the PDF changes or moves after it was opened, Send is blocked until it's opened again |
| No sender | If any Email row has no From address, Send stops before anything goes out (prints included) |
| From account | If a From account isn't set up in this user's Outlook, Send stops before anything goes out |
| One user at a time | A second user gets "already being sent by [name]" |
| Delayed delivery | In `send` mode, emails wait the Settings send delay in the Outbox so one can still be pulled back |

---

## Output and log files

```
Output\
  Runs\
    2026-11-06 wk_21_yr_26_pre-final\
      Emailed\
        joe@example.com.pdf
      Printed\
        No Sender - ACME COMPANY.pdf
      Not Sent\                       (skipped rows and errors)
        Untagged pages.pdf
      send_log.csv
  Recipients\
    joe@example.com\
      2026-10-30 statement.pdf
      2026-11-06 statement.pdf
```

- **Runs** has one folder per run with every statement that was built.
- **Recipients** has one folder per email address. Each emailed statement is saved as `<date> <file name sent>`, so older statements are easy to find later. If a row's To was edited, the copy goes under the edited address.

The logs folder holds:

| File | Purpose |
|---|---|
| `settings.json` | Default From, file name, subject, body and send delay set from the Settings button. Delete it to reset to no sender. |
| `recipient_history.json` | Every address emailed and its last page count. Drives the New recipient and page jump flags. |
| `sent_files.json` | Weekly files already sent, with date, user and run folder. Drives the already-sent warning. |
| `send_log_<timestamp>.csv` | Full log of each run: action, To, From, file name, subject, statements, pages, notes, status |

The statements folder holds `.pdfss.lock` while the app is open. It's removed on close, and a lock older than `STALE_LOCK_HOURS` is cleared automatically.

---

## Script settings (top of PDFSS.py)

Email defaults are set in the app with the Settings button. These are the fallback values and the technical settings.

| Setting | Default | What it does |
|---|---|---|
| `STATEMENTS_FOLDER` | `\\fileserver\share\Statements` | Where the file picker opens and the lock file lives |
| `OUTPUT_FOLDER` | `...\Statements\Output` | Runs and Recipients folders |
| `LOG_FOLDER` | `...\Statements\Logs` | Settings, history and logs |
| `DEFAULT_FILE_NAME` | `statement.pdf` | Used until Settings is saved |
| `DEFAULT_SUBJECT` | `Statements {date}` | Used until Settings is saved |
| `DEFAULT_BODY` | `Please find your statement(s) attached.` | Used until Settings is saved |
| `DATE_FORMAT` | `%m/%d/%Y` | `{date}` in the subject and body |
| `FILE_DATE_FORMAT` | `%m-%d-%Y` | `{date}` in the file name |
| `HEADER_BAND` | `0.20` | Top part of the page searched for the tag (0.20 = top 20%) |
| `TAG_RE` | Set in the script | Pattern for the hidden tag |
| `PRINT_VALUE` | `print` | Tag value that means No Sender |
| `START_RE` | `Page 1` | Marks the first page of each statement |
| `TOTALS_RE` | `Report Totals: N` | Statement count on the totals page, used for the statement count check |
| `CARRY_FORWARD` | `False` | Leave off. Every statement page is tagged, so untagged pages go to their own row instead of joining the previous statement. |
| `MOVE_AFTER_SEND` | `False` | Move the weekly PDF into its run folder after sending |
| `MAX_ATTACHMENT_MB` | `20` | Largest attachment allowed |
| `DELAY_MINUTES` | `10` | Send delay used until one is saved in Settings |
| `PRINT_DPI` | `300` | Print quality |
| `JUMP_FACTOR` / `JUMP_MIN_PAGES` | `2.0` / `5` | Page jump flag: at least 2x last run and at least 5 more pages |
| `STALE_LOCK_HOURS` | `4` | Lock files older than this are cleared |

---

## Building the .exe

1. Set `STATEMENTS_FOLDER`, `OUTPUT_FOLDER` and `LOG_FOLDER` to the real share paths. They are built into the .exe.
2. Put `build.bat` next to `PDFSS.py` and double-click it. The app is created at `dist\PDFSS.exe`.
3. Sign it with the internal code-signing cert, or allowlist it in Defender and AppLocker.
4. Copy it to the share and give each user a desktop shortcut.

To test the .exe, make a separate shortcut with the flags at the end of Target:

```
"PDFSS.exe" --mode dryrun --folder C:\StatementTest --output C:\StatementTest\Output --logs C:\StatementTest\Logs
```

After any change to the script, run `build.bat` again and replace the .exe on the share.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `python` opens the Microsoft Store or isn't found | Use `py` instead |
| `pip` is not recognized | Use `py -m pip install ...` |
| `No module named 'pymupdf'` | `py -m pip install pymupdf` |
| `No module named 'win32com'`, `win32print` or `PIL` | `py -m pip install pywin32 pillow` |
| "Can't reach the statements folder" and it closes | The share isn't reachable from this PC. For testing, pass `--folder`, `--output` and `--logs` with local paths. |
| "Already being sent by ..." | Someone else has it open. If not, delete `.pdfss.lock` from the statements folder. |
| "email(s) have no From address" | Pick a sender in Settings, or for that row with Edit |
| "From accounts aren't set up in this Outlook" | That user's Outlook doesn't have the account. Pick another From in Settings or Edit, or add the account to Outlook. |
| "No default printer could be found" | Set a default printer in Windows |
| Count mismatch on preview | The file may be incomplete or in a new layout. Don't send. Check the PDF. |
| Emails don't leave the Outbox | With delayed delivery on a non-Exchange account, Outlook must stay open until the delay passes |
| Every recipient flagged as new | Expected on the first real run, since there's no history yet |
