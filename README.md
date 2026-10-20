# AutoReport

Windows desktop tool for the POS support team. It signs in against Firebase,
pulls an Azure DevOps (TFS) work item plus the signed-in user's latest comment,
and turns that into the paperwork the team files back into TFS:

- **Generate Status Report** - fills `Report Template.docx` with the ticket
  number, subject, date and the `Root Cause / Preventive Action / Next Step /
  Status` fields scraped from the comment, then saves a `.docx`.
- **Create Azure Status Report** - files a Product Backlog Item back into
  `WebPOS_EES` with the same fields, plus billing/product/support type.
- **Create POT** - fills `POT_Template.docx` and embeds the raw ticket comment
  (images included) at `{{TICKETDATA}}`.
- **Ayala Checker** - reconciles a folder of hourly sales CSVs against the EOD
  CSV and writes an Excel workbook.
- **Missing Sales Generator** - produces the diagnostic SQL / helper script.
- Tools 3-10 are registered but unimplemented (they pop "not yet implemented").

Everything lives in a single `report.py`; there is no package layout.

## Requirements

- Windows (the app uses `winreg`, `ctypes` and `os.startfile`)
- Python 3.11 or newer for running from source
- A Firebase account on the `autoreport-db4a7` project, and
  `serviceAccountKey.json` beside the source or the exe

## Setup

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Copy `serviceAccountKey.json` (Firebase Console -> Project settings ->
Service accounts -> Generate new private key) into the repo root. The app
reads it by relative path, so the working directory matters: run it from the
folder that contains it.

```bat
python report.py
```

## Files the app expects next to it

| File | Purpose |
| --- | --- |
| `serviceAccountKey.json` | Firebase Admin credential - **never commit this** |
| `Report Template.docx` | Status report template |
| `POT_Template.docx` | Proof of Test template |
| `asset/death.ico` | Window / dialog icon |
| `asset/app.ico` | Installer icon |

Generated documents are written to a dated folder (`MM-DD-YYYY`) in the
working directory. `report_log.txt` is written alongside them.

## Templates

Placeholders are replaced in body paragraphs, table cells (nested tables
included) and text boxes, so a tag can be moved anywhere in the template.

Status report: `{{TICKETNUM}}`, `{{TICKETCONTENT}}`, `{{DATE}}`, `{{OS}}`,
`{{FS}}`, `{{CLIENT}}`, `{{RCA}}`, `{{PREVAC}}`, `{{NEXT}}`, `{{CURSTAT}}`,
`{{DISPLAYN}}`, `{{REPORT_CONTENT}}` (HTML body with images).

POT: `{{POT_TITLE}}`, `{{COMPANY}}`, `{{COMPANY_ID}}`, `{{COMPANY_NAME}}`,
`{{TICKET_NO}}`, `{{DATE}}`, `{{MODULE}}`, `{{URL}}`, `{{ISSUE_LIST}}`,
`{{DISPLAYN}}`, `{{TICKETDATA}}` (HTML body with images).

## Adding a tool

Append a dict to `TOOL_REGISTRY` near the top of `report.py`, then implement
the matching `run_<name>` method on `LoginWindow`:

```python
{
    "id": "my_tool",
    "name": "My Tool",
    "description": "What it does",
    "icon": "✨",
    "handler": "run_my_tool",
    "category": "Utilities",
},
```

The Tools page reads the registry at open time; no UI edits are needed.

## Building a release

The app updates itself from GitHub Releases on
`shinispades/AutoReport`. A release only reaches users when all three happen:

1. Bump `APP_VERSION` in `report.py`. The updater compares dotted numbers, so
   the tag must be strictly greater than the running version.
2. Build and package, e.g.
   `pyinstaller --noconfirm --windowed --icon asset/app.ico report.py`, then
   zip the `dist` folder. **Zip is preferred** - it is extracted in place and
   the app restarts itself. `.exe` assets are launched as an installer.
3. Publish a release whose tag matches (`v3.0.8.6` for `APP_VERSION = "3.0.8.6"`)
   with the zip attached.

Users are offered the update ~2 seconds after the window appears, and Settings
has a **Check for Updates** button. Locked files (the running exe) are staged as
`.new` and swapped by a generated batch script after the process exits, so tell
users to restart normally rather than killing the app mid-update.

Upload the zip as a release asset (not a repo file) so the API can report its
`size` and `digest`; both are verified before anything is unpacked.

### How an update protects itself

- **Download** - three attempts, written to a `.part` file, then checked against
the asset's advertised size and SHA-256 before it is used. A truncated or
altered download is thrown away, not installed.
- **Before replacing anything** - the payload must contain the app (the exe, or
`report.py` when running from source), and every file it will overwrite is
copied to `_prev/` first. If the backup fails, nothing is changed.
- **If a copy fails** - the backup is restored immediately and the dialog says
what went wrong. The install is never left half old and half new.
- **First start after an update** - writes `_update_started.flag`. Once a window
is on screen the markers are cleared. If the next start finds that flag with no
window ever recorded, the new build is assumed broken: `_prev/` is restored and
the app relaunches into the previous version, once (no rollback loop).
- **Failed swap** - if the batch script cannot replace a locked file after 20
tries it records the target in `_update_failed.txt` and still relaunches the
app. The next start logs it; the staged `.new` file is kept, not deleted.

`_prev/` holds one generation back and is refreshed on each update, so expect
it to use roughly the size of the app directory.

## Security notes

- **`serviceAccountKey.json` is a master key for the whole Firestore database.**
  It has to sit next to the exe for the app to work, which means anyone who
  runs the app can read every user's record - including their TFS personal
  access token, which is stored in plaintext. Keep the file out of git (see
  `.gitignore`) and treat any leaked copy as a full database compromise: rotate
  the key in the Firebase console.
- PATs should be scoped to the minimum work-item permissions the tool needs,
  and rotated on the schedule shown on the Settings page.
- `report_log.txt` records emails, ticket IDs and ticket titles in plaintext.
  It is useful for support but should not be attached to external tickets as-is.
- `pyrebase` is unmaintained and pulls in deprecated `gcloud` / `oauth2client`
  packages. If sign-in starts failing after a Python upgrade, move to the
  Firebase REST sign-in endpoint instead of pinning an older interpreter.
