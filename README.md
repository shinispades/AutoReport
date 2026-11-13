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
- **Scheduled Comment** - writes a TFS comment (text, links, images), stores it
  in Firestore, and `scheduled_comment_worker.py` on an always-on host posts it
  to the ticket at the time you picked. See
  [Scheduled comments](#scheduled-comments-worker-on-a-host).
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
2. Run `build.bat`. It creates `.venv`, installs the pinned requirements and
   PyInstaller, freezes `report.py` into `dist\Report.exe` (direct
   `pyinstaller --onefile` call, no spec to maintain) and copies the templates
   and `asset\` next to the exe. Then zip the `dist` folder. **Zip is
   preferred** - it is extracted in place and the app restarts itself. `.exe`
   assets are launched as an installer.
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

## Scheduled comments (worker on a host)

The **Scheduled Comment** tool writes a comment now and has it posted to the
ticket later. Two pieces are involved:

| Piece | Runs on | Needs |
| --- | --- | --- |
| `Scheduled Comment` tool inside `report.py` | your PC | nothing extra |
| `scheduled_comment_worker.py` | an always-on host | Python 3.11+, `requests` only |

The worker deliberately does not import `report.py`: that module needs PyQt6,
`winreg` and DPAPI, none of which exist in a Linux container.

### How a comment gets posted

1. You pick a work item, a local date/time and write the comment.
2. Any images are uploaded straight away as **work item attachments**, and the
   comment references their URLs. The host cannot read your disk, and an
   attachment URL keeps working however late the comment is posted.

### Writing the comment

The box behaves like Word rather than a plain text field:

| Type | Result |
| --- | --- |
| `*` or `-` + space at the start of a line | starts a bulleted list |
| `1.` (or `1)`) + space | starts a numbered list |
| Ctrl+B / Ctrl+I / Ctrl+U | bold / italic / underline for the selection or what is typed next |
| Ctrl+V with a picture on the clipboard | the picture is inserted and attached to the comment |

Enter on an empty list item ends the list, as in Word. Bare URLs are still
turned into links automatically.

The editor is rich text, so what TFS receives is real `<b>`, `<i>`, `<u>`,
`<ul>`/`<ol>` markup rather than Markdown. A picture pasted into the box is
written to a temp PNG under the system temp folder and uploaded as an
attachment when the schedule is saved.

A comment copied out of TFS is different: its pictures arrive as
`<img src="https://tfs/.../_apis/wit/attachments/...">`, URLs rather than
bytes, and Qt cannot load them (they need your token), so it draws an empty
box. The editor therefore downloads each one in the background with your PAT -
sent **only** when the picture's host matches `TFS_ORG_URL` - and swaps it into
the box, so it is visible while you write and is uploaded with the comment.
If a download fails, the URL is left in place and TFS still renders the
picture from the original attachment.

Only a fixed tag set (paragraphs, lists, `b`/`strong`, `i`/`em`, `u`, `s`,
`code`, links, `img` with an http(s) source) survives into the stored HTML, so
pasting from a browser or Word cannot inject scripts, styles or `<iframe>`s.

3. The comment lands in the `scheduledComments` Firestore collection with
   `status: pending` and a UTC `due_at`.
4. The worker polls every `POLL_SECONDS` (default 30), claims each due document
   with a compare-and-set write (so two instances can never post twice), posts
   it to `POST <project>/_apis/wit/workItems/{id}/comments`, then marks it
   `sent` or `failed` after `MAX_ATTEMPTS` tries.
5. The comment is authored with the **scheduling user's PAT**, which the app
   stores in the schedule document itself (`scheduledComments/{id}.pat`), so
   it appears under their name. Schedules saved by older builds have no `pat`
   field, and the worker then reads `users/{uid}` instead. The worker's own
   `TFS_PAT` is only used when `TFS_ALLOW_SHARED_PAT=1` is set, because it
   credits whoever owns that token.

### Which project a comment goes to

The **Project** picker in the dialog offers every entry in `TFS_TARGETS`
(`report.py`):

| Label | Project URL |
| --- | --- |
| `WebPOS` | `https://tfs.alliancewebpos.com/tfs/WebPOSCollection/WebPOS` |
| `WPS` | `https://tfs.alliancewebpos.com/tfs/WPSCollection/WPS` |

The chosen URL is stored on the schedule as `tfs_org_url` (with the label in
`tfs_target`), and the worker posts to that project. Images are uploaded to the
same project at compose time. Schedules saved by older builds carry no URL and
go to the worker's `TFS_ORG_URL`.

Add a project in two places: the `TFS_TARGETS` list in `report.py` and the
`TFS_TARGETS` dict in `scheduled_comment_worker.py`. The worker **refuses** a
schedule whose URL is not in its own list - a schedule carries a person's PAT,
so an unchecked URL would be a way to send that token to another server.

### Is the worker online?

A bot host usually exposes no public HTTP port, so the dialog cannot ping the
worker directly. The worker instead writes a heartbeat document on every poll:

```
workerStatus/<WORKER_ID>  ->  last_seen, started_at, host, version,
                              projects, pending, poll_seconds
```

The **Scheduled Comment** dialog reads the newest heartbeat on open, every 30
seconds, and when **Check** is pressed, and shows one of:

- **Worker online** - heartbeat within `WORKER_ONLINE_SECONDS` (120s), with the
  projects it serves and how long ago it was last seen.
- **Worker offline** - a heartbeat exists but is older than that. Scheduling is
  still allowed after a confirmation, because the comment will post when a
  worker comes back; the point is that nothing posts it right now.
- **Worker status unknown** - no heartbeat has been written yet, or the
  `workerStatus` read was denied. Check the worker log and the rules below.

Each worker needs its own `WORKER_ID` (default: the hostname) so two workers do
not overwrite each other's heartbeat. A missing heartbeat never stops comments
from being posted - it only affects what the app displays.

If the label reads "could not check ... Missing or insufficient permissions",
the published rules have no `workerStatus` block yet (or it is nested inside
another `match`). That is a rules problem, not a worker problem: the worker may
be running fine and posting comments while the app is still refused. Add the
block from step 5 below, press **Publish**, and press **Check** again. The
worker logs `Signed in as ... (uid ...)` on startup - that uid is what
`THE_WORKER_UID` must be replaced with, and the same log prints a heartbeat
warning if the worker's *write* is denied as well.

### Deploying the worker

A Discord bot host such as bot-hosting.net is just a container running one
long-lived process - no Discord token is involved. Any small Linux host works
the same way.

1. Create a **dedicated Firebase account** for the worker (e.g.
   `scheduler@yourcompany.com`). It has to read every user's pending
   schedules, so it must not be a person's account.
2. Upload `scheduled_comment_worker.py` and `requirements-worker.txt`, then
   `pip install -r requirements-worker.txt`.
3. Set the environment variables in the host's startup/env settings:

   | Variable | Meaning |
   | --- | --- |
   | `FIREBASE_EMAIL` | the worker account's email |
   | `FIREBASE_PASSWORD` | its password |
   | `TFS_PAT` | fallback PAT, used only when the schedule's author has none |
   | `TFS_ORG_URL` | optional fallback project URL for schedules saved without one, defaults to WebPOS |
   | `WORKER_ID` | optional name this worker's heartbeat is written under, default: the hostname |
   | `POLL_SECONDS` | optional, default `30` |
   | `MAX_ATTEMPTS` | optional, default `3` |
   | `TFS_STRICT_AUTHOR` | optional, `1` refuses to post a comment that would appear under an account other than the person who scheduled it |
   | `TFS_ALLOW_SHARED_PAT` | optional, `1` allows `TFS_PAT` to be used for schedules that carry no token of their own (default off) |

   The panel only passes the variables it defines - a `.env` file is not read
   by the container. The worker therefore loads a `.env` placed next to itself
   (or in the working directory) on startup, so either route works. Real
   environment variables win over the file.

4. Start it with `python scheduled_comment_worker.py`. `--once` posts whatever
   is due and exits, which is the quickest way to check the setup.
5. Add rules for the collection in Firestore:

   ```
   match /scheduledComments/{docId} {
     // Owner or worker only: each schedule carries the author's PAT.
     allow read: if request.auth != null
                 && (resource.data.owner_uid == request.auth.uid
                     || request.auth.uid == "THE_WORKER_UID");
     allow create: if request.auth != null
                   && request.resource.data.owner_uid == request.auth.uid;
     allow update: if request.auth != null
                   && (resource.data.owner_uid == request.auth.uid
                       || request.auth.uid == "THE_WORKER_UID"
                       || resource.data.status in ["pending", "sending"]);
     allow delete: if request.auth != null
                   && resource.data.owner_uid == request.auth.uid;
   }

   match /workerStatus/{docId} {
     // The app shows the worker as online/offline from these documents, so any
     // signed-in user reads them; only worker accounts may write.
     allow read: if request.auth != null;
     allow write: if request.auth != null
                  && request.auth.uid == "THE_WORKER_UID";
   }

   match /users/{uid} {
     // Each person reads and writes their own profile, as before.
     allow read, write: if request.auth != null && request.auth.uid == uid;
     // Optional. Only needed for schedules saved before COPY_PAT_INTO_SCHEDULE,
     // which carry no token of their own: without it those comments cannot be
     // posted as their author.
     allow read: if request.auth != null
                 && request.auth.uid == "THE_WORKER_UID";
   }
   ```

   The worker writes to documents it does not own, which is what the
   `request.auth.uid == "THE_WORKER_UID"` clause is for - replace it with the
   worker account's uid, or drop the clause and rely on the status check
   (broader: any signed-in user could then claim a pending schedule).
   The status list must include `"sending"`, not only `"pending"`: the worker
   first claims a document as `sending`, and without that the later `sent`
   write is denied - the comment posts but the schedule is retried later, and
   the ticket gets a duplicate.

   The `match /users/{uid}` block above is the existing one with a single line
   added; keep your current version and only add the second `allow read` if you
   still have schedules saved by an older build. The `scheduledComments` block
   matters either way: restricting `read` to the owner and the worker is what
   keeps each copied PAT private now that schedules carry one.

   The new `match` block must sit at the **top level** of the rules document,
   next to `match /users/{uid}`. Paste it *inside* that block and the real path
   becomes `users/{uid}/scheduledComments/...`, so every save is denied with
   "The database refused the request". Press **Publish** after editing -
   unsaved rules in the console are not in effect.

No composite Firestore index is needed: the app queries on `owner_uid`, the
worker on `status`, and the time filtering/ordering happens in Python.

### The comment appears as the wrong person

A TFS comment is authored by **whoever owns the PAT** used to post it - nothing
in the comment payload can change that. The worker picks a token in this order:

1. `pat` stored inside the schedule - written by the app when `COPY_PAT_INTO_SCHEDULE`
   is on, which is the case that always gives you the right author,
2. `users/{owner_uid}.pat` - schedules saved by older builds, and the copy
   that also lets the account sign in on another PC,
3. the worker's `TFS_PAT`, whose owner then gets credited - only with
   `TFS_ALLOW_SHARED_PAT=1`. Without it the schedule is left `pending` rather
   than posted under the wrong name.

So a comment under the `TFS_PAT` owner means step 1 and step 2 both came up
empty. The worker log says which one:

| Log line | Cause | Fix |
| --- | --- | --- |
| `Firestore refused users/... (403)` | the `users/{uid}` rule does not let the worker account read other people's documents (or `THE_WORKER_UID` in the rules was never replaced with the real uid) | add the second `allow read` from the rules block above, or re-save the schedule with `COPY_PAT_INTO_SCHEDULE` on |
| `NOT posting - no PAT stored for ...` | the schedule has no `pat` field (saved by a build older than `COPY_PAT_INTO_SCHEDULE`) and the worker could read nothing from `users/{uid}` | re-save the schedule in the current app |
| `WARNING ... owns the token used, but ... scheduled it` | the token found belongs to somebody else (a shared personal token handed round the team, or the wrong token copied into `users/{uid}`) | give each person their own TFS PAT |
| `WARNING TFS recorded ... as the author, not the scheduler` | the comment was posted and TFS recorded a different account - this is the definitive check | the PAT in `scheduledComments/{id}.pat` belongs to the wrong person |

Run the worker with `--check` to see this without posting anything:

```bash
python scheduled_comment_worker.py --check
```

It signs in, lists every pending/sending schedule with its author, the token
that would be used and the TFS account that token belongs to, and posts
nothing. It is safe on the live host while the worker is running.

The account lookup uses `_apis/connectionData`, which older on-premises TFS
installations answer only without an `api-version` or with an old one, so the
worker tries no version first and then a descending list. If the server
answers none of them (`Could not resolve a token's TFS account (...)`), the
worker still posts, and reports the author from the `createdBy` field of the
reply to its own comment POST - `posted as <account>`. That reply is the
recorded truth, so a wrong author is reported even when the lookup fails.

Every real post also logs `comment will appear as <account>`, so a wrong author
is visible even when `--check` is not run.

By default the worker will not post a comment as the wrong person: a schedule
whose author's token it cannot find stays `pending` (with the reason in
`last_error`) and is retried every poll, so it goes out by itself once the
token is there. `TFS_STRICT_AUTHOR=1` adds the same block for the case where a
token *was* found but belongs to somebody else.

Note that the worker's own `FIREBASE_EMAIL` is irrelevant to the author: the
worker can sign in as any account (even a person's) and the comment still
carries the author's PAT. Use a dedicated worker account anyway, so revoking
it does not disturb a person's sign-in.

### `TFS rejected the PAT (401)`

TFS never saw a token it accepts, so nothing was posted. The schedule holds a
**copy** of the author's PAT taken when it was saved (`scheduledComments/{id}.pat`),
which means a token that is valid in the app today can still be dead in a
schedule saved last month: PATs expire, and regenerating one invalidates the
old value everywhere it was copied.

Check the token against the project before changing anything else:

```bash
# the collection root - a 401 here means the token itself is bad for this server
curl -s -o /dev/null -w "%{http_code}\n" \
     -u :<PAT> https://tfs.alliancewebpos.com/tfs/WPSCollection/_apis/connectionData

# the project - a 401 here with 200 above means no access to that project
curl -s -o /dev/null -w "%{http_code}\n" \
     -u :<PAT> "https://tfs.alliancewebpos.com/tfs/WPSCollection/WPS/_apis/wit/workitems/1?api-version=7.0"
```

| Result | Meaning | Fix |
| --- | --- | --- |
| 401 on both | the PAT is expired, revoked, scoped to a different server, **or its account has no access to that collection** - on-premises TFS answers 401, not 403, for an identity that is not a member of the collection | check the token works elsewhere on the same server (a WebPOS schedule posting fine proves the token is alive), then add the account to the collection/project team or issue a PAT from an account that is in it |
| 200 on the collection, 401 on the project | the token is valid but that account has no work-item access to `WPS` | add the account to the project team, and check the PAT's scopes include Work Items (read & write) |
| 404 on the collection root | `WPSCollection` is not the collection's name/path on this server | fix the URL in `TFS_TARGETS` in both `report.py` and the worker |

The retry log now names the URL and keeps every candidate's status, e.g.
`the token was rejected (401)` for both the collection and the project, or a
404 for the collection followed by a 401 for the project. After `MAX_ATTEMPTS`
the schedule is marked `failed` with the same text in `last_error`, which is
what the app lists - nothing is retried behind your back.

A 401 at the **collection root** (`/tfs/<Collection>/_apis/connectionData`)
means the server never accepted the token, so the project URL and the
`TFS_TARGETS` entry are not implicated. Compare two schedules by the token
fingerprint the worker prints:

```
... (WI 50101 WebPOS): posting with the PAT stored with the schedule [token sha 3f9c1a2b (52 chars)]
... (WI 5510 WPS): posting with the PAT stored with the schedule [token sha 7d20e4f1 (52 chars)]
```

Same fingerprint and WPS still 401s - the account lacks access to that
collection. Different fingerprints - the failing schedule is holding a
different (older or other person's) token; re-save it in the app.

### Things to know

- The comment is authored by the **person who scheduled it**: the app copies
  their PAT into the schedule document, and the worker posts with it. That
  removes any dependency on the worker being allowed to read `users/{uid}` - a
  permission that is easy to get wrong and used to fail silently.
- A comment the worker cannot attribute to its author is **not** posted, so a
  ticket can no longer quietly receive a comment under the wrong name. The
  schedule waits in `pending` with the reason in `last_error`.
- **Where the PAT lives.** The app keeps it DPAPI-encrypted under
  `%LOCALAPPDATA%\AutoReport\credentials\` and, because `PAT_SYNC_TO_FIRESTORE`
  is on in `report.py`, also mirrors it into the user's own `users/{uid}`
  document. That second copy is what lets the same account sign in on another
  PC without re-entering the token, but it is plain text in Firestore: its only
  protection is the owner-only `users/{uid}` rule and whatever guards the
  project. Set `PAT_SYNC_TO_FIRESTORE = False` to keep the PAT on each machine
  only, at the cost of one prompt per new PC.
- Free hosts restart containers. Claims are atomic and attempts are counted, and
  a schedule left on `sending` by a restart is retried after 10 minutes, so a
  restart cannot silently swallow a comment - but run **one** worker instance.
- Cancelling a schedule after its images were uploaded leaves the attachments
  on the work item. An unused attachment is harmless; TFS has no cheap "delete
  orphan attachment" call.
- If the host is asleep at the scheduled moment, the comment goes out on the
  first poll after it wakes. The schedule is "no earlier than", not exact.
- The worker's PAT needs work item read/write on the project - a comment is a
  work item write.

## Security notes

- **`serviceAccountKey.json` is a master key for the whole Firestore database.**
  It has to sit next to the exe for the app to work, which means anyone who
  runs the app can read every user's record - including their TFS personal
  access token, which is stored in plaintext. Keep the file out of git (see
  `.gitignore`) and treat any leaked copy as a full database compromise: rotate
  the key in the Firebase console.
- **Every** scheduled comment now stores the author's PAT in its document
  (`COPY_PAT_INTO_SCHEDULE`), because the hosted worker has no other way to
  post it as that person. The token is plain text in Firestore: its only
  protection is the `read` rule on `scheduledComments`, which is why the rules
  above limit reads to the owner and the worker. Anyone who can read the whole
  collection is holding everybody's token.
- Scoped, per-person PATs matter more now than they did: a token kept in
  Firestore should not be able to do anything except comment on work items.
- PATs should be scoped to the minimum work-item permissions the tool needs,
  and rotated on the schedule shown on the Settings page.
- `report_log.txt` records emails, ticket IDs and ticket titles in plaintext.
  It is useful for support but should not be attached to external tickets as-is.
- `pyrebase` is unmaintained and pulls in deprecated `gcloud` / `oauth2client`
  packages. If sign-in starts failing after a Python upgrade, move to the
  Firebase REST sign-in endpoint instead of pinning an older interpreter.
