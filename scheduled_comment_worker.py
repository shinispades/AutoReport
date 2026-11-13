"""AutoReport - scheduled TFS comment worker.

Runs on a small always-on host (bot-hosting.net, a VPS, or a spare Windows PC).
It polls Firestore for comments the desktop app scheduled and posts each one
to TFS when its time arrives.

This is a separate file on purpose. report.py imports PyQt6, winreg and DPAPI,
so it cannot run on a Linux container. Nothing here needs anything but
`requests`.

Configuration (environment variables, set them in the host's panel):

    FIREBASE_EMAIL      Firebase account the worker signs in as (make a
                        dedicated one - see the README)
    FIREBASE_PASSWORD   its password
    TFS_PAT             fallback PAT, used only when the person who scheduled
                        a comment has no token in users/{uid}
    TFS_ORG_URL         fallback project URL for schedules that carry none
                        (older builds); default: WebPOS
    WORKER_ID           name this worker reports under (default: hostname)
    POLL_SECONDS        how often to look for due comments (default 30)
    MAX_ATTEMPTS        give up on a comment after this many tries (default 3)
    TFS_STRICT_AUTHOR   set to 1 to refuse to post a comment whose author
                        would not be the person who scheduled it
    TFS_ALLOW_SHARED_PAT
                        set to 1 to allow TFS_PAT to be used for schedules
                        that carry no token of their own. Off by default: that
                        comment would be credited to the TFS_PAT owner

Run once to test, then leave it running:

    python scheduled_comment_worker.py --check   # who would post what (no writes)
    python scheduled_comment_worker.py --once
    python scheduled_comment_worker.py

A schedule may target any of the projects in TFS_TARGETS below. The app writes
the project URL into the schedule (`tfs_org_url`) and this worker refuses any
URL outside that list - the schedule carries a person's PAT, so an unvetted
URL would be a way to send that token to another server.

Each poll also writes a heartbeat document to `workerStatus/{WORKER_ID}`, which
is how the desktop app tells the user whether a worker is actually running
before it lets them schedule a comment.

A comment is authored by whoever owns the PAT used to post it. When the
comment shows the wrong name, run --check on the host: it reports, for every
pending schedule, which PAT would be used and which TFS account that PAT
belongs to.
"""

import os
import hashlib
import re
import socket
import sys
import time
import traceback
from datetime import datetime, timezone

import requests
from requests.auth import HTTPBasicAuth

# Public Firebase web config - the apiKey is meant to be shipped in a client.
PROJECT_ID = "autoreport-db4a7"
API_KEY = "AIzaSyD-uAl1Ce-2NRoAVVYExH5-MSJ6K4y-LEo"
IDENTITY_API = "https://identitytoolkit.googleapis.com/v1/accounts"
SECURE_TOKEN_API = "https://securetoken.googleapis.com/v1/token"
DOCUMENTS_API = (
    "https://firestore.googleapis.com/v1/projects/"
    f"{PROJECT_ID}/databases/(default)/documents"
)

COLLECTION = "scheduledComments"
USERS_COLLECTION = "users"
DEFAULT_ORG_URL = "https://tfs.alliancewebpos.com/tfs/WebPOSCollection/WebPOS"

# Projects a schedule may name. The app shows these labels and writes the
# matching URL into the schedule; anything else is refused below instead of
# being sent the author's token.
TFS_TARGETS = {
    "WebPOS": "https://tfs.alliancewebpos.com/tfs/WebPOSCollection/WebPOS",
    "WPS": "https://tfs.alliancewebpos.com/tfs/WPSCollection/WPS",
}

# Where the heartbeat goes: one document per worker, read by the desktop app.
STATUS_COLLECTION = "workerStatus"
WORKER_VERSION = "1.0"

COMMENT_API_VERSION = "5.1-preview.3"

# connectionData takes no api-version at all on most servers, and refuses one
# newer than the server supports (an on-premises TFS answers 400 for 7.1 while
# Azure DevOps answers 404 for an unknown path). Try none first, then older
# versions, so this works against either kind of server.
CONNECTION_API_VERSIONS = (None, "7.1", "6.0", "5.1", "4.1", "3.2", "2.2", "1.0")
HTTP_TIMEOUT = 30

# How long a claimed ("sending") schedule may sit before it is retried. See
# is_stale_claim().
STALE_CLAIM_SECONDS = 600

# How long the PAT read from users/{uid} is reused, so a batch of due schedules
# from the same author costs one read instead of one per comment.
PAT_CACHE_SECONDS = 300

# PAT -> TFS account is a stable fact for the life of a token; cache it so a
# batch of due schedules costs one extra request, not one per comment.
IDENTITY_CACHE_SECONDS = 3600

# A failed lookup is cached only briefly: a TFS restart or a wrong version
# should not leave every author unverified for an hour.
IDENTITY_FAIL_CACHE_SECONDS = 300

_pat_cache = {}          # uid -> (fetched_at, pat)
_identity_cache = {}     # pat -> (fetched_at, identity dict or "")

# Schedules already reported as unpostable, so a document that waits for its
# author's token to be fixed does not print the same line every poll.
_refused = set()


def log(message):
    """Console log the host actually shows.

    flush matters: a Pterodactyl console buffers stdout, so without it the
    panel looks frozen for minutes at a time.
    """
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")
    print(f"{stamp} {message}", flush=True)


def load_env_file(path):
    """Read KEY=VALUE lines from a .env file into os.environ.

    Bot hosts (bot-hosting.net, Pterodactyl) start the container with the
    panel's variables and nothing else: a .env sitting in the folder is never
    loaded for you, so the worker reads it here. Values already present in the
    real environment win, so the panel stays the authority.

    Handles a UTF-8 BOM (Notepad's default when saving), CRLF line endings,
    `export ` prefixes, comments and quoted values. Returns the names of the
    keys it set.
    """
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            lines = f.readlines()
    except FileNotFoundError:
        return []
    except OSError as e:
        log(f"Could not read {path}: {e}")
        return []

    loaded = []
    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export "):].lstrip()
        key, sep, value = line.partition("=")
        key = key.strip()
        if not sep or not key:
            continue
        value = value.strip()
        # A password may legally contain '#' or quotes, so only strip a
        # matching pair around the whole value; no inline-comment handling.
        if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
            value = value[1:-1]
        if key not in os.environ:
            os.environ[key] = value
            loaded.append(key)
    return loaded


def utc_now():
    return datetime.now(timezone.utc)


def now_iso():
    """Timestamp format the app writes and this worker compares against.

    Fixed-width UTC, so plain string comparison is chronological - which is
    what makes the Firestore `due_at` range filter unnecessary (see
    fetch_pending).
    """
    return utc_now().strftime("%Y-%m-%dT%H:%M:%SZ")


# Reported in the heartbeat, so "last seen" can be read next to "running since".
STARTED_AT = now_iso()


# ================= FIREBASE AUTH =================

# Plain-English meaning of the Identity Toolkit error codes, so the host's
# console says what to fix instead of only "invalid credentials".
AUTH_HINTS = {
    "INVALID_LOGIN_CREDENTIALS": (
        "the email/password pair was rejected. Firebase answers this way both "
        "for a wrong password and for an account that does not exist, so check: "
        "the user exists under Authentication > Users in the Firebase console "
        "for this project, its password is current, and Email/Password is "
        "enabled under Authentication > Sign-in method"
    ),
    "EMAIL_NOT_FOUND": (
        "no user with that email exists in this Firebase project - create it in "
        "Authentication > Users"
    ),
    "INVALID_PASSWORD": "the password is wrong - reset it in the console",
    "USER_DISABLED": "the account is disabled in the Firebase console",
    "OPERATION_NOT_ALLOWED": (
        "Email/Password sign-in is switched off for this project - enable it "
        "under Authentication > Sign-in method"
    ),
    "TOO_MANY_ATTEMPTS_TRY_LATER": (
        "Firebase rate-limited the failed logins - wait a few minutes"
    ),
}


def sign_in(email, password):
    """Email/password sign-in. Returns (id_token, refresh_token)."""
    log(f"Signing in as {email}")
    resp = requests.post(
        f"{IDENTITY_API}:signInWithPassword?key={API_KEY}",
        json={"email": email, "password": password, "returnSecureToken": True},
        timeout=HTTP_TIMEOUT,
    )
    if resp.status_code >= 400:
        detail = ""
        try:
            detail = (resp.json().get("error") or {}).get("message", "")
        except ValueError:
            pass
        hint = AUTH_HINTS.get(detail, "see the Firebase console for this project")
        raise RuntimeError(
            f"Firebase sign-in failed ({resp.status_code}): {detail} - {hint}"
        )
    data = resp.json()
    # The uid is what the scheduledComments/workerStatus rules need for
    # THE_WORKER_UID; printing it saves hunting through the Auth console.
    log(f"Signed in as {email} (uid {data.get('localId', 'unknown')})")
    return data.get("idToken", ""), data.get("refreshToken", "")


def refresh_session(refresh_token):
    """ID tokens live for an hour; renew with the refresh token.

    Note the snake_case keys - the Secure Token API differs from Identity
    Toolkit's camelCase.
    """
    resp = requests.post(
        f"{SECURE_TOKEN_API}?key={API_KEY}",
        json={"grant_type": "refresh_token", "refresh_token": refresh_token},
        timeout=HTTP_TIMEOUT,
    )
    resp.raise_for_status()
    data = resp.json()
    return data.get("id_token", ""), data.get("refresh_token", refresh_token)


def auth_headers(id_token):
    return {
        "Authorization": f"Bearer {id_token}",
        "Content-Type": "application/json",
    }


# ================= FIRESTORE (REST) =================

def decode_value(value):
    """One Firestore REST Value -> Python value."""
    if not isinstance(value, dict):
        return ""
    for key, cast in (
        ("stringValue", str),
        ("booleanValue", bool),
        ("timestampValue", str),
    ):
        if key in value:
            return cast(value[key])
    if "integerValue" in value:
        try:
            return int(value["integerValue"])
        except (TypeError, ValueError):
            return 0
    if "doubleValue" in value:
        return value["doubleValue"]
    return ""


def decode_document(doc):
    """A Firestore document -> (id, flat dict, updateTime)."""
    fields = doc.get("fields") or {}
    data = {name: decode_value(value) for name, value in fields.items()}
    data["_id"] = (doc.get("name") or "").rsplit("/", 1)[-1]
    return data, doc.get("updateTime", "")


def fetch_pending(id_token):
    """Documents that are still pending, plus abandoned "sending" claims.

    Deliberately one field, no range and no order-by: combining `status` with
    a `due_at` comparison would need a composite index that has to be built in
    the console before the worker works at all. The small set comes back and
    the time comparison happens here instead.
    """
    body = {
        "structuredQuery": {
            "from": [{"collectionId": COLLECTION}],
            "where": {
                "fieldFilter": {
                    "field": {"fieldPath": "status"},
                    "op": "IN",
                    "value": {
                        "arrayValue": {
                            "values": [
                                {"stringValue": "pending"},
                                {"stringValue": "sending"},
                            ]
                        }
                    },
                }
            },
            "limit": 100,
        }
    }
    resp = requests.post(
        f"{DOCUMENTS_API}:runQuery",
        headers=auth_headers(id_token),
        json=body,
        timeout=HTTP_TIMEOUT,
    )
    if resp.status_code == 401:
        raise PermissionError("Firestore rejected the worker's token")
    resp.raise_for_status()

    documents = []
    for row in resp.json() or []:
        doc = row.get("document")
        if not doc:
            continue          # runQuery returns a bare readTime when a row is empty
        data, update_time = decode_document(doc)
        data["_update_time"] = update_time
        documents.append(data)
    return documents


def fetch_user_pat(id_token, uid):
    """Read users/{uid}.pat - the copy the desktop app keeps so the same
    account works on any PC.

    Returns "" when the document or the field is missing, or when the rules do
    not let the worker read another user's document. A missing PAT is not fatal
    here: the caller falls back to the shared TFS_PAT.
    """
    resp = requests.get(
        f"{DOCUMENTS_API}/{USERS_COLLECTION}/{uid}",
        headers=auth_headers(id_token),
        timeout=HTTP_TIMEOUT,
    )
    if resp.status_code == 404:
        return ""
    if resp.status_code in (401, 403):
        log(f"Firestore refused users/{uid} ({resp.status_code}). Working out "
            "who to post as needs read access to every user document - see the "
            "users/{uid} rule in the README. Falling back to TFS_PAT.")
        return ""
    if resp.status_code >= 400:
        log(f"Could not read users/{uid}: HTTP {resp.status_code}")
        return ""
    data, _ = decode_document(resp.json())
    return (data.get("pat") or "").strip()


def normalise_org_url(url):
    """Compare project URLs without caring about case or a trailing slash."""
    return (url or "").strip().rstrip("/").lower()


def target_url(doc):
    """The project URL this schedule must be posted to, or "" when refused.

    Schedules saved by older builds have no `tfs_org_url`, so those fall back
    to the worker's default. A URL that is not one of TFS_TARGETS is refused:
    the document carries the author's PAT, and posting it to a host the worker
    was not configured for would hand that token to that host.
    """
    wanted = normalise_org_url(doc.get("tfs_org_url"))
    if not wanted:
        return DEFAULT_ORG_URL
    for known in TFS_TARGETS.values():
        if normalise_org_url(known) == wanted:
            return known
    return ""


def target_label(doc, url):
    """Human name for a project URL - the schedule's own label when it has one."""
    for label, known in TFS_TARGETS.items():
        if normalise_org_url(known) == normalise_org_url(url):
            return label
    return (doc.get("tfs_target") or url.rsplit("/", 1)[-1]).strip()


def pat_identity(pat, org_url):
    """The TFS account a PAT belongs to - which is the author of the comment.

    GET .../_apis/connectionData answers with the account behind the token, so
    this can say "this comment will appear as X" before one is posted - which
    is what --check needs. After a real post, comment_author() reads the same
    fact from TFS's reply and is the authoritative check.

    Returns {"unique_name", "display_name"}, or "" when TFS cannot be asked.
    """
    if not pat:
        return ""
    now = time.time()
    cached = _identity_cache.get(pat)
    if cached:
        ttl = IDENTITY_CACHE_SECONDS if cached[1] else IDENTITY_FAIL_CACHE_SECONDS
        if now - cached[0] < ttl:
            return cached[1]

    identity = ""
    # Every candidate is recorded, not just the last one: a 404 on the
    # collection root followed by a 401 on the project URL is a different
    # problem from two 401s, and only the list tells them apart.
    problems = []
    for base in connection_candidates(org_url):
        for version in CONNECTION_API_VERSIONS:
            params = {} if version is None else {"api-version": version}
            try:
                resp = requests.get(
                    f"{base}/_apis/connectionData",
                    params=params,
                    auth=HTTPBasicAuth("", pat),
                    timeout=HTTP_TIMEOUT,
                )
            except requests.RequestException as e:
                problems.append(f"{base}: {e}")
                break                     # network is down, versions will not help
            if resp.status_code == 401:
                problems.append(f"{base}: the token was rejected (401)")
                break
            if resp.status_code >= 400:
                problems.append(f"{base}: HTTP {resp.status_code} "
                                f"(api-version {version or 'none'})")
                continue
            try:
                user = (resp.json() or {}).get("authenticatedUser") or {}
            except ValueError:
                problems.append(f"{base}: the reply was not JSON")
                continue
            unique_name = (user.get("uniqueName") or "").strip()
            if unique_name:
                identity = {
                    "unique_name": unique_name,
                    "display_name": (user.get("displayName") or "").strip(),
                }
                break
        if identity:
            break

    if not identity:
        # Not fatal: the comment is still posted, and the author is confirmed
        # from the reply TFS sends back for it (see comment_author()).
        log(f"Could not resolve a token's TFS account "
            f"({'; '.join(problems) or 'no answer'}) - posting anyway; the "
            "author is checked on the posted comment")

    _identity_cache[pat] = (now, identity)
    return identity


def connection_candidates(org_url):
    """Base URLs to ask for connectionData, collection first.

    The collection root is the address every server answers; the project URL
    under it may not exist (a renamed project) and an on-premises TFS replies
    404 there. Trying the collection first also saves a round of requests.
    """
    trimmed = org_url.rstrip("/")
    candidates = [trimmed]
    parts = trimmed.split("/")
    for cut in (1, 2):
        # Keep at least scheme://host/tfs/collection.
        if len(parts) - cut > 4:
            candidates.append("/".join(parts[:-cut]))
    return list(reversed(candidates))


def account_key(name):
    """The bare account name from any of the spellings TFS uses.

    On-premises TFS reports `SMB01\\jbarrientos` while the schedule carries
    `jbarrientos@asi-dev1.com`; a literal comparison of the two would flag
    every on-prem comment as a wrong author. Strip the domain first, then the
    address suffix, so both reduce to `jbarrientos`.
    """
    key = (name or "").strip().lower()
    key = key.rsplit("\\", 1)[-1]
    return key.split("@", 1)[0]


def author_matches(identity, owner_email):
    """True when the token owner is the person the schedule says it is.

    Compared case-insensitively, and also on the account name alone, because
    TFS reports the same person as `domain\\user`, `user@domain` or `user`
    depending on the collection.
    """
    if not identity or not owner_email:
        return True          # nothing to compare - do not block on a guess
    unique_name = identity.get("unique_name", "").strip().lower()
    owner = owner_email.strip().lower()
    if unique_name == owner:
        return True
    return account_key(unique_name) == account_key(owner)


def token_fingerprint(pat):
    """Identify a token without printing it.

    Two schedules that fail differently can be compared from the log: the same
    fingerprint means the same PAT (so a 401 is about access), a different one
    means the schedule is carrying an older or foreign token.
    """
    if not pat:
        return "no token"
    digest = hashlib.sha256(pat.strip().encode("utf-8")).hexdigest()[:8]
    return f"sha {digest} ({len(pat)} chars)"


def resolve_pat(id_token, doc, shared_pat):
    """Which PAT a schedule would be posted with, and where it came from.

    Order: the PAT copied into the document (older "post as me" schedules),
    then the scheduling account's own users/{uid}.pat, then the worker's
    shared TFS_PAT.
    """
    owner_uid = (doc.get("owner_uid") or "").strip()
    pat = (doc.get("pat") or "").strip()
    source = "the PAT stored with the schedule"
    if not pat and owner_uid:
        pat = owner_pat(id_token, owner_uid)
        source = "users/%s.pat (%s)" % (
            owner_uid, doc.get("owner_email") or "no email")
    if not pat:
        pat = shared_pat
        source = "the shared TFS_PAT"
    return owner_uid, pat, source


def worker_id():
    """Document id for this worker's heartbeat (Firestore ids allow no '/')."""
    raw = (os.environ.get("WORKER_ID") or socket.gethostname() or "worker").strip()
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "-", raw).strip("-")
    return (cleaned or "worker")[:60]


_status_warning_shown = False


def write_heartbeat(id_token, poll_seconds, pending):
    """Tell the app this worker is alive and which projects it serves.

    PATCH creates the document when it does not exist, so there is nothing to
    set up in the console. A rules problem is logged once and then ignored: a
    missing heartbeat must never stop comments from being posted.
    """
    global _status_warning_shown

    values = {
        "worker_id": worker_id(),
        "status": "online",
        "host": socket.gethostname(),
        "version": WORKER_VERSION,
        "projects": ", ".join(TFS_TARGETS),
        "poll_seconds": int(poll_seconds),
        "pending": int(pending),
        "started_at": STARTED_AT,
        "last_seen": now_iso(),
    }
    fields = {key: {"stringValue": "" if value is None else str(value)}
              for key, value in values.items()}
    fields["poll_seconds"] = {"integerValue": str(values["poll_seconds"])}
    fields["pending"] = {"integerValue": str(values["pending"])}

    doc_id = worker_id()
    try:
        resp = requests.patch(
            f"{DOCUMENTS_API}/{STATUS_COLLECTION}/{doc_id}",
            headers=auth_headers(id_token),
            params={"updateMask.fieldPaths": list(values.keys())},
            json={"fields": fields},
            timeout=HTTP_TIMEOUT,
        )
        if resp.status_code >= 400 and not _status_warning_shown:
            _status_warning_shown = True
            log(f"Could not write the worker heartbeat "
                f"({resp.status_code}) - the app will show this worker as "
                "offline until the workerStatus rules allow it. Posting is "
                "not affected.")
    except requests.RequestException as e:
        if not _status_warning_shown:
            _status_warning_shown = True
            log(f"Could not write the worker heartbeat ({e}) - the app will "
                "show this worker as offline. Posting is not affected.")


def owner_pat(id_token, uid):
    """users/{uid}.pat, cached briefly per author."""
    now = time.time()
    cached = _pat_cache.get(uid)
    if cached and now - cached[0] < PAT_CACHE_SECONDS:
        return cached[1]
    pat = fetch_user_pat(id_token, uid)
    _pat_cache[uid] = (now, pat)
    return pat


def patch_fields(id_token, doc_id, values, expect_update_time=None):
    """Update selected fields. Returns True on success.

    With expect_update_time set the write is a compare-and-set: it fails if
    the document changed since it was read. That is what stops two worker
    instances (a restart overlapping the old process) from posting the same
    comment twice.
    """
    params = {"updateMask.fieldPaths": list(values.keys())}
    if expect_update_time:
        params["currentDocument.updateTime"] = expect_update_time

    fields = {}
    for key, value in values.items():
        if isinstance(value, bool):
            fields[key] = {"booleanValue": value}
        elif isinstance(value, int):
            fields[key] = {"integerValue": str(value)}
        else:
            fields[key] = {"stringValue": "" if value is None else str(value)}

    resp = requests.patch(
        f"{DOCUMENTS_API}/{COLLECTION}/{doc_id}",
        headers=auth_headers(id_token),
        params=params,
        json={"fields": fields},
        timeout=HTTP_TIMEOUT,
    )
    if resp.status_code == 409:
        return False          # someone else got there first
    if resp.status_code == 404 and expect_update_time:
        return False          # deleted while we were looking at it
    resp.raise_for_status()
    return True


# ================= TFS =================

def comment_author(resp):
    """The account TFS recorded on the comment it just created.

    This is the authoritative answer to "who does this comment appear as", and
    it costs nothing: the POST reply already carries it. connectionData is only
    needed for --check, where there is no comment to read it from.
    """
    try:
        data = resp.json() or {}
    except ValueError:
        return ""
    created_by = data.get("createdBy") or {}
    unique_name = (created_by.get("uniqueName") or "").strip()
    if not unique_name:
        return ""
    return {
        "unique_name": unique_name,
        "display_name": (created_by.get("displayName") or "").strip(),
    }


def post_comment(pat, work_item_id, html, org_url):
    """Add a comment to a work item. Raises on failure."""
    url = (
        f"{org_url}/_apis/wit/workItems/{work_item_id}/comments"
        f"?api-version={COMMENT_API_VERSION}"
    )
    resp = requests.post(
        url, auth=HTTPBasicAuth("", pat), json={"text": html}, timeout=HTTP_TIMEOUT
    )
    if resp.status_code in (401, 403):
        # TFS answers both for a PAT problem: 401 when the token itself is
        # not accepted, 403 when it is accepted but the account is not allowed
        # to write here. An empty body is normal, so say what it means.
        raise RuntimeError(
            f"TFS rejected the PAT ({resp.status_code}) at {org_url} - it is "
            "expired, was regenerated after this comment was scheduled, or "
            "its account has no access to that project"
        )
    if resp.status_code >= 400:
        raise RuntimeError(
            f"TFS returned {resp.status_code}: {resp.text[:200]}"
        )
    return resp


# ================= SCHEDULING =================

def is_stale_claim(update_time, reference=None):
    """True when a "sending" claim is old enough to be considered abandoned.

    A host restart between the claim and the result used to leave a document
    stuck on "sending" forever: nothing queried that status, so the comment
    was never posted and never visibly failed. Anything claimed longer ago
    than STALE_CLAIM_SECONDS is retried instead.
    """
    try:
        claimed = datetime.fromisoformat(str(update_time).replace("Z", "+00:00"))
    except (TypeError, ValueError):
        return True           # unreadable: better a retry than a silent loss
    return ((reference or utc_now()) - claimed).total_seconds() >= STALE_CLAIM_SECONDS


def is_due(due_at, reference=None):
    """True when a schedule's due_at (UTC ISO) has arrived."""
    try:
        due = datetime.strptime(due_at, "%Y-%m-%dT%H:%M:%SZ").replace(
            tzinfo=timezone.utc
        )
    except (TypeError, ValueError):
        log(f"Rejected schedule with unreadable due_at: {due_at!r}")
        return False
    return due <= (reference or utc_now())


def leave_pending(id_token, doc_id, label, reason):
    """Undo our claim and keep a schedule retryable, explaining it once.

    A schedule that cannot be posted as its author is not a failure of the
    comment, it is a missing token: marking it failed would burn the schedule
    permanently (the worker cannot revive a failed document). Leaving it
    pending means it goes out by itself once the token is fixed.
    """
    patch_fields(id_token, doc_id, {"status": "pending", "last_error": reason[:400]})
    if doc_id not in _refused:
        _refused.add(doc_id)
        log(f"{label}: NOT posting - {reason}")


def run_once(id_token, shared_pat, fallback_org_url, max_attempts,
             strict_author=False, allow_shared_pat=False):
    """Post every pending comment that is due.

    Returns (posted, pending): the count the heartbeat reports, so the app can
    say how much work a live worker is carrying.
    """
    posted = 0
    pending = fetch_pending(id_token)

    due_docs = [doc for doc in pending if is_due(doc.get("due_at", ""))]
    if due_docs:
        log(f"{len(due_docs)} due of {len(pending)} pending schedule(s)")

    for doc in due_docs:
        doc_id = doc["_id"]
        work_item_id = str(doc.get("work_item_id", "")).strip()
        attempts = doc.get("attempts", 0) or 0
        owner_email = (doc.get("owner_email") or "").strip()

        # Which project this comment belongs to, decided by the schedule (the
        # app writes it). Older schedules have none and use the fallback, which
        # is this worker's own TFS_ORG_URL rather than something a document
        # asked for, so it is trusted without the allowlist below.
        stored_target = (doc.get("tfs_org_url") or "").strip()
        org_url = target_url(doc) if stored_target else fallback_org_url
        if not org_url:
            log(f"{COLLECTION}/{doc_id}: refusing unknown project "
                f"{doc.get('tfs_org_url')!r} - it is not in this worker's "
                "TFS_TARGETS")
            patch_fields(id_token, doc_id, {
                "status": "failed",
                "last_error": ("Unknown project "
                               f"{doc.get('tfs_org_url')}"),
            })
            continue

        label = (f"{COLLECTION}/{doc_id} (WI {work_item_id or '?'} "
                 f"{target_label(doc, org_url)})")

        # A fresh claim belongs to another instance (or to this one, moments
        # ago); only an abandoned one is retried here.
        if doc.get("status") == "sending" and not is_stale_claim(doc.get("_update_time", "")):
            continue

        if attempts >= max_attempts:
            log(f"{label}: {attempts} failed attempt(s) - marking failed")
            patch_fields(id_token, doc_id, {
                "status": "failed",
                "last_error": doc.get("last_error", "") or "too many attempts",
            })
            continue

        if not work_item_id:
            log(f"{label}: no work item id - marking failed")
            patch_fields(id_token, doc_id, {
                "status": "failed",
                "last_error": "Missing work item id",
            })
            continue

        # Claim first. If another instance already claimed it, skip.
        if not patch_fields(
            id_token, doc_id, {"status": "sending"},
            expect_update_time=doc.get("_update_time"),
        ):
            log(f"{label}: claimed by another instance - skipping")
            continue

        # Whose name the comment carries, in order of preference:
        #   1. the PAT stored with the schedule - what the app writes since
        #      COPY_PAT_INTO_SCHEDULE, so the comment is authored by the person
        #      who scheduled it without this worker needing to read their
        #      users/{uid} document,
        #   2. the PAT the scheduling account keeps in users/{uid} - schedules
        #      saved by older builds,
        #   3. the worker's shared TFS_PAT - only when TFS_ALLOW_SHARED_PAT is
        #      on, because that comment is credited to the token's owner.
        owner_uid, pat, source = resolve_pat(id_token, doc, shared_pat)
        author = owner_email or owner_uid or "the author"

        refusal = ""
        if not pat:
            refusal = f"no PAT available for {author}"
        elif source == "the shared TFS_PAT" and not allow_shared_pat:
            refusal = (f"no PAT stored for {author}; using the shared TFS_PAT "
                       "would credit that token's owner instead - store their "
                       "token (sign in once in the app) or set "
                       "TFS_ALLOW_SHARED_PAT=1")
        if refusal:
            leave_pending(id_token, doc_id, label, refusal)
            continue

        log(f"{label}: posting with {source} [token {token_fingerprint(pat)}]")

        # Name the account the comment will actually be written by. Without
        # this the author is only implied by which token happened to be used,
        # which is how a comment ends up under the wrong person unnoticed.
        identity = pat_identity(pat, org_url)
        if not identity:
            log(f"{label}: could not confirm who {source} belongs to before "
                "posting - it is checked on the comment TFS returns")
        else:
            log(f"{label}: comment will appear as {identity['unique_name']}"
                + (f" ({identity['display_name']})" if identity["display_name"] else ""))
            if not author_matches(identity, owner_email):
                mismatch = (f"{identity['unique_name']} owns the token used, but "
                            f"{owner_email or 'the scheduling user'} scheduled it")
                if strict_author:
                    leave_pending(id_token, doc_id, label,
                                  "wrong author: " + mismatch)
                    continue
                log(f"{label}: WARNING {mismatch} - posting anyway, set "
                    "TFS_STRICT_AUTHOR=1 to block this")

        html = doc.get("html") or ""
        if not html.strip():
            log(f"{label}: empty comment - marking failed")
            patch_fields(id_token, doc_id, {
                "status": "failed",
                "last_error": "Comment body is empty",
            })
            continue

        try:
            response = post_comment(pat, work_item_id, html, org_url)
        except Exception as e:
            attempts += 1
            give_up = attempts >= max_attempts
            # A token captured in the cache may have been replaced since; drop
            # it so the retry reads users/{uid} again.
            if owner_uid:
                _pat_cache.pop(owner_uid, None)
            log(f"{label}: attempt {attempts} failed at {org_url}: {e}")
            patch_fields(id_token, doc_id, {
                "status": "failed" if give_up else "pending",
                "attempts": attempts,
                "last_error": f"{type(e).__name__}: {e}"[:400],
            })
            continue

        # Confirm the author from TFS's own reply. This is the one check that
        # cannot be wrong, and it needs no extra request: if a token belonging
        # to somebody else was used, the ticket is already wrong - say so
        # rather than reporting a clean success.
        recorded = comment_author(response)
        author_note = ""
        if recorded:
            log(f"{label}: posted as {recorded['unique_name']}"
                + (f" ({recorded['display_name']})" if recorded["display_name"] else ""))
            if not author_matches(recorded, owner_email):
                author_note = (f"TFS recorded {recorded['unique_name']} as the "
                               "author, not the scheduler")
                log(f"{label}: WARNING {author_note}")

        patch_fields(id_token, doc_id, {
            "status": "sent",
            "attempts": attempts,
            "posted_at": now_iso(),
            "last_error": author_note,
        })
        posted += 1
        log(f"{label}: comment posted")

    return posted, len(pending)


def check(id_token, shared_pat, fallback_org_url, allow_shared_pat=False):
    """Report who each pending schedule would post as. Writes nothing.

    This is the answer to "why is my comment under the wrong name": either the
    schedule carries no token of its own and the author's users/{uid} document
    has none the worker can read, or the token that was found belongs to
    somebody else.
    """
    pending = fetch_pending(id_token)
    log(f"{len(pending)} schedule(s) pending or sending (nothing will be posted)")

    if shared_pat:
        identity = pat_identity(shared_pat, fallback_org_url)
        if identity:
            log(f"TFS_PAT belongs to {identity['unique_name']}"
                + (f" ({identity['display_name']})" if identity["display_name"] else ""))
        else:
            log("TFS_PAT: could not resolve its TFS account - check the token "
                "is valid for " + fallback_org_url)
    else:
        log("TFS_PAT is not set - a schedule with no stored PAT cannot post")

    wrong = 0
    for doc in sorted(pending, key=lambda d: str(d.get("due_at", ""))):
        doc_id = doc["_id"]
        owner_email = (doc.get("owner_email") or "").strip()
        owner_uid, pat, source = resolve_pat(id_token, doc, shared_pat)
        stored_target = (doc.get("tfs_org_url") or "").strip()
        org_url = target_url(doc) if stored_target else fallback_org_url
        head = (
            f"scheduledComments/{doc_id}: WI {doc.get('work_item_id') or '?'}, "
            f"due {doc.get('due_at') or '?'}, {doc.get('status') or '?'}"
        )
        log(head)
        log(f"    project: "
            + (target_label(doc, org_url) if org_url
               else f"UNKNOWN ({stored_target}) - would be refused"))
        log(f"    scheduled by: {owner_email or '(no email)'} "
            f"uid={owner_uid or '(missing owner_uid)'}")
        if not pat:
            log("    token: NONE - this schedule stays pending until a token exists")
            wrong += 1
            continue
        log(f"    token: {source} [token {token_fingerprint(pat)}]")
        if source == "the shared TFS_PAT" and not allow_shared_pat:
            log("    blocked: TFS_ALLOW_SHARED_PAT is not set, so this would not "
                "post at all rather than post under the wrong author")
            wrong += 1
            continue
        identity = pat_identity(pat, org_url or fallback_org_url)
        if not identity:
            log("    author: unknown here - when this posts, the worker reports "
                "who TFS actually recorded")
            continue
        if author_matches(identity, owner_email):
            log(f"    author: {identity['unique_name']} - correct")
            continue
        wrong += 1
        log(f"    author: {identity['unique_name']} - WRONG, this posts under "
            f"that account instead of {owner_email or 'the scheduling user'}")
        if source == "the shared TFS_PAT":
            log("    fix: the schedule carries no token and users/{uid} has no "
                "readable pat. Re-save the schedule from an app build with "
                "COPY_PAT_INTO_SCHEDULE on, or let the worker read users/{uid} "
                "(see the README).")

    if wrong:
        log(f"{wrong} of {len(pending)} schedule(s) would post under the wrong "
            "author or not at all. Set TFS_STRICT_AUTHOR=1 to block the ones "
            "that would still post under the wrong name.")
    elif pending:
        log("Every pending schedule would post under its author.")
    return 0


def main(argv):
    once = "--once" in argv
    check_only = "--check" in argv

    # Look for .env next to this script, then in the working directory.
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for candidate in (os.path.join(script_dir, ".env"),
                      os.path.join(os.getcwd(), ".env")):
        if os.path.isfile(candidate):
            loaded = load_env_file(candidate)
            log(f"Loaded {len(loaded)} setting(s) from {candidate}"
                + (f": {', '.join(loaded)}" if loaded else ""))
            break
    else:
        log("No .env file found - using the host's environment variables only")

    email = os.environ.get("FIREBASE_EMAIL", "").strip()
    password = os.environ.get("FIREBASE_PASSWORD", "")
    shared_pat = os.environ.get("TFS_PAT", "").strip()
    org_url = (os.environ.get("TFS_ORG_URL") or DEFAULT_ORG_URL).rstrip("/")
    poll_seconds = int(os.environ.get("POLL_SECONDS") or 30)
    max_attempts = int(os.environ.get("MAX_ATTEMPTS") or 3)
    truthy = ("1", "true", "yes", "on")
    strict_author = (os.environ.get("TFS_STRICT_AUTHOR") or "").strip().lower() in truthy
    # Off by default: a schedule with no token of its own must not be posted
    # under the TFS_PAT owner's name, which is what silently happened before.
    allow_shared_pat = (os.environ.get("TFS_ALLOW_SHARED_PAT") or "").strip().lower() in truthy

    if not email or not password:
        log("FIREBASE_EMAIL and FIREBASE_PASSWORD must be set - add them as "
            "host variables, or put them in a .env file next to this script")
        return 2
    if not shared_pat:
        log("TFS_PAT is not set - fine: every schedule posts with its author's "
            "own token. It is only needed for older schedules saved without "
            "one, and only when TFS_ALLOW_SHARED_PAT=1.")

    log(f"AutoReport comment worker starting (id={worker_id()}, "
        f"default={org_url}, poll={poll_seconds}s)")
    log("Projects: " + ", ".join(f"{name} -> {url}"
                                 for name, url in TFS_TARGETS.items()))
    try:
        id_token, refresh_token = sign_in(email, password)
    except RuntimeError as e:
        # A bad password will not fix itself; log it clearly and exit so the
        # host stops the restart loop instead of printing a traceback.
        log(str(e))
        return 2
    last_refresh = time.time()

    if check_only:
        # Same sign-in and Firestore read as a normal poll, but no claims and
        # no comments: safe to run on the live host while the worker runs.
        check(id_token, shared_pat, org_url, allow_shared_pat)
        return 0

    # First heartbeat before any poll, so a freshly started worker shows as
    # online in the app even when nothing is due.
    write_heartbeat(id_token, poll_seconds, 0)

    while True:
        try:
            # Refresh well before the one-hour token expiry.
            if time.time() - last_refresh > 3000:
                id_token, refresh_token = refresh_session(refresh_token)
                last_refresh = time.time()
                log("Session refreshed")

            posted, waiting = run_once(
                id_token, shared_pat, org_url, max_attempts, strict_author,
                allow_shared_pat)
            write_heartbeat(id_token, poll_seconds, waiting)
        except PermissionError:
            log("Firestore refused the token - refreshing the session")
            id_token, refresh_token = refresh_session(refresh_token)
            last_refresh = time.time()
        except requests.RequestException as e:
            log(f"Network problem, retrying next poll: {e}")
        except Exception:
            # A crash here restarts the process on the host, which loses the
            # session; log and keep looping instead.
            log("Unexpected error:\n" + traceback.format_exc())

        if once:
            return 0
        time.sleep(poll_seconds)


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        log("Stopped")
        sys.exit(0)
