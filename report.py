import sys
import os
import re
import json
import base64
import difflib
import time
import shutil
import tempfile
import winreg
import ctypes
import ctypes.wintypes
import logging
import asyncio
from datetime import datetime, timezone
from io import BytesIO
from logging.handlers import RotatingFileHandler
from zoneinfo import ZoneInfo

import requests
from requests.auth import HTTPBasicAuth
from bs4 import BeautifulSoup, NavigableString, Tag
from PIL import Image
from docx import Document
from docx.shared import Inches

from PyQt6 import QtWidgets, QtCore, QtGui, sip
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal, QUrl
from PyQt6.QtGui import QDesktopServices
from PyQt6.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QApplication,
    QDialog, QFormLayout, QLineEdit, QComboBox,
    QPushButton, QHBoxLayout, QMessageBox,
    QTextEdit, QFileDialog
)

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ================= APP PATHS =================
# Everything the app reads or writes is anchored to its own directory, never
# to the current working directory. A shortcut whose "Start in" differs from
# the install folder - or a source run launched from elsewhere - used to make
# the template "disappear" and scatter report folders wherever the process
# happened to start.

def _app_dir():
    """Directory the app runs from, whether frozen or run from source."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _app_dir()
LOG_PATH = os.path.join(APP_DIR, "report_log.txt")
TEMPLATE_PATH = os.path.join(APP_DIR, "Report Template.docx")
POT_TEMPLATE_PATH = os.path.join(APP_DIR, "POT_Template.docx")
REPORTS_DIR = os.path.join(APP_DIR, "Reports")
ASSET_DIR = os.path.join(APP_DIR, "asset")


def _safe_filename(name):
    """Strip characters Windows will not accept in a file name.

    The client name and ticket title come straight from user input; a stray
    "/" or ":" made doc.save() fail with a bare PermissionError.
    """
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name or ""))
    return cleaned.strip(" .") or "Untitled"

# ================= LOGGING SETUP =================
# INFO by default, and capped: the old DEBUG + plain FileHandler combination
# wrote full tracebacks, e-mail addresses and ticket titles to a file that grew
# without limit next to the exe.
_log_handlers = []
try:
    _log_handlers.append(RotatingFileHandler(
        LOG_PATH, maxBytes=5 * 1024 * 1024, backupCount=3,
        encoding="utf-8", delay=True,
    ))
except OSError:
    pass  # read-only install folder: the file log is best effort

# A windowed PyInstaller build has sys.stdout/sys.stderr set to None, so a
# console handler would raise on the first log record.
if getattr(sys, "stderr", None) is not None:
    _log_handlers.append(logging.StreamHandler(sys.stderr))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=_log_handlers,
)
logger = logging.getLogger("report")

org_url = "https://tfs.alliancewebpos.com/tfs/WebPOSCollection/WebPOS"
collection_url = "https://tfs.alliancewebpos.com/tfs/WebPOSCollection"
_TFS_HOST = collection_url.split("/tfs/")[0]
template_path = TEMPLATE_PATH

# ================= HTTP TIME-OUTS =================
# requests has no default timeout, so without these a stalled connection
# blocks a worker thread forever: no error, no cancel, no result.
HTTP_TIMEOUT = 30      # TFS / Azure DevOps REST calls
IMAGE_TIMEOUT = 20     # images embedded in a work item comment
DOWNLOAD_TIMEOUT = 120 # release assets: bigger, slower, worth waiting for

# Update bookkeeping, all inside the app directory.
UPDATE_PREV_DIR = "_prev"              # one generation back, for rollback
UPDATE_TRIAL_FILE = "_update_trial.json"    # written before the restart
UPDATE_STARTED_FILE = "_update_started.flag"  # written by the first start after it
UPDATE_OK_FILE = "_update_ok.flag"       # written once a window is up
UPDATE_FAILED_FILE = "_update_failed.txt"  # written by the swap script

# ================= AUTO-UPDATE CONFIG =================
APP_VERSION = "4.0.0"
GITHUB_REPO = "shinispades/AutoReport"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


# ================= TOOL REGISTRY =================
# To add a new tool, append a dict to this list:
#   {
#       "id": "unique_tool_id",          # unique identifier
#       "name": "Display Name",          # shown on the card
#       "description": "Short desc...",  # shown on the card
#       "icon": "🔧",                    # emoji or text icon
#       "handler": "method_name",        # method name on LoginWindow to call
#       "category": "General",           # optional category for grouping
#   }
# Then implement the handler method on LoginWindow.
# The Tools page auto-fetches all tools from this list.
#
# Tools listed here are still placeholders: the page marks them "coming soon"
# and they do nothing when clicked. Empty them out as each tool is built.

TOOL_REGISTRY = [
    # --- Tool 1 ---
    {
        "id": "tool_1",
        "name": "Missing Sales Generator",
        "description": "Generate diagnostic SQL and helper scripts for missing POS sales",
        "icon": "🗃️",
        "handler": "run_tool_1",
        "category": "Reports",
    },
    # --- Tool 2: Ayala Checker ---
    {
        "id": "tool_2",
        "name": "Ayala Checker",
        "description": "Compare hourly sales CSVs against EOD CSV and generate organized Excel report",
        "icon": "📊",
        "handler": "run_tool_2",
        "category": "Reports",
    },
    # --- Scheduled TFS Comment ---
    {
        "id": "scheduled_comment",
        "name": "Scheduled Comment",
        "description": "Write a TFS comment with images and links, then let the hosted worker post it at the time you choose",
        "icon": "⏰",
        "handler": "run_scheduled_comment",
        "category": "Reports",
    },
    # --- PTU Generator ---
    {
        "id": "ptu_generator",
        "name": "PTU Generator",
        "description": "Read a PTU permit from a PDF or an image and show the text the app can see in it",
        "icon": "🧾",
        "handler": "run_ptu_generator",
        "category": "Reports",
    },
    # --- Tool 3 ---
    {
        "id": "tool_3",
        "name": "Tool 3",
        "description": "Description for Tool 3",
        "icon": "📁",
        "handler": "run_tool_3",
        "category": "Reports",
    },
    # --- Tool 4 ---
    {
        "id": "tool_4",
        "name": "Tool 4",
        "description": "Description for Tool 4",
        "icon": "🔍",
        "handler": "run_tool_4",
        "category": "Reports",
    },
    # --- Tool 5 ---
    {
        "id": "tool_5",
        "name": "Tool 5",
        "description": "Description for Tool 5",
        "icon": "⚙️",
        "handler": "run_tool_5",
        "category": "Utilities",
    },
    # --- Tool 6 ---
    {
        "id": "tool_6",
        "name": "Tool 6",
        "description": "Description for Tool 6",
        "icon": "📝",
        "handler": "run_tool_6",
        "category": "Utilities",
    },
    # --- Tool 7 ---
    {
        "id": "tool_7",
        "name": "Tool 7",
        "description": "Description for Tool 7",
        "icon": "🗂️",
        "handler": "run_tool_7",
        "category": "Utilities",
    },
    # --- Tool 8 ---
    {
        "id": "tool_8",
        "name": "Tool 8",
        "description": "Description for Tool 8",
        "icon": "📈",
        "handler": "run_tool_8",
        "category": "Analytics",
    },
    # --- Tool 9 ---
    {
        "id": "tool_9",
        "name": "Tool 9",
        "description": "Description for Tool 9",
        "icon": "🛠️",
        "handler": "run_tool_9",
        "category": "Analytics",
    },
    # --- Tool 10 ---
    {
        "id": "tool_10",
        "name": "Tool 10",
        "description": "Description for Tool 10",
        "icon": "📦",
        "handler": "run_tool_10",
        "category": "Analytics",
    },
    # === ADD MORE TOOLS BELOW ===
    # {
    #     "available": True,
    #     "id": "my_new_tool",
    #     "name": "My New Tool",
    #     "description": "What it does",
    #     "icon": "✨",
    #     "handler": "run_my_new_tool",
    #     "category": "General",
    # },
]


UNFINISHED_TOOL_IDS = {f"tool_{n}" for n in range(3, 11)}


def _fetch_pat_expiry(pat):
    """
    Query the TFS token API to get the expiry date of the given PAT.
    Returns an ISO date string (e.g. '2025-08-01T00:00:00Z') or None.

    Only the token that is actually in use is reported. The old fallback to
    `pats[0]` handed back some other PAT's expiry date, which then appeared in
    Settings as if it belonged to the signed-in user - worse than no date.
    """
    if not pat:
        return None

    url = f"{collection_url}/_apis/tokens/pats?api-version=7.1-preview.1"
    logger.info("Fetching PAT expiry from %s", url)
    try:
        resp = requests.get(url, auth=HTTPBasicAuth("", pat), timeout=HTTP_TIMEOUT)
        if resp.status_code == 403:
            logger.warning("PAT lacks 'vso.tokens_manage' scope — cannot read expiry (403)")
            return None
        if resp.status_code != 200:
            logger.warning("Token API returned status %d: %s", resp.status_code, resp.text[:200])
            return None

        data = resp.json()
        pats = data.get("patTokens", data.get("value", []))
        logger.debug("Token API returned %d PAT(s)", len(pats))
        short = pat[:4]
        for token in pats:
            if (token.get("displayTokenId") or "").startswith(short):
                expiry = token.get("validTo", "")
                logger.info("PAT expiry found: %s", expiry)
                return expiry or None

        logger.warning("None of the %d PAT(s) matched the token in use - expiry unknown", len(pats))
        return None
    except Exception:
        logger.exception("Failed to fetch PAT expiry")
        return None


def _check_for_update():
    """
    Check GitHub for the latest release. Returns a dict with
    'version', 'download_url', 'asset_type', and 'body' if an update
    is available, or None if already up-to-date or the check fails.
    """
    logger.info("Checking for update: current=%s, repo=%s", APP_VERSION, GITHUB_REPO)
    try:
        resp = requests.get(GITHUB_API_URL, timeout=10, headers={
            "Accept": "application/vnd.github+json"
        })
        if resp.status_code != 200:
            logger.warning("GitHub API returned %d", resp.status_code)
            return None

        data = resp.json()
        tag = data.get("tag_name", "").lstrip("v")
        if not tag:
            logger.warning("No tag_name in release response")
            return None

        logger.info("Latest release: %s (current: %s)", tag, APP_VERSION)

        # Version comparison on the leading numeric part only. A tag like
        # "v3.1.0-beta" used to raise ValueError, fall back to [0], and compare
        # as "older" - which silently disabled updates for any prerelease tag.
        def parse_ver(v):
            match = re.match(r"(\d+(?:\.\d+)*)", (v or "").strip())
            return [int(p) for p in match.group(1).split(".")] if match else [0]

        def ver_key(v):
            """Zero-padded so 3.1 and 3.1.0 are the same version. Compared
            as plain lists, [3, 1, 0] > [3, 1] and the app offered an
            "update" to the build it was already running."""
            parts = parse_ver(v)
            while len(parts) < 4:
                parts.append(0)
            return parts

        if ver_key(tag) <= ver_key(APP_VERSION):
            logger.info("Already up-to-date")
            return None

        # Prefer a .zip (extracted in place), fall back to an .exe installer.
        # Whatever is picked, carry its advertised size and digest through:
        # they are the only way to tell a good download from a truncated one.
        assets = data.get("assets", [])
        chosen = None
        for wanted in (".zip", ".exe"):
            for asset in assets:
                if asset.get("name", "").lower().endswith(wanted):
                    chosen = asset
                    break
            if chosen:
                break

        if not chosen:
            logger.warning("No .zip or .exe asset found in release %s", tag)
            return None

        download_url = chosen.get("browser_download_url")
        if not download_url:
            logger.warning("Asset %r has no download URL", chosen.get("name"))
            return None

        asset_type = "zip" if chosen.get("name", "").lower().endswith(".zip") else "exe"
        logger.info("Update asset found: type=%s, url=%s", asset_type, download_url[:80])
        logger.info("Update asset metadata: size=%r digest=%r",
                    chosen.get("size"), chosen.get("digest"))
        return {
            "version": tag,
            "download_url": download_url,
            "asset_type": asset_type,
            "body": data.get("body", ""),
            "size": chosen.get("size") or None,
            "digest": chosen.get("digest") or None,
        }
    except Exception:
        logger.exception("Update check failed")
        return None


def _download_update(download_url, progress_callback=None,
                     expected_size=None, expected_digest=None, attempts=3,
                     local_name=None):
    """
    Download an update and verify it before returning the path.

    Retries transient failures, writes to a .part file so an interrupted
    download can never be mistaken for a finished one, and checks the size and
    SHA-256 the release advertises. Returns the path, or None on failure.

    local_name overrides the on-disk name, which is how an .exe release is
    saved without an .exe extension (see _staged_update_name).
    """
    import hashlib

    filename = (local_name or os.path.basename(download_url.split("?")[0])
                or "update.zip")
    # One private directory per download. The shared %TEMP% let a second
    # session - or another user on the same machine - collide with, or
    # pre-place, the file this process is about to execute.
    download_dir = tempfile.mkdtemp(prefix="autoreport_update_")
    tmp_path = os.path.join(download_dir, filename)
    part_path = tmp_path + ".part"

    for attempt in range(1, attempts + 1):
        logger.info("Downloading update (attempt %d/%d): %s", attempt, attempts, download_url)
        try:
            digest = hashlib.sha256()
            downloaded = 0
            with requests.get(download_url, stream=True, timeout=DOWNLOAD_TIMEOUT) as resp:
                resp.raise_for_status()
                total = int(resp.headers.get("content-length", 0))
                with open(part_path, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=65536):
                        if not chunk:
                            continue
                        f.write(chunk)
                        digest.update(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total:
                            progress_callback(downloaded, total)

            # A truncated download is still a valid .zip prefix far less often
            # than it is a corrupt one, so check both claims the release makes.
            if expected_size and downloaded != int(expected_size):
                raise ValueError(
                    f"size mismatch: got {downloaded} bytes, expected {expected_size}"
                )
            if expected_digest:
                expected_hex = str(expected_digest).split(":", 1)[-1].strip().lower()
                if digest.hexdigest() != expected_hex:
                    raise ValueError("SHA-256 mismatch - the download is corrupt or tampered with")

            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            os.replace(part_path, tmp_path)
            logger.info("Update downloaded to %s (%d bytes, sha256=%s)",
                        tmp_path, downloaded, digest.hexdigest())
            return tmp_path

        except Exception as e:
            logger.exception("Download attempt %d failed: %s", attempt, e)
            try:
                if os.path.exists(part_path):
                    os.remove(part_path)
            except OSError:
                logger.warning("Could not remove partial download: %s", part_path)
            if attempt >= attempts:
                return None
            time.sleep(2 * attempt)

    return None


def _clear_update_markers(app_dir):
    """
    Forget the in-flight update markers. The _prev/ backup is kept, and so is
    _update_failed.txt: _cleanup_update_artifacts still has to read and log it.
    """
    for name in (UPDATE_TRIAL_FILE, UPDATE_STARTED_FILE, UPDATE_OK_FILE):
        try:
            os.remove(os.path.join(app_dir, name))
        except OSError:
            pass


def _mark_update_ok():
    """
    Called once a window is actually on screen. Clearing the markers here is
    what tells the next start that the applied update is usable.
    """
    app_dir = _app_dir()
    trial_path = os.path.join(app_dir, UPDATE_TRIAL_FILE)
    if not os.path.isfile(trial_path):
        return

    try:
        os.remove(trial_path)
        _clear_update_markers(app_dir)
        logger.info("Update confirmed working - trial markers cleared")
        return
    except OSError:
        pass

    # Could not remove it. Leave a flag recording that a window really did
    # appear, so the next start does not roll a working build back.
    try:
        with open(os.path.join(app_dir, UPDATE_OK_FILE), "w", encoding="utf-8") as f:
            f.write(datetime.now().isoformat())
        logger.warning("Could not clear the update trial marker - wrote %s instead",
                       UPDATE_OK_FILE)
    except OSError:
        logger.warning("Could not write the update confirmation flag")


def _restore_from_backup(prev_dir, app_dir, names):
    """
    Copy backed-up originals back over whatever an update already replaced.
    Best effort: anything still locked is skipped and logged.
    """
    running_exe = (os.path.basename(sys.executable)
                   if getattr(sys, 'frozen', False) else "")
    restored = 0
    for name in names:
        src = os.path.join(prev_dir, name)
        dst = os.path.join(app_dir, name)
        if not os.path.exists(src):
            continue
        if running_exe and os.path.normcase(name) == os.path.normcase(running_exe):
            # The exe running this code cannot be written over, and it was
            # never replaced anyway - a locked file is staged as .new instead.
            # Counting it as a failed rollback only hid the real error.
            logger.info("Left the running %s in place - it was never replaced", name)
            continue
        try:
            if os.path.isdir(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst, ignore_errors=True)
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
            restored += 1
            logger.info("Rolled back: %s", name)
        except Exception:
            logger.exception("Rollback failed for %s - restore it from %s", name, prev_dir)
    return restored


def _restore_previous_version(app_dir):
    """
    Put the files saved in _prev/ back in place. Anything still locked - the
    running exe - is staged as .new and swapped in by the batch script after
    this process exits, which is the same mechanism a normal update uses.

    Returns the number of items handled.
    """
    prev_dir = os.path.join(app_dir, UPDATE_PREV_DIR)
    if not os.path.isdir(prev_dir):
        return 0

    handled = 0
    locked = []
    for name in os.listdir(prev_dir):
        src = os.path.join(prev_dir, name)
        dst = os.path.join(app_dir, name)
        try:
            if os.path.isdir(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
            handled += 1
        except PermissionError:
            staged = dst + ".new"
            try:
                if os.path.isdir(src):
                    if os.path.exists(staged):
                        shutil.rmtree(staged)
                    shutil.copytree(src, staged)
                else:
                    if os.path.exists(staged):
                        os.remove(staged)
                    shutil.copy2(src, staged)
                locked.append((staged, dst))
                handled += 1
            except Exception:
                logger.exception("Could not stage rollback for %s", dst)
        except Exception:
            logger.exception("Could not roll back %s", name)

    if locked:
        exe_name = os.path.basename(sys.executable) if getattr(sys, 'frozen', False) else None
        _write_swap_script(os.path.join(app_dir, "_update_swap.bat"), app_dir, exe_name, locked)
    return handled


def _recover_failed_update():
    """
    Run once at startup, before the UI is built.

    An update writes a trial marker before restarting. If the next launch never
    reaches a visible window, the marker is still there - which means the new
    build is broken. The previous version is restored and the app relaunches
    into it. Returns True when a relaunch is pending (the caller must exit).
    """
    app_dir = _app_dir()
    trial_path = os.path.join(app_dir, UPDATE_TRIAL_FILE)
    started_path = os.path.join(app_dir, UPDATE_STARTED_FILE)

    if not os.path.isfile(trial_path):
        # Nothing pending. Drop any stray markers from an interrupted run.
        _clear_update_markers(app_dir)
        return False

    # A window has already been on screen for this build - nothing to recover.
    if os.path.isfile(os.path.join(app_dir, UPDATE_OK_FILE)):
        _clear_update_markers(app_dir)
        return False

    # First start of the updated build. Record that it got this far and let it
    # run: the marker written here is what the NEXT start looks for. Without
    # this phase every successful update would be rolled back immediately,
    # because the trial marker is written before the restart, not after it.
    if not os.path.isfile(started_path):
        try:
            with open(started_path, "w", encoding="utf-8") as f:
                f.write(datetime.now().isoformat())
            logger.info("First start after an update - watching this build")
        except OSError:
            logger.warning("Could not write the update start marker")
        return False

    if not os.path.isdir(os.path.join(app_dir, UPDATE_PREV_DIR)):
        logger.warning("Update trial marker found, but there is no backup to restore")
        _clear_update_markers(app_dir)
        return False

    logger.error("The updated build never reached a window - restoring the previous version")
    handled = _restore_previous_version(app_dir)

    # Clear the marker *before* relaunching: if the restored build also fails,
    # the next start must not roll back again, or it would loop forever.
    _clear_update_markers(app_dir)

    if handled == 0:
        logger.error("Nothing could be restored from %s", UPDATE_PREV_DIR)
        return False

    bat_path = os.path.join(app_dir, "_update_swap.bat")
    if not os.path.isfile(bat_path):
        # Nothing was locked, so the files are already restored in place.
        logger.info("Rolled back %d item(s) without a relaunch script", handled)
        return False

    logger.info("Rolled back %d item(s); the relaunch script will start the restored build", handled)
    return True


def _apply_zip_update(zip_path, version=""):
    """
    Extract a zip update and replace files in the current app directory.

    Order is deliberate: validate the payload, back up everything it will
    touch, and only then replace anything. A failure at any point restores the
    backup, so a bad update leaves the old install exactly as it was instead of
    half old and half new.

    Locked files (the running exe) are staged as .new and swapped by a batch
    script after the app exits. Returns (ok, message).
    """
    import zipfile

    app_dir = _app_dir()
    exe_name = os.path.basename(sys.executable) if getattr(sys, 'frozen', False) else None

    logger.info("Applying zip update: %s -> %s", zip_path, app_dir)

    extract_dir = os.path.join(tempfile.gettempdir(), "update_extract")
    prev_dir = os.path.join(app_dir, UPDATE_PREV_DIR)
    # Defined before the try so the failure handler below can still clean up if
    # extraction itself is what raised.
    entries = []

    try:
        if os.path.exists(extract_dir):
            shutil.rmtree(extract_dir)
        os.makedirs(extract_dir)

        with zipfile.ZipFile(zip_path, 'r') as zf:
            members = zf.namelist()
            root_real = os.path.realpath(extract_dir)
            for name in members:
                member_real = os.path.realpath(os.path.join(extract_dir, name))
                if member_real != root_real and not member_real.startswith(root_real + os.sep):
                    logger.error("Rejecting unsafe zip member: %r", name)
                    shutil.rmtree(extract_dir, ignore_errors=True)
                    return False, "The update archive contains an unsafe path and was rejected."
            zf.extractall(extract_dir)
            logger.info("Extracted %d member(s) from zip", len(members))

        # Some release zips wrap everything in a single top-level folder.
        # Descend into it so the payload lands in the app directory itself.
        payload_dir = extract_dir
        entries = os.listdir(payload_dir)
        while len(entries) == 1 and os.path.isdir(os.path.join(payload_dir, entries[0])):
            payload_dir = os.path.join(payload_dir, entries[0])
            entries = os.listdir(payload_dir)
        logger.info("Update payload root: %s (%d top-level item(s))", payload_dir, len(entries))

        # Refuse a payload that clearly isn't this app. An empty or wrong zip
        # used to "apply" successfully and restart into the same version.
        anchor = exe_name or "report.py"
        if not entries or anchor not in entries:
            logger.error("Update payload has no %r - refusing to apply", anchor)
            shutil.rmtree(extract_dir, ignore_errors=True)
            return False, (
                f"The update archive does not contain {anchor}, so it was not "
                "applied. The current installation is unchanged."
            )

        # --- Back up everything the payload will overwrite ---
        # Start from an empty backup dir so it only ever holds the generation
        # immediately before this one: a stale entry from an older round would
        # otherwise be restored over a newer file.
        if os.path.isdir(prev_dir):
            shutil.rmtree(prev_dir, ignore_errors=True)
        os.makedirs(prev_dir, exist_ok=True)
        for item in entries:
            dst = os.path.join(app_dir, item)
            if not os.path.exists(dst):
                continue
            backup = os.path.join(prev_dir, item)
            try:
                if os.path.isdir(dst):
                    if os.path.exists(backup):
                        shutil.rmtree(backup)
                    shutil.copytree(dst, backup)
                else:
                    shutil.copy2(dst, backup)
            except Exception as e:
                logger.exception("Backup failed for %s - aborting before any change", item)
                shutil.rmtree(extract_dir, ignore_errors=True)
                return False, (
                    f"Could not back up {item} before updating ({e}). "
                    "Nothing was changed."
                )

        # --- Replace ---
        copied = 0
        locked_files = []
        for item in entries:
            src_item = os.path.join(payload_dir, item)
            dst = os.path.join(app_dir, item)
            try:
                if os.path.isdir(src_item):
                    if os.path.exists(dst):
                        shutil.rmtree(dst)
                    shutil.copytree(src_item, dst)
                else:
                    shutil.copy2(src_item, dst)
                copied += 1
                logger.debug("Replaced: %s", dst)
            except PermissionError:
                # File is locked (most likely the running exe) - stage as .new
                staged = dst + ".new"
                try:
                    if os.path.isdir(staged):
                        shutil.rmtree(staged)
                    elif os.path.exists(staged):
                        os.remove(staged)
                except OSError:
                    logger.warning("Could not remove stale .new file: %s", staged)
                try:
                    if os.path.isdir(src_item):
                        shutil.copytree(src_item, staged)
                    else:
                        shutil.copy2(src_item, staged)
                    locked_files.append((staged, dst))
                    copied += 1
                    logger.info("Staged locked file: %s -> %s", src_item, staged)
                except Exception:
                    logger.exception("Failed to stage locked file: %s -> %s", src_item, staged)
                    raise
            except Exception:
                # Anything unexpected part-way through: put the originals back.
                logger.exception("Update failed while replacing %s - rolling back", item)
                restored = _restore_from_backup(prev_dir, app_dir, entries)
                shutil.rmtree(extract_dir, ignore_errors=True)
                return False, (
                    f"Updating failed while replacing {item}. "
                    f"Restored {restored} item(s) from the previous version."
                )

        # Re-stage any .new file left over from an earlier attempt so the
        # swap script retries it instead of leaving it behind forever.
        for item in entries:
            dst = os.path.join(app_dir, item)
            staged = dst + ".new"
            if os.path.isfile(staged) and (staged, dst) not in locked_files:
                locked_files.append((staged, dst))

        if locked_files:
            bat_path = os.path.join(app_dir, "_update_swap.bat")
            _write_swap_script(bat_path, app_dir, exe_name, locked_files)
            logger.info("Created update swap script for %d locked file(s): %s",
                        len(locked_files), bat_path)

        # Trial marker: the next start has to reach a visible window, otherwise
        # _recover_failed_update() puts the previous version back.
        try:
            with open(os.path.join(app_dir, UPDATE_TRIAL_FILE), "w", encoding="utf-8") as f:
                json.dump({"version": version, "at": datetime.now().isoformat()}, f)
        except OSError:
            logger.warning("Could not write the update trial marker")

        try:
            os.remove(os.path.join(app_dir, UPDATE_FAILED_FILE))
        except OSError:
            pass

        shutil.rmtree(extract_dir, ignore_errors=True)
        try:
            os.remove(zip_path)
        except OSError:
            logger.warning("Could not remove downloaded update: %s", zip_path)
        logger.info("Zip update applied: %d item(s) processed, %d locked",
                    copied, len(locked_files))
        return True, ""

    except Exception as e:
        logger.exception("Failed to apply zip update")
        restored = 0
        if os.path.isdir(prev_dir):
            restored = _restore_from_backup(prev_dir, app_dir, os.listdir(prev_dir))
        # Drop anything this attempt staged as .new. Left behind, a release that
        # ships only templates would later swap an old exe in over the current
        # one, and the exe already has a matching copy in _prev/.
        for item in entries:
            staged = os.path.join(app_dir, item) + ".new"
            try:
                if os.path.isdir(staged):
                    shutil.rmtree(staged)
                elif os.path.exists(staged):
                    os.remove(staged)
                    logger.info("Removed staged file from the failed attempt: %s", staged)
            except OSError:
                logger.warning("Could not remove staged file: %s", staged)
        shutil.rmtree(extract_dir, ignore_errors=True)
        return False, f"The update could not be applied ({e}). Restored {restored} item(s) from _prev."



def _staged_update_name(version):
    """File name a downloaded .exe update is parked under.

    Deliberately without a .exe extension. The file is never launched under
    this name - the swap script renames it to the real exe name after this
    process exits - and an extension-based scanner heuristic on a freshly
    downloaded .exe is exactly what this avoids.
    """
    base = (os.path.basename(sys.executable) if getattr(sys, 'frozen', False)
            else "report.py")
    return f"{base}.update.v{version or 'new'}"


def _apply_exe_update(downloaded_path, version=""):
    """Park a verified .exe update beside the app for the swap script to rename.

    Only the executable is replaced: templates and `asset/` still come from a
    .zip release.

    Order is deliberate, and matches _apply_zip_update: back the current exe up
    into _prev/ first, then move the download into place under its non-.exe
    name, then write the script that does the rename. A build that never
    reaches a window is rolled back from _prev/ by _recover_failed_update().

    Returns (ok, message).
    """
    if not getattr(sys, 'frozen', False):
        return False, ("This copy runs from source, so a .exe update cannot "
                       "replace it. Publish a .zip release instead.")

    app_dir = _app_dir()
    exe_name = os.path.basename(sys.executable)
    target = os.path.join(app_dir, exe_name)
    prev_dir = os.path.join(app_dir, UPDATE_PREV_DIR)
    staged = os.path.join(app_dir, _staged_update_name(version))

    logger.info("Applying exe update: %s -> %s", downloaded_path, staged)

    if not os.path.isfile(target):
        return False, f"{exe_name} was not found in {app_dir}, so nothing was replaced."

    try:
        # One generation back only, same as the zip path: a stale entry from an
        # older round must not be restored over a newer file.
        if os.path.isdir(prev_dir):
            shutil.rmtree(prev_dir, ignore_errors=True)
        os.makedirs(prev_dir, exist_ok=True)
        shutil.copy2(target, os.path.join(prev_dir, exe_name))

        if os.path.exists(staged):
            os.remove(staged)
        shutil.move(downloaded_path, staged)
    except Exception as e:
        logger.exception("Could not stage the exe update")
        return False, f"The update could not be staged ({e}). Nothing was replaced."

    try:
        _write_swap_script(os.path.join(app_dir, "_update_swap.bat"),
                           app_dir, exe_name, [(staged, target)])
    except Exception as e:
        logger.exception("Could not write the swap script")
        return False, (f"The update was saved as {staged}, but the swap script "
                       f"could not be written ({e}).")

    # Trial marker: the next start has to reach a visible window, otherwise
    # _recover_failed_update() puts the exe saved in _prev/ back.
    try:
        with open(os.path.join(app_dir, UPDATE_TRIAL_FILE), "w", encoding="utf-8") as f:
            json.dump({"version": version, "at": datetime.now().isoformat()}, f)
    except OSError:
        logger.warning("Could not write the update trial marker")

    try:
        os.remove(os.path.join(app_dir, UPDATE_FAILED_FILE))
    except OSError:
        pass

    logger.info("Exe update staged for the swap script: %s", staged)
    return True, ""


def _bat_quote(path):
    """
    Quote a path for the generated swap script. Doubling any percent sign
    stops cmd.exe from reading it as an environment variable reference.
    """
    return '"%s"' % str(path).replace("%", "%%")


def _write_swap_script(bat_path, app_dir, exe_name, locked_files):
    """
    Write the batch script that replaces locked files after the app exits.

    cmd.exe quirks handled here:
      * CRLF line endings - LF-only .bat files are parsed unreliably.
      * Delays use `ping`, because `timeout` errors out with no console.
      * One `:replace` subroutine is `call`ed per file - batch labels must
        be unique, so a retry loop cannot be inlined more than once.
      * The app is relaunched with `start /d`, so it opens in its own folder
        even when the shortcut's "Start in" points somewhere else.
      * Arguments are quoted through _bat_quote(), which doubles percent
        signs so a literal "%" in the install path is not read as a variable.
      * Windows tools are called by full path: Git for Windows ships a Unix
        `find`, which shadows find.exe and would make the wait loop give up
        before the running app has released its files.
    """
    pid = os.getpid()
    target = os.path.join(app_dir, exe_name or "report.py")

    # Full paths for every Windows tool: a Unix `find` earlier in PATH (Git for
    # Windows ships one) would fail, and the wait loop would stop waiting for
    # the running app to exit.
    sys32 = os.path.join("%SystemRoot%", "System32")
    tasklist = os.path.join(sys32, "tasklist.exe")
    find = os.path.join(sys32, "find.exe")
    taskkill = os.path.join(sys32, "taskkill.exe")
    ping = os.path.join(sys32, "ping.exe")

    lines = [
        "@echo off",
        "rem AutoReport update swap script",
        # cmd.exe decodes the file with the *current* code page, so UTF-8 is
        # only safe with this as the first command. The script used to be
        # written with the ANSI code page, which mangles any install path
        # containing non-ASCII characters (C:\\Users\\Maria with an accent,
        # or a name in another script) - the swap then silently targeted a
        # path that does not exist.
        "chcp 65001 >nul",
        'set "PID=%s"' % pid,
        "",
        "rem --- wait up to ~40s for the running app to exit ---",
        "set _tries=0",
        ":waitloop",
        '"%s" /nh /fi "PID eq %%PID%%" 2>nul | "%s" /i "%%PID%%" >nul' % (tasklist, find),
        "if errorlevel 1 goto swap",
        "set /a _tries=%_tries%+1",
        "if %_tries% geq 20 goto forcekill",
        '"%s" -n 2 127.0.0.1 >nul' % ping,
        "goto waitloop",
        ":forcekill",
        '"%s" /f /pid %%PID%% >nul 2>&1' % taskkill,
        "",
        ":swap",
        "rem --- give the OS a moment to release file handles ---",
        '"%s" -n 3 127.0.0.1 >nul' % ping,
    ]
    failed_path = os.path.join(app_dir, UPDATE_FAILED_FILE)

    for staged, dst in locked_files:
        lines.append('call :replace %s %s' % (_bat_quote(staged), _bat_quote(dst)))
    lines += [
        "",
        # The app is relaunched even when a replace failed: a missing window is
        # worse than an update that did not take. The failure is recorded in
        # _update_failed.txt, which the next start reads and logs.
        'start "" /d %s %s' % (_bat_quote(app_dir), _bat_quote(target)),
        "exit /b 0",
        "",
        "rem --- replace one file, retrying while it stays locked ---",
        ":replace",
        "set _rtries=0",
        ":replace_loop",
        'move /y "%~1" "%~2" >nul 2>&1',
        'if not exist "%~1" exit /b 0',
        "set /a _rtries=%_rtries%+1",
        "if %_rtries% geq 20 goto replace_giveup",
        '"%s" -n 2 127.0.0.1 >nul' % ping,
        "goto replace_loop",
        ":replace_giveup",
        # %%~2: the trailing %% is consumed by the %-formatting below, otherwise
        # Python reads "%~" as a format spec and raises ValueError instead of
        # writing the script.
        'echo "%%~2">> %s' % _bat_quote(failed_path),
        "exit /b 1",
    ]

    data = "\r\n".join(lines) + "\r\n"
    # UTF-8 on purpose: the chcp 65001 above makes cmd read it that way.
    # "mbcs" is deprecated in 3.11+ and cannot represent every install path.
    with open(bat_path, "w", encoding="utf-8", newline="") as f:
        f.write(data)
    logger.info("Wrote swap script (%d file(s) to replace): %s",
                len(locked_files), bat_path)


# ================= FIREBASE =================
# Sign-in uses the public web config below: that apiKey is meant to be shipped
# and is safe in a client. Data access no longer uses the Admin SDK.
#
# Shipping serviceAccountKey.json inside the release handed every person who
# downloaded the app full read/write access to the whole Firestore project.
# Profiles are now read and written as the signed-in user, with that user's ID
# token, so Firestore security rules decide what is reachable. Requirements:
#
#   match /users/{uid} {
#     allow read, write: if request.auth != null && request.auth.uid == uid;
#   }
#
# The users/{uid} rules must stay owner-only: with PAT_SYNC_TO_FIRESTORE on,
# the TFS PAT is kept in that document. Anyone who can read the collection -
# a leaked key, a wrong rule - is then holding the token. See _save_pat_remote().
# The PTU tool's Gemini API key rides in the same owner-only document (field
# gemini_api_key), for exactly the same reason: it must follow the account
# without becoming readable by anyone but its owner. See _save_ai_key_remote().
#
# Auth is two plain HTTP endpoints (Identity Toolkit + Secure Token), so the
# unmaintained pyrebase package - and the gcloud/oauth2client/pycryptodome
# chain it drags in - is no longer a dependency at all.

_firebase_config = {
  "apiKey": "AIzaSyD-uAl1Ce-2NRoAVVYExH5-MSJ6K4y-LEo",
  "authDomain": "autoreport-db4a7.firebaseapp.com",
  "projectId": "autoreport-db4a7",
  "databaseURL": "",
  "storageBucket": "autoreport-db4a7.firebasestorage.app",
  "messagingSenderId": "173593376987",
  "appId": "1:173593376987:web:66412a51ebb20818a9b4fe",
  "measurementId": "G-F2VL1BBRGT"
}
FIRESTORE_API = (
    "https://firestore.googleapis.com/v1/projects/"
    f"{_firebase_config['projectId']}/databases/(default)/documents"
)


FIREBASE_API_KEY = _firebase_config["apiKey"]

# Scheduled comments: written on this PC, stored in Firestore, posted by
# scheduled_comment_worker.py running on an always-on host. The app writes
# them; the worker only reads pending ones and updates their status. Both talk
# to Firestore as an authenticated user, so the security rules decide who may
# touch what - still no service account key anywhere.
SCHEDULED_COMMENTS_COLLECTION = "scheduledComments"
IDENTITY_API = "https://identitytoolkit.googleapis.com/v1/accounts"
SECURE_TOKEN_API = "https://securetoken.googleapis.com/v1/token"


class FirebaseAuthError(Exception):
    """Sign-in / sign-up failure. `code` is Firebase's own error code."""

    def __init__(self, code, detail=""):
        self.code = code or "AUTH_FAILED"
        self.detail = detail or self.code
        super().__init__(self.code)


def _auth_post(url, payload):
    """POST to Identity Toolkit / Secure Token and return the JSON body."""
    try:
        resp = requests.post(url, json=payload, timeout=HTTP_TIMEOUT)
    except requests.RequestException as e:
        raise FirebaseAuthError(
            "NETWORK", f"Could not reach Firebase ({e.__class__.__name__})."
        ) from e

    try:
        data = resp.json()
    except ValueError:
        data = {}

    if resp.status_code >= 400:
        raw = (data.get("error") or {}).get("message") or f"HTTP {resp.status_code}"
        # Codes arrive as "WEAK_PASSWORD : Password should be at least 6 ..."
        logger.warning("Firebase auth error: %s", raw)
        raise FirebaseAuthError(raw.split(":", 1)[0].strip(), raw)
    return data


def _sign_in(email, password):
    """Email/password sign-in. Returns uid, id token and refresh token."""
    data = _auth_post(
        f"{IDENTITY_API}:signInWithPassword?key={FIREBASE_API_KEY}",
        {"email": email, "password": password, "returnSecureToken": True},
    )
    return {
        "uid": data.get("localId", ""),
        "id_token": data.get("idToken", ""),
        "refresh_token": data.get("refreshToken", ""),
    }


def _sign_up(email, password):
    """Create an account. Returns the same shape as _sign_in()."""
    data = _auth_post(
        f"{IDENTITY_API}:signUp?key={FIREBASE_API_KEY}",
        {"email": email, "password": password, "returnSecureToken": True},
    )
    return {
        "uid": data.get("localId", ""),
        "id_token": data.get("idToken", ""),
        "refresh_token": data.get("refreshToken", ""),
    }


class FirebaseError(Exception):
    """A Firestore failure with a message that is fit to show a user."""

    def __init__(self, message, status=None):
        super().__init__(message)
        self.status = status


def _fs_value(value):
    """Encode a Python value as a Firestore REST Value."""
    if value is None:
        return {"nullValue": None}
    if isinstance(value, bool):
        return {"booleanValue": value}
    if isinstance(value, int):
        return {"integerValue": str(value)}
    if isinstance(value, float):
        return {"doubleValue": value}
    return {"stringValue": str(value)}


def _fs_decode(value):
    """Decode one Firestore REST Value into a Python value."""
    if not isinstance(value, dict):
        return ""
    if "stringValue" in value:
        return value["stringValue"]
    if "booleanValue" in value:
        return bool(value["booleanValue"])
    if "integerValue" in value:
        try:
            return int(value["integerValue"])
        except (TypeError, ValueError):
            return 0
    if "doubleValue" in value:
        return value["doubleValue"]
    if "timestampValue" in value:
        return value["timestampValue"]
    return ""


def _fs_error_detail(resp):
    """Best-effort message from a Google API error body.

    runQuery streams its result, so a rules denial arrives as a JSON *array*
    holding one status object rather than a dict. Assuming a dict there raised
    AttributeError and hid the actual reason, which is the one thing the user
    needs to see.
    """
    try:
        body = resp.json()
    except ValueError:
        return ""
    if isinstance(body, list):
        body = next((item for item in body if isinstance(item, dict)), None)
    if not isinstance(body, dict):
        return ""
    error = body.get("error")
    if isinstance(error, list):
        error = next((item for item in error if isinstance(item, dict)), None)
    if isinstance(error, dict):
        return error.get("message", "") or ""
    if isinstance(error, str):
        return error
    return ""


def _fs_denied_error(collection, resp):
    """Build the user-facing error for a 401/403 from Firestore.

    The bare "rules must allow" sentence cannot tell apart the three things
    that produce a 403 here: the rules really do deny the call, the ID token
    was rejected (so request.auth is null), or the Cloud Firestore API is not
    enabled for the project. Google's own message does, so it is appended.
    """
    detail = _fs_error_detail(resp)
    if resp.status_code == 401:
        message = ("Your Firebase session was rejected. Sign in again to get "
                   "a fresh token.")
    else:
        message = (
            f"The database refused the request for '{collection}'. The "
            f"Firestore rules must allow the signed-in user to use that "
            f"collection."
        )
    if detail and detail not in message:
        message = f"{message} ({detail})"
    return FirebaseError(message, status=resp.status_code)


def _fs_request(method, path, id_token, payload=None, allow_missing=False):
    """One Firestore REST call, authenticated as the signed-in user.

    With allow_missing=True a 404 returns None instead of raising: a first-time
    user legitimately has no users/{uid} document yet.
    """
    if not id_token:
        raise FirebaseError("Your session expired. Please sign in again.")
    headers = {"Authorization": f"Bearer {id_token}"}
    if payload is not None:
        headers["Content-Type"] = "application/json"
    try:
        resp = requests.request(
            method, f"{FIRESTORE_API}/{path}", headers=headers,
            json=payload, timeout=HTTP_TIMEOUT,
        )
    except requests.RequestException as e:
        raise FirebaseError(f"Could not reach the database ({e.__class__.__name__}).") from e

    if resp.status_code == 404 and allow_missing:
        return None
    if resp.status_code in (401, 403):
        logger.warning("Firestore %s %s -> %d %s", method, path,
                       resp.status_code, _fs_error_detail(resp))
        # Name the collection that was actually refused. Pointing at
        # users/{uid} every time sent people hunting in the wrong rules block
        # when the failing call was for a different collection.
        collection = path.split("?", 1)[0].split("/", 1)[0]
        raise _fs_denied_error(collection, resp)
    if resp.status_code >= 400:
        detail = _fs_error_detail(resp)
        logger.error("Firestore %s %s -> %d %s", method, path, resp.status_code, detail)
        raise FirebaseError(
            f"Database error {resp.status_code}: {detail or resp.text[:120]}",
            status=resp.status_code,
        )
    return resp


def _load_profile(uid, id_token):
    """Read users/{uid}. Returns a dict (empty when the document is missing)."""
    if not uid:
        raise FirebaseError("No signed-in user.")
    resp = _fs_request("GET", f"users/{uid}", id_token, allow_missing=True)
    if resp is None:
        return {}
    fields = (resp.json() or {}).get("fields") or {}
    return {k: _fs_decode(v) for k, v in fields.items()}


def _update_profile(uid, id_token, values):
    """Merge `values` into users/{uid}, creating the document when absent.

    A value of None deletes that field.
    """
    if not uid:
        raise FirebaseError("No signed-in user.")
    mask = "&".join(f"updateMask.fieldPaths={k}" for k in values)
    payload = {"fields": {k: _fs_value(v) for k, v in values.items()}}
    try:
        _fs_request("PATCH", f"users/{uid}?{mask}", id_token, payload)
    except FirebaseError as e:
        if e.status != 404:
            raise
        # A masked update needs the document to exist; for a brand-new account
        # create it instead (nulls have nothing to delete).
        create = {k: v for k, v in values.items() if v is not None}
        if not create:
            return
        _fs_request(
            "PATCH", f"users/{uid}", id_token,
            {"fields": {k: _fs_value(v) for k, v in create.items()}},
        )


def _refresh_id_token(user_data):
    """Firebase ID tokens live for an hour; renew with the refresh token.

    Note the snake_case keys: the Secure Token API returns {id_token,
    refresh_token}, unlike Identity Toolkit's camelCase.
    """
    refresh_token = user_data.get("refresh_token")
    if not refresh_token:
        return user_data.get("id_token", "")
    try:
        data = _auth_post(
            f"{SECURE_TOKEN_API}?key={FIREBASE_API_KEY}",
            {"grant_type": "refresh_token", "refresh_token": refresh_token},
        )
        user_data["id_token"] = data.get("id_token", user_data.get("id_token", ""))
        user_data["refresh_token"] = data.get("refresh_token", refresh_token)
        logger.info("Firebase session refreshed")
        return user_data["id_token"]
    except Exception:
        logger.exception("Could not refresh the Firebase session")
        return user_data.get("id_token", "")


def _update_user_profile(user_data, values):
    """Write profile fields for the signed-in user, retrying once after a
    token refresh so an expired session does not fail a settings save."""
    uid = user_data.get("uid", "")
    try:
        _update_profile(uid, user_data.get("id_token", ""), values)
    except FirebaseError:
        token = _refresh_id_token(user_data)
        if not token:
            raise
        _update_profile(uid, token, values)


def _fs_create(collection, id_token, values):
    """Create a document with a server-assigned id. Returns that id."""
    if not id_token:
        raise FirebaseError("Your session expired. Please sign in again.")
    resp = _fs_request(
        "POST", collection, id_token,
        {"fields": {k: _fs_value(v) for k, v in values.items()}},
    )
    name = (resp.json() or {}).get("name", "")
    return name.rsplit("/", 1)[-1]


def _fs_query(collection, id_token, field=None, value=None, limit=100):
    """Query a collection and return documents as flat dicts, id included.

    Deliberately filtered on one field with no order-by: Firestore serves those
    from its automatic single-field indexes, so this works on a fresh project
    without anyone having to build a composite index first. Ordering is done in
    Python - there are only a handful of rows.
    """
    if not id_token:
        raise FirebaseError("Your session expired. Please sign in again.")

    structured = {
        "from": [{"collectionId": collection}],
        "limit": limit,
    }
    if field:
        structured["where"] = {
            "fieldFilter": {
                "field": {"fieldPath": field},
                "op": "EQUAL",
                "value": _fs_value(value),
            }
        }

    try:
        resp = requests.post(
            f"{FIRESTORE_API}:runQuery",
            headers={
                "Authorization": f"Bearer {id_token}",
                "Content-Type": "application/json",
            },
            json={"structuredQuery": structured},
            timeout=HTTP_TIMEOUT,
        )
    except requests.RequestException as e:
        raise FirebaseError(
            f"Could not reach the database ({e.__class__.__name__})."
        ) from e

    if resp.status_code in (401, 403):
        # Same wording as _fs_request: name the collection that was refused so
        # nobody goes looking in the users/{uid} rules block for a schedule
        # permission problem.
        logger.warning("Firestore query on %s -> %d %s", collection,
                       resp.status_code, _fs_error_detail(resp))
        raise _fs_denied_error(collection, resp)
    if resp.status_code >= 400:
        detail = _fs_error_detail(resp)
        logger.error("Firestore query on %s -> %d %s", collection, resp.status_code, detail)
        raise FirebaseError(
            f"Database error {resp.status_code}: {detail or resp.text[:120]}",
            status=resp.status_code,
        )

    documents = []
    for row in resp.json() or []:
        doc = row.get("document")
        if not doc:
            continue    # runQuery sends a bare readTime for rows with no document
        item = {k: _fs_decode(v) for k, v in (doc.get("fields") or {}).items()}
        item["id"] = (doc.get("name") or "").rsplit("/", 1)[-1]
        documents.append(item)
    return documents


def _fs_delete(collection, doc_id, id_token):
    """Delete one document."""
    if not id_token:
        raise FirebaseError("Your session expired. Please sign in again.")
    _fs_request("DELETE", f"{collection}/{doc_id}", id_token)


# ================= LOCAL SECRET STORAGE =================
# The TFS PAT is a credential. It is encrypted with DPAPI for the current
# Windows user and stored under %LOCALAPPDATA%, which keeps it out of reach of
# other Windows users on the same PC and of anything reading the database.
#
# PAT_SYNC_TO_FIRESTORE mirrors it into this user's own document so the same
# account works on any PC without re-entering the token. Trade-off: the PAT is
# then stored in Firestore in plain text and is only as private as the
# users/{uid} rules (owner-only) and the project's access controls. Set this
# to False to return to per-machine-only storage.
PAT_SYNC_TO_FIRESTORE = True

# A scheduled comment is posted later by the hosted worker, which needs the
# author's PAT at post time. Keeping the token only in users/{uid} makes that
# depend on the worker being allowed to read every user document - a rules
# edit that fails silently, and when it fails the comment is posted with the
# worker's own TFS_PAT and credited to that person instead.
#
# With this on, the token is also written into the schedule document itself
# (scheduledComments/{id}.pat), which the worker already has read access to,
# so a comment is credited to the user who scheduled it whatever the rules for
# users/{uid} say. The cost is one more copy of the secret in Firestore, so
# reads of scheduledComments must stay limited to the owner and the worker
# (see the rules in the README).
COPY_PAT_INTO_SCHEDULE = True

def _credential_dir():
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    path = os.path.join(base, "AutoReport", "credentials")
    os.makedirs(path, exist_ok=True)
    return path


class _DataBlob(ctypes.Structure):
    _fields_ = [
        ("cbData", ctypes.wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


def _dpapi(data, protect=True):
    """Encrypt (protect=True) or decrypt a blob for the current Windows user."""
    if not data:
        return b""
    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    buf = ctypes.create_string_buffer(data, len(data))
    src = _DataBlob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    dst = _DataBlob()
    call = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    if not call(ctypes.byref(src), None, None, None, None, 0, ctypes.byref(dst)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(dst.pbData, dst.cbData)
    finally:
        kernel32.LocalFree(dst.pbData)


def _pat_file(uid):
    return os.path.join(_credential_dir(), f"{uid}.bin")


def _save_pat_local(uid, pat):
    """Store this user's PAT encrypted for this Windows account."""
    if not uid or not pat:
        return False
    try:
        with open(_pat_file(uid), "wb") as f:
            f.write(_dpapi(pat.encode("utf-8"), protect=True))
        logger.info("PAT stored in the local credential store for uid=%s", uid)
        return True
    except Exception:
        logger.exception("Could not store the PAT locally")
        return False


def _save_pat_remote(uid, id_token, pat):
    """Mirror the PAT into users/{uid} so any PC can sign in with it.

    Best effort: a failure here (offline, rules rejecting the write) leaves the
    local DPAPI copy in charge, so login still works on this machine.
    """
    if not PAT_SYNC_TO_FIRESTORE:
        return False
    if not (uid and id_token and pat):
        return False
    try:
        _update_profile(uid, id_token, {"pat": pat})
        logger.info("PAT synced to Firestore for uid=%s", uid)
        return True
    except FirebaseError as e:
        logger.warning("Could not sync the PAT to Firestore: %s", e)
        return False


def _load_pat_local(uid):
    if not uid:
        return ""
    try:
        path = _pat_file(uid)
        if not os.path.isfile(path):
            return ""
        with open(path, "rb") as f:
            return _dpapi(f.read(), protect=False).decode("utf-8")
    except Exception:
        logger.exception("Could not read the stored PAT - it will have to be re-entered")
        return ""


def _delete_pat_local(uid):
    try:
        if uid:
            os.remove(_pat_file(uid))
    except OSError:
        pass



def replace_placeholders_in_paragraph(paragraph, replacements: dict):
    """
    Robustly replace {{PLACEHOLDER}} tokens in a paragraph, even when
    Word has split the token across multiple runs.
    """
    full_text = paragraph.text
    if not any(ph in full_text for ph in replacements):
        return  # nothing to do, skip untouched paragraphs (preserves formatting)

    logger.debug("Paragraph contains placeholder(s): %r", full_text[:120])

    # Apply all replacements to the combined text
    new_text = full_text
    for ph, value in replacements.items():
        new_text = new_text.replace(ph, value)

    if new_text == full_text:
        return

    # Wipe all runs except the first, then set the first run's text.
    # This collapses formatting to whatever the first run had, which is
    # usually fine for placeholder tokens (they're typically uniformly formatted).
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_text)


def replace_placeholders_in_doc(doc, replacements: dict):
    logger.info("Replacing placeholders in document (%d replacements)", len(replacements))
    logger.debug("Placeholder keys: %s", list(replacements.keys()))

    # Paragraphs outside tables
    for para in doc.paragraphs:
        replace_placeholders_in_paragraph(para, replacements)

    # Paragraphs inside tables (including nested tables)
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                for para in cell.paragraphs:
                    replace_placeholders_in_paragraph(para, replacements)
                # handle nested tables inside a cell
                for nested_table in cell.tables:
                    for nrow in nested_table.rows:
                        for ncell in nrow.cells:
                            for npara in ncell.paragraphs:
                                replace_placeholders_in_paragraph(npara, replacements)

    # --- Text boxes (w:txbxContent) ---
    # python-docx does NOT expose paragraphs inside text boxes via
    # doc.paragraphs/doc.tables, since they live inside drawing/VML
    # shapes (mc:AlternateContent -> mc:Choice/w:drawing and
    # mc:Fallback -> w:pict). Both copies must be patched identically,
    # so we walk the raw XML tree directly.
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph

    body = doc.element.body
    txbx_count = 0
    for txbx_content in body.iter(qn('w:txbxContent')):
        txbx_count += 1
        for p in txbx_content.iter(qn('w:p')):
            wrapped_para = Paragraph(p, None)
            replace_placeholders_in_paragraph(wrapped_para, replacements)
    logger.debug("Processed %d text box(es) in document", txbx_count)


def insert_html_into_placeholder(doc, placeholder, raw_html, pat=""):
    """
    Replace `placeholder` with rendered HTML (images included) wherever it
    lives: body paragraphs, table cells (nested tables too) or text boxes.
    Returns True when the placeholder was found and replaced.
    """
    logger.info("Looking for placeholder %r in document", placeholder)
    found = False

    def try_paragraph(paragraph):
        nonlocal found
        if placeholder in paragraph.text:
            logger.debug("Found %r in paragraph text", placeholder)
            # Remove the placeholder runs instead of blanking them: an emptied
            # run stayed at the front of the paragraph with its own formatting.
            for run in list(paragraph.runs):
                run._r.getparent().remove(run._r)
            add_html_content_to_paragraph(paragraph, raw_html, pat)
            found = True
            return True
        return False

    # Body paragraphs
    for para in doc.paragraphs:
        if try_paragraph(para):
            return True

    # Table cells, including tables nested inside a cell
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if try_paragraph(para):
                        return True
                for nested_table in cell.tables:
                    for nrow in nested_table.rows:
                        for ncell in nrow.cells:
                            for npara in ncell.paragraphs:
                                if try_paragraph(npara):
                                    return True

    # Text boxes live inside drawing/VML shapes, which doc.paragraphs and
    # doc.tables never expose, so walk the raw XML tree.
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph

    for txbx_content in doc.element.body.iter(qn('w:txbxContent')):
        for p in txbx_content.iter(qn('w:p')):
            if try_paragraph(Paragraph(p, None)):
                return True

    logger.warning(
        "Placeholder %r was NOT FOUND in any paragraph, table cell, or text box!",
        placeholder,
    )
    return False


# --- HTML to DOCX rendering ---
# Tags that end a block of text, and tags whose contents are not content.
_BLOCK_TAGS = {
    "p", "div", "li", "ul", "ol", "tr", "table", "blockquote",
    "h1", "h2", "h3", "h4", "h5", "h6",
}
_SKIP_TAGS = {"script", "style", "head"}

# One download per URL per session: the same screenshot is usually pasted into
# several comments, and it used to be re-fetched for every report.
_image_cache = {}


def _fetch_image_bytes(url, pat=""):
    """Download an image, resolving TFS-relative src values.

    The PAT is only ever sent to the TFS host itself - the old code attached it
    to whatever host an <img src> pointed at in the comment.
    """
    candidate = (url or "").strip()
    if candidate.startswith("//"):
        candidate = "https:" + candidate
    elif candidate.startswith("/"):
        candidate = _TFS_HOST + candidate

    if candidate in _image_cache:
        return _image_cache[candidate]

    if candidate.startswith(_TFS_HOST):
        resp = requests.get(candidate, auth=HTTPBasicAuth("", pat), timeout=IMAGE_TIMEOUT)
    else:
        resp = requests.get(candidate, timeout=IMAGE_TIMEOUT)
    resp.raise_for_status()

    if len(_image_cache) > 50:
        _image_cache.clear()
    _image_cache[candidate] = resp.content
    return resp.content


def _utc_now_iso():
    """UTC timestamp in the fixed format the worker compares against.

    Fixed width on purpose: plain string comparison is then chronological,
    which is how the worker decides what is due without a Firestore range
    index.
    """
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# Bare URLs in a comment become real links. Stopping at whitespace and quotes
# keeps a URL that ends a sentence from swallowing the full stop - trailing
# punctuation is put back outside the anchor.
_URL_PATTERN = re.compile(r"(https?://[^\s<>\"']+)", re.IGNORECASE)


def _escape_comment_text(text):
    return (
        (text or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _linkify_escaped(escaped):
    """Wrap bare URLs in already-escaped text with anchors."""
    def _link(match):
        url = match.group(1)
        trailing = ""
        while url and url[-1] in ".,;:!?":
            trailing = url[-1] + trailing
            url = url[:-1]
        return f'<a href="{url}">{url}</a>{trailing}'

    return _URL_PATTERN.sub(_link, escaped)


# Tags a comment may contain after the rich editor's HTML is normalised.
# Anything else is unwrapped to its text, or dropped along with its content.
_COMMENT_INLINE_TAGS = ("b", "strong", "i", "em", "u", "s", "strike", "code")
_COMMENT_URL_SCHEME = re.compile(r"^(?:https?|mailto):", re.IGNORECASE)
_COMMENT_IMAGE_URL = re.compile(r"^https?://", re.IGNORECASE)


def _qt_span_tags(style):
    """Qt writes bold/italic/underline as a styled <span>; name the tags.

    Qt emits `font-weight:600`, `font-style:italic` and `text-decoration:
    underline` with a space after the colon, so each is matched loosely.
    """
    opens, closes = "", ""
    weight = re.search(r"font-weight:\s*(bold|\d+)", style)
    if weight:
        value = weight.group(1)
        if value == "bold" or int(value) >= 600:
            opens += "<b>"
            closes = "</b>" + closes
    if re.search(r"font-style:\s*italic", style):
        opens += "<i>"
        closes = "</i>" + closes
    if re.search(r"text-decoration:\s*[^;]*underline", style):
        opens += "<u>"
        closes = "</u>" + closes
    if re.search(r"text-decoration:\s*[^;]*line-through", style):
        opens += "<s>"
        closes = "</s>" + closes
    return opens, closes


def _rich_comment_html(document_html, image_urls=None, placed=None):
    """Turn the editor's rich document into the HTML TFS stores.

    Only the tags listed above (plus paragraphs, lists and plain hyperlinks)
    survive, so a paste from a browser or Word cannot smuggle in styles,
    scripts or other references. Text is escaped and bare URLs are linked, so
    the result is safe whatever was pasted.

    A picture pasted into the box lives on this PC until it is uploaded as a
    work item attachment. `image_urls` maps the name Qt gives that picture (its
    temp file name) to the attachment URL, so the picture is written where it
    was in the comment rather than collected at the end; every name it uses is
    added to `placed`, so _comment_html does not append it a second time. An
    http(s) <img> that already points at a server is kept as it is, which is
    how a picture copied out of another ticket stays visible.
    """
    image_urls = image_urls or {}
    placed = placed if placed is not None else set()
    soup = BeautifulSoup(document_html or "", "html.parser")
    root = soup.body or soup
    out = []
    last_break = [True]

    def br():
        if not last_break[0]:
            out.append("<br>")
            last_break[0] = True

    def walk(node):
        for child in node.children:
            if isinstance(child, NavigableString):
                text = str(child)
                if text.strip():
                    out.append(_linkify_escaped(_escape_comment_text(text)))
                    last_break[0] = False
                continue
            if not isinstance(child, Tag):
                continue
            name = child.name.lower()
            if name in ("script", "style", "head", "title", "meta"):
                continue
            if name in ("p", "div", "blockquote"):
                br()
                walk(child)
                br()
                continue
            if name == "br":
                # Always emitted: an empty paragraph is how the editor stores
                # a blank line, and dropping it would join the paragraphs.
                out.append("<br>")
                last_break[0] = True
                continue
            if name == "span":
                opens, closes = _qt_span_tags((child.get("style") or "").lower())
                if opens:
                    out.append(opens)
                    walk(child)
                    out.append(closes)
                else:
                    walk(child)
                continue
            if name == "a":
                href = (child.get("href") or "").strip()
                if _COMMENT_URL_SCHEME.match(href):
                    out.append(f'<a href="{_escape_comment_text(href)}">')
                    walk(child)
                    out.append("</a>")
                else:
                    walk(child)
                continue
            if name == "img":
                src = (child.get("src") or "").strip()
                if _COMMENT_IMAGE_URL.match(src):
                    out.append(f'<img src="{_escape_comment_text(src)}">')
                    last_break[0] = False
                    continue
                # A picture pasted into the box is named after its file on
                # disk; that name is the key its upload is filed under, so it
                # can be written here instead of being gathered at the end.
                key = os.path.basename(src)
                uploaded = image_urls.get(key) or image_urls.get(src) or ""
                if uploaded:
                    out.append(f'<img src="{_escape_comment_text(uploaded)}">')
                    placed.add(key)
                    last_break[0] = False
                continue
            if name in ("ul", "ol"):
                br()
                out.append(f"<{name}>")
                walk(child)
                out.append(f"</{name}>")
                br()
                continue
            if name == "li":
                out.append("<li>")
                last_break[0] = True      # no break at the very start of an item
                walk(child)
                out.append("</li>")
                last_break[0] = False
                continue
            if name in _COMMENT_INLINE_TAGS:
                out.append(f"<{name}>")
                walk(child)
                out.append(f"</{name}>")
                continue
            walk(child)                    # unknown tag: keep only its text

    walk(root)
    result = "".join(out).strip()
    result = re.sub(r"(?:<br>)+(?=</li>)", "", result)
    result = re.sub(r"^(?:<br>)+", "", result)
    return re.sub(r"(?:<br>)+$", "", result)


def _comment_html(text, image_urls=None, rich_html=""):
    """Turn the compose box into the HTML TFS stores.

    `rich_html` is what the comment editor produces (paragraphs, lists,
    bold/italic/underline, links, pictures) and goes through
    _rich_comment_html, which keeps a fixed set of tags and nothing else.
    Without it the plain text is escaped first, so a comment can never inject
    markup: only the links wrapped below and the attachment images are real
    tags.

    `image_urls` maps a picture's file name to the attachment URL it was
    uploaded to, so a pasted picture is written back into the comment at the
    place it was typed. One that has no place left in the text - a picture
    deleted from the box after it was attached - is appended at the end.
    """
    keyed = isinstance(image_urls, dict)
    uploads = image_urls if keyed else {}
    unplaced = [] if keyed else [url for url in (image_urls or []) if url]
    placed = set()

    if rich_html:
        body = _rich_comment_html(rich_html, uploads, placed)
    else:
        escaped = _escape_comment_text(text)
        body = "<br>".join(
            re.split(r"\r\n|\r|\n", _linkify_escaped(escaped))
        )

    unplaced += [url for name, url in uploads.items()
                 if url and name not in placed]
    if unplaced:
        body += "".join(f'<br><img src="{url}">' for url in unplaced)
    return body.strip()


def _upload_comment_image(pat, work_item_id, file_path):
    """Upload an image as a work item attachment and return its URL.

    Images are attached at compose time on purpose: the worker on the host
    cannot read this PC's disk, and a TFS attachment URL keeps working no
    matter when the comment is finally posted.
    """
    from urllib.parse import quote

    name = os.path.basename(file_path)
    with open(file_path, "rb") as handle:
        payload = handle.read()

    url = f"{org_url}/_apis/wit/attachments?fileName={quote(name)}&api-version=5.1"
    logger.info("Uploading image for work item %s: %s (%d bytes)",
                work_item_id, name, len(payload))
    resp = requests.post(
        url,
        auth=HTTPBasicAuth("", pat),
        data=payload,
        headers={"Content-Type": "application/octet-stream"},
        timeout=DOWNLOAD_TIMEOUT,
    )
    if resp.status_code >= 400:
        raise RuntimeError(
            f"TFS rejected {name} ({resp.status_code}): {resp.text[:160]}"
        )
    return (resp.json() or {}).get("url", "")


def add_html_content_to_paragraph(paragraph, html_content, pat=""):
    """Render HTML (text, breaks, lists, bold/italic, images) into a paragraph.

    The old renderer added a break after *every* text node and another one per
    <p>, so each line of a comment gained a blank line under it; it also
    flattened the text of plain-text comments into one long line.
    """
    logger.debug("Rendering HTML content into paragraph (length=%d chars)", len(html_content or ""))
    soup = BeautifulSoup(html_content or "", "html.parser")
    img_count = 0
    img_fail_count = 0
    last_was_break = True   # never start with a break

    def add_break():
        nonlocal last_was_break
        if last_was_break:
            return
        paragraph.add_run().add_break()
        last_was_break = True

    def add_text(text, bold=False, italic=False):
        nonlocal last_was_break
        for index, line in enumerate(re.split(r"\r\n|\r|\n", str(text))):
            if index:
                add_break()
            collapsed = re.sub(r"[ \t]+", " ", line)
            if last_was_break:
                collapsed = collapsed.lstrip()
            if not collapsed.strip():
                continue
            run = paragraph.add_run(collapsed)
            if bold:
                run.bold = True
            if italic:
                run.italic = True
            last_was_break = False

    def add_image(url):
        nonlocal img_count, img_fail_count, last_was_break
        try:
            logger.debug("Fetching image: %s", (url or "")[:120])
            stream = BytesIO(_fetch_image_bytes(url, pat))
            with Image.open(stream) as probe:
                dpi = 96.0
                info = probe.info.get("dpi")
                if isinstance(info, (tuple, list)) and info and info[0]:
                    dpi = float(info[0]) or 96.0
                width_inches = probe.width / dpi
            stream.seek(0)
            if width_inches > 6:
                paragraph.add_run().add_picture(stream, width=Inches(6))
            else:
                paragraph.add_run().add_picture(stream)
            last_was_break = False
            img_count += 1
        except Exception as e:
            img_fail_count += 1
            logger.error("Error inserting image from %s: %s", (url or "")[:120], e)

    def walk(node, bold=False, italic=False):
        if isinstance(node, NavigableString):
            add_text(str(node), bold, italic)
            return
        if not isinstance(node, Tag) or node.name in _SKIP_TAGS:
            return

        name = node.name
        if name == "br":
            add_break()
            return
        if name == "img":
            add_image(node.get("src"))
            return

        if name in ("strong", "b"):
            bold = True
        elif name in ("em", "i"):
            italic = True
        elif name == "li":
            add_break()
            add_text("• ")

        for child in node.children:
            walk(child, bold, italic)

        if name in _BLOCK_TAGS:
            add_break()

    for child in soup.children:
        walk(child)

    logger.info("HTML rendering complete: %d image(s) inserted, %d failed",
                img_count, img_fail_count)


def _ui_alive(widget):
    """True while a widget still exists.

    A worker signal can arrive after the page that connected it has been torn
    down by clear_layout(), and writing to the dead C++ object raises
    "wrapped C/C++ object of type QTextEdit has been deleted".
    """
    if widget is None:
        return False
    try:
        return not sip.isdeleted(widget)
    except Exception:
        return True


# ================= CUSTOM DIALOG =================
class AppDialog(QtWidgets.QDialog):
    def __init__(self, title, message, parent=None):
        super().__init__(parent)

        self.setWindowIcon(QtGui.QIcon(os.path.join(ASSET_DIR, "death.ico")))
        self.setWindowTitle(title)
        # Minimum, not fixed: the longer error messages (Firestore rules,
        # "cannot create report folder") were clipped by a fixed 160px height.
        self.setMinimumSize(320, 160)
        self.setObjectName("Root")

        # Use the same dark/light styles as your main window
        if is_windows_dark_mode():
            self.setStyleSheet("""
            QDialog#Root {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #07101a, stop:1 #07121a);
                font-family: Inter, Segoe UI, Arial;
                color: #e6eef5;
            }
            QLabel {
                color: #e6eef5;
                background: transparent;
            }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white;
                border-radius: 10px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                opacity: 0.92;
            }
            """)
        else:
            self.setStyleSheet("""
            QDialog#Root {
                background: #f5f7fa;
                font-family: Inter, Segoe UI, Arial;
                color: #1f2933;
            }
            QLabel {
                color: #111827;
                background: transparent;
            }
            QPushButton {
                background: #ed3b3b;
                color: white;
                border-radius: 10px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background: #c62e2e;
            }
            """)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)
        label = QtWidgets.QLabel(message)
        label.setWordWrap(True)
        button = QtWidgets.QPushButton("OK")
        button.setFixedHeight(34)
        button.clicked.connect(self.accept)
        layout.addWidget(label)
        layout.addStretch()
        layout.addWidget(button)
        self.adjustSize()


class UpdateCheckWorker(QThread):
    """
    Runs the GitHub release check off the GUI thread, so the Settings button
    cannot freeze the window while the request is in flight.
    """
    found = pyqtSignal(object)   # update dict, or None when up to date

    def run(self):
        try:
            self.found.emit(_check_for_update())
        except Exception:
            logger.exception("Manual update check failed")
            self.found.emit(None)


class LoginWorker(QThread):
    """Signs in and loads the profile off the GUI thread.

    Three network round-trips (Firebase auth, Firestore profile, TFS PAT expiry)
    ran inline in login(), so the window stopped repainting and Windows drew it
    as "Not responding" for the whole sign-in.
    """
    succeeded = pyqtSignal(object)      # the current_user dict
    pat_required = pyqtSignal(object)   # profile without a stored PAT
    failed = pyqtSignal(str)            # message, or a known marker

    def __init__(self, email, password, parent=None):
        super().__init__(parent)
        self.email = email
        self.password = password

    def run(self):
        try:
            session = _sign_in(self.email, self.password)
            uid = session["uid"]
            id_token = session["id_token"]
            logger.info("Firebase auth success, uid=%s", uid)

            profile = _load_profile(uid, id_token)
            if not profile:
                self.failed.emit("USER_PROFILE_NOT_FOUND")
                return

            data = {
                "email": self.email,
                "uid": uid,
                "pat": "",
                "display_name": profile.get("display_name", ""),
                "unique_name": profile.get("unique_name", ""),
                "pat_expiry": profile.get("pat_expiry", ""),
                "id_token": id_token,
                "refresh_token": session["refresh_token"],
            }

            # Local store first: it is fast and works without Firestore. With
            # PAT_SYNC_TO_FIRESTORE on, users/{uid}.pat is the copy that lets
            # the same account sign in on another PC.
            pat = _load_pat_local(uid)
            remote_pat = (profile.get("pat") or "").strip()
            if not pat and remote_pat:
                pat = remote_pat
                if _save_pat_local(uid, pat):
                    logger.info("PAT restored from Firestore for uid=%s", uid)
            elif pat and pat != remote_pat:
                # The local copy is authoritative: publish it when the remote
                # one is empty (first login after upgrading) and refresh it
                # when the token was regenerated. The worker posts comments
                # with users/{uid}.pat, so a stale copy there would make every
                # schedule from this account fail with a 401.
                if _save_pat_remote(uid, id_token, pat) and remote_pat:
                    logger.info("Refreshed the Firestore PAT for uid=%s", uid)

            if not pat:
                logger.warning("No PAT available for uid=%s", uid)
                self.pat_required.emit(data)
                return

            data["pat"] = pat

            if not data["pat_expiry"]:
                expiry = _fetch_pat_expiry(pat)
                if expiry:
                    data["pat_expiry"] = expiry
                    try:
                        _update_profile(uid, id_token, {"pat_expiry": expiry})
                        logger.info("PAT expiry saved: %s", expiry)
                    except FirebaseError:
                        logger.warning("Could not save the PAT expiry")

            self.succeeded.emit(data)
        except FirebaseAuthError as e:
            # The code alone ("INVALID_LOGIN_CREDENTIALS", "NETWORK", ...) is
            # what the dialog maps to a readable sentence.
            logger.warning("Sign-in failed for %r: %s", self.email, e.detail)
            self.failed.emit(e.code)
        except Exception as e:
            logger.exception("Login failed for %r", self.email)
            self.failed.emit(str(e))


class UpdateDownloadWorker(QThread):
    """Downloads and verifies the release asset off the GUI thread.

    The download used to run inside UpdateDialog with processEvents() pumping
    the progress bar. That let the user hit Skip mid-transfer, after which the
    progress callback was still writing into the destroyed dialog.
    """
    progress = pyqtSignal(int, int)   # bytes done, bytes total
    downloaded = pyqtSignal(str)      # verified local path
    failed = pyqtSignal(str)

    def __init__(self, url, expected_size=None, expected_digest=None, parent=None,
                 local_name=None):
        super().__init__(parent)
        self.url = url
        self.expected_size = expected_size
        self.expected_digest = expected_digest
        self.local_name = local_name

    def run(self):
        try:
            path = _download_update(
                self.url,
                progress_callback=lambda done, total: self.progress.emit(done, total),
                expected_size=self.expected_size,
                expected_digest=self.expected_digest,
                local_name=self.local_name,
            )
        except Exception:
            logger.exception("Update download crashed")
            path = None
        if path:
            self.downloaded.emit(path)
        else:
            self.failed.emit(
                "Download failed after 3 attempts. Nothing was installed; "
                "the details are in report_log.txt."
            )


class UpdateDialog(QtWidgets.QDialog):
    """Dialog shown when an update is available."""
    def __init__(self, version, release_notes, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Update Available")
        self.setFixedSize(400, 280)
        self.setObjectName("UpdateDialog")
        self.download_url = None

        dark = is_windows_dark_mode()
        if dark:
            self.setStyleSheet("""
            QDialog#UpdateDialog {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #07101a, stop:1 #07121a);
                font-family: Inter, Segoe UI, Arial;
                color: #e6eef5;
            }
            QLabel { color: #e6eef5; background: transparent; }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white; border-radius: 10px; padding: 8px 16px;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton#SkipBtn {
                background: transparent; border: 1px solid rgba(255,255,255,0.15);
                color: #98a0a6;
            }
            QPushButton#SkipBtn:hover { border: 1px solid #98a0a6; color: #e6eef5; }
            QProgressBar {
                background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1);
                border-radius: 6px; text-align: center; color: #e6eef5; height: 20px;
            }
            QProgressBar::chunk { background: #ed3b3b; border-radius: 5px; }
            """)
        else:
            self.setStyleSheet("""
            QDialog#UpdateDialog {
                background: #f5f7fa;
                font-family: Inter, Segoe UI, Arial;
                color: #1f2933;
            }
            QLabel { color: #111827; background: transparent; }
            QPushButton {
                background: #ed3b3b; color: white; border-radius: 10px; padding: 8px 16px;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton#SkipBtn {
                background: transparent; border: 1px solid #d1d5db; color: #6b7280;
            }
            QPushButton#SkipBtn:hover { border: 1px solid #6b7280; color: #111827; }
            QProgressBar {
                background: #e5e7eb; border: 1px solid #d1d5db;
                border-radius: 6px; text-align: center; color: #111827; height: 20px;
            }
            QProgressBar::chunk { background: #ed3b3b; border-radius: 5px; }
            """)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        title = QtWidgets.QLabel(f"Version {version} is available")
        title.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(title)

        current_label = QtWidgets.QLabel(f"Current version: {APP_VERSION}")
        current_label.setStyleSheet("font-size: 12px; color: #6b7280;")
        layout.addWidget(current_label)

        if release_notes:
            notes_label = QtWidgets.QLabel(release_notes[:200])
            notes_label.setWordWrap(True)
            notes_label.setStyleSheet("font-size: 12px; color: #98a0a6;")
            layout.addWidget(notes_label)

        layout.addSpacing(8)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.status_label = QtWidgets.QLabel("")
        self.status_label.setStyleSheet("font-size: 11px; color: #6b7280;")
        self.status_label.setVisible(False)
        layout.addWidget(self.status_label)

        layout.addStretch()

        btn_layout = QtWidgets.QHBoxLayout()
        self.download_btn = QtWidgets.QPushButton("Download Update")
        self.download_btn.setFixedHeight(38)
        self.download_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.download_btn.clicked.connect(self._on_download)

        skip_btn = QtWidgets.QPushButton("Skip")
        skip_btn.setObjectName("SkipBtn")
        skip_btn.setFixedHeight(38)
        skip_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        skip_btn.clicked.connect(self.reject)

        btn_layout.addWidget(skip_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.download_btn)
        layout.addLayout(btn_layout)

    def _on_download(self):
        update_info = getattr(self.parent(), "_update_info", None) if self.parent() else None
        if not update_info:
            self.status_label.setVisible(True)
            self.status_label.setText("Error: No update info")
            return

        asset_type = update_info.get("asset_type", "exe")
        if asset_type != "zip" and not update_info.get("digest"):
            # An .exe is executed, not inspected: without the SHA-256 the
            # release can publish there is nothing to check but its length.
            self.status_label.setVisible(True)
            self.status_label.setText(
                "This release publishes no SHA-256 digest for the installer, so "
                "it cannot be verified. Ask for a .zip release."
            )
            return

        self.download_btn.setEnabled(False)
        self.download_btn.setText("Downloading...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setVisible(True)
        self.status_label.setText("Starting download...")

        self._download_worker = UpdateDownloadWorker(
            update_info["download_url"],
            expected_size=update_info.get("size"),
            expected_digest=update_info.get("digest"),
            # An .exe is saved under a name that does not end in .exe; the swap
            # script renames it to Report.exe after this process exits.
            local_name=(_staged_update_name(update_info.get("version", ""))
                        if asset_type != "zip" else None),
            parent=self,
        )
        self._download_worker.progress.connect(self._on_download_progress)
        self._download_worker.downloaded.connect(
            lambda path: self._install(path, update_info)
        )
        self._download_worker.failed.connect(self._on_download_failed)
        self._download_worker.start()

    def _on_download_progress(self, done, total):
        if not _ui_alive(self.progress_bar):
            return
        if total:
            self.progress_bar.setValue(int(done / total * 100))
            self.status_label.setText(
                f"{done / (1024 * 1024):.1f} MB / {total / (1024 * 1024):.1f} MB"
            )
        else:
            self.status_label.setText(f"{done / (1024 * 1024):.1f} MB")

    def _on_download_failed(self, message):
        if not _ui_alive(self.download_btn):
            return
        self.status_label.setText(message)
        self.download_btn.setEnabled(True)
        self.download_btn.setText("Retry")

    def closeEvent(self, event):
        """Refuse to close mid-download: the worker writes into this dialog."""
        worker = getattr(self, "_download_worker", None)
        if worker is not None and worker.isRunning():
            self.status_label.setText("Please wait - the download is still running.")
            event.ignore()
            return
        super().closeEvent(event)

    def reject(self):
        """Esc must not tear the dialog down while the download is in flight."""
        worker = getattr(self, "_download_worker", None)
        if worker is not None and worker.isRunning():
            self.status_label.setText("Please wait - the download is still running.")
            return
        super().reject()

    def _install(self, path, update_info):
        """Apply a verified download. Runs on the GUI thread, after the
        transfer has finished."""
        if not _ui_alive(self.download_btn):
            return

        self.progress_bar.setValue(100)

        if update_info.get("asset_type", "exe") == "zip":
            # Extract and replace files in-place
            self.status_label.setText("Verifying and installing update...")
            self.download_btn.setText("Applying update...")
            QtWidgets.QApplication.processEvents()

            ok, message = _apply_zip_update(path, update_info.get("version", ""))
            if ok:
                self.status_label.setText("Update applied! Restarting app...")
                QtWidgets.QApplication.processEvents()
                logger.info("Zip update applied, restarting app")
                self._restart_app()
            else:
                # Say what actually went wrong: the old build is still intact.
                self.status_label.setText(message or "Failed to apply update.")
                self._turn_download_button_into_close()
        else:
            # .exe - park it under a name that does not end in .exe, then let
            # the swap script rename it over the running exe once this process
            # has exited. The file is never launched under its download name.
            self.status_label.setText("Verifying and installing update...")
            self.download_btn.setText("Applying update...")
            QtWidgets.QApplication.processEvents()
            logger.info("Applying exe update: %s", path)

            ok, message = _apply_exe_update(path, update_info.get("version", ""))
            if ok:
                self.status_label.setText("Update applied! Restarting app...")
                QtWidgets.QApplication.processEvents()
                logger.info("Exe update staged, restarting app")
                self._restart_app()
            else:
                self.status_label.setText(message or "Failed to apply update.")
                self._turn_download_button_into_close()

    def _turn_download_button_into_close(self):
        """Swap the button's action for `accept`.

        Connecting accept() without disconnecting _on_download left both
        handlers on the same click, so "Close" started a second download.
        """
        if not _ui_alive(self.download_btn):
            return
        self.download_btn.setText("Close")
        self.download_btn.setEnabled(True)
        try:
            self.download_btn.clicked.disconnect()
        except TypeError:
            pass
        self.download_btn.clicked.connect(self.accept)

    def _restart_app(self):
        """Restart the application after a zip update."""
        import subprocess

        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
            exe_name = os.path.basename(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
            exe_name = None

        bat_path = os.path.join(app_dir, "_update_swap.bat")

        # Write a relaunch-only script if the update did not create one.
        # Starting the new copy while this process still holds the
        # single-instance mutex would make it refuse to start and leave the user
        # with no window at all - the script waits for this PID to disappear
        # first. This applies to source runs too, where exe_name is None and the
        # direct Popen below used to race the mutex.
        if not os.path.isfile(bat_path):
            try:
                _write_swap_script(bat_path, app_dir, exe_name, [])
            except Exception:
                logger.exception("Could not write the relaunch script")

        # The bat waits for this process to exit, then starts the new copy
        if os.path.isfile(bat_path):
            logger.info("Swap script found, launching: %s", bat_path)
            subprocess.Popen(["cmd", "/c", bat_path], shell=False, cwd=app_dir,
                             creationflags=subprocess.CREATE_NO_WINDOW)
            QtWidgets.QApplication.quit()
            return

        # No locked files — restart directly
        exe = sys.executable
        logger.info("Restarting: %s %s", exe, sys.argv)
        try:
            subprocess.Popen([exe] + sys.argv)
            QtWidgets.QApplication.quit()
        except Exception:
            logger.exception("Failed to restart app")
            self.status_label.setText("Update applied. Please restart manually.")
            self.download_btn.setText("Close")
            self.download_btn.setEnabled(True)
            self.download_btn.clicked.connect(self.accept)


class PBIFormDialog(QDialog):
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowTitle("Create Azure DevOps PBI")
            self.setFixedWidth(380)

            if is_windows_dark_mode():
                self.setStyleSheet("""
                QDialog { background: #0f1923; color: #e6eef5; font-family: Inter, Segoe UI, Arial; }
                QLabel { color: #e6eef5; }
                QLineEdit {
                    background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.06);
                    border-radius: 10px; padding: 12px; color: #e6eef5; font-size: 14px;
                }
                QLineEdit:focus { border: 1px solid #ed3b3b; }
                QComboBox {
                    background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.06);
                    border-radius: 10px; padding: 12px 36px 12px 12px; color: #e6eef5;
                    font-size: 14px; min-height: 20px;
                }
                QComboBox:focus { border: 1px solid #ed3b3b; }
                QComboBox::drop-down {
                    subcontrol-origin: padding; subcontrol-position: center right;
                    width: 32px; border: none;
                    border-left: 1px solid rgba(255,255,255,0.06);
                    border-top-right-radius: 10px; border-bottom-right-radius: 10px;
                }
                QComboBox::down-arrow { width: 10px; height: 10px; }
                QComboBox QAbstractItemView {
                    background: #15212c; border: 1px solid #2a3a48; border-radius: 8px;
                    padding: 4px; color: #e6eef5; selection-background-color: #1e3a50;
                    selection-color: #ffffff; outline: none;
                }
                QComboBox QAbstractItemView::item {
                    padding: 8px 12px; border-radius: 4px; min-height: 24px;
                }
                QComboBox QAbstractItemView::item:hover { background: #1e3a50; color: #ffffff; }
                QPushButton {
                    background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ed3b3b, stop:1 #c62e2e);
                    color: white; border-radius: 10px; padding: 12px 24px; font-size: 14px; font-weight: 600;
                }
                QPushButton:hover { background: #c62e2e; }
                """)
            else:
                self.setStyleSheet("""
                QDialog { background: #f5f7fa; color: #1f2933; font-family: Inter, Segoe UI, Arial; }
                QLabel { color: #111827; }
                QLineEdit {
                    background: white; border: 1px solid #d1d5db;
                    border-radius: 10px; padding: 12px; color: #111827; font-size: 14px;
                }
                QLineEdit:focus { border: 1px solid #ed3b3b; }
                QComboBox {
                    background: white; border: 1px solid #d1d5db;
                    border-radius: 10px; padding: 12px 36px 12px 12px; color: #111827;
                    font-size: 14px; min-height: 20px;
                }
                QComboBox:focus { border: 1px solid #ed3b3b; }
                QComboBox::drop-down {
                    subcontrol-origin: padding; subcontrol-position: center right;
                    width: 32px; border: none;
                    border-left: 1px solid #d1d5db;
                    border-top-right-radius: 10px; border-bottom-right-radius: 10px;
                }
                QComboBox::down-arrow { width: 10px; height: 10px; }
                QComboBox QAbstractItemView {
                    background: white; border: 1px solid #d1d5db; border-radius: 8px;
                    padding: 4px; color: #111827; selection-background-color: #f0f0f0;
                    selection-color: #111827; outline: none;
                }
                QComboBox QAbstractItemView::item {
                    padding: 8px 12px; border-radius: 4px; min-height: 24px;
                }
                QComboBox QAbstractItemView::item:hover { background: #f3f4f6; color: #111827; }
                QPushButton {
                    background: #ed3b3b; color: white; border-radius: 10px;
                    padding: 12px 24px; font-size: 14px; font-weight: 600;
                }
                QPushButton:hover { background: #c62e2e; }
                """)

            layout = QVBoxLayout(self)
            form = QFormLayout()

            # === Back Job (Yes / No) ===
            self.back_job = QComboBox()
            self.back_job.addItems(["Yes", "No"])

            # === Billing Status ===
            self.billing_status = QComboBox()
            self.billing_status.addItems([
                "APS",
                "Billed",
                "Pending From Previous Task",
                "Warranty"
            ])

            # === Product Type ===
            self.product_type = QComboBox()
            self.product_type.addItems([
                "CloudPOS",
                "HRMS",
                "Portfolio ERP/PY",
                "POS V3/V5",
                "SAP",
                "WebPOS"
            ])

            # === Support Type ===
            self.support_type = QComboBox()
            self.support_type.addItems([
                "Onsite",
                "Remote/Offsite"
            ])

            # === Confirmed By ===
            self.confirmed_by = QLineEdit()

            # Add fields to form
            form.addRow("Back Job:", self.back_job)
            form.addRow("Billing Status:", self.billing_status)
            form.addRow("Product Type:", self.product_type)
            form.addRow("Support Type:", self.support_type)
            form.addRow("Confirmed By:", self.confirmed_by)

            layout.addLayout(form)

            # === Buttons ===
            # btn_layout = QHBoxLayout()
            # self.ok_btn = QPushButton("Create PBI")
            # cancel_btn = QPushButton("Cancel")

            # self.ok_btn.clicked.connect(self.validate_and_accept)
            # cancel_btn.clicked.connect(self.reject)

            # btn_layout.addWidget(self.ok_btn)
            # btn_layout.addWidget(cancel_btn)
            # layout.addLayout(btn_layout)

            btn_layout = QHBoxLayout()

            self.ok_btn = QPushButton("Create PBI")
            cancel_btn = QPushButton("Cancel")

            self.ok_btn.clicked.connect(self.validate_and_accept)
            cancel_btn.clicked.connect(self.reject)

            # Keep buttons compact
            self.ok_btn.setFixedWidth(120)
            cancel_btn.setFixedWidth(90)

            btn_layout.addStretch()
            btn_layout.addWidget(self.ok_btn)
            btn_layout.addSpacing(10)
            btn_layout.addWidget(cancel_btn)
            btn_layout.addStretch()

            layout.addLayout(btn_layout)


        def validate_and_accept(self):
            """Basic validation like CLI version"""
            if not self.confirmed_by.text().strip():
                AppDialog(
                    "Validation Error",
                    "Confirmed By cannot be empty.",
                    self
                ).exec()
                return
            self.accept()

        def get_values(self):
            return {
                "back_job": self.back_job.currentText(),   # Yes / No (string, same as CLI)
                "billing_status": self.billing_status.currentText(),
                "product_type": self.product_type.currentText(),
                "support_type": self.support_type.currentText(),
                "confirmed_by": self.confirmed_by.text().strip()
            }



class PBICreatedDialog(QDialog):
    def __init__(self, pbi_id, pbi_url, parent=None):
        super().__init__(parent)

        self.setWindowTitle("PBI Created Successfully")
        self.setFixedWidth(420)

        if is_windows_dark_mode():
            self.setStyleSheet("""
            QDialog { background: #0f1923; color: #e6eef5; font-family: Inter, Segoe UI, Arial; }
            QLabel { color: #e6eef5; }
            QLabel a { color: #60a5fa; }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white; border-radius: 10px; padding: 8px 16px; font-size: 14px; font-weight: 600;
            }
            QPushButton:hover { background: #c62e2e; }
            """)
        else:
            self.setStyleSheet("""
            QDialog { background: #f5f7fa; color: #1f2933; font-family: Inter, Segoe UI, Arial; }
            QLabel { color: #111827; }
            QLabel a { color: #2563eb; }
            QPushButton {
                background: #ed3b3b; color: white; border-radius: 10px;
                padding: 8px 16px; font-size: 14px; font-weight: 600;
            }
            QPushButton:hover { background: #c62e2e; }
            """)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        pbi_id_label = QLabel(str(pbi_id))
        pbi_id_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )

        pbi_url_label = QLabel(f'<a href="{pbi_url}">{pbi_url}</a>')
        pbi_url_label.setOpenExternalLinks(True)
        pbi_url_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )

        form.addRow("PBI ID:", pbi_id_label)
        form.addRow("PBI URL:", pbi_url_label)

        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        open_btn = QPushButton("Open in Browser")
        close_btn = QPushButton("Close")

        open_btn.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl(pbi_url))
        )
        close_btn.clicked.connect(self.accept)

        btn_layout.addStretch()
        btn_layout.addWidget(open_btn)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)


class MissingSalesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Missing Sales Generator")
        self.setMinimumWidth(560)

        dark = is_windows_dark_mode()
        if dark:
            self.setStyleSheet("""
            QDialog { background: #0f1923; color: #e6eef5; font-family: Inter, Segoe UI, Arial; }
            QLabel { color: #e6eef5; background: transparent; }
            QLineEdit {
                background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.06);
                border-radius: 8px; padding: 10px 12px; color: #e6eef5; font-size: 13px;
            }
            QLineEdit:focus { border: 1px solid #ed3b3b; }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white; border-radius: 10px; padding: 10px 20px; font-size: 14px; font-weight: 600;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton#GhostBtn {
                background: transparent; border: 1px solid rgba(255,255,255,0.15);
                color: #98a0a6;
            }
            QPushButton#GhostBtn:hover { border: 1px solid #98a0a6; color: #e6eef5; }
            QTextEdit {
                background: #111a22; border: 1px solid #2a3a48; border-radius: 8px;
                color: #e6eef5; font-size: 12px; font-family: Consolas, monospace;
            }
            """)
        else:
            self.setStyleSheet("""
            QDialog { background: #f5f7fa; color: #1f2933; font-family: Inter, Segoe UI, Arial; }
            QLabel { color: #111827; background: transparent; }
            QLineEdit {
                background: white; border: 1px solid #d1d5db;
                border-radius: 8px; padding: 10px 12px; color: #111827; font-size: 13px;
            }
            QLineEdit:focus { border: 1px solid #ed3b3b; }
            QPushButton {
                background: #ed3b3b; color: white; border-radius: 10px;
                padding: 10px 20px; font-size: 14px; font-weight: 600;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton#GhostBtn {
                background: transparent; border: 1px solid #d1d5db; color: #6b7280;
            }
            QPushButton#GhostBtn:hover { border: 1px solid #6b7280; color: #111827; }
            QTextEdit {
                background: white; border: 1px solid #d1d5db; border-radius: 8px;
                color: #111827; font-size: 12px; font-family: Consolas, monospace;
            }
            """)

        tc = parent._theme_colors() if parent and hasattr(parent, "_theme_colors") else {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(0)

        # --- Header ---
        header = QLabel("Missing Sales Generator")
        header.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(header)
        layout.addSpacing(4)

        subtitle = QLabel("Generate diagnostic SQL and helper scripts for missing POS sales")
        subtitle.setStyleSheet("color: #6b7280; font-size: 12px;")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        layout.addSpacing(20)

        # --- Divider ---
        divider = QtWidgets.QFrame()
        divider.setFixedHeight(1)
        divider.setStyleSheet("background: rgba(255,255,255,0.08) if dark else #e5e7eb; border: none;")
        divider.setStyleSheet(f"background: {'rgba(255,255,255,0.08)' if dark else '#e5e7eb'}; border: none;")
        layout.addWidget(divider)
        layout.addSpacing(16)

        # --- Section: Terminal Info ---
        section1 = QLabel("TERMINAL INFORMATION")
        section1.setStyleSheet("color: #ed3b3b; font-size: 11px; font-weight: 700; letter-spacing: 1.5px;")
        layout.addWidget(section1)
        layout.addSpacing(12)

        field_label_style = "color: #98a0a6; font-size: 12px; font-weight: 600;" if dark else "color: #374151; font-size: 12px; font-weight: 600;"

        def make_field(label_text, widget):
            block = QVBoxLayout()
            block.setSpacing(5)
            lbl = QLabel(label_text)
            lbl.setStyleSheet(field_label_style)
            block.addWidget(lbl)
            block.addWidget(widget)
            return block

        self.company_id = QLineEdit()
        self.company_id.setPlaceholderText("WPOS-26012789")
        self.terminal = QLineEdit()
        self.terminal.setPlaceholderText("1 or 0001")
        self.year = QLineEdit()
        self.year.setPlaceholderText("YYYY")
        self.month = QLineEdit()
        self.month.setPlaceholderText("MM")

        grid1 = QtWidgets.QGridLayout()
        grid1.setHorizontalSpacing(16)
        grid1.setVerticalSpacing(12)
        grid1.addLayout(make_field("Company ID", self.company_id), 0, 0)
        grid1.addLayout(make_field("Terminal", self.terminal), 0, 1)
        grid1.addLayout(make_field("Missing Year", self.year), 1, 0)
        grid1.addLayout(make_field("Missing Month", self.month), 1, 1)
        layout.addLayout(grid1)

        layout.addSpacing(20)

        # --- Divider 2 ---
        divider2 = QtWidgets.QFrame()
        divider2.setFixedHeight(1)
        divider2.setStyleSheet(f"background: {'rgba(255,255,255,0.08)' if dark else '#e5e7eb'}; border: none;")
        layout.addWidget(divider2)
        layout.addSpacing(16)

        # --- Section: Connection ---
        section2 = QLabel("CONNECTION")
        section2.setStyleSheet("color: #ed3b3b; font-size: 11px; font-weight: 700; letter-spacing: 1.5px;")
        layout.addWidget(section2)
        layout.addSpacing(12)

        self.link = QLineEdit()
        self.link.setPlaceholderText("https://example.com")
        layout.addLayout(make_field("Company Link", self.link))

        layout.addSpacing(24)

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        copy_btn = QPushButton("\U0001f4cb Copy")
        copy_btn.setObjectName("GhostBtn")
        copy_btn.setFixedHeight(40)
        copy_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        copy_btn.clicked.connect(self.copy_to_clipboard)

        self.generate_btn = QPushButton("Generate & Save")
        self.generate_btn.setFixedHeight(40)
        self.generate_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.generate_btn.clicked.connect(self.generate_and_save)

        btn_layout.addStretch()
        btn_layout.addWidget(copy_btn)
        btn_layout.addWidget(self.generate_btn)
        layout.addLayout(btn_layout)

        layout.addSpacing(16)

        # --- Preview ---
        preview_label = QLabel("PREVIEW")
        preview_label.setStyleSheet("color: #ed3b3b; font-size: 11px; font-weight: 700; letter-spacing: 1.5px;")
        layout.addWidget(preview_label)
        layout.addSpacing(8)

        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setMinimumHeight(100)
        self.preview.setMaximumHeight(160)
        self.preview.setPlaceholderText("Generated script will appear here...")
        layout.addWidget(self.preview)

    def _pad_terminal(self, terminal):
        t = terminal.strip()
        if len(t) == 1:
            return "000" + t
        elif len(t) == 2:
            return "00" + t
        elif len(t) == 3:
            return "0" + t
        return t

    def _build_content(self):
        cid = self.company_id.text().strip()
        terminal = self._pad_terminal(self.terminal.text())
        year = self.year.text().strip()
        month = self.month.text().strip()
        link = self.link.text().strip()
        missing_date = year + month + "01"
        missing_date2 = year + "1231"

        lacking = (
            'WITH LACKING:\n'
            'SELECT fzcounter, fsale_date, MIN(fdocument_no) AS fmin, MAX(fdocument_no) AS fmax, '
            'SUM(fgross) AS fgross, COUNT(fdocument_no) AS fdocument_no, '
            'MAX(fdocument_no) - MIN(fdocument_no) + 1 AS Expected_OR_count, '
            'count(*)-(max(fdocument_no)-min(fdocument_no)+1) as missing_OR '
            'FROM pos_sale WHERE fcompanyid="' + cid + '" and ftermid="' + terminal + '" '
            'AND fsale_date>="' + missing_date + '" AND fdocument_no <> 0 '
            'GROUP BY fsale_date, fzcounter'
        )

        full_missing = (
            '\n\nFULL MISSING SCRIPT:\n'
            'select fsale_date, fzcounter, min(fdocument_no) as min_fdoc, max(fdocument_no) as max_fdoc, '
            'count(*) as trxcnt,(max(fdocument_no) - min(fdocument_no) + 1) as expected, '
            "case when count(*) = (max(fdocument_no) - min(fdocument_no) + 1) then 'Y' else 'N' end as same, "
            'count(*) - (max(fdocument_no) - min(fdocument_no) + 1) as lacking, sum(fgross) as fgross '
            'from pos_sale where fcompanyid="' + cid + '" AND ftermid="' + terminal + '" '
            "and fsale_date >= '" + missing_date + "' and fsale_date <= '" + missing_date2 + "' "
            "and fdocument_no <> '0' group by fsale_date, fzcounter having (same = 'N')"
        )

        excel = (
            '\n\nFOR EXCEL:\n'
            '=IF(E8-E7=1,"OK","MISSING")\n'
            '=IF(E8 - E7 > 1, IF(E8-E7 > 2, E7 + 1 & "-" & E8 - 1, E8 -1), "OK")'
        )

        today = datetime.now().strftime("%Y%m%d")
        select_or = (
            '\n\nSELECT OR/DATE SCRIPT:\n'
            'select * from pos_sale where fcompanyid="' + cid + '" and ftermid="' + terminal + '" '
            "and fsale_date='" + today + "' and fzcounter='2251'"
        )

        update_frecno = (
            '\n\nUPDATE FRECNO:\n'
            "update pos_sale set fzcounter='' where fcompanyid='" + cid + "' "
            "and ftermid='" + terminal + "' and fsale_date='date nga e set' and fzcounter='0'"
        )

        sql_delete = (
            '\n\nSQL Query Delete:\n'
            "DELETE from pos_sale where frecno not in ('', '', '',);\n"
            "DELETE from pos_sale_payment where frecno not in ('', '', '',);\n"
            "DELETE from pos_sale_product where frecno not in ('', '', '',);\n"
            "DELETE from pos_reading;\nDELETE from pos_reading_summary;"
        )

        sql_update = (
            '\n\nSQL Query Update:\n'
            'UPDATE pos_sale_payment\n'
            "SET frecno = \nCASE frecno\n    WHEN '' THEN ''\nELSE frecno\nEND;"
        )

        sql_transmit = (
            '\n\nSQL Transmit:\n'
            'update pos_sale set ftransmit="0000";\n'
            'update pos_sale_payment set ftransmit="0000";\n'
            'update pos_sale_product set ftransmit="0000";'
        )

        for_exception = '\n\nFOR EXCEPTION:\n' + link + '/appserv/app/batch/sys_get_exception.php'
        for_bypass = '\n\nFOR BYPASS:\n' + link + '/appserv/app/batch/sys_bypass_exception.php/'
        for_rereading = (
            '\n\nFOR REREADING:\n' + link +
            '/appserv/app/batch/fix/recreate_reading.php?fcompanyid=' + cid +
            '&ftermid=' + terminal + '&fsale_date=&fzcounter=&fend_date=&fcreate_flag=1'
        )

        return lacking + full_missing + excel + select_or + update_frecno + sql_delete + sql_update + sql_transmit + for_exception + for_bypass + for_rereading

    def _validate(self):
        if not self.company_id.text().strip():
            AppDialog("Validation Error", "Company ID cannot be empty.", self).exec()
            return False
        if not self.terminal.text().strip():
            AppDialog("Validation Error", "Terminal cannot be empty.", self).exec()
            return False
        if not self.year.text().strip():
            AppDialog("Validation Error", "Missing Year cannot be empty.", self).exec()
            return False
        if not self.month.text().strip():
            AppDialog("Validation Error", "Missing Month cannot be empty.", self).exec()
            return False
        return True

    def generate_and_save(self):
        if not self._validate():
            return
        content = self._build_content()
        self.preview.setPlainText(content)

        today_folder = datetime.now().strftime("%m-%d-%Y")
        os.makedirs(today_folder, exist_ok=True)
        default_path = os.path.join(today_folder, "Script.txt")

        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Script", default_path, "Text Files (*.txt);;All Files (*)"
        )
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            logger.info("Missing Sales script saved: %s", path)
            AppDialog("Success", f"Script saved to:\n{path}", self).exec()
        except Exception:
            logger.exception("Failed to save Missing Sales script")
            AppDialog("Error", "Failed to save script. Check logs.", self).exec()

    def copy_to_clipboard(self):
        if not self._validate():
            return
        content = self._build_content()
        self.preview.setPlainText(content)
        QtWidgets.QApplication.clipboard().setText(content)
        self.preview.append("\n--- Copied to clipboard ---")


class AyalaWorker(QThread):
    """
    Parses a folder of hourly sales CSVs and writes the reconciliation
    workbook, off the GUI thread. This used to run inline inside
    AyalaCheckerDialog.generate_report with periodic processEvents() calls,
    so a large folder froze the dialog until the whole report was built.

    `processor` is the owning AyalaCheckerDialog. Only its pure methods
    (parse_hourly_csv, parse_eod_csv, generate_excel) are called from here -
    none of them touch widgets.
    """
    log = pyqtSignal(str)
    succeeded = pyqtSignal(str)   # output path
    failed = pyqtSignal(str)

    def __init__(self, folder, output_name, processor, parent=None):
        super().__init__(parent)
        self.folder = folder
        self.output_name = output_name
        self.processor = processor

    @staticmethod
    def _file_number(name):
        """Trailing _<number> in a CSV name, or None when there isn't one."""
        stem = name[:-4] if name.lower().endswith(".csv") else name
        match = re.search(r"_(\d+)$", stem)
        return int(match.group(1)) if match else None

    def run(self):
        try:
            self.log.emit("Starting Ayala Checker processing...")
            self.log.emit(f"Folder: {self.folder}")

            csv_files = sorted(
                f for f in os.listdir(self.folder) if f.lower().endswith('.csv')
            )
            self.log.emit(f"Found {len(csv_files)} CSV file(s)")

            hourly_files = []
            eod_file = None
            for csv_file in csv_files:
                if csv_file.upper().startswith('EOD'):
                    eod_file = csv_file
                else:
                    hourly_files.append(csv_file)

            # Files whose names carry no trailing _<number> are processed last
            # instead of aborting the run with a ValueError, which is what the
            # old int(name.split('_')[-1]) sort did on any unexpected file.
            numbered = [f for f in hourly_files if self._file_number(f) is not None]
            unnumbered = [f for f in hourly_files if self._file_number(f) is None]
            numbered.sort(key=self._file_number)
            unnumbered.sort()

            if unnumbered:
                self.log.emit(
                    "⚠️ No transaction number in name, processing last: "
                    + ", ".join(unnumbered)
                )
            hourly_files = numbered + unnumbered

            self.log.emit(f"Hourly sales files: {len(hourly_files)}")
            self.log.emit(f"EOD file: {eod_file if eod_file else 'Not found'}")

            if not hourly_files:
                self.failed.emit("No hourly sales CSV files found in the folder.")
                return

            self.log.emit("")
            self.log.emit("Parsing hourly sales CSVs...")
            hourly_transactions = []
            header = {}
            for csv_file in hourly_files:
                file_header, transactions = self.processor.parse_hourly_csv(
                    os.path.join(self.folder, csv_file)
                )
                if not header and file_header:
                    header = file_header
                hourly_transactions.extend(transactions)
                self.log.emit(f"  {csv_file}: {len(transactions)} transaction(s)")

            self.log.emit("")
            self.log.emit(
                f"Total transactions from hourly CSVs: {len(hourly_transactions)}"
            )

            eod_data = {}
            if eod_file:
                self.log.emit("")
                self.log.emit("Parsing EOD CSV...")
                eod_data = self.processor.parse_eod_csv(
                    os.path.join(self.folder, eod_file)
                )
                self.log.emit(f"  EOD data loaded: {len(eod_data)} fields")

            self.log.emit("")
            self.log.emit("Generating Excel report...")
            output_name = self.output_name
            if not output_name.lower().endswith('.xlsx'):
                output_name += '.xlsx'
            output_path = os.path.join(self.folder, output_name)
            self.processor.generate_excel(
                hourly_transactions, eod_data, output_path, header
            )

            self.log.emit("")
            self.log.emit("✅ Report generated successfully!")
            self.log.emit(f"Output: {output_path}")
            self.succeeded.emit(output_path)

        except Exception as e:
            logger.exception("Ayala Checker error")
            self.failed.emit(str(e))


class AyalaCheckerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ayala Checker")
        self.setMinimumWidth(700)
        self.setMinimumHeight(600)

        dark = is_windows_dark_mode()
        if dark:
            self.setStyleSheet("""
            QDialog { background: #0f1923; color: #e6eef5; font-family: Inter, Segoe UI, Arial; }
            QLabel { color: #e6eef5; background: transparent; }
            QLineEdit {
                background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.06);
                border-radius: 8px; padding: 10px 12px; color: #e6eef5; font-size: 13px;
            }
            QLineEdit:focus { border: 1px solid #ed3b3b; }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white; border-radius: 10px; padding: 10px 20px; font-size: 14px; font-weight: 600;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton#GhostBtn {
                background: transparent; border: 1px solid rgba(255,255,255,0.15);
                color: #98a0a6;
            }
            QPushButton#GhostBtn:hover { border: 1px solid #98a0a6; color: #e6eef5; }
            QTextEdit {
                background: #111a22; border: 1px solid #2a3a48; border-radius: 8px;
                color: #e6eef5; font-size: 12px; font-family: Consolas, monospace;
            }
            """)
        else:
            self.setStyleSheet("""
            QDialog { background: #f5f7fa; color: #1f2933; font-family: Inter, Segoe UI, Arial; }
            QLabel { color: #111827; background: transparent; }
            QLineEdit {
                background: white; border: 1px solid #d1d5db;
                border-radius: 8px; padding: 10px 12px; color: #111827; font-size: 13px;
            }
            QLineEdit:focus { border: 1px solid #ed3b3b; }
            QPushButton {
                background: #ed3b3b; color: white; border-radius: 10px;
                padding: 10px 20px; font-size: 14px; font-weight: 600;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton#GhostBtn {
                background: transparent; border: 1px solid #d1d5db; color: #6b7280;
            }
            QPushButton#GhostBtn:hover { border: 1px solid #6b7280; color: #111827; }
            QTextEdit {
                background: white; border: 1px solid #d1d5db; border-radius: 8px;
                color: #111827; font-size: 12px; font-family: Consolas, monospace;
            }
            """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(0)

        # --- Header ---
        header = QLabel("Ayala Checker")
        header.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(header)
        layout.addSpacing(4)

        subtitle = QLabel("Compare hourly sales CSVs against EOD CSV and generate organized Excel report")
        subtitle.setStyleSheet("color: #6b7280; font-size: 12px;")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        layout.addSpacing(20)

        # --- Divider ---
        divider = QtWidgets.QFrame()
        divider.setFixedHeight(1)
        divider.setStyleSheet(f"background: {'rgba(255,255,255,0.08)' if dark else '#e5e7eb'}; border: none;")
        layout.addWidget(divider)
        layout.addSpacing(16)

        # --- Section: Folder Selection ---
        section1 = QLabel("FOLDER SELECTION")
        section1.setStyleSheet("color: #ed3b3b; font-size: 11px; font-weight: 700; letter-spacing: 1.5px;")
        layout.addWidget(section1)
        layout.addSpacing(12)

        field_label_style = "color: #98a0a6; font-size: 12px; font-weight: 600;" if dark else "color: #374151; font-size: 12px; font-weight: 600;"

        def make_field(label_text, widget):
            block = QVBoxLayout()
            block.setSpacing(5)
            lbl = QLabel(label_text)
            lbl.setStyleSheet(field_label_style)
            block.addWidget(lbl)
            block.addWidget(widget)
            return block

        # Folder path input with browse button
        folder_label = QLabel("Tenant Data Folder")
        folder_label.setStyleSheet(field_label_style)
        layout.addWidget(folder_label)
        layout.addSpacing(5)

        folder_layout = QHBoxLayout()
        self.folder_path = QLineEdit()
        self.folder_path.setPlaceholderText("Select folder containing CSV files...")
        self.folder_path.setReadOnly(True)
        folder_layout.addWidget(self.folder_path)

        browse_btn = QPushButton("Browse")
        browse_btn.setObjectName("GhostBtn")
        browse_btn.setFixedHeight(40)
        browse_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        browse_btn.clicked.connect(self.browse_folder)
        folder_layout.addWidget(browse_btn)

        layout.addLayout(folder_layout)
        layout.addSpacing(10)

        # Output file name
        self.output_name = QLineEdit()
        self.output_name.setPlaceholderText("Output filename (e.g., Ayala_Report.xlsx)")
        self.output_name.setText("Ayala_Report.xlsx")
        layout.addLayout(make_field("Output Filename", self.output_name))

        layout.addSpacing(20)

        # --- Divider 2 ---
        divider2 = QtWidgets.QFrame()
        divider2.setFixedHeight(1)
        divider2.setStyleSheet(f"background: {'rgba(255,255,255,0.08)' if dark else '#e5e7eb'}; border: none;")
        layout.addWidget(divider2)
        layout.addSpacing(16)

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.generate_btn = QPushButton("Generate Report")
        self.generate_btn.setFixedHeight(40)
        self.generate_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.generate_btn.clicked.connect(self.generate_report)

        btn_layout.addStretch()
        btn_layout.addWidget(self.generate_btn)
        layout.addLayout(btn_layout)

        layout.addSpacing(16)

        # --- Log Area ---
        log_label = QLabel("PROCESSING LOG")
        log_label.setStyleSheet("color: #ed3b3b; font-size: 11px; font-weight: 700; letter-spacing: 1.5px;")
        layout.addWidget(log_label)
        layout.addSpacing(8)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMinimumHeight(150)
        self.log_area.setPlaceholderText("Processing log will appear here...")
        layout.addWidget(self.log_area)

    def browse_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Select Tenant Data Folder"
        )
        if folder:
            self.folder_path.setText(folder)
            logger.info("Ayala Checker: Selected folder: %s", folder)

    def log(self, message):
        self.log_area.append(message)
        QtWidgets.QApplication.processEvents()

    def generate_report(self):
        folder = self.folder_path.text().strip()
        if not folder:
            AppDialog("Error", "Please select a folder first.", self).exec()
            return

        if not os.path.isdir(folder):
            AppDialog("Error", "Selected folder does not exist.", self).exec()
            return

        # The parsing and workbook writing run on AyalaWorker, so the dialog
        # stays responsive instead of freezing until the report is finished.
        self.generate_btn.setEnabled(False)
        self.generate_btn.setText("Processing...")

        self.worker = AyalaWorker(folder, self.output_name.text().strip(), self, self)
        self.worker.log.connect(self.log)
        self.worker.succeeded.connect(self._on_ayala_succeeded)
        self.worker.failed.connect(self._on_ayala_failed)
        self.worker.finished.connect(self._on_ayala_thread_finished)
        self.worker.start()

    def _on_ayala_succeeded(self, output_path):
        self.log("")
        self.log(f"✅ Report generated:\n{output_path}")
        AppDialog("Success", f"Report generated:\n{output_path}", self).exec()

    def _on_ayala_failed(self, message):
        self.log("")
        self.log(f"❌ Error: {message}")
        AppDialog("Error", f"Failed to generate report:\n{message}", self).exec()

    def _on_ayala_thread_finished(self):
        self.generate_btn.setEnabled(True)
        self.generate_btn.setText("Generate Report")

    def closeEvent(self, event):
        """Refuse to close while the worker is still writing the workbook."""
        worker = getattr(self, "worker", None)
        if worker is not None and worker.isRunning():
            self.log("Please wait — the report is still being generated.")
            event.ignore()
            return
        super().closeEvent(event)

    def reject(self):
        """Esc must not tear the dialog down mid-run either: the worker calls
        methods on this object, so it has to outlive the thread."""
        worker = getattr(self, "worker", None)
        if worker is not None and worker.isRunning():
            self.log("Please wait — the report is still being generated.")
            return
        super().reject()

    def parse_hourly_csv(self, file_path):
        """
        Parse one hourly sales CSV.

        Returns (header, transactions). The header block (CCCODE,
        MERCHANT_NAME, TRN_DATE, NO_TRN) is returned on its own: it used to be
        appended to the transaction list as element 0, which made every hourly
        file contribute a blank pseudo-transaction to the Excel report.
        """
        transactions = []
        header = {}
        current_transaction = None

        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue

                parts = line.split(',', 1)
                if len(parts) != 2:
                    continue

                key, value = parts

                # Header fields appear once, before the first transaction
                if key in ['CCCODE', 'MERCHANT_NAME', 'TRN_DATE', 'NO_TRN']:
                    if current_transaction is None:
                        header[key] = value
                    continue

                # Transaction start marker
                if key == 'CDATE' and current_transaction is None:
                    current_transaction = {'CDATE': value, 'items': []}
                    continue

                if current_transaction is not None:
                    # Item fields
                    if key in ['QTY', 'ITEMCODE', 'PRICE', 'LDISC']:
                        if key == 'QTY' and current_transaction.get('items') and 'QTY' in current_transaction['items'][-1]:
                            # Start new item
                            current_transaction['items'].append({key: value})
                        elif current_transaction['items']:
                            current_transaction['items'][-1][key] = value
                        else:
                            current_transaction['items'].append({key: value})
                        continue

                    # Check if this is a new transaction (CDATE appears again)
                    if key == 'CDATE' and 'CDATE' in current_transaction:
                        transactions.append(current_transaction)
                        current_transaction = {'CDATE': value, 'items': []}
                        continue

                    # Regular transaction field
                    current_transaction[key] = value

            # Don't forget the last transaction
            if current_transaction:
                transactions.append(current_transaction)

        return header, transactions

    def parse_eod_csv(self, file_path):
        """Parse an EOD CSV file and return dictionary of fields."""
        eod_data = {}
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.split(',', 1)
                if len(parts) == 2:
                    eod_data[parts[0]] = parts[1]
        return eod_data

    def generate_excel(self, hourly_transactions, eod_data, output_path, header=None):
        """Generate Excel file matching the Sample Output format."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Ayala Report"

        # Define styles
        header_font = Font(bold=True, size=11, name='Calibri')
        header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        normal_font = Font(size=11, name='Calibri')
        total_font = Font(bold=True, size=11, color="0000FF", name='Calibri')
        total_fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
        eod_font = Font(bold=True, size=11, color="006600", name='Calibri')
        eod_fill = PatternFill(start_color="E6FFE6", end_color="E6FFE6", fill_type="solid")

        # Header block comes from the CSV parser, not from a transaction
        header = header or {}

        # Row 1-4: Header fields
        header_fields = ['CCCODE', 'MERCHANT_NAME', 'TRN_DATE', 'NO_TRN']
        for row_idx, field in enumerate(header_fields, start=1):
            ws.cell(row=row_idx, column=1, value=field)
            ws.cell(row=row_idx, column=2, value=header.get(field, ''))
            ws.cell(row=row_idx, column=1).font = normal_font

        # Starting row for transactions
        start_row = 5

        # Define transaction fields to display
        transaction_fields = [
            'CDATE', 'TRN_TIME', 'TER_NO', 'TRANSACTION_NO',
            'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS',
            'VATEXEMPT_SLS', 'VATEXEMPT_AMNT', 'LOCAL_TAX',
            'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC',
            'STORE_DISC', 'OTHER_DISC', 'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR',
            'CASH_SLS', 'CARD_SLS', 'EPAY_SLS', 'DCARD_SLS', 'OTHERSL_SLS',
            'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS', 'AMEX_SLS',
            'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
            'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS',
            'VISADEBIT_SLS', 'PAYPAL_SLS', 'ONLINE_SLS',
            'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3', 'OPEN_SALES_4',
            'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
            'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
            'MOBILE_NO', 'NO_CUST', 'TRN_TYPE', 'SLS_FLAG', 'VAT_PCT', 'QTY_SLD'
        ]

        # Define numeric fields that should be converted to numbers
        numeric_fields = [
            'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS',
            'VATEXEMPT_SLS', 'VATEXEMPT_AMNT', 'LOCAL_TAX',
            'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC',
            'STORE_DISC', 'OTHER_DISC', 'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR',
            'CASH_SLS', 'CARD_SLS', 'EPAY_SLS', 'DCARD_SLS', 'OTHERSL_SLS',
            'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS', 'AMEX_SLS',
            'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
            'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS',
            'VISADEBIT_SLS', 'PAYPAL_SLS', 'ONLINE_SLS',
            'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3', 'OPEN_SALES_4',
            'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
            'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
            'NO_CUST', 'QTY_SLD', 'VAT_PCT', 'PRICE', 'LDISC', 'QTY'
        ]

        def convert_value(field, value):
            """Convert value to number if field is numeric, otherwise keep as string."""
            if field in numeric_fields and value:
                try:
                    if '.' in str(value):
                        return float(value)
                    else:
                        return int(value)
                except (ValueError, TypeError):
                    return value
            return value

        # Find max items per transaction
        max_items = max(len(t.get('items', [])) for t in hourly_transactions) if hourly_transactions else 0

        # Write transactions side by side
        num_transactions = len(hourly_transactions)
        for trn_idx, transaction in enumerate(hourly_transactions):
            col_start = 1 + (trn_idx * 2)  # Each transaction takes 2 columns

            # Write transaction fields
            for row_offset, field in enumerate(transaction_fields):
                row = start_row + row_offset
                ws.cell(row=row, column=col_start, value=field)
                value = transaction.get(field, '')
                ws.cell(row=row, column=col_start + 1, value=convert_value(field, value))
                ws.cell(row=row, column=col_start).font = normal_font

            # Write items
            item_start_row = start_row + len(transaction_fields)
            for item_idx in range(max_items):
                row = item_start_row + (item_idx * 4)  # 4 rows per item (QTY, ITEMCODE, PRICE, LDISC)
                if item_idx < len(transaction.get('items', [])):
                    item = transaction['items'][item_idx]
                    ws.cell(row=row, column=col_start, value='QTY')
                    ws.cell(row=row, column=col_start).font = normal_font
                    ws.cell(row=row, column=col_start + 1, value=convert_value('QTY', item.get('QTY', '')))
                    ws.cell(row=row + 1, column=col_start, value='ITEMCODE')
                    ws.cell(row=row + 1, column=col_start).font = normal_font
                    ws.cell(row=row + 1, column=col_start + 1, value=item.get('ITEMCODE', ''))
                    ws.cell(row=row + 2, column=col_start, value='PRICE')
                    ws.cell(row=row + 2, column=col_start).font = normal_font
                    ws.cell(row=row + 2, column=col_start + 1, value=convert_value('PRICE', item.get('PRICE', '')))
                    ws.cell(row=row + 3, column=col_start, value='LDISC')
                    ws.cell(row=row + 3, column=col_start).font = normal_font
                    ws.cell(row=row + 3, column=col_start + 1, value=convert_value('LDISC', item.get('LDISC', '')))

        # Add SUM column for hourly totals
        sum_col = 1 + (num_transactions * 2)
        ws.cell(row=4, column=sum_col, value="HOURLY TOTAL")
        ws.cell(row=4, column=sum_col).font = total_font
        ws.cell(row=4, column=sum_col).fill = total_fill

        # Fields that should have SUM formulas
        sum_fields = [
            'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS',
            'VATEXEMPT_SLS', 'VATEXEMPT_AMNT', 'LOCAL_TAX',
            'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC',
            'STORE_DISC', 'OTHER_DISC', 'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR',
            'CASH_SLS', 'CARD_SLS', 'EPAY_SLS', 'DCARD_SLS', 'OTHERSL_SLS',
            'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS', 'AMEX_SLS',
            'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
            'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS',
            'VISADEBIT_SLS', 'PAYPAL_SLS', 'ONLINE_SLS',
            'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3', 'OPEN_SALES_4',
            'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
            'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
            'NO_CUST', 'QTY_SLD'
        ]

        # Add SUM formulas for each transaction field
        for row_offset, field in enumerate(transaction_fields):
            row = start_row + row_offset
            ws.cell(row=row, column=sum_col, value=field)
            ws.cell(row=row, column=sum_col).font = normal_font

            if field in sum_fields:
                # Build SUM formula using range format
                first_col = get_column_letter(2)  # Column B
                last_col = get_column_letter(1 + (num_transactions * 2) - 1)  # Last transaction value column
                formula = f"=SUM({first_col}{row}:{last_col}{row})"
                ws.cell(row=row, column=sum_col + 1, value=formula)
                ws.cell(row=row, column=sum_col + 1).font = total_font

        # Add separator column (black)
        separator_col = sum_col + 2
        separator_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        separator_letter = get_column_letter(separator_col)
        ws.column_dimensions[separator_letter].width = 3.2

        # Add EOD data on the right side
        if eod_data:
            eod_col = sum_col + 3  # After separator column

            # EOD header starts 3 rows above the transaction data
            eod_header_row = start_row - 3  # Row 2 instead of row 5
            ws.cell(row=eod_header_row, column=eod_col, value="EOD VALUE")
            ws.cell(row=eod_header_row, column=eod_col).font = eod_font
            ws.cell(row=eod_header_row, column=eod_col).fill = eod_fill

            # Write ALL EOD fields starting from eod_header_row + 1
            eod_row = eod_header_row + 1
            for eod_field, eod_value in eod_data.items():
                ws.cell(row=eod_row, column=eod_col, value=eod_field)
                ws.cell(row=eod_row, column=eod_col).font = normal_font
                ws.cell(row=eod_row, column=eod_col + 1, value=convert_value(eod_field, eod_value))
                eod_row += 1

        # Auto-fit column widths (except separator)
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            # Skip separator column
            if column_letter == separator_letter:
                continue
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            ws.column_dimensions[column_letter].width = adjusted_width

        wb.save(output_path)
        logger.info("Ayala Checker: Excel report saved to %s", output_path)

    def compare_totals(self, hourly_transactions, eod_data):
        """Compare hourly totals against EOD values."""
        # Calculate totals from hourly transactions
        numeric_fields = [
            'GROSS_SLS', 'VAT_AMNT', 'VATABLE_SLS', 'NONVAT_SLS',
            'VATEXEMPT_SLS', 'VATEXEMPT_AMNT', 'LOCAL_TAX',
            'PWD_DISC', 'SNRCIT_DISC', 'EMPLO_DISC', 'AYALA_DISC',
            'STORE_DISC', 'OTHER_DISC', 'REFUND_AMT', 'SCHRGE_AMT', 'OTHER_SCHR',
            'CASH_SLS', 'CARD_SLS', 'EPAY_SLS', 'DCARD_SLS', 'OTHERSL_SLS',
            'CHECK_SLS', 'GC_SLS', 'MASTERCARD_SLS', 'VISA_SLS', 'AMEX_SLS',
            'DINERS_SLS', 'JCB_SLS', 'GCASH_SLS', 'PAYMAYA_SLS', 'ALIPAY_SLS',
            'WECHAT_SLS', 'GRAB_SLS', 'FOODPANDA_SLS', 'MASTERDEBIT_SLS',
            'VISADEBIT_SLS', 'PAYPAL_SLS', 'ONLINE_SLS',
            'OPEN_SALES', 'OPEN_SALES_2', 'OPEN_SALES_3', 'OPEN_SALES_4',
            'OPEN_SALES_5', 'OPEN_SALES_6', 'OPEN_SALES_7', 'OPEN_SALES_8',
            'OPEN_SALES_9', 'OPEN_SALES_10', 'OPEN_SALES_11', 'GC_EXCESS',
            'NO_CUST', 'QTY_SLD'
        ]

        hourly_totals = {}
        for field in numeric_fields:
            hourly_totals[field] = 0

        for transaction in hourly_transactions:
            for field in numeric_fields:
                try:
                    hourly_totals[field] += float(transaction.get(field, 0))
                except (ValueError, TypeError):
                    pass

        # EOD field mapping (some fields have different names)
        eod_field_mapping = {
            'OTHERSL_SLS': 'OTHER_SLS',  # EOD uses OTHER_SLS
        }

        # Compare with EOD
        self.log("\n--- Comparison Results ---")
        match_count = 0
        mismatch_count = 0

        for field, hourly_val in hourly_totals.items():
            # Skip fields that are always 0 or text
            if hourly_val == 0 and field not in ['NO_CUST']:
                continue

            # Get corresponding EOD field
            eod_field = eod_field_mapping.get(field, field)
            eod_val_str = eod_data.get(eod_field, '0')
            try:
                eod_val = float(eod_val_str)
            except ValueError:
                eod_val = 0

            diff = abs(hourly_val - eod_val)
            if diff < 1:  # Match if difference is less than 1
                status = "✅ MATCH"
                match_count += 1
            else:
                status = f"❌ MISMATCH (diff: {diff:.2f})"
                mismatch_count += 1

            self.log(f"{field}: Hourly={hourly_val:.2f}, EOD={eod_val:.2f} {status}")

        self.log(f"\nSummary: {match_count} matched, {mismatch_count} mismatched")


# ================= MAIN WINDOW =================
class ReportWorker(QThread):
        progress = pyqtSignal(str)  # signal to update status
        # NOT named `finished`: QThread already owns a zero-argument `finished`
        # signal and redefining it here shadowed that one, so the thread's real
        # completion was unobservable (and an emit with no argument could raise).
        done = pyqtSignal(str)
        report_saved = pyqtSignal(str)

        def __init__(self, work_item_id, client_name, location_type1, location_type2, user_data, template_path):
            super().__init__()
            self.work_item_id = work_item_id
            self.client_name = client_name
            self.location_type1 = location_type1
            self.location_type2 = location_type2
            self.user_data = user_data
            self.template_path = template_path

        def run(self):
            logger.info("=== ReportWorker started for work item %s ===", self.work_item_id)
            pat = self.user_data["pat"]
            my_display_name = self.user_data["display_name"]
            my_unique_name = self.user_data["unique_name"]

            self.progress.emit(f"Fetching work item {self.work_item_id}...")
            work_item_url = f"{org_url}/_apis/wit/workitems/{self.work_item_id}?api-version=5.1"
            try:
                logger.debug("GET %s", work_item_url)
                resp = requests.get(work_item_url, auth=HTTPBasicAuth("", pat), timeout=HTTP_TIMEOUT)
                resp.raise_for_status()
                work_item_data = resp.json()
                logger.info("Work item fetched: id=%s, title=%r", work_item_data.get("id"), work_item_data.get("fields", {}).get("System.Title", "")[:80])
            except Exception as e:
                logger.exception("Failed to fetch work item %s", self.work_item_id)
                self.done.emit(f"❌ Error fetching work item: {e}")
                return

            ticket_number = work_item_data.get("id")
            ticket_subject = work_item_data.get("fields", {}).get("System.Title", "")

            self.progress.emit("Fetching comments...")
            comments_url = f"{org_url}/_apis/wit/workItems/{self.work_item_id}/comments?api-version=5.1-preview.3"
            try:
                logger.debug("GET %s", comments_url)
                resp = requests.get(comments_url, auth=HTTPBasicAuth("", pat), timeout=HTTP_TIMEOUT)
                resp.raise_for_status()
                comments_data = resp.json()
                total_comments = len(comments_data.get("comments", []))
                logger.info("Fetched %d comment(s) for work item %s", total_comments, self.work_item_id)
            except Exception as e:
                logger.exception("Failed to fetch comments for work item %s", self.work_item_id)
                self.done.emit(f"❌ Error fetching comments: {e}")
                return

            my_comments = [
                c for c in comments_data.get("comments", [])
                if c.get("createdBy", {}).get("displayName") == my_display_name
                or c.get("createdBy", {}).get("uniqueName") == my_unique_name
            ]
            logger.debug("Matched %d comment(s) for user displayName=%r, uniqueName=%r", len(my_comments), my_display_name, my_unique_name)
            if not my_comments:
                logger.warning("No comments found for user %r on work item %s", my_display_name, self.work_item_id)
                self.done.emit("❌ No comments found for your account.")
                return

            latest_comment = sorted(my_comments, key=lambda x: x.get("createdDate", ""))[-1]
            self.progress.emit("Parsing comment...")

            iso_date = latest_comment.get("createdDate", "")
            logger.debug("Latest comment date (ISO): %s", iso_date)

            if iso_date:
                try:
                    # Handles timestamps WITH milliseconds
                    dt = datetime.strptime(iso_date, "%Y-%m-%dT%H:%M:%S.%fZ")
                except ValueError:
                    # Handles timestamps WITHOUT milliseconds
                    dt = datetime.strptime(iso_date, "%Y-%m-%dT%H:%M:%SZ")

                comment_date = dt.strftime("%m/%d/%Y")
            else:
                comment_date = ""

            logger.debug("Formatted comment date: %s", comment_date)

            raw_html = latest_comment.get("text", "")
            logger.debug("Raw HTML comment length: %d chars", len(raw_html))

            soup_for_text = BeautifulSoup(raw_html or "", "html.parser")
            clean_text = soup_for_text.get_text("\n", strip=True)

            # Known field labels (checked in order so longer names match first)
            known_labels = ["Root Cause", "Preventive Action", "Next Step", "Next Steps", "Status"]
            label_to_key = {
                "Root Cause": "Root Cause",
                "Preventive Action": "Preventive Action",
                "Next Step": "Next Step",
                "Next Steps": "Next Step",
                "Status": "Status",
            }

            def extract_all_fields(text):
                """Parse fields from comment text. Handles both
                'Label: value on same line' and 'Label:' on its own line
                with value on the next line."""
                lines = text.splitlines()
                fields = {}
                current_key = None
                current_value_lines = []

                def flush():
                    nonlocal current_key, current_value_lines
                    if current_key is not None:
                        fields[current_key] = "\n".join(current_value_lines).strip()
                    current_key = None
                    current_value_lines = []

                for line in lines:
                    stripped = line.strip()
                    matched = False
                    for label in known_labels:
                        after = stripped[len(label):]
                        if after.startswith(":"):
                            flush()
                            current_key = label_to_key[label]
                            remainder = after[1:].strip()
                            if remainder:
                                current_value_lines.append(remainder)
                            matched = True
                            break
                    if not matched and current_key is not None:
                        current_value_lines.append(stripped)

                flush()
                return fields

            fields = extract_all_fields(clean_text)
            root_cause = fields.get("Root Cause", "")
            preventive_action = fields.get("Preventive Action", "")
            next_step = fields.get("Next Step", "")
            status = fields.get("Status", "")

            logger.info("Extracted fields — Root Cause: %d chars, Preventive Action: %d chars, Next Step: %d chars, Status: %r",
                        len(root_cause), len(preventive_action), len(next_step), status)

            self.raw_html = raw_html
            self.root_cause = root_cause
            self.preventive_action = preventive_action
            self.next_step = next_step
            self.status = status

            # Anchored to the app directory, not the working directory.
            today_folder = os.path.join(REPORTS_DIR, datetime.now().strftime("%m-%d-%Y"))
            try:
                os.makedirs(today_folder, exist_ok=True)
            except OSError as e:
                # Without this the thread died silently and the Generate button
                # stayed disabled forever (the UI is only re-enabled by a
                # signal, and no signal used to be emitted here).
                logger.exception("Cannot create the report folder")
                self.done.emit(f"❌ Cannot create {today_folder}: {e}")
                return
            output_filename = (
                f"Ticket No. {_safe_filename(self.work_item_id)} - "
                f"Status Report for {_safe_filename(self.client_name)}.docx"
            )
            output_path = os.path.join(today_folder, output_filename)
            self.last_output_path = output_path
            logger.info("Output path: %s", output_path)

            self.progress.emit("Generating DOCX...")
            try:
                logger.info("Loading template: %s", self.template_path)
                if not os.path.isfile(self.template_path):
                    logger.error("Template file NOT FOUND: %s (cwd=%s)", self.template_path, os.getcwd())
                    self.done.emit(f"❌ Template not found: {self.template_path}")
                    return
                doc = Document(self.template_path)
                logger.info("Template loaded successfully (%d paragraph(s), %d table(s))", len(doc.paragraphs), len(doc.tables))
                placeholder_map = {
                    "{{TICKETNUM}}": str(ticket_number),
                    "{{DATE}}": comment_date,
                    "{{TICKETCONTENT}}": ticket_subject,
                    "{{RCA}}": root_cause,
                    "{{PREVAC}}": preventive_action,
                    "{{NEXT}}": next_step,
                    "{{CURSTAT}}": status,
                    "{{REPORT_CONTENT}}": raw_html,
                    "{{OS}}": self.location_type1,
                    "{{FS}}": self.location_type2,
                    "{{CLIENT}}": self.client_name,
                    "{{DISPLAYN}}": self.user_data["display_name"]  # new placeholder
                }

                # The HTML body is rendered run-by-run, so keep it out of the
                # plain-text pass.
                content_html = placeholder_map.pop("{{REPORT_CONTENT}}")

                # Shared engine: body paragraphs, table cells (nested included)
                # and text boxes are all patched. This used to walk doc.tables
                # only, so a placeholder moved out of a table silently shipped
                # as literal text.
                replace_placeholders_in_doc(doc, placeholder_map)
                insert_html_into_placeholder(doc, "{{REPORT_CONTENT}}", content_html, pat)

                doc.save(output_path)
                logger.info("Report saved successfully: %s", output_path)
                self.report_saved.emit(output_path)
                self.done.emit(f"✅ Report saved for ticket number: {self.work_item_id}")
            except Exception as e:
                logger.exception("Error generating report for work item %s", self.work_item_id)
                self.done.emit(f"❌ Error generating report: {e}")

class PotTicketDataWorker(QThread):
    """
    Fetches the latest comment (by the logged-in user) for a given
    TFS work item, so it can be inserted into {{TICKETDATA}}.
    """
    progress = pyqtSignal(str)
    done = pyqtSignal(str)       # emits raw_html on success
    error = pyqtSignal(str)      # emits error message on failure

    def __init__(self, work_item_id, user_data):
        super().__init__()
        self.work_item_id = work_item_id
        self.user_data = user_data
        self.raw_html = ""

    def run(self):
        logger.info("=== PotTicketDataWorker started for work item %s ===", self.work_item_id)

        pat = self.user_data["pat"]
        my_display_name = self.user_data["display_name"]
        my_unique_name = self.user_data.get("unique_name", "")

        self.progress.emit(f"Fetching comments for work item {self.work_item_id}...")
        comments_url = f"{org_url}/_apis/wit/workItems/{self.work_item_id}/comments?api-version=5.1-preview.3"

        try:
            logger.debug("GET %s", comments_url)
            resp = requests.get(comments_url, auth=HTTPBasicAuth("", pat), timeout=HTTP_TIMEOUT)
            resp.raise_for_status()
            comments_data = resp.json()
            total = len(comments_data.get("comments", []))
            logger.info("Fetched %d comment(s) for POT work item %s", total, self.work_item_id)
        except Exception as e:
            logger.exception("Failed to fetch comments for POT work item %s", self.work_item_id)
            self.error.emit(f"❌ Error fetching comments: {e}")
            return

        my_comments = [
            c for c in comments_data.get("comments", [])
            if c.get("createdBy", {}).get("displayName") == my_display_name
            or c.get("createdBy", {}).get("uniqueName") == my_unique_name
        ]
        logger.debug("Matched %d comment(s) for POT user %r", len(my_comments), my_display_name)

        if not my_comments:
            logger.warning("No comments found for user %r on POT work item %s", my_display_name, self.work_item_id)
            self.error.emit("❌ No comments found for your account on this ticket.")
            return

        latest_comment = sorted(my_comments, key=lambda x: x.get("createdDate", ""))[-1]
        self.raw_html = latest_comment.get("text", "")
        logger.info("POT comment fetched: %d chars of HTML", len(self.raw_html))

        self.done.emit(self.raw_html)


def _write_temp_image(image):
    """Save a QImage as a PNG in the temp folder; return the path ("" on failure).

    A picture that only exists inside the document cannot be uploaded: the
    worker on the host has no access to this PC, and clipboard data is gone
    once the dialog closes. A real file also survives a restart of the app.
    """
    if image is None or image.isNull():
        return ""
    directory = os.path.join(tempfile.gettempdir(), "AutoReport-pasted")
    try:
        os.makedirs(directory, exist_ok=True)
        name = f"pasted-{datetime.now().strftime('%Y%m%d-%H%M%S-%f')}.png"
        path = os.path.join(directory, name)
        if not image.save(path, "PNG"):
            logger.warning("Could not write a pasted image to %s", path)
            return ""
        return path
    except OSError:
        logger.exception("Could not store a pasted image")
        return ""


_CLIPBOARD_FILES = "AutoReport-clipboard"


def _clipboard_directory():
    """Where the pictures of a TFS copy are put for the paste target.

    TFS's comment editor ignores an <img> that arrives in pasted HTML, but it
    does upload picture *files* the way it uploads a pasted screenshot. So the
    files have to exist on disk while the clipboard holds them.
    """
    return os.path.join(tempfile.gettempdir(), _CLIPBOARD_FILES)


def _clear_clipboard_files():
    """Forget the previous copy: its files are no longer on the clipboard."""
    directory = _clipboard_directory()
    if not os.path.isdir(directory):
        return
    for name in os.listdir(directory):
        try:
            os.remove(os.path.join(directory, name))
        except OSError:
            logger.warning("Could not remove an old clipboard picture: %s", name)


_DRAFT_FILE = "draft.json"


def _draft_directory():
    """Where an unfinished scheduled comment waits between runs.

    %LOCALAPPDATA% rather than the temp folder: Windows is free to clean temp,
    and the whole point of the draft is to outlive whatever ended the app.
    """
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return os.path.join(base, "AutoReport", "scheduled-comment-draft")


def _clear_scheduled_draft():
    """Forget the local draft: it has been scheduled, or emptied."""
    directory = _draft_directory()
    if not os.path.isdir(directory):
        return
    try:
        shutil.rmtree(directory)
    except OSError:
        logger.warning("Could not clear the scheduled comment draft",
                       exc_info=True)


def _load_scheduled_draft():
    """The saved draft as a dict, or None when there is nothing to restore."""
    try:
        with open(os.path.join(_draft_directory(), _DRAFT_FILE),
                  encoding="utf-8") as handle:
            draft = json.load(handle)
    except (OSError, ValueError):
        return None
    return draft if isinstance(draft, dict) else None


class RemoteImageFetcher(QThread):
    """Download a picture referenced by pasted HTML so it can be displayed.

    A comment copied out of TFS carries <img src="https://tfs/..."> tags. Qt
    cannot fetch those (they need the reader's token), so it paints an empty
    box. Pulling the bytes down here fixes both halves of the problem: the
    picture shows in the editor, and it becomes a file that can be uploaded
    with the comment instead of a URL that may expire.
    """

    fetched = pyqtSignal(str, str)      # source URL, saved path ("" on failure)

    def __init__(self, url, pat="", parent=None):
        super().__init__(parent)
        self.url = url
        self.pat = pat

    def run(self):
        try:
            auth = None
            if self.pat:
                from urllib.parse import urlparse
                # Only hand the token to our own TFS; a picture hosted
                # anywhere else must not receive it.
                if (urlparse(self.url).netloc.lower()
                        == urlparse(org_url).netloc.lower()):
                    auth = HTTPBasicAuth("", self.pat)
            resp = requests.get(self.url, auth=auth, timeout=DOWNLOAD_TIMEOUT)
            if resp.status_code >= 400:
                logger.warning("Picture %s answered HTTP %s - left as a link",
                               self.url, resp.status_code)
                self.fetched.emit(self.url, "")
                return
            image = QtGui.QImage()
            if not image.loadFromData(resp.content):
                logger.warning("Picture %s is not a readable image", self.url)
                self.fetched.emit(self.url, "")
                return
            self.fetched.emit(self.url, _write_temp_image(image))
        except Exception:
            logger.exception("Could not download a pasted picture")
            self.fetched.emit(self.url, "")


class CommentEditor(QTextEdit):
    """The comment box: Word-style lists, character formatting, image paste.

    Three things a plain QTextEdit does not do, all of which the team expects
    from Word:

      * typing "*" (or "-" / "+") and then space at the start of a line
        starts a bulleted list, "1." and space a numbered one;
      * Ctrl+B / Ctrl+I / Ctrl+U format the selection, or whatever is typed
        next, and the result is sent to TFS as real <b>/<i>/<u>;
      * pasting a picture inserts it here and adds it to the comment's
        attachments, so it is uploaded instead of silently dropped;
      * a picture copied out of another ticket arrives as a URL, which Qt
        draws as an empty box - it is downloaded with your token first, so
        what you see is the picture itself.

    Pasted pictures are written to a temp PNG because an image embedded only in
    the document cannot be uploaded: the worker on the host has no access to
    this PC, and the clipboard data is gone once the dialog closes.
    """

    image_added = pyqtSignal(str)      # absolute path of a file to attach

    _BULLET_MARKERS = ("*", "-", "+", "\u2022")
    _NUMBER_MARKER = re.compile(r"\d{1,3}[.)]")

    def __init__(self, pat="", parent=None):
        super().__init__(parent)
        self.setAcceptRichText(True)
        self._pat = pat
        self._attached = []        # pasted paths, to avoid uploading twice
        self._fetchers = []        # running downloads, kept alive until done
        self._fetching = set()     # URLs already requested

    # --- images ---

    def canInsertFromMimeData(self, source):
        return source.hasImage() or super().canInsertFromMimeData(source)

    @staticmethod
    def _clipboard_image(source):
        """The clipboard's picture, however Qt hands it over."""
        data = source.imageData()
        if isinstance(data, QtGui.QImage):
            return data
        if isinstance(data, QtGui.QPixmap):
            return data.toImage()
        try:
            return QtGui.QImage(data)
        except TypeError:
            return QtGui.QImage()

    def insertFromMimeData(self, source):
        if source.hasImage():
            image = self._clipboard_image(source)
            path = _write_temp_image(image)
            if path and self._insert_local_image(path):
                if path not in self._attached:
                    self._attached.append(path)
                    self.image_added.emit(path)
                return
        super().insertFromMimeData(source)
        # A comment copied out of TFS arrives as HTML whose pictures are URLs,
        # not bytes: Qt draws them as empty boxes until they are fetched.
        if source.hasHtml():
            self._fetch_remote_images(source.html())

    @staticmethod
    def _remote_image_urls(html):
        urls = []
        for img in BeautifulSoup(html or "", "html.parser").find_all("img"):
            src = (img.get("src") or "").strip()
            if _COMMENT_IMAGE_URL.match(src) and src not in urls:
                urls.append(src)
        return urls

    def _fetch_remote_images(self, html):
        for url in self._remote_image_urls(html):
            if url in self._fetching:
                continue
            self._fetching.add(url)
            fetcher = RemoteImageFetcher(url, self._pat, self)
            fetcher.fetched.connect(self._on_image_fetched)
            fetcher.finished.connect(self._forget_fetcher)
            self._fetchers.append(fetcher)
            fetcher.start()

    def _forget_fetcher(self):
        sender = self.sender()
        if sender in self._fetchers:
            self._fetchers.remove(sender)

    def _on_image_fetched(self, url, path):
        """Swap a downloaded picture into the box where the empty box was."""
        self._fetching.discard(url)
        if not path:
            return          # the URL stays in the document and still posts
        if not self._replace_image(url, path):
            logger.warning("Downloaded %s but it is no longer in the comment", url)
            return
        if path not in self._attached:
            self._attached.append(path)
            self.image_added.emit(path)

    @staticmethod
    def _same_image(name, url):
        """Does this image fragment hold that URL? (Qt may normalise it.)"""
        name = (name or "").strip()
        url = (url or "").strip()
        return name == url or name.rstrip("/") == url.rstrip("/")

    def _replace_image(self, url, path):
        """Replace every fragment holding the URL with the downloaded file."""
        image = QtGui.QImage(path)
        if image.isNull():
            return False
        available = max(self.viewport().width() - 40, 160)
        if image.width() > available:
            image = image.scaledToWidth(
                available, QtCore.Qt.TransformationMode.SmoothTransformation
            )

        document = self.document()
        targets = []
        block = document.begin()
        while block.isValid():
            iterator = block.begin()
            while not iterator.atEnd():
                fragment = iterator.fragment()
                if fragment.isValid() and fragment.charFormat().isImageFormat():
                    name = fragment.charFormat().toImageFormat().name()
                    if self._same_image(name, url):
                        targets.append((fragment.position(), fragment.length()))
                iterator += 1
            block = block.next()

        # Back to front, so earlier positions stay valid.
        name = os.path.basename(path)
        for position, length in reversed(targets):
            cursor = QtGui.QTextCursor(document)
            cursor.setPosition(position)
            cursor.setPosition(position + length,
                               QtGui.QTextCursor.MoveMode.KeepAnchor)
            cursor.removeSelectedText()
            cursor.insertImage(image, name)
        return bool(targets)

    def _insert_local_image(self, path):
        """Show the picture in the editor, scaled to the box's width."""
        image = QtGui.QImage(path)
        if image.isNull():
            return False
        available = max(self.viewport().width() - 40, 160)
        if image.width() > available:
            image = image.scaledToWidth(
                available, QtCore.Qt.TransformationMode.SmoothTransformation
            )
        cursor = self.textCursor()
        cursor.insertImage(image, os.path.basename(path))
        self.setTextCursor(cursor)
        return True

    def restore_image(self, name, path):
        """Make a picture saved in the draft folder drawable again.

        toHtml() keeps only the image's file name, so a document reloaded from
        a draft has nothing to paint until the file is registered under that
        name, exactly as insertImage() does when the picture is pasted.
        """
        image = QtGui.QImage(path)
        if image.isNull():
            return False
        self.document().addResource(
            QtGui.QTextDocument.ResourceType.ImageResource, QUrl(name), image
        )
        return True

    def image_names(self):
        """File names of the pictures currently in the document.

        The dialog compares this against its attachment list: a picture
        deleted from the box must not be uploaded with the comment.
        """
        names = set()
        block = self.document().begin()
        while block.isValid():
            iterator = block.begin()
            while not iterator.atEnd():
                fragment = iterator.fragment()
                if fragment.isValid() and fragment.charFormat().isImageFormat():
                    name = fragment.charFormat().toImageFormat().name()
                    names.add(os.path.basename(name))
                iterator += 1
            block = block.next()
        return names

    def clear_attachments(self):
        """Forget pasted pictures and take them out of the document.

        Used when a saved schedule is cleared: a picture left visible in the
        box would look like it is still going to be uploaded with the next
        comment, so it is removed. The typed text is not touched.
        """
        self._attached = []
        document = self.document()
        block = document.begin()
        while block.isValid():
            images = []
            iterator = block.begin()
            while not iterator.atEnd():
                fragment = iterator.fragment()
                if fragment.isValid() and fragment.charFormat().isImageFormat():
                    images.append((fragment.position(), fragment.length()))
                iterator += 1
            # Back to front, so earlier positions stay valid.
            for position, length in reversed(images):
                cursor = QtGui.QTextCursor(document)
                cursor.setPosition(position)
                cursor.setPosition(position + length,
                                   QtGui.QTextCursor.MoveMode.KeepAnchor)
                cursor.removeSelectedText()
            block = block.next()

    # --- typing ---

    def keyPressEvent(self, event):
        if event.modifiers() == QtCore.Qt.KeyboardModifier.ControlModifier:
            key = event.key()
            if key == QtCore.Qt.Key.Key_B:
                self._toggle("bold")
                return
            if key == QtCore.Qt.Key.Key_I:
                self._toggle("italic")
                return
            if key == QtCore.Qt.Key.Key_U:
                self._toggle("underline")
                return
        if (event.key() == QtCore.Qt.Key.Key_Space
                and not self.textCursor().hasSelection()
                and self._autoformat_list()):
            return
        super().keyPressEvent(event)

    def _toggle(self, kind):
        """Word's Ctrl+B / Ctrl+I / Ctrl+U for the selection or the next word."""
        fmt = QtGui.QTextCharFormat()
        if kind == "bold":
            fmt.setFontWeight(400 if self.fontWeight() >= 700 else 700)
        elif kind == "italic":
            fmt.setFontItalic(not self.fontItalic())
        else:
            fmt.setFontUnderline(not self.fontUnderline())
        self.mergeCurrentCharFormat(fmt)
        self.setFocus()

    def _autoformat_list(self):
        """Turn a leading "* " / "- " / "1. " into a real list, Word-style."""
        cursor = self.textCursor()
        block_text = cursor.block().text()
        if cursor.positionInBlock() != len(block_text):
            return False           # only when the marker is the last thing typed
        marker = block_text.strip()
        if not marker:
            return False
        if marker in self._BULLET_MARKERS:
            style = QtGui.QTextListFormat.Style.ListDisc
        elif self._NUMBER_MARKER.fullmatch(marker):
            style = QtGui.QTextListFormat.Style.ListDecimal
        else:
            return False

        indent = max(len(block_text) - len(marker), 0)
        fmt = QtGui.QTextListFormat()
        fmt.setStyle(style)
        fmt.setIndent(1 + indent // 2)

        cursor.beginEditBlock()
        cursor.movePosition(QtGui.QTextCursor.MoveOperation.StartOfBlock)
        cursor.movePosition(QtGui.QTextCursor.MoveOperation.EndOfBlock,
                            QtGui.QTextCursor.MoveMode.KeepAnchor)
        cursor.removeSelectedText()
        cursor.createList(fmt)
        cursor.endEditBlock()
        self.setTextCursor(cursor)
        return True


class ScheduleCommentWorker(QThread):
    """Uploads any images, builds the comment HTML and stores the schedule.

    Everything that touches the network happens here: each image upload can
    take seconds, and the dialog has to stay responsive while it runs.
    """
    progress = pyqtSignal(str)
    saved = pyqtSignal(str)      # Firestore document id
    failed = pyqtSignal(str)

    def __init__(self, user_data, work_item_id, due_at, body, image_paths,
                 parent=None, body_html=""):
        super().__init__(parent)
        self.user_data = user_data
        self.work_item_id = work_item_id
        self.due_at = due_at
        self.body = body
        self.body_html = body_html
        self.image_paths = image_paths or []

    def run(self):
        try:
            pat = self.user_data.get("pat", "")
            if not pat:
                self.failed.emit("No TFS token on this PC. Sign in again to store one.")
                return

            # Keyed by the name the editor gives each picture, so _comment_html
            # can put it back where it was typed instead of at the end.
            uploads = {}
            image_urls = []
            for index, path in enumerate(self.image_paths, 1):
                self.progress.emit(f"Uploading image {index}/{len(self.image_paths)}...")
                url = _upload_comment_image(pat, self.work_item_id, path)
                image_urls.append(url)
                uploads[os.path.basename(path)] = url

            html = _comment_html(self.body, uploads, self.body_html)
            if not html.strip():
                self.failed.emit("The comment is empty.")
                return

            self.progress.emit("Saving the schedule...")

            # The object-replacement character marks where an inline picture
            # sits in the editor. It is meaningless in the text copy kept for
            # the list preview, where it shows up as a box.
            body_text = (self.body or "").replace("\ufffc", "").strip()

            values = {
                "owner_uid": self.user_data.get("uid", ""),
                "owner_email": self.user_data.get("email", ""),
                "work_item_id": str(self.work_item_id),
                "due_at": self.due_at,
                "body_text": body_text,
                "html": html,
                "image_urls": "\n".join(image_urls),
                "status": "pending",
                "created_at": _utc_now_iso(),
                "attempts": 0,
                "last_error": "",
                "posted_at": "",
            }
            if COPY_PAT_INTO_SCHEDULE:
                # The worker posts with this, so the comment is authored by the
                # person who scheduled it even when the worker is not allowed
                # to read that person's users/{uid} document.
                values["pat"] = pat

            token = _refresh_id_token(self.user_data) or self.user_data.get("id_token", "")

            # Keep users/{uid} in step too: with PAT_SYNC_TO_FIRESTORE on it is
            # the copy that lets this account sign in on another PC, and it is
            # what posts schedules saved by builds older than
            # COPY_PAT_INTO_SCHEDULE. A failure here no longer affects the
            # author, because the schedule carries its own token.
            if PAT_SYNC_TO_FIRESTORE and pat:
                if not _save_pat_remote(self.user_data.get("uid", ""), token, pat):
                    logger.warning("users/%s.pat could not be updated",
                                   self.user_data.get("uid", ""))

            doc_id = _fs_create(SCHEDULED_COMMENTS_COLLECTION, token, values)
            logger.info("Scheduled comment saved: doc=%s wi=%s due=%s",
                        doc_id, self.work_item_id, self.due_at)
            self.saved.emit(doc_id)
        except Exception as e:
            logger.exception("Could not save the scheduled comment")
            self.failed.emit(str(e))


class ScheduledCommentsListWorker(QThread):
    """Reads this user's schedules back from Firestore."""
    loaded = pyqtSignal(object)   # list of dicts
    failed = pyqtSignal(str)

    def __init__(self, user_data, parent=None):
        super().__init__(parent)
        self.user_data = user_data

    def run(self):
        try:
            uid = self.user_data.get("uid", "")
            if not uid:
                self.failed.emit("Not signed in.")
                return
            token = _refresh_id_token(self.user_data) or self.user_data.get("id_token", "")
            documents = _fs_query(
                SCHEDULED_COMMENTS_COLLECTION, token,
                field="owner_uid", value=uid, limit=200,
            )
            documents.sort(key=lambda d: str(d.get("due_at", "")), reverse=True)
            self.loaded.emit(documents)
        except Exception as e:
            logger.exception("Could not load scheduled comments")
            self.failed.emit(str(e))


class DeleteSchedulesWorker(QThread):
    """Deletes a batch of schedules, one request at a time.

    The deletes used to run on the GUI thread, so a slow Firestore call froze
    the dialog - indistinguishable from the app hanging. Here the requests run
    off the GUI thread and every row reports back, so the panel can show that
    work is happening rather than looking stuck.
    """

    progress = pyqtSignal(int, int)          # done, total
    done = pyqtSignal(object, object)        # deleted docs, [(doc, error)]

    def __init__(self, documents, user_data, parent=None):
        super().__init__(parent)
        self.documents = list(documents)
        self.user_data = dict(user_data or {})

    def run(self):
        token = (_refresh_id_token(self.user_data)
                 or self.user_data.get("id_token", ""))
        deleted = []
        failed = []
        total = len(self.documents)
        for index, doc in enumerate(self.documents, start=1):
            try:
                _fs_delete(SCHEDULED_COMMENTS_COLLECTION, doc.get("id", ""), token)
                deleted.append(doc)
                logger.info("Scheduled comment deleted: %s (status=%s)",
                            doc.get("id"), doc.get("status") or "?")
            except Exception as e:
                logger.exception("Could not delete the schedule %s", doc.get("id"))
                failed.append((doc, str(e)))
            self.progress.emit(index, total)
        self.done.emit(deleted, failed)


def _comment_plain_text(html):
    """Readable text for the plain-text half of the clipboard.

    Pictures are dropped (they ride along in the HTML half), list items become
    "\u2022" lines and <br> becomes a newline, so a paste into a plain text field
    still reads like the comment.
    """
    soup = BeautifulSoup(html or "", "html.parser")
    for img in soup.find_all("img"):
        img.decompose()
    for br in soup.find_all("br"):
        br.replace_with("\n")
    for item in soup.find_all("li"):
        item.insert_before("\u2022 ")
    text = soup.get_text()
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line)


class CopyCommentWorker(QThread):
    """Puts a saved comment on the clipboard, pictures and all.

    Firestore keeps only the comment HTML and the URLs of its attachments, and
    which of those the paste target understands decides the mode:

    * ``embed_pictures=True`` downloads each picture (with the same PAT the
      editor uses) and writes it into the HTML as a data URI, so the clipboard
      is self-contained - Word, Docs and chat windows get the pixels.
    * ``embed_pictures=False`` leaves the attachment URLs alone and also
      writes the pictures to disk so they can ride on the clipboard as files.
      The TFS comment editor drops every ``<img>`` out of pasted HTML, but it
      uploads picture files the way it uploads a pasted screenshot.

    Read-only: it never writes to Firestore or TFS, so a pending schedule is
    untouched.
    """

    # html, text, pictures, missing, files (paths for TFS to upload)
    ready = pyqtSignal(str, str, int, int, list)
    failed = pyqtSignal(str)

    _MIME_BY_FORMAT = {
        "PNG": "image/png", "JPEG": "image/jpeg", "GIF": "image/gif",
        "BMP": "image/bmp", "TIFF": "image/tiff", "WEBP": "image/webp",
    }

    def __init__(self, document, pat, embed_pictures=True, parent=None):
        super().__init__(parent)
        self.document = dict(document or {})
        self.pat = pat or ""
        self.embed_pictures = embed_pictures

    def _html(self):
        html = str(self.document.get("html", "") or "").strip()
        if html:
            return html
        # Schedules saved before rich text only have the plain body.
        text = str(self.document.get("body_text", "") or "").strip()
        if not text:
            raise RuntimeError("That schedule has no comment to copy.")
        return "<br>".join(
            _linkify_escaped(_escape_comment_text(line))
            for line in re.split(r"\r\n|\r|\n", text)
        )

    @classmethod
    def _mime(cls, payload):
        try:
            with Image.open(BytesIO(payload)) as probe:
                return cls._MIME_BY_FORMAT.get(
                    (probe.format or "").upper(), "image/png"
                )
        except Exception:
            return "image/png"

    def _stage_files(self, images):
        """Download each picture to a file and return the paths that worked."""
        _clear_clipboard_files()
        directory = _clipboard_directory()
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError:
            logger.exception("Could not create %s", directory)
            return []

        extensions = {
            "image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif",
            "image/bmp": ".bmp", "image/tiff": ".tiff",
            "image/webp": ".webp",
        }
        paths = []
        for index, image in enumerate(images, 1):
            src = (image.get("src") or "").strip()
            if not src or src.startswith("data:"):
                continue
            try:
                payload = _fetch_image_bytes(src, self.pat)
                suffix = extensions.get(self._mime(payload), ".png")
                path = os.path.join(directory, f"picture-{index}{suffix}")
                with open(path, "wb") as handle:
                    handle.write(payload)
                paths.append(path)
            except Exception as e:
                logger.warning(
                    "Could not stage a picture for the TFS clipboard: %s", e
                )
        return paths

    def run(self):
        try:
            html = self._html()
            soup = BeautifulSoup(html, "html.parser")
            images = soup.find_all("img")
            if not images:
                # A comment saved without rich text still lists its uploads.
                container = soup.body or soup
                for url in str(self.document.get("image_urls", "")).splitlines():
                    if url.strip():
                        image = soup.new_tag("img", src=url.strip())
                        container.append(image)
                        images.append(image)

            if not self.embed_pictures:
                # The stored attachment URLs are kept as they are, and the
                # pictures are also staged as real files: TFS's comment editor
                # drops <img> out of pasted HTML but uploads picture files the
                # way it uploads a pasted screenshot.
                files = self._stage_files(images)
                copied = str(soup)
                self.ready.emit(
                    copied, _comment_plain_text(copied), len(images),
                    max(len(images) - len(files), 0), files,
                )
                return

            embedded = 0
            missing = 0
            for image in images:
                src = (image.get("src") or "").strip()
                if not src or src.startswith("data:"):
                    continue
                try:
                    payload = _fetch_image_bytes(src, self.pat)
                    encoded = base64.b64encode(payload).decode("ascii")
                    image["src"] = (
                        f"data:{self._mime(payload)};base64,{encoded}"
                    )
                    embedded += 1
                except Exception as e:
                    # The URL is left in place: the comment still copies, and
                    # the original attachment keeps working.
                    missing += 1
                    logger.warning(
                        "Could not fetch a picture for the clipboard: %s", e
                    )

            copied = str(soup)
            self.ready.emit(
                copied, _comment_plain_text(copied), embedded, missing, []
            )
        except Exception as e:
            logger.exception("Could not copy the scheduled comment")
            self.failed.emit(str(e))


class ScheduledCommentDialog(QDialog):
    """Compose a TFS comment now, have it posted at a chosen time.

    The comment is stored in Firestore; scheduled_comment_worker.py on the
    host posts it. Images go up as work item attachments straight away,
    because the host has no access to this PC's disk.
    """

    def __init__(self, user_data, parent=None):
        super().__init__(parent)
        self.user_data = dict(user_data or {})
        self.image_paths = []
        self._documents = []
        self._save_worker = None
        self._list_worker = None
        self._copy_worker = None
        self._delete_worker = None
        self._copy_embed = True

        # An unfinished comment is kept on disk: a mis-clicked close, a crash
        # or a Windows update must not cost the text and the pictures that are
        # already in the box. See _save_draft / _restore_draft.
        self._draft_timer = QtCore.QTimer(self)
        self._draft_timer.setSingleShot(True)
        self._draft_timer.setInterval(700)
        self._draft_timer.timeout.connect(self._save_draft)

        self.setWindowTitle("Scheduled Comment")
        self.setObjectName("ScheduledComment")
        self.setMinimumSize(980, 620)
        self.setStyleSheet(self._style())

        # Compose on the left, saved schedules down the right-hand side: the
        # list is a reference you glance at while writing, and its rows carry
        # their own right-click actions.
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(18)

        compose = QtWidgets.QVBoxLayout()
        compose.setSpacing(10)

        title = QtWidgets.QLabel("Schedule a TFS comment")
        title.setStyleSheet("font-size: 20px; font-weight: 700;")
        compose.addWidget(title)

        subtitle = QtWidgets.QLabel(
            "Saved in the cloud now; the hosted worker posts it to the work item "
            "at the time you pick."
        )
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("font-size: 12px; color: #6b7280;")
        compose.addWidget(subtitle)

        form = QtWidgets.QFormLayout()
        form.setSpacing(8)
        form.setContentsMargins(0, 6, 0, 6)

        self.work_item_input = QLineEdit()
        self.work_item_input.setPlaceholderText("e.g. 123456")
        self.work_item_input.setMaximumWidth(240)
        self.work_item_input.textChanged.connect(self._schedule_draft_save)
        form.addRow("Work Item ID:", self.work_item_input)

        self.when_input = QtWidgets.QDateTimeEdit(
            QtCore.QDateTime.currentDateTime().addSecs(600)
        )
        self.when_input.setCalendarPopup(True)
        self.when_input.setDisplayFormat("MM/dd/yyyy hh:mm AP")
        self.when_input.setMaximumWidth(240)
        self.when_input.dateTimeChanged.connect(self._schedule_draft_save)
        form.addRow("Post at (local time):", self.when_input)

        compose.addLayout(form)

        comment_label = QtWidgets.QLabel("Comment")
        comment_label.setObjectName("sectionLabel")
        compose.addWidget(comment_label)

        self.body_input = CommentEditor(pat=self.user_data.get("pat", ""))
        self.body_input.setPlaceholderText(
            "Write the comment. '*' then space starts a bullet, '1.' then space a "
            "numbered list, Ctrl+B / Ctrl+I / Ctrl+U set bold, italic and "
            "underline, and a picture can be pasted straight in."
        )
        self.body_input.image_added.connect(self._on_pasted_image)
        self.body_input.textChanged.connect(self._on_editor_changed)
        compose.addWidget(self.body_input, stretch=1)

        editor_hint = QtWidgets.QLabel(
            "Ctrl+B bold  ·  Ctrl+I italic  ·  Ctrl+U underline  ·  "
            "* then space = bullet  ·  1. then space = numbered"
        )
        editor_hint.setWordWrap(True)
        editor_hint.setStyleSheet("font-size: 11px; color: #6b7280;")
        compose.addWidget(editor_hint)

        # Pictures are pasted into the box, not picked from a dialog: the box is
        # the only place a picture can be positioned among the text.
        self.images_label = QtWidgets.QLabel("No pictures pasted")
        self.images_label.setStyleSheet("font-size: 12px; color: #6b7280;")
        compose.addWidget(self.images_label)

        self.author_label = QtWidgets.QLabel(
            "This comment is posted as you: the worker uses the TFS token "
            "stored with your account."
        )
        self.author_label.setWordWrap(True)
        self.author_label.setStyleSheet("font-size: 12px; color: #6b7280;")
        compose.addWidget(self.author_label)

        buttons = QtWidgets.QHBoxLayout()
        buttons.setContentsMargins(0, 6, 0, 0)
        self.schedule_btn = QPushButton("Schedule Comment")
        self.schedule_btn.clicked.connect(self.schedule)
        close_btn = QPushButton("Close")
        close_btn.setObjectName("ghost")
        close_btn.clicked.connect(self.reject)
        buttons.addWidget(self.schedule_btn)
        buttons.addStretch()
        buttons.addWidget(close_btn)
        compose.addLayout(buttons)

        self.status_label = QtWidgets.QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("font-size: 12px; color: #6b7280;")
        compose.addWidget(self.status_label)

        layout.addLayout(compose, stretch=1)

        # --- saved schedules, as a panel down the right-hand side ---
        side = QtWidgets.QFrame()
        side.setObjectName("sidePanel")
        side.setFixedWidth(320)
        side_layout = QtWidgets.QVBoxLayout(side)
        side_layout.setContentsMargins(16, 16, 16, 16)
        side_layout.setSpacing(10)

        side_header = QtWidgets.QHBoxLayout()
        side_title = QtWidgets.QLabel("Scheduled")
        side_title.setObjectName("sideTitle")
        side_header.addWidget(side_title)
        side_header.addStretch()
        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.setObjectName("ghost")
        self.refresh_btn.clicked.connect(self.refresh_list)
        side_header.addWidget(self.refresh_btn)
        side_layout.addLayout(side_header)

        self.list_view = QtWidgets.QListWidget()
        self.list_view.setObjectName("scheduleList")
        # Shift-click selects a range and Ctrl-click picks rows out of one, so
        # a batch of old schedules can be deleted in a single go.
        self.list_view.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection
        )
        # Copy and delete live on the rows themselves: right-click a schedule.
        self.list_view.setContextMenuPolicy(
            QtCore.Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.list_view.customContextMenuRequested.connect(
            self._show_schedule_menu
        )
        side_layout.addWidget(self.list_view, stretch=1)

        # Runs while a copy, delete or refresh is in flight: a few Firestore
        # round-trips look like a frozen window otherwise.
        self.busy_bar = QtWidgets.QProgressBar()
        self.busy_bar.setObjectName("busyBar")
        self.busy_bar.setRange(0, 0)          # indeterminate: no known length
        self.busy_bar.setTextVisible(False)
        self.busy_bar.setFixedHeight(4)
        self.busy_bar.setVisible(False)
        side_layout.addWidget(self.busy_bar)

        # List messages belong to the list, not to the compose area.
        self.list_status = QtWidgets.QLabel("")
        self.list_status.setWordWrap(True)
        self.list_status.setStyleSheet("font-size: 11.5px; color: #6b7280;")
        side_layout.addWidget(self.list_status)

        hint = QtWidgets.QLabel(
            "Right-click a schedule to copy or delete it. Shift-click or "
            "Ctrl-click to select several."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("font-size: 11.5px; color: #6b7280;")
        side_layout.addWidget(hint)

        layout.addWidget(side)

        self.refresh_list()
        self._restore_draft()
        self.showMaximized()

    def _style(self):
        """Same dark/light treatment as the rest of the app."""
        if is_windows_dark_mode():
            return """
            QDialog#ScheduledComment {
                background: #0f1923; color: #e6eef5;
                font-family: Inter, Segoe UI, Arial;
            }
            QLabel { color: #e6eef5; }
            QLabel#sectionLabel {
                font-size: 11px; font-weight: 700; color: #98a0a6;
            }
            QLabel#sideTitle { font-size: 14px; font-weight: 700; }
            QFrame#sidePanel {
                background: rgba(255,255,255,0.03);
                border: 1px solid rgba(255,255,255,0.08);
                border-radius: 14px;
            }
            QLineEdit, QTextEdit, QDateTimeEdit, QListWidget {
                background: rgba(255,255,255,0.04);
                border: 1px solid rgba(255,255,255,0.08);
                border-radius: 10px; padding: 8px;
                color: #e6eef5; font-size: 13px;
            }
            QLineEdit:focus, QTextEdit:focus, QDateTimeEdit:focus {
                border: 1px solid #ed3b3b;
            }
            QListWidget#scheduleList {
                background: transparent; border: none; padding: 0px;
                outline: none;
            }
            QListWidget#scheduleList::item {
                padding: 7px 10px;
                border-radius: 8px;
                border-left: 3px solid transparent;
                color: #c9d4de;
            }
            QListWidget#scheduleList::item:hover {
                background: rgba(255,255,255,0.06);
            }
            QListWidget#scheduleList::item:selected {
                border-left: 3px solid #ed3b3b;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 rgba(237,59,59,0.34),
                    stop:1 rgba(237,59,59,0.06));
                color: #ffffff;
            }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white; border-radius: 10px;
                padding: 8px 16px; font-size: 13px; font-weight: 600;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton#ghost {
                background: transparent;
                border: 1px solid rgba(255,255,255,0.16);
                color: #e6eef5; padding: 7px 14px;
            }
            QPushButton#ghost:hover { border-color: #ed3b3b; color: #ffffff; }
            QPushButton:disabled {
                background: rgba(255,255,255,0.08); color: #98a0a6;
                border-color: transparent;
            }
            QPushButton#ghost:disabled {
                background: transparent; color: #98a0a6;
                border-color: rgba(255,255,255,0.08);
            }
            QProgressBar {
                background: rgba(255,255,255,0.08); border: none;
                border-radius: 2px;
            }
            QProgressBar::chunk { background: #ed3b3b; border-radius: 2px; }
            QMenu {
                background: #16232f; color: #e6eef5;
                border: 1px solid rgba(255,255,255,0.12);
                border-radius: 10px; padding: 6px;
            }
            QMenu::item { padding: 7px 18px; border-radius: 7px; }
            QMenu::item:selected {
                background: rgba(237,59,59,0.25); color: #ffffff;
            }
            QMenu::separator {
                height: 1px; background: rgba(255,255,255,0.10);
                margin: 5px 8px;
            }
            """
        return """
        QDialog#ScheduledComment {
            background: #f5f7fa; color: #1f2933;
            font-family: Inter, Segoe UI, Arial;
        }
        QLabel { color: #111827; }
        QLabel#sectionLabel {
            font-size: 11px; font-weight: 700; color: #6b7280;
        }
        QLabel#sideTitle { font-size: 14px; font-weight: 700; }
        QFrame#sidePanel {
            background: #ffffff; border: 1px solid #e5e7eb;
            border-radius: 14px;
        }
        QLineEdit, QTextEdit, QDateTimeEdit, QListWidget {
            background: white; border: 1px solid #d1d5db;
            border-radius: 10px; padding: 8px;
            color: #111827; font-size: 13px;
        }
        QLineEdit:focus, QTextEdit:focus, QDateTimeEdit:focus {
            border: 1px solid #ed3b3b;
        }
        QListWidget#scheduleList {
            background: transparent; border: none; padding: 0px;
            outline: none;
        }
        QListWidget#scheduleList::item {
            padding: 7px 10px;
            border-radius: 8px;
            border-left: 3px solid transparent;
            color: #374151;
        }
        QListWidget#scheduleList::item:hover {
            background: rgba(17,24,39,0.05);
        }
        QListWidget#scheduleList::item:selected {
            border-left: 3px solid #ed3b3b;
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(237,59,59,0.20),
                stop:1 rgba(237,59,59,0.03));
            color: #111827;
        }
        QPushButton {
            background: #ed3b3b; color: white; border-radius: 10px;
            padding: 8px 16px; font-size: 13px; font-weight: 600;
        }
        QPushButton:hover { background: #c62e2e; }
        QPushButton#ghost {
            background: transparent; border: 1px solid #d1d5db;
            color: #374151; padding: 7px 14px;
        }
        QPushButton#ghost:hover { border-color: #ed3b3b; color: #ed3b3b; }
        QPushButton:disabled { background: #d1d5db; color: #6b7280; }
        QPushButton#ghost:disabled {
            background: transparent; color: #9ca3af; border-color: #e5e7eb;
        }
        QProgressBar {
            background: #e5e7eb; border: none; border-radius: 2px;
        }
        QProgressBar::chunk { background: #ed3b3b; border-radius: 2px; }
        QMenu {
            background: #ffffff; color: #111827;
            border: 1px solid #e5e7eb;
            border-radius: 10px; padding: 6px;
        }
        QMenu::item { padding: 7px 18px; border-radius: 7px; }
        QMenu::item:selected {
            background: rgba(237,59,59,0.12); color: #111827;
        }
        QMenu::separator {
            height: 1px; background: #e5e7eb; margin: 5px 8px;
        }
        """

    # --- compose ---

    def clear_images(self):
        self.image_paths = []
        if _ui_alive(self.body_input):
            self.body_input.clear_attachments()
        self._update_images_label()

    def _on_pasted_image(self, path):
        """A picture pasted into the box is an attachment like any other."""
        if path and path not in self.image_paths:
            self.image_paths.append(path)
            self._update_images_label()
            self._schedule_draft_save()

    def _on_editor_changed(self):
        """Re-check the attachments once the edit that caused this is done.

        Replacing a downloaded picture briefly removes it from the document,
        so pruning immediately could drop an attachment that is still there.
        """
        QtCore.QTimer.singleShot(0, self._sync_pasted_images)
        self._schedule_draft_save()

    def _sync_pasted_images(self):
        """Keep the attachments to the pictures still in the box.

        There is no "Clear images" button any more, so deleting a picture is
        how it is dropped: leaving its file attached would upload a picture the
        comment no longer shows.
        """
        if not self.image_paths or not _ui_alive(self.body_input):
            return
        present = self.body_input.image_names()
        kept = [p for p in self.image_paths if os.path.basename(p) in present]
        if len(kept) != len(self.image_paths):
            self.image_paths = kept
            self._update_images_label()
            self._schedule_draft_save()

    # --- local draft ---

    def _schedule_draft_save(self):
        """Save shortly after the typing stops, not on every keystroke."""
        if _ui_alive(self.body_input):
            self._draft_timer.start()

    def _save_draft(self):
        """Keep a local copy of the comment that has not been scheduled yet.

        Until Schedule Comment is pressed Firestore has never seen this text,
        so this folder is the only copy: the editor HTML plus a real copy of
        every picture, because the ones in the temp folder can be cleaned up
        by Windows while the app is closed.
        """
        if not _ui_alive(self.body_input):
            return
        try:
            html = self.body_input.toHtml()
            names = self.body_input.image_names()
            paths = [p for p in self.image_paths
                     if os.path.basename(p) in names and os.path.exists(p)]

            # An empty box means there is nothing to come back to.
            if not self.body_input.toPlainText().strip() and not paths:
                _clear_scheduled_draft()
                return

            directory = _draft_directory()
            os.makedirs(directory, exist_ok=True)
            kept = set()
            for path in paths:
                name = os.path.basename(path)
                target = os.path.join(directory, name)
                kept.add(name)
                if (os.path.abspath(path) != os.path.abspath(target)
                        and not os.path.exists(target)):
                    shutil.copy2(path, target)

            # Pictures deleted from the box, or left behind by a save that
            # never finished, must not pile up in the draft folder.
            for name in os.listdir(directory):
                if name == _DRAFT_FILE or name in kept:
                    continue
                try:
                    os.remove(os.path.join(directory, name))
                except OSError:
                    pass

            draft = {
                "body_html": html,
                "work_item_id": self.work_item_input.text().strip(),
                "when": self.when_input.dateTime().toString(
                    "yyyy-MM-ddTHH:mm:ss"
                ),
                "images": sorted(kept),
                "saved_at": _utc_now_iso(),
            }
            target = os.path.join(directory, _DRAFT_FILE)
            temporary = target + ".tmp"
            with open(temporary, "w", encoding="utf-8") as handle:
                json.dump(draft, handle)
            os.replace(temporary, target)
        except OSError:
            logger.warning("Could not save the scheduled comment draft",
                           exc_info=True)

    def _restore_draft(self):
        """Bring back the comment the previous session never finished."""
        draft = _load_scheduled_draft()
        if not draft:
            return
        html = str(draft.get("body_html", "") or "").strip()
        if not html:
            return

        # setHtml() alone cannot draw the pictures: it only knows their file
        # names, so each one is loaded from the draft folder under that name.
        self.body_input.setHtml(html)
        directory = _draft_directory()
        paths = []
        for name in self.body_input.image_names():
            candidate = os.path.join(directory, name)
            if self.body_input.restore_image(name, candidate):
                paths.append(candidate)
        self.image_paths = paths
        self._update_images_label()

        work_item_id = str(draft.get("work_item_id", "") or "").strip()
        if work_item_id:
            self.work_item_input.setText(work_item_id)

        # A draft left overnight would fail "pick a time in the future", so a
        # time that has passed moves to now + 10 minutes.
        when = QtCore.QDateTime.fromString(
            str(draft.get("when", "") or ""), "yyyy-MM-ddTHH:mm:ss"
        )
        if when.isValid() and when < QtCore.QDateTime.currentDateTime():
            when = QtCore.QDateTime.currentDateTime().addSecs(600)
        if when.isValid():
            self.when_input.setDateTime(when)

        self.status_label.setText(
            "Restored your unfinished comment and its pictures from last time."
        )
        logger.info("Restored the scheduled comment draft (%d picture(s))",
                    len(paths))

    def _update_images_label(self):
        if not self.image_paths:
            self.images_label.setText("No pictures pasted")
            return
        names = ", ".join(os.path.basename(p) for p in self.image_paths[:3])
        extra = ""
        if len(self.image_paths) > 3:
            extra = f" +{len(self.image_paths) - 3} more"
        self.images_label.setText(
            f"{len(self.image_paths)} picture(s): {names}{extra}"
        )

    def schedule(self):
        work_item_id = self.work_item_input.text().strip()
        body = self.body_input.toPlainText().strip()

        if not work_item_id.isdigit():
            AppDialog("Validation Error", "Enter a numeric work item ID.", self).exec()
            return
        if not body and not self.image_paths:
            AppDialog("Validation Error", "Write a comment or paste a picture first.", self).exec()
            return
        if not self.user_data.get("pat"):
            AppDialog("Missing Token", "No TFS token on this PC. Sign in again to store one.", self).exec()
            return

        # The picker is local time; Firestore stores UTC so the worker can
        # compare against its own clock wherever it runs.
        local_text = self.when_input.dateTime().toString("yyyy-MM-ddTHH:mm:ss")
        local_dt = datetime.strptime(local_text, "%Y-%m-%dT%H:%M:%S")
        if local_dt < datetime.now():
            AppDialog("Validation Error", "Pick a time in the future.", self).exec()
            return
        due_utc = local_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        self._set_busy(True)
        self.status_label.setText("Uploading and saving...")

        self._save_worker = ScheduleCommentWorker(
            self.user_data, work_item_id, due_utc, body,
            list(self.image_paths), self,
            body_html=self.body_input.toHtml(),
        )
        self._save_worker.progress.connect(self.status_label.setText)
        self._save_worker.saved.connect(self._on_saved)
        self._save_worker.failed.connect(self._on_failed)
        self._save_worker.start()

    def _set_busy(self, busy):
        for widget in (self.schedule_btn, self.refresh_btn):
            if _ui_alive(widget):
                widget.setEnabled(not busy)
        if _ui_alive(self.schedule_btn):
            self.schedule_btn.setText("Saving..." if busy else "Schedule Comment")

    def _on_saved(self, doc_id):
        self._set_busy(False)
        logger.info("Schedule stored: %s", doc_id)
        self.status_label.setText(
            "Scheduled. The worker will post it to the work item at that time."
        )
        self.body_input.clear()
        self.clear_images()
        _clear_scheduled_draft()
        self.refresh_list()

    def _on_failed(self, message):
        self._set_busy(False)
        logger.warning("Scheduling failed: %s", message)
        self.status_label.setText(f"Could not schedule: {message}")

    # --- saved schedules ---

    def _set_list_busy(self, busy, message=None):
        """Show that the panel is working, not hanging.

        Copying a comment (which may download pictures), deleting rows and
        loading the list all talk to Firestore, so each one can take a few
        seconds. The bar runs meanwhile and the list and Refresh button are
        locked: a second action started mid-flight would fight the first for
        the same rows.
        """
        if _ui_alive(self.busy_bar):
            self.busy_bar.setVisible(busy)
        for widget in (self.list_view, self.refresh_btn):
            if _ui_alive(widget):
                widget.setEnabled(not busy)
        if message is not None and _ui_alive(self.list_status):
            self.list_status.setText(message)

    def refresh_list(self):
        if not self.user_data.get("uid"):
            self._set_list_busy(False, "Not signed in.")
            return
        self._set_list_busy(True, "Loading schedules...")
        self._list_worker = ScheduledCommentsListWorker(self.user_data, self)
        self._list_worker.loaded.connect(self._on_list_loaded)
        self._list_worker.failed.connect(self._on_list_failed)
        self._list_worker.start()

    def _on_list_loaded(self, documents):
        self._documents = documents or []
        self.list_view.clear()
        for doc in self._documents:
            item = QtWidgets.QListWidgetItem(self._summary(doc))
            item.setSizeHint(QtCore.QSize(0, 52))
            self.list_view.addItem(item)
        if not self._documents:
            note = "No scheduled comments yet."
        else:
            note = f"{len(self._documents)} schedule(s)"
        self._set_list_busy(False, note)

    def _on_list_failed(self, message):
        logger.warning("Could not load schedules: %s", message)
        self._set_list_busy(False, f"Could not load schedules: {message}")

    def _summary(self, doc):
        """Ticket number and the time it goes out - nothing else.

        A body preview made every row hard to scan; the panel is for spotting
        which ticket is queued and when, not for reading the comments.
        """
        return f"WI {doc.get('work_item_id', '?')}\n{self._due_text(doc)}"

    def _due_text(self, doc):
        """The schedule's due time in local time, as shown in the list."""
        due = str(doc.get("due_at", ""))
        try:
            local = datetime.strptime(due, "%Y-%m-%dT%H:%M:%SZ").replace(
                tzinfo=timezone.utc
            ).astimezone()
            return local.strftime("%m/%d/%Y %I:%M %p")
        except (TypeError, ValueError):
            return due or "?"

    def _is_future(self, doc):
        """True while the schedule's due time has not arrived yet.

        Unreadable timestamps are not treated as future: they cannot be
        posted, so holding on to them would leave a row nobody can remove.
        """
        try:
            due = datetime.strptime(
                str(doc.get("due_at", "")), "%Y-%m-%dT%H:%M:%SZ"
            ).replace(tzinfo=timezone.utc)
        except (TypeError, ValueError):
            return False
        return due >= datetime.now(timezone.utc)

    def _selected_docs(self):
        """The schedules behind the highlighted rows, in list order."""
        rows = sorted(self.list_view.row(item)
                      for item in self.list_view.selectedItems())
        return [self._documents[r] for r in rows
                if 0 <= r < len(self._documents)]

    def _show_schedule_menu(self, pos):
        """Right-click menu for the highlighted schedules.

        A right-click inside an existing highlight keeps it, so one row can be
        dropped from a batch; anywhere else selects just the row under the
        cursor, which is what a right-click on a fresh row means.
        """
        item = self.list_view.itemAt(pos)
        if item is None:
            return
        if not item.isSelected():
            self.list_view.clearSelection()
            item.setSelected(True)
            self.list_view.setCurrentItem(item)

        count = len(self._selected_docs())

        menu = QtWidgets.QMenu(self.list_view)
        menu.setObjectName("scheduleMenu")
        copy_tfs = menu.addAction("Copy for TFS")
        copy_tfs.setToolTip(
            "Paste into the TFS comment box: pictures go on the clipboard as "
            "files, which is the only image form that editor keeps."
        )
        copy_general = menu.addAction("Copy for general")
        copy_general.setToolTip(
            "Paste into Word, Docs or a chat window: pictures are embedded in "
            "the copied HTML."
        )
        # The clipboard holds one comment, so copying needs exactly one row.
        for action in (copy_tfs, copy_general):
            action.setEnabled(count == 1)
        menu.addSeparator()
        delete = menu.addAction(
            "Delete schedule" if count <= 1 else f"Delete {count} schedules"
        )

        chosen = menu.exec(self.list_view.viewport().mapToGlobal(pos))
        if chosen is copy_tfs:
            self.copy_selected(embed_pictures=False)
        elif chosen is copy_general:
            self.copy_selected(embed_pictures=True)
        elif chosen is delete:
            self.cancel_selected()

    def copy_selected(self, embed_pictures=True):
        """Put the selected schedule's comment on the clipboard.

        Two flavours, because paste targets disagree about pictures:

        * embed_pictures=False - leave the attachment URLs the comment was
          saved with and stage the pictures as files. The TFS comment editor
          ignores <img> in pasted HTML and strips data URIs, but it uploads
          picture files the way it uploads a pasted screenshot, so this is the
          one that lands in a ticket with its pictures.
        * embed_pictures=True - download the pictures and inline them, so the
          clipboard is self-contained for Word, Docs, chat and so on.

        Both only read: no schedule document is written. One row only - the
        clipboard has room for a single comment.
        """
        if self._copy_worker is not None and self._copy_worker.isRunning():
            self.list_status.setText("Still copying - give it a second.")
            return

        docs = self._selected_docs()
        if len(docs) != 1:
            self.list_status.setText(
                "Select one schedule to copy." if docs
                else "Right-click a schedule to copy it."
            )
            return

        self._copy_embed = embed_pictures
        self._set_list_busy(
            True,
            "Copying the comment and its pictures..." if embed_pictures
            else "Copying the comment and preparing its pictures...",
        )

        self._copy_worker = CopyCommentWorker(
            docs[0], self.user_data.get("pat", ""),
            embed_pictures, self
        )
        self._copy_worker.ready.connect(self._on_copied)
        self._copy_worker.failed.connect(self._on_copy_failed)
        self._copy_worker.start()

    def _on_copied(self, html, text, pictures, missing, files):
        mime = QtCore.QMimeData()
        mime.setHtml(html)
        mime.setText(text)
        if files:
            # The file list is what lets the TFS comment box upload the
            # pictures itself, the same way it handles a pasted screenshot.
            mime.setUrls([QtCore.QUrl.fromLocalFile(path) for path in files])
        QtWidgets.QApplication.clipboard().setMimeData(mime)

        if self._copy_embed:
            note = (f"Comment copied with {pictures} embedded picture(s)."
                    if pictures else "Comment copied.")
            if missing:
                note += f" {missing} picture(s) could not be downloaded."
        else:
            note = (f"Comment copied for TFS with {pictures} picture(s)."
                    if pictures else "Comment copied for TFS.")
            if missing:
                note += f" {missing} picture(s) could not be read."
        self._set_list_busy(False, note)
        logger.info(
            "Scheduled comment copied (%s): %d picture(s), %d missing",
            "embedded" if self._copy_embed else "TFS links",
            pictures, missing,
        )

    def _on_copy_failed(self, message):
        logger.warning("Could not copy the scheduled comment: %s", message)
        self._set_list_busy(False, f"Could not copy: {message}")

    def cancel_selected(self):
        """Delete every highlighted schedule, refusing the whole batch if any
        of them is still in the future.

        A schedule is deleted after its time has passed - it is either already
        posted (deleting the record leaves the ticket comment alone) or the
        worker has given up on it. Deleting a future one would silently cancel
        a comment somebody is still waiting for, so the batch is refused and
        the rows that caused it are named instead.
        """
        docs = self._selected_docs()
        if not docs:
            self.list_status.setText("Right-click a schedule to delete it.")
            return

        upcoming = [doc for doc in docs if self._is_future(doc)]
        if upcoming:
            rows = "\n".join(
                f"\u2022 WI {doc.get('work_item_id', '?')} - {self._due_text(doc)}"
                for doc in upcoming
            )
            one = len(upcoming) == 1
            if len(upcoming) == len(docs):
                tail = "Nothing else was selected, so nothing was deleted."
            else:
                tail = ("Deselect " + ("it" if one else "them")
                        + " to delete the rest of the selection.")
            AppDialog(
                "Cannot Delete",
                ("Nothing was deleted. "
                 + ("This schedule is" if one else f"These {len(upcoming)} schedules are")
                 + " still in the future, and a schedule can only be deleted "
                 "after its time has passed:\n\n" + rows + "\n\n" + tail),
                self,
            ).exec()
            return

        self._set_list_busy(
            True,
            "Deleting the schedule..." if len(docs) == 1
            else f"Deleting {len(docs)} schedules...",
        )
        self._delete_worker = DeleteSchedulesWorker(docs, self.user_data, self)
        self._delete_worker.progress.connect(self._on_delete_progress)
        self._delete_worker.done.connect(self._on_deleted)
        self._delete_worker.start()

    def _on_delete_progress(self, done, total):
        if _ui_alive(self.list_status):
            self.list_status.setText(
                f"Deleting {done} of {total}..." if total > 1
                else "Deleting the schedule..."
            )

    def _on_deleted(self, deleted, failed):
        total = len(deleted) + len(failed)
        posted = sum(1 for doc in deleted
                     if str(doc.get("status", "")) == "sent")
        if failed:
            rows = "\n".join(f"WI {doc.get('work_item_id', '?')}: {message}"
                             for doc, message in failed)
            AppDialog(
                "Delete Failed",
                (f"Deleted {len(deleted)} of {total} schedule(s).\n\n"
                 "Could not delete:\n" + rows),
                self,
            ).exec()
        if len(deleted) == 1:
            note = ("Deleted. That comment was already posted, so the ticket "
                    "keeps it." if posted else "Schedule deleted.")
        elif deleted:
            note = f"Deleted {len(deleted)} schedules."
        else:
            note = "Nothing was deleted."
        self._set_list_busy(False, note)
        self.refresh_list()

    # --- lifetime ---

    def _worker_running(self):
        for worker in (self._save_worker, self._list_worker,
                       self._copy_worker, self._delete_worker):
            if worker is not None and worker.isRunning():
                return True
        return False

    def reject(self):
        """Closing mid-upload would destroy a running QThread."""
        if self._worker_running():
            self.status_label.setText("Still working - give it a second.")
            return
        # Whatever ends the dialog, the last edit is already on disk.
        self._draft_timer.stop()
        self._save_draft()
        super().reject()

    def closeEvent(self, event):
        if self._worker_running():
            event.ignore()
            self.status_label.setText("Still working - give it a second.")
            return
        self._draft_timer.stop()
        self._save_draft()
        super().closeEvent(event)


# ================= PTU READING =================
# The PTU tool reads a permit from a PDF or a picture. Two engines do the work:
#   * a PDF's own text layer, via pypdf - exact and instant;
#   * Windows OCR (Windows.Media.Ocr) for scans, photos and screenshots, which
#     runs offline on the language packs already installed on the PC.
# Both are optional imports: the app must still start - and the tool still open
# with an explanation - on a machine where a build step skipped them.

try:
    from pypdf import PdfReader
    _PDF_IMPORT_ERROR = ""
except Exception as e:     # noqa: BLE001 - absence must not stop the app
    PdfReader = None
    _PDF_IMPORT_ERROR = f"{type(e).__name__}: {e}"

try:
    import pypdfium2 as pdfium
    _PDFIUM_IMPORT_ERROR = ""
except Exception as e:     # noqa: BLE001
    pdfium = None
    _PDFIUM_IMPORT_ERROR = f"{type(e).__name__}: {e}"

# The reason is kept, not just a None: "OCR is not available" is useless on a
# machine where the package is installed but a DLL did not load.
try:
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    _OCR_IMPORT_ERROR = ""
except Exception as e:     # noqa: BLE001
    OcrEngine = None
    _OCR_IMPORT_ERROR = f"{type(e).__name__}: {e}"
    logger.warning("Windows OCR unavailable: %s", _OCR_IMPORT_ERROR)

PTU_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
PTU_SUFFIXES = PTU_IMAGE_SUFFIXES | {".pdf"}
PTU_MIN_TEXT_CHARS = 40          # below this many characters a page is a scan
PTU_OCR_MAX_SIDE = 4000          # past this OCR is slower, not more accurate
PTU_OCR_UPSCALE_BELOW = 1200     # small screenshots read better a little larger
PTU_PDF_RENDER_SCALE = 2.0       # about 144 dpi: enough for 6-8 pt table text

# The five header fields, with the label wording a permit actually uses and the
# variants OCR returns. Order matters only for how they are listed in the UI.
PTU_FIELD_SPECS = (
    (
        "proprietor",
        "Proprietor",
        r"granted\s+to|grantee|proprietor(?:'s)?\s+name|"
        r"name\s+of\s+(?:the\s+)?(?:taxpayer|proprietor|owner)|"
        r"taxpayer(?:'s)?\s+name|registered\s+owner",
    ),
    (
        "business_name",
        "Business Name",
        r"business\s+name|trade\s+name(?:\s*/\s*business\s+name)?|"
        r"name\s+of\s+(?:the\s+)?(?:business|establishment)|"
        r"establishment(?:'s)?\s+name|registered\s+name",
    ),
    (
        "address",
        "Address",
        r"business\s+address|registered\s+address|"
        r"address\s+of\s+(?:the\s+)?(?:business|establishment)|address",
    ),
    (
        "tin",
        "TIN",
        r"taxpayer(?:'s)?\s+identification\s+number|tin\s*(?:no\.?|number)?",
    ),
    (
        "effective",
        "Effective",
        r"effective\s+(?:date|on)|date\s+of\s+effectivity|"
        r"effectivity(?:\s+date)?|valid\s+(?:from|beginning|start)",
    ),
)

# The printed permit reads as a sentence - "Permit hereby granted to <name>,
# with business name <business>, address of <address>, ... effective <date>" -
# so between a label and its value sit these connective words. They are peeled
# off the front of a value, longest phrase first.
PTU_VALUE_FILLER = {
    "proprietor": (
        r"permit\s+is\s+hereby\s+granted\s+to",
        r"permit\s+hereby\s+granted\s+to",
        r"permit\s+is\s+granted\s+to",
        r"permit\s+granted\s+to",
        r"is\s+hereby\s+granted\s+to",
        r"hereby\s+granted\s+to",
        r"granted\s+unto",
        r"granted\s+to",
        r"granted",
    ),
    "business_name": (
        r"with\s+the\s+business\s+name",
        r"with\s+business\s+name",
        r"under\s+the\s+business\s+name",
        r"under\s+business\s+name",
        r"with\s+the\s+trade\s+name",
        r"with\s+trade\s+name",
        r"business\s+name",
        r"trade\s+name",
    ),
    "address": (
        r"with\s+address\s+of",
        r"with\s+business\s+address",
        r"business\s+address",
        r"registered\s+address",
        r"address\s+of",
        r"located\s+at",
        r"of",
        r"at",
    ),
    "tin": (
        r"taxpayer(?:'s)?\s+identification\s+number",
        r"tin\s+no\.?",
        r"tin\s+number",
    ),
    "effective": (
        r"date\s+of\s+effectivity",
        r"effective\s+date",
        r"effective\s+on",
        r"effectivity(?:\s+date)?",
        r"valid\s+from",
        r"effective",
        r"date",
        r"on",
    ),
}

_PTU_MONTHS = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
# 3-3-3 and 3-3-3-5 (branch) as printed; OCR may swap the dashes for spaces.
_PTU_TIN_RE = re.compile(r"\b(\d{3})[-\u2013\u2014\s.]?(\d{3})[-\u2013\u2014\s.]?(\d{3})(?:[-\u2013\u2014\s.]?(\d{3,5}))?\b")
_PTU_DATE_RE = re.compile(
    r"(?i)\b(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}"
    r"|\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?,?\s+\d{4}"
    r"|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b"
)


def _ptu_cut_at_other_label(text, key):
    """Text up to the next label of a *different* field.

    The same field's label is skipped on purpose: prose such as "permit hereby
    granted to" is also the proprietor label, and cutting there would leave an
    empty value.
    """
    cut = len(text)
    for other_key, _label, pattern in PTU_FIELD_SPECS:
        if other_key == key:
            continue
        match = re.search(r"(?i)\b(?:" + pattern + r")\b", text)
        if match and match.start() < cut:
            cut = match.start()
    return text[:cut]


def _ptu_strip_filler(key, text):
    """Drop the connectives the printed sentence puts between label and value.

    "Business Name: With business name JJ-..." leaves "JJ-..."; "Address : of
    1354 ..." leaves "1354 ...". Longest phrase first, so "with business name"
    wins over "business name".
    """
    text = text.lstrip(" \t\r\n:;,.|_\u2013\u2014-")
    for _ in range(6):
        for phrase in PTU_VALUE_FILLER.get(key, ()):
            # The hyphen is last in the class on purpose: anywhere else it
            # would make a range and eat the whole value.
            match = re.match(r"(?i)" + phrase + r"\b[\s:;,.|_\u2013\u2014-]*", text)
            if match and match.end() > 0:
                text = text[match.end():]
                break
        else:
            break
    return text.strip()


def _ptu_month_name(word):
    """The month a token stands for, tolerating one misread letter: "Apnl"."""
    token = re.sub(r"[^a-z]", "", word.lower())
    if len(token) < 3:
        return ""
    for name in _PTU_MONTHS:
        if token in (name.lower(), name[:3].lower()):
            return name
    # Only months that start with the same letter are candidates, so a name
    # like "Janna" can never be read as July.
    candidates = [name for name in _PTU_MONTHS if name[0].lower() == token[0]]
    if not candidates:
        return ""
    best = max(
        candidates,
        key=lambda name: difflib.SequenceMatcher(None, token, name.lower()).ratio(),
    )
    if difflib.SequenceMatcher(None, token, best.lower()).ratio() >= 0.6:
        return best
    return ""


def _ptu_find_date(value):
    """The date in a value, repaired when OCR misread the month.

    "Apnl 27. 2026" comes back as "April 27, 2026": the month name is the one
    Windows knows, while day and year are the numbers OCR actually read.
    """
    strict = _PTU_DATE_RE.search(value)
    if strict:
        return strict.group(0).strip(" ,.")

    words = re.findall(r"[A-Za-z]+|\d{1,4}", value)
    numbers = [int(word) for word in words if word.isdigit()]
    years = [number for number in numbers if 1900 <= number <= 2099]
    days = [number for number in numbers if 1 <= number <= 31 and number not in years]
    for word in words:
        if not word.isalpha():
            continue
        month = _ptu_month_name(word)
        if not month:
            continue
        if years and days:
            return f"{month} {days[0]}, {years[0]}"
        if years:
            return f"{month} {years[0]}"

    stripped = value.strip(" ,.")
    return stripped if any(char.isdigit() for char in stripped) else ""


def _ptu_finalize(key, value):
    """One line, no wrapping punctuation; names stop at their first comma."""
    value = re.sub(r"\s+", " ", value).strip(" \t:;,-\u2013\u2014.|_")
    if key in ("proprietor", "business_name"):
        value = re.split(r"\s*[,;]\s*", value)[0].strip(" \t:;,-\u2013\u2014.|_")
    return value


def _ptu_format_tin(match):
    """A TIN the way it is printed: 682-090-387-00001."""
    return "-".join(group for group in match.groups() if group)


def _ptu_extract_fields(text):
    """Proprietor, business name, address, TIN and effective date from OCR text.

    Worked out on the whole text, not line by line: on a skewed scan a label and
    its value rarely share a line, and the printed permit states each value as a
    sentence that repeats the label ("Business Name: With business name ...").
    For each field, the first label whose remainder yields a plausible value
    wins; nothing is inferred from position alone.
    """
    flat = re.sub(r"\s+", " ", text or "").strip()
    values = {key: "" for key, _label, _pattern in PTU_FIELD_SPECS}
    if not flat:
        return values

    for key, _label, pattern in PTU_FIELD_SPECS:
        label_re = re.compile(r"(?i)\b(?:" + pattern + r")\b")
        for match in label_re.finditer(flat):
            value = _ptu_cut_at_other_label(flat[match.end():], key)
            value = _ptu_strip_filler(key, value)
            if key == "tin":
                found = _PTU_TIN_RE.search(value) or _PTU_TIN_RE.search(flat)
                if not found:
                    continue
                value = _ptu_format_tin(found)
            elif key == "effective":
                value = _ptu_find_date(value)
                if not value:
                    continue
            else:
                value = _ptu_finalize(key, value)
                if not value:
                    continue
            values[key] = value
            break

    # A TIN is the one field with a shape of its own: if its label was lost to
    # OCR, the number itself is still findable.
    if not values["tin"]:
        found = _PTU_TIN_RE.search(flat)
        if found:
            values["tin"] = _ptu_format_tin(found)

    return values


def _ptu_prepare_image(image):
    """A PIL image as PNG bytes, sized for the OCR engine.

    Small screenshots are doubled: the engine reads 6-8 pt table text far more
    reliably at twice the pixels, and a photo of a permit is rarely large
    enough to lose detail that way. Very large pictures are scaled down to the
    engine's comfortable range instead.
    """
    prepared = image.convert("RGB")
    width, height = prepared.size
    longest = max(width, height)
    if not longest:
        factor = 1.0
    elif longest <= PTU_OCR_UPSCALE_BELOW:
        factor = 2.0
    elif longest > PTU_OCR_MAX_SIDE:
        factor = PTU_OCR_MAX_SIDE / longest
    else:
        factor = 1.0
    if factor != 1.0:
        prepared = prepared.resize(
            (max(1, round(width * factor)), max(1, round(height * factor))),
            Image.Resampling.LANCZOS,
        )
    buffer = BytesIO()
    prepared.save(buffer, format="PNG")
    return buffer.getvalue()


async def _ptu_ocr_async(png_bytes):
    """Run Windows OCR over one PNG and return its lines joined by newlines."""
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(png_bytes)
    await writer.store_async()
    writer.detach_stream()
    stream.seek(0)

    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()

    # The user's own languages first (so a Filipino/English PC OCRs the same
    # way its Photos app does), then any recogniser the PC has installed.
    engine = OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        for language in OcrEngine.available_recognizer_languages:
            engine = OcrEngine.try_create_from_language(language)
            if engine is not None:
                break
    if engine is None:
        raise RuntimeError(
            "Windows OCR has no language installed. Add one under Settings > "
            "Time & language > Language & region."
        )

    result = await engine.recognize_async(bitmap)
    return "\n".join(line.text for line in result.lines)


def _ptu_ocr_png(png_bytes):
    """Windows OCR on one PNG. Blocking: call it from the worker thread."""
    if OcrEngine is None:
        raise RuntimeError(
            "OCR is not available in the Python running this app. Install the "
            "packages with `pip install -r requirements.txt` (then run "
            f"build.bat if you use the exe). Reason: {_OCR_IMPORT_ERROR or 'unknown'}"
        )
    return asyncio.run(_ptu_ocr_async(png_bytes))


def _ptu_read_image(path):
    """Text of one picture, plus a note saying which engine read it."""
    with Image.open(path) as image:
        png = _ptu_prepare_image(image)
    text = _ptu_ocr_png(png).strip()
    return text, "Read with Windows OCR (offline)."


def _ptu_read_pdf(path):
    """Text of a PDF, mixing its text layer and OCR page by page.

    A page with a real text layer is used as-is: that is exact. A page with
    almost no text is treated as a scan, rendered to pixels and OCR'd, which is
    what makes a printed-then-scanned permit readable at all.
    """
    if PdfReader is None:
        raise RuntimeError(
            "Reading PDFs is not available in the Python running this app. "
            "Install pypdf with `pip install -r requirements.txt` "
            f"(then run build.bat if you use the exe). Reason: {_PDF_IMPORT_ERROR}"
        )

    reader = PdfReader(path)
    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt("")
        except Exception:
            raise RuntimeError("That PDF is password protected.")

    document = None
    pages = []
    from_text = 0
    from_ocr = 0
    skipped = 0

    for index, page in enumerate(reader.pages):
        try:
            text = (page.extract_text() or "").strip()
        except Exception:
            logger.exception("Could not extract page %d of %s", index + 1, path)
            text = ""

        if len(text) >= PTU_MIN_TEXT_CHARS:
            pages.append((index + 1, text))
            from_text += 1
            continue

        if pdfium is None:
            skipped += 1
            pages.append((index + 1, text))
            continue

        try:
            if document is None:
                document = pdfium.PdfDocument(path)
            bitmap = document[index].render(scale=PTU_PDF_RENDER_SCALE)
            image = bitmap.to_pil()
            try:
                text = _ptu_ocr_png(_ptu_prepare_image(image)).strip()
            finally:
                image.close()
            pages.append((index + 1, text))
            from_ocr += 1
        except Exception as e:
            logger.exception("Could not read page %d of %s as a scan", index + 1, path)
            pages.append((index + 1, f"(this page could not be read: {e})"))
            skipped += 1

    notes = []
    if from_text:
        notes.append(f"{from_text} page(s) read from the PDF's text layer.")
    if from_ocr:
        notes.append(f"{from_ocr} scanned page(s) read with Windows OCR.")
    if skipped and pdfium is None:
        notes.append(
            f"{skipped} page(s) have no text layer and could not be read "
            f"because pypdfium2 is missing ({_PDFIUM_IMPORT_ERROR})."
        )

    if len(pages) > 1:
        body = "\n\n".join(
            f"--- Page {number} ---\n{text}" for number, text in pages
        )
    elif pages:
        body = pages[0][1]
    else:
        body = ""
    return body, " ".join(notes) or "The PDF has no readable pages."


def _ptu_read_file(path):
    """(text, note) for one PDF or picture. Raises RuntimeError with a reason."""
    suffix = os.path.splitext(path)[1].lower()
    if suffix not in PTU_SUFFIXES:
        raise RuntimeError("Use a PDF or a picture (PNG, JPG, BMP, TIFF, WEBP).")
    if not os.path.exists(path):
        raise RuntimeError("That file is no longer there.")
    if suffix == ".pdf":
        return _ptu_read_pdf(path)
    return _ptu_read_image(path)


# ================= GEMINI (OPTIONAL) =================
# Local OCR reads characters; a vision model reads the page. This permit is a
# skewed, sentence-shaped form, so the model is clearly better at the five
# header fields - at the cost of the image leaving this PC. That trade is the
# user's to make: the tool is only opened once a key exists, and the tier text
# in PtuApiKeyDialog says where the image goes.
#
# REST only, through `requests` (already a dependency): no google-genai SDK to
# freeze into the build, and no extra wheels.
#
# Google retires model ids (gemini-2.5-flash now answers 404), so the default is
# the current stable Flash and a 404 walks the fallback list instead of failing
# the read. Reported as "Fields read by Gemini (<model>)" so the log says which
# one answered.
PTU_GEMINI_MODEL = (
    os.environ.get("AUTOREPORT_GEMINI_MODEL", "").strip() or "gemini-3.8-flash"
)
PTU_GEMINI_FALLBACK_MODELS = ("gemini-3.7-flash", "gemini-2.5-flash")
# 503 ("high demand") is routine on a new Flash model, and a short wait usually
# turns it into a success. 429 is retried the same way: it is normally a
# per-minute limit rather than an exhausted quota, and the status line still
# explains it if the last attempt fails.
PTU_AI_RETRY_STATUSES = frozenset({429, 500, 502, 503, 504})
PTU_AI_RETRY_WAIT = 3           # seconds; attempt n waits n * this
PTU_AI_ATTEMPTS = (3, 2)        # per model, in order; the default gets one more
# Asking the API which models exist beats hard-coding ids Google retires, and
# leaves room to fail over when the newest Flash is the one that is busy.
PTU_GEMINI_MODELS_URL = "https://generativelanguage.googleapis.com/v1beta/models"
PTU_AI_MODEL_WORDS = ("tts", "live", "image", "embedding", "audio")
PTU_AI_MODELS = []              # filled by the first read, reused after that
PTU_AI_KEY_URL = "https://aistudio.google.com/apikey"
PTU_GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
)
PTU_AI_MAX_SIDE = 2400      # a 2400 px page keeps table text legible
PTU_AI_MAX_PAGES = 3        # a permit is one or two pages in practice
PTU_AI_PROMPT = (
    "This is a Philippine BIR permit to use sales machine. Read the page and "
    "return JSON only, with exactly these keys and no others:\n"
    '{"proprietor": "", "business_name": "", "address": "", "tin": "", '
    '"effective": ""}\n'
    "proprietor: the name the permit is granted to (after GRANTED TO or "
    "PROPRIETOR).\n"
    "business_name: the business or trade name.\n"
    "address: the full registered address, exactly as printed.\n"
    "tin: the taxpayer identification number, formatted 000-000-000-00000.\n"
    "effective: the effective date, formatted Month DD, YYYY.\n"
    "Do not repeat the label, do not explain, do not add keys. Use an empty "
    "string for any value that is not printed on the page."
)


def _ai_key_file():
    return os.path.join(_credential_dir(), "gemini_api_key.bin")


def _save_ai_key(key):
    """Store the Gemini key for this Windows user, like the TFS PAT."""
    if not key:
        return False
    try:
        with open(_ai_key_file(), "wb") as f:
            f.write(_dpapi(key.encode("utf-8"), protect=True))
        return True
    except Exception:
        logger.exception("Could not store the Gemini API key locally")
        return False


def _load_ai_key():
    try:
        path = _ai_key_file()
        if not os.path.isfile(path):
            return ""
        with open(path, "rb") as f:
            return _dpapi(f.read(), protect=False).decode("utf-8")
    except Exception:
        logger.exception("Could not read the stored Gemini API key")
        return ""


PTU_AI_KEY_FIELD = "gemini_api_key"


def _load_ai_key_remote(user_data):
    """This user's Gemini key from users/{uid}, or "" when there is none.

    The key lives on the owner-only user document - the same place as the TFS
    PAT - so it follows the account to any PC. Raises FirebaseError when the
    document cannot be read, so the caller can tell "no key saved" apart from
    "could not ask".
    """
    uid = (user_data or {}).get("uid", "")
    if not uid:
        return ""
    token = user_data.get("id_token", "")
    try:
        profile = _load_profile(uid, token)
    except FirebaseError:
        # An hour-old token is the common cause; one refresh, then give up.
        token = _refresh_id_token(user_data) or token
        profile = _load_profile(uid, token)
    return str(profile.get(PTU_AI_KEY_FIELD) or "").strip()


def _save_ai_key_remote(user_data, key):
    """Mirror the Gemini key into users/{uid}.

    Best effort, like the PAT: offline or a rules refusal leaves the DPAPI copy
    on this PC in charge, so the tool keeps working here.
    """
    uid = (user_data or {}).get("uid", "")
    if not (uid and key):
        return False
    try:
        _update_user_profile(user_data, {PTU_AI_KEY_FIELD: key})
        logger.info("Gemini API key synced to Firestore for uid=%s", uid)
        return True
    except Exception as e:
        logger.warning("Could not sync the Gemini key to Firestore: %s", e)
        return False


def _delete_ai_key():
    """Remove the DPAPI copy from this PC. True when one was actually deleted."""
    path = _ai_key_file()
    if not os.path.isfile(path):
        return False
    try:
        os.remove(path)
        logger.info("Gemini API key removed from this PC")
        return True
    except Exception:
        logger.exception("Could not remove the Gemini API key from this PC")
        return False


def _delete_ai_key_remote(user_data):
    """Drop gemini_api_key from users/{uid}: the field goes, the doc stays."""
    uid = (user_data or {}).get("uid", "")
    if not uid:
        return False
    try:
        # None on the field path is how the masked update deletes a field.
        _update_user_profile(user_data, {PTU_AI_KEY_FIELD: None})
        logger.info("Gemini API key removed from Firestore for uid=%s", uid)
        return True
    except Exception as e:
        logger.warning("Could not remove the Gemini key from Firestore: %s", e)
        return False


def _ptu_ai_image_bytes(image):
    """One page as JPEG for the model."""
    prepared = image.convert("RGB")
    width, height = prepared.size
    longest = max(width, height)
    if longest > PTU_AI_MAX_SIDE:
        scale = PTU_AI_MAX_SIDE / longest
        prepared = prepared.resize(
            (max(1, round(width * scale)), max(1, round(height * scale))),
            Image.Resampling.LANCZOS,
        )
    # JPEG on purpose: a 2400 px page is about 1 MB instead of 6 MB, and the
    # model gains nothing from the extra bytes.
    buffer = BytesIO()
    prepared.save(buffer, format="JPEG", quality=90, optimize=True)
    return buffer.getvalue()


def _ptu_ai_images(path):
    """The pages of the file as JPEG bytes, for sending to the model."""
    suffix = os.path.splitext(path)[1].lower()
    if suffix != ".pdf":
        with Image.open(path) as image:
            return [_ptu_ai_image_bytes(image)]
    if pdfium is None:
        raise RuntimeError(
            "Sending a PDF to Gemini needs pypdfium2 to render its pages. "
            "Run: pip install -r requirements.txt"
        )
    document = pdfium.PdfDocument(path)
    images = []
    for index in range(min(len(document), PTU_AI_MAX_PAGES)):
        bitmap = document[index].render(scale=PTU_PDF_RENDER_SCALE)
        image = bitmap.to_pil()
        try:
            images.append(_ptu_ai_image_bytes(image))
        finally:
            image.close()
    return images


def _ptu_ai_parse_json(answer):
    """The JSON object in the model's answer, tolerating a stray code fence."""
    text = answer.strip()
    text = text.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    try:
        return json.loads(text)
    except ValueError:
        start, end = text.find("{"), text.rfind("}")
        if start != -1 and end > start:
            try:
                return json.loads(text[start:end + 1])
            except ValueError:
                pass
    raise RuntimeError("Gemini's answer was not the JSON that was asked for.")


def _ptu_ai_error(response):
    """A readable reason for a failed Gemini call (the body is a JSON error)."""
    status = response.status_code
    if status in (401, 403):
        return "The Gemini API key was rejected. Check it in Google AI Studio."
    if status == 429:
        return "Gemini's rate limit or quota was reached. Try again in a minute."
    if status == 503:
        return (
            "Gemini is busy right now (503, high demand) and every retry met "
            "the same answer. The permit is fine - try the same file again in a "
            "minute."
        )
    if status in (500, 502, 504):
        return (
            f"Gemini had a server error ({status}) and the retries did not "
            "clear it. Try the same file again in a minute."
        )
    snippet = re.sub(r"\s+", " ", response.text or "")[:300].strip()
    return f"Gemini returned {status}. {snippet}".strip()


def _ptu_ai_post(model, api_key, payload):
    """One generateContent call for one model id."""
    try:
        return requests.post(
            PTU_GEMINI_URL.format(model=model),
            headers={"x-goog-api-key": api_key, "Content-Type": "application/json"},
            json=payload,
            timeout=(15, 120),
        )
    except requests.RequestException as e:
        # No URL or key in the message: the status line is visible, the log is
        # not always.
        raise RuntimeError(
            f"Could not reach Gemini ({e.__class__.__name__}). Check the "
            "internet connection."
        )


def _ptu_ai_discovered_models(api_key):
    """Flash text models this key can call, as reported by the API itself.

    Discovered once per launch: the list changes only when Google ships or
    retires a model, and a permit read should not pay for a listing every time.
    """
    if PTU_AI_MODELS:
        return list(PTU_AI_MODELS)
    body = {}
    try:
        resp = requests.get(
            PTU_GEMINI_MODELS_URL,
            headers={"x-goog-api-key": api_key},
            timeout=15,
        )
        if resp.status_code == 200:
            body = resp.json() or {}
        else:
            logger.warning("Listing Gemini models answered %d", resp.status_code)
    except Exception:
        logger.exception("Could not list the Gemini models")
    found = []
    for item in body.get("models", []) or []:
        name = str(item.get("name", "")).removeprefix("models/")
        methods = [str(m).lower() for m in item.get("supportedGenerationMethods") or []]
        lowered = name.lower()
        if not name or "generatecontent" not in methods:
            continue
        if "flash" not in lowered or any(w in lowered for w in PTU_AI_MODEL_WORDS):
            continue
        found.append(name)
    PTU_AI_MODELS.extend(found[:4])
    return list(PTU_AI_MODELS)


def _ptu_ai_post_retrying(model, api_key, payload, attempts):
    """POST once, waiting out the statuses that mean "busy", not "wrong".

    A 5xx or a 429 says nothing about the page, so the request is repeated a
    couple of times before the model is treated as unavailable. Sleeping here is
    fine: this runs in PtuExtractionWorker, not on the GUI thread.
    """
    response = None
    for attempt in range(max(1, attempts)):
        response = _ptu_ai_post(model, api_key, payload)
        if response.status_code not in PTU_AI_RETRY_STATUSES:
            return response
        if attempt + 1 < attempts:
            wait = PTU_AI_RETRY_WAIT * (attempt + 1)
            logger.warning(
                "Gemini %s answered %d; retrying in %ss",
                model, response.status_code, wait,
            )
            time.sleep(wait)
    return response


def _ptu_ai_read(path, api_key):
    """(values, model) read by Gemini from the page image itself.

    Two kinds of failure are worked around instead of reported: a 404 (the model
    id was retired or is not on this key's plan) moves to the next candidate, and
    a 5xx/429 is retried on the same one first, because those are demand spikes
    rather than anything wrong with the page or the key.
    """
    images = _ptu_ai_images(path)
    if not images:
        raise RuntimeError("No page could be prepared to send.")
    parts = [{"text": PTU_AI_PROMPT}]
    for data in images:
        parts.append({
            "inline_data": {
                "mime_type": "image/jpeg",
                "data": base64.b64encode(data).decode("ascii"),
            }
        })
    payload = {
        "contents": [{"role": "user", "parts": parts}],
        # No temperature: the Gemini 3 family ignores or rejects the older
        # sampling knobs, and the prompt already pins the output shape.
        "generationConfig": {"responseMimeType": "application/json"},
    }
    models = [PTU_GEMINI_MODEL]
    models += _ptu_ai_discovered_models(api_key)
    models += list(PTU_GEMINI_FALLBACK_MODELS)
    models = list(dict.fromkeys(models))[:5]
    response = None
    model = models[0]
    for index, candidate in enumerate(models):
        attempts = PTU_AI_ATTEMPTS[min(index, len(PTU_AI_ATTEMPTS) - 1)]
        response = _ptu_ai_post_retrying(candidate, api_key, payload, attempts)
        if response.status_code == 200:
            model = candidate
            break
        if response.status_code in PTU_AI_RETRY_STATUSES:
            # Busy is about the model, not the page: the next one may have
            # capacity, and 3.8 Flash is the newest, so it is the busiest.
            logger.warning(
                "Gemini model %s is busy (%d); trying the next one",
                candidate, response.status_code,
            )
            continue
        if response.status_code == 404:
            logger.warning(
                "Gemini model %s is not available for this key; trying the next",
                candidate,
            )
            continue
        # A rejected key or a bad request fails the same way on every model.
        model = candidate
        break
    else:
        if response is not None and response.status_code in PTU_AI_RETRY_STATUSES:
            raise RuntimeError(_ptu_ai_error(response))
        raise RuntimeError(
            "None of these Gemini models is available for this key: "
            f"{', '.join(models)}. Check the key's project in Google AI Studio, "
            "or set AUTOREPORT_GEMINI_MODEL to a model the project can use."
        )
    if response.status_code != 200:
        raise RuntimeError(_ptu_ai_error(response))

    try:
        body = response.json()
    except ValueError:
        raise RuntimeError("Gemini's response was not JSON.")
    block = (body.get("promptFeedback") or {}).get("blockReason")
    if block:
        raise RuntimeError(f"Gemini would not read the page ({block}).")
    candidates = body.get("candidates") or []
    if not candidates:
        raise RuntimeError("Gemini returned no result for the page.")
    response_parts = (candidates[0].get("content") or {}).get("parts", [])
    # Gemini 3 can return its reasoning as extra parts. The JSON is in the parts
    # that are not marked as thoughts, and a thought containing a brace would
    # confuse the brace-slice fallback in _ptu_ai_parse_json.
    answer = "".join(
        part.get("text", "") for part in response_parts if not part.get("thought")
    ).strip() or "".join(part.get("text", "") for part in response_parts).strip()
    if not answer:
        finish = candidates[0].get("finishReason") or "unknown"
        raise RuntimeError(f"Gemini returned no text ({finish}).")

    parsed = _ptu_ai_parse_json(answer)
    values = {}
    for key, _label, _pattern in PTU_FIELD_SPECS:
        raw = parsed.get(key)
        values[key] = _ptu_finalize(key, raw) if isinstance(raw, str) else ""
    tin = _PTU_TIN_RE.search(values["tin"])
    values["tin"] = _ptu_format_tin(tin) if tin else values["tin"]
    if values["effective"]:
        values["effective"] = (
            _ptu_find_date(values["effective"]) or values["effective"]
        )
    return values, model


class PtuExtractionWorker(QThread):
    """Reads one PTU file off the GUI thread: OCR and the AI call are slow."""

    done = pyqtSignal(str, str, object)     # text, note, AI fields (or None)
    failed = pyqtSignal(str)

    def __init__(self, path, parent=None, api_key=""):
        super().__init__(parent)
        self.path = path
        self.api_key = api_key or ""

    def run(self):
        text, note = "", ""
        try:
            text, note = _ptu_read_file(self.path)
        except Exception as e:
            logger.exception("Could not read the PTU file %s", self.path)
            # With a key the model can still read the page, and its answer is
            # the better of the two - so a local failure is only fatal offline.
            if not self.api_key:
                self.failed.emit(str(e))
                return
            note = f"Local reading failed ({e})."

        fields = None
        if self.api_key:
            try:
                fields, model = _ptu_ai_read(self.path, self.api_key)
                note = f"{note} Fields read by Gemini ({model}).".strip()
            except Exception as e:
                logger.exception("Gemini could not read the PTU file %s", self.path)
                note = f"{note} AI read failed: {e}".strip()
        self.done.emit(text, note, fields)


class PtuApiKeyDialog(QDialog):
    """The gate in front of the PTU tool: a Gemini key, or no tool.

    The tool reads permits with Gemini, so it is useless without a key. This
    asks for one, points at the page that creates it, and stores it both on this
    PC (DPAPI, like the TFS token) and on the signed-in user's Firestore
    document - so it is asked once per account, not once per machine.
    """

    def __init__(self, parent=None, user_data=None, note=""):
        super().__init__(parent)
        self.setWindowIcon(QtGui.QIcon(os.path.join(ASSET_DIR, "death.ico")))
        self.setWindowTitle("PTU Generator - Gemini API key needed")
        self.setObjectName("Root")
        self.setMinimumWidth(580)
        self.saved_key = ""
        self.user_data = user_data or {}

        if is_windows_dark_mode():
            self.setStyleSheet("""
            QDialog#Root {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #07101a, stop:1 #07121a);
                font-family: Inter, Segoe UI, Arial;
                color: #e6eef5;
            }
            QLabel { color: #e6eef5; background: transparent; }
            QLabel#noteLabel { color: #f6c177; }
            QLineEdit {
                background: rgba(255,255,255,0.06);
                border: 1px solid rgba(255,255,255,0.14);
                border-radius: 8px;
                padding: 8px 10px;
                color: #e6eef5;
                selection-background-color: #ed3b3b;
            }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white;
                border: none;
                border-radius: 10px;
                padding: 8px 16px;
            }
            QPushButton:disabled { background: rgba(255,255,255,0.12); color: #9fb0be; }
            QPushButton#ghost {
                background: transparent;
                border: 1px solid rgba(255,255,255,0.18);
                color: #e6eef5;
            }
            QPushButton#ghost:hover { background: rgba(255,255,255,0.06); }
            """)
        else:
            self.setStyleSheet("""
            QDialog#Root {
                background: #f5f7fa;
                font-family: Inter, Segoe UI, Arial;
                color: #1f2933;
            }
            QLabel { color: #111827; background: transparent; }
            QLabel#noteLabel { color: #92400e; }
            QLineEdit {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 8px;
                padding: 8px 10px;
                color: #111827;
                selection-background-color: #ed3b3b;
            }
            QPushButton {
                background: #ed3b3b;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 8px 16px;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton:disabled { background: #e5e7eb; color: #9ca3af; }
            QPushButton#ghost {
                background: #ffffff;
                border: 1px solid #d1d5db;
                color: #111827;
            }
            QPushButton#ghost:hover { background: #eef1f5; }
            """)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("This tool needs a Gemini API key")
        title.setStyleSheet("font-size: 16px; font-weight: 700;")
        layout.addWidget(title)

        if note:
            note_label = QtWidgets.QLabel(note)
            note_label.setWordWrap(True)
            note_label.setObjectName("noteLabel")
            layout.addWidget(note_label)

        where = (
            "It goes into your account as well as this PC, so the tool opens "
            "straight away on any PC you sign in to."
            if self.user_data
            else "It is kept encrypted on this PC."
        )
        body = QtWidgets.QLabel(
            "PTU Generator reads the permit with Google Gemini, so it does not "
            "open until there is a key. Create one in Google AI Studio (the "
            "free tier is enough) and paste it below. " + where
        )
        body.setWordWrap(True)
        layout.addWidget(body)

        link = QtWidgets.QLabel(f'<a href="{PTU_AI_KEY_URL}">{PTU_AI_KEY_URL}</a>')
        link.setOpenExternalLinks(True)
        link.setTextInteractionFlags(
            QtCore.Qt.TextInteractionFlag.TextBrowserInteraction
        )
        layout.addWidget(link)

        self.key_edit = QtWidgets.QLineEdit()
        self.key_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self.key_edit.setPlaceholderText("Paste the key from Google AI Studio")
        self.key_edit.returnPressed.connect(self._save)
        layout.addWidget(self.key_edit)

        self.status_label = QtWidgets.QLabel(
            "The permit image is sent to Google only while the fields are read."
        )
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        buttons = QtWidgets.QHBoxLayout()
        open_btn = QPushButton("Open Google AI Studio")
        open_btn.setObjectName("ghost")
        open_btn.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl(PTU_AI_KEY_URL))
        )
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("ghost")
        cancel_btn.clicked.connect(self.reject)
        self.save_btn = QPushButton("Save key")
        self.save_btn.clicked.connect(self._save)
        buttons.addWidget(open_btn)
        buttons.addStretch()
        buttons.addWidget(cancel_btn)
        buttons.addWidget(self.save_btn)
        layout.addLayout(buttons)

    def _save(self):
        key = self.key_edit.text().strip()
        if not key:
            self.status_label.setText(
                "Paste the key first, or open Google AI Studio to create one."
            )
            self.key_edit.setFocus()
            return
        saved_here = _save_ai_key(key)
        saved_remote = _save_ai_key_remote(self.user_data, key)
        if not (saved_here or saved_remote):
            self.status_label.setText(
                "The key could not be saved on this PC - see the log for why."
            )
            return
        if saved_here and self.user_data and not saved_remote:
            # Still usable here; the account copy is what carries it to another
            # PC. _save_ai_key_remote already logged the reason.
            logger.warning("Gemini key saved on this PC only, not to the account")
        self.saved_key = key
        self.accept()


class PtuGeneratorDialog(QDialog):
    """Read a PTU permit from a PDF or an image and show what it contains.

    Gemini reads the page image and returns the five header fields; Windows OCR
    runs alongside it to show the text found on this PC and to fill any field
    the model left empty.
    """

    def __init__(self, parent=None, api_key="", user_data=None):
        super().__init__(parent)
        self._worker = None
        self.user_data = user_data or {}
        # The key already accepted for this session; a successful read remembers
        # it on the PC and on the account, but not twice for the same key.
        self._synced_key = ""
        self.setWindowTitle("PTU Generator")
        self.setObjectName("PtuGenerator")
        self.setMinimumSize(900, 760)
        self.setStyleSheet(self._style())
        self.setAcceptDrops(True)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("PTU Generator")
        title.setStyleSheet("font-size: 20px; font-weight: 700;")
        layout.addWidget(title)

        subtitle = QtWidgets.QLabel(
            "Drop a PTU permit (PDF or image) below, paste a screenshot with "
            "Ctrl+V, or browse for the file. Gemini reads the page and the fields "
            "it finds appear underneath, with the text read on this PC. The file "
            "itself is never changed."
        )
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        self.drop_zone = QtWidgets.QFrame()
        self.drop_zone.setObjectName("dropZone")
        self.drop_zone.setMinimumHeight(96)
        self.drop_zone.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        drop_layout = QtWidgets.QVBoxLayout(self.drop_zone)
        drop_layout.setContentsMargins(16, 14, 16, 14)
        self.drop_label = QtWidgets.QLabel(
            "Drop a PDF or an image here\n"
            "or press Ctrl+V to paste, or click to browse"
        )
        self.drop_label.setObjectName("dropLabel")
        self.drop_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        drop_layout.addWidget(self.drop_label)
        self.drop_zone.mousePressEvent = lambda event: self._browse()
        layout.addWidget(self.drop_zone)

        buttons = QtWidgets.QHBoxLayout()
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.clicked.connect(self._browse)
        self.paste_btn = QPushButton("Paste")
        self.paste_btn.setObjectName("ghost")
        self.paste_btn.clicked.connect(self._paste)
        self.copy_btn = QPushButton("Copy text")
        self.copy_btn.setObjectName("ghost")
        self.copy_btn.clicked.connect(self._copy_text)
        buttons.addWidget(self.browse_btn)
        buttons.addWidget(self.paste_btn)
        buttons.addStretch()
        buttons.addWidget(self.copy_btn)
        layout.addLayout(buttons)

        # The permit is read by Gemini, so a key is what gets this tool opened at
        # all. The row stays so the key can be replaced. It is stored the way
        # the TFS PAT is - DPAPI, per Windows user - never bundled into the exe.
        key_row = QtWidgets.QHBoxLayout()
        key_row.setSpacing(10)
        key_label = QtWidgets.QLabel("Gemini API key")
        self.ai_key_edit = QtWidgets.QLineEdit()
        self.ai_key_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self.ai_key_edit.setText(api_key or _load_ai_key())
        self.ai_key_edit.setPlaceholderText("Paste the key from Google AI Studio")
        self.getkey_btn = QPushButton("Get a key")
        self.getkey_btn.setObjectName("ghost")
        self.getkey_btn.setToolTip(PTU_AI_KEY_URL)
        self.getkey_btn.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl(PTU_AI_KEY_URL))
        )
        self.forget_btn = QPushButton("Forget key")
        self.forget_btn.setObjectName("ghost")
        self.forget_btn.setToolTip(
            "Delete the saved key from this PC and from your account, then paste "
            "a different one."
        )
        self.forget_btn.clicked.connect(self._forget_key)
        key_row.addWidget(key_label)
        key_row.addWidget(self.ai_key_edit, stretch=1)
        key_row.addWidget(self.getkey_btn)
        key_row.addWidget(self.forget_btn)
        layout.addLayout(key_row)

        # The five header fields, read-only: select a value to copy it, and a
        # field the OCR could not pin down says "not found" instead of lying.
        self.fields_frame = QtWidgets.QFrame()
        self.fields_frame.setObjectName("fieldsBox")
        fields_form = QtWidgets.QFormLayout(self.fields_frame)
        fields_form.setContentsMargins(14, 12, 14, 12)
        fields_form.setHorizontalSpacing(14)
        fields_form.setVerticalSpacing(8)
        fields_form.setLabelAlignment(
            QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter
        )
        self.field_edits = {}
        for key, label, _pattern in PTU_FIELD_SPECS:
            edit = QtWidgets.QLineEdit()
            edit.setReadOnly(True)
            edit.setPlaceholderText("not found")
            self.field_edits[key] = edit
            fields_form.addRow(label, edit)
        layout.addWidget(self.fields_frame)

        self.preview = QTextEdit()
        self.preview.setObjectName("ptuPreview")
        self.preview.setReadOnly(True)
        self.preview.setAcceptDrops(False)
        self.preview.setPlaceholderText(
            "The text found in the permit appears here."
        )
        self.preview.setFont(QtGui.QFont("Consolas", 10))
        layout.addWidget(self.preview, stretch=1)

        self.busy_bar = QtWidgets.QProgressBar()
        self.busy_bar.setObjectName("busyBar")
        self.busy_bar.setRange(0, 0)
        self.busy_bar.setTextVisible(False)
        self.busy_bar.setFixedHeight(4)
        self.busy_bar.setVisible(False)
        layout.addWidget(self.busy_bar)

        self.status_label = QtWidgets.QLabel(
            "Supported files: PDF, PNG, JPG, BMP, TIFF, WEBP."
        )
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        # Say what is missing before the user drops a file on it, with the real
        # import error: a missing package and a broken install read differently.
        if OcrEngine is None:
            self.status_label.setText(
                "Image and scan reading is off - the Windows OCR packages are "
                f"not installed in this Python ({_OCR_IMPORT_ERROR}). Run: "
                "pip install -r requirements.txt"
            )
        elif PdfReader is None or pdfium is None:
            self.status_label.setText(
                "Scanned PDF pages cannot be read - install the PDF packages "
                "with: pip install -r requirements.txt"
            )

        close_row = QtWidgets.QHBoxLayout()
        close_row.addStretch()
        close_btn = QPushButton("Close")
        close_btn.setObjectName("ghost")
        close_btn.clicked.connect(self.reject)
        close_row.addWidget(close_btn)
        layout.addLayout(close_row)

        # Ctrl+V works wherever the focus is, including the read-only preview.
        paste_shortcut = QtGui.QShortcut(
            QtGui.QKeySequence.StandardKey.Paste, self
        )
        paste_shortcut.setContext(QtCore.Qt.ShortcutContext.WindowShortcut)
        paste_shortcut.activated.connect(self._paste)

    # --- input ---

    @staticmethod
    def _path_from_mime(mime):
        """The first dropped file this tool can read, or ""."""
        if mime is None:
            return ""
        for url in mime.urls():
            local = url.toLocalFile()
            if local and os.path.splitext(local)[1].lower() in PTU_SUFFIXES:
                return local
        return ""

    def dragEnterEvent(self, event):
        if self._path_from_mime(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        path = self._path_from_mime(event.mimeData())
        if not path:
            event.ignore()
            return
        event.acceptProposedAction()
        self._start(path)

    def keyPressEvent(self, event):
        if event.matches(QtGui.QKeySequence.StandardKey.Paste):
            self._paste()
            event.accept()
            return
        super().keyPressEvent(event)

    def _browse(self):
        if self._worker is not None and self._worker.isRunning():
            return
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Open a PTU",
            "",
            "PTU files (*.pdf *.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp)",
        )
        if path:
            self._start(path)

    def _paste(self):
        mime = QtWidgets.QApplication.clipboard().mimeData()
        path = self._path_from_mime(mime)
        if path:
            self._start(path)
            return
        if mime is not None and mime.hasImage():
            image = CommentEditor._clipboard_image(mime)
            path = _write_temp_image(image)
            if path:
                self._start(path)
                return
            self._set_busy(False, "That picture could not be written to disk.")
            return
        self._set_busy(False, "The clipboard holds no picture or PDF file.")

    # --- reading ---

    def _ai_key(self):
        """The key to send with, or "" - with a prompt when it is missing."""
        key = self.ai_key_edit.text().strip()
        if not key:
            self.status_label.setText(
                "Paste a Gemini API key on the row above before reading a file."
            )
        return key

    def _start(self, path):
        if self._worker is not None and self._worker.isRunning():
            self.status_label.setText(
                "Still reading the previous file - give it a second."
            )
            return
        key = self._ai_key()
        if not key:
            return
        self.preview.clear()
        self._show_fields({})
        self._set_busy(
            True, f"Reading {os.path.basename(path)} with Gemini..."
        )
        # Keep it on this PC straight away (instant); the account copy waits for
        # a read that proves the key works, so a typo never lands in Firestore.
        self._remember_key(key)
        self._worker = PtuExtractionWorker(path, self, api_key=key)
        self._worker.done.connect(self._on_read)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _forget_key(self):
        """Delete the saved key here and on the account, then ask for another."""
        if self._worker is not None and self._worker.isRunning():
            self.status_label.setText(
                "Wait for the current read to finish before removing the key."
            )
            return
        answer = QtWidgets.QMessageBox.question(
            self,
            "Remove the Gemini key",
            "Delete the saved Gemini API key?\n\n"
            "It is removed from this PC and from your account, and the tool "
            "will ask for a key again before it can read a permit.",
        )
        if answer != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        removed_here = _delete_ai_key()
        removed_remote = _delete_ai_key_remote(self.user_data)
        self.ai_key_edit.clear()
        self._synced_key = ""
        if removed_here and removed_remote:
            self.status_label.setText("Gemini key removed from this PC and your account.")
        elif removed_here:
            self.status_label.setText(
                "Gemini key removed from this PC; your account could not be "
                "updated, so it may still be there."
            )
        elif removed_remote:
            self.status_label.setText("Gemini key removed from your account.")
        else:
            self.status_label.setText(
                "No saved key was found - paste one above to use the tool."
            )

    def _remember_key(self, key, verified=False):
        """Keep the key: on this PC always, on the account once it is proven.

        verified=False is the instant local save; verified=True is the Firestore
        write that a successful Gemini read earns, so a mistyped key is never
        pushed to the account.
        """
        if not key:
            return
        _save_ai_key(key)
        if verified and key != self._synced_key:
            self._synced_key = key
            _save_ai_key_remote(self.user_data, key)

    def _on_read(self, text, note, ai_fields=None):
        values = _ptu_extract_fields(text)
        if ai_fields:
            # The model's answer wins where it has one; anything it left empty
            # keeps the local result rather than losing it.
            for key, value in ai_fields.items():
                if value:
                    values[key] = value
            self._remember_key(self.ai_key_edit.text().strip(), verified=True)
        found = sum(1 for value in values.values() if value)
        summary = f"{found}/5 fields found."
        self._set_busy(False, f"{note} {summary}" if note else summary)
        if text:
            self.preview.setPlainText(text)
        elif ai_fields:
            self.preview.setPlainText(
                "No local text was read from this file. The fields above were "
                "read by Gemini."
            )
        else:
            self.preview.setPlainText("No text could be read from this file.")
        self._show_fields(values)

    def _show_fields(self, values):
        for key, _label, _pattern in PTU_FIELD_SPECS:
            edit = self.field_edits.get(key)
            if not edit or not _ui_alive(edit):
                continue
            edit.setText(values.get(key, ""))
            edit.setCursorPosition(0)

    def _on_failed(self, message):
        self._set_busy(False, f"Could not read that file: {message}")

    def _set_busy(self, busy, message=None):
        if _ui_alive(self.busy_bar):
            self.busy_bar.setVisible(busy)
        for widget in (
            self.browse_btn,
            self.paste_btn,
            self.drop_zone,
            self.ai_key_edit,
            self.getkey_btn,
            self.forget_btn,
        ):
            if _ui_alive(widget):
                widget.setEnabled(not busy)
        if message is not None and _ui_alive(self.status_label):
            self.status_label.setText(message)

    def _copy_text(self):
        text = self.preview.toPlainText()
        if not text:
            self.status_label.setText("Nothing to copy yet.")
            return
        QtWidgets.QApplication.clipboard().setText(text)
        self.status_label.setText("Copied the readable text to the clipboard.")

    # --- lifetime ---

    def _reading(self):
        return self._worker is not None and self._worker.isRunning()

    def reject(self):
        if self._reading():
            self.status_label.setText("Still reading - give it a second.")
            return
        super().reject()

    def closeEvent(self, event):
        if self._reading():
            event.ignore()
            self.status_label.setText("Still reading - give it a second.")
            return
        super().closeEvent(event)

    def _style(self):
        if is_windows_dark_mode():
            return """
            QDialog#PtuGenerator {
                background: #0f1923; color: #e6eef5;
                font-family: Inter, Segoe UI, Arial;
            }
            QLabel { color: #e6eef5; }
            QLabel#statusLabel { font-size: 12px; color: #98a0a6; }
            QLabel#dropLabel { font-size: 13px; color: #c9d4de; }
            QFrame#dropZone {
                background: rgba(255,255,255,0.03);
                border: 1px dashed rgba(255,255,255,0.22);
                border-radius: 12px;
            }
            QFrame#dropZone:hover { border-color: #ed3b3b; }
            QFrame#fieldsBox {
                background: rgba(255,255,255,0.03);
                border: 1px solid rgba(255,255,255,0.08);
                border-radius: 10px;
            }
            QLineEdit {
                background: rgba(255,255,255,0.02);
                border: 1px solid rgba(255,255,255,0.10);
                border-radius: 6px; padding: 5px 8px;
                color: #e6eef5; font-size: 13px;
                selection-background-color: rgba(237,59,59,0.45);
            }
            QTextEdit {
                background: rgba(255,255,255,0.04);
                border: 1px solid rgba(255,255,255,0.08);
                border-radius: 10px; padding: 8px;
                color: #e6eef5; font-size: 13px;
                selection-background-color: rgba(237,59,59,0.45);
            }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white; border-radius: 10px;
                padding: 8px 16px; font-size: 13px; font-weight: 600;
            }
            QPushButton:hover { background: #c62e2e; }
            QPushButton#ghost {
                background: transparent;
                border: 1px solid rgba(255,255,255,0.16);
                color: #e6eef5; padding: 7px 14px;
            }
            QPushButton#ghost:hover { border-color: #ed3b3b; color: #ffffff; }
            QPushButton:disabled { background: rgba(255,255,255,0.08); color: #98a0a6; }
            QProgressBar {
                background: rgba(255,255,255,0.08); border: none;
                border-radius: 2px;
            }
            QProgressBar::chunk { background: #ed3b3b; border-radius: 2px; }
            """
        return """
        QDialog#PtuGenerator {
            background: #f5f7fa; color: #1f2933;
            font-family: Inter, Segoe UI, Arial;
        }
        QLabel { color: #111827; }
        QLabel#statusLabel { font-size: 12px; color: #6b7280; }
        QLabel#dropLabel { font-size: 13px; color: #374151; }
        QFrame#dropZone {
            background: #ffffff;
            border: 1px dashed #cbd5e1;
            border-radius: 12px;
        }
        QFrame#dropZone:hover { border-color: #ed3b3b; }
        QFrame#fieldsBox {
            background: #ffffff; border: 1px solid #e5e7eb;
            border-radius: 10px;
        }
        QLineEdit {
            background: #f9fafb; border: 1px solid #e5e7eb;
            border-radius: 6px; padding: 5px 8px;
            color: #111827; font-size: 13px;
            selection-background-color: rgba(237,59,59,0.25);
        }
        QTextEdit {
            background: #ffffff;
            border: 1px solid #d1d5db;
            border-radius: 10px; padding: 8px;
            color: #111827; font-size: 13px;
            selection-background-color: rgba(237,59,59,0.25);
        }
        QPushButton {
            background: #ed3b3b; color: white; border-radius: 10px;
            padding: 8px 16px; font-size: 13px; font-weight: 600;
        }
        QPushButton:hover { background: #c62e2e; }
        QPushButton#ghost {
            background: transparent; border: 1px solid #d1d5db;
            color: #374151; padding: 7px 14px;
        }
        QPushButton#ghost:hover { border-color: #ed3b3b; color: #ed3b3b; }
        QPushButton:disabled { background: #d1d5db; color: #6b7280; }
        QProgressBar {
            background: #e5e7eb; border: none; border-radius: 2px;
        }
        QProgressBar::chunk { background: #ed3b3b; border-radius: 2px; }
        """


class LoginWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Report — Marilag")
        self.setMinimumSize(460, 600)
        self.resize(460, 600)
        self.setObjectName("Root")
        self.apply_style()
        self.main_layout = QtWidgets.QVBoxLayout(self)
        self.main_layout.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        self.show_login()

    def _theme_colors(self):
        """Return a dict of theme-aware colors for inline widget styling."""
        if is_windows_dark_mode():
            return {
                "outline_border": "rgba(255,255,255,0.18)",
                "outline_text": "#e6eef5",
                "outline_hover_border": "#ed3b3b",
                "outline_hover_text": "#ffffff",
                "cancel_hover_border": "#98a0a6",
                "divider": "rgba(255,255,255,0.08)",
                "field_bg": "#15212c",
                "field_border": "#2a3a48",
                "field_text": "#e6eef5",
                "field_focus_border": "#ed3b3b",
                "section_label": "#ed3b3b",
                "field_label": "#98a0a6",
                "back_color": "#98a0a6",
                "back_hover": "#ff4d4d",
                "output_bg": "#111a22",
                "output_text": "#e6eef5",
            }
        else:
            return {
                "outline_border": "#d1d5db",
                "outline_text": "#1f2933",
                "outline_hover_border": "#ed3b3b",
                "outline_hover_text": "#ed3b3b",
                "cancel_hover_border": "#6b7280",
                "divider": "#e5e7eb",
                "field_bg": "#ffffff",
                "field_border": "#d1d5db",
                "field_text": "#111827",
                "field_focus_border": "#ed3b3b",
                "section_label": "#ed3b3b",
                "field_label": "#374151",
                "back_color": "#6b7280",
                "back_hover": "#ed3b3b",
                "output_bg": "#ffffff",
                "output_text": "#111827",
            }

    def set_page_size(self, width, height):
        """
        Resize the window for the current page. Called at the top of
        each show_xxx_page() method so different pages can use
        different window sizes (e.g. POT's wider two-column form).
        """
        self.setMinimumSize(width, height)
        self.resize(width, height)

    def apply_style(self):
        if is_windows_dark_mode():

            # DARK THEME
            self.setStyleSheet("""
            #Root {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #07101a, stop:1 #07121a);
                font-family: Inter, Segoe UI, Arial;
                color: #e6eef5;
            }

            QLabel#Title {
                font-size: 22px;
                font-weight: 700;
                color: #f8fbfc;
            }

            QLabel#Subtitle {
                font-size: 13px;
                color: #98a0a6;
            }

            QWidget#Card {
                background: rgba(255,255,255,0.03);
                border: 1px solid rgba(255,255,255,0.05);
                border-radius: 14px;
            }

            QLineEdit {
                background: rgba(255,255,255,0.02);
                border: 1px solid rgba(255,255,255,0.06);
                border-radius: 10px;
                padding: 12px;
                font-size: 14px;
                color: #eaf6fb;
            }

            QLineEdit:focus {
                border: 1px solid #ed3b3b;
            }

            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #ed3b3b, stop:1 #c62e2e);
                border: none;
                border-radius: 10px;
                padding: 12px;
                font-size: 15px;
                font-weight: 700;
                color: white;
            }

            QPushButton:hover {
                opacity: 0.92;
            }

            QRadioButton {
                color: #e6eef5;
                spacing: 6px;
            }

            QRadioButton::indicator {
                width: 14px;
                height: 14px;
                border-radius: 7px;
                border: 2px solid #6b7280;
                background-color: transparent;
            }

            QRadioButton::indicator:checked {
                background-color: #DA3535;
                border: 2px solid #c62e2e;
            }

            QComboBox {
                background: rgba(255,255,255,0.02);
                border: 1px solid rgba(255,255,255,0.06);
                border-radius: 10px;
                padding: 12px 36px 12px 12px;
                font-size: 14px;
                color: #e6eef5;
                min-height: 20px;
            }

            QComboBox:focus {
                border: 1px solid #ed3b3b;
            }

            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 32px;
                border: none;
                border-left: 1px solid rgba(255,255,255,0.06);
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
            }

            QComboBox::down-arrow {
                width: 10px;
                height: 10px;
            }

            QComboBox QAbstractItemView {
                background: #15212c;
                border: 1px solid #2a3a48;
                border-radius: 8px;
                padding: 4px;
                color: #e6eef5;
                selection-background-color: #1e3a50;
                selection-color: #ffffff;
                outline: none;
            }

            QComboBox QAbstractItemView::item {
                padding: 8px 12px;
                border-radius: 4px;
                min-height: 24px;
            }

            QComboBox QAbstractItemView::item:hover {
                background: #1e3a50;
                color: #ffffff;
            }

            QDateEdit {
                background: rgba(255,255,255,0.02);
                border: 1px solid rgba(255,255,255,0.06);
                border-radius: 10px;
                padding: 12px 36px 12px 12px;
                font-size: 14px;
                color: #e6eef5;
                min-height: 20px;
            }

            QDateEdit:focus {
                border: 1px solid #ed3b3b;
            }

            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 32px;
                border: none;
                border-left: 1px solid rgba(255,255,255,0.06);
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
            }

            QDateEdit::down-arrow {
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #98a0a6;
            }
            """)

        else:

            # LIGHT THEME
            self.setStyleSheet("""
            #Root {
                background: #f5f7fa;
                font-family: Inter, Segoe UI, Arial;
                color: #1f2933;
            }

            QLabel#Title {
                font-size: 22px;
                font-weight: 700;
                color: #111827;
            }

            QLabel#Subtitle {
                font-size: 13px;
                color: #6b7280;
            }

            QWidget#Card {
                background: white;
                border: 1px solid #e5e7eb;
                border-radius: 14px;
            }

            QLineEdit {
                background: white;
                border: 1px solid #d1d5db;
                border-radius: 10px;
                padding: 12px;
                font-size: 14px;
                color: #111827;
            }

            QLineEdit:focus {
                border: 1px solid #ed3b3b;
            }

            QPushButton {
                background: #ed3b3b;
                border: none;
                border-radius: 10px;
                padding: 12px;
                font-size: 15px;
                font-weight: 700;
                color: white;
            }

            QPushButton:hover {
                background: #c62e2e;
            }

            QRadioButton {
                color: #374151;
                spacing: 6px;
            }

            QRadioButton::indicator {
                width: 14px;
                height: 14px;
                border-radius: 7px;
                border: 2px solid #9ca3af;
                background-color: transparent;
            }

            QRadioButton::indicator:checked {
                background-color: #ed3b3b;
                border: 2px solid #c62e2e;
            }

            QComboBox {
                background: white;
                border: 1px solid #d1d5db;
                border-radius: 10px;
                padding: 12px 36px 12px 12px;
                font-size: 14px;
                color: #111827;
                min-height: 20px;
            }

            QComboBox:focus {
                border: 1px solid #ed3b3b;
            }

            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 32px;
                border: none;
                border-left: 1px solid #d1d5db;
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
            }

            QComboBox::down-arrow {
                width: 10px;
                height: 10px;
            }

            QComboBox QAbstractItemView {
                background: white;
                border: 1px solid #d1d5db;
                border-radius: 8px;
                padding: 4px;
                color: #111827;
                selection-background-color: #f0f0f0;
                selection-color: #111827;
                outline: none;
            }

            QComboBox QAbstractItemView::item {
                padding: 8px 12px;
                border-radius: 4px;
                min-height: 24px;
            }

            QComboBox QAbstractItemView::item:hover {
                background: #f3f4f6;
                color: #111827;
            }

            QDateEdit {
                background: white;
                border: 1px solid #d1d5db;
                border-radius: 10px;
                padding: 12px 36px 12px 12px;
                font-size: 14px;
                color: #111827;
                min-height: 20px;
            }

            QDateEdit:focus {
                border: 1px solid #ed3b3b;
            }

            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 32px;
                border: none;
                border-left: 1px solid #d1d5db;
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
            }

            QDateEdit::down-arrow {
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #6b7280;
            }
            """)

    #DARK THEME
    

    def confirm_generate_report(self):
        dialog = QtWidgets.QMessageBox(self)
        dialog.setWindowTitle("Confirm Action")
        dialog.setText("Do you want to create a status report?")
        dialog.setIcon(QtWidgets.QMessageBox.Icon.Question)

        dialog.setStandardButtons(
            QtWidgets.QMessageBox.StandardButton.Yes |
            QtWidgets.QMessageBox.StandardButton.No
        )

        if is_windows_dark_mode():
            dialog.setStyleSheet("""
            QMessageBox {
                background: #07101a;
                font-family: Inter, Segoe UI, Arial;
                color: #e6eef5;
            }
            QLabel {
                color: #e6eef5;
                background: transparent;
            }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #ed3b3b, stop:1 #c62e2e);
                color: white;
                border-radius: 8px;
                padding: 6px 16px;
                min-width: 60px;
            }
            QPushButton:hover {
                background: #c62e2e;
            }
            """)
        else:
            dialog.setStyleSheet("""
            QMessageBox {
                background: #f5f7fa;
                font-family: Inter, Segoe UI, Arial;
                color: #1f2933;
            }
            QLabel {
                color: #111827;
                background: transparent;
            }
            QPushButton {
                background: #ed3b3b;
                color: white;
                border-radius: 8px;
                padding: 6px 16px;
                min-width: 60px;
            }
            QPushButton:hover {
                background: #c62e2e;
            }
            """)

        return dialog.exec() == QtWidgets.QMessageBox.StandardButton.Yes



    def clear_layout(self):
        while self.main_layout.count():
            item = self.main_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def build_card(self, title_text, subtitle_text, button_text, handler, pat=False, display=False, screenshot=None):
        card = QtWidgets.QWidget()
        card.setObjectName("Card")
        card.setFixedWidth(360)
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(16)

        # Optional screenshot at top
        if screenshot:
            pixmap = QtGui.QPixmap(screenshot)
            if not pixmap.isNull():
                pixmap = pixmap.scaledToWidth(320, QtCore.Qt.TransformationMode.SmoothTransformation)
                img_label = QtWidgets.QLabel()
                img_label.setPixmap(pixmap)
                img_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                img_label.setContentsMargins(0, 0, 0, 12)
                layout.addWidget(img_label)
            else:
                logger.warning("Screenshot not found: %s", screenshot)

        # Title
        title = QtWidgets.QLabel(title_text)
        title.setObjectName("Title")
        layout.addWidget(title)

        # Subtitle / description with word wrap and selectable text
        subtitle = QtWidgets.QLabel(subtitle_text)
        subtitle.setObjectName("Subtitle")
        subtitle.setWordWrap(True)  # Wrap long text
        subtitle.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)  # Make text copyable
        layout.addWidget(subtitle)

        layout.addSpacing(6)

        # Input fields
        if pat:
            self.pat_input = QtWidgets.QLineEdit()
            self.pat_input.setPlaceholderText("Enter your PAT token")
            layout.addWidget(self.pat_input)
        elif display:
            self.display_input = QtWidgets.QLineEdit()
            self.display_input.setPlaceholderText("Enter Azure Display Name")
            layout.addWidget(self.display_input)
        else:
            self.username = QtWidgets.QLineEdit()
            self.username.setPlaceholderText("Email")
            self.password = QtWidgets.QLineEdit()
            self.password.setPlaceholderText("Password")
            self.password.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
            layout.addWidget(self.username)
            layout.addWidget(self.password)

        layout.addSpacing(10)

        button = QtWidgets.QPushButton(button_text)
        button.clicked.connect(handler)
        button.setFixedHeight(40)

        # Store current page button so we can change text like "Creating Account..."
        self.action_button = button

        layout.addWidget(button)

        for field in card.findChildren(QtWidgets.QLineEdit):
            field.returnPressed.connect(button.click)

        self.main_layout.addWidget(card, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)



    # ----------------- PAGES -----------------
    def show_create_account(self):
        self.clear_layout()
        self.build_card(
            "Create your account",
            "Set a username and password for this app",
            "Next",
            self.next_to_pat
        )

    def next_to_pat(self):
        u = self.username.text().strip()
        p = self.password.text()
        if len(u) < 3 or len(p) < 6:
            AppDialog("Error", "Username must be 3+ chars\nPassword must be 6+ chars", self).exec()
            return
        self.temp_username = u
        self.temp_password = p
        self.show_pat_page()

    def show_pat_page(self):
        self.clear_layout()
        self.build_card(
            "Create PAT Token",
            "Create PAT by accessing the link below: \nhttps://tfs.alliancewebpos.com/tfs/WebPOSCollection/_usersSettings/tokens",
            "Next",           # button text
            self.save_pat,    # handler
            pat=True,
            screenshot=os.path.join(ASSET_DIR, "picture01.png")
        )

    def save_pat(self):
        pat = self.pat_input.text().strip()
        if len(pat) < 6:
            AppDialog("Error", "PAT must be at least 6 characters", self).exec()
            return
        self.temp_pat = pat
        self.show_display_name_page()

    def show_display_name_page(self):
        self.clear_layout()
        self.build_card(
            "Azure Display Name",
            "Enter the display name used in Azure for your account",
            "Finish Setup",
            self.save_display_name,
            display=True,
            screenshot=os.path.join(ASSET_DIR, "picture1.png")
        )

    # Replace the save_display_name method with this:
    def save_display_name(self):
        display_name = self.display_input.text().strip()
        if len(display_name) < 1:
            AppDialog("Error", "Display Name cannot be empty", self).exec()
            return

        self.temp_display_name = display_name  # store for next step
        self.show_unique_name_page()  # move to unique name page instead of saving

    def show_unique_name_page(self):
        self.clear_layout()
        self.build_card(
            "Unique Name",
            "Enter a unique name for this account (used for identification in the app)",
            "Finish Setup",
            self.save_unique_name,
            display=True,  # reuse the input field type
            screenshot=os.path.join(ASSET_DIR, "picture2.png")
        )

    def save_unique_name(self):
        unique_name = self.display_input.text().strip()

        if len(unique_name) < 1:
            AppDialog("Error", "Unique Name cannot be empty", self).exec()
            return

        logger.info("Account creation attempt for email=%r, display_name=%r", self.temp_username, self.temp_display_name)
        try:
            self.action_button.setEnabled(False)
            self.action_button.setText("Creating Account...")
            QApplication.processEvents()

            session = _sign_up(self.temp_username, self.temp_password)
            uid = session["uid"]
            id_token = session["id_token"]
            logger.info("Firebase user created, uid=%s", uid)

            self.action_button.setText("Saving User Data...")
            QApplication.processEvents()

            # Stored on this PC (DPAPI) and, when syncing is enabled, also in
            # users/{uid} so the account can be used on another PC.
            if not _save_pat_local(uid, self.temp_pat):
                logger.error("PAT could not be stored locally for uid=%s", uid)
                self.action_button.setEnabled(True)
                self.action_button.setText("Finish Setup")
                AppDialog(
                    "Error",
                    "Your PAT could not be stored securely on this PC.\n"
                    "Please try again.",
                    self,
                ).exec()
                return

            _save_pat_remote(uid, id_token, self.temp_pat)

            _update_profile(uid, id_token, {
                "display_name": self.temp_display_name,
                "unique_name": unique_name,
            })
            logger.info("Profile saved for uid=%s", uid)

            # Try to fetch PAT expiry and save it
            self.action_button.setText("Checking PAT expiry...")
            QApplication.processEvents()
            pat_expiry = _fetch_pat_expiry(self.temp_pat)
            if pat_expiry:
                _update_profile(uid, id_token, {"pat_expiry": pat_expiry})
                logger.info("PAT expiry saved: %s", pat_expiry)
            else:
                logger.info("PAT expiry could not be determined (will retry on login)")

            self.action_button.setText("Success!")
            QApplication.processEvents()

            AppDialog("Success", "Account created successfully!", self).exec()
            self.show_login()

        except Exception as e:
            self.action_button.setEnabled(True)
            self.action_button.setText("Finish Setup")

            error_message = str(e)
            logger.warning("Account creation failed for %r: %s", self.temp_username, error_message)

            if "EMAIL_EXISTS" in error_message:
                friendly_message = "This email is already registered.\nPlease login instead."
            elif "INVALID_EMAIL" in error_message:
                friendly_message = "Please enter a valid email address."
            elif "WEAK_PASSWORD" in error_message:
                friendly_message = "Password is too weak.\nPlease use at least 6 characters."
            else:
                friendly_message = "Registration failed.\nPlease try again later."

            AppDialog("Registration Error", friendly_message, self).exec()

    def show_main_menu_page(self):
        self.clear_layout()
        self.set_page_size(460, 620)

        card = QtWidgets.QWidget()
        card.setObjectName("Card")
        card.setFixedWidth(380)

        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(36, 36, 36, 32)
        layout.setSpacing(0)

        tc = self._theme_colors()

        # --- Top bar: Eyebrow + Settings gear icon ---
        top_bar = QtWidgets.QHBoxLayout()
        top_bar.setSpacing(0)

        eyebrow = QtWidgets.QLabel("WORKSPACE")
        eyebrow.setStyleSheet("""
            color: #ed3b3b;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 2px;
        """)
        top_bar.addWidget(eyebrow)
        top_bar.addStretch()

        settings_icon = QtWidgets.QToolButton()
        settings_icon.setText("\u2699")
        settings_icon.setToolTip("Settings")
        settings_icon.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        settings_icon.clicked.connect(self.show_settings_page)
        settings_icon.setStyleSheet(f"""
            QToolButton {{
                border: none;
                font-size: 18px;
                color: {tc['back_color']};
                padding: 0;
            }}
            QToolButton:hover {{
                color: {tc['back_hover']};
            }}
        """)
        top_bar.addWidget(settings_icon)

        layout.addLayout(top_bar)
        layout.addSpacing(8)

        # --- Title ---
        title = QtWidgets.QLabel("What are you creating?")
        title.setObjectName("Title")
        title.setWordWrap(True)
        layout.addWidget(title)
        layout.addSpacing(6)

        # --- Subtitle ---
        subtitle = QtWidgets.QLabel("Pick a document type to get started.")
        subtitle.setObjectName("Subtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        layout.addSpacing(28)

        # --- Divider ---
        divider = QtWidgets.QFrame()
        divider.setFixedHeight(1)
        divider.setStyleSheet(f"background: {tc['divider']}; border: none;")
        layout.addWidget(divider)
        layout.addSpacing(24)

        # --- Primary action: Create Status Report ---
        status_btn = QtWidgets.QPushButton("Create Status Report")
        status_btn.setFixedHeight(46)
        status_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        status_btn.clicked.connect(self.show_report_page)
        layout.addWidget(status_btn)

        layout.addSpacing(12)

        # --- Secondary action: Create POT (outlined style) ---
        pot_btn = QtWidgets.QPushButton("Create POT")
        pot_btn.setFixedHeight(46)
        pot_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        pot_btn.clicked.connect(self.show_pot_page)
        pot_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: 1.5px solid {tc['outline_border']};
                border-radius: 10px;
                padding: 12px;
                font-size: 15px;
                font-weight: 600;
                color: {tc['outline_text']};
            }}
            QPushButton:hover {{
                border: 1.5px solid {tc['outline_hover_border']};
                color: {tc['outline_hover_text']};
            }}
        """)
        layout.addWidget(pot_btn)

        layout.addSpacing(12)

        # --- Tertiary action: Tools (outlined style) ---
        tools_btn = QtWidgets.QPushButton("Tools")
        tools_btn.setFixedHeight(46)
        tools_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        tools_btn.clicked.connect(self.show_tools_page)
        tools_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: 1.5px solid {tc['outline_border']};
                border-radius: 10px;
                padding: 12px;
                font-size: 15px;
                font-weight: 600;
                color: {tc['outline_text']};
            }}
            QPushButton:hover {{
                border: 1.5px solid {tc['outline_hover_border']};
                color: {tc['outline_hover_text']};
            }}
        """)
        layout.addWidget(tools_btn)

        # --- Footer: signed in as ---
        if hasattr(self, "current_user") and self.current_user.get("display_name"):
            layout.addSpacing(28)
            footer_divider = QtWidgets.QFrame()
            footer_divider.setFixedHeight(1)
            footer_divider.setStyleSheet(f"background: {tc['divider']}; border: none;")
            layout.addWidget(footer_divider)
            layout.addSpacing(14)

            footer = QtWidgets.QLabel(f"Signed in as {self.current_user['display_name']}")
            footer.setStyleSheet("color: #6b7280; font-size: 12px;")
            footer.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(footer)

            build_label = QtWidgets.QLabel(f"Release v{APP_VERSION}")
            build_label.setStyleSheet("color: #6b7280; font-size: 11px;")
            build_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(build_label)

        self.main_layout.addWidget(
            card,
            alignment=QtCore.Qt.AlignmentFlag.AlignCenter
        )


    def show_settings_page(self):
        self.clear_layout()
        self.set_page_size(460, 660)

        card = QtWidgets.QWidget()
        card.setObjectName("Card")
        card.setFixedWidth(380)

        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(36, 36, 36, 32)
        layout.setSpacing(0)

        tc = self._theme_colors()

        # --- Top Bar (Back Arrow) ---
        top_bar = QtWidgets.QHBoxLayout()
        back_btn = QtWidgets.QToolButton()
        back_btn.setText("\u2190")
        back_btn.setToolTip("Back to Main Menu")
        back_btn.clicked.connect(self.show_main_menu_page)
        back_btn.setStyleSheet(f"""
            QToolButton {{
                border: none;
                font-size: 18px;
                font-weight: bold;
                color: {tc['back_color']};
            }}
            QToolButton:hover {{
                color: {tc['back_hover']};
            }}
        """)
        back_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        top_bar.addWidget(back_btn)
        top_bar.addStretch()
        layout.addLayout(top_bar)
        layout.addSpacing(4)

        # --- Title ---
        title = QtWidgets.QLabel("Settings")
        title.setObjectName("Title")
        layout.addWidget(title)
        layout.addSpacing(6)

        subtitle = QtWidgets.QLabel("Manage your account and PAT settings.")
        subtitle.setObjectName("Subtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        layout.addSpacing(28)

        # --- Divider ---
        divider = QtWidgets.QFrame()
        divider.setFixedHeight(1)
        divider.setStyleSheet(f"background: {tc['divider']}; border: none;")
        layout.addWidget(divider)
        layout.addSpacing(20)

        # --- PAT Expiry Section ---
        section_label = QtWidgets.QLabel("PAT EXPIRY")
        section_label.setStyleSheet(f"""
            color: {tc['section_label']};
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 1.5px;
        """)
        layout.addWidget(section_label)
        layout.addSpacing(12)

        field_label = QtWidgets.QLabel("PAT Token Expiry Date")
        field_label.setStyleSheet(f"color: {tc['field_label']}; font-size: 12px; font-weight: 600;")
        layout.addWidget(field_label)
        layout.addSpacing(6)

        self.settings_expiry_input = QtWidgets.QDateEdit()
        self.settings_expiry_input.setCalendarPopup(True)
        self.settings_expiry_input.setDate(QtCore.QDate.currentDate().addYears(1))
        self.settings_expiry_input.setStyleSheet(f"""
            QDateEdit {{
                background-color: {tc['field_bg']};
                border: 1px solid {tc['field_border']};
                border-radius: 8px;
                padding: 9px 36px 9px 10px;
                color: {tc['field_text']};
                font-size: 13px;
            }}
            QDateEdit:focus {{
                border: 1px solid {tc['field_focus_border']};
            }}
            QDateEdit::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 28px;
                border: none;
                border-left: 1px solid {tc['field_border']};
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }}
            QDateEdit::down-arrow {{
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid {tc['field_label']};
            }}
        """)
        layout.addWidget(self.settings_expiry_input)

        layout.addSpacing(8)

        # Show current saved expiry
        current_expiry = self.current_user.get("pat_expiry", "") if hasattr(self, "current_user") else ""
        if current_expiry:
            expiry_display = QtWidgets.QLabel(f"Current: {current_expiry}")
            expiry_display.setStyleSheet("color: #6b7280; font-size: 11px;")
            layout.addWidget(expiry_display)
            # Try to pre-fill the date picker with the saved value
            try:
                dt = datetime.fromisoformat(current_expiry.replace("Z", "+00:00"))
                self.settings_expiry_input.setDate(QtCore.QDate(dt.year, dt.month, dt.day))
            except Exception:
                pass
        else:
            expiry_display = QtWidgets.QLabel("No expiry date saved yet")
            expiry_display.setStyleSheet("color: #6b7280; font-size: 11px;")
            layout.addWidget(expiry_display)

        layout.addSpacing(24)

        # --- Save Button ---
        btn_layout = QtWidgets.QHBoxLayout()
        save_btn = QtWidgets.QPushButton("Save")
        save_btn.setFixedHeight(42)
        save_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(self.save_pat_expiry)
        save_btn.setFixedWidth(120)

        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

        layout.addSpacing(12)

        # --- Status label ---
        self.settings_status = QtWidgets.QLabel("")
        self.settings_status.setWordWrap(True)
        if is_windows_dark_mode():
            self.settings_status.setStyleSheet("color: #4ade80; font-weight: 600; font-size: 12.5px;")
        else:
            self.settings_status.setStyleSheet("color: #16a34a; font-weight: 600; font-size: 12.5px;")
        layout.addWidget(self.settings_status)

        # --- Application / updates ---
        layout.addSpacing(24)

        divider2 = QtWidgets.QFrame()
        divider2.setFixedHeight(1)
        divider2.setStyleSheet(f"background: {tc['divider']}; border: none;")
        layout.addWidget(divider2)
        layout.addSpacing(20)

        app_label = QtWidgets.QLabel("APPLICATION")
        app_label.setStyleSheet(f"""
            color: {tc['section_label']};
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 1.5px;
        """)
        layout.addWidget(app_label)
        layout.addSpacing(12)

        version_label = QtWidgets.QLabel(f"Installed version: v{APP_VERSION}")
        version_label.setStyleSheet(f"color: {tc['field_label']}; font-size: 12px;")
        layout.addWidget(version_label)
        layout.addSpacing(10)

        update_btn_layout = QtWidgets.QHBoxLayout()
        self.check_update_btn = QtWidgets.QPushButton("Check for Updates")
        self.check_update_btn.setFixedHeight(38)
        self.check_update_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.check_update_btn.clicked.connect(self.check_for_updates_manually)
        update_btn_layout.addWidget(self.check_update_btn)
        update_btn_layout.addStretch()
        layout.addLayout(update_btn_layout)

        layout.addSpacing(8)

        self.update_status_label = QtWidgets.QLabel("")
        self.update_status_label.setWordWrap(True)
        self.update_status_label.setStyleSheet("color: #6b7280; font-size: 11.5px;")
        layout.addWidget(self.update_status_label)

        self.main_layout.addWidget(card, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)


    def save_pat_expiry(self):
        expiry_date = self.settings_expiry_input.date()
        expiry_str = expiry_date.toString("yyyy-MM-dd") + "T00:00:00Z"

        uid = self.current_user.get("uid", "") if hasattr(self, "current_user") else ""
        if not uid:
            self.settings_status.setText("Not logged in.")
            return

        logger.info("Saving PAT expiry: %s", expiry_str)
        try:
            _update_user_profile(self.current_user, {"pat_expiry": expiry_str})
            self.current_user["pat_expiry"] = expiry_str
            logger.info("PAT expiry saved for uid=%s", uid)
            self.settings_status.setText(f"Saved! Expiry: {expiry_date.toString('MM/dd/yyyy')}")
        except FirebaseError as e:
            logger.exception("Failed to save PAT expiry")
            self.settings_status.setText(str(e))
        except Exception:
            logger.exception("Failed to save PAT expiry")
            self.settings_status.setText("Failed to save. Check logs.")


    def check_for_updates_manually(self):
        """Settings button: look for a release now instead of on next start."""
        if getattr(self, "_update_check_running", False):
            return
        self._update_check_running = True
        self.check_update_btn.setEnabled(False)
        self.check_update_btn.setText("Checking...")
        self.update_status_label.setStyleSheet("color: #6b7280; font-size: 11.5px;")
        self.update_status_label.setText("Looking for a newer release...")

        self._update_worker = UpdateCheckWorker(self)
        self._update_worker.found.connect(self._on_manual_update_result)
        self._update_worker.finished.connect(self._on_manual_update_finished)
        self._update_worker.start()

    def _on_manual_update_finished(self):
        self._update_check_running = False
        if hasattr(self, "check_update_btn"):
            self.check_update_btn.setEnabled(True)
            self.check_update_btn.setText("Check for Updates")

    def _on_manual_update_result(self, update):
        if not hasattr(self, "update_status_label"):
            return
        if not update:
            self.update_status_label.setText(
                f"You are on the latest version (v{APP_VERSION})."
            )
            return

        self.update_status_label.setText(f"Version {update['version']} is available.")
        self._update_info = update
        dlg = UpdateDialog(update["version"], update.get("body", ""), parent=self)
        dlg.exec()


    # ================= TOOLS PAGE =================
    def show_tools_page(self):
        self.clear_layout()
        self.set_page_size(480, 620)

        tc = self._theme_colors()

        outer = QtWidgets.QVBoxLayout()
        outer.setContentsMargins(24, 20, 24, 20)
        outer.setSpacing(0)

        # --- Top Bar (Back Arrow) ---
        top_bar = QtWidgets.QHBoxLayout()
        back_btn = QtWidgets.QToolButton()
        back_btn.setText("←")
        back_btn.setToolTip("Back to Main Menu")
        back_btn.clicked.connect(self.show_main_menu_page)
        back_btn.setStyleSheet(f"""
            QToolButton {{
                border: none;
                font-size: 18px;
                font-weight: bold;
                color: {tc['back_color']};
            }}
            QToolButton:hover {{
                color: {tc['back_hover']};
            }}
        """)
        back_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        top_bar.addWidget(back_btn)
        top_bar.addStretch()
        outer.addLayout(top_bar)
        outer.addSpacing(4)

        # --- Title ---
        title = QtWidgets.QLabel("Tools")
        title.setObjectName("Title")
        outer.addWidget(title)
        outer.addSpacing(2)

        subtitle = QtWidgets.QLabel(f"{len(TOOL_REGISTRY)} tools available")
        subtitle.setObjectName("Subtitle")
        outer.addWidget(subtitle)
        outer.addSpacing(12)

        # --- Search bar ---
        self.tools_search = QtWidgets.QLineEdit()
        self.tools_search.setPlaceholderText("Search tools...")
        self.tools_search.setClearButtonEnabled(True)
        self.tools_search.setStyleSheet(f"""
            QLineEdit {{
                background-color: {tc['field_bg']};
                border: 1px solid {tc['field_border']};
                border-radius: 8px;
                padding: 8px 10px;
                color: {tc['field_text']};
                font-size: 13px;
            }}
            QLineEdit:focus {{
                border: 1px solid {tc['field_focus_border']};
            }}
        """)
        outer.addWidget(self.tools_search)
        outer.addSpacing(12)

        # --- Scrollable list ---
        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("""
            QScrollArea { border: none; background: transparent; }
            QScrollBar:vertical { width: 5px; background: transparent; }
            QScrollBar::handle:vertical { background: rgba(128,128,128,0.3); border-radius: 2px; min-height: 20px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
        """)

        scroll_widget = QtWidgets.QWidget()
        scroll_widget.setStyleSheet("background: transparent;")
        self.tools_list_layout = QtWidgets.QVBoxLayout(scroll_widget)
        self.tools_list_layout.setSpacing(6)
        self.tools_list_layout.setContentsMargins(2, 2, 2, 2)

        self._populate_tools_list(tc)

        scroll_area.setWidget(scroll_widget)
        outer.addWidget(scroll_area)

        container = QtWidgets.QWidget()
        container.setLayout(outer)
        self.main_layout.addWidget(container)

        # --- Connect search to filter ---
        self.tools_search.textChanged.connect(lambda text: self._filter_tools(text, tc))

    def _populate_tools_list(self, tc, filter_text=""):
        """Populate the tools list from TOOL_REGISTRY. Auto-fetches all tools."""
        while self.tools_list_layout.count():
            item = self.tools_list_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        found = 0
        for tool in TOOL_REGISTRY:
            if filter_text.lower() not in tool["name"].lower() and \
               filter_text.lower() not in tool.get("description", "").lower() and \
               filter_text.lower() not in tool.get("category", "").lower():
                continue

            row = self._create_tool_row(tool, tc)
            self.tools_list_layout.addWidget(row)
            found += 1

        if found == 0:
            no_results = QtWidgets.QLabel("No tools found")
            no_results.setStyleSheet(f"color: {tc['back_color']}; font-size: 13px; padding: 20px;")
            no_results.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.tools_list_layout.addWidget(no_results)

        self.tools_list_layout.addStretch()

    def _filter_tools(self, text, tc):
        """Re-populate list when search text changes."""
        self._populate_tools_list(tc, filter_text=text)

    def _create_tool_row(self, tool, tc):
        """Create a compact horizontal row for a single tool."""
        row = QtWidgets.QFrame()
        row.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        row.setMinimumHeight(56)
        row.setStyleSheet(f"""
            QFrame {{
                background-color: {tc['field_bg']};
                border: 1px solid {tc['field_border']};
                border-radius: 8px;
            }}
            QFrame:hover {{
                border: 1px solid {tc['field_focus_border']};
            }}
        """)

        row_layout = QtWidgets.QHBoxLayout(row)
        row_layout.setContentsMargins(14, 8, 14, 8)
        row_layout.setSpacing(12)

        # Icon
        icon_label = QtWidgets.QLabel(tool.get("icon", "🔧"))
        icon_label.setFixedWidth(28)
        icon_label.setStyleSheet("font-size: 18px; background: transparent; border: none;")
        row_layout.addWidget(icon_label)

        # Text block (name + description stacked)
        text_block = QtWidgets.QVBoxLayout()
        text_block.setSpacing(2)

        name_label = QtWidgets.QLabel(tool["name"])
        name_label.setStyleSheet(f"""
            font-size: 13px; font-weight: 600; color: {tc['field_text']};
            background: transparent; border: none;
        """)
        text_block.addWidget(name_label)

        desc_label = QtWidgets.QLabel(tool.get("description", ""))
        desc_label.setStyleSheet(f"""
            font-size: 11px; color: {tc['back_color']};
            background: transparent; border: none;
        """)
        desc_label.setWordWrap(True)
        text_block.addWidget(desc_label)

        row_layout.addLayout(text_block, stretch=1)

        # Arrow indicator
        arrow = QtWidgets.QLabel("›")
        arrow.setStyleSheet(f"""
            font-size: 18px; color: {tc['back_color']};
            background: transparent; border: none;
        """)
        row_layout.addWidget(arrow)

        # Click handler
        handler_name = tool.get("handler", "")
        available = tool.get("available", tool.get("id") not in UNFINISHED_TOOL_IDS)
        if not available:
            # A placeholder row that looks identical to a working one is a lie
            # to the user: mute it, label it, and leave it inert.
            name_label.setStyleSheet(f"""
                font-size: 13px; font-weight: 600; color: {tc['back_color']};
                background: transparent; border: none;
            """)
            desc_label.setText(f"{tool.get('description', '')} (coming soon)")
            row.setCursor(QtCore.Qt.CursorShape.ArrowCursor)
            row.mousePressEvent = lambda event: None
            arrow.setText("")
        else:
            row.mousePressEvent = lambda event, h=handler_name: self._on_tool_clicked(h)

        return row

    def _on_tool_clicked(self, handler_name):
        """Dispatch to the tool's handler method on this window."""
        if not handler_name:
            AppDialog("Coming Soon", "This tool is not yet implemented.", self).exec()
            return
        handler = getattr(self, handler_name, None)
        if handler and callable(handler):
            logger.info("Tool clicked: %s", handler_name)
            handler()
        else:
            AppDialog("Error", f"Handler '{handler_name}' not found.", self).exec()

    # ================= TOOL HANDLERS =================
    # Implement each tool's logic below.
    # The handler name must match the "handler" key in TOOL_REGISTRY.

    def run_tool_1(self):
        dialog = MissingSalesDialog(self)
        dialog.exec()

    def run_tool_2(self):
        """Ayala Checker - Compare hourly sales vs EOD"""
        dialog = AyalaCheckerDialog(self)
        dialog.exec()

    def run_scheduled_comment(self):
        """Scheduled Comment - compose now, posted to TFS by the hosted worker."""
        user_data = getattr(self, "current_user", None)
        if not user_data:
            AppDialog("Not Signed In", "Sign in before scheduling a comment.", self).exec()
            return
        ScheduledCommentDialog(user_data, self).exec()

    def run_ptu_generator(self):
        """PTU Generator - read a permit from a PDF or an image with Gemini.

        The key comes from this PC first, then from users/{uid} in Firestore, so
        signing in on a new machine is enough to open the tool. Only when neither
        has one does the gate appear and ask for it.
        """
        user_data = getattr(self, "current_user", None)
        note = ""
        key = _load_ai_key()
        if not key and user_data:
            try:
                key = _load_ai_key_remote(user_data)
            except Exception as e:
                logger.warning("Could not read the Gemini key from Firestore: %s", e)
                note = (
                    f"Your account could not be checked for a saved key ({e}). "
                    "If you already saved one there, sign in again or try when "
                    "you are online."
                )
            if key:
                # Cache it so the next open needs no network.
                _save_ai_key(key)
        if not key:
            gate = PtuApiKeyDialog(self, user_data=user_data, note=note)
            if gate.exec() != QtWidgets.QDialog.DialogCode.Accepted:
                return
            key = gate.saved_key
            if not key:
                return
        PtuGeneratorDialog(self, api_key=key, user_data=user_data).exec()

    def run_tool_3(self):
        """Tool 3 — TODO: implement"""
        AppDialog("Tool 3", "Tool 3 is not yet implemented.", self).exec()

    def run_tool_4(self):
        """Tool 4 — TODO: implement"""
        AppDialog("Tool 4", "Tool 4 is not yet implemented.", self).exec()

    def run_tool_5(self):
        """Tool 5 — TODO: implement"""
        AppDialog("Tool 5", "Tool 5 is not yet implemented.", self).exec()

    def run_tool_6(self):
        """Tool 6 — TODO: implement"""
        AppDialog("Tool 6", "Tool 6 is not yet implemented.", self).exec()

    def run_tool_7(self):
        """Tool 7 — TODO: implement"""
        AppDialog("Tool 7", "Tool 7 is not yet implemented.", self).exec()

    def run_tool_8(self):
        """Tool 8 — TODO: implement"""
        AppDialog("Tool 8", "Tool 8 is not yet implemented.", self).exec()

    def run_tool_9(self):
        """Tool 9 — TODO: implement"""
        AppDialog("Tool 9", "Tool 9 is not yet implemented.", self).exec()

    def run_tool_10(self):
        """Tool 10 — TODO: implement"""
        AppDialog("Tool 10", "Tool 10 is not yet implemented.", self).exec()


    def show_login(self):
        self.clear_layout()
        self.set_page_size(460, 600)

        card = QtWidgets.QWidget()
        card.setObjectName("Card")
        card.setFixedWidth(360)

        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(16)

        title = QtWidgets.QLabel("Sign in")
        title.setObjectName("Title")
        layout.addWidget(title)

        subtitle = QtWidgets.QLabel("Enter your email and password to continue")
        subtitle.setObjectName("Subtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        self.username = QtWidgets.QLineEdit()
        self.username.setPlaceholderText("Email")

        self.password = QtWidgets.QLineEdit()
        self.password.setPlaceholderText("Password")
        self.password.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)

        layout.addWidget(self.username)
        layout.addWidget(self.password)

        self.login_btn = QtWidgets.QPushButton("Login")
        self.login_btn.clicked.connect(self.login)
        self.login_btn.setFixedHeight(40)
        layout.addWidget(self.login_btn)

        register_label = QtWidgets.QLabel(
            'Don\'t have an account? '
            '<a href="#">Create one</a>'
        )

        register_label.setObjectName("Subtitle")
        register_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        register_label.setTextFormat(QtCore.Qt.TextFormat.RichText)
        register_label.setTextInteractionFlags(
            QtCore.Qt.TextInteractionFlag.TextBrowserInteraction
        )

        register_label.setOpenExternalLinks(False)

        register_label.linkActivated.connect(
            lambda _: self.show_create_account()
        )

        layout.addWidget(register_label)

        build_label = QtWidgets.QLabel(f"Release v{APP_VERSION}")
        build_label.setStyleSheet("color: #6b7280; font-size: 11px;")
        build_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(build_label)

        self.username.returnPressed.connect(self.login_btn.click)
        self.password.returnPressed.connect(self.login_btn.click)

        self.main_layout.addWidget(
            card,
            alignment=QtCore.Qt.AlignmentFlag.AlignCenter
        )

    def login(self):
        email = self.username.text().strip()
        password = self.password.text()

        if not email or not password:
            AppDialog("Error", "Email and password cannot be empty.", self).exec()
            return

        logger.info("Login attempt for email=%r", email)
        self.login_btn.setEnabled(False)
        self.login_btn.setText("Signing in...")

        # Auth, profile and PAT expiry all happen on the worker: doing them here
        # froze the window (three network round-trips) on every sign-in.
        self._login_worker = LoginWorker(email, password, self)
        self._login_worker.succeeded.connect(self._on_login_succeeded)
        self._login_worker.pat_required.connect(self._on_login_needs_pat)
        self._login_worker.failed.connect(self._on_login_failed)
        self._login_worker.start()

    def _reset_login_button(self):
        if _ui_alive(getattr(self, "login_btn", None)):
            self.login_btn.setEnabled(True)
            self.login_btn.setText("Login")

    def _on_login_succeeded(self, user_data):
        self.current_user = user_data
        if _ui_alive(getattr(self, "password", None)):
            self.password.clear()
        logger.info("Signed in as %r (uid=%s)",
                    user_data.get("email"), user_data.get("uid"))

        # The user may have navigated away while the worker was running.
        if not _ui_alive(getattr(self, "login_btn", None)):
            return

        AppDialog(
            "Welcome",
            f"Hello {user_data.get('display_name', '')}!\nYou have successfully logged in.",
            self
        ).exec()
        self.show_main_menu_page()

    def _on_login_needs_pat(self, pending):
        """No PAT anywhere - a brand-new account, or one created before syncing
        was enabled. Ask once, keep it locally, and publish it to users/{uid}."""
        if not _ui_alive(getattr(self, "login_btn", None)):
            return
        self._reset_login_button()
        pat, ok = QtWidgets.QInputDialog.getText(
            self,
            "Personal Access Token",
            "Your TFS token is not stored on this PC.\nEnter it to continue:",
            QtWidgets.QLineEdit.EchoMode.Password,
        )
        if not ok or not pat.strip():
            return
        if not _save_pat_local(pending.get("uid", ""), pat.strip()):
            AppDialog("Error", "The PAT could not be stored securely on this PC.", self).exec()
            return
        _save_pat_remote(pending.get("uid", ""), pending.get("id_token", ""), pat.strip())
        pending["pat"] = pat.strip()
        self._on_login_succeeded(pending)

    def _on_login_failed(self, error_message):
        self._reset_login_button()
        logger.warning("Login failed: %s", error_message)

        if "INVALID_LOGIN_CREDENTIALS" in error_message or "INVALID_PASSWORD" in error_message:
            friendly_message = "Invalid email or password."
        elif error_message == "NETWORK" or "Could not reach" in error_message:
            friendly_message = "Could not reach the internet.\nCheck your connection and try again."
        elif error_message == "TOO_MANY_ATTEMPTS_TRY_LATER":
            friendly_message = "Too many attempts.\nWait a few minutes and try again."
        elif "EMAIL_NOT_FOUND" in error_message:
            friendly_message = "No account found with this email."
        elif "USER_DISABLED" in error_message:
            friendly_message = "This account has been disabled."
        elif "USER_PROFILE_NOT_FOUND" in error_message:
            friendly_message = "Account found, but profile data is missing."
        elif error_message.startswith(("The database", "Database error", "Could not reach")):
            # Already written for a human - hiding it is what made a rules
            # problem look like a wrong password.
            friendly_message = error_message
        else:
            friendly_message = "Login failed.\nPlease try again later."

        AppDialog("Login Error", friendly_message, self).exec()

    def handle_report_finished(self, msg):
        if not _ui_alive(getattr(self, "output_area", None)):
            logger.info("Report finished after the page was closed: %s", msg)
            return
        self.output_area.append(msg)

        # Re-enable inputs
        self.generate_report_btn.setEnabled(True)
        self.generate_report_btn.setText("Generate Report")
        self.work_item_input.setEnabled(True)
        self.client_input.setEnabled(True)
        self.onsite_radio.setEnabled(True)
        self.offsite_radio.setEnabled(True)

        if msg.startswith("✅ Report saved"):
            # Show and enable the "Create Azure Status Report" button
            if hasattr(self, "create_azure_btn") and self.create_azure_btn is not None:
                self.create_azure_btn.setVisible(True)
                self.create_azure_btn.setEnabled(True)

            # Save report data for later use in PBI creation
            self.last_report_data = {
                "work_item_id": self.worker.work_item_id,
                "client_name": self.worker.client_name,
                "root_cause": getattr(self.worker, "root_cause", ""),
                "preventive_action": getattr(self.worker, "preventive_action", ""),
                "next_step": getattr(self.worker, "next_step", ""),
                "raw_html": getattr(self.worker, "raw_html", "")
            }

    def show_pot_page(self):
        self.clear_layout()
        self.set_page_size(700, 760)

        # Card container
        card = QtWidgets.QWidget()
        card.setObjectName("Card")
        card.setFixedWidth(620)

        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(32, 28, 32, 28)
        layout.setSpacing(0)

        tc = self._theme_colors()

        # ===== Top Bar (Back Arrow + Folder Icon) =====
        top_bar = QtWidgets.QHBoxLayout()
        back_btn = QtWidgets.QToolButton()
        back_btn.setText("←")
        back_btn.setToolTip("Back to Main Menu")
        back_btn.clicked.connect(self.show_main_menu_page)
        back_btn.setStyleSheet(f"""
            QToolButton {{
                border: none;
                font-size: 18px;
                font-weight: bold;
                color: {tc['back_color']};
            }}
            QToolButton:hover {{
                color: {tc['back_hover']};
            }}
        """)
        back_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)

        self.pot_folder_btn = QtWidgets.QToolButton()
        self.pot_folder_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.StandardPixmap.SP_DirOpenIcon))
        self.pot_folder_btn.setToolTip("Open POT Folder")
        self.pot_folder_btn.clicked.connect(self.open_report_folder)
        self.pot_folder_btn.setEnabled(False)

        top_bar.addWidget(back_btn)
        top_bar.addStretch()
        top_bar.addWidget(self.pot_folder_btn)
        layout.addLayout(top_bar)
        layout.addSpacing(4)

        # ===== Title & Subtitle =====
        title = QtWidgets.QLabel("Create POT")
        title.setObjectName("Title")
        layout.addWidget(title)
        layout.addSpacing(4)

        subtitle = QtWidgets.QLabel("Fill in the details below to generate a Proof of Test document")
        subtitle.setObjectName("Subtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        layout.addSpacing(24)

        # Shared field style so every input matches regardless of section
        field_style = f"""
            QLineEdit {{
                background-color: {tc['field_bg']};
                border: 1px solid {tc['field_border']};
                border-radius: 8px;
                padding: 9px 10px;
                color: {tc['field_text']};
                font-size: 13px;
            }}
            QLineEdit:focus {{
                border: 1px solid {tc['field_focus_border']};
            }}
            QDateEdit {{
                background-color: {tc['field_bg']};
                border: 1px solid {tc['field_border']};
                border-radius: 8px;
                padding: 9px 36px 9px 10px;
                color: {tc['field_text']};
                font-size: 13px;
            }}
            QDateEdit:focus {{
                border: 1px solid {tc['field_focus_border']};
            }}
            QDateEdit::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 28px;
                border: none;
                border-left: 1px solid {tc['field_border']};
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }}
            QDateEdit::down-arrow {{
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid {tc['field_label']};
            }}
        """
        section_label_style = f"""
            color: {tc['section_label']};
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 1.5px;
        """
        field_label_style = f"color: {tc['field_label']}; font-size: 12px; font-weight: 600;"

        def make_field_block(label_text, widget):
            block = QtWidgets.QVBoxLayout()
            block.setSpacing(6)
            lbl = QtWidgets.QLabel(label_text)
            lbl.setStyleSheet(field_label_style)
            widget.setStyleSheet(field_style)
            block.addWidget(lbl)
            block.addWidget(widget)
            return block

        # ===== Section 1: Ticket Information =====
        ticket_section_label = QtWidgets.QLabel("TICKET INFORMATION")
        ticket_section_label.setStyleSheet(section_label_style)
        layout.addWidget(ticket_section_label)
        layout.addSpacing(12)

        self.pot_title_input = QtWidgets.QLineEdit()
        self.ticket_input = QtWidgets.QLineEdit()
        self.date_input = QtWidgets.QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDate(QtCore.QDate.currentDate())
        self.module_input = QtWidgets.QLineEdit()
        self.url_input = QtWidgets.QLineEdit()

        ticket_grid = QtWidgets.QGridLayout()
        ticket_grid.setHorizontalSpacing(20)
        ticket_grid.setVerticalSpacing(16)
        ticket_grid.addLayout(make_field_block("POT Title", self.pot_title_input), 0, 0, 1, 2)
        ticket_grid.addLayout(make_field_block("Ticket No.", self.ticket_input), 1, 0)
        ticket_grid.addLayout(make_field_block("Date", self.date_input), 1, 1)
        ticket_grid.addLayout(make_field_block("Affected Module / Screens", self.module_input), 2, 0)
        ticket_grid.addLayout(make_field_block("URL", self.url_input), 2, 1)
        layout.addLayout(ticket_grid)

        layout.addSpacing(24)
        divider1 = QtWidgets.QFrame()
        divider1.setFixedHeight(1)
        divider1.setStyleSheet(f"background: {tc['divider']}; border: none;")
        layout.addWidget(divider1)
        layout.addSpacing(20)

        # ===== Section 2: Company Information =====
        company_section_label = QtWidgets.QLabel("COMPANY INFORMATION")
        company_section_label.setStyleSheet(section_label_style)
        layout.addWidget(company_section_label)
        layout.addSpacing(12)

        self.company_input = QtWidgets.QLineEdit()
        self.company_id_input = QtWidgets.QLineEdit()
        self.company_name_input = QtWidgets.QLineEdit()
        self.issue_list_input = QtWidgets.QLineEdit()

        company_grid = QtWidgets.QGridLayout()
        company_grid.setHorizontalSpacing(20)
        company_grid.setVerticalSpacing(16)
        company_grid.addLayout(make_field_block("Company", self.company_input), 0, 0)
        company_grid.addLayout(make_field_block("Company ID", self.company_id_input), 0, 1)
        company_grid.addLayout(make_field_block("Company Name", self.company_name_input), 1, 0)
        company_grid.addLayout(make_field_block("Issue List", self.issue_list_input), 1, 1)
        layout.addLayout(company_grid)

        layout.addSpacing(28)

        # ===== Buttons =====
        btn_layout = QtWidgets.QHBoxLayout()
        btn_layout.setSpacing(12)

        cancel_btn = QtWidgets.QPushButton("Cancel")
        cancel_btn.setFixedHeight(42)
        cancel_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        cancel_btn.clicked.connect(self.show_main_menu_page)
        cancel_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: 1.5px solid {tc['outline_border']};
                border-radius: 10px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: 600;
                color: {tc['outline_text']};
            }}
            QPushButton:hover {{
                border: 1.5px solid {tc['cancel_hover_border']};
                color: {tc['outline_hover_text']};
            }}
        """)

        self.generate_btn = QtWidgets.QPushButton("Generate POT")
        self.generate_btn.setFixedHeight(42)
        self.generate_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        self.generate_btn.clicked.connect(self.generate_pot)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(self.generate_btn)
        layout.addLayout(btn_layout)

        layout.addSpacing(16)

        # ===== Status label =====
        self.status_label = QtWidgets.QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        if is_windows_dark_mode():
            self.status_label.setStyleSheet("color: #4ade80; font-weight: 600; font-size: 12.5px;")
        else:
            self.status_label.setStyleSheet("color: #16a34a; font-weight: 600; font-size: 12.5px;")
        layout.addWidget(self.status_label)

        self.main_layout.addWidget(card, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

    def generate_pot(self):
        work_item_id = self.ticket_input.text().strip()

        if not work_item_id:
            if hasattr(self, "status_label"):
                self.status_label.setText("❌ Ticket No. is required to fetch TFS comment data.")
            return

        if not hasattr(self, "current_user"):
            if hasattr(self, "status_label"):
                self.status_label.setText("❌ No logged-in user data available.")
            return

        # Disable button + show feedback while fetching
        self.generate_btn.setEnabled(False)
        self.generate_btn.setText("Fetching ticket data...")
        if hasattr(self, "status_label"):
            self.status_label.setText("")
        QApplication.processEvents()

        self.pot_ticket_worker = PotTicketDataWorker(work_item_id, self.current_user)
        self.pot_ticket_worker.progress.connect(self._on_pot_ticket_progress)
        self.pot_ticket_worker.done.connect(self._on_pot_ticket_data_fetched)
        self.pot_ticket_worker.error.connect(self._on_pot_ticket_data_error)
        self.pot_ticket_worker.start()


    def _on_pot_ticket_progress(self, msg):
        if _ui_alive(getattr(self, "status_label", None)):
            self.status_label.setText(msg)

    def _on_pot_ticket_data_error(self, msg):
        if not _ui_alive(getattr(self, "generate_btn", None)):
            logger.info("POT fetch failed after the page was closed: %s", msg)
            return
        self.generate_btn.setEnabled(True)
        self.generate_btn.setText("Generate POT")
        if _ui_alive(getattr(self, "status_label", None)):
            self.status_label.setText(msg)


    def _on_pot_ticket_data_fetched(self, raw_html):
        logger.info("POT ticket data received: %d chars of HTML", len(raw_html))
        if not _ui_alive(getattr(self, "generate_btn", None)):
            logger.info("POT data arrived after the page was closed - discarded")
            return
        self.generate_btn.setEnabled(True)
        self.generate_btn.setText("Generate POT")

        # --- Load the Word template ---
        pot_template_path = POT_TEMPLATE_PATH
        logger.info("Loading POT template: %s (cwd=%s)", pot_template_path, os.getcwd())
        if not os.path.isfile(pot_template_path):
            logger.error("POT template NOT FOUND: %s", pot_template_path)
            if hasattr(self, "status_label"):
                self.status_label.setText(f"❌ Template not found: {pot_template_path}")
            return
        try:
            doc = Document(pot_template_path)
            logger.info("POT template loaded (%d paragraph(s), %d table(s))", len(doc.paragraphs), len(doc.tables))
        except Exception:
            logger.exception("Failed to load POT template: %s", pot_template_path)
            if hasattr(self, "status_label"):
                self.status_label.setText(f"❌ Error loading template: {pot_template_path}")
            return

        # --- Get logged-in user's display name ---
        display_name = self.current_user.get("display_name", "") if hasattr(self, "current_user") else ""

        # --- Prepare text-only replacements dictionary ---
        replacements = {
            "{{POT_TITLE}}": self.pot_title_input.text(),
            "{{COMPANY}}": self.company_input.text(),
            "{{TICKET_NO}}": self.ticket_input.text(),
            "{{DATE}}": self.date_input.date().toString("MM-dd-yyyy"),
            "{{MODULE}}": self.module_input.text(),
            "{{URL}}": self.url_input.text(),
            "{{ISSUE_LIST}}": self.issue_list_input.text(),
            "{{COMPANY_ID}}": self.company_id_input.text(),
            "{{COMPANY_NAME}}": self.company_name_input.text(),
            "{{DISPLAYN}}": display_name,
        }
        logger.debug("POT replacements: %s", {k: v[:40] if isinstance(v, str) else v for k, v in replacements.items()})

        # --- Replace all plain-text placeholders everywhere (paragraphs + tables + text boxes) ---
        replace_placeholders_in_doc(doc, replacements)

        # --- Insert TICKETDATA (raw HTML, with images) into its paragraph ---
        self._insert_html_into_placeholder(doc, "{{TICKETDATA}}", raw_html)

        # --- Generate output file name safely ---
        ticket = _safe_filename(self.ticket_input.text().strip().replace(" ", "_"))
        company = _safe_filename(self.company_input.text().strip().replace(" ", "_"))
        output_filename = f"TFS{ticket}_WebPOS_{company}_POT_Internal_Passv1.0.docx"

        # Same dated folder as the status reports, under the app directory
        today_folder = os.path.join(REPORTS_DIR, datetime.now().strftime("%m-%d-%Y"))
        os.makedirs(today_folder, exist_ok=True)
        output_path = os.path.join(today_folder, output_filename)

        try:
            doc.save(output_path)
            logger.info("POT saved successfully: %s", output_path)
        except Exception:
            logger.exception("Failed to save POT document: %s", output_path)
            if hasattr(self, "status_label"):
                self.status_label.setText(f"❌ Error saving POT: {output_path}")
            return

        # Track for the Open Folder button
        self.last_output_path = output_path
        if hasattr(self, "pot_folder_btn"):
            self.pot_folder_btn.setEnabled(True)

        if hasattr(self, "status_label"):
            self.status_label.setText(f"✅ POT generated and saved to {output_path}")
        else:
            logger.info("POT generated: %s", output_path)


    def _insert_html_into_placeholder(self, doc, placeholder, raw_html):
        """
        Delegates to the shared HTML renderer, which covers body paragraphs,
        table cells (nested too) and text boxes.
        """
        pat = self.current_user.get("pat", "") if hasattr(self, "current_user") else ""
        return insert_html_into_placeholder(doc, placeholder, raw_html, pat)

    def open_report_folder(self):
        """Open the folder holding the last generated document.

        Uses the logger instead of print(): a windowed build has no stdout.
        """
        path = getattr(self, "last_output_path", "")
        if not path:
            logger.info("Open folder clicked, but nothing has been generated yet")
            return
        folder = os.path.dirname(os.path.abspath(path))
        if os.path.isdir(folder):
            os.startfile(folder)  # Windows-only app; opens Explorer
        else:
            logger.warning("Report folder does not exist: %s", folder)
    
    

    #CREATE REPORT PAGE
    def show_report_page(self):
        self.clear_layout()
        self.set_page_size(460, 700)

        card = QtWidgets.QWidget()
        card.setObjectName("Card")
        card.setFixedWidth(400)
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(12)


        tc = self._theme_colors()

        #FOLDER LAYOUT
        top_layout = QtWidgets.QHBoxLayout()

        back_btn = QtWidgets.QToolButton()
        back_btn.setText("←")
        back_btn.setToolTip("Back to Main Menu")
        back_btn.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        back_btn.clicked.connect(self.show_main_menu_page)
        back_btn.setStyleSheet(f"""
            QToolButton {{
                border: none;
                font-size: 18px;
                font-weight: bold;
                color: {tc['back_color']};
            }}
            QToolButton:hover {{
                color: {tc['back_hover']};
            }}
        """)

        title = QtWidgets.QLabel("Generate Status Report")
        title.setObjectName("Title")

        self.open_folder_btn = QtWidgets.QToolButton()
        self.open_folder_btn.setIcon(self.style().standardIcon(QtWidgets.QStyle.StandardPixmap.SP_DirOpenIcon))
        self.open_folder_btn.setToolTip("Open Report Folder")
        self.open_folder_btn.clicked.connect(self.open_report_folder)

        top_layout.addWidget(back_btn)
        top_layout.addWidget(title)
        top_layout.addStretch()
        top_layout.addWidget(self.open_folder_btn)

        layout.addLayout(top_layout)
        # Subtitle
        subtitle = QtWidgets.QLabel("Fill in the details below to create a report")
        subtitle.setObjectName("Subtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        # Work Item ID
        self.work_item_input = QtWidgets.QLineEdit()
        self.work_item_input.setPlaceholderText("Work Item ID")
        layout.addWidget(self.work_item_input)

        # Client Name
        self.client_input = QtWidgets.QLineEdit()
        self.client_input.setPlaceholderText("Client Name")
        layout.addWidget(self.client_input)

        # Location type
        loc_layout = QtWidgets.QHBoxLayout()
        loc_label = QtWidgets.QLabel("Location Type:")
        loc_label.setStyleSheet(f"color: {tc['field_label']};")
        loc_layout.addWidget(loc_label)

        self.onsite_radio = QtWidgets.QRadioButton("Onsite")
        self.offsite_radio = QtWidgets.QRadioButton("Offsite / Remote")
        self.offsite_radio.setChecked(True)
        loc_layout.addWidget(self.onsite_radio)
        loc_layout.addWidget(self.offsite_radio)
        layout.addLayout(loc_layout)

        # Generate Button
        self.generate_report_btn = QtWidgets.QPushButton("Generate Report")
        self.generate_report_btn.clicked.connect(self.generate_report)
        layout.addWidget(self.generate_report_btn)

        for field in card.findChildren(QtWidgets.QLineEdit):
            field.returnPressed.connect(self.generate_report_btn.click)

        # Output / status area
        self.output_area = QtWidgets.QTextEdit()
        self.output_area.setReadOnly(True)
        self.output_area.setStyleSheet(f"background: {tc['output_bg']}; color: {tc['output_text']};")
        layout.addWidget(self.output_area)

        self.status_label = QLabel("")  # initially empty
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        if is_windows_dark_mode():
            self.status_label.setStyleSheet("color: #4ade80; font-weight: bold;")
        else:
            self.status_label.setStyleSheet("color: #16a34a; font-weight: bold;")
        layout.addWidget(self.status_label)  # add it to the bottom

        # Create the button but keep it hidden initially
        self.create_azure_btn = QtWidgets.QPushButton("Create Azure Status Report")
        self.create_azure_btn.setVisible(False)  # hidden until DOCX is saved
        self.create_azure_btn.clicked.connect(self.on_create_azure_btn_clicked)
        layout.addWidget(self.create_azure_btn, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        self.main_layout.addWidget(card, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

    def on_create_azure_btn_clicked(self):
        # Make sure we have report data
        if not hasattr(self, "last_report_data"):
            self.output_area.append("❌ No report data available to create PBI.")
            return
        user_data = self.current_user
        pat = user_data["pat"]

        # Build a default report title
        report_title = f"Ticket No. {self.last_report_data['work_item_id']} - Status Report for {self.last_report_data['client_name']}"

        # Open the PBI input dialog
        dialog = PBIFormDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            values = dialog.get_values()  # get user input from the dialog

            # Call your PBI creation method with GUI values
            self.create_pbi_interactively_tfs(
                pat,
                report_title,
                self.last_report_data["root_cause"],
                self.last_report_data["preventive_action"],
                self.last_report_data["next_step"],
                self.last_report_data["raw_html"],
                back_job=values["back_job"],
                billing_status=values["billing_status"],
                product_type=values["product_type"],
                support_type=values["support_type"],
                confirmed_by=values["confirmed_by"]
            )
    


    def create_pbi_interactively_tfs(
            self,
            pat,
            report_title,
            root_cause,
            preventive_action,
            next_step,
            raw_html,
            back_job,
            billing_status,
            product_type,
            support_type,
            confirmed_by,
            timezone_str="Asia/Kuala_Lumpur",
        ):
        """
        Create a Product Backlog Item (PBI) in TFS (Azure DevOps Server)
        using GUI-provided values and show result in a popup.
        """
        logger.info("=== Creating PBI: %r ===", report_title)

        # --- Static config ---
        collection_url = "https://tfs.alliancewebpos.com/tfs/WebPOSCollection"
        project = "WebPOS_EES"

        user_data = self.current_user
        assigned_to = user_data["display_name"]

        # --- Get current time ---
        try:
            # zoneinfo (stdlib) replaces pytz; Windows needs the tzdata package,
            # which is now in requirements.txt.
            now = datetime.now(ZoneInfo(timezone_str))
            current_time = now.strftime("%m/%d/%Y %I:%M %p %Z").lstrip("0").replace("/0", "/")
        except Exception as e:
            now = datetime.now()
            current_time = now.strftime("%m/%d/%Y %I:%M %p")
            logger.warning("Timezone '%s' not found, using local time: %s", timezone_str, e)

        # --- REST API URL ---
        url = f"{collection_url}/{project}/_apis/wit/workitems/$Product%20Backlog%20Item?api-version=5.1"
        headers = {'Content-Type': 'application/json-patch+json'}

        # --- Build request body ---
        body = [
            {"op": "add", "path": "/fields/System.Title", "value": report_title},
            {"op": "add", "path": "/fields/System.Description", "value": raw_html},
            {"op": "add", "path": "/fields/System.AssignedTo", "value": assigned_to},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReportBillingStatus", "value": billing_status},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReportProducttype", "value": product_type},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReportSupportType", "value": support_type},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReportConfirmedBy", "value": confirmed_by},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReport_RootCause", "value": root_cause or "N/A"},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReport_PreventiveAction", "value": preventive_action or "N/A"},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReport_NextSteps", "value": next_step or "N/A"},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReportBackjob", "value": back_job},
            {"op": "add", "path": "/fields/Alliance.Groundup.Workitem.StatusReportDate", "value": current_time}
        ]
        logger.debug("PBI body fields: billing=%r, product=%r, support=%r, confirmed_by=%r, back_job=%r",
                      billing_status, product_type, support_type, confirmed_by, back_job)

        # --- Make API call ---
        try:
            logger.debug("POST %s", url)
            response = requests.post(
                url,
                auth=HTTPBasicAuth('', pat),
                headers=headers,
                json=body,
                timeout=HTTP_TIMEOUT
            )

            if response.status_code in (200, 201):
                data = response.json()

                pbi_id = data.get("id")

                # Prefer human-friendly web URL
                pbi_url = (
                    data.get("_links", {})
                        .get("html", {})
                        .get("href")
                    or data.get("url")
                )

                logger.info("PBI created successfully: id=%s, url=%s", pbi_id, pbi_url)

                dialog = PBICreatedDialog(pbi_id, pbi_url, self)
                dialog.exec()

                self.reset_generate_report_page()

            else:
                logger.error("PBI creation failed: status=%d, body=%s", response.status_code, response.text[:500])
                AppDialog(
                    "PBI Creation Failed",
                    f"Status Code: {response.status_code}\n\n{response.text[:300]}",
                    self
                ).exec()

        except Exception as e:
            logger.exception("Error creating PBI")
            AppDialog("PBI Creation Error", f"Error creating PBI: {e}", self).exec()

    
    def reset_generate_report_page(self):
        # Clear text fields
        self.work_item_input.clear()
        self.client_input.clear()

        # Reset radio buttons
        self.onsite_radio.setChecked(True)
        self.offsite_radio.setChecked(False)

        # Clear output area
        self.output_area.clear()

        # Hide Azure button again
        if hasattr(self, "create_azure_btn"):
            self.create_azure_btn.setVisible(False)
            self.create_azure_btn.setEnabled(False)

        # Clear saved report data
        if hasattr(self, "last_report_data"):
            del self.last_report_data

        # Optional status message
        self.status_label.setText("")


    def generate_report(self):
        work_item_id = self.work_item_input.text().strip()
        client_name = self.client_input.text().strip()
        location_type1 = "X" if self.onsite_radio.isChecked() else ""
        location_type2 = "X" if self.offsite_radio.isChecked() else ""

        if not work_item_id or not client_name:
            self.output_area.append("Work Item ID and Client Name cannot be empty.")
            return

        # Disable inputs to prevent duplicate submissions
        self.generate_report_btn.setEnabled(False)
        self.generate_report_btn.setText("Generating...")
        self.work_item_input.setEnabled(False)
        self.client_input.setEnabled(False)
        self.onsite_radio.setEnabled(False)
        self.offsite_radio.setEnabled(False)

        self.output_area.append("Starting report generation...")

        # Create worker
        self.worker = ReportWorker(
            work_item_id, client_name, location_type1, location_type2, self.current_user, template_path
        )

        # Connect signals
        self.worker.progress.connect(self._on_report_progress)
        self.worker.done.connect(self.handle_report_finished)
        self.worker.report_saved.connect(self.on_report_saved)

        # Start worker
        self.worker.start()

    def _on_report_progress(self, message):
        """Progress arrives from a worker, possibly after this page is gone."""
        if _ui_alive(getattr(self, "output_area", None)):
            self.output_area.append(message)
        else:
            logger.debug("Report progress after the page was closed: %s", message)

    def on_report_saved(self, path):
        self.last_output_path = path
        if _ui_alive(getattr(self, "open_folder_btn", None)):
            self.open_folder_btn.setEnabled(True)
        logger.info("Report saved at: %s", path)


 
# ================= SINGLE INSTANCE =================
# No global namespace prefix: the mutex stays inside one Windows session, so
# two RDP sessions can each keep one copy running.
APP_MUTEX_NAME = "AutoReport-Report-Marilag-SingleInstance"
MAIN_WINDOW_TITLE = "Report — Marilag"

# Kept alive for the lifetime of the process so the mutex is only released
# when the app exits.
_mutex_handle = None


def _acquire_single_instance_lock():
    """
    Hold a named Windows mutex so only one copy of the app runs at a time.
    Returns True for the first instance and False when another one already
    holds the mutex. Windows releases the mutex when the process exits, even
    if the app crashes, so a stale lock can never block a fresh start.
    """
    global _mutex_handle
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        kernel32.CreateMutexW.argtypes = [
            ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p
        ]
        handle = kernel32.CreateMutexW(None, False, APP_MUTEX_NAME)
        err = ctypes.get_last_error()

        if not handle:
            logger.warning("Could not create the single-instance mutex (error %d) — continuing", err)
            return True

        ERROR_ALREADY_EXISTS = 183
        if err == ERROR_ALREADY_EXISTS:
            kernel32.CloseHandle(ctypes.c_void_p(handle))
            logger.info("Another instance already holds mutex %s — exiting", APP_MUTEX_NAME)
            return False

        _mutex_handle = handle
        logger.info("Single-instance mutex acquired: %s", APP_MUTEX_NAME)
        return True
    except Exception:
        logger.exception("Single-instance check failed — continuing")
        return True


def _focus_running_instance():
    """
    Best-effort: raise the window of the instance that is already running, so
    the user lands on the app instead of only seeing the message box.
    """
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        user32.FindWindowW.restype = ctypes.c_void_p
        user32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
        hwnd = user32.FindWindowW(None, MAIN_WINDOW_TITLE)
        if not hwnd:
            logger.debug("Window %r not found — nothing to focus", MAIN_WINDOW_TITLE)
            return False
        SW_RESTORE = 9
        user32.ShowWindow(ctypes.c_void_p(hwnd), SW_RESTORE)
        user32.SetForegroundWindow(ctypes.c_void_p(hwnd))
        logger.info("Raised the window of the running instance")
        return True
    except Exception:
        logger.exception("Could not focus the running instance")
        return False


# detect windows theme
def _cleanup_update_artifacts():
    """
    Remove what the updater leaves behind once it has finished, without ever
    destroying work that is still pending.

      * _update_swap.bat   - it already ran before this process started
      * update_extract/    - scratch space from an interrupted extraction
      * _update_failed.txt - read and logged once, then removed

    *.new files are deliberately KEPT: they belong to a swap that has not
    completed, and deleting them is what used to make an update disappear
    silently between launches.
    """
    app_dir = _app_dir()

    failed_path = os.path.join(app_dir, UPDATE_FAILED_FILE)
    if os.path.isfile(failed_path):
        try:
            with open(failed_path, "r", encoding="utf-8", errors="replace") as f:
                detail = f.read().strip()
            logger.warning("A previous update could not replace: %s", detail or "(unnamed)")
        except OSError:
            logger.warning("A previous update left a failure marker behind")
        try:
            os.remove(failed_path)
        except OSError:
            logger.warning("Could not remove %s", UPDATE_FAILED_FILE)

    # Each download now gets its own mkdtemp() folder (autoreport_update_*);
    # collect any that an earlier run left behind.
    temp_root = tempfile.gettempdir()
    stale_dirs = []
    try:
        stale_dirs = [os.path.join(temp_root, n) for n in os.listdir(temp_root)
                      if n.startswith("autoreport_update_")]
    except OSError:
        pass

    for path in ([os.path.join(app_dir, "_update_swap.bat"),
                  os.path.join(tempfile.gettempdir(), "update_extract")] + stale_dirs):
        try:
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            elif os.path.exists(path):
                os.remove(path)
            logger.info("Cleaned up leftover update path: %s", path)
        except OSError:
            logger.warning("Could not clean up %s", path)

    try:
        pending = [n for n in os.listdir(app_dir) if n.endswith(".new")]
    except OSError:
        pending = []
    if pending:
        logger.info("Keeping %d staged update file(s) for a pending swap: %s",
                    len(pending), ", ".join(pending))



_DARK_MODE = None


def is_windows_dark_mode():
    """Cached theme lookup.

    This used to open (and never close) the registry key on every dialog and
    every stylesheet build, leaking a handle each time.
    """
    global _DARK_MODE
    if _DARK_MODE is not None:
        return _DARK_MODE

    key = None
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        )
        # 0 = dark, 1 = light
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        _DARK_MODE = (value == 0)
    except Exception:
        _DARK_MODE = True
    finally:
        if key is not None:
            key.Close()
    return _DARK_MODE

# ------------------- Splash Screen -------------------
class SplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedSize(300, 150)
        if is_windows_dark_mode():
            self.setStyleSheet("""
                background-color: #111a22;
                color: #e6eef5;
                font-family: Inter, Segoe UI, Arial;
                font-size: 14px;
            """)
        else:
            self.setStyleSheet("""
                background-color: #f5f7fa;
                color: #111827;
                font-family: Inter, Segoe UI, Arial;
                font-size: 14px;
            """)
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label = QLabel("Loading application...")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
        self.setLayout(layout)    



def _offer_update(window, update):
    """Show the update dialog when there is one and the window still exists."""
    if not update:
        logger.info("No update available")
        return
    if not _ui_alive(window):
        logger.info("Update %s available, but the window is gone", update.get("version"))
        return
    logger.info("Update available: %s", update["version"])
    window._update_info = update
    UpdateDialog(update["version"], update["body"], parent=window).exec()


if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    # Ensure relative paths resolve from the exe's directory, not the
    # current working directory (matters when launching via shortcut).
    if getattr(sys, 'frozen', False):
        os.chdir(os.path.dirname(sys.executable))

    app = QApplication(sys.argv)

    # Only one copy of the app may run at a time. Checked before any cleanup,
    # so a second launch cannot delete a pending update payload.
    if not _acquire_single_instance_lock():
        _focus_running_instance()
        AppDialog("Already Running", "Application is already running.").exec()
        sys.exit(0)

    # If the previous launch was the first one after an update and never got a
    # window on screen, the new build is broken: put the old files back and
    # relaunch into them. This must run before _cleanup_update_artifacts(),
    # which removes the swap script the rollback depends on.
    if _recover_failed_update():
        logger.info("Failed update rolled back - exiting so the restored build starts")
        sys.exit(0)

    _cleanup_update_artifacts()

    logger.info("=== Application starting (cwd=%s) ===", os.getcwd())
    logger.info("Python %s, frozen=%s", sys.version, getattr(sys, 'frozen', False))

    app.setWindowIcon(QtGui.QIcon(os.path.join(ASSET_DIR, "death.ico")))

    # Show splash screen first
    splash = SplashScreen()
    splash.show()

    # Use QTimer to simulate loading
    def load_main_window():
        logger.info("Loading main window")
        # Here you can import heavy modules if needed (e.g., pandas)
        window = LoginWindow()
        window.show()
        splash.close()
        app.main_window = window  # keep reference
        logger.info("Main window loaded and shown")

        # A window is on screen, so the build applied by the last update works.
        # This clears the trial marker that _recover_failed_update() checks.
        _mark_update_ok()

        # Check for updates after the UI is visible, on a worker. The HTTPS
        # request used to run on the GUI thread here, so a slow network stalled
        # the app right after the splash screen disappeared.
        def check_update():
            if not _ui_alive(window):
                return
            worker = UpdateCheckWorker(window)
            worker.found.connect(lambda update: _offer_update(window, update))
            worker.start()
            window._startup_update_worker = worker   # keep the QThread alive

        QTimer.singleShot(2000, check_update)

    # Start loading after 100ms
    QTimer.singleShot(100, load_main_window)

    sys.exit(app.exec())
