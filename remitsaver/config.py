# config.py
# Load/save payers.json and settings.json from %APPDATA%\RemitSaver\

import os
import json
import logging
from typing import Any, Dict, List

logger = logging.getLogger("RemitSaver.config")

# ── Storage directory ──────────────────────────────────────────────────────────
def _appdata_dir() -> str:
    base = os.environ.get("APPDATA", os.path.expanduser("~"))
    path = os.path.join(base, "RemitSaver")
    os.makedirs(path, exist_ok=True)
    return path


APPDATA_DIR   = _appdata_dir()
PAYERS_FILE   = os.path.join(APPDATA_DIR, "payers.json")
SETTINGS_FILE = os.path.join(APPDATA_DIR, "settings.json")

# ── Default structures ────────────────────────────────────────────────────────

DEFAULT_SETTINGS: Dict[str, Any] = {
    "default_save_root":        os.path.join(os.path.expanduser("~"), "RemitSaver"),
    "log_file_path":            os.path.join(APPDATA_DIR, "remitsaver.log"),
    "skip_inline_attachments":  True,
    "auto_run_on_startup":      False,
    "scan_unread_only":         False,
}

# A payer dict looks like:
# {
#   "name":              "Nationwide",
#   "sender_filters":    ["@nationwide.com"],
#   "watch_folder_path": "Inbox\\Nationwide",   # EntryID or display path string
#   "watch_folder_id":   "",                     # Outlook EntryID (preferred)
#   "save_path":         "C:\\Remittances\\Nationwide",
#   "extensions":        ["pdf", "xlsx"],        # or ["all"]
#   "amount_location":   "auto",                 # auto | subject | body | attachment
#   "amount_strategy":   "largest",              # largest | last
# }

PAYER_DEFAULTS: Dict[str, Any] = {
    "name":              "",
    "sender_filters":    [],
    "watch_folder_path": "Inbox",
    "watch_folder_id":   "",
    "save_path":         "",
    "extensions":        ["all"],
    "amount_location":   "auto",
    "amount_strategy":   "largest",
}


# ── Payer helpers ─────────────────────────────────────────────────────────────

def load_payers() -> List[Dict[str, Any]]:
    """Return list of payer dicts from payers.json (creates empty file if absent)."""
    if not os.path.exists(PAYERS_FILE):
        save_payers([])
        return []
    try:
        with open(PAYERS_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, list):
            logger.warning("payers.json is not a list – resetting.")
            return []
        # Back-fill any missing keys with defaults
        result = []
        for p in data:
            merged = dict(PAYER_DEFAULTS)
            merged.update(p)
            result.append(merged)
        return result
    except Exception:
        logger.exception("Failed to load payers.json")
        return []


def save_payers(payers: List[Dict[str, Any]]) -> None:
    """Persist payer list to payers.json."""
    try:
        with open(PAYERS_FILE, "w", encoding="utf-8") as fh:
            json.dump(payers, fh, indent=2, ensure_ascii=False)
    except Exception:
        logger.exception("Failed to save payers.json")


def get_payer_by_name(name: str) -> Dict[str, Any] | None:
    for p in load_payers():
        if p["name"] == name:
            return p
    return None


def upsert_payer(payer: Dict[str, Any]) -> None:
    """Add or replace payer (matched by name)."""
    payers = load_payers()
    for i, p in enumerate(payers):
        if p["name"] == payer["name"]:
            payers[i] = payer
            save_payers(payers)
            return
    payers.append(payer)
    save_payers(payers)


def delete_payer(name: str) -> None:
    payers = [p for p in load_payers() if p["name"] != name]
    save_payers(payers)


# ── Settings helpers ──────────────────────────────────────────────────────────

def load_settings() -> Dict[str, Any]:
    """Return settings dict (creates file with defaults if absent)."""
    if not os.path.exists(SETTINGS_FILE):
        save_settings(DEFAULT_SETTINGS)
        return dict(DEFAULT_SETTINGS)
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        merged = dict(DEFAULT_SETTINGS)
        merged.update(data)
        return merged
    except Exception:
        logger.exception("Failed to load settings.json")
        return dict(DEFAULT_SETTINGS)


def save_settings(settings: Dict[str, Any]) -> None:
    """Persist settings dict to settings.json."""
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
            json.dump(settings, fh, indent=2, ensure_ascii=False)
    except Exception:
        logger.exception("Failed to save settings.json")


def update_setting(key: str, value: Any) -> None:
    s = load_settings()
    s[key] = value
    save_settings(s)
