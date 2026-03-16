# extractor.py
# Amount & company extraction logic, Outlook folder scanning, file saving.

from __future__ import annotations

import io
import logging
import os
import re
import tempfile
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger("RemitSaver.extractor")

# ── Amount regex ──────────────────────────────────────────────────────────────
# Matches: $1,234.56  $1234.56  1,234.56  1234.56
_AMOUNT_RE = re.compile(
    r"\$?\s?(\d{1,3}(?:,\d{3})*\.\d{2})"   # grouped with commas
    r"|\$?\s?(\d+\.\d{2})",                  # plain decimal
    re.MULTILINE,
)

# Words to strip from sender display name when building company name
_FILLER_WORDS = re.compile(
    r"\b(?:Inc|LLC|Corp|LLP|Ltd|Co|N\.?A\.?|Group|Holdings|"
    r"Incorporated|Corporation|Limited|Company|Services|Solutions|"
    r"Financial|Insurance|Bank|Trust|Capital|Management)\b",
    re.IGNORECASE,
)

# ── Amount helpers ────────────────────────────────────────────────────────────

def _parse_amounts(text: str) -> List[float]:
    """Extract all dollar amounts from *text* and return as floats."""
    results = []
    for m in _AMOUNT_RE.finditer(text):
        raw = m.group(1) or m.group(2)
        try:
            results.append(float(raw.replace(",", "")))
        except ValueError:
            pass
    return results


def _pick_amount(amounts: List[float], strategy: str) -> Optional[float]:
    if not amounts:
        return None
    if strategy == "last":
        return amounts[-1]
    return max(amounts)   # default: largest


# ── Company-name derivation ───────────────────────────────────────────────────

def _derive_company_name(sender_name: str, sender_email: str) -> str:
    """
    Return a clean company name string.
    Priority: display name → domain of email address.
    """
    name = (sender_name or "").strip()
    if not name or "@" in name:
        # Display name is empty or looks like an email
        # Extract domain, drop TLD
        match = re.search(r"@([\w.-]+)", sender_email or "")
        if match:
            domain_parts = match.group(1).split(".")
            # Drop the TLD (last part), capitalize each segment
            name = " ".join(p.capitalize() for p in domain_parts[:-1])
        else:
            name = "Unknown"
    else:
        # Strip filler corporate words
        name = _FILLER_WORDS.sub("", name)
    # Collapse extra whitespace and strip punctuation at edges
    name = re.sub(r"\s{2,}", " ", name).strip(" ,;.")
    return name or "Unknown"


# ── Attachment content readers ────────────────────────────────────────────────

def _amounts_from_pdf(data: bytes) -> List[float]:
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        return _parse_amounts(text)
    except Exception as exc:
        logger.debug("PDF read failed: %s", exc)
        return []


def _amounts_from_excel(data: bytes, ext: str) -> List[float]:
    amounts: List[float] = []
    try:
        if ext in ("xlsx", "xlsm", "xlam"):
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for cell in row:
                        if cell is None:
                            continue
                        if isinstance(cell, (int, float)):
                            amounts.append(float(cell))
                        else:
                            amounts.extend(_parse_amounts(str(cell)))
        elif ext in ("xls",):
            import xlrd
            wb = xlrd.open_workbook(file_contents=data)
            for sh in wb.sheets():
                for r in range(sh.nrows):
                    for c in range(sh.ncols):
                        cell = sh.cell(r, c)
                        if cell.ctype == xlrd.XL_CELL_NUMBER:
                            amounts.append(cell.value)
                        else:
                            amounts.extend(_parse_amounts(str(cell.value)))
    except Exception as exc:
        logger.debug("Excel read failed: %s", exc)
    return amounts


def _amounts_from_csv_txt(data: bytes) -> List[float]:
    try:
        text = data.decode("utf-8", errors="replace")
        return _parse_amounts(text[:50_000])
    except Exception:
        return []


def _amounts_from_attachment(attachment) -> List[float]:
    """Read an Outlook MailItem attachment and extract amounts."""
    try:
        ext = os.path.splitext(attachment.FileName)[1].lower().lstrip(".")
        # Save to a temp file, read bytes
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix="." + ext)
        tmp.close()
        attachment.SaveAsFile(tmp.name)
        try:
            with open(tmp.name, "rb") as fh:
                data = fh.read()
        finally:
            os.unlink(tmp.name)

        if ext == "pdf":
            return _amounts_from_pdf(data)
        if ext in ("xlsx", "xlsm", "xlam", "xls"):
            return _amounts_from_excel(data, ext)
        if ext in ("csv", "txt"):
            return _amounts_from_csv_txt(data)
    except Exception as exc:
        logger.debug("Attachment read error (%s): %s", getattr(attachment, "FileName", "?"), exc)
    return []


# ── Filename sanitizer ────────────────────────────────────────────────────────

_UNSAFE = re.compile(r'[\\/:*?"<>|]')

def _safe_name(s: str) -> str:
    return _UNSAFE.sub("_", s).strip()


def _unique_path(directory: str, filename: str) -> str:
    """Return *directory/filename* making it unique via (1), (2)…."""
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(directory, filename)
    n = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{base} ({n}){ext}")
        n += 1
    return candidate


# ── Amount extraction from a MailItem ────────────────────────────────────────

def extract_amount(mail_item, payer: Dict[str, Any]) -> Optional[float]:
    """
    Extract a dollar amount from *mail_item* according to payer's strategy.
    Returns float or None.
    """
    location = payer.get("amount_location", "auto")
    strategy = payer.get("amount_strategy", "largest")

    candidates: Dict[str, List[float]] = {
        "subject":    [],
        "body":       [],
        "attachment": [],
    }

    def _try_subject():
        candidates["subject"] = _parse_amounts(mail_item.Subject or "")

    def _try_body():
        body = (mail_item.Body or "")[:3000]
        candidates["body"] = _parse_amounts(body)

    def _try_attachments():
        try:
            for att in mail_item.Attachments:
                if att.Type == 1:  # olByValue
                    candidates["attachment"].extend(_amounts_from_attachment(att))
        except Exception as exc:
            logger.debug("Attachment iteration error: %s", exc)

    if location == "auto":
        _try_attachments()
        _try_subject()
        _try_body()
        all_amounts = (
            candidates["attachment"]
            + candidates["subject"]
            + candidates["body"]
        )
        return _pick_amount(all_amounts, strategy)

    loc_map = {
        "subject":    _try_subject,
        "body":       _try_body,
        "attachment": _try_attachments,
    }
    runner = loc_map.get(location)
    if runner:
        runner()
        all_amounts = candidates[location]
        return _pick_amount(all_amounts, strategy)

    return None


# ── Folder resolution ─────────────────────────────────────────────────────────

def _resolve_folder(application, payer: Dict[str, Any]):
    """
    Return the Outlook MAPIFolder for the payer's watch folder.
    Tries EntryID first, then display-path walk.
    Returns None on failure.
    """
    try:
        entry_id = payer.get("watch_folder_id", "")
        if entry_id:
            ns = application.GetNamespace("MAPI")
            folder = ns.GetFolderFromID(entry_id)
            return folder
    except Exception:
        pass

    # Fall back to path walk: "Inbox\Subfolder\…"
    try:
        path = payer.get("watch_folder_path", "Inbox")
        ns = application.GetNamespace("MAPI")
        parts = path.replace("/", "\\").split("\\")
        folder = ns.GetDefaultFolder(6)  # 6 = olFolderInbox
        # If first part matches "Inbox" skip it, otherwise start from root
        if parts and parts[0].lower() in ("inbox",):
            parts = parts[1:]
        for part in parts:
            folder = folder.Folders[part]
        return folder
    except Exception as exc:
        logger.warning("Could not resolve folder for payer '%s': %s", payer.get("name"), exc)
        return None


# ── Sender filter matching ────────────────────────────────────────────────────

def _matches_sender(mail_item, sender_filters: List[str]) -> bool:
    """Return True if any filter substring is found in the sender address or name."""
    if not sender_filters:
        return True  # no filter = accept all
    sender_addr = (getattr(mail_item, "SenderEmailAddress", "") or "").lower()
    sender_name = (getattr(mail_item, "SenderName", "") or "").lower()
    for flt in sender_filters:
        flt_lower = flt.strip().lower()
        if flt_lower and (flt_lower in sender_addr or flt_lower in sender_name):
            return True
    return False


# ── Extension filter ──────────────────────────────────────────────────────────

def _ext_allowed(filename: str, extensions: List[str]) -> bool:
    if not extensions or extensions == ["all"] or "all" in extensions:
        return True
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    return ext in [e.lower().strip(".") for e in extensions]


# ── Core extraction for one payer ────────────────────────────────────────────

class ExtractionResult:
    def __init__(self):
        self.saved   = 0
        self.skipped = 0
        self.errors  = 0
        self.details: List[str] = []   # human-readable log lines

    def log(self, msg: str):
        self.details.append(msg)
        logger.info(msg)


def run_extraction_for_payer(
    payer: Dict[str, Any],
    application,
    settings: Dict[str, Any],
    unread_only: bool = False,
    dry_run: bool = False,
    progress_cb: Optional[Callable[[str], None]] = None,
    limit: Optional[int] = None,
) -> ExtractionResult:
    """
    Scan the payer's watch folder and save attachments renamed to AMOUNT COMPANY.ext.

    :param dry_run:     If True, compute filenames but do not actually write files.
    :param progress_cb: Optional callback receiving status strings.
    :param limit:       If set, process at most this many matching emails.
    """
    result = ExtractionResult()

    def status(msg: str):
        result.log(msg)
        if progress_cb:
            progress_cb(msg)

    payer_name = payer.get("name", "Unknown")
    status(f"[{payer_name}] Starting extraction…")

    folder = _resolve_folder(application, payer)
    if folder is None:
        status(f"[{payer_name}] ERROR: Could not resolve Outlook folder.")
        result.errors += 1
        return result

    save_path = payer.get("save_path") or settings.get("default_save_root", "")
    if not save_path:
        status(f"[{payer_name}] ERROR: No save path configured.")
        result.errors += 1
        return result

    if not dry_run:
        os.makedirs(save_path, exist_ok=True)

    extensions   = payer.get("extensions", ["all"])
    filters      = payer.get("sender_filters", [])
    skip_inline  = settings.get("skip_inline_attachments", True)

    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)   # newest first
    except Exception as exc:
        status(f"[{payer_name}] ERROR: Cannot access folder items: {exc}")
        result.errors += 1
        return result

    processed = 0
    for mail in items:
        try:
            # Only MailItem (Class=43)
            if getattr(mail, "Class", None) != 43:
                continue
            if unread_only and not mail.UnRead:
                continue
            if not _matches_sender(mail, filters):
                continue

            sender_name  = getattr(mail, "SenderName", "") or ""
            sender_email = getattr(mail, "SenderEmailAddress", "") or ""
            company      = _derive_company_name(sender_name, sender_email)
            amount       = extract_amount(mail, payer)
            amount_str   = f"{amount:.2f}" if amount is not None else "0.00"
            subject      = getattr(mail, "Subject", "") or ""

            attachments = mail.Attachments
            if attachments.Count == 0:
                status(f"[{payer_name}] No attachments in: {subject!r}")
                result.skipped += 1
                continue

            for att in attachments:
                try:
                    att_name = att.FileName
                    # Skip inline attachments (ContentId set = inline)
                    if skip_inline and att.Type != 1:  # olByValue = 1
                        result.skipped += 1
                        continue
                    if not _ext_allowed(att_name, extensions):
                        result.skipped += 1
                        continue

                    _, ext = os.path.splitext(att_name)
                    new_name = _safe_name(f"{amount_str} {company}{ext}")

                    if dry_run:
                        status(f"[DRY-RUN] Would save: {new_name}  (from: {subject!r})")
                        result.saved += 1
                        continue

                    dest = _unique_path(save_path, new_name)
                    att.SaveAsFile(dest)
                    status(
                        f"[{payer_name}] Saved: {os.path.basename(dest)}"
                        f"  (source: {subject!r})"
                    )
                    _write_log_entry(settings, payer_name, dest, sender_email, subject)
                    result.saved += 1

                except Exception as exc:
                    status(f"[{payer_name}] ERROR saving attachment: {exc}")
                    result.errors += 1

            processed += 1
            if limit and processed >= limit:
                break

        except Exception as exc:
            status(f"[{payer_name}] ERROR processing mail: {exc}")
            result.errors += 1

    status(
        f"[{payer_name}] Done — saved: {result.saved}, "
        f"skipped: {result.skipped}, errors: {result.errors}"
    )
    return result


# ── Structured log writer ─────────────────────────────────────────────────────

def _write_log_entry(
    settings: Dict[str, Any],
    payer: str,
    filename: str,
    sender_email: str,
    subject: str,
) -> None:
    """Append a structured CSV-style line to remitsaver.log."""
    from datetime import datetime
    log_path = settings.get("log_file_path", "")
    if not log_path:
        return
    try:
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = (
            f"{ts} | INFO | {payer} | {os.path.basename(filename)} "
            f"| {sender_email} | {subject}\n"
        )
        with open(log_path, "a", encoding="utf-8") as fh:
            fh.write(line)
    except Exception:
        pass
