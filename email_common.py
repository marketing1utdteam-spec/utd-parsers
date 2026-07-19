#!/usr/bin/env python3
"""
email_common.py — reusable email-automation library for UTD outreach bots.

Shared building blocks for autoresponders / senders that run on GitHub Actions:
  • SMTP send via smtplib (Gmail app-password) with in-thread replies + attachments
  • IMAP read via imaplib (Gmail app-password): fetch INBOX since N days, parse
  • Google Sheets via gspread + service account (creds from env GOOGLE_CREDENTIALS_JSON)
  • Claude call via the Anthropic HTTP API (key from env ANTHROPIC_API_KEY)
  • Incoming-mail classifier (bounce / send_failed / auto_reply / human)
  • PII-safe dedup state: raw Message-IDs are SHA256-hashed before being written
    to data/*.json (repo is PUBLIC — never commit real emails/addresses).

Kept dependency-light on purpose: only stdlib + requests + gspread + google-auth
(all already in requirements.txt). No new packages required.
"""

import os
import re
import ssl
import json
import time
import base64
import smtplib
import imaplib
import hashlib
import mimetypes
from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
from email.utils import parsedate_to_datetime, formatdate, make_msgid, formataddr
from datetime import datetime, timedelta, timezone

try:
    import requests
except Exception:  # pragma: no cover
    requests = None


# ═══════════════════════════════════════════════════════════════════
#   SMTP  —  send / reply in-thread / attach docx
# ═══════════════════════════════════════════════════════════════════

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587


def _looks_html(s):
    """Heuristic: does this body contain real HTML markup?"""
    return bool(s) and bool(re.search(
        r"<(?:p|br|div|table|tr|td|a|h[1-6]|ul|ol|li|strong|b|em|i|span|img|body|html)[\s>/]",
        s, re.I))


def send_email(account, to, subject, body,
               in_reply_to=None, references=None, attachment_path=None,
               reply_to=None, from_name=None, html=None):
    """Send an email via Gmail SMTP + STARTTLS.

    account         : {"user": "x@gmail.com", "password": "<app-password>"}
    to              : str or list of recipients
    in_reply_to     : Message-ID of the email we are replying to (threads it)
    references      : existing References header value (chain), optional
    attachment_path : path to a file to attach (e.g. a .docx), optional
    from_name       : display name for the From header (e.g. "Sergey | UTD Web")
    html            : explicit HTML body. If omitted but `body` looks like HTML,
                      it is sent as HTML automatically with a plain-text fallback.

    Returns the Message-ID of the sent message.
    """
    user = account["user"]
    pw = account["password"]
    if isinstance(to, (list, tuple)):
        to_list = [t for t in to if t]
    else:
        to_list = [to]

    # TEST_RECIPIENT safety valve: when set, EVERY email from EVERY module is
    # redirected to this address, with the original recipient noted in the
    # subject. Lets us do live end-to-end tests without emailing real leads.
    test_rcpt = os.environ.get("TEST_RECIPIENT", "").strip()
    if test_rcpt:
        subject = "[TEST → %s] %s" % (", ".join(to_list), subject)
        to_list = [test_rcpt]

    # Decide plain vs HTML. If caller passed html=..., use it; else auto-detect.
    if html is None and _looks_html(body):
        html = body

    msg = EmailMessage()
    msg["From"] = formataddr((from_name, user)) if from_name else user
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg_id = make_msgid()
    msg["Message-ID"] = msg_id
    if reply_to:
        msg["Reply-To"] = reply_to

    # Threading headers so the reply lands inside the same Gmail conversation.
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        refs = (references + " " + in_reply_to) if references else in_reply_to
        msg["References"] = refs.strip()
    elif references:
        msg["References"] = references

    if html:
        # Plain-text alternative first, then the HTML part.
        text_fallback = _html_to_text(html) if body is None or _looks_html(body) else body
        msg.set_content(text_fallback or _html_to_text(html))
        msg.add_alternative(html, subtype="html")
    else:
        msg.set_content(body)

    if attachment_path and os.path.exists(attachment_path):
        ctype, _ = mimetypes.guess_type(attachment_path)
        if ctype is None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        with open(attachment_path, "rb") as f:
            data = f.read()
        msg.add_attachment(data, maintype=maintype, subtype=subtype,
                           filename=os.path.basename(attachment_path))

    ctx = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as s:
        s.ehlo()
        s.starttls(context=ctx)
        s.login(user, pw)
        s.send_message(msg)
    return msg_id


# ═══════════════════════════════════════════════════════════════════
#   IMAP  —  fetch INBOX, parse messages + attachments
# ═══════════════════════════════════════════════════════════════════

IMAP_HOST = "imap.gmail.com"
IMAP_PORT = 993


def _decode_part_text(part):
    try:
        payload = part.get_content()
        if isinstance(payload, bytes):
            return payload.decode(part.get_content_charset() or "utf-8", "replace")
        return payload or ""
    except Exception:
        raw = part.get_payload(decode=True)
        if raw is None:
            return ""
        try:
            return raw.decode(part.get_content_charset() or "utf-8", "replace")
        except Exception:
            return raw.decode("utf-8", "replace")


def _html_to_text(html):
    html = re.sub(r"<style[\s\S]*?</style>", " ", html, flags=re.I)
    html = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
    html = re.sub(r"<[^>]+>", " ", html)
    return html


def _extract_body(eml):
    """Return best-effort plain-text body (prefer text/plain, fall back html)."""
    text_plain, text_html = "", ""
    if eml.is_multipart():
        for part in eml.walk():
            if part.get_content_maintype() == "multipart":
                continue
            disp = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disp:
                continue
            ctype = part.get_content_type()
            if ctype == "text/plain" and not text_plain:
                text_plain = _decode_part_text(part)
            elif ctype == "text/html" and not text_html:
                text_html = _decode_part_text(part)
    else:
        ctype = eml.get_content_type()
        if ctype == "text/html":
            text_html = _decode_part_text(eml)
        else:
            text_plain = _decode_part_text(eml)
    body = text_plain or _html_to_text(text_html)
    return re.sub(r"\s+", " ", body).strip()


def _extract_attachments(eml):
    """Return list of {'filename', 'size', 'data'(bytes)} for real attachments."""
    out = []
    if not eml.is_multipart():
        return out
    for part in eml.walk():
        disp = (part.get("Content-Disposition") or "").lower()
        filename = part.get_filename()
        if "attachment" in disp or (filename and part.get_content_maintype() != "multipart"):
            if not filename:
                continue
            try:
                data = part.get_payload(decode=True) or b""
            except Exception:
                data = b""
            out.append({"filename": filename, "size": len(data), "data": data})
    return out


def extract_email_address(s):
    if not s:
        return ""
    m = re.search(r"<([^>\s@]+@[^>\s]+)>", s)
    if m:
        return m.group(1).lower()
    m = re.search(r"[\w.+%-]+@[\w.-]+\.[a-z]{2,}", s, re.I)
    return m.group(0).lower() if m else ""


def _extract_dsn_text(eml):
    """Concatenate text from ALL text/* and message/delivery-status parts.
    Bounce DSNs keep the failed recipient in a message/delivery-status part
    (Final-Recipient:) that the normal body extractor skips, so we grab it here."""
    chunks = []
    parts = eml.walk() if eml.is_multipart() else [eml]
    for part in parts:
        ctype = part.get_content_type()
        if ctype.startswith("text/") or ctype == "message/delivery-status":
            try:
                chunks.append(_decode_part_text(part))
            except Exception:
                pass
    return "\n".join(c for c in chunks if c)


_FINAL_RCPT_RE = re.compile(
    r"(?:Final-Recipient|Original-Recipient)\s*:\s*(?:rfc822;)?\s*<?([\w.+%-]+@[\w.-]+\.[a-z]{2,})>?",
    re.I)


def _gm_thrid_from_fetch(meta_bytes):
    """Pull the Gmail thread id (X-GM-THRID) out of a FETCH response header line."""
    try:
        s = meta_bytes.decode("utf-8", "replace") if isinstance(meta_bytes, bytes) else str(meta_bytes)
    except Exception:
        return ""
    m = re.search(r"X-GM-THRID\s+(\d+)", s)
    return m.group(1) if m else ""


def fetch_inbox(account, since_days=3, mailbox="INBOX", unseen_only=False,
                limit=int(os.environ.get("INBOX_FETCH_LIMIT") or "400")):
    """Fetch messages from a Gmail mailbox over IMAP.

    Returns a list of dicts:
      {message_id, from, from_email, to, subject, date, in_reply_to,
       references, return_path, auto_submitted, x_failed_recipients,
       final_recipients:[...], gm_thrid, body, dsn_text, attachments:[...],
       attachment_names:[...], has_attachments:bool}
    Newest first. Read-only-ish: uses PEEK so it never changes \\Seen.
    """
    user = account["user"]
    pw = account["password"]
    results = []
    since = (datetime.now(timezone.utc) - timedelta(days=since_days)).strftime("%d-%b-%Y")
    crit = ["SINCE", since]
    if unseen_only:
        crit = ["UNSEEN"] + crit

    M = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    try:
        M.login(user, pw)
        M.select(mailbox)
        typ, data = M.search(None, *crit)
        if typ != "OK":
            return results
        ids = data[0].split()
        ids = ids[-limit:] if limit else ids
        for num in reversed(ids):  # newest first
            # Ask for the Gmail thread id alongside the raw message.
            typ, msg_data = M.fetch(num, "(X-GM-THRID BODY.PEEK[])")
            if typ != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1]
            gm_thrid = _gm_thrid_from_fetch(msg_data[0][0])
            eml = BytesParser(policy=policy.default).parsebytes(raw)
            frm = str(eml.get("From", ""))
            date_hdr = str(eml.get("Date", ""))
            try:
                dt = parsedate_to_datetime(date_hdr) if date_hdr else None
            except Exception:
                dt = None
            attachments = _extract_attachments(eml)
            dsn_text = _extract_dsn_text(eml)
            final_rcpts = [m.lower() for m in _FINAL_RCPT_RE.findall(dsn_text)]
            results.append({
                "message_id": str(eml.get("Message-ID", "")).strip(),
                "from": frm,
                "from_email": extract_email_address(frm),
                "to": str(eml.get("To", "")),
                "subject": str(eml.get("Subject", "")) or "(no subject)",
                "date": dt.isoformat() if dt else date_hdr,
                "in_reply_to": str(eml.get("In-Reply-To", "")).strip(),
                "references": str(eml.get("References", "")).strip(),
                "return_path": str(eml.get("Return-Path", "")).strip(),
                "auto_submitted": str(eml.get("Auto-Submitted", "")).strip().lower(),
                "x_failed_recipients": str(eml.get("X-Failed-Recipients", "")).strip(),
                "final_recipients": final_rcpts,
                "gm_thrid": gm_thrid,
                "dsn_text": dsn_text[:6000],
                "body": _extract_body(eml)[:4000],
                "attachments": attachments,
                "attachment_names": [a["filename"] for a in attachments],
                "has_attachments": len(attachments) > 0,
                "imap_uid": num.decode() if isinstance(num, bytes) else str(num),
            })
    finally:
        try:
            M.close()
        except Exception:
            pass
        try:
            M.logout()
        except Exception:
            pass
    return results


def fetch_thread(account, gm_thrid, own_addresses=(), all_mail="[Gmail]/All Mail",
                 max_msgs=40):
    """Fetch the whole Gmail conversation (both directions) for a thread id.

    Uses the Gmail IMAP extension X-GM-THRID to search "[Gmail]/All Mail", which
    contains our sent messages too, so the model sees the full history — not just
    the single inbound email. Returns a chronological list (oldest first) of:
      {date, from, from_email, subject, direction('sent'|'received'),
       snippet, attachment_names, has_attachments}
    Returns [] on any failure (caller falls back to the single message).
    """
    if not gm_thrid:
        return []
    own = {a.lower() for a in own_addresses}
    own.add((account.get("user", "") or "").lower())
    msgs = []
    M = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    try:
        M.login(account["user"], account["password"])
        # All Mail label name can vary; try a couple of common ones.
        for box in (all_mail, "[Gmail]/All Mail", "[Google Mail]/All Mail"):
            typ, _ = M.select(box, readonly=True)
            if typ == "OK":
                break
        else:
            return []
        typ, data = M.uid("SEARCH", None, "X-GM-THRID", str(gm_thrid))
        if typ != "OK" or not data or not data[0]:
            return []
        uids = data[0].split()[-max_msgs:]  # most-recent max_msgs (search returns oldest-first); recent msgs matter for a reply
        for uid in uids:
            typ, md = M.uid("FETCH", uid, "(BODY.PEEK[])")
            if typ != "OK" or not md or not md[0]:
                continue
            eml = BytesParser(policy=policy.default).parsebytes(md[0][1])
            frm = str(eml.get("From", ""))
            frm_email = extract_email_address(frm)
            date_hdr = str(eml.get("Date", ""))
            try:
                dt = parsedate_to_datetime(date_hdr) if date_hdr else None
            except Exception:
                dt = None
            atts = _extract_attachments(eml)
            body = _extract_body(eml)
            msgs.append({
                "date": dt.isoformat() if dt else date_hdr,
                "_sortkey": dt.timestamp() if dt else 0,
                "from": frm,
                "from_email": frm_email,
                "subject": str(eml.get("Subject", "")) or "(no subject)",
                "direction": "sent" if frm_email in own else "received",
                "snippet": body[:1000],
                "body": body,
                "attachment_names": [a["filename"] for a in atts],
                "has_attachments": len(atts) > 0,
            })
    except Exception:
        return []
    finally:
        try:
            M.close()
        except Exception:
            pass
        try:
            M.logout()
        except Exception:
            pass
    msgs.sort(key=lambda m: m.get("_sortkey", 0))
    for m in msgs:
        m.pop("_sortkey", None)
    return msgs


def format_thread_readable(thread, per_msg=1400):
    """Human-readable rendering of a fetch_thread() list for hand-off / report
    emails: every message on its own, clearly separated, with WHO wrote it, WHEN,
    the subject, attachments, and its text — so nothing blurs into one block."""
    if not thread:
        return "(история переписки недоступна)"
    out = []
    for i, m in enumerate(thread, 1):
        who = ("МЫ → клиенту (UTD)" if m.get("direction") == "sent"
               else f"КЛИЕНТ → нам ({m.get('from_email', '')})")
        atts = m.get("attachment_names") or []
        att = ("\nВложения: " + ", ".join(atts)) if atts else ""
        text = (m.get("body") or m.get("snippet") or "").strip()
        if len(text) > per_msg:
            text = text[:per_msg] + " […]"
        out.append(
            "─────────── письмо %d ───────────\n"
            "Когда: %s\n"
            "Кто:   %s\n"
            "Тема:  %s%s\n\n"
            "%s" % (i, str(m.get("date", ""))[:16], who,
                    m.get("subject", ""), att, text or "(пусто)")
        )
    return "\n\n".join(out)


_DSN_PHRASE_RE = re.compile(
    r"(?:following recipient failed|tried to reach|couldn't be delivered to|"
    r"delivery to the following recipient|wasn'?t delivered to|"
    r"recipient address rejected|does not exist|could not be delivered to)"
    r"[^\w]{0,40}([\w.+%-]+@[\w.-]+\.[a-z]{2,})", re.I)

_DAEMON_RE = re.compile(r"mailer-daemon|postmaster|no-?reply|donotreply|googlemail\.com$", re.I)


def extract_failed_recipient(msg, own_addresses=()):
    """Find the real recipient whose message bounced (NOT the mailer-daemon).

    Order of preference (ported/expanded from the n8n PARSE_JS failedRcpt logic):
      1. X-Failed-Recipients header
      2. Final-Recipient / Original-Recipient lines in the delivery-status part
      3. a known bounce phrase followed by an address
      4. first plausible address in the DSN/body that is neither ours nor a daemon
    Returns "" when nothing usable is found (caller then leaves the sheet alone).
    """
    own = {a.lower() for a in own_addresses}

    def _clean(addr):
        addr = (addr or "").strip().strip("<>").lower()
        if not addr or "@" not in addr:
            return ""
        if addr in own or _DAEMON_RE.search(addr):
            return ""
        return addr

    # 1) X-Failed-Recipients header (comma/space separated)
    for cand in re.split(r"[,;\s]+", msg.get("x_failed_recipients", "") or ""):
        c = _clean(cand)
        if c:
            return c
    # 2) Final-Recipient / Original-Recipient
    for cand in msg.get("final_recipients", []) or []:
        c = _clean(cand)
        if c:
            return c
    text = (msg.get("dsn_text") or "") + "\n" + (msg.get("body") or "")
    # 3) phrase-anchored address
    for m in _DSN_PHRASE_RE.finditer(text):
        c = _clean(m.group(1))
        if c:
            return c
    # 4) first plausible non-daemon, non-own address
    for m in re.finditer(r"[\w.+%-]+@[\w.-]+\.[a-z]{2,}", text, re.I):
        c = _clean(m.group(0))
        if c:
            return c
    return ""


# ═══════════════════════════════════════════════════════════════════
#   Google Sheets  —  gspread + service account
# ═══════════════════════════════════════════════════════════════════

_GC = None


def _get_gspread_client():
    """Return a cached authorised gspread client.

    Credentials come from env GOOGLE_CREDENTIALS_JSON (raw JSON string, as set
    in GitHub Secrets) or from a file at GOOGLE_CREDS_FILE. Matches the pattern
    used by b2b_harvester / run.py.
    """
    global _GC
    if _GC is not None:
        return _GC
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    raw = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
    creds_file = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")
    info = None
    if raw:
        info = json.loads(raw)
    elif os.path.exists(creds_file):
        with open(creds_file, "r", encoding="utf-8") as f:
            info = json.load(f)
    else:
        raise RuntimeError("No Google credentials (GOOGLE_CREDENTIALS_JSON / GOOGLE_CREDS_FILE)")
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    _GC = gspread.authorize(creds)
    return _GC


def _open_ws(sheet_id, tab):
    gc = _get_gspread_client()
    ss = gc.open_by_key(sheet_id)
    return ss.worksheet(tab)


def open_worksheet(sheet_id, tab):
    """Public: open a worksheet ONCE and reuse the handle for the whole run
    (read + batched write). Costs ~2 API calls (open_by_key + worksheet)."""
    return _open_ws(sheet_id, tab)


CLOSED_HEADER = ["Date", "Chain", "Contact", "Company/Store", "Account",
                 "Outcome", "Details / review"]


def append_closed(sheet_id, tab, row):
    """Append ONE completed chain to the shared 'Closed' log so every finished
    deal (whether or not a notification email was sent) is kept with a short
    review + full contact — the team works these ready people later. `row` is a
    dict keyed by CLOSED_HEADER. Native append_row auto-grows the grid."""
    if not sheet_id:
        return False
    try:
        gc = _get_gspread_client()
        ss = gc.open_by_key(sheet_id)
        try:
            ws = ss.worksheet(tab)
        except Exception:
            ws = ss.add_worksheet(title=tab, rows=2000, cols=len(CLOSED_HEADER))
            with_sheets_backoff(lambda: ws.append_row(
                CLOSED_HEADER, value_input_option="USER_ENTERED"))
        head = ws.row_values(1)
        if not head:
            with_sheets_backoff(lambda: ws.append_row(
                CLOSED_HEADER, value_input_option="USER_ENTERED"))
            head = CLOSED_HEADER
        with_sheets_backoff(lambda: ws.append_row(
            [row.get(h, "") for h in head], value_input_option="USER_ENTERED"))
        return True
    except Exception as e:
        print(f"  [CLOSED] append failed: {e}")
        return False


NOTABLE_HEADER = ["Date", "Chain", "Account", "From", "Subject", "Why notable"]


def append_notable(sheet_id, tab, row):
    """Append ONE unusual/strange email to the shared 'Notable' log so the team can
    later ask 'who wrote us X' and get the address + which mailbox it hit. Creates
    the tab + header on first use. `row` is a dict keyed by NOTABLE_HEADER names.
    Private sheet, so storing real addresses is fine. Uses native append_row
    (auto-grows the grid). Returns True on success."""
    if not sheet_id:
        return False
    try:
        gc = _get_gspread_client()
        ss = gc.open_by_key(sheet_id)
        try:
            ws = ss.worksheet(tab)
        except Exception:
            ws = ss.add_worksheet(title=tab, rows=2000, cols=len(NOTABLE_HEADER))
            with_sheets_backoff(lambda: ws.append_row(
                NOTABLE_HEADER, value_input_option="USER_ENTERED"))
        head = ws.row_values(1)
        if not head:
            with_sheets_backoff(lambda: ws.append_row(
                NOTABLE_HEADER, value_input_option="USER_ENTERED"))
            head = NOTABLE_HEADER
        with_sheets_backoff(lambda: ws.append_row(
            [row.get(h, "") for h in head], value_input_option="USER_ENTERED"))
        return True
    except Exception as e:
        print(f"  [NOTABLE] append failed: {e}")
        return False


def _is_rate_error(e):
    err = str(e).lower()
    return any(k in err for k in ("429", "quota", "rate", "resource_exhausted", "503", "500"))


def _retry_after_seconds(e, default):
    m = re.search(r"retry[- ]?after['\"]?\s*[:=]\s*['\"]?(\d+)", str(e), re.I)
    if m:
        try:
            return float(m.group(1))
        except Exception:
            pass
    return default


def with_sheets_backoff(fn, retries=5, base=10, cap=30):
    """Run a Sheets call with exponential backoff on 429/quota (10s..30s, x5)."""
    last = None
    for attempt in range(1, retries + 1):
        try:
            return fn()
        except Exception as e:
            last = e
            if attempt >= retries or not _is_rate_error(e):
                raise
            wait = _retry_after_seconds(e, min(cap, base * attempt))
            time.sleep(min(wait, cap * 2))
    if last:
        raise last


def _records_from_values(values):
    """Build list-of-dicts from raw rows, tolerating empty/duplicate headers.

    gspread's get_all_records() raises on blank or duplicate header cells; real
    UTD sheets have both. This mirrors the n8n Sheets node: the first row is the
    header, real named columns are preserved verbatim, blank headers become
    unique `_colN` keys, and duplicate names get a numeric suffix.
    """
    if not values:
        return []
    header = values[0]
    keys, seen = [], {}
    for i, h in enumerate(header):
        name = (h or "").strip()
        if not name:
            key = "_col%d" % i
        elif name in seen:
            seen[name] += 1
            key = "%s_%d" % (name, seen[name])
        else:
            seen[name] = 0
            key = name
        keys.append(key)
    rows = []
    for r in values[1:]:
        rows.append({k: (r[i] if i < len(r) else "") for i, k in enumerate(keys)})
    return rows


def read_rows_ws(ws):
    """One read call: worksheet rows as a list of dicts, with 429 backoff.

    Tolerant of blank/duplicate header cells (see _records_from_values)."""
    return with_sheets_backoff(lambda: _records_from_values(ws.get_all_values()))


def batch_update_cells(ws, cell_updates, retries=5):
    """Write MANY cells in a SINGLE Sheets API call (values.batchUpdate).

    cell_updates: list of {"range": "C5", "values": [[value]]} dicts.
    One network call for the whole run (plus retries), instead of one call per
    contact — this is the fix for the per-email 429 storm.
    """
    if not cell_updates:
        return 0
    def _do():
        return ws.batch_update(cell_updates, value_input_option="USER_ENTERED")
    with_sheets_backoff(_do, retries=retries)
    return len(cell_updates)


def read_rows(sheet_id, tab):
    """Return worksheet rows as a list of dicts (header row = keys).

    Tolerant of blank/duplicate header cells (see _records_from_values)."""
    ws = _open_ws(sheet_id, tab)
    return with_sheets_backoff(lambda: _records_from_values(ws.get_all_values()))


def read_header(sheet_id, tab):
    ws = _open_ws(sheet_id, tab)
    return ws.row_values(1)


def update_row_by_match(sheet_id, tab, match_col, match_value, updates,
                        retries=3):
    """Gap-safe update of an existing row.

    Finds the FIRST data row whose `match_col` equals `match_value`, then writes
    each column in `updates` (dict {column_name: value}) with a targeted
    worksheet.update on that single cell. We deliberately use update() (not
    append_rows) — the influencer-scraper bug was caused by append leaving gaps;
    a positional update never shifts rows.

    Returns the 1-based row number updated, or None if no match was found.
    """
    ws = _open_ws(sheet_id, tab)
    header = ws.row_values(1)
    if match_col not in header:
        return None
    match_idx = header.index(match_col) + 1  # 1-based col
    col_values = ws.col_values(match_idx)
    target_row = None
    norm = str(match_value).strip().lower()
    for i, v in enumerate(col_values[1:], start=2):  # skip header, 1-based rows
        if str(v).strip().lower() == norm:
            target_row = i
            break
    if target_row is None:
        return None

    for col_name, value in updates.items():
        if col_name not in header:
            continue
        col_idx = header.index(col_name) + 1
        a1 = gspread_a1(target_row, col_idx)
        for attempt in range(1, retries + 1):
            try:
                ws.update(a1, [[value]], value_input_option="USER_ENTERED")
                break
            except Exception as e:
                err = str(e).lower()
                if any(k in err for k in ("429", "quota", "rate")):
                    time.sleep(20 * attempt)
                else:
                    time.sleep(5 * attempt)
    return target_row


# ── Universal bounce quarantine ─────────────────────────────────────
# When ANY chain sees a delivery-failure DSN for an address, flag that
# contact 'Bounced' in EVERY CRM sheet/tab it might live in, so no chain
# ever emails or follows-up a dead/invalid address again. Config via env:
#   BOUNCE_SHEETS = "sheetId:tab, sheetId:tab, ..."  (defaults below)
def _bounce_targets():
    raw = os.environ.get("BOUNCE_SHEETS", "").strip()
    if raw:
        out = []
        for pair in raw.split(","):
            pair = pair.strip()
            if ":" in pair:
                sid, tab = pair.split(":", 1)
                out.append((sid.strip(), tab.strip()))
        return out
    b2b = os.environ.get("B2B_SHEET_ID", "")
    ecom = os.environ.get("ECOM_SHEET_ID", b2b)
    infl = os.environ.get("INFL_SHEET_ID", "")
    t = []
    if b2b: t.append((b2b, os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")))
    if ecom: t.append((ecom, os.environ.get("ECOM_SHEET_TAB", "Ecom Contacts")))
    if infl: t.append((infl, "Sheet1"))
    return t


def mark_bounced_everywhere(email, dry_run=False, when=None):
    """Set Status='Bounced' for `email` in every configured CRM sheet/tab.
    Returns count of sheets where a row was updated. Safe/no-op on failure."""
    if not email:
        return 0
    when = when or datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    hits = 0
    for sid, tab in _bounce_targets():
        try:
            updates = {"Status": "Bounced"}
            hdr = read_header(sid, tab)
            if "Date Replied" in hdr:
                updates["Date Replied"] = when
            if dry_run:
                print(f"  [BOUNCE] DRY_RUN would flag {email} Bounced in {tab}")
                hits += 1
                continue
            n = update_row_by_match(sid, tab, "Email", email, updates)
            if n:
                print(f"  [BOUNCE] {email} -> Bounced in {tab}")
                hits += 1
        except Exception as e:
            print(f"  [BOUNCE] skip {tab}: {e}")
    return hits


def gspread_a1(row, col):
    """Convert (row, col) 1-based to an A1 reference (col letters + row)."""
    letters = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


# ═══════════════════════════════════════════════════════════════════
#   Claude  —  Anthropic Messages API
# ═══════════════════════════════════════════════════════════════════

ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"
DEFAULT_MODEL = "claude-sonnet-5"


CLAUDE_CALLS = {"n": 0}  # per-run telemetry: visible in GHA logs


def call_claude(system, user, model=DEFAULT_MODEL, max_tokens=1500,
                api_key=None, timeout=90, retries=2):
    """Call the Anthropic Messages API and return the assistant text.

    Retries up to `retries` times on API error / empty response (so a transient
    5xx or timeout does not look like an unparseable answer to the caller).
    Returns "" only after all attempts fail — the caller MUST treat "" as
    "undecided, try again next run", never as an escalation.
    """
    if requests is None:
        return ""
    CLAUDE_CALLS["n"] += 1
    print(f"  [CLAUDE] call #{CLAUDE_CALLS['n']} model={model} in~{(len(system)+len(user))//4}tok")
    key = api_key or os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        return ""
    payload = {
        "model": model,
        "max_tokens": max_tokens,
        "system": system,
        "messages": [{"role": "user", "content": user}],
    }
    headers = {
        "x-api-key": key,
        "anthropic-version": ANTHROPIC_VERSION,
        "content-type": "application/json",
    }
    for attempt in range(retries + 1):
        try:
            r = requests.post(ANTHROPIC_URL, headers=headers,
                              data=json.dumps(payload), timeout=timeout)
            if r.status_code == 200:
                data = r.json()
                u = data.get("usage", {}) or {}
                CLAUDE_CALLS["in"] = CLAUDE_CALLS.get("in", 0) + int(u.get("input_tokens", 0) or 0)
                CLAUDE_CALLS["out"] = CLAUDE_CALLS.get("out", 0) + int(u.get("output_tokens", 0) or 0)
                print(f"  [CLAUDE usage] in={u.get('input_tokens', 0)} out={u.get('output_tokens', 0)} "
                      f"cache_read={u.get('cache_read_input_tokens', 0)} "
                      f"stop={data.get('stop_reason')}")
                parts = data.get("content", [])
                text = "".join(p.get("text", "") for p in parts
                               if p.get("type") == "text").strip()
                if text:
                    return text
                # 200 but no text (e.g. stop_reason 'refusal' / all-thinking).
                # Do NOT retry: an identical request returns the same empty result,
                # so retrying only triples the cost. Return "" now — the caller's
                # next-run retry + attempt cap handle any genuine transient.
                print(f"  [CLAUDE empty-200 -> no retry] stop={data.get('stop_reason')} content_types="
                      f"{[p.get('type') for p in parts]}")
                return ""
            if attempt < retries:
                if r.status_code in (429, 529):
                    # Rate limited / overloaded: honour Retry-After, else long backoff.
                    try:
                        wait = float(r.headers.get("retry-after", ""))
                    except Exception:
                        wait = 0
                    time.sleep(max(wait, 10 * (attempt + 1)))
                else:
                    time.sleep(3 * (attempt + 1))
        except Exception:
            if attempt < retries:
                time.sleep(5 * (attempt + 1))
    return ""


# ═══════════════════════════════════════════════════════════════════
#   Incoming-mail classifier  (ported from n8n PARSE_JS)
# ═══════════════════════════════════════════════════════════════════

_BOUNCE_FROM = re.compile(r"mailer-daemon|postmaster|mail delivery", re.I)
_BOUNCE_SUBJ = re.compile(
    r"^(undeliver|delivery status notification|mail delivery failed|failure notice|"
    r"returned mail|delivery incomplete|address not found|message not delivered)", re.I)
_AUTO_SUBJ = re.compile(
    r"^(automatic reply|auto[- ]?reply|autoreply|out of (the )?office|автоответ|"
    r"abwesenheit|réponse automatique|respuesta automática|risposta automatica)", re.I)
_AUTO_SENDER = re.compile(r"no-?reply@|donotreply@|notifications?@", re.I)
_SEND_FAILED_BODY = re.compile(
    r"максимально допустимое число писем|sending quota|quota exceeded|"
    r"reached a limit for sending|Сообщение заблокировано|message rejected|"
    r"отправка временно заблокирована", re.I)


def classify_incoming(msg, own_addresses):

    # Gmail OUTBOUND policy block ("Message rejected", answer/69585): permanent.
    # Re-sending identical content makes reputation worse -> treat as bounce
    # so the reanimator never requeues it.
    _blob = (msg.get("subject", "") or "") + " " + (msg.get("body", "") or "")
    if "mailer-daemon" in (msg.get("from_email", "") or "") and (
            "answer/69585" in _blob or "Message rejected" in _blob or
            "Сообщение заблокировано" in _blob):
        return "bounce"
    """Classify a parsed message dict (from fetch_inbox).

    Returns one of:
      'bounce'      — dead-address DSN (never resurrect the contact)
      'send_failed' — our send blocked by quota/limit (return contact to queue)
      'auto_reply'  — ticket/OOO/no-reply/own-address/unparseable sender
      'human'       — a real human reply that needs AI handling

    own_addresses : iterable of our own mailbox addresses (lowercased).
    """
    frm = msg.get("from", "") or ""
    subject = msg.get("subject", "") or ""
    sender = (msg.get("from_email", "") or "").lower()
    return_path = (msg.get("return_path", "") or "").strip()
    auto_sub = (msg.get("auto_submitted", "") or "").lower()
    body = msg.get("body", "") or ""
    own = {a.lower() for a in own_addresses}

    cat = "human"
    if (_BOUNCE_FROM.search(frm) or _BOUNCE_SUBJ.search(subject)
            or return_path == "<>"):
        cat = "bounce"
    elif sender in own:
        cat = "auto_reply"  # our own address looping back — not a real prospect
    elif auto_sub and auto_sub != "no":
        cat = "auto_reply"
    elif _AUTO_SUBJ.search(subject):
        cat = "auto_reply"
    elif _AUTO_SENDER.search(sender):
        cat = "auto_reply"
    elif not sender:
        cat = "auto_reply"

    # A "bounce" that is really a quota/block on OUR send = send_failed.
    if cat == "bounce" and _SEND_FAILED_BODY.search(body):
        cat = "send_failed"
    return cat


# ═══════════════════════════════════════════════════════════════════
#   PII-safe dedup state
# ═══════════════════════════════════════════════════════════════════

def hash_id(s):
    """SHA256 of a normalised string. Used to persist processed Message-IDs as
    one-way hashes in the PUBLIC repo (never store raw Message-IDs/emails)."""
    return hashlib.sha256(str(s).lower().strip().encode()).hexdigest()


def load_state(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"processed_ids": [], "updated": ""}


def save_state(path, state):
    state["updated"] = datetime.now(timezone.utc).isoformat()
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=0)


def is_processed(state, message_id):
    if not message_id:
        return False
    return hash_id(message_id) in set(state.get("processed_ids", []))


def mark_processed(state, message_id):
    if not message_id:
        return
    h = hash_id(message_id)
    ids = state.setdefault("processed_ids", [])
    if h not in ids:
        ids.append(h)
    # keep the file bounded (last ~5000 ids is plenty for a 3-day window)
    if len(ids) > 5000:
        state["processed_ids"] = ids[-5000:]


def bump_attempt(state, message_id, limit=3):
    """Count how many runs have failed to get a usable Claude answer for this
    message (empty response / unparseable). Returns True once the count reaches
    `limit` — the caller should then GIVE UP (mark it processed) instead of
    re-calling Claude on it every single cycle forever.

    This stops the 'empty AI response -> retry next run' branch from turning a
    handful of bad messages into hundreds of paid Claude calls per day.
    """
    if not message_id:
        return False
    h = hash_id(message_id)
    att = state.setdefault("attempts", {})
    att[h] = int(att.get(h, 0)) + 1
    if len(att) > 5000:  # bound the dict
        for k in list(att)[:len(att) - 5000]:
            att.pop(k, None)
    return att[h] >= limit
