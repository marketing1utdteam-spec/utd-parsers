#!/usr/bin/env python3
"""
influencer_reminders.py — UTD influencer outreach reply-watcher / manager alert.

Faithful port of the n8n workflow INFL_reminders (x3ccDxXtZBo1iRWy) to plain
Python for GitHub Actions. TEMPLATE-BASED (no Claude): scan the outreach inbox for
replies from creators we contacted (Status "Sent" or "Follow-up Sent"), match each
reply to a contact by Thread ID (then sender email as fallback), skip
auto-replies/bounces, notify the manager, and mark the row "Replied".

n8n mapping:
  «Every 15 Minutes»           → schedule
  «Get Creator Contacts»       → ec.read_rows(SHEET_ID, TAB)
  «Build Gmail Search Query»   → build_sent_contacts()  (Status in Sent/Follow-up Sent + isValid)
  «Get Inbox Replies»          → ec.fetch_inbox(ACCOUNT, since_days=35)  (in:inbox newer_than:35d)
  «Match Replies to Contacts»  → match_replies()  (thread-id then email; skip auto/bounce; dedup)
  «Notify Manager»             → ec.send_email(ACCOUNT, MANAGER, ...)  (plain text)
  «Mark as Replied»            → update Status/Date Replied by Email

Safety:
  • DRY_RUN=true (default) prints the manager alert + intended sheet write, sends nothing.
  • Handled reply messages are deduped via a SHA256-hashed state file (repo is PUBLIC).

Env:  GOOGLE_CREDENTIALS_JSON, GMAIL_APP_PW_SERGEY, INFL_GMAIL_USER,
      INFL_MANAGER_EMAIL, INFL_SHEET_ID, INFL_SHEET_TAB,
      INFL_LOOKBACK_DAYS, DRY_RUN, STATE_DIR
"""

import os
import re
from datetime import datetime, timezone

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get(
    "INFL_SHEET_ID", "12IiHIsdibJPRGYNyZfrvdmBDY9OjmsokdmL4GgWg4qQ")
SHEET_TAB = os.environ.get("INFL_SHEET_TAB", "Sheet1")

INFL_GMAIL_USER = os.environ.get("INFL_GMAIL_USER", "sergey.utd@gmail.com")
ACCOUNT = {"user": INFL_GMAIL_USER, "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")}

# «Match Replies to Contacts»: const MANAGER = 'serhii.smortkin.utd@gmail.com'
MANAGER = os.environ.get("INFL_MANAGER_EMAIL", "serhii.smortkin.utd@gmail.com")

# Our own outreach mailbox — inbound from it is our own sent mail, not a reply.
OWN_ADDRESSES = [INFL_GMAIL_USER, "sergey.utd@gmail.com", "serge.utd@gmail.com",
                 "serhii.smortkin.utd@gmail.com"]

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# «Build Gmail Search Query»: in:inbox newer_than:35d
LOOKBACK_DAYS = int(os.environ.get("INFL_LOOKBACK_DAYS", "35"))

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "influencer_reminders_state.json")


# ═══════════════════════════════════════════════════════════════════
#   EMAIL VALIDATION  (verbatim from «Build Gmail Search Query»)
# ═══════════════════════════════════════════════════════════════════

_BAD = ['.png', '.jpg', '.webp', '@sentry', 'ingest.sentry',
        'your-company', 'john@company', 'you@company', 'your@email',
        'name@email', 'example@', 'placeholder', '%20', 'sentry.io', 'abc@company']


def is_valid(e):
    if not e or "@" not in e or " " in e or len(e) < 7:
        return False
    if any(b in e.lower() for b in _BAD):
        return False
    parts = e.split("@")
    if len(parts) != 2 or "." not in parts[1]:
        return False
    if re.match(r"^[0-9a-f]{20,}$", parts[0], re.I):
        return False
    return True


# ═══════════════════════════════════════════════════════════════════
#   SENT CONTACTS  (verbatim from «Build Gmail Search Query»)
# ═══════════════════════════════════════════════════════════════════

def build_sent_contacts(rows):
    """Creators we have contacted and are awaiting a reply from."""
    sent = []
    for r in rows:
        s = str(r.get("Status", "")).strip()
        e = str(r.get("Email", "")).strip()
        if (s == "Sent" or s == "Follow-up Sent") and is_valid(e):
            sent.append({
                "email": e.lower(),
                "company": r.get("Channel") or "Unknown",
                "thread_id": str(r.get("Thread ID", "")).strip(),
            })
    return sent


# ═══════════════════════════════════════════════════════════════════
#   MATCH REPLIES  (ported from «Match Replies to Contacts»)
# ═══════════════════════════════════════════════════════════════════

def match_replies(msgs, contacts, state):
    """Match inbound messages to contacted creators.

    Strategy 1: Thread ID — the stored Thread ID (the Message-ID we saved when we
      sent) appears in the reply's References/In-Reply-To chain.
    Strategy 2: sender email equals a contact email (fallback).
    Auto-replies, bounces and our own loop-back mail are dropped via
    ec.classify_incoming (only 'human' passes). Deduped per contact (seen) and
    per message across runs (hashed state).
    """
    seen = set()
    replies = []
    for msg in msgs:
        sender = (msg.get("from_email", "") or "").lower()

        # Strategy 1: Thread ID match.
        refs = (msg.get("references", "") + " " + msg.get("in_reply_to", "")).strip()
        match = None
        method = ""
        for c in contacts:
            tid = c.get("thread_id", "")
            if tid and (tid in refs or tid == msg.get("gm_thrid", "")):
                match = c
                method = "thread_id"
                break
        # Strategy 2: sender email fallback.
        if not match and sender:
            for c in contacts:
                if c["email"] == sender:
                    match = c
                    method = "email"
                    break

        if not match:
            continue
        if match["email"] in seen:
            continue  # deduplicate per contact

        # Skip auto-replies / bounces / own loop-back (bounce/send_failed/auto_reply).
        if ec.classify_incoming(msg, OWN_ADDRESSES) != "human":
            continue

        # Cross-run dedup: this exact reply message already handled?
        if ec.is_processed(state, msg.get("message_id", "")):
            continue

        seen.add(match["email"])
        replies.append({
            "manager_email": MANAGER,
            "contact_email": match["email"],
            "contact_name": match["company"],
            "match_method": method,
            "msg_snippet": (msg.get("body", "") or "")[:400],
            "msg_subject": msg.get("subject", "") or "No Subject",
            "msg_from": msg.get("from", ""),
            "msg_date": msg.get("date", ""),
            "msg_id": msg.get("message_id", ""),
            "msg_thread": msg.get("gm_thrid", ""),
        })
    return replies


# ═══════════════════════════════════════════════════════════════════
#   MANAGER NOTIFICATION  (verbatim template from «Notify Manager»)
# ═══════════════════════════════════════════════════════════════════

def build_notification(rep):
    subject = f"↩️ Creator replied: {rep['contact_name'] or rep['contact_email']}"
    body = (
        "Hi,\n\n"
        "A Shopify creator has replied to your review collab outreach!\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📺 Channel/Name:  {rep['contact_name']}\n"
        f"📧 Their email:   {rep['contact_email']}\n"
        f"📅 Date:          {rep['msg_date']}\n"
        f"📌 Subject:       {rep['msg_subject']}\n"
        f"🔗 Match method:  {rep['match_method']}\n\n"
        "Message preview:\n"
        f"{rep['msg_snippet']}\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Open Gmail and reply to continue the conversation.\n\n"
        "— Automated notification"
    )
    return subject, body


# ═══════════════════════════════════════════════════════════════════
#   ACTIONS  (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def _now_ts():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _print_draft(to, subject, body):
    print("\n" + "=" * 70)
    print("[DRAFT · manager alert]  DRY_RUN — not sent")
    print(f"  from   : {ACCOUNT['user']}")
    print(f"  to     : {to}")
    print(f"  subject: {subject}")
    print("  body:")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def run_once():
    print(f"=== UTD influencer outreach REPLY-WATCHER | DRY_RUN={DRY_RUN} | "
          f"lookback {LOOKBACK_DAYS}d | {datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)

    try:
        rows = ec.read_rows(SHEET_ID, SHEET_TAB)
    except Exception as e:
        print(f"⚠️  Could not read creator sheet: {e}")
        return {"parser": "influencer_reminders", "dry_run": DRY_RUN,
                "error": str(e), "notified": 0}

    contacts = build_sent_contacts(rows)
    print(f"Creator CRM: {len(rows)} rows read, {len(contacts)} awaiting-reply contacts.")
    if not contacts:
        print("No sent emails yet — stop.")
        return {"parser": "influencer_reminders", "dry_run": DRY_RUN,
                "has_sent": False, "notified": 0, "rows": len(rows)}

    if not ACCOUNT["password"]:
        print(f"⚠️  No app-password for {ACCOUNT['user']} — cannot read inbox.")
        return {"parser": "influencer_reminders", "dry_run": DRY_RUN,
                "notified": 0, "rows": len(rows), "error": "no app-password"}

    try:
        msgs = ec.fetch_inbox(ACCOUNT, since_days=LOOKBACK_DAYS, unseen_only=False)
    except Exception as e:
        print(f"⚠️  IMAP error for {ACCOUNT['user']}: {e}")
        return {"parser": "influencer_reminders", "dry_run": DRY_RUN,
                "notified": 0, "rows": len(rows), "error": str(e)}

    print(f">>> {ACCOUNT['user']}: {len(msgs)} inbox messages in last {LOOKBACK_DAYS}d")
    replies = match_replies(msgs, contacts, state)
    if not replies:
        print("No new replies — stop.")
        return {"parser": "influencer_reminders", "dry_run": DRY_RUN,
                "notified": 0, "rows": len(rows)}

    notified = 0
    for rep in replies:
        subject, body = build_notification(rep)
        print(f"\n· reply from {rep['contact_email']} | {rep['contact_name']} "
              f"| via {rep['match_method']}")
        if DRY_RUN:
            _print_draft(rep["manager_email"], subject, body)
            print(f"[SHEET] DRY_RUN — would set Status='Replied', "
                  f"Date Replied='{_now_ts()}' for {rep['contact_email']}")
        else:
            ec.send_email(ACCOUNT, rep["manager_email"], subject, body)
            ec.update_row_by_match(SHEET_ID, SHEET_TAB, "Email", rep["contact_email"], {
                "Status": "Replied",
                "Date Replied": _now_ts(),
            })
            ec.mark_processed(state, rep["msg_id"])
            print(f"[SHEET] Status='Replied' written for {rep['contact_email']}")
        notified += 1

    if not DRY_RUN:
        ec.save_state(STATE_FILE, state)

    print(f"\n=== done. replies notified: {notified} ===")
    return {"parser": "influencer_reminders", "dry_run": DRY_RUN,
            "notified": notified, "rows": len(rows), "contacts": len(contacts)}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
