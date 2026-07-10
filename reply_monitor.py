#!/usr/bin/env python3
"""
reply_monitor.py — UTD outreach "Reply Monitor": alert the manager on prospect replies.

Port of the n8n workflow REPLY_MON (54zBYG8SibTjMqUT) to plain Python for
GitHub Actions.

What it does (mirrors the n8n nodes):
  • «Get All Contacts»  — read the B2B CRM sheet.
  • «Build Gmail Search Query» — keep contacts whose Status is 'Sent' or
    'Follow-up Sent' AND whose Email passes the isValid() filter; poll the inbox
    broadly (newer_than:35d) so we catch replies even from a different address.
  • «Match Replies to Contacts» — match each inbound message to a sent contact,
    STRATEGY 1 by Thread ID, STRATEGY 2 by sender email; skip auto-replies /
    bounces / system mail; deduplicate one reply per contact.
  • «Notify Manager» — email the manager (MANAGER_EMAIL) the VERBATIM alert.
  • «Mark as Replied in Sheet» — set Status='Replied' + Date Replied on the row.

Reuses email_common state (SHA256-hashed Message-IDs) so the SAME reply is never
reported twice across runs — on top of the per-run one-reply-per-contact dedup.

Safety:
  • DRY_RUN=true (default) prints the alert + sheet update, sends/writes nothing.

Usage:  python reply_monitor.py
Env:    GMAIL_APP_PW_SERGEY, GMAIL_APP_PW_SERGE, GOOGLE_CREDENTIALS_JSON,
        DRY_RUN, B2B_SHEET_ID, B2B_SHEET_TAB, MANAGER_EMAIL, LOOKBACK_DAYS,
        STATE_DIR
"""

import os
import re
from datetime import datetime, timezone

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get("B2B_SHEET_ID", "")
SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")

# «Match Replies to Contacts»: const MANAGER = <manager email>
MANAGER_EMAIL = os.environ.get("MANAGER_EMAIL", "")

# «Build Gmail Search Query»: in:inbox newer_than:35d
LOOKBACK_DAYS = int(os.environ.get("LOOKBACK_DAYS", "35"))

# Our own outreach mailboxes — inbound to poll AND own-address loop guard.
OWN_ADDRESSES = [a for a in (
    os.environ.get("UTD_MAIL_SERGEY", ""),
    os.environ.get("UTD_MAIL_SERGE", ""),
    os.environ.get("UTD_MAIL_SERHII", ""),
) if a]

# The physical mailboxes we poll for replies (both B2B outreach boxes).
ACCOUNTS = [a for a in (
    {"user": os.environ.get("UTD_MAIL_SERGEY", ""), "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")},
    {"user": os.environ.get("UTD_MAIL_SERGE", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGE", "")},
) if a["user"]]

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "reply_monitor_state.json")

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# «Build Gmail Search Query»: BAD list + isValid(), VERBATIM.
_BAD = ['.png', '.jpg', '.webp', '@sentry', 'ingest.sentry',
        'your-company', 'john@company', 'you@company', 'your@email',
        'name@email', 'example@', 'placeholder', '%20', 'sentry.io', 'abc@company']


def is_valid_email(e):
    """Port of isValid() from «Build Gmail Search Query» (verbatim logic)."""
    if not e or "@" not in e or " " in e or len(e) < 7:
        return False
    el = e.lower()
    if any(b in el for b in _BAD):
        return False
    parts = e.split("@")
    if len(parts) != 2 or "." not in parts[1]:
        return False
    if re.match(r"^[0-9a-f]{20,}$", parts[0], re.I):
        return False
    return True


# ═══════════════════════════════════════════════════════════════════
#   Notification (VERBATIM from «Notify Manager»)
# ═══════════════════════════════════════════════════════════════════

def build_notification(reply):
    """Return (subject, body) for the manager alert — text VERBATIM from n8n."""
    subject = f"↩️ Reply from {reply['contact_company']} — Action Required"
    body = (
        "Hi,\n\n"
        "A potential partner has replied to your outreach and is waiting for a response!\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        f"🏢 Company:     {reply['contact_name']}\n"
        f"📧 Their email: {reply['contact_email']}\n"
        f"📅 Date:        {reply['msg_date']}\n"
        f"📌 Subject:     {reply['msg_subject']}\n\n"
        "Message preview:\n"
        f"{reply['msg_snippet']}\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "Open Gmail and reply to continue the conversation.\n\n"
        "— Automated notification"
    )
    return subject, body


# ═══════════════════════════════════════════════════════════════════
#   Contact / reply matching
# ═══════════════════════════════════════════════════════════════════

def build_sent_contacts(rows):
    """«Build Gmail Search Query»: keep Status in {Sent, Follow-up Sent} + valid Email."""
    sent = []
    for r in rows:
        status = str(r.get("Status", "")).strip()
        email = str(r.get("Email", "")).strip()
        if status in ("Sent", "Follow-up Sent") and is_valid_email(email):
            sent.append({
                "email": email.lower(),
                "company": str(r.get("Company", "")).strip() or "Unknown",
                "thread_id": str(r.get("Thread ID", "")).strip(),
                "last_msg": str(r.get("Last Msg ID", "")).strip(),
            })
    return sent


def match_reply(msg, contacts):
    """STRATEGY 1: match by Thread ID token (In-Reply-To / References contain our
    stored Thread ID / Last Msg ID). STRATEGY 2: fall back to sender email."""
    refs = (msg.get("references", "") + " " + msg.get("in_reply_to", "")).strip()
    if refs:
        for c in contacts:
            for tok in (c.get("thread_id"), c.get("last_msg")):
                if tok and tok in refs:
                    return c, "thread_id"
    sender = (msg.get("from_email", "") or "").lower()
    if sender:
        for c in contacts:
            if c["email"] == sender:
                return c, "email"
    return None, ""


# ═══════════════════════════════════════════════════════════════════
#   Actions (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def notify_manager(account, reply):
    subject, body = build_notification(reply)
    if DRY_RUN:
        print("\n" + "=" * 70)
        print("[NOTIFY]  DRY_RUN — not sent")
        print(f"  from   : {account['user']}")
        print(f"  to     : {reply['manager_email']}")
        print(f"  subject: {subject}")
        print("  body:")
        for line in body.splitlines():
            print("    " + line)
        print("=" * 70)
        return
    ec.send_email(account, reply["manager_email"], subject, body)


def mark_replied(contact_email):
    """«Mark as Replied in Sheet»: Status='Replied' + Date Replied, match by Email."""
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    if DRY_RUN:
        print(f"  [SHEET] DRY_RUN — would set Status='Replied', "
              f"Date Replied='{now}' for {contact_email}")
        return
    row = ec.update_row_by_match(SHEET_ID, SHEET_TAB, "Email", contact_email,
                                 {"Status": "Replied", "Date Replied": now})
    if row:
        print(f"  [SHEET] marked Replied (row {row}) for {contact_email}")
    else:
        print(f"  [SHEET] {contact_email} not found → not marked")


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def run_once():
    print(f"=== UTD Reply Monitor | DRY_RUN={DRY_RUN} | lookback {LOOKBACK_DAYS}d | "
          f"{datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)

    # «Get All Contacts» + «Build Gmail Search Query»
    try:
        rows = ec.read_rows(SHEET_ID, SHEET_TAB)
    except Exception as e:
        print(f"⚠️  Could not read CRM sheet: {e}")
        return {"parser": "reply_monitor", "dry_run": DRY_RUN, "error": str(e),
                "notified": 0}

    contacts = build_sent_contacts(rows)
    print(f"CRM: {len(rows)} rows read, {len(contacts)} sent-contacts to watch.")
    if not contacts:
        print("No 'Sent'/'Follow-up Sent' contacts — nothing to monitor.")
        return {"parser": "reply_monitor", "dry_run": DRY_RUN, "notified": 0,
                "sent_contacts": 0}

    default_account = next((a for a in ACCOUNTS if a["password"]), None)
    seen = set()   # «Match Replies to Contacts»: one reply per contact per run
    notified = 0
    scanned = 0

    for account in ACCOUNTS:
        if not account["password"]:
            print(f"⚠️  No app-password for {account['user']} — skipping mailbox.")
            continue
        try:
            msgs = ec.fetch_inbox(account, since_days=LOOKBACK_DAYS, unseen_only=False)
        except Exception as e:
            print(f"⚠️  IMAP error for {account['user']}: {e}")
            continue
        scanned += len(msgs)
        print(f"\n>>> {account['user']}: {len(msgs)} inbox messages in last {LOOKBACK_DAYS}d")

        for msg in msgs:
            mid = msg.get("message_id", "")

            # Cross-run dedup: this exact reply already reported before.
            if ec.is_processed(state, mid):
                continue

            contact, method = match_reply(msg, contacts)
            if not contact:
                continue
            if contact["email"] in seen:   # one reply per contact per run
                continue

            # Skip auto-replies / bounces / system mail (bounce, OOO, no-reply…).
            if ec.classify_incoming(msg, OWN_ADDRESSES) != "human":
                continue

            seen.add(contact["email"])

            reply = {
                "manager_email": MANAGER_EMAIL,
                "contact_email": contact["email"],
                "contact_name": contact["company"],
                "contact_company": contact["company"],
                "match_method": method,
                "msg_snippet": (msg.get("body", "") or "")[:400],
                "msg_subject": msg.get("subject", "") or "No Subject",
                "msg_from": msg.get("from", ""),
                "msg_date": msg.get("date", ""),
                "msg_id": mid,
                "msg_thread": msg.get("gm_thrid", ""),
            }

            print(f"\n· reply from {reply['contact_email']} "
                  f"({reply['contact_name']}) matched by {method}")
            notify_manager(account, reply)
            mark_replied(contact["email"])
            notified += 1

            if not DRY_RUN:
                ec.mark_processed(state, mid)

    if not DRY_RUN:
        ec.save_state(STATE_FILE, state)

    print(f"\n=== done. scanned {scanned} messages, notified {notified} reply(ies). ===")
    return {"parser": "reply_monitor", "dry_run": DRY_RUN,
            "sent_contacts": len(contacts), "scanned": scanned,
            "notified": notified, "manager": MANAGER_EMAIL}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
