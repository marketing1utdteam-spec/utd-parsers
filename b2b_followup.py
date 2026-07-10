#!/usr/bin/env python3
"""
b2b_followup.py — UTD "Referral program" B2B follow-up sender.

Faithful Python port of the n8n workflow B2B_followup (HhCE6aXjSRhUJX27) for
GitHub Actions. Once a day it finds contacts that were emailed ("Sent") more
than 7 days ago and never replied, sends ONE fixed follow-up email, and marks
the row "Follow-up Sent".

n8n → Python node mapping:
  Get All Contacts              → ec.open_worksheet + ec.read_rows_ws (ONE read)
  Find Contacts Needing Follow-Up → find_needing_followup()  (Status 'Sent',
                                   valid email, not competitor, Date Sent < now-7d)
  Build Follow-Up Email         → build_followup_email()  (fixed body, no AI)
  Send Follow-Up Email          → ec.send_email()
  Mark as Follow-Up Sent        → Status "Follow-up Sent" / Date Sent, match Email
  Daily at 9 AM                 → one follow-up per run (see B2B_FOLLOWUP_LIMIT)

There is NO Claude call in the n8n follow-up chain — the body is a static
template — so this module sends it verbatim (no ANTHROPIC_API_KEY needed to run,
though the shared config still reads it for parity with the sender).

Safety:
  • DRY_RUN=true (default) prints the drafted email + intended sheet writes and
    does NOT send or write.
  • Followed-up emails are SHA256-hashed into data/b2b_followup_state.json so the
    same lead is never followed up twice.

Usage:  python b2b_followup.py
Env:    GOOGLE_CREDENTIALS_JSON, GMAIL_APP_PW_SERGEY, B2B_SHEET_ID,
        B2B_SHEET_TAB, DRY_RUN, B2B_FOLLOWUP_LIMIT, STATE_DIR
        (ANTHROPIC_API_KEY read for parity but unused)
"""

import os
import re
from datetime import datetime, timezone, timedelta

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get(
    "B2B_SHEET_ID", "1ggMS5Hko2jCY5eqcPvasBy3P6hAwbw8rldr4cS3Zeo4")
SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")

# Cold-outreach mailbox — "Sergey | UTD Web" (same box as b2b_sender).
ACCOUNT = {
    "user": os.environ.get("B2B_SENDER_USER", "sergey.utd@gmail.com"),
    "password": os.environ.get("GMAIL_APP_PW_SERGEY", ""),
}
SENDER_NAME = "Sergey | UTD Web"  # n8n Gmail senderName (see assumptions)

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# n8n picks the FIRST qualifying row per daily trigger (one follow-up per run).
FOLLOWUP_LIMIT = int(os.environ.get("B2B_FOLLOWUP_LIMIT", "1"))

# Silence window before a follow-up is due (n8n: 7 days).
FOLLOWUP_AFTER_DAYS = int(os.environ.get("B2B_FOLLOWUP_AFTER_DAYS", "7"))

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "b2b_followup_state.json")


# ═══════════════════════════════════════════════════════════════════
#   LEAD SELECTION  (verbatim from «Find Contacts Needing Follow-Up»)
# ═══════════════════════════════════════════════════════════════════

# SHOPIFY THEME DEVELOPER BLOCKLIST — competitors, never pitch to them.
THEME_COMPETITOR_DOMAINS = ['shopify.com', 'archetypethemes.co', 'pixelunion.net', 'outofthesandbox.com', 'cleancanvas.co.nz', 'maestrooo.com', 'troopthemes.com', 'groupthought.com', 'eightthemes.com', 'weareunderground.com', 'corknine.com', 'switchthemes.co', 'safeasmilk.nl', 'krownthemes.com', 'milehighthemes.com', 'fluorescent.ca', 'trailblaze.media', 'trailblazethemes.com', 'invisiblethemes.com', 'presidio.build', 'pagemilldesign.com', 'brickspacelab.com', 'roartheme.com', 'boostertheme.com', 'wetheme.com', 'woolman.io', 'the4.co', 'stylehatch.co', 'staylime.com', 'redplugdesign.com', 'p-themes.com', 'superfinedigital.com', 'bsscommerce.com', 'emthemes.net', 'adornthemes.com', 'fuelthemes.net', 'webibazaar.com', 'slashthemes.in', 'designthemes.com', 'cssigniter.com', 'swissuplabs.com', 'foxecom.com', 'shinedesigninfo.com', 'digifist.com', 'thethemegoal.com', 'karmoon.design', 'envora.com', 'muupthemes.com', 'coquelicotthemes.com', 'barracuda.design', 'shopidevs.com', 'softalithemes.com', 'mpthemez.com', 'agnisoftware.com', 'harmoniks.com', 'openthinking.net', 'archer-commerce.com', 'nethypeco.com', 'saleshunter.io', 'kumi.studio', 'boostifythemes.com', 'templatemonster.com', 'themeforest.net', 'envato.com', 'themeisle.com', 'elegantthemes.com', 'utdweb.team']

# NOTE: the follow-up node ships a SHORTER BAD list than the sender — reproduced
# verbatim so behaviour matches the original workflow exactly.
BAD = ['denvdavydov', 'smortkin', 'utdweb.team', 'utd.agency', 'its_always_teatime', 'noreply@utd', '.png', '.jpg', '.webp', '@sentry', 'ingest.sentry',
       'your-company', 'john@company', 'you@company', 'example@', 'placeholder', '%20']


def is_theme_competitor(email):
    domain = (email.split("@")[1] if "@" in email else "").lower()
    return any(domain == d or domain.endswith("." + d) for d in THEME_COMPETITOR_DOMAINS)


def is_bad(e):
    el = e.lower()
    return any(b.lower() in el for b in BAD)


def is_valid(e):
    # follow-up node's isValid — note: no www-domain check (unlike the sender).
    if not e or "@" not in e or " " in e or len(e) < 7:
        return False
    if is_bad(e):
        return False
    parts = e.split("@")
    if len(parts) != 2 or "." not in parts[1]:
        return False
    if re.match(r"^[0-9a-f]{20,}$", parts[0], re.I):
        return False
    return True


def _parse_date_sent(value):
    """Parse the Date Sent cell like `new Date(d)` — return a UTC datetime or None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d",
                "%Y-%m-%dT%H:%M:%S%z"):
        try:
            dt = datetime.strptime(s[:len(fmt) + 6], fmt)
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None


def find_needing_followup(rows):
    """Reproduce «Find Contacts Needing Follow-Up»: Status 'Sent' + valid email +
    not competitor + Date Sent present + Date Sent older than the cutoff.
    Returns an ordered list of {row_number, email, company_name} (first = next up)."""
    cutoff = datetime.now(timezone.utc) - timedelta(days=FOLLOWUP_AFTER_DAYS)
    out = []
    for row_number, r in rows:
        s = str(r.get("Status", "")).strip()
        e = str(r.get("Email", "")).strip()
        d = r.get("Date Sent")
        if s != "Sent" or not is_valid(e) or is_theme_competitor(e) or not d:
            continue
        dt = _parse_date_sent(d)
        if not dt or not (dt < cutoff):
            continue
        out.append({
            "row_number": row_number,
            "email": e,
            "company_name": str(r.get("Company", "")) or "Unknown",
        })
    return out


# ═══════════════════════════════════════════════════════════════════
#   BUILD FOLLOW-UP EMAIL  (verbatim from «Build Follow-Up Email»)
# ═══════════════════════════════════════════════════════════════════

def build_followup_email(contact):
    body = (
        'Hi ' + contact["company_name"] + ' team,\n\n'
        'I am following up on my email from last week regarding a potential partnership.\n\n'
        'To briefly recap: we are UTD Web, a Shopify theme studio. Our themes are available on Shopify\'s official '
        'Theme Store (themes.shopify.com/themes?q=UTD). Through our UTD Referral program, web agencies that build '
        'client stores using our themes receive a referral fee on each purchase, with no additional overhead.\n\n'
        'If the timing is not right or this falls outside your area of work, please disregard this message entirely. '
        'However, if there is any interest, I would be glad to answer questions or arrange a short call at a time that suits you.\n\n'
        'Best regards,\nSergey\nUTD Web · utdweb.team'
    )
    return {
        "email": contact["email"],
        "company_name": contact["company_name"],
        "email_subject": "Follow-up: UTD Referral partnership enquiry",
        "email_body": body,
    }


# ═══════════════════════════════════════════════════════════════════
#   SEND + SHEET WRITE  (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def _print_draft(to, subject, body):
    print("\n" + "=" * 70)
    print("[DRAFT · b2b follow-up]  DRY_RUN — not sent")
    print(f"  from   : {SENDER_NAME} <{ACCOUNT['user']}>")
    print(f"  to     : {to}")
    print(f"  subject: {subject}")
    print("  body:")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def send_and_mark(ws, header, contact, drafted, state, stats):
    to = contact["email"]
    subject = drafted["email_subject"]
    body = drafted["email_body"]
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    if DRY_RUN:
        _print_draft(to, subject, body)
        print(f"  [SHEET] DRY_RUN — would set row {contact['row_number']} → "
              f"Status='Follow-up Sent', Date Sent='{now}' (match Email={to})")
        stats["sent"] += 1
        return

    ec.send_email(ACCOUNT, to, subject, body)
    updates = {"Status": "Follow-up Sent", "Date Sent": now}
    cell_updates = []
    for col, val in updates.items():
        if col in header:
            a1 = ec.gspread_a1(contact["row_number"], header.index(col) + 1)
            cell_updates.append({"range": a1, "values": [[val]]})
    try:
        ec.batch_update_cells(ws, cell_updates)
    except Exception as e:
        stats["errors"].append(f"sheet write failed for {to}: {e}")
    ec.mark_processed(state, to)
    stats["sent"] += 1
    print(f"  [SHEET] row {contact['row_number']} → Status='Follow-up Sent'")


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def run_once():
    print(f"=== UTD B2B follow-up | DRY_RUN={DRY_RUN} | limit={FOLLOWUP_LIMIT} | "
          f"after {FOLLOWUP_AFTER_DAYS}d | {datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)
    stats = {"considered": 0, "sent": 0, "skipped": 0, "errors": []}

    if not DRY_RUN and not ACCOUNT["password"]:
        stats["errors"].append("No GMAIL_APP_PW_SERGEY — cannot send.")
        print("⚠️  No app-password (GMAIL_APP_PW_SERGEY) — aborting live send.")
        return stats

    try:
        ws = ec.open_worksheet(SHEET_ID, SHEET_TAB)
        records = ec.read_rows_ws(ws)
        header = list(records[0].keys()) if records else ws.row_values(1)
    except Exception as e:
        stats["errors"].append(f"CRM read failed: {e}")
        print(f"⚠️  Could not read CRM sheet: {e}")
        return stats

    rows = list(enumerate(records, start=2))  # (row_number, record)
    print(f"CRM: {len(records)} rows read in 1 call.")

    due = find_needing_followup(rows)
    print(f"{len(due)} contact(s) need a follow-up.")

    sent_this_run = 0
    for contact in due:
        if sent_this_run >= max(1, FOLLOWUP_LIMIT):
            break
        stats["considered"] += 1

        if ec.is_processed(state, contact["email"]):
            print(f"· {contact['email']} already followed up (state) → skip.")
            stats["skipped"] += 1
            continue

        print(f"\n· follow-up row {contact['row_number']} → {contact['email']} "
              f"({contact['company_name']})")
        drafted = build_followup_email(contact)
        try:
            send_and_mark(ws, header, contact, drafted, state, stats)
            sent_this_run += 1
        except Exception as e:
            stats["errors"].append(f"send failed for {contact['email']}: {e}")
            print(f"  !! send error: {e}")

    if not due:
        print("No follow-ups needed → stop.")

    if not DRY_RUN:
        ec.save_state(STATE_FILE, state)

    print(f"\n=== done. {stats} ===")
    return stats


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
