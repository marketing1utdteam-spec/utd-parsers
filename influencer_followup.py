#!/usr/bin/env python3
"""
influencer_followup.py — UTD influencer outreach follow-up (7-day silence nudge).

Faithful port of the n8n workflow INFL_followup (0PvM3G4jjZYFTg3R) to plain Python
for GitHub Actions. TEMPLATE-BASED (no Claude): find a creator we emailed at least
7 days ago who is still at Status "Sent", send the fixed follow-up email, and mark
the row "Follow-up Sent".

n8n mapping:
  «Daily at 9 AM»                     → schedule (one follow-up per run)
  «Get Creator Contacts»              → ec.read_rows(SHEET_ID, TAB)
  «Find Contacts Needing Follow-Up»   → pick_followup() (Status=='Sent' + Date Sent < now-7d)
  «Build Follow-Up Email»             → SUBJECT / build_body()
  «Send Follow-Up Email»              → ec.send_email (senderName "Sergey | UTD Web")
  «Mark as Follow-Up Sent»            → update Status/Date Sent by Email

Safety:
  • DRY_RUN=true (default) prints the draft + intended sheet write, sends nothing.
  • Followed-up emails are deduped via a SHA256-hashed state file (repo is PUBLIC).

Env:  GOOGLE_CREDENTIALS_JSON, GMAIL_APP_PW_SERGEY, INFL_GMAIL_USER,
      INFL_SHEET_ID, INFL_SHEET_TAB, DRY_RUN, STATE_DIR
"""

import os
import re
from datetime import datetime, timezone, timedelta

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get("INFL_SHEET_ID", "")
SHEET_TAB = os.environ.get("INFL_SHEET_TAB", "Sheet1")

INFL_GMAIL_USER = os.environ.get("INFL_GMAIL_USER", os.environ.get("UTD_MAIL_SERGEY", ""))
ACCOUNT = {"user": INFL_GMAIL_USER, "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")}

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# Days of silence before a follow-up (verbatim: 7 * 24 * 60 * 60 * 1000 ms).
FOLLOWUP_DAYS = int(os.environ.get("INFL_FOLLOWUP_DAYS", "7"))

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "influencer_followup_state.json")


# ═══════════════════════════════════════════════════════════════════
#   EMAIL TEMPLATE  (verbatim from «Build Follow-Up Email»)
# ═══════════════════════════════════════════════════════════════════

SUBJECT = "Follow-up: Shopify theme review collab — UTD Web"

BODY_HTML = "Hi,<br><br>I wanted to follow up on my previous message regarding a potential collaboration with UTD Web.<br><br>As a quick reminder: we're a Shopify theme development team with 5 themes and 25 presets on the <a href='https://themes.shopify.com/themes?page=1&q=utd'>Shopify Theme Store</a> — <a href='https://themes.shopify.com/themes/gain'>Gain</a>, <a href='https://themes.shopify.com/themes/ultra'>Ultra</a>, <a href='https://themes.shopify.com/themes/boutique'>Boutique</a>, <a href='https://themes.shopify.com/themes/allure'>Allure</a>, and <a href='https://themes.shopify.com/themes/victory'>Victory</a>. We're looking to partner with Shopify-focused creators for sponsored reviews and showcases.<br><br>We're happy to provide full theme access and compensate you for your time. If you're interested, I'd be glad to share more details or answer any questions.<br><br>Best regards,<br>Sergey<br>UTD Web<br><a href='https://utdweb.team'>utdweb.team</a>"


def build_body(channel):
    """«Build Follow-Up Email» ignores `channel` in the body (kept for parity)."""
    return BODY_HTML


# ═══════════════════════════════════════════════════════════════════
#   EMAIL VALIDATION  (verbatim from «Find Contacts Needing Follow-Up»)
# ═══════════════════════════════════════════════════════════════════

_BAD = ['denvdavydov', 'smortkin', 'utdweb.team', 'utd.agency', 'its_always_teatime',
        '.png', '.jpg', '.webp', '@sentry', 'your-company', 'you@company',
        'example@', 'placeholder', '%20', 'noreply']


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
#   SELECTION  (verbatim logic from «Find Contacts Needing Follow-Up»)
# ═══════════════════════════════════════════════════════════════════

def _parse_date_sent(value):
    """Parse a 'Date Sent' cell the way JS `new Date(d)` would accept it."""
    if not value:
        return None
    s = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[:len(fmt) + 2], fmt).replace(tzinfo=timezone.utc)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        return None


def pick_followup(rows, state):
    """Return the FIRST contact needing a follow-up, or None.

    Filter (verbatim): Status == 'Sent' AND the email is valid AND 'Date Sent'
    exists AND 'Date Sent' is older than the cutoff (now - FOLLOWUP_DAYS). n8n
    takes rows[0] (the first match), so we do too. A hashed-state guard prevents
    re-following-up the same address across runs.
    """
    cutoff = datetime.now(timezone.utc) - timedelta(days=FOLLOWUP_DAYS)
    for r in rows:
        s = str(r.get("Status", "")).strip()
        e = str(r.get("Email", "")).strip()
        d = r.get("Date Sent")
        if s != "Sent" or not is_valid(e) or not d:
            continue
        dt = _parse_date_sent(d)
        if not dt or not (dt < cutoff):
            continue
        if ec.is_processed(state, e.lower()):
            continue
        return {"email": e, "channel": r.get("Channel") or r.get("Name") or ""}
    return None


# ═══════════════════════════════════════════════════════════════════
#   ACTIONS  (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def _now_ts():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _print_draft(to, subject, body):
    print("\n" + "=" * 70)
    print("[DRAFT · follow-up]  DRY_RUN — not sent")
    print(f"  from   : {ACCOUNT['user']} (Sergey | UTD Web)")
    print(f"  to     : {to}")
    print(f"  subject: {subject}")
    print("  body (HTML):")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def run_once():
    print(f"=== UTD influencer outreach FOLLOW-UP | DRY_RUN={DRY_RUN} | "
          f"silence {FOLLOWUP_DAYS}d | {datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)

    try:
        rows = ec.read_rows(SHEET_ID, SHEET_TAB)
    except Exception as e:
        print(f"⚠️  Could not read creator sheet: {e}")
        return {"parser": "influencer_followup", "dry_run": DRY_RUN,
                "error": str(e), "sent": 0}

    print(f"Creator CRM: {len(rows)} rows read.")
    cand = pick_followup(rows, state)
    if not cand:
        print("No follow-ups needed — stop.")
        return {"parser": "influencer_followup", "dry_run": DRY_RUN,
                "found": False, "sent": 0, "rows": len(rows)}

    email = cand["email"]
    channel = cand["channel"]
    body = build_body(channel)
    now = _now_ts()

    print(f"\n· follow-up → {email} | channel={channel or '-'}")
    sent = 0
    if DRY_RUN:
        _print_draft(email, SUBJECT, body)
        print(f"[SHEET] DRY_RUN — would set Status='Follow-up Sent', "
              f"Date Sent='{now}' for {email}")
    else:
        ec.send_email(ACCOUNT, email, SUBJECT, body)
        ec.update_row_by_match(SHEET_ID, SHEET_TAB, "Email", email, {
            "Status": "Follow-up Sent",
            "Date Sent": now,
        })
        ec.mark_processed(state, email.lower())
        ec.save_state(STATE_FILE, state)
        sent = 1
        print(f"[SHEET] Status='Follow-up Sent' written for {email}")

    return {"parser": "influencer_followup", "dry_run": DRY_RUN, "found": True,
            "sent": sent, "email": email, "channel": channel, "rows": len(rows)}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
