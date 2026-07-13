#!/usr/bin/env python3
"""
influencer_sender.py — UTD "Shopify theme review collab" influencer outreach sender.

Faithful port of the n8n workflow INFL_send (jD4nxeD600T7DJPy) to plain Python for
GitHub Actions. This chain is TEMPLATE-BASED (no Claude): it reads the creator CRM
sheet, picks ONE uncontacted, valid creator email at random, sends the fixed
outreach email, and marks the row as "Sent".

n8n mapping:
  «Every 60 Minutes»            → schedule (one send per run)
  «Get Creator Contacts»        → ec.read_rows(SHEET_ID, TAB)
  «Pick Next Uncontacted»       → pick_next() (BAD/isValid + Status=='' filter + random pick)
  «Build Outreach Email»        → SUBJECT / build_body()
  «Send Outreach Email»         → ec.send_email (senderName "Sergey | UTD Web")
  «Mark as Sent»                → update Status/Date Sent/Thread ID by Email

Safety:
  • DRY_RUN=true (default) prints the draft + intended sheet write, sends nothing.
  • Already-attempted emails are deduped via a SHA256-hashed state file
    (repo is PUBLIC — no raw addresses committed).

Env:  GOOGLE_CREDENTIALS_JSON, GMAIL_APP_PW_SERGEY, INFL_GMAIL_USER,
      INFL_SHEET_ID, INFL_SHEET_TAB, DRY_RUN, STATE_DIR
"""

import os
import re
import random
from datetime import datetime, timezone

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get("INFL_SHEET_ID", "")
SHEET_TAB = os.environ.get("INFL_SHEET_TAB", "Sheet1")

# Outreach mailbox — "Sergey | UTD Web". app-password via the email_common
# account convention (same env var as agency_autoresponder's Sergey mailbox).
INFL_GMAIL_USER = os.environ.get("INFL_GMAIL_USER", os.environ.get("UTD_MAIL_SERGEY", ""))
ACCOUNT = {"user": INFL_GMAIL_USER, "password": os.environ.get("SENDER_APP_PW") or os.environ.get("GMAIL_APP_PW_SERGEY", "")}

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "influencer_sender_state.json")


# ═══════════════════════════════════════════════════════════════════
#   EMAIL TEMPLATE  (verbatim from «Build Outreach Email»)
# ═══════════════════════════════════════════════════════════════════

SUBJECT = "Shopify theme review collab — UTD Web"

BODY_HTML = "Hi,<br><br>I've been watching your Shopify theme reviews for a while and appreciate that you focus on real merchant use cases rather than just design comparisons.<br><br>I'm Sergey from UTD Web. We're a Shopify theme development team with 5 themes and 25 presets currently available in the <a href='https://themes.shopify.com/themes?page=1&q=utd'>Shopify Theme Store</a>. We've been building themes for over 4 years, but we've spent far more time developing products than promoting them, which is probably why we haven't appeared in many review videos yet.<br><br>One thing that makes our themes different is the amount of functionality merchants get out of the box. Features like upsells, cross-sells, promotional blocks, conversion-focused sections, and other sales tools are built directly into the theme, helping reduce the need for additional apps and monthly app subscriptions.<br><br>Our themes include:<br><ul><li><a href='https://themes.shopify.com/themes/gain'>Gain</a></li><li><a href='https://themes.shopify.com/themes/ultra'>Ultra</a></li><li><a href='https://themes.shopify.com/themes/boutique'>Boutique</a></li><li><a href='https://themes.shopify.com/themes/allure'>Allure</a></li><li><a href='https://themes.shopify.com/themes/victory'>Victory</a></li></ul>You can learn more about us here:<br><a href='https://utdweb.team'>https://utdweb.team</a><br><br>We're currently looking to partner with Shopify-focused creators for sponsored reviews and showcases, and we'd love to explore a collaboration with you.<br><br>We're happy to provide full access to our themes and compensate you for your time and work. We'd also be interested in learning about your rates, preferred video formats, and sponsorship options.<br><br>If this sounds interesting, I'd be happy to share more details.<br><br>Best regards,<br>Sergey<br>UTD Web"


VARIATION_SYS = (
    "You rewrite an outreach email template so every copy is worded differently "
    "while meaning EXACTLY the same. Rules:\n"
    "- Keep the same structure, the same paragraphs in the same order, the same "
    "HTML format with <br> breaks, the same <ul><li> theme list and ALL links "
    "EXACTLY as they are (do not touch URLs or <a> tags).\n"
    "- Keep every fact identical: 5 themes, 25 presets, 4+ years, full theme "
    "access, paid work, ask for rates/formats/sponsorship options.\n"
    "- Vary the wording naturally (roughly a third of the phrasing) so no two "
    "copies are identical: SIMPLE everyday English, short sentences, no idioms, "
    "no hype words (exclusive, exciting, game-changer, handpicked, curated, "
    "unique opportunity), never an em dash.\n"
    "- Same approximate length. Signature stays exactly: Best regards,<br>"
    "Sergey<br>UTD Web\n"
    "Output ONLY the rewritten HTML body, nothing else.")


def build_body(channel):
    """Claude paraphrases the approved template per letter (Gmail flags mass
    identical content as spam: real 'Message rejected' blocks were observed).
    Meaning/structure/links stay identical; wording varies. Fallback = the
    verbatim template when Claude is unavailable or breaks a link."""
    try:
        import email_common as _ec
        varied = _ec.call_claude(VARIATION_SYS, BODY_HTML,
                                 model="claude-sonnet-5", max_tokens=1200)
        if varied and varied.count("<a href=") >= 6 and "utdweb.team" in varied                 and "themes/gain" in varied and "themes/victory" in varied                 and "—" not in varied:
            return varied.strip()
        if varied:
            print("  [VARIATION] draft rejected by link/dash check -> template")
    except Exception as e:
        print(f"  [VARIATION] unavailable ({e}) -> template")
    return BODY_HTML


# ═══════════════════════════════════════════════════════════════════
#   EMAIL VALIDATION  (verbatim from «Pick Next Uncontacted»)
# ═══════════════════════════════════════════════════════════════════

_BAD = ['denvdavydov', 'smortkin', 'utdweb.team', 'utd.agency', 'its_always_teatime',
        '.png', '.jpg', '.jpeg', '.webp', '.gif', '.svg', '@2x', '@sentry', 'your-company',
        'you@company', 'you@yourcompany', 'example@', 'placeholder', '%20', 'noreply', 'no-reply',
        '@company.com', 'robertsmith@']
_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def is_valid(e):
    if not e or " " in e or len(e) < 7:
        return False
    el = e.lower()
    if any(b in el for b in _BAD):
        return False
    if not _RE.match(e):
        return False
    parts = e.split("@")
    if len(parts) != 2 or "." not in parts[1]:
        return False
    if len(parts[0]) < 2 or len(parts[1]) < 6 or len(parts[1].split(".")[-1]) < 2:
        return False  # junk guard (e.g. 7@g.ebe)
    if re.match(r"^[0-9a-f]{20,}$", parts[0], re.I):
        return False
    if re.search(r"\.{2,}", e) or e.startswith(".") or ".@" in e or "@." in e:
        return False
    return True


# ═══════════════════════════════════════════════════════════════════
#   SELECTION  (verbatim logic from «Pick Next Uncontacted»)
# ═══════════════════════════════════════════════════════════════════

def pick_next(rows, state):
    """Return one uncontacted candidate dict, or None.

    Filter (verbatim): Status is empty AND the email is valid AND the email is
    NOT already present in any row that has a non-empty Status (sentSheet) AND it
    has not been attempted before (the n8n `$getWorkflowStaticData.attempted`
    session guard → persisted here as hashed state). One random candidate is
    picked, matching the n8n `rows[Math.floor(Math.random()*rows.length)]`.
    """
    sent_sheet = set(
        str(r.get("Email", "")).strip().lower()
        for r in rows if str(r.get("Status", "")).strip() != ""
    )
    candidates = []
    for r in rows:
        s = str(r.get("Status", "")).strip()
        e = str(r.get("Email", "")).strip()
        el = e.lower()
        if s == "" and is_valid(e) and el not in sent_sheet and not ec.is_processed(state, el):
            candidates.append(r)
    if not candidates:
        return None
    r = random.choice(candidates)
    em = str(r["Email"]).strip()
    return {
        "email": em,
        "email_raw": r["Email"],
        "channel": r.get("Channel") or r.get("Name") or "",
        "row_number": r.get("row_number"),
    }


# ═══════════════════════════════════════════════════════════════════
#   ACTIONS  (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def _now_ts():
    # n8n: new Date().toISOString().slice(0,19).replace('T',' ')  (UTC)
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _print_draft(to, subject, body):
    print("\n" + "=" * 70)
    print("[DRAFT · outreach]  DRY_RUN — not sent")
    print(f"  from   : {ACCOUNT['user']} (Sergey | UTD Web)")
    print(f"  to     : {to}")
    print(f"  subject: {subject}")
    print("  body (HTML):")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def run_once():
    print(f"=== UTD influencer outreach SENDER | DRY_RUN={DRY_RUN} | "
          f"{datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)

    try:
        rows = ec.read_rows(SHEET_ID, SHEET_TAB)
    except Exception as e:
        print(f"⚠️  Could not read creator sheet: {e}")
        return {"parser": "influencer_sender", "dry_run": DRY_RUN,
                "error": str(e), "sent": 0}

    print(f"Creator CRM: {len(rows)} rows read.")
    cand = pick_next(rows, state)
    if not cand:
        print("No uncontacted contacts left — stop.")
        return {"parser": "influencer_sender", "dry_run": DRY_RUN,
                "found": False, "sent": 0, "rows": len(rows)}

    email = cand["email"]
    channel = cand["channel"]
    body = build_body(channel)
    now = _now_ts()

    print(f"\n· picked {email} | channel={channel or '-'}")
    sent = 0
    if DRY_RUN:
        _print_draft(email, SUBJECT, body)
        print(f"[SHEET] DRY_RUN — would set Status='Sent', Date Sent='{now}', "
              f"Thread ID=<sent msg-id> for {email}")
    else:
        msg_id = ec.send_email(ACCOUNT, email, SUBJECT, body)
        ec.update_row_by_match(SHEET_ID, SHEET_TAB, "Email", email, {
            "Status": "Sent",
            "Date Sent": now,
            "Thread ID": msg_id,
        })
        ec.mark_processed(state, email)
        ec.save_state(STATE_FILE, state)
        sent = 1
        print(f"[SHEET] Status='Sent' written for {email}")

    return {"parser": "influencer_sender", "dry_run": DRY_RUN, "found": True,
            "sent": sent, "email": email, "channel": channel, "rows": len(rows)}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
