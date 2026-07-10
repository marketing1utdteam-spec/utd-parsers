#!/usr/bin/env python3
"""
b2b_followup.py — UTD "Referral program" B2B follow-up sender (Claude-first).

Evolution of the n8n workflow B2B_followup (HhCE6aXjSRhUJX27) for GitHub
Actions. Once a day it finds contacts that were emailed ("Sent") more than
7 days ago and never replied, asks Claude to write a calm personalized
follow-up (grounded in a fresh site scrape + the actual prior correspondence
pulled over IMAP when available), sends it IN-THREAD when the row has a
Thread ID, and marks the row "Follow-up Sent".

Pipeline per contact:
  Get All Contacts                → ec.open_worksheet + ec.read_rows_ws (ONE read)
  Find Contacts Needing Follow-Up → find_needing_followup()  (Status 'Sent',
                                    valid email, not competitor, Date Sent < now-7d)
  Assemble Context                → fresh site scrape (b2b_sender helper) +
                                    prior thread via ec.fetch_thread (IMAP,
                                    Message-ID → X-GM-THRID resolve) or a
                                    faithful summary when IMAP is unavailable
  Claude — Write Follow-Up        → ec.call_claude (claude-sonnet-5, 1000 tok)
                                    canon: plain human voice, explicit reminder
                                    of the first email, ONE new angle, no em
                                    dashes, no hype, simple reply-inviting ask
  Fallback                        → static template body when Claude fails
  Send Follow-Up Email            → ec.send_email (in-thread reply when the
                                    row has a Thread ID, else "Re: <subject>")
  Mark as Follow-Up Sent          → Status "Follow-up Sent" / Date Sent

Safety:
  • DRY_RUN=true (default) prints the FULL built Claude system+user prompt,
    the drafted email and intended sheet writes, and does NOT send or write.
  • Followed-up emails are SHA256-hashed into data/b2b_followup_state.json so
    the same lead is never followed up twice.

Usage:  python b2b_followup.py
Env:    GOOGLE_CREDENTIALS_JSON, ANTHROPIC_API_KEY, GMAIL_APP_PW_SERGEY,
        B2B_SHEET_ID, B2B_SHEET_TAB, DRY_RUN, B2B_FOLLOWUP_LIMIT, STATE_DIR
"""

import os
import re
import imaplib
from datetime import datetime, timezone, timedelta

import email_common as ec
from b2b_sender import (fetch_company_website, _clean_site_text,
                        print_prompt_for_review, strip_em_dashes,
                        strip_trailing_signoff, clean_company_name,
                        ensure_greeting)


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get("B2B_SHEET_ID", "")
SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")

# Cold-outreach mailbox — "Sergey | UTD Web" (same box as b2b_sender).
ACCOUNT = {
    "user": os.environ.get("B2B_SENDER_USER", os.environ.get("UTD_MAIL_SERGEY", "")),
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

MODEL = "claude-sonnet-5"
MAX_TOKENS = 1000

SIG = '\n\nBest regards,\nSergey\nUTD Web | utdweb.team'


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
    Returns an ordered list of contact dicts (first = next up), now carrying the
    extra context columns (Website, Thread ID, Date Sent) for the Claude draft."""
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
        website = str(r.get("Website", "")).strip()
        url = ""
        if website:
            url = website if website.startswith("http") else "https://" + website
        # Sheet header is "Company Name"; "Company" kept as a fallback.
        company = clean_company_name(
            str(r.get("Company Name", "") or r.get("Company", "")))
        out.append({
            "row_number": row_number,
            "email": e,
            "company_name": company,
            "website": url,
            "thread_id": str(r.get("Thread ID", "")).strip(),
            "date_sent": str(d).strip(),
            "orig_subject": str(r.get("Subject", "") or r.get("Email Subject", "")).strip(),
        })
    return out


# ═══════════════════════════════════════════════════════════════════
#   PRIOR CORRESPONDENCE  (IMAP thread via email_common.fetch_thread)
# ═══════════════════════════════════════════════════════════════════

_MSGID_HDR_RE = re.compile(rb"Message-ID:\s*(<[^>]+>)", re.I)


def resolve_thread(account, thread_ref):
    """Return (gm_thrid, reply_to_msg_id) for the sheet's Thread ID value.

    The column holds either a Gmail API thread id (HEX, legacy n8n rows) or the
    SMTP Message-ID of our first email (rows sent by b2b_sender). Both are
    resolved over IMAP into the numeric X-GM-THRID that ec.fetch_thread needs,
    plus the Message-ID of the newest message in the thread (the right target
    for an in-thread reply). When IMAP is unavailable, returns ("", message_id)
    so a Message-ID row can still be replied to in-thread. Never raises."""
    ref = (thread_ref or "").strip()
    msg_id_ref = ""
    if "@" in ref:
        msg_id_ref = ref if ref.startswith("<") else "<%s>" % ref
    gm = ""
    if ref.isdigit():
        gm = ref
    elif not msg_id_ref and re.fullmatch(r"[0-9a-fA-F]{10,20}", ref):
        try:
            gm = str(int(ref, 16))  # Gmail API hex thread id -> X-GM-THRID
        except Exception:
            gm = ""
    if not account.get("password") or (not gm and not msg_id_ref):
        return "", msg_id_ref
    M = None
    try:
        M = imaplib.IMAP4_SSL(ec.IMAP_HOST, ec.IMAP_PORT)
        M.login(account["user"], account["password"])
        for box in ("[Gmail]/All Mail", "[Google Mail]/All Mail"):
            typ, _ = M.select(box, readonly=True)
            if typ == "OK":
                break
        else:
            return "", msg_id_ref
        # Message-ID row: look up its numeric thread id first.
        if not gm and msg_id_ref:
            typ, data = M.uid("SEARCH", None, "HEADER", "Message-ID",
                              msg_id_ref.strip("<>"))
            if typ == "OK" and data and data[0]:
                uid = data[0].split()[0]
                typ, md = M.uid("FETCH", uid, "(X-GM-THRID)")
                if typ == "OK" and md and md[0]:
                    meta = md[0][0] if isinstance(md[0], tuple) else md[0]
                    gm = ec._gm_thrid_from_fetch(meta)
        # Newest message in the thread = what the reply should reference.
        last_mid = ""
        if gm:
            typ, data = M.uid("SEARCH", None, "X-GM-THRID", gm)
            if typ == "OK" and data and data[0]:
                last_uid = data[0].split()[-1]
                typ, md = M.uid("FETCH", last_uid,
                                "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])")
                if typ == "OK" and md and md[0] and isinstance(md[0], tuple):
                    m = _MSGID_HDR_RE.search(md[0][1] or b"")
                    if m:
                        last_mid = m.group(1).decode("ascii", "replace")
        return gm, (last_mid or msg_id_ref)
    except Exception:
        return "", msg_id_ref
    finally:
        if M is not None:
            try:
                M.logout()
            except Exception:
                pass


def fetch_prior_thread(contact):
    """Pull the real correspondence for this contact when we can (app password
    present + row has a Thread ID). Returns (thread_list, reply_to_msg_id):
    a chronological fetch_thread list ([] when IMAP is unavailable / thread not
    found) and the Message-ID an in-thread reply should reference ("" if none)."""
    if not contact.get("thread_id"):
        return [], ""
    gm_thrid, reply_mid = resolve_thread(ACCOUNT, contact["thread_id"])
    if not gm_thrid:
        return [], reply_mid
    thread = ec.fetch_thread(ACCOUNT, gm_thrid, own_addresses=[ACCOUNT["user"]])
    return thread, reply_mid


def format_thread_history(thread, max_chars=4000):
    """Render a fetch_thread() list into a compact chronology for the model."""
    if not thread:
        return ""
    lines = []
    for m in thread:
        who = "UTD (us)" if m.get("direction") == "sent" else "Prospect"
        atts = m.get("attachment_names") or []
        att_note = f" | attachments: {', '.join(atts)}" if atts else ""
        snippet = (m.get("snippet") or "").strip()
        lines.append(f"- [{m.get('date', '')}] {who} | subject: "
                     f"{m.get('subject', '')}{att_note}\n  {snippet}")
    return "\n".join(lines)[:max_chars]


def correspondence_summary(contact):
    """Faithful summary used when the real thread cannot be pulled over IMAP."""
    subj = contact.get("orig_subject") or "UTD Referral program: partnership invitation"
    return (f"first outreach sent on {contact.get('date_sent') or 'an earlier date'}, "
            f"subject \"{subj}\", referral-program invitation, no reply received")


# ═══════════════════════════════════════════════════════════════════
#   BUILD CLAUDE REQUEST  (canon follow-up prompt)
# ═══════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = (
    'You are Sergey, a partnership manager at UTD Web (utdweb.team), a Shopify theme studio with 25 themes '
    'published on Shopify\'s official Theme Store (https://themes.shopify.com/themes?q=UTD).\n'
    '\n'
    'You write a follow-up email to a web or ecommerce agency that received our UTD Referral program '
    'invitation and has not replied. Program in one line: agencies that build client stores on UTD themes earn '
    'a commission on theme sales, plus priority technical support, early access to new theme versions, and '
    'access to UTD services (custom development, content, digital marketing, SEO).\n'
    '\n'
    'GOAL: SELL. This follow-up must add NEW selling substance, not nudge politely. Get a reply and move this '
    'agency toward the next funnel step (program details memo, then agreement). All persuasion happens by email: '
    'never suggest a call or a meeting, of any length, ever. Confident and direct, never whiny or guilt-tripping.\n'
    '\n'
    'VOICE:\n'
    '- You are a real person: Sergey, a normal guy who works at an IT company, writing an ordinary work email to '
    'another business owner. You are NOT a marketer and this is not ad copy. Plain everyday English, simple '
    'words, natural flow.\n'
    '- Confident and direct: state what they get and why it makes them money. Aggression through concreteness '
    'and numbers, never through hype adjectives.\n'
    '- Read-aloud test: if you would not say a sentence out loud to a colleague, rewrite it until you would.\n'
    '- Zero filler, zero marketing-speak, no dramatic one-liners, no "That\'s where X fits" constructions, no '
    'rhetorical devices. Just say what you mean, simply. Never "I hope this finds you well".\n'
    '- Length: exactly as long as needed to fully convey the point. No hard word cap; longer is fine when the '
    'substance requires it. But every sentence must carry real information, nothing decorative.\n'
    '\n'
    'STRUCTURE (mandatory, SHORT paragraphs, each its own idea, blank line between; never one big block):\n'
    '1. Greeting line: "Hi [Company Name] team," then a blank line.\n'
    '2. Reminder paragraph, one or two sentences max. The FIRST sentence must explicitly remind them of the '
    'earlier email: "I emailed you last week about our referral program for agencies..."\n'
    '3. NEW-substance paragraph(s): bring numbers and facts that were NOT in the first email. Best material: the '
    'money math (our themes sell for $100 to $340, flagship Impression at $340; put our themes into ten client '
    'builds a year and the commissions on those sales stack on top of normal project fees, exact rate in the '
    'memo) OR the time math (a from-scratch storefront theme is weeks of dev time, starting from a ready '
    'Shopify-reviewed theme is days, and our priority support takes setup questions off their plate). Pick '
    'whichever angle the first email did NOT use. Split into two short paragraphs if both fit naturally. Never '
    'repeat the first email\'s pitch wholesale.\n'
    '4. Closing paragraph: one simple question inviting a reply.\n'
    '- NO signature, sign-off, subject line, or footer. Output ONLY the email body, starting with the greeting.\n'
    '\n'
    'HARD RULES:\n'
    '1. Never use an em dash anywhere in the letter. Use a comma, colon, or period instead.\n'
    '2. Forbidden words: exclusive, exciting, game-changer, handpicked, curated, unique opportunity. No hype, no '
    'corporate slop ("in today\'s world", "take it to the next level", "seamless", "revolutionary").\n'
    '3. Never invent features, numbers, prices, or client results. The only links allowed: https://utdweb.team '
    'and https://themes.shopify.com/themes?q=UTD.\n'
    '4. Do NOT state exact commission percentages or thresholds. The numbers come later in the program memo.\n'
    '5. NO calls or meetings, ever. Do not offer or ask for a call. The only help-offer allowed is in the spirit '
    'of "reply and I\'ll walk you through it". Close with a simple question inviting a reply, never a "yes or '
    'no?" style demand.\n'
    '6. If the prospect\'s website or their emails are clearly in a language other than English, write the '
    'entire email in that language. Otherwise write in English.'
)


def build_claude_request(contact, site_text, history_text):
    """Return (system, user_prompt) for the canon follow-up letter."""
    if history_text:
        prior = ('Prior correspondence (chronological, oldest first):\n' + history_text)
    else:
        prior = 'Prior correspondence (summary): ' + correspondence_summary(contact)
    dt = _parse_date_sent(contact.get("date_sent"))
    days = (datetime.now(timezone.utc) - dt).days if dt else FOLLOWUP_AFTER_DAYS
    user_prompt = (
        'Company: ' + contact["company_name"] +
        '\nWebsite: ' + (contact.get("website") or "unknown") +
        '\nFirst email sent: ' + (contact.get("date_sent") or "unknown") +
        f' ({days} days ago, no reply)' +
        '\n\nWebsite content (fresh scrape, may be partial):\n' +
        (site_text or 'Site content unavailable.') +
        '\n\n' + prior +
        '\n\nWrite the follow-up email body now, in your natural human voice, as long as it '
        'needs to be and no longer. Start by explicitly reminding them of the earlier email, '
        'add ONE new concrete angle in its own paragraph, and close with a simple question '
        'inviting a reply. Output ONLY the body, starting with the greeting. '
        'No subject line, no signature.'
    )
    return SYSTEM_PROMPT, user_prompt


def clean_ai_body(text, company_name=""):
    """Normalize the Claude draft: drop stray SUBJECT/BODY labels, trailing
    sign-offs and em dashes, guarantee the greeting line, then append the
    fixed canon signature."""
    body = (text or "").strip()
    body = re.sub(r"^SUBJECT:.*\n", "", body, count=1, flags=re.I)
    body = re.sub(r"^BODY:\s*\n?", "", body, count=1, flags=re.I).strip()
    body = strip_trailing_signoff(body)
    body = strip_em_dashes(body)
    body = ensure_greeting(body, company_name)
    return body + SIG


# ═══════════════════════════════════════════════════════════════════
#   FALLBACK  (existing static template, used when Claude fails)
# ═══════════════════════════════════════════════════════════════════

def _fallback_body(contact):
    return (
        'Hi ' + contact["company_name"] + ' team,\n\n'
        'I am following up on my email from last week regarding a potential partnership.\n\n'
        'To briefly recap: we are UTD Web, a Shopify theme studio. Our themes are available on Shopify\'s official '
        'Theme Store (https://themes.shopify.com/themes?q=UTD). Through our UTD Referral program, web agencies that build '
        'client stores using our themes receive a referral fee on each purchase, with no additional overhead.\n\n'
        'If the timing is not right or this falls outside your area of work, please disregard this message entirely. '
        'However, if there is any interest, reply here and I will answer your questions and walk you through the details.'
        + SIG
    )


def _original_subject(contact, thread):
    """Best-known subject of the first outreach: real thread > sheet > default."""
    for m in thread or []:
        if m.get("direction") == "sent":
            s = (m.get("subject") or "").strip()
            if s:
                return re.sub(r"^(?:re|fwd?):\s*", "", s, flags=re.I)
    if contact.get("orig_subject"):
        return contact["orig_subject"]
    return "UTD Referral program: partnership invitation"


def build_followup_email(contact):
    """Assemble the full follow-up: context (site + prior thread) → Claude draft
    per canon → fallback template on failure. Returns
    {email, company_name, email_subject, email_body, in_reply_to}."""
    # 1) Fresh site scrape (same helper the cold sender uses).
    site_text = ""
    if contact.get("website"):
        site_text = _clean_site_text(fetch_company_website(contact["website"]))

    # 2) Prior correspondence: real IMAP thread when possible, else summary.
    thread, reply_mid = fetch_prior_thread(contact)
    history_text = format_thread_history(thread)
    if thread:
        print(f"  [thread] pulled {len(thread)} prior message(s) over IMAP")
    else:
        print("  [thread] IMAP thread unavailable -> using summary of first outreach")

    # 3) Claude draft (canon), with the static template as fallback.
    system, user_prompt = build_claude_request(contact, site_text, history_text)
    if DRY_RUN:
        # Review aid: dump the exact prompts before the (possibly key-less) call.
        print_prompt_for_review("b2b follow-up", system, user_prompt)
    ai_text = ec.call_claude(system, user_prompt, model=MODEL, max_tokens=MAX_TOKENS)
    if ai_text:
        body = clean_ai_body(ai_text, contact["company_name"])
    else:
        print("[claude unavailable -> fallback used]")
        body = _fallback_body(contact)

    # 4) Threading: reply inside the original conversation when we know the
    #    Message-ID to reference (resolved over IMAP, or stored on the row).
    in_reply_to = reply_mid
    subject = "Re: " + strip_em_dashes(_original_subject(contact, thread))

    return {
        "email": contact["email"],
        "company_name": contact["company_name"],
        "email_subject": subject,
        "email_body": body,
        "in_reply_to": in_reply_to,
    }


# ═══════════════════════════════════════════════════════════════════
#   SEND + SHEET WRITE  (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def _print_draft(to, subject, body, in_reply_to=""):
    print("\n" + "=" * 70)
    print("[DRAFT · b2b follow-up]  DRY_RUN — not sent")
    print(f"  from   : {SENDER_NAME} <{ACCOUNT['user']}>")
    print(f"  to     : {to}")
    print(f"  subject: {subject}")
    print(f"  thread : {'in-thread reply to ' + in_reply_to if in_reply_to else 'fresh email (no reply Message-ID available)'}")
    print("  body:")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def send_and_mark(ws, header, contact, drafted, state, stats):
    to = contact["email"]
    subject = drafted["email_subject"]
    body = drafted["email_body"]
    in_reply_to = drafted.get("in_reply_to") or None
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    if DRY_RUN:
        _print_draft(to, subject, body, in_reply_to or "")
        print(f"  [SHEET] DRY_RUN — would set row {contact['row_number']} → "
              f"Status='Follow-up Sent', Date Sent='{now}' (match Email={to})")
        stats["sent"] += 1
        return

    # send_email builds the References header from in_reply_to itself.
    ec.send_email(ACCOUNT, to, subject, body, in_reply_to=in_reply_to)
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
              f"({contact['company_name']}) | site={contact.get('website') or '-'} | "
              f"thread={'yes' if contact.get('thread_id') else 'no'}")
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
