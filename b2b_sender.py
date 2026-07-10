#!/usr/bin/env python3
"""
b2b_sender.py — UTD "Referral program" B2B cold-outreach sender.

Faithful Python port of the n8n workflow B2B_send (Oc9j9TViJKwBNpeV) for
GitHub Actions. One run picks ONE uncontacted web agency from the CRM sheet,
scrapes its website, asks Claude to write a four-paragraph partnership
invitation, sends it over SMTP (Gmail app-password) and marks the row "Sent".

n8n → Python node mapping:
  Get Contacts from Sheet   → ec.open_worksheet + ec.read_rows_ws (ONE read)
  Pick Next Uncontacted     → pick_next_uncontacted()  (random valid row)
  Fetch Company Website     → fetch_company_website()
  Build Claude Request      → build_claude_request()   (SYSTEM/user verbatim)
  Claude — Write Email      → ec.call_claude(model claude-sonnet-5, 900 tok)
  Parse + Clean Email       → parse_and_clean()  (+ no-AI fallback text)
  Send Outreach Email       → ec.send_email()
  Mark as Sent in Sheet     → Status "Sent" / Date Sent / Thread ID
  Every 20 Minutes          → one send per run (see B2B_SEND_LIMIT)

Safety:
  • DRY_RUN=true (default) prints the drafted email + intended sheet writes and
    does NOT send or write.
  • Sent emails are also SHA256-hashed into data/b2b_sender_state.json so the
    same lead is never emailed twice (belt-and-braces over the Status column).

Usage:  python b2b_sender.py
Env:    GOOGLE_CREDENTIALS_JSON, ANTHROPIC_API_KEY, GMAIL_APP_PW_SERGEY,
        B2B_SHEET_ID, B2B_SHEET_TAB, DRY_RUN, B2B_SEND_LIMIT, STATE_DIR
"""

import os
import re
import json
import random
from datetime import datetime, timezone

import email_common as ec

try:
    import requests
except Exception:  # pragma: no cover
    requests = None


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get(
    "B2B_SHEET_ID", "1ggMS5Hko2jCY5eqcPvasBy3P6hAwbw8rldr4cS3Zeo4")
SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")

# Cold-outreach mailbox — "Sergey | UTD Web". Same app-password env as the
# autoresponder's sergey.utd@gmail.com box (email_common account convention).
ACCOUNT = {
    "user": os.environ.get("B2B_SENDER_USER", "sergey.utd@gmail.com"),
    "password": os.environ.get("GMAIL_APP_PW_SERGEY", ""),
}
SENDER_NAME = "Sergey | UTD Web"  # n8n Gmail senderName (see assumptions)

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# n8n runs the send chain on a schedule and picks ONE random uncontacted row per
# trigger. We keep that "one per run" default; raise B2B_SEND_LIMIT to batch.
SEND_LIMIT = int(os.environ.get("B2B_SEND_LIMIT", "1"))

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "b2b_sender_state.json")

MODEL = "claude-sonnet-5"
MAX_TOKENS = 900


# ═══════════════════════════════════════════════════════════════════
#   LEAD SELECTION  (verbatim from «Pick Next Uncontacted» code node)
# ═══════════════════════════════════════════════════════════════════

# SHOPIFY THEME DEVELOPER BLOCKLIST — never pitch competitors
THEME_COMPETITOR_DOMAINS = ['shopify.com', 'archetypethemes.co', 'pixelunion.net', 'outofthesandbox.com', 'cleancanvas.co.nz', 'maestrooo.com', 'troopthemes.com', 'groupthought.com', 'eightthemes.com', 'weareunderground.com', 'corknine.com', 'switchthemes.co', 'safeasmilk.nl', 'krownthemes.com', 'milehighthemes.com', 'fluorescent.ca', 'trailblaze.media', 'trailblazethemes.com', 'invisiblethemes.com', 'presidio.build', 'pagemilldesign.com', 'brickspacelab.com', 'roartheme.com', 'boostertheme.com', 'wetheme.com', 'woolman.io', 'the4.co', 'stylehatch.co', 'staylime.com', 'redplugdesign.com', 'p-themes.com', 'superfinedigital.com', 'bsscommerce.com', 'emthemes.net', 'adornthemes.com', 'fuelthemes.net', 'webibazaar.com', 'slashthemes.in', 'designthemes.com', 'cssigniter.com', 'swissuplabs.com', 'foxecom.com', 'shinedesigninfo.com', 'digifist.com', 'thethemegoal.com', 'karmoon.design', 'envora.com', 'muupthemes.com', 'coquelicotthemes.com', 'barracuda.design', 'shopidevs.com', 'softalithemes.com', 'mpthemez.com', 'agnisoftware.com', 'harmoniks.com', 'openthinking.net', 'archer-commerce.com', 'nethypeco.com', 'saleshunter.io', 'kumi.studio', 'boostifythemes.com', 'templatemonster.com', 'themeforest.net', 'envato.com', 'themeisle.com', 'elegantthemes.com', 'utdweb.team']

BAD = ['denvdavydov', 'smortkin', 'utdweb.team', 'utd.agency', 'its_always_teatime', 'noreply@utd', '.png', '.jpg', '.webp', '.jpeg', '.svg', '.gif', '@sentry', 'ingest.sentry', 'ingest.us.sentry', 'your-company', 'john@company', 'jane@company', 'you@company', 'your@email', 'name@email', 'example@', 'placeholder', 'mymail@mailservice', 'johnsmith@email', '@text.tld', 'user@domain', 'abc@company', 'you@yourstore', 'john@johnson', 'jane@mycompany', 'joe.smith@email', 'contact@www.', 'info@www.', 'ihelp@www.', '%20', 'a%20href', 'sentry.io', 'sentry-next', 'rm463@sasktel']


def is_theme_competitor(email):
    domain = (email.split("@")[1] if "@" in email else "").lower()
    return any(domain == d or domain.endswith("." + d) for d in THEME_COMPETITOR_DOMAINS)


def is_bad(e):
    el = e.lower()
    return any(b.lower() in el for b in BAD)


def is_valid(e):
    if not e or "@" not in e or " " in e or len(e) < 7:
        return False
    if is_bad(e):
        return False
    parts = e.split("@")
    if len(parts) != 2:
        return False
    domain = parts[1]
    if "." not in domain or domain.startswith("www."):
        return False
    if re.match(r"^[0-9a-f]{20,}$", parts[0], re.I):
        return False
    return True


def pick_next_uncontacted(rows):
    """Reproduce «Pick Next Uncontacted»: Status empty + valid email + Website
    present + not a theme competitor, then choose ONE row at random.

    rows: list of (row_number, record_dict). Returns a contact dict
    {row_number, email, company_name, website} or None when nothing qualifies.
    """
    candidates = []
    for row_number, r in rows:
        s = str(r.get("Status", "")).strip()
        e = str(r.get("Email", "")).strip()
        w = str(r.get("Website", "")).strip()
        if s == "" and is_valid(e) and w != "" and not is_theme_competitor(e):
            candidates.append((row_number, r))
    if not candidates:
        return None
    row_number, r = candidates[random.randrange(len(candidates))]
    website = str(r.get("Website", "")).strip()
    url = website if website.startswith("http") else "https://" + website
    return {
        "row_number": row_number,
        "email": str(r.get("Email", "")).strip(),
        "company_name": str(r.get("Company", "")) or "Unknown",
        "website": url,
    }


# ═══════════════════════════════════════════════════════════════════
#   FETCH COMPANY WEBSITE  (verbatim from «Fetch Company Website»)
# ═══════════════════════════════════════════════════════════════════

def fetch_company_website(url):
    """GET the company site (allowUnauthorizedCerts, <=3 redirects, 12s timeout).
    Returns the raw HTML string, or "" on any failure."""
    if requests is None:
        return ""
    try:
        r = requests.get(url, timeout=12, verify=False, allow_redirects=True,
                         headers={"User-Agent": "Mozilla/5.0 (compatible; UTD-Outreach/1.0)"})
        return r.text or ""
    except Exception:
        return ""


# ═══════════════════════════════════════════════════════════════════
#   BUILD CLAUDE REQUEST  (SYSTEM + user prompt verbatim from spec)
# ═══════════════════════════════════════════════════════════════════

def _clean_site_text(html):
    """Reproduce the HTML→text cleanup from «Build Claude Request»."""
    site_text = "Site content unavailable."
    try:
        if html and len(html) > 80:
            t = html
            t = re.sub(r"<script[\s\S]*?</script>", "", t, flags=re.I)
            t = re.sub(r"<style[\s\S]*?</style>", "", t, flags=re.I)
            t = re.sub(r"<nav[\s\S]*?</nav>", "", t, flags=re.I)
            t = re.sub(r"<footer[\s\S]*?</footer>", "", t, flags=re.I)
            t = re.sub(r"<[^>]+>", " ", t)
            t = re.sub(r"&nbsp;|&amp;|&quot;", " ", t)
            t = re.sub(r"\s+", " ", t).strip()[:2500]
            site_text = t
    except Exception:
        pass
    return site_text


SYSTEM_PROMPT = (
    'You are Sergey, a partnership manager at UTD Web (utdweb.team). '
    'UTD Web is one of the leading Shopify theme studios with 25 themes published on Shopify\'s official Theme Store '
    '(themes.shopify.com/themes?q=UTD), trusted by thousands of merchants worldwide. '
    '\n\n'
    'You are writing a formal, professional partnership invitation to a web agency. '
    'The email must follow a precise four-paragraph structure — do not deviate from it. '
    '\n\n'
    'TONE: Professional, respectful, measured. No hype. No urgency. No pushy language. '
    'Sound like a senior business representative, not a marketer. '
    '\n\n'
    'FOUR-PARAGRAPH STRUCTURE — follow exactly in this order:\n'
    '\n'
    'PARAGRAPH 1 — Why we selected this company:\n'
    '  • 1–2 sentences with a specific, concrete observation about their work based on the website content.\n'
    '  • Final sentence (exact phrasing): "That is why we selected your company for our UTD Referral program."\n'
    '\n'
    'PARAGRAPH 2 — About UTD Web:\n'
    '  • State that UTD Web is among the leading Shopify theme studios.\n'
    '  • Mention 25 themes on Shopify\'s official Theme Store — reference the link naturally: themes.shopify.com/themes?q=UTD\n'
    '  • One brief reason why working with UTD is a sound choice (quality, merchant trust, ongoing investment).\n'
    '  • Maximum 2–3 sentences.\n'
    '\n'
    'PARAGRAPH 3 — UTD Referral program benefits (2 sentences maximum):\n'
    '  Summarise ALL of the following into no more than 2 sentences:\n'
    '  — Commission on every theme sale\n'
    '  — Complimentary top-tier technical support\n'
    '  — Early access to new theme versions before public release\n'
    '  — 10% discount on premium support packages\n'
    '  — Access to additional services: custom development, content, digital marketing, SEO\n'
    '\n'
    'PARAGRAPH 4 — Closing:\n'
    '  • One sentence on why this program suits their type of agency (based on their website).\n'
    '  • One sentence inviting a reply or brief conversation.\n'
    '\n'
    'HARD RULES:\n'
    '  1. Greeting: "Hi [Company Name] team," — actual company name, never a personal name.\n'
    '  2. Total body under 200 words — write tight, no padding.\n'
    '  3. NO signature, sign-off, or footer — added separately.\n'
    '  4. Do not use words: exclusive, exciting, game-changer, handpicked, curated, unique opportunity.'
)


def build_claude_request(contact, html):
    """Return (system, user_prompt) exactly as «Build Claude Request» builds them."""
    site_text = _clean_site_text(html)
    user_prompt = (
        'Company: ' + contact["company_name"] +
        '\nWebsite: ' + contact["website"] +
        '\nWebsite content:\n' + site_text +
        '\n\nWrite the invitation email following the four-paragraph structure. Under 200 words. '
        'Reply ONLY in this exact format:\n'
        'SUBJECT: [subject line]\n'
        'BODY:\n'
        '[four-paragraph email body — no signature, no sign-off]'
    )
    return SYSTEM_PROMPT, user_prompt


# ═══════════════════════════════════════════════════════════════════
#   PARSE + CLEAN EMAIL  (verbatim from «Parse + Clean Email»)
# ═══════════════════════════════════════════════════════════════════

def _fallback_text(company_name):
    """The no-AI fallback body from the «Parse + Clean Email» catch block
    (used when Claude returns nothing / an unparseable response)."""
    return (
        'SUBJECT: UTD Referral Program — Partnership Invitation\n'
        'BODY:\n'
        'Hi ' + company_name + ' team,\n\n'
        'We have been following your work and are impressed by the quality and consistency of your projects. '
        'That is why we selected your company for our UTD Referral program.\n\n'
        'UTD Web is among the leading Shopify theme studios, with 25 themes published on Shopify\'s official '
        'Theme Store (themes.shopify.com/themes?q=UTD). Our themes are trusted by thousands of merchants '
        'worldwide, and we invest continuously in performance, design quality, and merchant success.\n\n'
        'Through UTD Referral, partners earn a commission on every theme sale and receive complimentary '
        'top-tier technical support, early access to new theme versions prior to public release, and the '
        'ability to contribute to our product roadmap — alongside a 10% discount on premium support packages '
        'and access to custom development, content, marketing, and SEO services.\n\n'
        'Given the nature of your work, we believe this program aligns well with your existing offering. '
        'We would be glad to provide further details and look forward to hearing from you.'
    )


SIG = '\n\nBest regards,\nSergey\nUTD Web · utdweb.team'


def parse_and_clean(text, company_name):
    """Reproduce «Parse + Clean Email»: extract SUBJECT/BODY, strip any trailing
    sign-off/footer, collapse blank lines, append the fixed signature.
    Returns {email_subject, email_body}."""
    if not text or not text.strip():
        text = _fallback_text(company_name)

    subm = re.search(r"SUBJECT:\s*(.+?)(?:\n|$)", text, re.I)
    bodm = re.search(r"BODY:\n([\s\S]+)", text, re.I)
    subject = subm.group(1).strip() if subm else "UTD Referral Program — Partnership Invitation"
    if bodm:
        body = bodm.group(1).strip()
    else:
        body = re.sub(r"SUBJECT:.+?\n", "", text, count=1, flags=re.I)
        body = re.sub(r"BODY:", "", body, count=1, flags=re.I).strip()

    # SAFE STRIP: only trailing sign-off / footer anchored to end of string.
    body = re.sub(
        r"\n{1,3}(?:best\s*regards?|sincerely|cheers|warm\s*regards?|regards?|yours\s+faithfully|yours\s+sincerely)[,.]?[\s\S]*$",
        "", body, flags=re.I)
    body = re.sub(r"\n{1,3}\[your[\s_]*name\][\s\S]*$", "", body, flags=re.I)
    body = re.sub(r"\n{1,3}-{2,}\s*\n[\s\S]*$", "", body)
    body = re.sub(r"\n{1,3}this\s+email\s+was\s+sent[\s\S]*$", "", body, flags=re.I)
    body = re.sub(r"\n{3,}", "\n\n", body).strip()

    return {"email_subject": subject, "email_body": body + SIG}


def build_email_template(contact):
    """Alternate plain «Build Email (template)» node (kept for parity; not in the
    default AI path). Returns {email_subject, email_body}."""
    raw = (contact.get("company_name") or "").strip()
    company = raw if (raw and raw.lower() != "unknown" and len(raw) < 60) else "your team"
    subject = "Partnership opportunity — UTD Referral program"
    body = (
        f"Hi {company},\n\n"
        "I'm reaching out from UTD Web (utdweb.team), a Shopify theme studio with 25 themes published on "
        "Shopify's official Theme Store (https://themes.shopify.com/themes?q=UTD), trusted by thousands of "
        "merchants worldwide.\n\n"
        "We're inviting selected web and ecommerce agencies to our UTD Referral program: agencies that build "
        "client stores using our themes earn a referral fee on every purchase, with no additional overhead on "
        "your side.\n\n"
        "If this fits how you work, I would be glad to share the details or arrange a short call at a time that "
        "suits you. If the timing is not right, please feel free to disregard this message.\n\n"
        "Best regards,\nSergey\nUTD Web - utdweb.team"
    )
    return {"email_subject": subject, "email_body": body}


# ═══════════════════════════════════════════════════════════════════
#   SEND + SHEET WRITE  (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def _print_draft(account, to, subject, body):
    print("\n" + "=" * 70)
    print("[DRAFT · b2b send]  DRY_RUN — not sent")
    print(f"  from   : {SENDER_NAME} <{account['user']}>")
    print(f"  to     : {to}")
    print(f"  subject: {subject}")
    print("  body:")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def send_and_mark(ws, header, contact, drafted, state, stats):
    """Send the outreach email and mark the row Sent (unless DRY_RUN)."""
    to = contact["email"]
    subject = drafted["email_subject"]
    body = drafted["email_body"]
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    if DRY_RUN:
        _print_draft(ACCOUNT, to, subject, body)
        print(f"  [SHEET] DRY_RUN — would set row {contact['row_number']} → "
              f"Status='Sent', Date Sent='{now}', Thread ID='<msg-id>'")
        stats["sent"] += 1
        return

    msg_id = ec.send_email(ACCOUNT, to, subject, body)
    # Persist to CRM: Status / Date Sent / Thread ID on the exact row.
    updates = {"Status": "Sent", "Date Sent": now, "Thread ID": msg_id or ""}
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
    print(f"  [SHEET] row {contact['row_number']} → Status='Sent' (msg {msg_id})")


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def run_once():
    print(f"=== UTD B2B cold-sender | DRY_RUN={DRY_RUN} | limit={SEND_LIMIT} | "
          f"{datetime.now(timezone.utc).isoformat()} ===")
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

    # (row_number, record) — header is row 1, data starts at row 2.
    rows = list(enumerate(records, start=2))
    print(f"CRM: {len(records)} rows read in 1 call.")

    for _ in range(max(1, SEND_LIMIT)):
        contact = pick_next_uncontacted(rows)
        if not contact:
            print("No uncontacted contacts left → stop.")
            break
        stats["considered"] += 1

        # State-level dedup (never email the same lead twice).
        if ec.is_processed(state, contact["email"]):
            print(f"· {contact['email']} already emailed (state) → skip.")
            stats["skipped"] += 1
            # prevent re-picking this same row in this run
            rows = [(rn, r) for (rn, r) in rows if rn != contact["row_number"]]
            continue

        print(f"\n· picked row {contact['row_number']} → {contact['email']} "
              f"({contact['company_name']}) | {contact['website']}")

        html = fetch_company_website(contact["website"])
        system, user_prompt = build_claude_request(contact, html)
        ai_text = ec.call_claude(system, user_prompt, model=MODEL, max_tokens=MAX_TOKENS)
        drafted = parse_and_clean(ai_text, contact["company_name"])

        try:
            send_and_mark(ws, header, contact, drafted, state, stats)
        except Exception as e:
            stats["errors"].append(f"send failed for {contact['email']}: {e}")
            print(f"  !! send error: {e}")

        # Do not re-pick this row within the same run.
        rows = [(rn, r) for (rn, r) in rows if rn != contact["row_number"]]

    if not DRY_RUN:
        ec.save_state(STATE_FILE, state)

    print(f"\n=== done. {stats} ===")
    return stats


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
