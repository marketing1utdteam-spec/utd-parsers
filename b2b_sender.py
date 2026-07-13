#!/usr/bin/env python3
"""
b2b_sender.py — UTD "Referral program" B2B cold-outreach sender.

Faithful Python port of the n8n workflow B2B_send (Oc9j9TViJKwBNpeV) for
GitHub Actions. One run picks ONE uncontacted web agency from the CRM sheet,
scrapes its website, asks Claude to write a personalized referral-program
invitation (canon: plain human voice, ONE concrete site observation, length as
needed, paragraph format, no em dashes, no hype words, no invented facts,
email-only ask, never a call), sends it over SMTP (Gmail app-password) and
marks the row "Sent".

n8n → Python node mapping:
  Get Contacts from Sheet   → ec.open_worksheet + ec.read_rows_ws (ONE read)
  Pick Next Uncontacted     → pick_next_uncontacted()  (random valid row)
  Fetch Company Website     → fetch_company_website()
  Build Claude Request      → build_claude_request()   (SYSTEM/user verbatim)
  Claude — Write Email      → ec.call_claude(model claude-sonnet-5, 1300 tok)
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
import hashlib
from datetime import datetime, timezone

import email_common as ec

try:
    import requests
except Exception:  # pragma: no cover
    requests = None


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get("B2B_SHEET_ID", "")
SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")

# Cold-outreach mailbox — "Sergey | UTD Web". Same app-password env as the
# autoresponder's primary mailbox (email_common account convention).
ACCOUNT = {
    "user": os.environ.get("B2B_SENDER_USER", os.environ.get("UTD_MAIL_SERGEY", "")),
    "password": os.environ.get("SENDER_APP_PW") or os.environ.get("GMAIL_APP_PW_SERGEY", ""),
}
SENDER_NAME = "Sergey | UTD Web"  # n8n Gmail senderName (see assumptions)

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# n8n runs the send chain on a schedule and picks ONE random uncontacted row per
# trigger. We keep that "one per run" default; raise B2B_SEND_LIMIT to batch.
SEND_LIMIT = int(os.environ.get("B2B_SEND_LIMIT", "1"))

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "b2b_sender_state.json")

MODEL = "claude-sonnet-5"
MAX_TOKENS = 1300


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
    # junk guard: sane local part + real-looking domain + valid-length TLD
    if len(parts[0]) < 2 or len(domain) < 6 or len(domain.split(".")[-1]) < 2:
        return False
    if re.match(r"^[0-9a-f]{20,}$", parts[0], re.I):
        return False
    return True


def clean_company_name(name):
    """Company names in the sheet come from page <title> tags and often carry
    taglines ("Monforte Studio | Web Design and Branding"). Keep the brand part
    so the greeting reads "Hi Monforte Studio team,". Shared with b2b_followup."""
    n = (name or "").strip()
    n = re.split(r"\s*[|•]\s*|\s+[-–—]\s+", n)[0].strip()
    return n or "Unknown"


def pick_next_uncontacted(rows):
    """Reproduce «Pick Next Uncontacted»: Status empty + valid email + Website
    present + not a theme competitor, then choose ONE row at random.

    rows: list of (row_number, record_dict). Returns a contact dict
    {row_number, email, company_name, website} or None when nothing qualifies.
    """
    # Owner override: a row with Status='NEXT' is sent FIRST, bypassing the
    # validity blocklists (deliberate manual insert, e.g. an internal test).
    for row_number, r in rows:
        if str(r.get("Status", "")).strip() == "NEXT" and "@" in str(r.get("Email", "")):
            website = str(r.get("Website", "")).strip()
            return {
                "row_number": row_number,
                "email": str(r.get("Email", "")).strip(),
                "company_name": clean_company_name(str(r.get("Company Name", "") or r.get("Company", ""))),
                "website": website if website.startswith("http") else "https://" + website,
            }
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
    # Sheet header is "Company Name"; "Company" kept as a fallback.
    company = clean_company_name(
        str(r.get("Company Name", "") or r.get("Company", "")))
    return {
        "row_number": row_number,
        "email": str(r.get("Email", "")).strip(),
        "company_name": company,
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
    'You are Sergey, a partnership manager at UTD Web (utdweb.team), a Shopify theme studio: 6 themes with 30 '
    'ready-made store designs on Shopify\'s official Theme Store (https://themes.shopify.com/themes?q=UTD), '
    'used by thousands of merchants.\n'
    '\n'
    'You write a cold outreach email inviting a web or ecommerce agency to the UTD Referral program.\n'
    '\n'
    'GOAL: SELL participation in the referral program. Get a reply and move them to the next step (program memo, '
    'then agreement). All persuasion happens by email: never suggest a call or a meeting, ever.\n'
    '\n'
    'VOICE (this is the most important part):\n'
    '- Write like a normal person, as if the email was typed by hand in a few minutes. It must NOT read like AI '
    'text or ad copy.\n'
    '- SIMPLE ENGLISH: the reader is often not a native English speaker. Use common everyday words and short, '
    'simple sentences. No idioms, no slang, no fancy phrases (never "caught my eye", "worth a look" and '
    'similar). If a 12-year-old would not understand a sentence, rewrite it.\n'
    '- CLARITY ABOVE ALL: each paragraph instantly understandable on first read.\n'
    '- Maximum concreteness, zero water: names, numbers, dollars, hours. Cut every sentence that does not carry '
    'a fact or move the deal.\n'
    '- Paragraphs are SHORT: 1-3 sentences each. Never merge two topics into one paragraph.\n'
    '- No hype adjectives, no rhetorical devices, no "I hope this finds you well".\n'
    '\n'
    'STRUCTURE (mandatory, each numbered item is its OWN paragraph, blank line between):\n'
    '1. Greeting line: "Hi [Company Name] team," then a blank line.\n'
    '2. ABOUT THEM ONLY. Start this paragraph with a simple opener like "I was on your website and saw..." or '
    '"I visited your site and noticed...". Then 2-3 more simple sentences: what their company does, one real '
    'specific thing from their site (a project, a service, a client), and one more sentence saying plainly WHY '
    'that made us write to them (what about their work fits our program). Plain words anyone understands. '
    'Nothing about UTD in this paragraph. If the website content could not be loaded, NEVER mention errors, 403s or that you could not open the site: write the observation from the company name and what kind of agency they are, or keep the paragraph to one modest sentence.\n'
    '3. ABOUT US ONLY. Who UTD Web is and why we matter: we build Shopify themes, 6 themes with 30 ready-made '
    'designs live on Shopify\'s official Theme Store (link https://themes.shopify.com/themes?q=UTD), Shopify '
    'reviews every theme before listing, thousands of merchants run on them, and conversion features (upsells, '
    'cross-sells, promo blocks) are built in so stores need fewer paid apps. Our site: https://utdweb.team.\n'
    '4. THE OFFER + THE MONEY, concrete. Start with the offer: "We are inviting agencies like yours to our '
    'referral program." HOW IT WORKS (important, never say "bring your clients to us" or imply we take over '
    'the client): YOUR client buys the theme themselves through the official Shopify Theme Store, you keep '
    'the client and the project; we then help set up their store and give them the best support they can '
    'get, free of charge, and you get the commission on that sale. The math, plainly: commission starts at '
    '10% and grows to 15% as you sell more. Our flagship Impression costs $340, so ten client stores on '
    'Impression = $340 back at the starting rate, up to $510 at the top rate, on top of your normal project '
    'fees. Plus priority support and early access to new theme versions for you.\n'
    '5. THE HOURS AND THE MONEY SAVED, concrete (may be two short paragraphs). First the hours: a theme from '
    'scratch takes one developer roughly 100+ hours; with a ready theme it is a couple of days, so around '
    '80-100 hours of one specialist saved on every project. Then the dollars: a developer who sets up Shopify '
    'themes typically costs $40-60 per hour in the US, so that is roughly $3,200-6,000 saved on every single '
    'project. Then HOW our support removes even more work, 2-3 simple sentences: we help install and set up '
    'the theme, we help build the client sites on it, we explain how the theme works and what is inside it '
    '(sections, upsell blocks, settings), we answer your developer\'s questions fast and fix problems '
    'ourselves, so your team never sits reading documentation or debugging.\n'
    '6. Closing ask, one line, answerable by email ("Reply and I\'ll send the program memo").\n'
    '- NO signature, sign-off, or footer after the body. It is added separately.\n'
    '\n'
    'HARD RULES:\n'
    '1. Never use an em dash anywhere in the letter. Use a comma, colon, or period instead.\n'
    '2. Forbidden words: exclusive, exciting, game-changer, handpicked, curated, unique opportunity. No hype, no '
    'corporate slop ("in today\'s world", "take it to the next level", "seamless", "revolutionary").\n'
    '3. Never invent features, numbers, prices, or client results. The only links allowed: https://utdweb.team '
    'and https://themes.shopify.com/themes?q=UTD.\n'
    '4. Do NOT state exact commission percentages or thresholds. Say there is a commission on theme sales; the '
    'exact numbers come later in the program memo.\n'
    '5. NO calls or meetings, ever. Do not offer or ask for a call. End with ONE simple ask that can be answered '
    'by email, in the spirit of "reply and I\'ll send over the program details" or "reply and I\'ll walk you '
    'through how it works".\n'
    '6. Subject line MUST be unique to THIS agency — tie it to their niche, their work, or the one '
    'observation you make in the email. Vary it every single time; two different agencies must never get '
    'the same subject. Do NOT use the boilerplate phrase "UTD referral program for agencies" or any fixed '
    'template line. Properly capitalized, natural, like a real work email subject, never clickbait. '
    'STYLE examples only (invent your own, do not copy): "A partner idea for your Shopify builds", '
    '"Earning on the themes you already recommend", "For [what they do] — a referral angle".\n'
    '7. If the website content is clearly in a language other than English, write the entire email in that '
    'language. Otherwise write in English.'
)


def build_claude_request(contact, html):
    """Return (system, user_prompt) for the canon referral-invitation letter."""
    site_text = _clean_site_text(html)
    user_prompt = (
        'Company: ' + contact["company_name"] +
        '\nWebsite: ' + contact["website"] +
        '\nWebsite content (scraped, may be partial):\n' + site_text +
        '\n\nWrite the referral-program invitation email now, in your natural human voice, '
        'as long as it needs to be and no longer. '
        'Ground it in ONE concrete observation from the website content above, mentioned naturally. '
        'Reply ONLY in this exact format:\n'
        'SUBJECT: [a subject SPECIFIC to this agency — tie it to their niche or your observation; '
        'unique, not a generic/boilerplate line, no em dash, no forbidden words]\n'
        'BODY:\n'
        '[email body: greeting line, blank line, paragraphs separated by blank lines, no signature, no sign-off]'
    )
    return SYSTEM_PROMPT, user_prompt


def print_prompt_for_review(tag, system, user):
    """DRY_RUN aid: dump the exact SYSTEM + user prompt so the letter brief can
    be reviewed without an ANTHROPIC_API_KEY. Shared with b2b_followup."""
    bar = "-" * 70
    print("\n" + bar)
    print(f"[PROMPT REVIEW · {tag}] SYSTEM prompt sent to Claude:")
    print(bar)
    print(system)
    print(bar)
    print(f"[PROMPT REVIEW · {tag}] USER prompt sent to Claude:")
    print(bar)
    print(user)
    print(bar)


def strip_em_dashes(text):
    """Copy canon: never an em dash in an outgoing letter. Defensive scrub in
    case the model slips one through."""
    if not text:
        return text
    return text.replace(" — ", " - ").replace("—", "-")


def ensure_greeting(body, company_name):
    """Format canon: every email starts with a greeting line. If the model's
    body does not open with "Hi ", prepend "Hi [company] team,". Shared with
    b2b_followup."""
    b = (body or "").lstrip()
    if not b.startswith("Hi "):
        b = "Hi " + (company_name or "there") + " team,\n\n" + b
    return b


def strip_trailing_signoff(body):
    """Remove any trailing sign-off / footer the model appended despite the
    rules (the fixed signature is added separately). Shared with b2b_followup."""
    body = re.sub(
        r"\n{1,3}(?:best\s*regards?|sincerely|cheers|warm\s*regards?|regards?|yours\s+faithfully|yours\s+sincerely)[,.]?[\s\S]*$",
        "", body, flags=re.I)
    body = re.sub(r"\n{1,3}\[your[\s_]*name\][\s\S]*$", "", body, flags=re.I)
    body = re.sub(r"\n{1,3}-{2,}\s*\n[\s\S]*$", "", body)
    body = re.sub(r"\n{1,3}this\s+email\s+was\s+sent[\s\S]*$", "", body, flags=re.I)
    return re.sub(r"\n{3,}", "\n\n", body).strip()


# ═══════════════════════════════════════════════════════════════════
#   PARSE + CLEAN EMAIL  (verbatim from «Parse + Clean Email»)
# ═══════════════════════════════════════════════════════════════════

# Varied fallback subjects (used only when Claude fails to return a subject) —
# picked by a stable hash of the company so no two agencies get the same line.
_FALLBACK_SUBJECTS = [
    "A referral idea for your agency",
    "Earning on the Shopify themes you recommend",
    "A partner angle for your Shopify builds",
    "For agencies that build Shopify stores",
    "A theme referral idea worth a look",
    "Partnering with UTD on theme sales",
    "An idea for the stores you build",
    "A commission angle on Shopify themes",
]


def _fallback_subject(company_name):
    h = int(hashlib.sha256((company_name or "x").encode("utf-8")).hexdigest(), 16)
    return _FALLBACK_SUBJECTS[h % len(_FALLBACK_SUBJECTS)]


def _fallback_text(company_name):
    """The no-AI fallback body from the «Parse + Clean Email» catch block
    (used when Claude returns nothing / an unparseable response)."""
    return (
        'SUBJECT: ' + _fallback_subject(company_name) + '\n'
        'BODY:\n'
        'Hi ' + company_name + ' team,\n\n'
        'Your work caught our attention, and we selected your company for the UTD Referral program.\n\n'
        'UTD Web is a Shopify theme studio with 6 themes and 30 ready-made designs on Shopify\'s official '
        'Theme Store (https://themes.shopify.com/themes?q=UTD). Our themes are used by thousands of merchants '
        'worldwide, and we invest continuously in performance, design quality, and merchant success.\n\n'
        'Through UTD Referral, partner agencies earn a commission on every theme sale and receive priority '
        'technical support, early access to new theme versions before public release, and access to our '
        'custom development, content, marketing, and SEO services.\n\n'
        'If this fits how you build client stores, reply and I will send over the program details. '
        'Would a short overview be useful?'
    )


SIG = '\n\nBest regards,\nSergey\nUTD Web | utdweb.team'


def parse_and_clean(text, company_name):
    """Reproduce «Parse + Clean Email»: extract SUBJECT/BODY, strip any trailing
    sign-off/footer, collapse blank lines, append the fixed signature.
    Returns {email_subject, email_body}."""
    if not text or not text.strip():
        text = _fallback_text(company_name)

    subm = re.search(r"SUBJECT:\s*(.+?)(?:\n|$)", text, re.I)
    bodm = re.search(r"BODY:\n([\s\S]+)", text, re.I)
    subject = subm.group(1).strip() if subm else _fallback_subject(company_name)
    if bodm:
        body = bodm.group(1).strip()
    else:
        body = re.sub(r"SUBJECT:.+?\n", "", text, count=1, flags=re.I)
        body = re.sub(r"BODY:", "", body, count=1, flags=re.I).strip()

    # SAFE STRIP: only trailing sign-off / footer anchored to end of string.
    body = strip_trailing_signoff(body)

    # Copy canon: no em dashes anywhere in the outgoing letter.
    subject = strip_em_dashes(subject)
    body = strip_em_dashes(body)

    # Format canon: the body must always open with a greeting line.
    body = ensure_greeting(body, company_name)

    return {"email_subject": subject, "email_body": body + SIG}


def build_email_template(contact):
    """Alternate plain «Build Email (template)» node (kept for parity; not in the
    default AI path). Returns {email_subject, email_body}."""
    raw = (contact.get("company_name") or "").strip()
    company = raw if (raw and raw.lower() != "unknown" and len(raw) < 60) else "your team"
    subject = _fallback_subject(raw or company)
    body = (
        f"Hi {company},\n\n"
        "I'm reaching out from UTD Web (utdweb.team), a Shopify theme studio with 6 themes and 30 ready-made designs on "
        "Shopify's official Theme Store (https://themes.shopify.com/themes?q=UTD), trusted by thousands of "
        "merchants worldwide.\n\n"
        "We're inviting selected web and ecommerce agencies to our UTD Referral program: agencies that build "
        "client stores using our themes earn a referral fee on every purchase, with no additional overhead on "
        "your side.\n\n"
        "If this fits how you work, reply and I will send over the details and walk you through how it works. "
        "If the timing is not right, please feel free to disregard this message.\n\n"
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
        if DRY_RUN:
            # Review aid: dump the exact prompts before the (possibly key-less) call.
            print_prompt_for_review("b2b send", system, user_prompt)
        ai_text = ec.call_claude(system, user_prompt, model=MODEL, max_tokens=MAX_TOKENS)
        if not ai_text:
            print("[claude unavailable -> fallback used]")
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
