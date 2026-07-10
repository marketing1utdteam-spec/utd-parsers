#!/usr/bin/env python3
"""
ecom_sender.py — UTD Shopify-theme cold-outreach sequence for eCommerce stores.

Faithful port of the n8n workflow ECOM_seq (g6sutQHdbTXsPaN5) to plain Python
for GitHub Actions. Reads the "Ecom Contacts" CRM sheet (the one the ecom
harvester fills), picks ONE due store contact per run, drafts a per-merchant
email with Claude that references the store + industry and pitches UTD themes,
sends it (new thread for the first touch, in-thread reply for follow-ups over
SMTP), and writes the send state back to the sheet.

Sequence (Status column drives the touch, exactly as in the n8n «Pick Next»):
  Status ''         → touch 1 (first cold email, new thread)      → Status "Sent"
  Status 'Sent'      + Date Sent >= 4d → touch 2 (bump, reply)     → "Followup1"
  Status 'Followup1' + Date Sent >= 4d → touch 3 (value, reply)    → "Followup2"
  Status 'Followup2' + Date Sent >= 6d → touch 4 (breakup, reply)  → "Sequence Done"

Follow-ups (touch>1) are prioritised over new leads; within the chosen tier one
contact is picked at random (mirrors the n8n «Pick Next» code node).

Safety:
  • DRY_RUN=true (default) prints the draft + intended sheet write, sends nothing.
  • email_common state dedup (SHA256 of email|touch) stops a double-send of the
    same touch to the same contact (repo is PUBLIC — no raw addresses committed).

Usage:  python ecom_sender.py
Env:    GOOGLE_CREDENTIALS_JSON, ANTHROPIC_API_KEY, GMAIL_APP_PW_SERGEY,
        ECOM_SHEET_ID, ECOM_SHEET_TAB, DRY_RUN, STATE_DIR, LEAD_LIMIT
"""

import os
import re
import json
import random
import imaplib
from datetime import datetime, timezone

import email_common as ec

try:
    import requests
except Exception:  # pragma: no cover
    requests = None


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

# Same env var names / defaults as ecom_harvester.py so sender + harvester
# read and write the SAME sheet tab. (In production ECOM_SHEET_ID is supplied
# via the GitHub Actions env / a GitHub Variable and ECOM_SHEET_TAB is
# "Ecom Contacts" — see the harvester's REQUIRED-no-default convention.)
SHEET_ID = os.environ.get("ECOM_SHEET_ID", "").strip()
SHEET_TAB = os.environ.get("ECOM_SHEET_TAB", "Contacts")

# The mailbox we send from — persona "Sergey" (email_common "account" convention:
# {"user","password"} with the Gmail app-password pulled from env by name).
ACCOUNT = {
    "user": os.environ.get("ECOM_SENDER_EMAIL", os.environ.get("UTD_MAIL_SERGEY", "")),
    "password": os.environ.get("GMAIL_APP_PW_SERGEY", ""),
}
# Gmail "From" display name used by the n8n Gmail node (senderName).
SENDER_NAME = "Sergey | UTD Web"

# Our own mailbox addresses — used to tag directions when reading the thread
# back over IMAP (fetch_thread marks messages from these as 'sent').
OWN_ADDRESSES = [a for a in (
    os.environ.get("UTD_MAIL_SERGEY", ""),
    os.environ.get("UTD_MAIL_SERGI", ""),
    os.environ.get("UTD_MAIL_SERGE", ""),
    os.environ.get("UTD_MAIL_SERHII", ""),
) if a]

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "ecom_sender_state.json")

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# n8n «Расписание» fires once and «Pick Next» sends to ONE contact per run.
# LEAD_LIMIT keeps that default (1) but allows a small batch per GHA run.
LEAD_LIMIT = int(os.environ.get("LEAD_LIMIT", "1"))

# HTTP fetch of the store homepage for touch-1 personalisation (n8n
# «Fetch Store Website»: allowUnauthorizedCerts, maxRedirects 3, timeout 12s).
FETCH_TIMEOUT = 12
FETCH_MAX_REDIRECTS = 3
_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


# ═══════════════════════════════════════════════════════════════════
#   CATALOG + PROMPTS  (verbatim from the n8n «Build Claude Request» node)
# ═══════════════════════════════════════════════════════════════════

CATALOG = {
    "Impression": {"price": "$340", "link": "https://themes.shopify.com/themes/impression",
                   "pitch": "flagship premium theme, EU translations, cross-selling, mega menu, size chart, pre-order",
                   "presets": ["Etoile", "Felix", "Mimi", "Reflex"]},
    "Victory": {"price": "$320", "link": "https://themes.shopify.com/themes/victory",
                "pitch": "built for sports, events and active brands: store locator, event calendar, age verifier, countdowns",
                "presets": ["Athletica", "Flip", "Nitro", "Roast"]},
    "Boutique": {"price": "$160", "link": "https://themes.shopify.com/themes/boutique",
                 "pitch": "made for boutiques and premium brands, elegant product-first layouts",
                 "presets": ["Aurum", "Jade", "Noom", "Reflections"]},
    "Ultra": {"price": "$100", "link": "https://themes.shopify.com/themes/ultra",
              "pitch": "multi-purpose workhorse for tech, furniture, auto and toys",
              "presets": ["Grace", "Grip", "Harbor", "Sprout"]},
    "Allure": {"price": "$100", "link": "https://themes.shopify.com/themes/allure",
               "pitch": "versatile and affordable, great for beauty and lifestyle stores",
               "presets": ["Bijou", "Carrara", "Pristine", "Stitch"]},
    "Gain": {"price": "$100", "link": "https://themes.shopify.com/themes/gain",
             "pitch": "premium minimalist design that keeps the focus on products",
             "presets": ["Lace", "Maison", "Mio", "Sable"]},
}
# Every theme also ships 4 extra ready-made designs (presets), real names above.
# Preset demo URL pattern: <theme link>/presets/<preset-name-lowercase>


def _preset_lines(names):
    """Render the given themes' presets with their real demo URLs."""
    out = []
    for name in names:
        t = CATALOG[name]
        pres = ", ".join(
            p + " (" + t["link"] + "/presets/" + p.lower() + ")" for p in t.get("presets", []))
        out.append("- " + name + " presets: " + pres)
    return "\n".join(out)


def _tref(name):
    """How a theme must appear ANYWHERE it is named: Name ($price, link)."""
    t = CATALOG[name]
    return name + " (" + t["price"] + ", " + t["link"] + ")"


def _tline(name):
    """Full catalog line for the prompt: Name ($price, link): pitch."""
    t = CATALOG[name]
    return name + " (" + t["price"] + ", " + t["link"] + "): " + t["pitch"]


def _names_with_links(csv_names, fallback_names):
    """Render a comma-separated theme-name list (e.g. the sheet's Suggested
    Themes) with each theme's price + Theme Store link attached."""
    names = [s.strip() for s in (csv_names or "").split(",") if s.strip()]
    if not names:
        names = list(fallback_names)
    return ", ".join(_tref(n) if n in CATALOG else n for n in names)

MAP = {
    "Fashion & Apparel": ["Impression", "Boutique", "Gain", "Allure", "Victory"],
    "Jewelry & Accessories": ["Boutique", "Allure", "Gain", "Impression", "Ultra"],
    "Beauty & Cosmetics": ["Allure", "Boutique", "Gain", "Impression", "Victory"],
    "Sports & Fitness": ["Victory", "Ultra", "Impression", "Gain", "Allure"],
    "Food & Beverage": ["Victory", "Allure", "Gain", "Ultra", "Impression"],
    "Electronics & Tech": ["Ultra", "Impression", "Gain", "Victory", "Allure"],
    "Home & Furniture": ["Ultra", "Gain", "Allure", "Impression", "Boutique"],
    "Kids & Toys": ["Ultra", "Allure", "Victory", "Gain", "Impression"],
    "Pets": ["Ultra", "Allure", "Victory", "Gain", "Impression"],
    "Health & Supplements": ["Victory", "Gain", "Allure", "Ultra", "Impression"],
    "Art & Crafts": ["Allure", "Gain", "Boutique", "Ultra", "Impression"],
    "Auto & Moto": ["Ultra", "Victory", "Impression", "Gain", "Allure"],
}

REGISTRY = "https://themes.shopify.com/themes?q=UTD"
SITE = "https://utdweb.team"

FACTS = (
    "EVIDENCE ARSENAL. You may cite ONLY these studies and cases, always naming the "
    "source in plain words (no URLs needed for the studies). NEVER invent or "
    "estimate other numbers, and never claim a specific result for OUR themes: the "
    "cases prove the MECHANISM (faster, better-laid-out storefronts sell more), and "
    "our theme is how they get that mechanism.\n"
    "STUDIES:\n"
    "- Google/SOASTA (2017): mobile bounce probability rises 32% as load time goes "
    "from 1 to 3 seconds.\n"
    "- Deloitte and Google, 'Milliseconds Make Millions' (2020): a 0.1 second speed "
    "improvement lifted retail conversions by about 8.4% and average order value by "
    "over 9%.\n"
    "- Portent (2022): sites loading in 1 second convert about 2.5x better than "
    "sites loading in 5 seconds.\n"
    "- Baymard Institute: ~70% of carts are abandoned; their checkout-usability "
    "research puts the recoverable conversion gain for an average large store at "
    "about 35% from better checkout design alone.\n"
    "BUSINESS CASES (real, published by Google/web.dev):\n"
    "- Vodafone: made its landing page 31% faster and sales went up 8%.\n"
    "- Rakuten 24: invested in Core Web Vitals and got +33% conversion rate and "
    "+53% revenue per visitor.\n"
    "- Swappie: sped up its mobile site and grew mobile revenue by 42%.\n"
    "UTD VALUE FACTS (for the is-it-worth-the-money argument, NOT the main pain):\n"
    "- Upsell, cross-sell and promo blocks are built into the theme, which replaces "
    "several paid apps at $15-50/month each: the theme usually pays for itself "
    "within months.\n"
    "- One-time purchase (prices per theme below), updates included, listed on the "
    "official Shopify Theme Store, meaning Shopify reviewed the code before "
    "listing.\n"
    "HOW TO USE EVIDENCE: weave it INTO the argument as proof, at the exact moment "
    "you make the claim (\"stores that fixed exactly this saw...: Vodafone made its "
    "pages 31% faster and sales rose 8%\"). NEVER park a stat in its own detached "
    "paragraph, never dump several stats in a row. 1-2 pieces of evidence per "
    "email, each doing persuasion work.")

BASE = ("You are Sergey, a normal guy who works at UTD Web, an IT company that "
        "builds themes for the official Shopify Theme Store. You are writing an "
        "ordinary work email to the owner of a real Shopify store.\n"
        "\n"
        "GOAL: SELL the theme. Not present, not inform: convince this merchant "
        "that switching to ONE specific UTD theme will raise their sales, and "
        "back every claim with evidence. Be assertive: make the claim, prove "
        "it, tell them what to do next. The aggression lives in confidence, "
        "concreteness and proof, never in hype adjectives or pressure "
        "phrases. Everything happens by email: NEVER suggest a call, a "
        "meeting or a calendar link. It is fine to say 'if anything is "
        "unclear, just reply and I will help you set it up'.\n"
        "\n"
        "VOICE (the most important part): write like a normal person typing "
        "an email by hand. SIMPLE ENGLISH: many store owners are not native "
        "English speakers, so use common everyday words and short, simple "
        "sentences. No idioms, no slang, no fancy phrases (never 'caught my "
        "eye', 'worth a look', 'quick one'). If a 12-year-old would not "
        "understand a sentence, rewrite it. It must NOT read like AI text, "
        "a script, or ad copy. Maximum concreteness, zero water. Never use "
        "an em dash. Forbidden words: exclusive, exciting, game-changer, "
        "handpicked, curated, unique opportunity.\n"
        "\n"
        "LENGTH AND PARAGRAPHS (hard rules): the whole body stays SHORT, "
        "roughly 100-160 words. Paragraphs of 1-2 sentences, one idea each, "
        "blank line between. NEVER a big block of text: if a paragraph has "
        "3 or more sentences, split it.\n"
        "\n"
        "THE PAIN IS SALES. Speed, page performance and how products, "
        "upsells and checkout are laid out move the store's sales; that is "
        "the ground you argue on. Prove a claim with ONE study or business "
        "case woven into the same sentence (one short clause, source named), "
        "never a detached stats paragraph. App costs are only the "
        "pays-for-itself argument, never the pain.\n"
        "\n"
        "ONE MAIN THEME, AND THE RIGHT PRESET. Pick the single best-fit "
        "theme and sell THAT one: name + link + 1-2 concrete features tied "
        "to their store. Every theme ships 4 extra ready-made designs "
        "(presets); when a preset clearly matches the merchant's niche by "
        "its name and purpose (Roast for coffee, Athletica for sportswear, "
        "Bijou or Aurum for jewelry, Maison for home goods, Sprout for "
        "garden/eco, Nitro or Grip for auto and gear), point at that preset "
        "with its demo link instead of the generic theme demo. Never invent "
        "visual details of a preset; just say it is the ready-made design "
        "aimed at that kind of store and give the link. Other themes get "
        "ONE short line as alternatives, never equal billing.\n"
        "\n"
        "ENDING (hard rule): finish with an EASY, harmless way to continue "
        "the conversation. NEVER ask the merchant about their metrics, "
        "speed, conversion or anything they will not know off-hand, and no "
        "provocative questions. Good endings: offer to do something for "
        "them ('want me to check your store's speed and send you the "
        "numbers?', 'want me to point out which preset fits your catalog? "
        "just reply'), or an easy preference question ('which of the two "
        "looks closer to what you want?').\n"
        "\n"
        "FORMAT, mandatory for EVERY email including follow-up replies:\n"
        "- line 1: a greeting, 'Hi [store or person name] team,' or 'Hi "
        "there,' if there is no name\n"
        "- blank line\n"
        "- body split into short paragraphs by meaning, never one big mixed "
        "block\n"
        "- do NOT add a farewell or signature, it is appended separately.\n"
        "\n"
        "LINKS: every single time you name a theme, put its Theme Store link "
        "right next to it. Each theme page has a live preview/demo, so the "
        "merchant can click around a real working store. Other allowed links: " +
        REGISTRY + " (full UTD catalog) and " + SITE + ". No other links.\n"
        "\n" + FACTS + "\n"
        "\n"
        "Never invent features, prices or numbers; use only the facts given "
        "here. Never disparage their current theme. Write in English unless "
        "the merchant wrote to us in another language in this thread, then "
        "use their language.")

MODEL = "claude-sonnet-5"
MAX_TOKENS = 1300

# Signature appended to every body by the n8n «Parse Email» node.
SIGNATURE = "\n\nBest regards,\nSergey\nUTD Web | utdweb.team"


# ═══════════════════════════════════════════════════════════════════
#   LEAD SELECTION  (verbatim logic from the n8n «Pick Next» code node)
# ═══════════════════════════════════════════════════════════════════

# Junk / own-domain substrings that disqualify an email (n8n BAD list).
BAD = ['denvdavydov', 'smortkin', 'utdweb.team', 'utd.agency', '.png', '.jpg',
       '@sentry', 'your-company', 'example@', 'placeholder', '%20', 'noreply',
       'no-reply', 'support@', '@shopify.com']
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def _email_ok(e):
    return bool(e) and bool(EMAIL_RE.match(e)) and not any(b in e.lower() for b in BAD)


def _days_since(d):
    """Days since a 'YYYY-MM-DD HH:MM:SS' date string; 999 when unparseable
    (mirrors the n8n days() helper: Date.parse with ' '→'T')."""
    s = str(d or "").replace(" ", "T")
    try:
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
    except Exception:
        return 999
    return (datetime.now(timezone.utc) - dt).total_seconds() / 86400.0


def select_leads(rows):
    """Return the eligible leads, ordered/pooled as «Pick Next» does.

    Filter: valid on-domain email, Platform == 'Shopify', a Website, and a touch
    that is due by Status + Date-Sent age:
      Status ''        → touch 1
      'Sent'    & >=4d → touch 2
      'Followup1' & >=4d → touch 3
      'Followup2' & >=6d → touch 4
    Follow-ups (touch>1) are prioritised; the pool is then those follow-ups only
    (if any is due) else all touch-1 leads.
    """
    elig = []
    for i, r in enumerate(rows, start=2):  # sheet row (header is row 1)
        st = str(r.get("Status", "")).strip()
        email = str(r.get("Email", "")).strip()
        website = str(r.get("Website", "")).strip()
        if not _email_ok(email) or str(r.get("Platform", "")).strip() != "Shopify" or not website:
            continue
        ds = _days_since(r.get("Date Sent"))
        touch = 0
        if st == "":
            touch = 1
        elif st == "Sent" and ds >= 4:
            touch = 2
        elif st == "Followup1" and ds >= 4:
            touch = 3
        elif st == "Followup2" and ds >= 6:
            touch = 4
        if not touch:
            continue
        url = website if website.startswith("http") else "https://" + website
        elig.append({
            "found": True, "touch": touch, "email": email,
            "store_name": str(r.get("Store Name", "")).strip(),
            "website": url,
            "industry": str(r.get("Industry", "Other")).strip() or "Other",
            "current_theme": str(r.get("Current Theme", "")).strip(),
            "suggested": str(r.get("Suggested Themes", "")).strip(),
            "thread_id": str(r.get("Thread ID", "")).strip(),
            "last_msg": str(r.get("Last Msg ID", "")).strip(),
            "status": st,
            "date_sent": str(r.get("Date Sent", "")).strip(),
            "row_number": i,
        })
    if not elig:
        return []
    # follow-ups (touch>1) first, then new leads
    elig.sort(key=lambda e: 0 if e["touch"] > 1 else 1)
    if elig[0]["touch"] > 1:
        pool = [e for e in elig if e["touch"] > 1]
    else:
        pool = elig
    return pool


# ═══════════════════════════════════════════════════════════════════
#   STORE-SITE FETCH  (n8n «Fetch Store Website» + strip in «Build …»)
# ═══════════════════════════════════════════════════════════════════

def fetch_site_text(url):
    """Fetch the store homepage and reduce it to <=2000 chars of visible text.
    Matches the n8n strip: drop <script>/<style>, strip tags, collapse spaces.
    Only used for touch 1; returns '' on any failure (touch>1 never fetches)."""
    if requests is None or not url:
        return ""
    try:
        sess = requests.Session()
        sess.max_redirects = FETCH_MAX_REDIRECTS
        resp = sess.get(url, headers={"User-Agent": _UA}, timeout=FETCH_TIMEOUT,
                        allow_redirects=True, verify=False)
        html = resp.text or ""
    except Exception:
        return ""
    if not html or len(html) <= 80:
        return ""
    html = re.sub(r"<script[\s\S]*?</script>", "", html, flags=re.I)
    html = re.sub(r"<style[\s\S]*?</style>", "", html, flags=re.I)
    html = re.sub(r"<[^>]+>", " ", html)
    html = re.sub(r"\s+", " ", html).strip()
    return html[:2000]


# ═══════════════════════════════════════════════════════════════════
#   THREAD CONTEXT  (touches 2-4: feed Claude the ACTUAL prior emails)
# ═══════════════════════════════════════════════════════════════════

def _resolve_gm_thrid(account, thread_ref):
    """Turn the sheet's Thread ID value into a numeric Gmail thread id.

    The column may hold either a numeric X-GM-THRID (n8n era) or the RFC
    Message-ID of our first send (this sender writes sent_msg_id there).
    A Message-ID is resolved with ONE IMAP search in All Mail; '' on failure."""
    ref = (thread_ref or "").strip()
    if not ref:
        return ""
    if ref.isdigit():
        return ref
    if "@" not in ref:
        return ""
    mid = ref if ref.startswith("<") else "<%s>" % ref
    M = imaplib.IMAP4_SSL(ec.IMAP_HOST, ec.IMAP_PORT)
    try:
        M.login(account["user"], account["password"])
        for box in ("[Gmail]/All Mail", "[Google Mail]/All Mail"):
            typ, _ = M.select(box, readonly=True)
            if typ == "OK":
                break
        else:
            return ""
        typ, data = M.uid("SEARCH", None, "HEADER", "Message-ID", '"%s"' % mid)
        if typ != "OK" or not data or not data[0]:
            return ""
        uid = data[0].split()[0]
        typ, md = M.uid("FETCH", uid, "(X-GM-THRID)")
        if typ != "OK" or not md or not md[0]:
            return ""
        meta = md[0][0] if isinstance(md[0], tuple) else md[0]
        return ec._gm_thrid_from_fetch(meta)
    except Exception:
        return ""
    finally:
        try:
            M.close()
        except Exception:
            pass
        try:
            M.logout()
        except Exception:
            pass


def format_thread_history(thread, budget=1500):
    """Render a fetch_thread() list into a compact chronology for the model,
    trimmed to ~`budget` chars total (oldest first, like agency_autoresponder)."""
    if not thread:
        return ""
    per = max(160, budget // max(len(thread), 1) - 40)
    lines = []
    for m in thread:
        who = "UTD (us)" if m.get("direction") == "sent" else "Merchant"
        snippet = (m.get("snippet") or "").strip()[:per]
        lines.append("- [%s] %s: %s" % (str(m.get("date", ""))[:10], who, snippet))
    return "\n".join(lines)[:budget]


def summarize_prior_touches(c):
    """Structured fallback when the real thread cannot be fetched: reconstruct
    what each earlier touch said from the touch templates + the row state
    (Suggested Themes / Status / Date Sent)."""
    themes = c.get("suggested") or "the themes recommended for their industry"
    sums = {
        1: ("Touch 1 (first cold email): named the likely pain for their " +
            (c.get("industry") or "niche") + " store, introduced Sergey from "
            "UTD Web (official Shopify Theme Store developer), recommended all "
            "5 industry-fit themes with prices and Theme Store links (top 2-3 "
            "with reasons, the rest as alternatives): " + themes + ". Pointed "
            "them at the live demos on each theme page and ended with a "
            "question about what annoys them in their current setup."),
        2: ("Touch 2 (follow-up reply): reminded them of the first email, "
            "added one speed/conversion stat tied to their likely pain, "
            "nudged them to open the live demo of the best-fit theme (with "
            "its link), asked an easy question."),
        3: ("Touch 3 (value follow-up): did the cost math, built-in upsell/"
            "cross-sell/promo blocks replace several paid apps ($15-50/month "
            "each) while the theme is a one-time purchase with updates "
            "included, re-named the top themes with links: " + themes + "."),
    }
    lines = [sums[t] for t in range(1, c["touch"]) if t in sums]
    tail = "The merchant has not replied so far."
    if c.get("date_sent"):
        tail = ("Last email sent " + c["date_sent"] +
                " (CRM status '" + (c.get("status") or "") + "'). " + tail)
    lines.append(tail)
    return "\n".join(lines)


def build_thread_context(c):
    """Prior correspondence for touches 2-4. Prefers the ACTUAL thread over
    IMAP (needs the Gmail app-password + a Thread ID on the row, ~1500 chars);
    falls back to a structured summary of the earlier touches."""
    history = ""
    if ACCOUNT.get("user") and ACCOUNT.get("password") and c.get("thread_id"):
        try:
            thrid = _resolve_gm_thrid(ACCOUNT, c["thread_id"])
            if thrid:
                history = format_thread_history(
                    ec.fetch_thread(ACCOUNT, thrid, OWN_ADDRESSES))
        except Exception as e:
            print(f"  (thread fetch failed, using structured summary: {e})")
    if history:
        print("  [THREAD] using actual prior emails from IMAP "
              f"({len(history)} chars)")
    else:
        history = summarize_prior_touches(c)
        print("  [THREAD] IMAP unavailable/empty → structured summary of "
              "earlier touches")
    return history


# ═══════════════════════════════════════════════════════════════════
#   BUILD CLAUDE REQUEST  (per-touch prompts; 2-4 carry the prior thread)
# ═══════════════════════════════════════════════════════════════════

def build_request(c, site_text, history=""):
    """Return (system, user, primary, alt) for the contact's touch. Touches
    2-4 embed `history` (the actual prior emails, or the structured summary)
    so the letter builds on what was already said instead of repeating it."""
    order = MAP.get(c["industry"]) or ["Impression", "Ultra", "Allure", "Gain", "Victory"]
    primary, alt = order[:3], order[3:5]

    store = c["store_name"] or c["website"]
    touch = c["touch"]
    suggested_links = _names_with_links(c.get("suggested"), primary)

    if touch == 1:
        sys = (BASE +
               "\n\nThis is the FIRST email, a new thread. Paragraph plan (each "
               "its own SHORT paragraph, 1-3 sentences):\n"
               "1. Greeting line, blank line.\n"
               "2. THEM ONLY, detailed (2-3 simple sentences, may be two "
               "paragraphs): start with something like 'I was on your site "
               "and saw that it runs on Shopify.' Then describe what you "
               "actually saw there in plain words: what they sell, a "
               "collection or promo they run right now, how the store is "
               "set up. It must read like a person really browsed their "
               "store and is telling them what he noticed.\n"
               "3. US, one or two sentences: I'm Sergey from UTD Web, we make "
               "Shopify themes, they're on the official Theme Store (" +
               REGISTRY + ").\n"
               "4. THE PITCH, 2-3 sentences: the ONE best theme (or its "
               "best-matching preset) for this store, with link and price, "
               "1-2 features tied to their store, and one short woven proof "
               "clause from the EVIDENCE ARSENAL. Candidates:\n" +
               "\n".join("- " + _tline(n) for n in primary) + "\n" +
               _preset_lines(primary) + "\n"
               "5. Value line, 1 sentence: upsells/cross-sells/promo blocks "
               "are built in, that usually replaces $15-50/month of apps, so "
               "it pays for itself.\n"
               "6. Alternatives, 1 short sentence: " +
               ", ".join(_tref(n) for n in (primary[1:] + alt)) + ".\n"
               "7. Easy close per the ENDING rule (offer to do something for "
               "them, or an easy preference question; never ask about their "
               "numbers).\n"
               "SUBJECT: natural and properly capitalized, naming their niche."
               "\nOutput:\nSUBJECT: [subject]\nBODY:\n[body]")
        user = ("Store: " + store + "\nWebsite: " + c["website"] + "\nIndustry: " +
                c["industry"] + "\n\nSite content:\n" + (site_text or "(unavailable)") +
                "\n\nWrite the email.")
    elif touch == 2:
        sys = (BASE +
               "\n\nThis is the SECOND email, a reply in the SAME thread (they "
               "have not answered yet). You are shown the previous emails: build "
               "on them, never repeat the same pitch or reuse the same wording.\n"
               "Start with the greeting line, it is mandatory even in a reply. "
               "The FIRST sentence after the greeting reminds them of the "
               "previous email in plain words ('I emailed you last week about "
               "themes for your store, one more thing worth knowing').\n"
               "MAIN THEME LOCK (hard rule): the main theme of this email is "
               "THE SAME theme we recommended in the first email (the first "
               "name in 'Themes we suggested earlier'). Remind them of it by "
               "name: 'Last week I recommended [Theme] for your store.' "
               "NEVER switch the main recommendation to a different theme. "
               "If you mention a preset, it must be a preset OF THAT SAME "
               "theme.\n"
               "Then ONE new point in 1-2 simple sentences: a fresh reason "
               "this theme fits their store, with one short proof clause "
               "from the EVIDENCE ARSENAL woven into the sentence.\n"
               "Then one short paragraph: there are alternatives too, name "
               "1-2 other themes from our catalog with their links, one "
               "sentence only.\n"
               "Then one sentence pointing at the demo (what to open).\n"
               "Close per the ENDING rule: offer to do something for them "
               "or an easy preference question. NEVER ask about their "
               "metrics or speed.\n"
               "Total body 70-120 words, paragraphs of 1-2 sentences.\n"
               "SUBJECT: short, natural, properly capitalized (the reply keeps "
               "the thread subject, but output one anyway)."
               "\nOutput:\nSUBJECT: [subject]\nBODY:\n[body]")
        user = ("Store: " + store + "\nIndustry: " + c["industry"] +
                "\nThemes we suggested earlier: " + suggested_links +
                "\n\nPrevious emails in this thread:\n" + (history or "(unavailable)") +
                "\n\nWrite the follow-up. It must build on the thread above "
                "without repeating it.")
    elif touch == 3:
        sys = (BASE +
               "\n\nThis is the THIRD email, a reply in the same thread. You are "
               "shown the previous emails: take a DIFFERENT angle from both "
               "earlier notes, never repeat what was already said.\n"
               "Start with the greeting line, it is mandatory even in a reply. "
               "The FIRST sentence after the greeting must remind them that you "
               "wrote earlier about themes for their store. Calm, never pushy.\n"
               "MAIN THEME LOCK (hard rule): keep selling THE SAME theme "
               "from the earlier emails (the first name in the suggested "
               "list); remind them of it by name. Never switch the main "
               "recommendation.\n"
               "The angle of this email is money, said simply: the theme "
               "already includes what stores usually pay apps for (upsells, "
               "cross-sells, promo blocks, $15-50/month each), and it is a "
               "one-time price with updates included, so it pays for itself "
               "in a few months. Use the actual price of the main theme, "
               "keep its link. One short woven proof clause from the "
               "EVIDENCE ARSENAL max. Simple words, 1-2 sentences per "
               "paragraph.\n"
               "Close per the ENDING rule: friendly, decision-helping, "
               "email-only ('happy to help you set it up over email'), never "
               "a question about their numbers.\n"
               "Total body 60-110 words.\n"
               "SUBJECT: short, natural, properly capitalized."
               "\nOutput:\nSUBJECT: [subject]\nBODY:\n[body]")
        user = ("Store: " + store + "\nIndustry: " + c["industry"] + "\nTop themes: " +
                suggested_links +
                "\n\nPrevious emails in this thread:\n" + (history or "(unavailable)") +
                "\n\nWrite the value follow-up. It must add something new versus "
                "the thread above.")
    else:
        sys = (BASE +
               "\n\nThis is the FOURTH and final email, a calm breakup in the "
               "same thread. You are shown the previous emails: acknowledge the "
               "sequence naturally, do not re-pitch and do not repeat earlier "
               "lines.\n"
               "Start with the greeting line, it is mandatory even in a reply. "
               "The FIRST sentence after the greeting should reference that you "
               "have written a few times about themes for their store.\n"
               "Say plainly that you will stop emailing so you are not adding "
               "noise to their inbox, that there is no deadline on any of this, "
               "and leave the catalog link " + REGISTRY + " for whenever it "
               "fits. If one theme clearly fit them best, you may name it once, "
               "WITH its Theme Store link. No guilt, no pressure.\n"
               "Close by saying they can reply any time and you will help them "
               "figure things out or set the theme up.\n"
               "SUBJECT: short, natural, properly capitalized."
               "\nOutput:\nSUBJECT: [subject]\nBODY:\n[body]")
        user = ("Store: " + store +
                "\nTop themes we suggested: " + suggested_links +
                "\n\nPrevious emails in this thread:\n" + (history or "(unavailable)") +
                "\n\nWrite the breakup email.")

    return sys, user, primary, alt


# ═══════════════════════════════════════════════════════════════════
#   PARSE EMAIL  (verbatim SUBJECT/BODY parse + fallback templates)
# ═══════════════════════════════════════════════════════════════════

def parse_email(c, ai_text, primary, alt):
    """Parse Claude's SUBJECT/BODY output; fall back to the hard-coded per-touch
    template (verbatim from «Parse Email») when the model output is unusable.
    Appends the signature and returns the send + sheet payload dict."""
    text = ai_text or ""
    subject, body = "", ""
    if text:
        sm = re.search(r"SUBJECT:\s*(.+?)(?:\n|$)", text, re.I)
        bm = re.search(r"BODY:\s*([\s\S]+)", text, re.I)
        subject = sm.group(1).strip() if sm else ""
        body = bm.group(1).strip() if bm else ""

    if not subject or not body:
        p, a = primary, alt
        touch = c["touch"]
        if touch == 1:
            subject = "Shopify themes for your " + (
                c["industry"].lower() if c["industry"] != "Other" else "store") + " store"
            body = ("Hi " + (c["store_name"] or "there") + " team,\n\n"
                    "I came across your store and wanted to reach out. I am Sergey from "
                    "UTD Web, we build themes for the official Shopify Theme Store with "
                    "upsell, cross-sell and promo blocks built in, so they replace "
                    "several paid apps.\n\n"
                    "For a store in your niche I would look first at " + _tref(p[0]) +
                    " or " + _tref(p[1]) + ": " + CATALOG[p[0]]["pitch"] + ".\n\n"
                    "Also worth a look: " + _tref(p[2]) + ", " +
                    ", ".join(_tref(n) for n in a) + ". Every link opens the Theme "
                    "Store page with a live demo, so you can click through a real "
                    "store. Full catalog: " + REGISTRY + " and " + SITE + "\n\n"
                    "What annoys you most in your current setup? If anything is "
                    "unclear, just reply and I will help you figure it out.")
        elif touch == 2:
            subject = "Following up on themes for your store"
            body = ("Hi " + (c["store_name"] or "there") + " team,\n\n"
                    "I emailed you last week about Shopify themes for your store and "
                    "wanted to check back.\n\n"
                    "If you have a minute, open the live demo of " + _tref(p[0]) +
                    " and click through a product page. That is the quickest way to "
                    "see if it fits.\n\n"
                    "Did you get a chance to look, or is there anything I can answer?")
        elif touch == 3:
            subject = "The math on theme vs paid apps"
            body = ("Hi " + (c["store_name"] or "there") + " team,\n\n"
                    "I wrote earlier about themes for your store, one more thing worth "
                    "mentioning.\n\n"
                    "Our themes have upsell, cross-sell and promo blocks built in, "
                    "which replaces several paid apps (typical Shopify apps run "
                    "$15-50/month each). The theme itself is a one-time purchase and "
                    "includes updates. For your niche the best fits are " + _tref(p[0]) +
                    " and " + _tref(p[1]) + ", both pages have live demos.\n\n"
                    "Which of the two looks closer to what you want?")
        else:
            subject = "Last note from me"
            body = ("Hi " + (c["store_name"] or "there") + " team,\n\n"
                    "I have written a few times about themes for your store, so this "
                    "is my last note, I do not want to add noise to your inbox.\n\n"
                    "There is no deadline on any of this. If you ever want to explore "
                    "a new theme, the full catalog is here: " + REGISTRY + ".\n\n"
                    "Reply any time and I will help you set things up. Wishing you a "
                    "great season.")

    # Hard format guard: EVERY touch must open with a greeting line. If the
    # model (or a stray parse) dropped it, prepend a neutral one in code.
    if not re.match(r"^\s*(hi|hello|hey|dear|good\s(morning|afternoon|evening))\b",
                    body, re.I):
        body = "Hi there,\n\n" + body

    body += SIGNATURE
    next_status = ("Sent" if c["touch"] == 1 else
                   "Followup1" if c["touch"] == 2 else
                   "Followup2" if c["touch"] == 3 else "Sequence Done")
    suggested = ", ".join(primary + alt) if c["touch"] == 1 else c["suggested"]
    return {
        "email": c["email"], "subject": subject, "body": body, "touch": c["touch"],
        "next_status": next_status, "thread_id": c["thread_id"],
        "last_msg": c["last_msg"], "row_number": c["row_number"], "suggested": suggested,
    }


# ═══════════════════════════════════════════════════════════════════
#   SEND + SHEET  (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

# n8n «Update Sheet» writes these columns (matched by row_number).
SHEET_WRITE_COLS = ("Status", "Date Sent", "Thread ID", "Last Msg ID", "Suggested Themes")


def _print_prompt(system, user):
    """DRY_RUN: dump the FULL Claude prompt (system + user), clearly delimited,
    BEFORE the call — so a dry run shows exactly what the model will get."""
    print("\n" + "-" * 70)
    print("[CLAUDE PROMPT · DRY_RUN]")
    print("--- SYSTEM " + "-" * 59)
    print(system)
    print("--- USER " + "-" * 61)
    print(user)
    print("--- END PROMPT " + "-" * 55)


def _print_draft(payload, mode):
    print("\n" + "=" * 70)
    print(f"[DRAFT · touch {payload['touch']} · {mode}]  DRY_RUN — not sent")
    print(f"  from    : {SENDER_NAME} <{ACCOUNT['user']}>")
    print(f"  to      : {payload['email']}")
    print(f"  subject : {payload['subject']}")
    print(f"  status→ : {payload['next_status']}")
    print("  body:")
    for line in payload["body"].splitlines():
        print("    " + line)
    print("=" * 70)


def send_and_update(payload, sheet):
    """Send the email (new thread for touch 1, in-thread reply otherwise) and
    queue the CRM row update. Returns the sent Message-ID (or '')."""
    is_new = payload["touch"] == 1
    mode = "Send New" if is_new else "Send Reply"

    if DRY_RUN:
        _print_draft(payload, mode)
        _queue_sheet_write(sheet, payload, sent_msg_id="")
        return ""

    if is_new:
        sent_msg_id = ec.send_email(ACCOUNT, payload["email"], payload["subject"],
                                    payload["body"])
    else:
        # Reply in the same thread: In-Reply-To / References carry the chain.
        sent_msg_id = ec.send_email(
            ACCOUNT, payload["email"], payload["subject"], payload["body"],
            in_reply_to=payload["last_msg"] or None,
            references=payload["thread_id"] or payload["last_msg"] or None)
    _queue_sheet_write(sheet, payload, sent_msg_id=sent_msg_id)
    return sent_msg_id


def _queue_sheet_write(sheet, payload, sent_msg_id):
    """Accumulate the A1 cell updates for this contact's row (flushed in one
    batch call). Mirrors «Update Sheet»: Status / Date Sent / Thread ID /
    Last Msg ID / Suggested Themes matched by row_number."""
    header = sheet["header"]
    row = payload["row_number"]
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    values = {
        "Status": payload["next_status"],
        "Date Sent": now,
        "Thread ID": payload["thread_id"] or sent_msg_id,
        "Last Msg ID": sent_msg_id or payload["last_msg"],
        "Suggested Themes": payload["suggested"],
    }
    for col in SHEET_WRITE_COLS:
        if col not in header:
            continue
        a1 = ec.gspread_a1(row, header.index(col) + 1)
        sheet["cell_updates"].append({"range": a1, "values": [[values[col]]]})
    print(f"  [SHEET] queued row {row}: Status='{payload['next_status']}', "
          f"Suggested Themes='{payload['suggested']}'")


def flush_sheet(sheet):
    """Write ALL queued cell updates in a single Sheets batchUpdate call."""
    updates = sheet["cell_updates"]
    if not updates:
        print("\n[SHEET] no updates to flush.")
        return
    if DRY_RUN:
        print(f"\n[SHEET] DRY_RUN — would batch-write {len(updates)} cells in ONE call:")
        for u in updates:
            print(f"    {u['range']} = {u['values'][0][0]!r}")
        return
    if not sheet["ws"]:
        print("\n[SHEET] no worksheet handle — cannot flush.")
        return
    try:
        n = ec.batch_update_cells(sheet["ws"], updates)
        print(f"\n[SHEET] batch-wrote {n} cells in ONE call.")
    except Exception as e:
        print(f"\n⚠️  [SHEET] batch update failed after retries: {e}")


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def run_once():
    print(f"=== UTD ecom cold-outreach sender | DRY_RUN={DRY_RUN} | "
          f"limit {LEAD_LIMIT} | {datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)
    sheet = {"ws": None, "header": [], "cell_updates": []}

    # CRM snapshot: open + read ONCE (429 backoff inside email_common).
    rows = []
    try:
        ws = ec.open_worksheet(SHEET_ID, SHEET_TAB)
        rows = ec.read_rows_ws(ws)
        sheet["ws"] = ws
        sheet["header"] = list(rows[0].keys()) if rows else ws.row_values(1)
    except Exception as e:
        print(f"⚠️  Could not read CRM sheet: {e}")
        return {"parser": "ecom_sender", "dry_run": DRY_RUN, "error": str(e),
                "sent": 0, "selected": 0}

    pool = select_leads(rows)
    print(f"CRM: {len(rows)} rows read, {len(pool)} due in the winning tier.")
    if not pool:
        print("Nothing due — stop.")
        return {"parser": "ecom_sender", "dry_run": DRY_RUN, "eligible": 0,
                "selected": 0, "sent": 0}

    # «Pick Next» picks one at random; LEAD_LIMIT lets a GHA run send a few.
    random.shuffle(pool)
    stats = {}
    selected = 0
    sent = 0
    for c in pool:
        if selected >= LEAD_LIMIT:
            break
        # Dedup: this exact touch to this contact was already sent in a prior run.
        dedup_id = f"ecom:{c['email'].lower()}:touch{c['touch']}"
        if ec.is_processed(state, dedup_id):
            continue
        selected += 1

        site_text = fetch_site_text(c["website"]) if c["touch"] == 1 else ""
        # Touches 2-4: pull the ACTUAL prior correspondence (or a structured
        # summary) so the letter builds on what was already said.
        history = build_thread_context(c) if c["touch"] > 1 else ""
        sys, user, primary, alt = build_request(c, site_text, history)
        if DRY_RUN:
            _print_prompt(sys, user)
        ai_text = ec.call_claude(sys, user, model=MODEL, max_tokens=MAX_TOKENS)
        if not ai_text:
            print("[claude unavailable -> fallback/skip]")
        payload = parse_email(c, ai_text, primary, alt)

        print(f"\n· touch {payload['touch']} → {payload['email']} | "
              f"{c['store_name'] or c['website']} [{c['industry']}] "
              f"| status→{payload['next_status']}")
        send_and_update(payload, sheet)
        stats[payload["next_status"]] = stats.get(payload["next_status"], 0) + 1
        if not DRY_RUN:
            ec.mark_processed(state, dedup_id)
            sent += 1

    flush_sheet(sheet)
    if not DRY_RUN:
        ec.save_state(STATE_FILE, state)

    print(f"\n=== done. selected={selected} sent={sent} statuses={stats} ===")
    return {"parser": "ecom_sender", "dry_run": DRY_RUN,
            "eligible": len(pool), "selected": selected, "sent": sent,
            "statuses": stats, "sheet_cells_queued": len(sheet["cell_updates"])}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
