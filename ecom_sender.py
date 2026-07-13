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
    "password": os.environ.get("SENDER_APP_PW") or os.environ.get("GMAIL_APP_PW_SERGEY", ""),
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

# ONE specialized preset per industry (theme, preset) — the preset IS the product
# we sell first; names are the real Theme Store presets.
# ═══════════════════════════════════════════════════════════════════
#   PRESET REGISTRY — каждый пресет: ОДНА основная индустрия + ДВЕ
#   второстепенные (правится в одном месте, здесь).
#   t1 продаёт пресет, чья ОСНОВНАЯ индустрия = реальная индустрия магазина
#   (по контенту сайта). t2 напоминает главный пресет и добавляет ДВА пресета,
#   у которых индустрия магазина стоит во ВТОРОСТЕПЕННЫХ.
# ═══════════════════════════════════════════════════════════════════

PRESETS = {
    # Victory ($320)
    "Victory":   {"theme": "Victory",    "primary": "Food & Beverage",       "secondary": ["Sports & Fitness", "Health & Supplements"]},
    "Athletica": {"theme": "Victory",    "primary": "Sports & Fitness",      "secondary": ["Health & Supplements", "Fashion & Apparel"]},
    "Nitro":     {"theme": "Victory",    "primary": "Auto & Moto",           "secondary": ["Electronics & Tech", "Sports & Fitness"]},
    "Roast":     {"theme": "Victory",    "primary": "Food & Beverage",       "secondary": ["Health & Supplements", "Art & Crafts"]},
    "Flip":      {"theme": "Victory",    "primary": "Sports & Fitness",      "secondary": ["Kids & Toys", "Auto & Moto"]},
    # Ultra ($100)
    "Ultra":     {"theme": "Ultra",      "primary": "Electronics & Tech",    "secondary": ["Auto & Moto", "Kids & Toys"]},
    "Grip":      {"theme": "Ultra",      "primary": "Sports & Fitness",      "secondary": ["Auto & Moto", "Electronics & Tech"]},
    "Harbor":    {"theme": "Ultra",      "primary": "Home & Furniture",      "secondary": ["Art & Crafts", "Electronics & Tech"]},
    "Sprout":    {"theme": "Ultra",      "primary": "Pets",                  "secondary": ["Health & Supplements", "Kids & Toys"]},
    "Grace":     {"theme": "Ultra",      "primary": "Beauty & Cosmetics",    "secondary": ["Fashion & Apparel", "Jewelry & Accessories"]},
    # Gain ($100)
    "Gain":      {"theme": "Gain",       "primary": "Fashion & Apparel",     "secondary": ["Beauty & Cosmetics", "Electronics & Tech"]},
    "Lace":      {"theme": "Gain",       "primary": "Fashion & Apparel",     "secondary": ["Jewelry & Accessories", "Beauty & Cosmetics"]},
    "Maison":    {"theme": "Gain",       "primary": "Home & Furniture",      "secondary": ["Art & Crafts", "Kids & Toys"]},
    "Mio":       {"theme": "Gain",       "primary": "Kids & Toys",           "secondary": ["Pets", "Fashion & Apparel"]},
    "Sable":     {"theme": "Gain",       "primary": "Beauty & Cosmetics",    "secondary": ["Fashion & Apparel", "Health & Supplements"]},
    # Allure ($100)
    "Allure":    {"theme": "Allure",     "primary": "Beauty & Cosmetics",    "secondary": ["Fashion & Apparel", "Health & Supplements"]},
    "Bijou":     {"theme": "Allure",     "primary": "Jewelry & Accessories", "secondary": ["Fashion & Apparel", "Beauty & Cosmetics"]},
    "Carrara":   {"theme": "Allure",     "primary": "Art & Crafts",          "secondary": ["Home & Furniture", "Jewelry & Accessories"]},
    "Pristine":  {"theme": "Allure",     "primary": "Beauty & Cosmetics",    "secondary": ["Health & Supplements", "Fashion & Apparel"]},
    "Stitch":    {"theme": "Allure",     "primary": "Fashion & Apparel",     "secondary": ["Art & Crafts", "Kids & Toys"]},
    # Boutique ($160)
    "Boutique":  {"theme": "Boutique",   "primary": "Fashion & Apparel",     "secondary": ["Jewelry & Accessories", "Beauty & Cosmetics"]},
    "Aurum":     {"theme": "Boutique",   "primary": "Jewelry & Accessories", "secondary": ["Fashion & Apparel", "Beauty & Cosmetics"]},
    "Jade":      {"theme": "Boutique",   "primary": "Beauty & Cosmetics",    "secondary": ["Jewelry & Accessories", "Health & Supplements"]},
    "Noom":      {"theme": "Boutique",   "primary": "Home & Furniture",      "secondary": ["Electronics & Tech", "Art & Crafts"]},
    "Reflections":{"theme": "Boutique",  "primary": "Jewelry & Accessories", "secondary": ["Beauty & Cosmetics", "Fashion & Apparel"]},
    # Impression ($340)
    "Impression":{"theme": "Impression", "primary": "Fashion & Apparel",     "secondary": ["Beauty & Cosmetics", "Home & Furniture"]},
    "Etoile":    {"theme": "Impression", "primary": "Fashion & Apparel",     "secondary": ["Jewelry & Accessories", "Beauty & Cosmetics"]},
    "Felix":     {"theme": "Impression", "primary": "Pets",                  "secondary": ["Kids & Toys", "Food & Beverage"]},
    "Mimi":      {"theme": "Impression", "primary": "Kids & Toys",           "secondary": ["Fashion & Apparel", "Pets"]},
    "Reflex":    {"theme": "Impression", "primary": "Electronics & Tech",    "secondary": ["Auto & Moto", "Sports & Fitness"]},
}


def preset_url(preset):
    t = CATALOG[PRESETS[preset]["theme"]]
    return t["link"] + "/presets/" + preset.lower()


def preset_ref(preset):
    """How a preset appears anywhere. Base presets (name == theme) read as the
    theme itself; named presets read as 'Name design of Theme'."""
    th = PRESETS[preset]["theme"]
    if preset == th:
        return "the " + th + " theme (" + CATALOG[th]["price"] + ", " + CATALOG[th]["link"] + ")"
    return (preset + " design of " + th + " (" + CATALOG[th]["price"] + ", " +
            preset_url(preset) + ")")


def primary_preset(industry):
    """The ONE preset whose PRIMARY industry matches (first match wins)."""
    for pr, info in PRESETS.items():
        if info["primary"] == industry:
            return pr
    return "Ultra"


def secondary_presets(industry, exclude=()):
    """TWO presets that carry this industry in their SECONDARY industries."""
    out = []
    for pr, info in PRESETS.items():
        if pr in exclude or PRESETS[pr]["theme"] in [PRESETS.get(e, {}).get("theme") for e in exclude]:
            continue
        if industry in info["secondary"]:
            out.append(pr)
        if len(out) == 2:
            break
    return out


def registry_table():
    """Whole registry rendered for the t1 prompt: Claude picks by REAL industry."""
    lines = []
    for pr, info in PRESETS.items():
        lines.append("- " + preset_ref(pr) + " | main industry: " + info["primary"] +
                     " | also fits: " + ", ".join(info["secondary"]))
    return "\n".join(lines)


def main_pick(industry):
    """(theme, preset, preset_url) by industry PRIMARY match (deterministic)."""
    pr = primary_preset(industry)
    return PRESETS[pr]["theme"], pr, preset_url(pr)


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
        "ONE PRESET IS THE PRODUCT. The user message names THE preset we "
        "sell to this store (each industry has its specialized ready-made "
        "design). SELL THAT PRESET as the product: its name, its demo link, "
        "the parent theme's price and 1-2 features tied to their store. "
        "Say plainly this is the design made for stores like theirs and "
        "the one you would install. The parent theme is mentioned once as "
        "context ('the [Preset] design of our [Theme] theme'). Never sell "
        "the bare theme when a preset is given, never name presets of "
        "other themes, never invent visual details. Other themes get ONE "
        "short line as alternatives at the end.\n"
        "\n"
        "ENDING (hard rule): finish with an EASY, harmless way to continue "
        "the conversation. NEVER ask the merchant about their metrics or "
        "anything they will not know off-hand. NEVER offer to test their "
        "site speed (if we have speed numbers, they are already in the "
        "letter). Good endings: offer something concrete we will do for "
        "them by email ('want me to send a short list of what would change "
        "on your store with this design? just reply'), an easy "
        "preference question ('which look is closer to what you want?'), or "
        "an offer to advise if they have questions about which theme fits "
        "their store (they can also just pick one themselves on the Theme "
        "Store).\n"
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
MAX_TOKENS = 2000

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
    if not e or not EMAIL_RE.match(e) or any(b in e.lower() for b in BAD):
        return False
    dom = e.split("@")[1]
    # junk guard: real-looking domain + valid-length TLD (e.g. 7@g.ebe)
    return len(e.split("@")[0]) >= 2 and len(dom) >= 6 and len(dom.split(".")[-1]) >= 2


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
    # Owner override: rows with Status='NEXT' are sent FIRST as touch 1,
    # bypassing the blocklists (deliberate manual insert, e.g. internal test).
    for i, r in enumerate(rows, start=2):
        if str(r.get("Status", "")).strip() == "NEXT" and "@" in str(r.get("Email", "")):
            website = str(r.get("Website", "")).strip()
            url = website if website.startswith("http") else "https://" + website
            return [{
                "found": True, "touch": 1,
                "email": str(r.get("Email", "")).strip(),
                "store_name": str(r.get("Store Name", "")).strip(),
                "website": url,
                "industry": str(r.get("Industry", "Other")).strip() or "Other",
                "current_theme": str(r.get("Current Theme", "")).strip(),
                "suggested": str(r.get("Suggested Themes", "")).strip(),
                "thread_id": str(r.get("Thread ID", "")).strip(),
                "last_msg": str(r.get("Last Msg ID", "")).strip(),
                "status": "", "date_sent": "",
                "row_number": i,
            }]
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


def measure_pagespeed(url):
    """Real Google PageSpeed (mobile) for the store: (score/100, lcp_seconds).
    Returns (None, None) on any failure — the letter then must NOT claim we
    tested anything."""
    if not requests or not url:
        return None, None
    try:
        api = ("https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
               "?strategy=mobile&category=performance&url=" + url)
        # A real key lifts the strict keyless quota (which returns 429). Prefer
        # a dedicated PAGESPEED_API_KEY; otherwise reuse the first Google API
        # key we already carry — PageSpeed is the same key type and the API is
        # now enabled on those projects.
        key = os.environ.get("PAGESPEED_API_KEY", "").strip()
        if not key:
            key = (os.environ.get("GOOGLE_API_KEYS", "").split(",") or [""])[0].strip()
        if key:
            api += "&key=" + key
        d = None
        for _attempt in range(3):
            r = requests.get(api, timeout=70)
            if r.status_code == 200:
                d = r.json(); break
            if r.status_code == 429:  # keyless daily quota; a key fixes this
                import time as _t; _t.sleep(2 * (_attempt + 1)); continue
            break
        if d is None:
            return None, None
        lh = d.get("lighthouseResult", {})
        score = lh.get("categories", {}).get("performance", {}).get("score")
        lcp = lh.get("audits", {}).get("largest-contentful-paint", {}).get("numericValue")
        return (round(score * 100) if score is not None else None,
                round(lcp / 1000.0, 1) if lcp else None)
    except Exception:
        return None, None

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


def followup_main(c):
    """Main preset for touches 2-4: the preset RECORDED at touch 1 (in the
    sheet's Suggested Themes) if present, else the registry primary match."""
    for token in re.split(r"[,@|]", str(c.get("suggested") or "")):
        token = token.strip()
        if token in PRESETS:
            return token
    return primary_preset(c["industry"])


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
               "2. THEM, detailed (2-4 simple sentences): start with 'I was "
               "on your site and saw that it runs on Shopify.' Then what you "
               "actually saw (products, promo, how the store is set up). "
               "CRITICAL: never stop at describing what they sell, they know "
               "their own store. The LAST sentence of this paragraph must be "
               "an INSIGHT they do not already know: what in their current "
               "setup is likely costing them orders or what opportunity "
               "their catalog has (grounded in what you saw, no invented "
               "numbers). This sentence is the reason the whole email "
               "exists.\n"
               "3. THE BRIDGE + WHY LISTEN, 2 sentences: connect directly to "
               "the insight ('That is exactly why I am writing to you.') and "
               "give them a reason to keep reading: I'm Sergey from UTD Web, "
               "our themes are on Shopify's official Theme Store (" +
               REGISTRY + "), Shopify reviews every theme before listing "
               "and thousands of stores run on them.\n"
               "4. THE PITCH, 2-3 sentences. FIRST decide the store's REAL "
               "industry from the site content (products they actually "
               "sell). The CRM industry field is only a hint and is "
               "sometimes wrong: a store selling vehicle awnings and car "
               "gear is Auto & Moto even if the CRM says Sports. Then pick "
               "from the PRESET REGISTRY in the user message THE preset "
               "whose MAIN industry matches the store's real industry, and "
               "sell it: preset name + demo link + parent theme price, 1-2 "
               "features tied to their store, one short woven proof clause "
               "from the EVIDENCE ARSENAL. If several presets share that "
               "main industry, take the one listed FIRST in the registry. "
               "NEVER name any other preset anywhere in the letter: the "
               "chosen preset is the only preset word allowed; alternatives "
               "are THEMES only. If the user message contains REAL "
               "PageSpeed numbers, state them plainly as something we "
               "measured; if not, do NOT claim we tested anything.\n"
               "5. Value line, 1 sentence: upsells/cross-sells/promo blocks "
               "are built in, that usually replaces $15-50/month of apps, so "
               "it pays for itself.\n"
               "6. CLOSING PARAGRAPH (mandatory, always include it, 2-3 short "
               "sentences). First mention our OTHER themes as options they can "
               "also look at: " +
               ", ".join(_tref(n) for n in (primary[1:] + alt)) + ". Then give "
               "them BOTH easy paths, no pressure: they can reply and we will "
               "talk it through, or they can just pick a theme themselves on "
               "the official Theme Store. End by offering that we are happy to "
               "advise if they have any questions about which theme fits their "
               "store. Warm and low-pressure; never ask about their numbers.\n"
               "SUBJECT: natural and properly capitalized, naming their niche."
               "\nOutput (the PRESET line is mandatory, it is machine-read):"
               "\nPRESET: [exact preset name you chose from the registry]"
               "\nSUBJECT: [subject]\nBODY:\n[body]")
        speed = ""
        if c.get("psi_score") is not None:
            speed = ("\nREAL PageSpeed numbers we measured for their site "
                     "(state them in the letter): mobile performance score " +
                     str(c["psi_score"]) + " of 100" +
                     (", largest content loads in " + str(c["psi_lcp"]) + "s"
                      if c.get("psi_lcp") else "") + ".")
        user = ("Store: " + store + "\nWebsite: " + c["website"] +
                "\nCRM industry (hint only, may be wrong): " + c["industry"] +
                "\n\nPRESET REGISTRY (pick ONE whose main industry = the store's "
                "REAL industry):\n" + registry_table() +
                speed +
                "\n\nSite content:\n" + (site_text or "(unavailable)") +
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
               "PRESET LOCK (hard rule): we sell ONE product to this store "
               "across the whole sequence: the preset named in the user "
               "message as 'MAIN preset'. Remind them of it by name: 'Last "
               "week I recommended the [Preset] design for your store.' "
               "Keep selling THAT preset with its demo link. NEVER switch "
               "to another theme or preset, never attach the preset to a "
               "different theme. If earlier emails named something else, "
               "still sell the MAIN preset given now.\n"
               "Then ONE new point in 1-2 simple sentences: a fresh reason "
               "this theme fits their store, with one short proof clause "
               "from the EVIDENCE ARSENAL woven into the sentence.\n"
               "Then one short paragraph: offer the TWO additional designs "
               "named in the user message as options that also fit their "
               "industry, each with its demo link, one sentence only.\n"
               "Then one sentence pointing at the demo (what to open).\n"
               "Close per the ENDING rule: offer to do something for them "
               "or an easy preference question. NEVER ask about their "
               "metrics or speed.\n"
               "Total body 70-120 words, paragraphs of 1-2 sentences.\n"
               "SUBJECT: short, natural, properly capitalized (the reply keeps "
               "the thread subject, but output one anyway)."
               "\nOutput:\nSUBJECT: [subject]\nBODY:\n[body]")
        m_preset = followup_main(c)
        extras = secondary_presets(c["industry"], exclude=(m_preset,))
        user = ("Store: " + store + "\nIndustry: " + c["industry"] +
                "\nMAIN preset (remind them of it, keep selling it): " +
                preset_ref(m_preset) +
                "\nADDITIONAL options (exactly these two, each with its link, "
                "offered as designs that also fit their industry): " +
                "; ".join(preset_ref(x) for x in extras) +
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
               "PRESET LOCK (hard rule): keep selling the ONE preset named "
               "in the user message as 'MAIN preset'; remind them of it by "
               "name with its demo link. Never switch to another theme or "
               "preset.\n"
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
        m_preset = followup_main(c)
        user = ("Store: " + store + "\nIndustry: " + c["industry"] +
                "\nMAIN preset (remind them of it, keep selling it): " +
                preset_ref(m_preset) +
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

ALL_PRESET_NAMES = {pr: th for th, t in CATALOG.items() for pr in t.get("presets", [])}


def validate_body(c, body, allowed_presets):
    """Code-level guard (not a prompt): every link must be one of ours, and any
    preset named in the body must be in allowed_presets AND carry its own
    theme's URL. Returns an error string or None."""
    allowed = {REGISTRY, SITE} | {t["link"] for t in CATALOG.values()} | {
        t["link"] + "/presets/" + pr.lower()
        for th, t in CATALOG.items() for pr in t.get("presets", [])}
    for link in re.findall(r"https?://[^\s)>\]]+", body):
        if link.rstrip('.,;:') not in allowed:
            return "disallowed link: " + link
    for pr, th in ALL_PRESET_NAMES.items():
        if pr.lower() == th.lower():
            continue  # base-name presets equal the theme name
        if re.search(r"\b" + re.escape(pr) + r"\b", body):
            if pr not in allowed_presets:
                return "preset not allowed here: " + pr
            # if its URL appears, it must be the CORRECT theme's preset URL
            wrong = [u for u in re.findall(r"https?://themes\.shopify\.com/themes/[a-z]+/presets/" + pr.lower(), body)
                     if u != preset_url(pr)]
            if wrong:
                return "preset %s linked to wrong theme: %s" % (pr, wrong[0])
    return None


def parse_email(c, ai_text, primary, alt):
    """Parse Claude's SUBJECT/BODY output; fall back to the hard-coded per-touch
    template (verbatim from «Parse Email») when the model output is unusable.
    Appends the signature and returns the send + sheet payload dict."""
    text = (ai_text or "").strip()
    subject, body, chosen_preset = "", "", ""
    if text:
        pm = re.search(r"PRESET:\s*([A-Za-z]+)", text)
        if pm and pm.group(1) in PRESETS:
            chosen_preset = pm.group(1)
        # strip the machine-only PRESET line so it never leaks into the body
        text = re.sub(r"^\s*PRESET:.*$", "", text, flags=re.I | re.M).strip()
        sm = re.search(r"SUBJECT:\s*(.+?)(?:\n|$)", text, re.I)
        bm = re.search(r"BODY:\s*([\s\S]+)", text, re.I)
        if sm and bm:
            subject = sm.group(1).strip()
            body = bm.group(1).strip()
        else:
            # TOLERANT: Claude wrote a real letter without the rigid markers.
            # NEVER discard good Claude text for the wooden fallback — use it.
            # Derive the subject from a leading "Subject:"-ish line if present,
            # otherwise from the greeting/niche; the rest is the body.
            lines = text.splitlines()
            if sm:  # only SUBJECT: present
                subject = sm.group(1).strip()
                body = re.sub(r"^\s*SUBJECT:.*$", "", text, flags=re.I | re.M).strip()
            elif lines and re.match(r"(?i)^\s*(hi|hello|hey)\b", lines[0]):
                body = text
                subject = ("Shopify themes for your "
                           + (c["industry"].lower() if c["industry"] != "Other" else "store")
                           + " store")
            elif len(text) > 60:  # a body without a greeting -> add one, keep text
                body = "Hi " + (c["store_name"] or "there") + " team,\n\n" + text
                subject = ("Shopify themes for your "
                           + (c["industry"].lower() if c["industry"] != "Other" else "store")
                           + " store")

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

    # Code-level link/preset guard: a letter with a foreign link, a preset of
    # the wrong theme, or the wrong preset NEVER goes out. AI draft -> fallback.
    if c["touch"] == 1:
        main_pr = chosen_preset or primary_preset(c["industry"])
        allowed_pr = {main_pr}
    else:
        main_pr = followup_main(c)
        allowed_pr = {main_pr} | set(secondary_presets(c["industry"], exclude=(main_pr,)))
    err = validate_body(c, body, allowed_pr)
    if err and ai_text:
        print("  [GUARD] AI draft rejected (" + err + ") -> safe fallback template")
        return parse_email(c, "", primary, alt)

    body += SIGNATURE
    next_status = ("Sent" if c["touch"] == 1 else
                   "Followup1" if c["touch"] == 2 else
                   "Followup2" if c["touch"] == 3 else "Sequence Done")
    if c["touch"] == 1:
        suggested = (main_pr + " @ " + PRESETS[main_pr]["theme"] + ", " +
                     ", ".join(n for n in (primary + alt) if n != PRESETS[main_pr]["theme"]))
    else:
        suggested = c["suggested"]
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
        if c["touch"] == 1:
            c["psi_score"], c["psi_lcp"] = measure_pagespeed(c["website"])
            if c["psi_score"] is not None:
                print(f"  [PSI] mobile score={c['psi_score']}/100 lcp={c.get('psi_lcp')}s")
            else:
                print("  [PSI] unavailable -> letter must not claim we tested")
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
