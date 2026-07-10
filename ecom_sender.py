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
    "Impression": {"price": "$340", "pitch": "flagship premium theme, EU translations, cross-selling, mega menu, size chart, pre-order"},
    "Victory": {"price": "$320", "pitch": "built for sports, events and active brands: store locator, event calendar, age verifier, countdowns"},
    "Boutique": {"price": "$160", "pitch": "made for boutiques and premium brands, elegant product-first layouts"},
    "Ultra": {"price": "$100", "pitch": "multi-purpose workhorse for tech, furniture, auto and toys"},
    "Allure": {"price": "$100", "pitch": "versatile and affordable, great for beauty and lifestyle stores"},
    "Gain": {"price": "$100", "pitch": "premium minimalist design that keeps the focus on products"},
}

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

BASE = ("You are Sergey from UTD Web, an official Shopify Theme Store developer "
        "(6 theme families, built-in conversion features: upsells, cross-sells, "
        "promo sections). You write to the owner of a real Shopify store. "
        "GOAL: move the merchant toward choosing and buying the UTD theme that "
        "fits their store. Concrete benefits to lean on: conversion features "
        "built in (fewer paid apps), fast themes, official Shopify Theme Store "
        "status (safe purchase, preview before publish, products stay in place). "
        "STYLE: human and specific. Hook in the first 3 words. Lead with a "
        "concrete observation or fact, never a generic compliment. No hype, no "
        "corporate slop. Never use an em dash. Forbidden words: exclusive, "
        "exciting, game-changer, handpicked, curated, unique opportunity. "
        "Never invent features, prices or numbers; use only the facts given "
        "here. Never disparage their current theme. The only links allowed: " +
        SITE + " and " + REGISTRY + ". Write in English unless the merchant "
        "wrote to us in another language in this thread, then use their "
        "language. NO signature (it is added separately).")

MODEL = "claude-sonnet-5"
MAX_TOKENS = 900

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
        1: ("Touch 1 (first cold email): one specific observation about their "
            "store, introduced UTD Web (official Shopify Theme Store developer, "
            "conversion features built in), recommended these themes for their " +
            (c.get("industry") or "niche") + " niche with prices: " + themes +
            ". Shared the catalog links and offered questions or a demo store."),
        2: ("Touch 2 (short bump in the same thread): lightly referenced the "
            "first note, offered to send a live demo store of the best-fit "
            "theme, asked if they had questions."),
        3: ("Touch 3 (value follow-up): explained UTD themes ship with upsells, "
            "cross-sells and promo sections built in (usually fewer paid apps), "
            "switching is low-risk (preview before publish, products stay in "
            "place), re-named the top themes: " + themes + "."),
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

    def tline(n):
        return n + " (" + CATALOG[n]["price"] + "): " + CATALOG[n]["pitch"]

    store = c["store_name"] or c["website"]
    touch = c["touch"]

    if touch == 1:
        sys = (BASE +
               "\n\nFIRST cold email. Structure: (1) one SPECIFIC observation about "
               "their store from the site content; (2) one sentence who we are; (3) "
               "confidently recommend 2-3 themes for their industry, each a concrete "
               "reason tied to their store: " + " | ".join(tline(n) for n in primary) +
               "; (4) one line naming the remaining alternatives so 5 total: " +
               ", ".join(n + " (" + CATALOG[n]["price"] + ")" for n in alt) +
               "; (5) links to full catalog: " + REGISTRY + " and " + SITE +
               "; (6) soft close offering questions or a demo store. Under 160 words."
               "\nOutput:\nSUBJECT: [subject naming their niche]\nBODY:\n[body]")
        user = ("Store: " + store + "\nWebsite: " + c["website"] + "\nIndustry: " +
                c["industry"] + "\n\nSite content:\n" + (site_text or "(unavailable)") +
                "\n\nWrite the email.")
    elif touch == 2:
        sys = (BASE +
               "\n\nSECOND email, a short follow-up in the SAME thread (no reply yet). "
               "60-90 words, hard cap 120. You are shown the previous emails in this "
               "thread: BUILD on them, never repeat the same pitch or reuse the same "
               "wording. Reference the first note lightly. Offer to send a live demo "
               "store of the best-fit theme for their niche, and ask if they had any "
               "questions. One theme name max. No subject needed (reply keeps thread "
               "subject) but still output SUBJECT line as short."
               "\nOutput:\nSUBJECT: [short]\nBODY:\n[body]")
        user = ("Store: " + store + "\nIndustry: " + c["industry"] +
                "\nThemes we suggested earlier: " + (c["suggested"] or ", ".join(primary)) +
                "\n\nPrevious emails in this thread:\n" + (history or "(unavailable)") +
                "\n\nWrite a short bump that builds on the thread above without "
                "repeating it.")
    elif touch == 3:
        sys = (BASE +
               "\n\nTHIRD email, follow-up in the same thread. 70-110 words, hard cap "
               "120. You are shown the previous emails in this thread: take a DIFFERENT "
               "angle from both earlier notes, never repeat what was already said. Focus "
               "on ONE concrete value point relevant to their store: UTD themes have "
               "conversion features built in (upsells, cross-sells, promo sections) "
               "which usually cut spend on extra apps, and switching a theme is low-risk "
               "(preview before publish, products stay). Re-name the top 2 themes for "
               "their niche. End with a light question that nudges them to pick a theme."
               "\nOutput:\nSUBJECT: [short]\nBODY:\n[body]")
        user = ("Store: " + store + "\nIndustry: " + c["industry"] + "\nTop themes: " +
                (c["suggested"] or ", ".join(primary)) +
                "\n\nPrevious emails in this thread:\n" + (history or "(unavailable)") +
                "\n\nWrite the value follow-up. It must add something new versus the "
                "thread above.")
    else:
        sys = (BASE +
               "\n\nFOURTH and final email, a polite breakup in the same thread. 45-70 "
               "words, hard cap 120. You are shown the previous emails in this thread: "
               "acknowledge the sequence naturally, do not re-pitch and do not repeat "
               "earlier lines. Say you will assume the timing is not right and will not "
               "keep following up, leave the catalog link " + REGISTRY + " for whenever "
               "it fits, wish them well. Warm, no guilt, no pressure."
               "\nOutput:\nSUBJECT: [short]\nBODY:\n[body]")
        user = ("Store: " + store +
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
                c["industry"].lower() if c["industry"] != "Other" else "store")
            body = ("Hi " + (c["store_name"] or "there") + " team,\n\n"
                    "I came across your store and wanted to reach out. I am Sergey from "
                    "UTD Web, we build themes for the official Shopify Theme Store with "
                    "conversion features built in (upsells, cross-sells, promo sections)."
                    "\n\nFor a store in your niche I would confidently suggest " + p[0] +
                    " (" + CATALOG[p[0]]["price"] + ") or " + p[1] + " (" +
                    CATALOG[p[1]]["price"] + "): " + CATALOG[p[0]]["pitch"] + ".\n\n"
                    "Also worth a look: " + p[2] + ", " + ", ".join(a) + ". Full catalog: " +
                    REGISTRY + " and " + SITE + "\n\n"
                    "Happy to answer questions or share a demo store.")
        elif touch == 2:
            subject = "Quick follow-up"
            body = ("Hi " + (c["store_name"] or "there") + ",\n\n"
                    "Just following up on my note about UTD themes for your store. Happy "
                    "to send you a live demo of " + p[0] + " so you can see it with content "
                    "like yours. Any questions so far?")
        elif touch == 3:
            subject = "One more thought"
            body = ("Hi " + (c["store_name"] or "there") + ",\n\n"
                    "One thing worth mentioning: our themes have upsells, cross-sells and "
                    "promo sections built in, which usually means fewer paid apps. Switching "
                    "is low-risk since you preview before publishing and your products stay "
                    "in place. " + p[0] + " and " + p[1] + " fit your niche well. Worth a "
                    "quick look?")
        else:
            subject = "Closing the loop"
            body = ("Hi " + (c["store_name"] or "there") + ",\n\n"
                    "I will assume the timing is not right and stop following up. If you "
                    "ever want to explore a new theme, the full catalog is here: " +
                    REGISTRY + ". Wishing you a great quarter.")

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
