#!/usr/bin/env python3
"""
ecom_autoresponder.py — UTD eCommerce (Shopify store owner) inbound autoresponder.

Faithful port of the n8n workflow «ECOM_auto» (ak23YgaINtgxkbjM) to plain Python
for GitHub Actions. Reads one or more Gmail mailboxes over IMAP, classifies each
new inbound reply to our UTD-themes cold outreach, drives the store-owner funnel
with Claude, and replies in-thread over SMTP (Gmail app-passwords).

n8n flow reproduced (node → step):
  «Входящие»(getAll) → «Разбор»(classify) → «Ecom Contacts»(read CRM) →
  «Обогащение»(match CRM, skip non-CRM humans, dedup on Last Msg ID) →
  «Нужен AI?»(if preCategory empty) → «Собрать запрос»/«Claude»/«Итог AI»
    (AI branch) OR «Без AI» (bounce/send_failed/auto_reply branch) →
  «Слияние» → «Маршрут»(switch) → per-route: reply + sheet status + mark read.

Route → CRM Status (VERBATIM from the n8n sheet nodes):
  respond      → "Replied"      (send reply, write Status + Date Replied)
  decline      → "Declined"     (write Status + Date Replied)
  bounce       → "Bounced"      (write Status + Date Replied)
  send_failed  → "Send Failed"  (write Status + Date Replied)
  escalate     → (no write)     (n8n «Эскалация: непрочитано» — leave for a human)
  spam/ignore  → (no write)     (n8n «Прочитано: игнор» — just mark handled)
  auto_reply   → (no write)     (routes to the ignore output in the n8n switch)

Safety / GHA specifics:
  • DRY_RUN=true (default) prints drafts + intended sheet writes; nothing is sent.
  • Processed Message-IDs are SHA256-hashed into <STATE_DIR>/ecom_autoresponder_state.json
    (repo is PUBLIC — no raw emails/addresses committed). This replaces the n8n
    markAsRead trick as the dedup mechanism.

Usage:  python ecom_autoresponder.py
Env:    GMAIL_APP_PW_SERGEY, GMAIL_APP_PW_SERGI, GMAIL_APP_PW_SERGE,
        ANTHROPIC_API_KEY, GOOGLE_CREDENTIALS_JSON,
        ECOM_SHEET_ID, ECOM_SHEET_TAB, DRY_RUN, LOOKBACK_DAYS, STATE_DIR
"""

import os
import re
import time
import json
from datetime import datetime, timezone

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

# Env var names match ecom_harvester.py (ECOM_SHEET_ID / ECOM_SHEET_TAB).
# Defaults mirror the n8n «ECOM_auto» workflow (spreadsheet id + "Ecom Contacts"
# tab), which is the sheet this autoresponder actually operates on.
SHEET_ID = os.environ.get("ECOM_SHEET_ID", "").strip()
SHEET_TAB = os.environ.get("ECOM_SHEET_TAB", "Ecom Contacts")

# Our own mailbox addresses — inbound from these is a loop, not a prospect.
# (VERBATIM from the n8n «Разбор» OWN list.)
OWN_ADDRESSES = [a for a in (
    os.environ.get("UTD_MAIL_SERGEY", ""),
    os.environ.get("UTD_MAIL_SERGI", ""),
    os.environ.get("UTD_MAIL_SERGE", ""),
    os.environ.get("UTD_MAIL_SERHII", ""),
) if a]

# Mailboxes to process. Skipped automatically when no app-password is set.
ACCOUNTS = [a for a in (
    {"user": os.environ.get("UTD_MAIL_SERGEY", ""), "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")},
    {"user": os.environ.get("UTD_MAIL_SERGI", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGI", "")},
    {"user": os.environ.get("UTD_MAIL_SERGE", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGE", "")},
) if a["user"]]

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "ecom_autoresponder_state.json")

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")
# Scan the WHOLE inbox over a wide window (not just unread) — dedup is by hashed
# Message-ID in state, so already-read replies we missed before are still caught.
LOOKBACK_DAYS = int(os.environ.get("LOOKBACK_DAYS", "60"))

# CRM columns the sheet is expected to carry (for reference / matching).
EXPECTED_COLUMNS = [
    "Email", "Thread ID", "Store Name", "Industry",
    "Suggested Themes", "Last Msg ID", "Status", "Date Replied",
]

# Switch order from the n8n «Маршрут» node. A category not in this list falls
# through to the "ignore" output (mark handled, no reply, no sheet write) — this
# is exactly what happens to 'auto_reply' and 'ignore' (spam) in the workflow.
ROUTE_OUTPUTS = ["respond", "decline", "bounce", "send_failed"]

# Team report inboxes — where a CLOSED ecom deal (merchant committed to buying) is
# reported. Same list as the influencer report / signed-contract alert.
_DEFAULT_REPORT_TO = ("denvdavydov@gmail.com,marketing@utdweb.team,"
                      "denys.davydov.utd@gmail.com,sergey.smortkin.utd@gmail.com,"
                      "george.smortkin@gmail.com")
REPORT_TO = [x.strip() for x in
             (os.environ.get("REPORT_NOTIFY_TO") or os.environ.get("SIGNED_NOTIFY_TO")
              or _DEFAULT_REPORT_TO).split(",") if x.strip()]

# Shared "Notable emails" log (unusual/strange messages → address + mailbox).
NOTABLE_SHEET_ID = os.environ.get("NOTABLE_SHEET_ID") or os.environ.get("B2B_SHEET_ID", "")
NOTABLE_TAB = os.environ.get("NOTABLE_TAB", "Notable")
CLOSED_TAB = os.environ.get("CLOSED_TAB", "Closed")

# ── In-run caches (populated once per run; NO per-email Sheets calls) ──
# _THREAD_CACHE: {(account_user, gm_thrid): history_text} so each Gmail thread
# is fetched at most once per run (same pattern as agency_autoresponder).
_SHEET = {"ws": None, "header": [], "email_to_row": {}}
_PENDING = {}
_THREAD_CACHE = {}


# ═══════════════════════════════════════════════════════════════════
#   PROMPTS  (ported from the n8n «Собрать запрос» node; voice reworked 2026-07)
# ═══════════════════════════════════════════════════════════════════

MODEL = "claude-sonnet-5"
MAX_TOKENS = 1300

SYSTEM_PROMPT = (
"You are Sergey, a normal guy from UTD Web (an IT company that builds Shopify themes), replying to a Shopify store owner who answered our cold email about UTD themes. UTD themes (official Shopify Theme Store, developer UTD Web, links https://themes.shopify.com/themes?q=UTD and https://utdweb.team): 6 theme families, 30 presets. Impression $340 (premium flagship, EU translations, cross-sell, mega menu, size chart, pre-order). Victory $320 (sports/events/active: store locator, event calendar, age verifier, countdowns). Boutique $160 (boutiques/premium brands). Ultra $100 (tech/furniture/auto/toys). Allure $100 (beauty/lifestyle). Gain $100 (minimalist). All themes: built-in upsells, cross-sells, promo sections; EU translations; preview before publish; products stay in place when switching. Themes are bought on the official Shopify Theme Store; UTD can help with setup and customization on request.\n\n"
"PRESET REGISTRY (a preset is a ready design variant of ONE theme; a preset belongs to that ONE theme only; NEVER attribute a preset to a different theme, a wrong pair makes a broken demo link):\n"
"- Victory ($320): Victory, Athletica, Nitro, Roast, Flip\n"
"- Ultra ($100): Ultra, Grace, Grip, Harbor, Sprout\n"
"- Gain ($100): Gain, Lace, Maison, Mio, Sable\n"
"- Allure ($100): Allure, Bijou, Carrara, Pristine, Stitch\n"
"- Boutique ($160): Boutique, Aurum, Jade, Noom, Reflections\n"
"- Impression ($340): Impression, Etoile, Felix, Mimi, Reflex\n"
"LINKS: theme page = https://themes.shopify.com/themes/<theme-lowercase>. Preset demo = https://themes.shopify.com/themes/<theme-lowercase>/presets/<preset-lowercase> (example: preset Roast of theme Victory = https://themes.shopify.com/themes/victory/presets/roast). EVERY time you name a theme or a preset, put its link right next to it. A demo = the live preview on that page. The only links allowed in a reply: https://utdweb.team, https://themes.shopify.com/themes?q=UTD, and these theme/preset pages.\n\n"
"FUNNEL GOAL: SELL, do not just inform: convince this merchant to choose and buy ONE UTD theme. Stick to ONE preset/theme across the whole thread: the one the thread already points to (the preset we suggested earlier or the one they asked about). Do not switch to a new theme in every email; if the thread does not point anywhere yet, ask one easy question about their store to pick the right one. Argue from THEIR SALES (speed, layout, checkout drive orders) and prove claims with the evidence below woven into the argument. Build the argument around THEIR pain: a slow store, monthly app fees stacking up, weak conversion, a clunky theme. Value line you may use: built-in upsell, cross-sell and promo blocks replace apps that cost $15-50 per month, so the one-time theme price pays for itself. Official Shopify Theme Store status means a safe purchase: preview before publish, products stay in place. NEVER ask the merchant about their metrics, speed scores or conversion numbers (they do not know them). Answer their questions accurately using ONLY the facts above. Do not invent facts, features, prices or numbers. Do not offer discounts. Never offer or suggest a call or meeting: everything is handled by email; you may offer help by email ('reply and I'll walk you through it'). You handle the whole conversation yourself and NEVER hand it to a human: if they ask about custom development, a specific dated call, contracts, or something not covered here, do not refuse or stall, answer plainly from the facts you have and, for anything you truly cannot cover, say you will note it and the team will follow up by email, then keep guiding them toward choosing and buying the theme.\n\n"
"EVIDENCE you may cite (real and named; use AT MOST ONE per reply; never cite any other statistic; NEVER park it in its own paragraph, weave it into the sales argument at the moment you make the claim): Google/SOASTA 2017: bounce probability grows 32% as mobile load goes 1s to 3s. Deloitte and Google 'Milliseconds Make Millions' 2020: a 0.1s speed improvement lifted retail conversions ~8.4%. Portent 2022: 1s sites convert ~2.5x better than 5s sites. Baymard Institute: ~70% of carts are abandoned; better checkout design alone recovers ~35% conversion for an average large store. Business cases (published by Google/web.dev): Vodafone made pages 31% faster and sales rose 8%; Rakuten 24 invested in Core Web Vitals and got +33% conversion and +53% revenue per visitor; Swappie grew mobile revenue 42% after speeding up its mobile site. These prove the MECHANISM; never claim a specific result for our themes.\n\n"
"Return STRICT JSON: {\"category\":\"interested|question|decline|spam\",\"deal_closed\":false,\"handoff\":false,\"handoff_note\":\"\",\"notable\":false,\"notable_reason\":\"\",\"note\":\"<short RU>\",\"reply_body\":\"<reply or empty>\"}\n"
"- notable: set true if the email is UNUSUAL or memorable — strange/funny, an unexpected or unusual request, a big/well-known store or person, an angry or odd tone, anything worth finding again later. notable_reason = one short RUSSIAN line why. Ordinary replies are notable=false.\n"
"- deal_closed: set true ONLY when the merchant clearly commits to going ahead with a theme (they say they will buy it, are purchasing it, ask exactly how to complete the purchase, or confirm they picked it). A general 'interested' is NOT closed. When true, still write a warm reply that helps them finish the purchase.\n"
"- handoff: set true whenever your reply tells the merchant that 'the team' will follow up on something you could not resolve yourself (custom development, a dated call, a contract point, anything not in the facts). Write handoff_note in RUSSIAN as a DETAILED brief for the team (4-7 sentences, not one line): (1) что за магазин и кто пишет, (2) что именно просят — со всеми деталями: цифры, требования, ссылки, (3) почему это вне того, что ты можешь сам, (4) какое решение или действие нужно от команды, (5) важный контекст из переписки. Развёрнуто и конкретно, чтобы человек всё понял без открытия оригинала. Leave false if you handled it fully. You still send the reply either way.\n"
"- interested/question: reply_body required. WRITING RULES:\n"
"  * Reply in the LANGUAGE of the incoming email.\n"
"  * SIMPLE ENGLISH for non-native readers (and the same simple wording in any other language): common everyday words, and write in LONG, flowing, simple sentences that go straight to the point (never short choppy ones) — real people write long simple sentences, not staccato fragments. No idioms, no slang, no fancy phrases ('caught my eye', 'worth a look' and anything similar are forbidden). If a 12-year-old would not understand a sentence, rewrite it.\n"
"  * Write like a normal person typing an email by hand. If a sentence reads like AI or a sales script, rewrite it. Zero filler, maximum concreteness. Never open with a generic compliment.\n"
"  * FORMAT (mandatory): line 1 is a greeting; then a blank line; then the body grouped by meaning into a few paragraphs with a blank line between them, each paragraph written as LONG, flowing, simple sentences that get straight to the point, never short choppy ones. Then a blank line, the farewell and the signature.\n"
"  * This is a reply inside a thread: the first sentence after the greeting refers naturally to what they wrote or to the earlier exchange. Add only NEW substance; never repeat what was already said (use the thread history).\n"
"  * Never use an em dash. Forbidden words: exclusive, exciting, game-changer, handpicked, curated, unique opportunity.\n"
"  * Recommend the ONE preset/theme of this thread (with its link) tied to their pain with one concrete reason. End by moving them one step closer to buying: offer to do something concrete by email (for example: 'reply with your store link and I will say which preset fits best') or ask one easy preference question (for example: which of two presets they like more). Never end with a question about their metrics.\n"
"  * End with exactly:\n"
"Best regards,\n"
"Sergey\n"
"UTD Web | utdweb.team\n"
"- decline/spam: reply_body empty.\n"
"Output ONLY the JSON."
)


def build_user_prompt(msg, store, industry, suggested, history=""):
    """The user message fed to Claude. Now carries the FULL prior thread
    (fetched over IMAP via fetch_thread) so the reply builds on what was
    already said instead of re-pitching from scratch."""
    return (
        "Incoming reply:\n"
        "From: " + msg["from"] + "\n"
        "Store: " + (store or "?") + "\n"
        "Industry: " + (industry or "?") + "\n"
        "Themes we suggested: " + (suggested or "?") + "\n\n"
        "Previous emails in this thread (oldest first):\n" +
        (history or "(no prior history found)") + "\n\n"
        "Latest incoming message body:\n" + msg["body"] + "\n\n"
        "Classify and reply, building on the thread above. Output ONLY JSON."
    )


def format_thread_history(thread, budget=4000):
    """Render a fetch_thread() list into a compact chronology for the model
    (oldest first), trimmed to ~`budget` chars total."""
    if not thread:
        return ""
    lines = []
    for m in thread:
        who = "UTD (us)" if m.get("direction") == "sent" else "Merchant"
        snippet = (m.get("snippet") or "").strip()
        lines.append("- [%s] %s: %s" % (str(m.get("date", ""))[:10], who, snippet))
    return "\n".join(lines)[:budget]


def get_thread_history(account, msg, contact):
    """Fetch the whole Gmail conversation for this message (both directions),
    at most once per thread per run. Thread id: the message's X-GM-THRID first,
    then the CRM row's Thread ID when numeric. '' when nothing can be fetched."""
    thrid = str(msg.get("gm_thrid", "") or "").strip()
    if not thrid and contact:
        t = str(contact.get("thread", "") or "").strip()
        if t.isdigit():
            thrid = t
    if not thrid:
        return ""
    cache_key = (account["user"], thrid)
    if cache_key in _THREAD_CACHE:
        return _THREAD_CACHE[cache_key]
    history = ""
    try:
        history = format_thread_history(
            ec.fetch_thread(account, thrid, OWN_ADDRESSES))
    except Exception as e:
        print(f"  (thread fetch failed, using single message: {e})")
    _THREAD_CACHE[cache_key] = history
    return history


# ═══════════════════════════════════════════════════════════════════
#   AI / non-AI result parsing  (ported from «Итог AI» and «Без AI»)
# ═══════════════════════════════════════════════════════════════════

def _clean_reply(text):
    """Canon guard: no em/en dashes in an outgoing reply. Digit ranges keep a
    plain hyphen ($15-50), any other dash becomes a comma pause."""
    t = (text or "").strip()
    if not t:
        return t
    t = re.sub(r"(?<=\d)\s*[—–]\s*(?=\d)", "-", t)
    t = re.sub(r"\s*[—–]\s*", ", ", t)
    return t


def parse_ai_result(text):
    """Parse Claude's strict-JSON output into a routing decision.

    Faithful to the n8n «Итог AI» node: on any parse failure (missing/invalid
    JSON, or an empty API response) the category DEFAULTS to 'escalate' — the
    workflow never leaves a message undecided, it hands it to a human.
    """
    cat, reply, note, closed = None, "", "AI не разобрал", False
    handoff, handoff_note = False, ""
    notable, notable_reason = False, ""
    try:
        m = re.search(r"\{[\s\S]*\}", text or "")
        p = json.loads(m.group(0))
        if p.get("category") in ("interested", "question", "decline", "spam"):
            cat = p["category"]
        reply = _clean_reply(p.get("reply_body"))
        note = (p.get("note") or "").strip() or note
        closed = bool(p.get("deal_closed"))
        handoff = bool(p.get("handoff"))
        handoff_note = (p.get("handoff_note") or "").strip()
        notable = bool(p.get("notable"))
        notable_reason = (p.get("notable_reason") or "").strip()
    except Exception:
        pass

    # We never hand off to a human. If unparseable, or it wants to reply but drafted
    # nothing, leave it for the next run to retry cleanly.
    if cat is None or (cat in ("interested", "question") and not reply):
        return None

    route = "respond" if cat in ("interested", "question") else (
        "ignore" if cat == "spam" else "decline")
    status = "Replied" if route == "respond" else (
        "Declined" if cat == "decline" else "")
    return {"category": route, "ai_category": cat, "note": note,
            "reply_body": reply, "new_status": status, "deal_closed": closed,
            "handoff": handoff, "handoff_note": handoff_note,
            "notable": notable, "notable_reason": notable_reason}


def non_ai_result(pre_category):
    """Map a classifier pre-category to a route (ported from «Без AI»)."""
    mapping = {"bounce": "bounce", "send_failed": "send_failed", "auto_reply": "auto_reply"}
    route = mapping.get(pre_category, "ignore")
    status = {"bounce": "Bounced", "send_failed": "Send Failed",
              "auto_reply": "Auto Reply"}.get(pre_category, "")
    return {"category": route, "ai_category": pre_category, "note": "",
            "reply_body": "", "new_status": status}


# ═══════════════════════════════════════════════════════════════════
#   CRM matching  (ported from «Обогащение»)
# ═══════════════════════════════════════════════════════════════════

def build_index(rows):
    """Index CRM rows by Email and by Thread ID (Gmail thread id)."""
    by_email, by_thread = {}, {}
    for r in rows:
        info = {
            "email": str(r.get("Email", "")).strip().lower(),
            "store": str(r.get("Store Name", "")).strip(),
            "industry": str(r.get("Industry", "")).strip(),
            "suggested": str(r.get("Suggested Themes", "")).strip(),
            "last": str(r.get("Last Msg ID", "")).strip(),
            "thread": str(r.get("Thread ID", "")).strip(),
            "status": str(r.get("Status", "")).strip(),
        }
        if info["email"]:
            by_email[info["email"]] = info
        if info["thread"]:
            by_thread[info["thread"]] = info
    return by_email, by_thread


def match_contact(msg, by_email, by_thread):
    """Match an inbound message to a CRM row: Thread ID first, then sender email
    (VERBATIM ordering from «Обогащение»: byThread[threadId] || byEmail[email])."""
    thrid = str(msg.get("gm_thrid", "")).strip()
    if thrid and thrid in by_thread:
        return by_thread[thrid]
    return by_email.get(msg.get("from_email", ""), None)


# ═══════════════════════════════════════════════════════════════════
#   Actions (send + sheet), guarded by DRY_RUN
# ═══════════════════════════════════════════════════════════════════

def _reply_subject(subject):
    s = subject or ""
    return s if s.lower().startswith("re:") else "Re: " + s


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


def _print_draft(kind, account, to, subject, body):
    print("\n" + "=" * 70)
    print(f"[DRAFT · {kind}]  DRY_RUN — not sent")
    print(f"  from   : {account['user']}")
    print(f"  to     : {to}")
    print(f"  subject: {subject}")
    print("  body:")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def do_reply(account, msg, decision):
    """Send the funnel reply in-thread (n8n «Автоответ» gmail reply node)."""
    subject = _reply_subject(msg["subject"])
    if DRY_RUN:
        _print_draft("reply", account, msg["from_email"], subject,
                     decision["reply_body"])
        return
    ec.send_email(account, msg["from_email"], subject, decision["reply_body"],
                  in_reply_to=msg["message_id"], references=msg["references"])


def _readable_thread(account, msg, contact):
    """Full conversation rendered as clearly-separated messages (who/when/what)."""
    try:
        return ec.format_thread_readable(
            ec.fetch_thread(account, str(msg.get("gm_thrid", "") or ""), OWN_ADDRESSES))
    except Exception:
        return ec.format_thread_readable([{
            "direction": "received", "from_email": msg.get("from_email", ""),
            "date": msg.get("date", ""), "subject": msg.get("subject", ""),
            "body": msg.get("body", "")}])


def send_handoff_alert(account, msg, contact_email, contact, decision):
    """The agent told the merchant 'the team' will follow up on something it could
    not resolve → email the team a DETAILED brief + the separated thread."""
    store = (contact.get("store") if contact else "") or contact_email
    subject = f"⚠️ Требуется команда (ecom) — {store}"
    body = (
        "Агент продолжил продажу сам, но не смог что-то решить и сказал мерчанту, "
        "что команда свяжется. Ниже — что нужно сделать.\n\n"
        "════════ КРАТКО ════════\n"
        f"Магазин:  {store}\n"
        f"Email:    {contact_email}\n"
        f"Ящик:     {account.get('user', '')}\n"
        f"Тема:     {msg.get('subject', '')}\n\n"
        "ЧТО НУЖНО / ПРОБЛЕМА (развёрнуто от агента):\n"
        f"{decision.get('handoff_note') or '(см. переписку ниже)'}\n\n"
        "════════ ЧТО АГЕНТ ОТВЕТИЛ ════════\n"
        f"{decision.get('reply_body', '')}\n\n"
        "════════ ПОЛНАЯ ПЕРЕПИСКА (по письмам) ════════\n"
        f"{_readable_thread(account, msg, contact)}\n\n"
        "— Автоматическая передача от UTD ecom agent"
    )
    if DRY_RUN:
        print(f"  [HANDOFF] DRY_RUN — would alert team re {contact_email}")
        return
    ec.send_email(account, REPORT_TO, subject, body, from_name="UTD Ecom Agent")
    print(f"  [HANDOFF] team alerted for {contact_email}: {(decision.get('handoff_note') or '')[:60]}")


def send_deal_report(account, msg, contact_email, contact, decision):
    """Deal closed (merchant committed to buying) → email the finished conversation
    to the team report inboxes. Sent once per contact."""
    store = (contact.get("store") if contact else "") or "(store)"
    subject = f"✅ Ecom-сделка закрыта — {store}"
    body = (
        "Мерчант решил купить тему UTD — диалог готов, команда может помочь "
        "завершить покупку.\n\n"
        f"Магазин: {store}\n"
        f"Email:   {contact_email}\n\n"
        "════════ ПОЛНАЯ ПЕРЕПИСКА (по письмам) ════════\n"
        f"{_readable_thread(account, msg, contact)}\n\n"
        "— Автоматический отчёт от UTD ecom agent"
    )
    if DRY_RUN:
        print(f"  [REPORT] DRY_RUN — would send closed-deal report to {REPORT_TO}")
        return
    ec.send_email(account, REPORT_TO, subject, body, from_name="UTD Ecom Agent")
    print(f"  [REPORT] closed-deal report sent to {len(REPORT_TO)} recipients for {contact_email}")


def enqueue_update(email, updates):
    """Queue a per-column update for a contact (merged by email; last write wins).
    NOTHING hits the Sheets API here — everything is flushed once at run end."""
    if not email or not updates:
        return
    key = str(email).strip().lower()
    row = _SHEET["email_to_row"].get(key)
    if not row:
        # Contact is not in the CRM sheet — nothing to update (we never add rows).
        print(f"  [SHEET] {key} not found in CRM → skipped (no row to update)")
        return
    bucket = _PENDING.setdefault(key, {})
    for col, val in updates.items():
        if col in _SHEET["header"]:
            bucket[col] = val


def write_status(email, status):
    """Queue Status + Date Replied for a contact matched by Email (VERBATIM the
    columns written by every status node in the n8n workflow)."""
    if not status or not email:
        return
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    enqueue_update(email, {"Status": status, "Date Replied": now})
    print(f"  [SHEET] queued Status='{status}' for {email}")


def flush_updates():
    """Write ALL queued contact updates in a SINGLE Sheets batchUpdate call."""
    if not _PENDING:
        print("\n[SHEET] no updates to flush.")
        return
    header = _SHEET["header"]
    cell_updates = []
    for email, cols in _PENDING.items():
        row = _SHEET["email_to_row"].get(email)
        if not row:
            continue
        for col, val in cols.items():
            if col not in header:
                continue
            a1 = ec.gspread_a1(row, header.index(col) + 1)
            cell_updates.append({"range": a1, "values": [[val]]})
    if DRY_RUN:
        print(f"\n[SHEET] DRY_RUN — would batch-write {len(cell_updates)} cells "
              f"across {len(_PENDING)} contacts in ONE call:")
        for email, cols in _PENDING.items():
            print(f"    {email}: {cols}")
        return
    if not _SHEET["ws"]:
        print("\n[SHEET] no worksheet handle — cannot flush.")
        return
    try:
        n = ec.batch_update_cells(_SHEET["ws"], cell_updates)
        print(f"\n[SHEET] batch-wrote {n} cells across {len(_PENDING)} contacts in ONE call.")
    except Exception as e:
        print(f"\n⚠️  [SHEET] batch update failed after retries: {e}")


# ═══════════════════════════════════════════════════════════════════
#   Per-message processing
# ═══════════════════════════════════════════════════════════════════

def process_message(account, msg, by_email, by_thread, state, stats):
    mid = msg.get("message_id", "")

    # Dedup: already processed in a previous run (hashed Message-ID)?
    if ec.is_processed(state, mid):
        return

    contact = match_contact(msg, by_email, by_thread)
    pre = ec.classify_incoming(msg, OWN_ADDRESSES)

    # «Обогащение» rule: not in the eCom base AND a plain human reply → ignore
    # entirely (never spend an AI call on strangers who are not in the CRM).
    if pre == "human" and contact is None:
        return

    # «Обогащение» dedup: this exact message is the last one we already acted on.
    if contact and contact.get("last") and contact["last"] == mid:
        if not DRY_RUN:
            ec.mark_processed(state, mid)
        return

    if pre != "human":
        decision = non_ai_result(pre)
    else:
        store = contact["store"] if contact else ""
        industry = contact["industry"] if contact else ""
        suggested = contact["suggested"] if contact else ""
        # Pull the WHOLE thread (both directions) so the model reads the real
        # history, not just this one inbound email (agency_autoresponder pattern).
        history = get_thread_history(account, msg, contact)
        user = build_user_prompt(msg, store, industry, suggested, history)
        if DRY_RUN:
            _print_prompt(SYSTEM_PROMPT, user)
        # Gentle pacing between Claude calls to avoid 429 on big batches.
        time.sleep(0.7)
        ai_text = ec.call_claude(SYSTEM_PROMPT, user, model=MODEL, max_tokens=MAX_TOKENS)
        decision = parse_ai_result(ai_text) if (ai_text or "").strip() else None
        if decision is None:
            # Empty or unparseable AI answer. Normally we leave it for the next
            # run — but a message that fails EVERY cycle would be re-sent to
            # Claude forever (hundreds of paid calls/day). Cap it: after 3 runs
            # give up, mark processed, and log it to Notable for a manual look.
            reason = "empty AI response" if not (ai_text or "").strip() else "unparseable AI result"
            print(f"\n· {account['user']} ← {msg['from_email']} | {msg['subject'][:60]!r}")
            if ec.bump_attempt(state, mid, limit=3):
                ec.mark_processed(state, mid)
                print(f"  GAVE UP after 3 tries ({reason}) → marked processed to stop the retry loop.")
                try:
                    ec.append_notable(NOTABLE_SHEET_ID, NOTABLE_TAB, {
                        "Date": (msg.get("date", "") or "")[:16], "Chain": "Ecom",
                        "Account": account.get("user", ""), "From": msg.get("from_email", ""),
                        "Subject": msg.get("subject", ""),
                        "Why notable": f"Claude {reason} 3x — auto-skipped, needs manual look"})
                except Exception:
                    pass
            else:
                print(f"  UNDECIDED ({reason}) → retry next run, not marked processed.")
            return

    # «Маршрут» switch: a category not in ROUTE_OUTPUTS falls to the ignore output.
    route = decision["category"] if decision["category"] in ROUTE_OUTPUTS else "ignore"
    stats[route] = stats.get(route, 0) + 1
    print(f"\n· {account['user']} ← {msg['from_email']} | {msg['subject'][:60]!r}")
    print(f"  route={route} status→{decision['new_status'] or '-'} | {decision['note']}")

    if decision.get("notable"):
        tgt = (contact.get("email") if contact else "") or msg.get("from_email", "")
        ec.append_notable(NOTABLE_SHEET_ID, NOTABLE_TAB, {
            "Date": (msg.get("date", "") or "")[:16], "Chain": "Ecom",
            "Account": account.get("user", ""), "From": tgt,
            "Subject": msg.get("subject", ""), "Why notable": decision.get("notable_reason", "")})
        print(f"  [NOTABLE] logged: {decision.get('notable_reason','')[:70]}")

    if route == "respond":
        do_reply(account, msg, decision)
        target = (contact.get("email") if contact else "") or msg["from_email"]
        # «Статус: диалог»: Status = new_status || 'Replied'
        write_status(target, decision["new_status"] or "Replied")
        # Deal closed = merchant committed to buying the theme. Report the finished
        # conversation to the team inboxes ONCE per contact.
        if decision.get("deal_closed") and target \
                and not ec.is_processed(state, "reported:" + target.lower()):
            write_status(target, "Deal Closed")
            send_deal_report(account, msg, target, contact, decision)
            ec.append_closed(NOTABLE_SHEET_ID, CLOSED_TAB, {
                "Date": (msg.get("date", "") or "")[:16], "Chain": "Ecom",
                "Contact": target,
                "Company/Store": (contact.get("store") if contact else "") or "",
                "Account": account.get("user", ""),
                "Outcome": "Deal closed (merchant buying a theme)",
                "Details / review": (decision.get("note") or "")[:900]})
            ec.mark_processed(state, "reported:" + target.lower())
        # Agent deferred something to "the team" → send us the summary + problem.
        if decision.get("handoff"):
            send_handoff_alert(account, msg, target, contact, decision)
    elif route == "decline":
        target = (contact.get("email") if contact else "") or msg["from_email"]
        write_status(target, "Declined")
    elif route in ("bounce", "send_failed"):
        # The DSN comes FROM mailer-daemon, not the contact. Pull the real failed
        # recipient out of the bounce body/headers and flag THAT contact.
        rcpt = ec.extract_failed_recipient(msg, OWN_ADDRESSES)
        if rcpt:
            print(f"  failed recipient = {rcpt} → Status '{decision['new_status']}'")
            write_status(rcpt, decision["new_status"])
            if route == "bounce":
                ec.mark_bounced_everywhere(rcpt, dry_run=DRY_RUN)
        else:
            print("  failed recipient not found → sheet left untouched.")
    # 'ignore' (spam / auto_reply) → mark handled, nothing else. No escalate route:
    # the agent answers every real message itself and reports only on a closed deal.

    if not DRY_RUN:
        ec.mark_processed(state, mid)


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def run_once():
    print(f"=== UTD eCom autoresponder | DRY_RUN={DRY_RUN} | "
          f"lookback {LOOKBACK_DAYS}d | {datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)

    # Reset in-run caches / pending writes.
    _PENDING.clear()
    _THREAD_CACHE.clear()
    _SHEET["ws"] = None
    _SHEET["header"] = []
    _SHEET["email_to_row"] = {}

    # CRM snapshot: open the worksheet ONCE and read ONCE (with 429 backoff).
    rows = []
    try:
        ws = ec.open_worksheet(SHEET_ID, SHEET_TAB)
        rows = ec.read_rows_ws(ws)
        _SHEET["ws"] = ws
        _SHEET["header"] = list(rows[0].keys()) if rows else ws.row_values(1)
        e2r = {}
        for i, r in enumerate(rows, start=2):
            em = str(r.get("Email", "")).strip().lower()
            if em and em not in e2r:
                e2r[em] = i
        _SHEET["email_to_row"] = e2r
    except Exception as e:
        print(f"⚠️  Could not read CRM sheet: {e}")

    by_email, by_thread = build_index(rows)
    print(f"CRM: {len(rows)} rows read in 1 call, {len(by_email)} emails indexed, "
          f"{len(by_thread)} threads indexed, {len(_SHEET['email_to_row'])} rows mapped.")

    stats = {}
    for account in ACCOUNTS:
        if not account["password"]:
            print(f"⚠️  No app-password for {account['user']} — skipping mailbox.")
            continue
        try:
            # Scan the WHOLE inbox (read + unread) over the lookback window.
            msgs = ec.fetch_inbox(account, since_days=LOOKBACK_DAYS, unseen_only=False)
        except Exception as e:
            print(f"⚠️  IMAP error for {account['user']}: {e}")
            continue
        print(f"\n>>> {account['user']}: {len(msgs)} messages in last {LOOKBACK_DAYS}d")
        for msg in msgs:
            try:
                process_message(account, msg, by_email, by_thread, state, stats)
            except Exception as e:
                print(f"  !! error on message {msg.get('message_id','?')}: {e}")

    # Flush ALL queued CRM updates in a single Sheets batchUpdate call.
    flush_updates()

    if not DRY_RUN:
        ec.save_state(STATE_FILE, state)
    print(f"\n=== done. routes: {stats} ===")
    return {"parser": "ecom_autoresponder", "routes": stats,
            "dry_run": DRY_RUN, "sheet_writes_queued": len(_PENDING)}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
