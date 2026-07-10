#!/usr/bin/env python3
"""
influencer_autoresponder.py — UTD influencer / content-creator inbound autoresponder.

Port of the n8n workflow INFL_auto (hWvHzjmYCJfBvLLW) to plain Python for GitHub
Actions. Mirrors agency_autoresponder.py structurally: poll the outreach mailbox(es)
over IMAP, keep ONLY the influencer chain ("shopify theme review collab"), match the
sender to the CRM, classify the reply with Claude, extract/merge rate-card data, send
a Claude-drafted reply in-thread over SMTP, and write the resulting status back to two
tabs of the influencer CRM spreadsheet — all guarded by DRY_RUN.

The n8n chain uses TWO tabs of the same spreadsheet:
  • "Sheet1"  (Контакты)  — one row per creator: Email, Name/Channel/Company Name,
                            Thread ID, Status. Status is set to Declined / Bounced /
                            Send Failed / Auto Reply here.
  • "Pricing" (Прайсы)    — the rate-card CRM: Email, Name, Platform, Contact Status,
                            all Price* / audience / notes columns, Last Msg ID.
                            appendOrUpdate (a new creator gets a fresh row).

Faithfulness notes (kept verbatim from the n8n code nodes):
  • Claude system + user prompts, model "claude-sonnet-5", max_tokens 1400.
  • Categories: interested | question | decline | spam | escalate.
  • Route mapping and the 17 rate-card data fields + merge (new value overrides old
    only when non-empty).
  • Contact Status values written to Pricing: Data Complete / Negotiating / Declined
    (Escalated is COMPUTED but never written — escalate is a no-op that leaves the
    email for a human, exactly like the n8n «Эскалация: оставить непрочитанным» noOp).
  • Sheet1 Status values: Declined / Bounced / Send Failed / Auto Reply.

Safety / porting deviations (documented, consistent with agency_autoresponder.py):
  • DRY_RUN=true (default) prints drafts + intended writes; nothing is sent/written.
  • Dedup is by SHA256-hashed Message-ID in data/influencer_autoresponder_state.json
    (repo is PUBLIC). The n8n Pricing "Last Msg ID" guard is ALSO reproduced.
  • We do NOT mark messages read over IMAP (PEEK); state dedup replaces markAsRead.
  • For bounce / send_failed the DSN comes FROM mailer-daemon, so we resolve the real
    failed recipient with ec.extract_failed_recipient (as agency does) before flagging
    Sheet1 — the n8n version matched on the daemon address and silently no-op'd.
  • An EMPTY Claude response (transient API failure) is treated as UNDECIDED (retry
    next run, not marked processed), never as an escalation. A NON-empty but
    unparseable response falls back to escalate, exactly like the n8n «Итог AI».

Usage:  python influencer_autoresponder.py
Env:    GMAIL_APP_PW_SERGEY, GMAIL_APP_PW_SERGI, GMAIL_APP_PW_SERHII,
        ANTHROPIC_API_KEY, GOOGLE_CREDENTIALS_JSON, DRY_RUN,
        INFL_SHEET_ID, INFL_SHEET_TAB, INFL_CONTACTS_TAB, LOOKBACK_DAYS, STATE_DIR
"""

import os
import re
import json
import time
from datetime import datetime, timezone

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

# Influencer CRM spreadsheet (from the n8n Google Sheets nodes).
SHEET_ID = os.environ.get(
    "INFL_SHEET_ID", "12IiHIsdibJPRGYNyZfrvdmBDY9OjmsokdmL4GgWg4qQ")
# INFL_SHEET_TAB = the rate-card tab that is written on every dialogue turn.
PRICING_TAB = os.environ.get("INFL_SHEET_TAB", "Pricing")
# The contacts tab (Sheet1) that carries Declined/Bounced/Send Failed/Auto Reply.
CONTACTS_TAB = os.environ.get("INFL_CONTACTS_TAB", "Sheet1")

# Our own mailbox addresses (verbatim OWN list from the n8n «Разбор писем» node) —
# inbound from these is a loop, not a creator.
OWN_ADDRESSES = [
    "sergey.utd@gmail.com",
    "sergi.utd@gmail.com",
    "serhii.smortkin.utd@gmail.com",
]
_OWN_SET = {a.lower() for a in OWN_ADDRESSES}

# Physical mailboxes to scan (only those with an app-password are processed).
# email_common "account" convention: {"user", "password"} with pw from env.
ACCOUNTS = [
    {"user": "sergey.utd@gmail.com",         "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")},
    {"user": "sergi.utd@gmail.com",          "password": os.environ.get("GMAIL_APP_PW_SERGI", "")},
    {"user": "serhii.smortkin.utd@gmail.com", "password": os.environ.get("GMAIL_APP_PW_SERHII", "")},
]

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "data", "influencer_autoresponder_state.json")

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")
# Scan the WHOLE inbox over a wide window (dedup is by hashed Message-ID in state).
LOOKBACK_DAYS = int(os.environ.get("LOOKBACK_DAYS", "60"))

# ONLY the influencer chain is handled here (mirror image of the agency responder,
# which SKIPS this marker). Everything else in the mailbox is ignored.
INFLUENCER_MARKER = re.compile(r"shopify theme review collab", re.I)

# The 17 rate-card data fields, in the exact n8n order (drives collected/merged/JSON).
DATA_KEYS = [
    "price_article", "price_youtube_video", "price_video_mention", "price_shorts",
    "price_social_post", "price_story_mention", "price_newsletter", "packages",
    "usage_rights", "affiliate_revshare", "audience_size", "audience_geo",
    "expected_views", "channel_links", "media_kit", "platform", "notes",
]

# Pricing-sheet column name for each data key (n8n «Обогащение» / «Прайс: upsert»).
PRICING_COL_BY_KEY = {
    "price_article":       "Price Article",
    "price_youtube_video": "Price YouTube Video",
    "price_video_mention": "Price Video Mention",
    "price_shorts":        "Price Shorts/Reels",
    "price_social_post":   "Price Social Post",
    "price_story_mention": "Price Story Mention",
    "price_newsletter":    "Price Newsletter",
    "packages":            "Packages",
    "usage_rights":        "Usage Rights",
    "affiliate_revshare":  "Affiliate/RevShare",
    "audience_size":       "Audience Size",
    "audience_geo":        "Audience Geo",
    "expected_views":      "Expected Views",
    "channel_links":       "Channel Links",
    "media_kit":           "Media Kit",
    "platform":            "Platform",
    "notes":               "Notes",
}

# ── In-run caches (populated once per run; NO per-email Sheets calls) ──
_CONTACTS = {"ws": None, "header": [], "email_to_row": {}}
_PRICING = {"ws": None, "header": [], "email_to_row": {}}
_PENDING_CONTACTS = {}   # {email_lower: {"Status": value}}
_PENDING_PRICING = {}    # {email_lower: {column_name: value}}


# ═══════════════════════════════════════════════════════════════════
#   PROMPTS  (verbatim from the n8n «Собрать запрос Claude» code node)
# ═══════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = (
"You are Sergey, a partnerships manager at UTD Web (utdweb.team), talking with a content creator or influencer who replied to our outreach about a collaboration around UTD, a studio with 25 Shopify themes on the official Shopify Theme Store (https://themes.shopify.com/themes?q=UTD). Instagram: @utd_web_team.\n\n"
"YOUR GOAL: find out which content formats this creator actually offers (their platforms differ: some are YouTube channels, some are blogs only, some are Instagram or TikTok), then collect their rates for EVERY format they offer, from this menu:\n"
"- dedicated article or blog post about UTD\n"
"- dedicated YouTube video about UTD\n"
"- mention or integration of UTD inside one of their regular videos\n"
"- YouTube Shorts / Reels / short vertical video\n"
"- post about UTD on their social networks\n"
"- mention of UTD in stories or inside another post\n"
"- inclusion in their newsletter\n"
"PLUS always try to learn: size of their ACTIVE audience; audience geography; realistic views or traffic our content would get; links to their channels and examples of previous brand integrations; a media kit if they have one; bundle or package offers; the cost of content usage rights or whitelisting; whether they work on affiliate or revenue-share terms.\n\n"
"IMPORTANT: do NOT ask for rates of formats the creator clearly does not offer. A blog-only creator has no YouTube prices. First understand what they do, then dig deeper into what is available. Extract every piece of data they volunteer even if you did not ask.\n\n"
"ALREADY COLLECTED data will be provided with each email. Ask ONLY for what is missing AND applicable. If everything applicable is collected, thank them warmly and say the team will review the options and get back to them shortly. Do not negotiate discounts, do not commit to any purchase, budget, or timeline. If they ask for OUR budget, politely say the budget depends on their formats and rates, and ask for their rate card instead. If they ask questions about UTD, answer briefly: official Shopify Theme Store developer, 25 themes, based in Belgium, links utdweb.team and themes.shopify.com/themes?q=UTD only.\n\n"
"YOUR TASK: read one incoming email and return STRICT JSON:\n"
"{\"category\":\"interested|question|decline|spam|escalate\",\"note\":\"<one short sentence in RUSSIAN for the manager>\",\"reply_body\":\"<reply text or empty>\",\"data\":{\"price_article\":\"\",\"price_youtube_video\":\"\",\"price_video_mention\":\"\",\"price_shorts\":\"\",\"price_social_post\":\"\",\"price_story_mention\":\"\",\"price_newsletter\":\"\",\"packages\":\"\",\"usage_rights\":\"\",\"affiliate_revshare\":\"\",\"audience_size\":\"\",\"audience_geo\":\"\",\"expected_views\":\"\",\"channel_links\":\"\",\"media_kit\":\"\",\"platform\":\"\",\"notes\":\"\"},\"data_complete\":false}\n\n"
"DATA RULES: fill data fields with values extracted from THIS email verbatim (e.g. \"$300\", \"1500 EUR\", \"120k subscribers\", \"30-50k views per video\"). Leave a field as empty string if this email does not mention it. Use notes for anything relevant that does not fit (bundles, requirements, free product requests). Set data_complete=true only when, combining already-collected data and this email, you have prices for ALL formats this creator offers (judge by their platform and their own words) AND audience_size AND expected_views. A blog-only creator with just an article price and audience data can be complete.\n\n"
"CATEGORY RULES:\n"
"- interested / question: continue the dialogue per the goal. reply_body required.\n"
"- decline: not interested or asks to stop. reply_body empty.\n"
"- spam: unrelated or automated mail. reply_body empty.\n"
"- escalate: contracts, calls with specific times, legal or payment questions, aggressive negotiation, whitelabel or revenue-share partnership proposals, paid research platforms, or anything you cannot answer from the facts above. reply_body empty.\n\n"
"REPLY RULES: reply in the language of the incoming email. Under 150 words. Friendly, human, creator-outreach tone, no corporate stiffness, no hype words. Never use em dashes. When asking for rates, ask as a short list. End with:\n"
"Best regards,\n"
"Sergey\n"
"UTD Web | utdweb.team\n\n"
"Output ONLY the JSON object."
)


def build_user_prompt(msg, contact_name, collected):
    """Verbatim reproduction of the n8n user-message assembly.
    collected is serialized with JS JSON.stringify semantics (no spaces)."""
    col_json = json.dumps(collected, ensure_ascii=False, separators=(",", ":"))
    return (
        "Incoming email:\n"
        "From: " + (msg.get("from", "") or "") + "\n"
        "Subject: " + (msg.get("subject", "") or "") + "\n"
        "Creator (from CRM): " + (contact_name or "(unknown)") + "\n\n"
        "ALREADY COLLECTED: " + col_json + "\n\n"
        "Body:\n" + (msg.get("body", "") or "") + "\n\n"
        "Classify, extract data, draft the reply. Output ONLY the JSON object."
    )


# ═══════════════════════════════════════════════════════════════════
#   AI result parsing  (ported from the n8n «Итог AI» code node)
# ═══════════════════════════════════════════════════════════════════

def parse_ai_result(text, collected):
    """Parse Claude's strict JSON into a routing decision + merged rate card.

    Returns:
      • None  → UNDECIDED (empty API response): leave the email for the next run,
                do NOT mark processed. (Safety layer over the n8n behaviour.)
      • dict  → a decision. A NON-empty but unparseable response falls back to
                category 'escalate' with note 'AI не смог разобрать письмо',
                exactly like the n8n «Итог AI» defaults.
    """
    if not text or not text.strip():
        return None  # transient API failure → retry next run, never escalate

    # n8n «Итог AI» defaults.
    cat = "escalate"
    reply = ""
    note = "AI не смог разобрать письмо"
    data = {}
    complete = False
    try:
        m = re.search(r"\{[\s\S]*\}", text)
        p = json.loads(m.group(0))
        if p.get("category") in ("interested", "question", "decline", "spam", "escalate"):
            cat = p["category"]
        reply = (p.get("reply_body") or "").strip()
        note = (p.get("note") or "").strip() or note
        data = p.get("data") or {}
        complete = bool(p.get("data_complete"))
    except Exception:
        pass

    # interested/question with no drafted reply → escalate to a human.
    if cat in ("interested", "question") and not reply:
        cat = "escalate"

    route = "respond" if cat in ("interested", "question") \
        else ("ignore" if cat == "spam" else cat)  # decline | escalate

    # Merge: new value overrides the old one only when non-empty.
    merged = {}
    for k in DATA_KEYS:
        nv = str((data.get(k) if isinstance(data, dict) else "") or "").strip()
        merged[k] = nv or str(collected.get(k, "") or "")

    pricing_status = "Declined" if route == "decline" \
        else ("Escalated" if route == "escalate"
              else ("Data Complete" if complete else "Negotiating"))

    return {"category": route, "ai_category": cat, "note": note, "reply_body": reply,
            "merged": merged, "data_complete": complete,
            "pricing_status_new": pricing_status}


def non_ai_result(pre_category):
    """Ported from the n8n «Без AI — категория» code node."""
    mapping = {"bounce": "bounce", "send_failed": "send_failed", "auto_reply": "auto_reply"}
    route = mapping.get(pre_category, "ignore")
    note = {
        "bounce": "Отбойник (мёртвый адрес)",
        "send_failed": "Письмо не ушло (лимит/блок Gmail)",
        "auto_reply": "Стандартный автоответ (тикет/OOO)",
    }.get(pre_category, "Служебное")
    return {"category": route, "ai_category": pre_category, "note": note,
            "reply_body": "", "merged": None, "data_complete": False,
            "pricing_status_new": ""}


# ═══════════════════════════════════════════════════════════════════
#   CRM matching + enrichment  (ported from «Контакты» + «Обогащение»)
# ═══════════════════════════════════════════════════════════════════

def build_contact_index(contact_rows):
    """Index Sheet1 contacts by Thread ID and by Email (name = Name|Channel|Company)."""
    by_thread, by_email = {}, {}
    for r in contact_rows:
        info = {
            "email": str(r.get("Email", "") or "").strip().lower(),
            "name": str(r.get("Name", "") or r.get("Channel", "") or r.get("Company Name", "") or "").strip(),
        }
        t = str(r.get("Thread ID", "") or "").strip()
        if t:
            by_thread[t] = info
        if info["email"]:
            by_email[info["email"]] = info
    return by_thread, by_email


def build_pricing_index(pricing_rows):
    """Index Pricing rows by Email (skipping the '(init)' placeholder)."""
    by_email = {}
    for p in pricing_rows:
        e = str(p.get("Email", "") or "").strip().lower()
        if e and e != "(init)":
            by_email[e] = p
    return by_email


def collected_from_pricing(prow):
    """Pull the 17 already-collected data fields out of a Pricing row dict."""
    prow = prow or {}
    return {k: str(prow.get(PRICING_COL_BY_KEY[k], "") or "") for k in DATA_KEYS}


def match_contact(msg, by_thread, by_email):
    """Prefer Thread ID match (Gmail thread id), then sender email — as in n8n.
    IMAP exposes X-GM-THRID as a decimal; the sheet may store the Gmail API hex
    thread id, so we try both representations before falling back to email."""
    thrid = str(msg.get("gm_thrid", "") or "").strip()
    if thrid:
        if thrid in by_thread:
            return by_thread[thrid]
        try:
            hx = format(int(thrid), "x")
            if hx in by_thread:
                return by_thread[hx]
        except Exception:
            pass
    return by_email.get(msg.get("from_email", ""), None)


# ═══════════════════════════════════════════════════════════════════
#   Pre-classification  (mirror of the n8n «Разбор писем» category logic)
# ═══════════════════════════════════════════════════════════════════

def pre_classify(msg):
    """Return bounce | send_failed | auto_reply | own | human.

    Uses ec.classify_incoming for bounce/send_failed/auto_reply/human, then adds
    the n8n 'own' distinction (own-address senders → route 'ignore', not a sheet
    write). Bounce still takes priority over own, matching the n8n ordering."""
    pre = ec.classify_incoming(msg, OWN_ADDRESSES)
    if pre not in ("bounce", "send_failed"):
        sender = (msg.get("from_email", "") or "").lower()
        if sender in _OWN_SET:
            return "own"
    return pre


# ═══════════════════════════════════════════════════════════════════
#   Actions (send + sheet), guarded by DRY_RUN
# ═══════════════════════════════════════════════════════════════════

def _reply_subject(subject):
    s = subject or ""
    return s if s.lower().startswith("re:") else "Re: " + s


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
    """Send the Claude-drafted reply in-thread (n8n «Автоответ», emailType text)."""
    subject = _reply_subject(msg["subject"])
    if DRY_RUN:
        _print_draft("reply", account, msg["from_email"], subject, decision["reply_body"])
        return
    ec.send_email(account, msg["from_email"], subject, decision["reply_body"],
                  in_reply_to=msg["message_id"], references=msg["references"])


def enqueue_contact_status(email, status):
    """Queue a Sheet1 Status update (only existing rows are updated, like n8n 'update')."""
    if not email or not status:
        return
    key = str(email).strip().lower()
    if key not in _CONTACTS["email_to_row"]:
        print(f"  [Sheet1] {key} not found → skipped (no row to update)")
        return
    _PENDING_CONTACTS.setdefault(key, {})["Status"] = status
    print(f"  [Sheet1] queued Status='{status}' for {key}")


def enqueue_pricing_upsert(contact_email, contact_name, decision, msg):
    """Queue a full Pricing appendOrUpdate row (n8n «Прайс: upsert» / «Прайс: отказ»)."""
    email = str(contact_email or "").strip().lower()
    if not email:
        print("  [Pricing] no contact email → skipped")
        return
    merged = decision["merged"] or {}
    cols = {
        "Email": contact_email,
        "Name": contact_name,
        "Platform": merged.get("platform", ""),
        "Contact Status": decision["pricing_status_new"],
        "Last Msg ID": msg["message_id"],
    }
    for k in DATA_KEYS:
        cols[PRICING_COL_BY_KEY[k]] = merged.get(k, "")
    _PENDING_PRICING[email] = cols  # last write wins for this email in a run
    print(f"  [Pricing] queued Contact Status='{decision['pricing_status_new']}' for {email}")


def flush_contacts():
    """Write all queued Sheet1 Status updates in ONE batchUpdate call."""
    if not _PENDING_CONTACTS:
        print("\n[Sheet1] no updates to flush.")
        return
    header = _CONTACTS["header"]
    cell_updates = []
    for email, cols in _PENDING_CONTACTS.items():
        row = _CONTACTS["email_to_row"].get(email)
        if not row:
            continue
        for col, val in cols.items():
            if col not in header:
                continue
            a1 = ec.gspread_a1(row, header.index(col) + 1)
            cell_updates.append({"range": a1, "values": [[val]]})
    if DRY_RUN:
        print(f"\n[Sheet1] DRY_RUN — would batch-write {len(cell_updates)} cells "
              f"across {len(_PENDING_CONTACTS)} contacts:")
        for email, cols in _PENDING_CONTACTS.items():
            print(f"    {email}: {cols}")
        return
    if not _CONTACTS["ws"]:
        print("\n[Sheet1] no worksheet handle — cannot flush.")
        return
    try:
        n = ec.batch_update_cells(_CONTACTS["ws"], cell_updates)
        print(f"\n[Sheet1] batch-wrote {n} cells across {len(_PENDING_CONTACTS)} contacts.")
    except Exception as e:
        print(f"\n⚠️  [Sheet1] batch update failed: {e}")


def flush_pricing():
    """Write all queued Pricing rows: existing → batchUpdate cells, new → append_rows."""
    if not _PENDING_PRICING:
        print("\n[Pricing] no updates to flush.")
        return
    header = _PRICING["header"]
    cell_updates, new_rows, new_emails = [], [], []
    for email, cols in _PENDING_PRICING.items():
        row = _PRICING["email_to_row"].get(email)
        if row:
            for col, val in cols.items():
                if col not in header:
                    continue
                a1 = ec.gspread_a1(row, header.index(col) + 1)
                cell_updates.append({"range": a1, "values": [[val]]})
        else:
            new_rows.append([cols.get(h, "") for h in header])
            new_emails.append(email)
    if DRY_RUN:
        print(f"\n[Pricing] DRY_RUN — would UPDATE {len(cell_updates)} cells and "
              f"APPEND {len(new_rows)} new rows:")
        for email, cols in _PENDING_PRICING.items():
            tag = "update" if email in _PRICING["email_to_row"] else "append"
            print(f"    [{tag}] {email}: {cols}")
        return
    if not _PRICING["ws"]:
        print("\n[Pricing] no worksheet handle — cannot flush.")
        return
    try:
        if cell_updates:
            ec.batch_update_cells(_PRICING["ws"], cell_updates)
        if new_rows:
            ec.with_sheets_backoff(
                lambda: _PRICING["ws"].append_rows(new_rows, value_input_option="USER_ENTERED"))
        print(f"\n[Pricing] wrote {len(cell_updates)} cells + appended {len(new_rows)} rows.")
    except Exception as e:
        print(f"\n⚠️  [Pricing] flush failed: {e}")


# ═══════════════════════════════════════════════════════════════════
#   Per-message processing
# ═══════════════════════════════════════════════════════════════════

def process_message(account, msg, by_thread, by_email, pricing_by_email, state, stats):
    mid = msg.get("message_id", "")

    # ONLY the influencer chain — everything else is ignored (mirror of agency).
    if not INFLUENCER_MARKER.search(msg.get("subject", "") or ""):
        return

    # Dedup: already processed in a previous run (hashed Message-ID)?
    if ec.is_processed(state, mid):
        return

    # Enrichment (runs for ALL messages, like the n8n «Обогащение» node).
    contact = match_contact(msg, by_thread, by_email)
    contact_email = (contact["email"] if contact else "") or msg.get("from_email", "") \
        or ec.extract_failed_recipient(msg, OWN_ADDRESSES)
    contact_name = contact["name"] if contact else ""
    prow = pricing_by_email.get(str(contact_email).strip().lower(), {})

    # Pricing "Last Msg ID" guard (faithful to the n8n «Обогащение» dedup).
    if prow and str(prow.get("Last Msg ID", "")).strip() == mid and mid:
        if not DRY_RUN:
            ec.mark_processed(state, mid)
        return

    collected = collected_from_pricing(prow)

    # Route: non-AI (bounce/send_failed/auto_reply/own) vs AI (human).
    pre = pre_classify(msg)
    if pre != "human":
        decision = non_ai_result(pre)
    else:
        user = build_user_prompt(msg, contact_name, collected)
        time.sleep(0.7)  # gentle pacing to avoid 429 on big batches
        ai_text = ec.call_claude(SYSTEM_PROMPT, user, model="claude-sonnet-5",
                                 max_tokens=1400)
        decision = parse_ai_result(ai_text, collected)
        if decision is None:
            print(f"\n· {account['user']} ← {msg['from_email']} | {msg['subject'][:60]!r}")
            print("  UNDECIDED (empty AI response) → left for next run, not marked processed.")
            return

    route = decision["category"]
    stats[route] = stats.get(route, 0) + 1
    print(f"\n· {account['user']} ← {msg['from_email']} | {msg['subject'][:60]!r}")
    print(f"  route={route} status→{decision['pricing_status_new'] or '-'} | {decision['note']}")

    if route == "respond":
        # interested/question → send reply + Pricing upsert (Negotiating/Data Complete).
        do_reply(account, msg, decision)
        enqueue_pricing_upsert(contact_email, contact_name, decision, msg)
    elif route == "decline":
        # Sheet1 Status='Declined' + Pricing upsert (Contact Status='Declined'). No reply.
        enqueue_contact_status(contact_email, "Declined")
        enqueue_pricing_upsert(contact_email, contact_name, decision, msg)
    elif route in ("bounce", "send_failed"):
        # DSN is FROM mailer-daemon — resolve the real failed recipient, flag Sheet1.
        rcpt = ec.extract_failed_recipient(msg, OWN_ADDRESSES) or contact_email
        status = {"bounce": "Bounced", "send_failed": "Send Failed"}[route]
        if rcpt:
            print(f"  failed recipient = {rcpt} → Sheet1 Status '{status}'")
            enqueue_contact_status(rcpt, status)
        else:
            print("  failed recipient not found → Sheet1 left untouched.")
    elif route == "auto_reply":
        enqueue_contact_status(contact_email, "Auto Reply")
    elif route == "escalate":
        # n8n «Эскалация: оставить непрочитанным» noOp — no reply, no sheet write.
        print("  ESCALATE → left for a human (no reply, no sheet write).")
    # 'own' / 'ignore' → nothing (n8n «Прочитано: игнор»).

    if not DRY_RUN:
        ec.mark_processed(state, mid)


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def _load_tab(sheet_id, tab):
    """Open + read a worksheet ONCE, returning (rows, ws, header, email_to_row)."""
    ws = ec.open_worksheet(sheet_id, tab)
    rows = ec.read_rows_ws(ws)
    header = list(rows[0].keys()) if rows else ws.row_values(1)
    e2r = {}
    for i, r in enumerate(rows, start=2):  # header is row 1, data starts at row 2
        em = str(r.get("Email", "") or "").strip().lower()
        if em and em not in e2r:
            e2r[em] = i
    return rows, ws, header, e2r


def run_once():
    print(f"=== UTD influencer autoresponder | DRY_RUN={DRY_RUN} | "
          f"lookback {LOOKBACK_DAYS}d | {datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)

    # Reset in-run caches / pending writes.
    _PENDING_CONTACTS.clear()
    _PENDING_PRICING.clear()
    for cache in (_CONTACTS, _PRICING):
        cache["ws"] = None
        cache["header"] = []
        cache["email_to_row"] = {}

    # Read both CRM tabs ONCE (429-safe).
    contact_rows, pricing_rows = [], []
    try:
        contact_rows, _CONTACTS["ws"], _CONTACTS["header"], _CONTACTS["email_to_row"] = \
            _load_tab(SHEET_ID, CONTACTS_TAB)
    except Exception as e:
        print(f"⚠️  Could not read contacts tab '{CONTACTS_TAB}': {e}")
    try:
        pricing_rows, _PRICING["ws"], _PRICING["header"], _PRICING["email_to_row"] = \
            _load_tab(SHEET_ID, PRICING_TAB)
    except Exception as e:
        print(f"⚠️  Could not read pricing tab '{PRICING_TAB}': {e}")

    by_thread, by_email = build_contact_index(contact_rows)
    pricing_by_email = build_pricing_index(pricing_rows)
    print(f"CRM: {len(contact_rows)} contacts ({len(by_email)} emails, "
          f"{len(by_thread)} threads), {len(pricing_by_email)} pricing rows.")

    stats = {}
    for account in ACCOUNTS:
        if not account["password"]:
            print(f"⚠️  No app-password for {account['user']} — skipping mailbox.")
            continue
        try:
            msgs = ec.fetch_inbox(account, since_days=LOOKBACK_DAYS, unseen_only=False)
        except Exception as e:
            print(f"⚠️  IMAP error for {account['user']}: {e}")
            continue
        print(f"\n>>> {account['user']}: {len(msgs)} messages in last {LOOKBACK_DAYS}d")
        for msg in msgs:
            try:
                process_message(account, msg, by_thread, by_email,
                                 pricing_by_email, state, stats)
            except Exception as e:
                print(f"  !! error on message {msg.get('message_id','?')}: {e}")

    # Flush queued CRM updates (Sheet1 statuses + Pricing upserts).
    flush_contacts()
    flush_pricing()

    if not DRY_RUN:
        ec.save_state(STATE_FILE, state)
    print(f"\n=== done. routes: {stats} ===")
    return {"parser": "influencer_autoresponder", "routes": stats,
            "dry_run": DRY_RUN,
            "contacts_writes_queued": len(_PENDING_CONTACTS),
            "pricing_writes_queued": len(_PENDING_PRICING)}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
