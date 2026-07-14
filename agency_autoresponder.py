#!/usr/bin/env python3
"""
agency_autoresponder.py — UTD "Agency Partner Program" B2B inbound autoresponder.

Port of the n8n prototype (build_autopilot.py) to plain Python for GitHub Actions.
Reads two Gmail mailboxes over IMAP, classifies each new inbound email, drives the
B2B partner funnel with Claude, and replies in-thread over SMTP (app-passwords).

Funnel (Status column drives the stage):
  ''/Sent/Follow-up Sent/Replied  → qualify          → Status "Qualifying"
  Qualifying (>=3 of 4 answered)  → memo (attach)     → Status "Memo Sent"
  Memo Sent                       → agreement_ready    → Status "Agreement Sent"
                                    (attach the Agreement, push to sign)
  Agreement Offered               → agreement_ready / agreement_offer
  Agreement Sent + signed doc     → signed             → Status "Agreement Signed"
  Agreement Sent / Signed + Qs    → info (answer contract questions any time)

New rules (Valeriy, 2026-07-08):
  (a) answer ANY contract questions both BEFORE and AFTER signing (info/agreement stages).
  (b) after signed: send the partner a thank-you + next steps, AND email a deal review
      (with the signed doc attached) to the review recipients.

Safety:
  • DRY_RUN=true (default) prints drafts instead of sending / writing to the sheet.
  • Processed Message-IDs are SHA256-hashed into data/agency_autoresponder_state.json
    (repo is PUBLIC — no raw emails/addresses committed).
  • Influencer-chain emails ("shopify theme review collab") are skipped entirely.

Usage:  python agency_autoresponder.py
Env:    GMAIL_APP_PW_SERGEY, GMAIL_APP_PW_SERGE, ANTHROPIC_API_KEY,
        GOOGLE_CREDENTIALS_JSON, DRY_RUN, DOC_MEMO_PATH, DOC_CONTRACT_PATH,
        LOOKBACK_DAYS, REMINDER_DAYS, MAX_REMINDERS, STATE_DIR

Stage is inferred by Claude from the FULL thread history (fetched via Gmail
X-GM-THRID), not just the CRM Status column, which can be stale.
"""

import os
import re
import time
import json
import tempfile
from datetime import datetime, timezone

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get("AGENCY_SHEET_ID", os.environ.get("B2B_SHEET_ID", ""))
SHEET_TAB = os.environ.get("AGENCY_SHEET_TAB", "IT Companies — Emails")

# Deal-review recipients after a signed agreement (Valeriy 2026-07-06/07-08).
REVIEW_TO = [a for a in (
    os.environ.get("UTD_MAIL_SERHII", ""),
    os.environ.get("UTD_MAIL_DENYS", ""),
) if a]

# Our own mailbox addresses — inbound from these is a loop, not a prospect.
OWN_ADDRESSES = [a for a in (
    os.environ.get("UTD_MAIL_SERGEY", ""),
    os.environ.get("UTD_MAIL_SERGE", ""),
    os.environ.get("UTD_MAIL_SERHII", ""),
) if a]

# The two physical mailboxes to process.
ACCOUNTS = [a for a in (
    {"user": os.environ.get("UTD_MAIL_SERGEY", ""), "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")},
    {"user": os.environ.get("UTD_MAIL_SERGE", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGE", "")},
) if a["user"]]

_STATE_DIR = os.environ.get("STATE_DIR", "./data")
STATE_FILE = os.path.join(_STATE_DIR, "agency_autoresponder_state.json")

# Documents (committed in docs/; overridable via env for local paths).
_REPO = os.path.dirname(os.path.abspath(__file__))
DOC_MEMO_PATH = os.environ.get(
    "DOC_MEMO_PATH", os.path.join(_REPO, "docs", "UTD_Agency_Partner_Program_Overview.docx"))
DOC_CONTRACT_PATH = os.environ.get(
    "DOC_CONTRACT_PATH", os.path.join(_REPO, "docs", "UTD_Agency_Partner_Program_Agreement.docx"))

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")
# Scan the WHOLE inbox over a wide window (not just unread) — dedup is by hashed
# Message-ID in state, so already-read replies we missed before are still caught.
LOOKBACK_DAYS = int(os.environ.get("LOOKBACK_DAYS", "60"))
# Silence-based follow-up reminders.
REMINDER_DAYS = int(os.environ.get("REMINDER_DAYS", "3"))
MAX_REMINDERS = int(os.environ.get("MAX_REMINDERS", "2"))
# Stages that receive silence reminders (contact is mid-dialogue).
REMINDER_STAGES = {"Qualifying", "Memo Sent", "Agreement Sent"}

INFLUENCER_MARKER = re.compile(r"shopify theme review collab", re.I)

# stage → CRM Status
STATUS_BY_STAGE = {
    "qualify": "Qualifying",
    "memo": "Memo Sent",
    "agreement_offer": "Agreement Offered",
    "agreement_ready": "Agreement Sent",
    "signed": "Agreement Signed",
}

# ── In-run caches (populated once per run; NO per-email Sheets/IMAP calls) ──
# _SHEET: the worksheet handle + header + email→row map, read ONCE at startup.
# _PENDING: {email_lower: {column: value}} accumulated during the run, then
#           written in ONE values.batchUpdate call at the end (429-safe).
# _THREAD_CACHE: {(account_user, gm_thrid): history_text} so each Gmail thread
#           is fetched at most once per run.
_SHEET = {"ws": None, "header": [], "email_to_row": {}}
_PENDING = {}
_THREAD_CACHE = {}


# ═══════════════════════════════════════════════════════════════════
#   KNOWLEDGE BASE + PROMPTS  (verbatim from the n8n prototype)
# ═══════════════════════════════════════════════════════════════════

KB = (
"UTD AGENCY PARTNER PROGRAM - VERIFIED FACTS (the ONLY source of truth):\n"
"- Program owner: UTD Web, a company registered in Belgium, developer of official Shopify themes. Sites: https://utdweb.team and https://themes.shopify.com/themes?q=UTD\n"
"- Invitation-only: UTD individually selects agencies based on professional reputation and technical expertise. Not publicly available.\n"
"- How it works: the participant recommends, purchases and implements UTD Shopify themes for their own clients as part of their web development services. Theme purchases MUST be made through the official Shopify Theme Store using the end client's own Shopify account. Purchases through any other channel are not eligible for commission.\n"
"- Commission (monthly, % of gross sale price, only confirmed non-refunded purchases): 1 to 10 confirmed sales in a calendar month = 10% per theme; 11 to 20 = 13%; 21 or more = 15%. Once a volume threshold is reached, that rate applies to ALL confirmed eligible sales of that month.\n"
"- Payout timing: commission is released after a sixty (60) day holdback from each purchase date (Shopify's standard refund window). If a refund happens within 60 days, that commission is not paid. Late refunds after payout are offset against future commissions. Payment transfer fees and currency conversion costs are borne by the participant.\n"
"- Before any commission is processed (at the close of the first active month), three steps are required: 1) signing a Non-Disclosure Agreement (NDA), 2) signing a Cooperation Agreement, 3) submitting business details (legal company name, registered address, tax ID where applicable, payment account).\n"
"- Monthly reporting: by the 3rd calendar day of the following month the participant submits a report listing Shopify store identifiers (store URLs or account references) for client purchases of that month. UTD cross-references this with its own Shopify Theme Store sales data; only purchases confirmed in both records are eligible. Late reports may be deferred to the next payment cycle.\n"
"- Fraud monitoring: coordinated purchase-and-refund cycles or purchases from accounts with no commercial activity can lead to cancelled commissions and removal from the program, with written notice and an opportunity to respond.\n"
"- Additional benefits: access to UTD's senior technical support channel; a Premium Support team is available (cost shared on request) and participants get a 10% discount on any premium support package; early access email summaries of new theme versions before public release (on request); ability to submit feature suggestions to the product team; additional services on a per-project basis (custom development, content production, digital marketing, SEO).\n"
"- Co-marketing: UTD does not promote referral partners directly, but partners are welcome to use UTD materials and brand assets to promote the partnership on their own channels and owned resources.\n"
"- Data protection: GDPR, the Belgian Act of 30 July 2018 and CCPA/CPRA apply where relevant. All client data exchanged is confidential and an NDA is signed before any data exchange.\n"
"- Legal nature: the informational overview document (the memo) is not a contract or offer. Formal participation is established only through the separate Cooperation Agreement and NDA. UTD may revise program terms with advance notice. Either party may withdraw in writing; sales confirmed before withdrawal are still paid under the holdback rules. Participation creates no employment, agency or partnership relationship.\n"
)

QUALIFY_QUESTIONS = (
"QUALIFICATION QUESTIONS (must be answered by the prospect BEFORE we send the memo):\n"
"1. Which themes are you currently using or reselling for your clients?\n"
"2. Roughly how many themes do you purchase for clients per month?\n"
"3. Do you already work with other theme providers?\n"
"4. How do you typically work with clients: do you purchase the theme on their behalf and then set up the store, or is your model different?\n"
)

SYSTEM_PROMPT = (
"You are Sergey, a partnership manager at UTD Web (utdweb.team), handling inbound replies in the UTD Agency Partner Program outreach mailbox.\n\n"
+ KB + "\n" + QUALIFY_QUESTIONS +
"\nDETERMINE THE REAL FUNNEL STAGE FROM THE THREAD HISTORY, not only from the CRM status "
"(the CRM status can be stale or wrong). You are given the full conversation history for this "
"thread (both our sent messages and the prospect's replies, with attachments). Read it and infer "
"where we actually are:\n"
"- HARD RULE (memo is sent at most once): if the thread ALREADY contains a message from us with a "
"memo/overview attachment AND the prospect later replied with any substantive text, the stage is "
"ALWAYS \"agreement_ready\" (attach the agreement and push to sign). NEVER choose \"memo\" again in "
"that case, even if the CRM status still says 'Sent' or 'Qualifying'. Only choose \"memo\" when the "
"thread does NOT already contain a memo attachment from us.\n"
"- If the history shows the program memo/overview was ALREADY sent (by us, manually or automatically) "
"and the prospect has since answered the qualification questions (even partially, at least 3 of 4), "
"the real stage is \"agreement_ready\": attach the agreement and push to sign. Do NOT re-send the "
"qualification questions in this case, even if the CRM status still says 'Sent'.\n"
"- If the history shows the agreement was ALREADY sent, do NOT send it again: answer their questions "
"and keep pushing toward the signature (stage \"agreement_ready\" without re-introducing it, or \"info\").\n"
"- If the history shows a signed document was already returned, treat further messages as \"info\".\n"
"- Only fall back to the CRM status below when the thread history is empty or inconclusive.\n\n"
"CONVERSATION FUNNEL (fallback when thread history is inconclusive; driven by contact_status):\n"
"- Statuses 'Sent', 'Follow-up Sent', 'Replied' or empty = PRE-QUALIFICATION. If the prospect is interested or asks questions, reply at stage \"qualify\".\n"
"- Status 'Qualifying' = we already asked the qualification questions. If the email answers most of them (at least 3 of 4, even briefly), reply at stage \"memo\". If they have not answered yet, stay at stage \"qualify\" and politely repeat what we need.\n"
"- Status 'Memo Sent' = they received the detailed memo and replied. Move them forward: reply at stage \"agreement_ready\" (the Agency Partner Agreement will be ATTACHED to your reply). Even if their reply has no clear intent, this is the moment to send the agreement and invite them to sign. Answer any questions with full detail, then present the agreement.\n"
"- Status 'Agreement Offered' = we already explained the agreement step. If they confirm they are ready to sign or ask to receive the agreement, reply at stage \"agreement_ready\". If they still have questions, stay at stage \"agreement_offer\" and answer them.\n"
"- Status 'Agreement Ready' or 'Agreement Sent' = the official agreement is being processed. If the incoming email HAS ATTACHMENTS and the text indicates they are returning the signed document, reply at stage \"signed\". Otherwise answer their questions at stage \"info\".\n"
"- Status 'Agreement Signed' = fully onboarded. Reply at stage \"info\" with full detail.\n\n"
"CONTRACT QUESTIONS RULE: you must fully answer ANY question about the agreement or program terms at ANY time, both BEFORE and AFTER signing, using the verified facts above (commission tiers, holdback, reporting, withdrawal, NDA). Never refuse a contract question because it is 'too early' or 'already signed'.\n\n"
"SELLING STANCE (mandatory): every reply must move the deal one step forward: toward the memo, then the agreement, then the signature. Never just inform. Close every reply with the next step.\n"
"SELLING ARGUMENTS (all true; concrete numbers sell better than adjectives, use them; pick the 1-2 arguments that fit what this prospect said, do not dump all of them into one email):\n"
"- Money on every project: commission starts at 10% and grows to 15% with monthly volume. Our flagship theme Impression costs $340, so ten client stores on it return the agency $340 to 510.\n"
"- Time: building a theme from scratch takes 100+ hours of one developer. With our theme the site is ready in a couple of days. That saves roughly 80-100 hours per project; at $40-60 per hour that is about $3,200-6,000 saved on one project.\n"
"- Support: we help install and set up the theme, we help build client sites, we explain what is inside the theme, we answer the agency's developers fast, and we fix issues ourselves.\n"
"CRITICAL FRAMING (never break it): never say or imply 'bring your clients to us', 'refer your clients', or 'pass us the project'. The client buys the theme themselves through the official Shopify Theme Store. The agency KEEPS the client and the project and earns the commission on top. The agency builds and sets up the client's store as usual, and we give our best support for free.\n\n"
"STAGE RULES:\n"
"- qualify: thank them for the reply and answer their questions. Headline numbers are fine here: commission starts at 10% of each theme's sale price and grows to 15% with monthly volume; payouts are monthly, accounting for Shopify's refund window; purchases go through the official Shopify Theme Store using the client's own account; we do not promote partners directly but they may use our materials and brand assets on their own channels. Keep the full tier table, the exact holdback length and the reporting details for the memo. Then ask the four qualification questions as a short list, and close by saying that once we have their answers we will send the detailed program memo.\n"
"- memo: thank them for the answers, refer to one specific detail from their reply, say the detailed program overview (memo) is attached to this email and covers commission tiers, payouts, reporting and onboarding, and invite any questions. Do not paste the full terms into the email body. Close by naming the next step: once they have looked at the memo, we send the agreement to sign.\n"
"- agreement_offer: answer their questions precisely using the facts above (exact numbers allowed now). Then push toward signing as the next step.\n"
"- agreement_ready: the Agency Partner Agreement is ATTACHED to this email (do NOT say it will come separately; it is attached now). Ask them to review, sign, and return it, mentioning they only need to sign their side (UTD does not countersign). Argue briefly why signing is worth it, using the SELLING ARGUMENTS above tailored to what they said (commission money, hours saved, our support). Signing is low commitment: no fee, no exclusivity, they can withdraw anytime in writing. Ask them to fill legal company name and registered address in the signature block. Keep it encouraging, not pushy.\n"
"- signed: thank them for returning the signed document, welcome them to the UTD Agency Partner Program, confirm receipt and that their participation is now finalised, and outline what happens next: their first active month starts now, they report client store identifiers by the 3rd calendar day of the following month, and commission is released after the sixty day holdback. Answer any open questions.\n"
"- info: answer their questions precisely using the facts above, including exact numbers, and still end with a small next step when one exists.\n\n"
"YOUR TASK: read one incoming email and return STRICT JSON:\n"
'{"category":"interested|question|decline|spam|escalate","stage":"qualify|memo|agreement_offer|agreement_ready|signed|info","note":"<one short sentence in RUSSIAN summarising the email for the manager>","reply_body":"<full reply text, or empty string>"}\n\n'
"IMPORTANT: choose stage \"signed\" ONLY when the email has attachments (has_attachments is true). A promise to sign without an attached document is NOT \"signed\".\n\n"
"CATEGORY RULES:\n"
"- interested: they want to join, learn more, or continue the conversation. Reply per the stage rules.\n"
"- question: they ask questions about the program. Reply per the stage rules.\n"
"- decline: not interested or asks to stop emailing. reply_body must be empty.\n"
"- spam: unrelated marketing, newsletters, automated notifications. reply_body must be empty.\n"
"- escalate: use whenever a correct answer needs a human: requests for custom commission terms, contract or NDA documents beyond our standard agreement, scheduling a call at a specific time, payment problems, complaints, legal interpretation, press, or a question that the facts above cannot answer. reply_body must be empty.\n\n"
"REPLY RULES (for interested and question):\n"
"- Use ONLY the facts above. Never invent facts, numbers, dates, names or prices. Only links allowed: https://utdweb.team and https://themes.shopify.com/themes?q=UTD\n"
"- Reply in the LANGUAGE of the incoming email.\n"
"- SIMPLE ENGLISH for non-native readers (and the same simple wording in any other language): common everyday words, and write in LONG, flowing, simple sentences that go straight to the point (never short choppy ones) — real people write long simple sentences, not staccato fragments. No idioms, no slang, no fancy phrases ('caught my eye', 'worth a look' and anything similar are forbidden). If a 12-year-old would not understand a sentence, rewrite it.\n"
"- Write like a normal person typing an email by hand. If a sentence reads like AI or a script, rewrite it. Zero filler, maximum concreteness: numbers, facts, the next step.\n"
"- The reply is as long as it needs to be to fully answer and move things forward, no longer.\n"
"- FORMAT (mandatory): line 1 is a greeting; then a blank line; then the body grouped by meaning into a few paragraphs with a blank line between them, each paragraph written as LONG, flowing, simple sentences that get straight to the point, never short choppy ones. Then a blank line, the farewell and the signature.\n"
"- If the thread already has earlier messages from us, the first sentence after the greeting must naturally refer to the earlier exchange. Add only NEW substance; never repeat what was already said (use the thread history).\n"
"- Never offer or suggest a call or meeting. Everything is handled by email; you may offer help by email ('reply and I'll walk you through it'). If THEY push for a call at a specific time, that is escalate per the category rules.\n"
"- Never use em dashes. Never use the words: exclusive, exciting, game-changer, handpicked, curated, unique opportunity.\n"
"- Do not promise anything beyond the facts.\n"
"- End the reply with exactly:\nBest regards,\nSergey\nUTD Web | utdweb.team\n\n"
"Output ONLY the JSON object, no markdown fences, no commentary."
)

REVIEW_SYSTEM = (
"Ты аналитик отдела партнёрств UTD Web. Партнёр подписал Cooperation Agreement по партнёрской "
"программе (Shopify-темы). Напиши РЕВЬЮ сделки на русском для команды: 1) кто партнёр (компания, "
"контакт); 2) ход переговоров по этапам (квалификация, мемо, соглашение); 3) ключевые ответы "
"партнёра на квалификационные вопросы (какие темы использует, объём покупок тем в месяц, работает "
"ли с другими вендорами, модель работы с клиентами); 4) выводы: перспективность партнёра, ожидаемый "
"объём, риски; 5) следующие шаги по программе (первый активный месяц, отчёт до 3-го числа, holdback "
"60 дней). Пиши развёрнутыми предложениями, без длинных тире, без воды. Только текст ревью."
)


# ═══════════════════════════════════════════════════════════════════
#   AI result parsing  (ported from AI_RESULT_JS)
# ═══════════════════════════════════════════════════════════════════

def _clean_reply(text):
    """Canon guard: no em/en dashes in an outgoing reply. Digit ranges keep a
    plain hyphen (10-15%), any other dash becomes a comma pause."""
    t = (text or "").strip()
    if not t:
        return t
    t = re.sub(r"(?<=\d)\s*[—–]\s*(?=\d)", "-", t)
    t = re.sub(r"\s*[—–]\s*", ", ", t)
    return t


def parse_ai_result(text, has_attachments, prev_status):
    """Parse Claude's strict-JSON output into a routing decision.

    Returns a decision dict on success, or None when the model output could not
    be parsed (missing/invalid JSON, or an empty API response passed in as "").
    A None result means "undecided" — the caller must leave the email for the
    next run and NOT mark it processed. There is deliberately NO fallback that
    turns a parse failure into an escalation (real Escalated only comes from a
    successfully parsed category=escalate).
    """
    if not text or not text.strip():
        return None
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        return None
    try:
        p = json.loads(m.group(0))
    except Exception:
        return None
    if not isinstance(p, dict) or "category" not in p:
        return None

    cat = p["category"] if p.get("category") in (
        "interested", "question", "decline", "spam", "escalate") else "escalate"
    stage = p["stage"] if p.get("stage") in (
        "qualify", "memo", "agreement_offer", "agreement_ready", "signed", "info") else ""
    reply = _clean_reply(p.get("reply_body"))
    note = (p.get("note") or "").strip() or "разобрано AI"

    # interested/question with no drafted reply → escalate to a human
    if cat in ("interested", "question") and not reply:
        cat = "escalate"
    # "signed" is impossible without an attachment — hard guard over the AI
    if stage == "signed" and not has_attachments:
        stage = "info"

    if cat in ("interested", "question"):
        route = "respond"
    elif cat == "spam":
        route = "ignore"
    else:
        route = cat  # decline | escalate
    if route == "respond" and stage == "signed":
        route = "signed"

    has_memo = route == "respond" and stage == "memo"
    has_contract = route == "respond" and stage == "agreement_ready"
    if route in ("respond", "signed"):
        new_status = STATUS_BY_STAGE.get(stage) or prev_status or "Replied"
    else:
        new_status = ""
    return {"category": route, "ai_category": cat, "stage": stage,
            "has_memo": has_memo, "has_contract": has_contract,
            "new_status": new_status, "note": note, "reply_body": reply}


def non_ai_result(pre_category):
    """Map a classifier pre-category (bounce/send_failed/auto_reply) to a route."""
    mapping = {"bounce": "bounce", "send_failed": "send_failed", "auto_reply": "auto_reply"}
    route = mapping.get(pre_category, "ignore")
    note = {
        "bounce": "Отбойник (мёртвый адрес)",
        "send_failed": "Письмо НЕ ушло (лимит/блок Gmail) — контакт вернуть в очередь",
        "auto_reply": "Стандартный автоответ (тикет-система/OOO) — контакт сомнительный",
    }.get(pre_category, "Служебное письмо")
    status = {"bounce": "Bounced", "send_failed": "Send Failed", "auto_reply": "Auto Reply"}.get(pre_category, "")
    return {"category": route, "ai_category": pre_category, "stage": "",
            "has_memo": False, "has_contract": False, "new_status": status,
            "note": note, "reply_body": ""}


# ═══════════════════════════════════════════════════════════════════
#   CRM matching
# ═══════════════════════════════════════════════════════════════════

def build_index(rows):
    """Index CRM rows by Email and by Last Msg ID / Thread ID for thread match."""
    by_email, by_token = {}, {}
    for r in rows:
        info = {
            "status": str(r.get("Status", "")).strip(),
            "company": str(r.get("Company", "")).strip(),
            "last_msg": str(r.get("Last Msg ID", "")).strip(),
            "thread_id": str(r.get("Thread ID", "")).strip(),
            "email": str(r.get("Email", "")).strip().lower(),
        }
        if info["email"] and info["email"] not in by_email:
            by_email[info["email"]] = info
        for tok in (info["last_msg"], info["thread_id"]):
            if tok:
                by_token[tok] = info
    return by_email, by_token


def match_contact(msg, by_email, by_token):
    """Match an inbound message to a CRM row. Prefer thread-token match
    (In-Reply-To / References contain our stored Message-ID), then sender email."""
    refs = (msg.get("references", "") + " " + msg.get("in_reply_to", "")).strip()
    if refs:
        for tok, info in by_token.items():
            if tok and tok in refs:
                return info
    return by_email.get(msg.get("from_email", ""), None)


# ═══════════════════════════════════════════════════════════════════
#   Actions (send + sheet), guarded by DRY_RUN
# ═══════════════════════════════════════════════════════════════════

def _reply_subject(subject):
    s = subject or ""
    return s if s.lower().startswith("re:") else "Re: " + s


def _print_draft(kind, account, to, subject, attachment, body):
    print("\n" + "=" * 70)
    print(f"[DRAFT · {kind}]  DRY_RUN — not sent")
    print(f"  from      : {account['user']}")
    print(f"  to        : {to}")
    print(f"  subject   : {subject}")
    print(f"  attachment: {attachment or '(none)'}")
    print("  body:")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def do_reply(account, msg, decision):
    """Send the funnel reply (with memo/contract attachment when required)."""
    attachment = None
    if decision["has_memo"]:
        attachment = DOC_MEMO_PATH
    elif decision["has_contract"]:
        attachment = DOC_CONTRACT_PATH
    subject = _reply_subject(msg["subject"])
    if DRY_RUN:
        _print_draft(f"reply/{decision['stage']}", account, msg["from_email"],
                     subject, attachment, decision["reply_body"])
        return
    ec.send_email(account, msg["from_email"], subject, decision["reply_body"],
                  in_reply_to=msg["message_id"], references=msg["references"],
                  attachment_path=attachment)


def do_signed(account, msg, decision, contact):
    """Signed branch: thank-you to the partner + deal review with the signed doc."""
    # 1) thank-you / next steps to the partner (Claude-drafted at stage 'signed')
    subject = _reply_subject(msg["subject"])
    if DRY_RUN:
        _print_draft("signed/thank-you", account, msg["from_email"], subject,
                     None, decision["reply_body"])
    else:
        ec.send_email(account, msg["from_email"], subject, decision["reply_body"],
                      in_reply_to=msg["message_id"], references=msg["references"])

    # 2) deal review to the team, with the signed document attached
    company = (contact or {}).get("company", "") or msg["from_email"]
    review_text = build_review_text(msg, contact)
    signed_path = _dump_largest_attachment(msg)
    review_subject = f"✅ Подписанный контракт: {company} — ревью сделки"
    if DRY_RUN:
        _print_draft("signed/review", account, ", ".join(REVIEW_TO), review_subject,
                     signed_path, review_text)
    else:
        ec.send_email(account, REVIEW_TO, review_subject, review_text,
                      attachment_path=signed_path)


def build_review_text(msg, contact):
    company = (contact or {}).get("company", "") or "не определена"
    status = (contact or {}).get("status", "") or "нет"
    user = (f"Данные CRM: компания {company}, email {msg['from_email']}, "
            f"статус был: {status}.\n\nПоследнее письмо партнёра:\n{msg['body']}")
    review = ec.call_claude(REVIEW_SYSTEM, user, max_tokens=2000)
    if not review:
        review = ("Автоматическое ревью не сформировано (AI недоступен). Данные сделки:\n\n"
                  f"Компания: {company}\nEmail: {msg['from_email']}\n"
                  f"Статус до подписания: {status}\n\n"
                  f"Последнее письмо партнёра:\n{msg['body'][:2000]}")
    review += "\n\nПодписанный документ во вложении."
    return review


def _dump_largest_attachment(msg):
    """Write the largest incoming attachment (the signed doc) to a temp file."""
    atts = msg.get("attachments", [])
    if not atts:
        return None
    best = max(atts, key=lambda a: a.get("size", 0))
    if not best.get("data"):
        return None
    safe_name = re.sub(r"[^\w.\-]+", "_", best["filename"] or "signed_document")
    path = os.path.join(tempfile.gettempdir(), f"utd_signed_{safe_name}")
    with open(path, "wb") as f:
        f.write(best["data"])
    return path


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


def update_status(msg, contact, decision):
    """Queue Status + Date Replied/Sent + Last Msg ID for the CRM row."""
    if not decision["new_status"]:
        return
    email = (contact.get("email") if contact else "") or msg["from_email"]
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    enqueue_update(email, {
        "Status": decision["new_status"],
        "Date Replied": now,
        "Date Sent": now,            # our last outbound timestamp (drives reminders)
        "Last Msg ID": msg["message_id"],
    })
    print(f"  [SHEET] queued Status='{decision['new_status']}' for {email}")


def set_contact_status(email, status):
    """Queue ONLY the Status for a contact matched by Email (bounce /
    send_failed / decline / auto_reply — no reply, no timestamps)."""
    if not status or not email:
        return
    enqueue_update(email, {"Status": status})
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


def format_thread_history(thread):
    """Render a fetch_thread() list into a compact chronology for the model,
    calling out memo/agreement attachments explicitly so the stage is obvious."""
    if not thread:
        return ""
    lines = []
    for m in thread:
        who = "UTD (us)" if m.get("direction") == "sent" else "Prospect"
        atts = m.get("attachment_names") or []
        att_note = ""
        if atts:
            joined = ", ".join(atts)
            low = joined.lower()
            tag = ""
            if "agreement" in low:
                tag = " [AGREEMENT DOCUMENT]"
            elif "overview" in low or "memo" in low or "informational" in low:
                tag = " [MEMO/OVERVIEW]"
            att_note = f" | attachments: {joined}{tag}"
        snippet = (m.get("snippet") or "").strip()
        lines.append(f"- [{m.get('date','')}] {who}: {snippet}{att_note}")
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════
#   Outbound bookkeeping in state (PII-safe: contacts keyed by hashed email)
# ═══════════════════════════════════════════════════════════════════

def record_outbound(state, email, stage_status, account_user, is_reminder=False):
    """Remember our last outbound to a contact so silence reminders can be paced.
    Stored under a SHA256 of the email (repo is public — no raw addresses)."""
    if DRY_RUN or not email:
        return
    key = ec.hash_id(email)
    contacts = state.setdefault("contacts", {})
    c = contacts.setdefault(key, {})
    c["last_sent"] = datetime.now(timezone.utc).isoformat()
    c["stage"] = stage_status
    c["account"] = account_user
    if is_reminder:
        rem = c.setdefault("reminders", {})
        rem[stage_status] = rem.get(stage_status, 0) + 1


def _reminder_count(state, email, stage_status):
    c = state.get("contacts", {}).get(ec.hash_id(email), {})
    return c.get("reminders", {}).get(stage_status, 0)


def _parse_dt(value):
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


def _last_outbound_dt(state, row, email):
    """When did we last send to this contact? Prefer state, fall back to sheet."""
    c = state.get("contacts", {}).get(ec.hash_id(email), {})
    dt = _parse_dt(c.get("last_sent"))
    if dt:
        return dt
    return _parse_dt(row.get("Date Sent")) or _parse_dt(row.get("Date Replied"))


REMINDER_BODY = {
    "Qualifying": (
        "Hi,\n\n"
        "I wanted to follow up on my last note about the UTD Agency Partner Program.\n\n"
        "Whenever you have a moment, it would help to know which themes you currently use for clients, "
        "roughly how many themes you buy per month, whether you already work with other theme providers, "
        "and how you usually handle theme purchases for your clients.\n\n"
        "Once I have that I will send you the full program overview.\n\n"
        "Best regards,\nSergey\nUTD Web | utdweb.team"),
    "Memo Sent": (
        "Hi,\n\n"
        "I am following up on the program overview I sent earlier. Did you get a chance to look it over?\n\n"
        "I am happy to answer any questions about commission, payouts or reporting. If it looks like a fit, "
        "I can send over the Agency Partner Agreement so you can get started.\n\n"
        "Best regards,\nSergey\nUTD Web | utdweb.team"),
    "Agreement Sent": (
        "Hi,\n\n"
        "I am following up on the Agency Partner Agreement I sent. To start earning commission on confirmed "
        "theme purchases you only need to sign your side and send it back.\n\n"
        "There is no fee and no exclusivity, and you can withdraw in writing at any time. If any clause "
        "needs clarifying, reply and I will walk you through it.\n\n"
        "Best regards,\nSergey\nUTD Web | utdweb.team"),
}


def send_reminders(rows, state, stats):
    """Scan the CRM for mid-dialogue contacts who went silent and nudge them
    toward the next step. Max MAX_REMINDERS per stage, spaced REMINDER_DAYS apart."""
    now = datetime.now(timezone.utc)
    default_account = next((a for a in ACCOUNTS if a["password"]), None)
    for row in rows:
        status = str(row.get("Status", "")).strip()
        email = str(row.get("Email", "")).strip().lower()
        if status not in REMINDER_STAGES or not email:
            continue
        last_out = _last_outbound_dt(state, row, email)
        if not last_out:
            continue  # no timestamp to measure silence from
        days_silent = (now - last_out).total_seconds() / 86400.0
        if days_silent < REMINDER_DAYS:
            continue
        if _reminder_count(state, email, status) >= MAX_REMINDERS:
            continue

        # pick the mailbox we used before, else the first available account
        stored_acc = state.get("contacts", {}).get(ec.hash_id(email), {}).get("account")
        account = next((a for a in ACCOUNTS if a["user"] == stored_acc), None) or default_account
        if not account or not account["password"]:
            continue

        body = REMINDER_BODY.get(status, "")
        thread_ref = str(row.get("Last Msg ID", "")).strip()
        references = str(row.get("Thread ID", "")).strip() or thread_ref
        subject = "Re: UTD Agency Partner Program"
        attachment = DOC_CONTRACT_PATH if status == "Agreement Sent" else None
        n = _reminder_count(state, email, status) + 1
        stats["reminder"] = stats.get("reminder", 0) + 1

        print(f"\n· REMINDER #{n} → {email} | stage={status} | silent {days_silent:.1f}d")
        if DRY_RUN:
            _print_draft(f"reminder/{status} #{n}", account, email, subject, attachment, body)
            continue
        ec.send_email(account, email, subject, body,
                      in_reply_to=thread_ref or None, references=references or None,
                      attachment_path=attachment)
        record_outbound(state, email, status, account["user"], is_reminder=True)
        enqueue_update(email, {"Date Sent": now.strftime("%Y-%m-%d %H:%M:%S")})


# ═══════════════════════════════════════════════════════════════════
#   Per-message processing
# ═══════════════════════════════════════════════════════════════════

def process_message(account, msg, by_email, by_token, state, stats):
    mid = msg.get("message_id", "")

    # Influencer chain is handled elsewhere — never touch it here.
    if INFLUENCER_MARKER.search(msg.get("subject", "")):
        return

    # Dedup: already processed in a previous run (hashed Message-ID)?
    if ec.is_processed(state, mid):
        return

    contact = match_contact(msg, by_email, by_token)

    # Dedup: this exact message was the last one we already acted on.
    if contact and contact.get("last_msg") and contact["last_msg"] == mid:
        if not DRY_RUN:
            ec.mark_processed(state, mid)
        return

    pre = ec.classify_incoming(msg, OWN_ADDRESSES)

    if pre != "human":
        decision = non_ai_result(pre)
    else:
        prev_status = contact["status"] if contact else ""
        company = contact["company"] if contact else ""

        # Pull the WHOLE thread (both directions) so the model reads the real
        # history, not just this one email + a possibly-stale CRM status.
        # Each thread is fetched at most ONCE per run (cached by gm_thrid).
        history = ""
        thrid = msg.get("gm_thrid", "")
        cache_key = (account["user"], thrid)
        if thrid and cache_key in _THREAD_CACHE:
            history = _THREAD_CACHE[cache_key]
        else:
            try:
                thread = ec.fetch_thread(account, thrid, OWN_ADDRESSES)
                history = format_thread_history(thread)
                if thrid:
                    _THREAD_CACHE[cache_key] = history
            except Exception as e:
                print(f"  (thread fetch failed, using single message: {e})")

        user = (
            "Incoming email:\n"
            f"From: {msg['from']}\n"
            f"Subject: {msg['subject']}\n"
            f"contact_status (CRM, may be stale): {prev_status or '(none)'}\n"
            f"Company (from CRM): {company or '(unknown)'}\n"
            f"has_attachments: "
            f"{('true (' + ', '.join(msg['attachment_names']) + ')') if msg['has_attachments'] else 'false'}\n\n"
            f"THREAD HISTORY (oldest first; use this to decide the real stage):\n"
            f"{history or '(no prior history found)'}\n\n"
            f"Latest incoming message body:\n{msg['body']}\n\n"
            "Decide the real funnel stage from the thread history above, classify it, and, if "
            "appropriate, draft the reply. Output ONLY the JSON object."
        )
        # Gentle pacing between Claude calls to avoid 429 on big batches.
        time.sleep(0.7)
        ai_text = ec.call_claude(SYSTEM_PROMPT, user, max_tokens=1500)
        decision = parse_ai_result(ai_text, msg["has_attachments"], prev_status)

        # Undecided (API error / unparseable JSON) is NOT an escalation.
        # Leave the email untouched so the next run retries it.
        if decision is None:
            print(f"\n· {account['user']} ← {msg['from_email']} | {msg['subject'][:60]!r}")
            print("  UNDECIDED (no valid AI result) → left for next run, not marked processed.")
            return

    route = decision["category"]
    stats[route] = stats.get(route, 0) + 1
    print(f"\n· {account['user']} ← {msg['from_email']} | {msg['subject'][:60]!r}")
    print(f"  route={route} stage={decision['stage'] or '-'} "
          f"status→{decision['new_status'] or '-'} | {decision['note']}")

    if route == "respond":
        do_reply(account, msg, decision)
        update_status(msg, contact, decision)
        record_outbound(state, msg["from_email"], decision["new_status"], account["user"])
    elif route == "signed":
        do_signed(account, msg, decision, contact)
        update_status(msg, contact, decision)
        record_outbound(state, msg["from_email"], decision["new_status"], account["user"])
    elif route in ("bounce", "send_failed"):
        # The DSN comes FROM mailer-daemon, not the contact. Pull the real failed
        # recipient out of the bounce body/headers and flag THAT contact.
        rcpt = ec.extract_failed_recipient(msg, OWN_ADDRESSES)
        if rcpt:
            print(f"  failed recipient = {rcpt} → Status '{decision['new_status']}'")
            set_contact_status(rcpt, decision["new_status"])
            if route == "bounce":
                ec.mark_bounced_everywhere(rcpt, dry_run=DRY_RUN)
        else:
            print("  failed recipient not found → sheet left untouched.")
    elif route == "decline":
        target = (contact.get("email") if contact else "") or msg["from_email"]
        set_contact_status(target, "Declined")
    elif route == "auto_reply":
        target = (contact.get("email") if contact else "") or msg["from_email"]
        set_contact_status(target, decision["new_status"] or "Auto Reply")
    elif route == "escalate":
        # A real, parsed escalation: no reply, no notification. The email stays
        # in the inbox for a human. We DO mark it processed (decision was made).
        print("  ESCALATE → left for a human (no reply, no notification).")
    # 'ignore' → nothing

    if not DRY_RUN:
        ec.mark_processed(state, mid)


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def run_once():
    print(f"=== UTD Agency Partner autoresponder | DRY_RUN={DRY_RUN} | "
          f"lookback {LOOKBACK_DAYS}d | reminders every {REMINDER_DAYS}d x{MAX_REMINDERS} | "
          f"{datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)

    # Reset in-run caches / pending writes.
    _PENDING.clear()
    _THREAD_CACHE.clear()
    _SHEET["ws"] = None
    _SHEET["header"] = []
    _SHEET["email_to_row"] = {}

    # CRM snapshot: open the worksheet ONCE and read ONCE (with 429 backoff).
    # Everything after this uses the in-memory snapshot — no per-email reads.
    rows = []
    try:
        ws = ec.open_worksheet(SHEET_ID, SHEET_TAB)
        rows = ec.read_rows_ws(ws)
        _SHEET["ws"] = ws
        _SHEET["header"] = list(rows[0].keys()) if rows else ws.row_values(1)
        # Map Email → sheet row number (header is row 1, data starts at row 2).
        e2r = {}
        for i, r in enumerate(rows, start=2):
            em = str(r.get("Email", "")).strip().lower()
            if em and em not in e2r:
                e2r[em] = i
        _SHEET["email_to_row"] = e2r
    except Exception as e:
        print(f"⚠️  Could not read CRM sheet: {e}")

    by_email, by_token = build_index(rows)
    print(f"CRM: {len(rows)} rows read in 1 call, {len(by_email)} emails indexed, "
          f"{len(_SHEET['email_to_row'])} rows mapped.")

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
                process_message(account, msg, by_email, by_token, state, stats)
            except Exception as e:
                print(f"  !! error on message {msg.get('message_id','?')}: {e}")

    # Silence-based follow-up reminders across the CRM.
    if rows:
        print("\n--- reminders (silence follow-ups) ---")
        try:
            send_reminders(rows, state, stats)
        except Exception as e:
            print(f"⚠️  reminder pass error: {e}")

    # Flush ALL queued CRM updates in a single Sheets batchUpdate call.
    flush_updates()

    if not DRY_RUN:
        ec.save_state(STATE_FILE, state)
    print(f"\n=== done. routes: {stats} ===")
    return {"parser": "agency_autoresponder", "routes": stats,
            "dry_run": DRY_RUN, "sheet_writes_queued": len(_PENDING)}


if __name__ == "__main__":
    run_once()
