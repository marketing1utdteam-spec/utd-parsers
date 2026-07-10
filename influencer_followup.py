#!/usr/bin/env python3
"""
influencer_followup.py — UTD influencer outreach follow-up (7-day silence nudge).

CLAUDE-FIRST rewrite (2026-07). The FIRST outreach email stays a fixed template
(influencer_sender.py); every SUBSEQUENT letter is written by Claude with real
context: the creator's CRM row fields + the actual Gmail thread (when we have an
app password and a stored Thread ID) or a faithful summary of the first email.
The old static template is kept verbatim as the fallback when Claude fails.

Flow:
  «Daily at 9 AM»                     → schedule (one follow-up per run)
  «Get Creator Contacts»              → ec.read_rows(SHEET_ID, TAB)
  «Find Contacts Needing Follow-Up»   → pick_followup() (Status=='Sent' + Date Sent < now-7d)
  «Build Follow-Up Email»             → Claude (claude-sonnet-5) per copy canon,
                                        fallback = legacy static template
  «Send Follow-Up Email»              → ec.send_email; in-thread reply
                                        (In-Reply-To/References) when the row
                                        has a Thread ID, else "Re: …" subject
  «Mark as Follow-Up Sent»            → update Status/Date Sent by Email

Safety:
  • DRY_RUN=true (default) prints the FULL Claude system+user prompt, the draft
    and the intended sheet write — sends nothing, writes nothing.
  • Followed-up emails are deduped via a SHA256-hashed state file (repo is PUBLIC).
  • Claude drafts are post-validated (no em dash, no hype words, only allowed
    links, exact signature) — any violation falls back to the static template.

Env:  GOOGLE_CREDENTIALS_JSON, GMAIL_APP_PW_SERGEY, INFL_GMAIL_USER,
      INFL_SHEET_ID, INFL_SHEET_TAB, ANTHROPIC_API_KEY, DRY_RUN, STATE_DIR
"""

import os
import re
import imaplib
from datetime import datetime, timezone, timedelta

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

SHEET_ID = os.environ.get("INFL_SHEET_ID", "")
SHEET_TAB = os.environ.get("INFL_SHEET_TAB", "Sheet1")

INFL_GMAIL_USER = os.environ.get("INFL_GMAIL_USER", os.environ.get("UTD_MAIL_SERGEY", ""))
ACCOUNT = {"user": INFL_GMAIL_USER, "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")}

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# Days of silence before a follow-up.
FOLLOWUP_DAYS = int(os.environ.get("INFL_FOLLOWUP_DAYS", "7"))

_STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(_STATE_DIR, "influencer_followup_state.json")

CLAUDE_MODEL = "claude-sonnet-5"
CLAUDE_MAX_TOKENS = 1200

# Subject of the FIRST outreach email (influencer_sender.py) — referenced in
# the summary context and used for the "Re:" reply subject.
FIRST_SUBJECT = "Shopify theme review collab — UTD Web"
REPLY_SUBJECT = "Re: " + FIRST_SUBJECT

SIGNATURE = "Best regards,\nSergey\nUTD Web | utdweb.team"


# ═══════════════════════════════════════════════════════════════════
#   FALLBACK EMAIL TEMPLATE  (legacy static body — used when Claude fails)
# ═══════════════════════════════════════════════════════════════════

FALLBACK_SUBJECT = "Follow-up: Shopify theme review collab — UTD Web"

FALLBACK_BODY_HTML = "Hi,<br><br>I wanted to follow up on my previous message regarding a potential collaboration with UTD Web.<br><br>As a quick reminder: we're a Shopify theme development team with 5 themes and 25 presets on the <a href='https://themes.shopify.com/themes?page=1&q=utd'>Shopify Theme Store</a> — <a href='https://themes.shopify.com/themes/gain'>Gain</a>, <a href='https://themes.shopify.com/themes/ultra'>Ultra</a>, <a href='https://themes.shopify.com/themes/boutique'>Boutique</a>, <a href='https://themes.shopify.com/themes/allure'>Allure</a>, and <a href='https://themes.shopify.com/themes/victory'>Victory</a>. We're looking to partner with Shopify-focused creators for sponsored reviews and showcases.<br><br>We're happy to provide full theme access and compensate you for your time. If you're interested, I'd be glad to share more details or answer any questions.<br><br>Best regards,<br>Sergey<br>UTD Web<br><a href='https://utdweb.team'>utdweb.team</a>"


# ═══════════════════════════════════════════════════════════════════
#   CLAUDE SYSTEM PROMPT  (copy canon)
# ═══════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You write follow-up emails for UTD Web, a Shopify theme development studio with 5 themes and 25 presets on the official Shopify Theme Store (Gain, Ultra, Boutique, Allure, Victory). You write as Sergey, a normal guy from the UTD Web team writing an ordinary work email. The recipient is a content creator / influencer we already emailed once about a sponsored theme-review collaboration and who has not replied.

GOAL: get a reply and move the conversation toward collecting the creator's full rate card (formats + prices for their platforms). This is a calm check-in, not a hard sell and not a pressure play.

VOICE:
- Plain everyday English, simple words, natural flow. Read-aloud test: if you would not say a sentence out loud to a colleague, rewrite it.
- Open naturally and get to the point in the first sentence. No clever hooks, no fragmented openers.
- Zero filler, no marketing-speak, no dramatic one-liners.
- Calm and unhurried. No pushy asks like "yes or no?", no guilt-tripping about the silence.
- The email is as long as it needs to be to make the point, no longer.

FORMAT (mandatory):
- Line 1: a greeting (e.g. "Hi <name>," or just "Hi,").
- Blank line.
- Body split into short paragraphs by meaning, one idea per paragraph, blank lines between them. Never one big mixed paragraph.
- Blank line, then the farewell and signature.
- The FIRST sentence after the greeting must remind them of the earlier email, e.g. "I emailed you a couple of weeks ago about reviewing our Shopify themes and wanted to check back."

CONTENT RULES:
- Include ONE concrete observation about THIS creator (their channel, platform, or what they review), taken ONLY from the context provided. Never invent facts, names, or numbers.
- Add exactly ONE new angle versus the first email (for example: we compensate for time, we are happy to hear their rates, full theme access up front).
- End the body with a clear, simple ask about their rates and formats.
- Never offer or suggest a call or meeting. Everything happens over email; you may offer help by email ("reply and I'll walk you through it").
- Never use an em dash character anywhere in the letter.
- No hype words: exclusive, exciting, game-changer, handpicked, curated, unique opportunity. No corporate slop.
- Whenever you mention a theme by name, include its Theme Store link: Gain https://themes.shopify.com/themes/gain, Ultra https://themes.shopify.com/themes/ultra, Boutique https://themes.shopify.com/themes/boutique, Allure https://themes.shopify.com/themes/allure, Victory https://themes.shopify.com/themes/victory. Other allowed links: https://utdweb.team and https://themes.shopify.com/themes?q=UTD. No other links.
- Sign off EXACTLY with:
Best regards,
Sergey
UTD Web | utdweb.team

OUTPUT: plain text of the email body only. No subject line, no preamble, no markdown, no commentary."""


# ═══════════════════════════════════════════════════════════════════
#   EMAIL VALIDATION  (verbatim from «Find Contacts Needing Follow-Up»)
# ═══════════════════════════════════════════════════════════════════

_BAD = ['denvdavydov', 'smortkin', 'utdweb.team', 'utd.agency', 'its_always_teatime',
        '.png', '.jpg', '.webp', '@sentry', 'your-company', 'you@company',
        'example@', 'placeholder', '%20', 'noreply']


def is_valid(e):
    if not e or "@" not in e or " " in e or len(e) < 7:
        return False
    if any(b in e.lower() for b in _BAD):
        return False
    parts = e.split("@")
    if len(parts) != 2 or "." not in parts[1]:
        return False
    if re.match(r"^[0-9a-f]{20,}$", parts[0], re.I):
        return False
    return True


# ═══════════════════════════════════════════════════════════════════
#   SELECTION  (same filter as before: Status=='Sent' + 7d silence)
# ═══════════════════════════════════════════════════════════════════

def _parse_date_sent(value):
    """Parse a 'Date Sent' cell the way JS `new Date(d)` would accept it."""
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


def pick_followup(rows, state):
    """Return the FIRST contact needing a follow-up, or None.

    Filter (unchanged): Status == 'Sent' AND the email is valid AND 'Date Sent'
    exists AND 'Date Sent' is older than the cutoff (now - FOLLOWUP_DAYS).
    A hashed-state guard prevents re-following-up the same address across runs.
    Returns the FULL row too, so the Claude context can use every sheet field.
    """
    cutoff = datetime.now(timezone.utc) - timedelta(days=FOLLOWUP_DAYS)
    for r in rows:
        s = str(r.get("Status", "")).strip()
        e = str(r.get("Email", "")).strip()
        d = r.get("Date Sent")
        if s != "Sent" or not is_valid(e) or not d:
            continue
        dt = _parse_date_sent(d)
        if not dt or not (dt < cutoff):
            continue
        if ec.is_processed(state, e.lower()):
            continue
        return {
            "email": e,
            "channel": str(r.get("Channel") or r.get("Name")
                           or r.get("Creator / Blog Name") or r.get("Company Name") or "").strip(),
            "date_sent": str(d).strip(),
            "date_sent_dt": dt,
            "thread_id": str(r.get("Thread ID", "") or "").strip(),
            "row": r,
        }
    return None


# ═══════════════════════════════════════════════════════════════════
#   CONTEXT ASSEMBLY  (sheet row + real thread OR faithful summary)
# ═══════════════════════════════════════════════════════════════════

# Sheet columns worth showing the model (skipped when empty). Covers the
# naming variants used across the influencer sheet / scraper versions.
_CONTEXT_COLS = ["Name", "Channel", "Creator / Blog Name", "Company Name", "Company",
                 "Platform", "Profile / Blog URL", "Profile URL", "Website",
                 "Email Found At", "Shopify Content Links", "Activity Signals",
                 "Notes", "Subscribers", "Date Found"]


def creator_context(row, email):
    """Render the CRM row's non-empty fields as bullet lines for the prompt."""
    lines = ["- Email: " + email]
    seen = {"email"}
    for col in _CONTEXT_COLS:
        v = str(row.get(col, "") or "").strip()
        if not v or col.lower() in seen:
            continue
        seen.add(col.lower())
        lines.append("- %s: %s" % (col, v[:300]))
    return "\n".join(lines)


def _resolve_gm_thrid(account, thread_id):
    """Turn the sheet's Thread ID into a numeric Gmail thread id (X-GM-THRID).

    influencer_sender stores the SMTP Message-ID (<...@...>) in the Thread ID
    column, while fetch_thread needs Gmail's numeric thread id — so when the
    value is not already numeric we look the message up in All Mail by its
    Message-ID header and read the X-GM-THRID attribute. Returns "" on failure.
    """
    tid = (thread_id or "").strip()
    if not tid:
        return ""
    if tid.isdigit():
        return tid
    # n8n-era rows store the Gmail API thread id (hex) — that is just the hex
    # form of the numeric X-GM-THRID.
    if "@" not in tid and re.fullmatch(r"[0-9a-fA-F]{10,}", tid):
        try:
            return str(int(tid, 16))
        except Exception:
            pass
    M = None
    try:
        M = imaplib.IMAP4_SSL(ec.IMAP_HOST, ec.IMAP_PORT)
        M.login(account["user"], account["password"])
        for box in ("[Gmail]/All Mail", "[Google Mail]/All Mail"):
            typ, _ = M.select(box, readonly=True)
            if typ == "OK":
                break
        else:
            return ""
        typ, data = M.uid("SEARCH", None, "HEADER", "Message-ID", tid)
        if typ != "OK" or not data or not data[0]:
            return ""
        uid = data[0].split()[-1]
        typ, md = M.uid("FETCH", uid, "(X-GM-THRID)")
        if typ != "OK" or not md:
            return ""
        blob = b" ".join(p if isinstance(p, bytes) else (p[0] or b"")
                         for p in md if p)
        m = re.search(rb"X-GM-THRID\s+(\d+)", blob)
        return m.group(1).decode() if m else ""
    except Exception:
        return ""
    finally:
        if M is not None:
            try:
                M.logout()
            except Exception:
                pass


def format_thread(thread, limit=6, snippet_chars=400):
    """Render a fetch_thread() list into a trimmed chronology for the model."""
    lines = []
    for m in thread[-limit:]:
        who = "UTD (us)" if m.get("direction") == "sent" else "Creator"
        snippet = (m.get("snippet") or "").strip()[:snippet_chars]
        lines.append("[%s] %s | %s: %s" % (m.get("date", ""), who,
                                           m.get("subject", ""), snippet))
    return "\n".join(lines)


def build_correspondence_context(cand):
    """Real Gmail thread when possible, else a faithful summary of email #1."""
    if ACCOUNT["password"] and cand["thread_id"]:
        try:
            thrid = _resolve_gm_thrid(ACCOUNT, cand["thread_id"])
            if thrid:
                thread = ec.fetch_thread(ACCOUNT, thrid, (ACCOUNT["user"],))
                if thread:
                    return format_thread(thread), "thread"
        except Exception as e:
            print(f"  [thread] fetch failed ({e}) -> summary context")
    summary = ("first outreach sent %s, subject '%s', sponsored-review "
               "collaboration pitch for our 5 Shopify themes (Gain, Ultra, "
               "Boutique, Allure, Victory), no reply received"
               % (cand["date_sent"] or "over %d days ago" % FOLLOWUP_DAYS,
                  FIRST_SUBJECT))
    return summary, "summary"


def build_user_prompt(cand, correspondence):
    days_silent = ""
    if cand.get("date_sent_dt"):
        days_silent = str((datetime.now(timezone.utc) - cand["date_sent_dt"]).days)
    parts = [
        "Write the follow-up email for this creator.",
        "",
        "CREATOR (from our CRM sheet):",
        creator_context(cand["row"], cand["email"]),
        "",
        "CORRESPONDENCE SO FAR:",
        correspondence,
    ]
    if days_silent:
        parts += ["", "Days since our first email with no reply: " + days_silent]
    return "\n".join(parts)


# ═══════════════════════════════════════════════════════════════════
#   CLAUDE DRAFT + POST-VALIDATION (canon enforcement, fallback on breach)
# ═══════════════════════════════════════════════════════════════════

_BANNED_WORDS = ["exclusive", "exciting", "game-changer", "game changer",
                 "handpicked", "curated", "unique opportunity"]
_ALLOWED_LINKS = ("https://utdweb.team",
                  "https://themes.shopify.com/themes?q=UTD",
                  "https://themes.shopify.com/themes/gain",
                  "https://themes.shopify.com/themes/ultra",
                  "https://themes.shopify.com/themes/boutique",
                  "https://themes.shopify.com/themes/allure",
                  "https://themes.shopify.com/themes/victory")


def validate_draft(text):
    """Clean a Claude draft and enforce the copy canon.

    Returns the cleaned plain-text body, or "" when the draft breaks a hard
    rule (caller then uses the static fallback template).
    """
    t = (text or "").strip()
    t = re.sub(r"^```[a-zA-Z]*\s*", "", t)
    t = re.sub(r"\s*```$", "", t).strip()
    if not t:
        return ""
    # Canon: never an em dash in the letter (belt and braces after the prompt).
    t = t.replace("—", "-")
    low = t.lower()
    if any(b in low for b in _BANNED_WORDS):
        return ""
    # Only the two allowed links.
    for url in re.findall(r"https?://[^\s>\"')]+", t):
        if not any(url.rstrip(".,;") == a or url.rstrip(".,;").startswith(a)
                   for a in _ALLOWED_LINKS):
            return ""
    # No length cap: the email is as long as it needs to be (voice spec 2026-07).
    # Exact signature — append it if the model forgot.
    if "utdweb.team" not in low or "best regards" not in low:
        t = t.rstrip() + "\n\n" + SIGNATURE
    return t


def draft_followup(cand):
    """Ask Claude for the follow-up; return (subject, body, source).

    source: 'claude' | 'fallback'. In DRY_RUN the FULL prompt is printed first.
    On any Claude failure or canon violation the legacy static template is used.
    """
    correspondence, ctx_kind = build_correspondence_context(cand)
    user_prompt = build_user_prompt(cand, correspondence)

    if DRY_RUN:
        print("\n" + "-" * 70)
        print(f"[CLAUDE PROMPT · context={ctx_kind}]  (printed because DRY_RUN)")
        print("--- SYSTEM " + "-" * 59)
        print(SYSTEM_PROMPT)
        print("--- USER " + "-" * 61)
        print(user_prompt)
        print("-" * 70)

    raw = ec.call_claude(SYSTEM_PROMPT, user_prompt,
                         model=CLAUDE_MODEL, max_tokens=CLAUDE_MAX_TOKENS)
    body = validate_draft(raw)
    if body:
        return REPLY_SUBJECT, body, "claude"
    print("[claude unavailable -> fallback used]")
    return FALLBACK_SUBJECT, FALLBACK_BODY_HTML, "fallback"


# ═══════════════════════════════════════════════════════════════════
#   ACTIONS  (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def _now_ts():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def _print_draft(to, subject, body, source, in_reply_to):
    print("\n" + "=" * 70)
    print(f"[DRAFT · follow-up · {source}]  DRY_RUN — not sent")
    print(f"  from       : {ACCOUNT['user']} (Sergey | UTD Web)")
    print(f"  to         : {to}")
    print(f"  subject    : {subject}")
    print(f"  in-reply-to: {in_reply_to or '(none — new thread)'}")
    print("  body:")
    for line in (body or "").splitlines():
        print("    " + line)
    print("=" * 70)


def run_once():
    print(f"=== UTD influencer outreach FOLLOW-UP (Claude-first) | DRY_RUN={DRY_RUN} | "
          f"silence {FOLLOWUP_DAYS}d | {datetime.now(timezone.utc).isoformat()} ===")
    state = ec.load_state(STATE_FILE)

    try:
        rows = ec.read_rows(SHEET_ID, SHEET_TAB)
    except Exception as e:
        print(f"⚠️  Could not read creator sheet: {e}")
        return {"parser": "influencer_followup", "dry_run": DRY_RUN,
                "error": str(e), "sent": 0}

    print(f"Creator CRM: {len(rows)} rows read.")
    cand = pick_followup(rows, state)
    if not cand:
        print("No follow-ups needed — stop.")
        return {"parser": "influencer_followup", "dry_run": DRY_RUN,
                "found": False, "sent": 0, "rows": len(rows)}

    email = cand["email"]
    channel = cand["channel"]
    print(f"\n· follow-up → {email} | channel={channel or '-'} | "
          f"thread_id={'yes' if cand['thread_id'] else 'no'}")

    subject, body, source = draft_followup(cand)
    now = _now_ts()

    # In-thread reply headers: the stored Thread ID is the Message-ID of our
    # first email — usable directly as In-Reply-To/References when it looks
    # like an RFC 5322 Message-ID.
    in_reply_to = None
    tid = cand["thread_id"]
    if tid and "@" in tid:
        in_reply_to = tid if tid.startswith("<") else "<%s>" % tid.strip("<>")

    sent = 0
    if DRY_RUN:
        _print_draft(email, subject, body, source, in_reply_to)
        print(f"[SHEET] DRY_RUN — would set Status='Follow-up Sent', "
              f"Date Sent='{now}' for {email}")
    else:
        # send_email appends in_reply_to to References itself, so passing
        # in_reply_to alone yields both threading headers correctly.
        ec.send_email(ACCOUNT, email, subject, body,
                      in_reply_to=in_reply_to,
                      from_name="Sergey | UTD Web")
        ec.update_row_by_match(SHEET_ID, SHEET_TAB, "Email", email, {
            "Status": "Follow-up Sent",
            "Date Sent": now,
        })
        ec.mark_processed(state, email.lower())
        ec.save_state(STATE_FILE, state)
        sent = 1
        print(f"[SHEET] Status='Follow-up Sent' written for {email} ({source})")

    return {"parser": "influencer_followup", "dry_run": DRY_RUN, "found": True,
            "sent": sent, "email": email, "channel": channel,
            "draft_source": source, "rows": len(rows)}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
