#!/usr/bin/env python3
"""
answer_unanswered.py — one-off: reply to a CURATED allow-list of real people who
answered our outreach but never got a reply (the autoresponder skipped them under
the old thinking bug). Junk (auto-acks, newsletters, promos, job-seekers) is NOT
in the list, so nothing wrong goes out. Replies in-thread, in Sergey's voice,
grounded in the actual conversation. DRY_RUN=true prints drafts and sends nothing.
"""
import os
import re
import email_common as ec

# email -> goal hint for the reply (which chain / what we want next)
ALLOW = {
    # influencer / creator collab (rate cards, "yes I'll do it", travelling, etc.)
    "office@shopioso.com":        "influencer",
    "kadirkoseoglu.01@gmail.com": "influencer",
    "imran.themescode@gmail.com": "influencer",
    "contact.support012@gmail.com":"influencer",
    "justin@spodmedia.com":       "influencer",
    "media.erisbusiness@gmail.com":"influencer",
    "business@smrtcontent.com":   "influencer",
    "hujjat@codeinspire.io":      "influencer",
    "support@appstoreresearch.com":"influencer",
    # ecom merchants
    "customerservice.intl@kizik.com":"ecom",
    "crew@bruntworkwear.com":     "ecom",
    # agency / referral program
    "info@ecomxagency.com":       "agency",
    "hello@ecomxagency.com":      "agency",
    "support@webcomforts.com":    "agency",
    "info@gadlio.com":            "agency",
}

GOALS = {
    "influencer": ("They are a content creator replying about a sponsored Shopify theme-review "
                   "collaboration. Goal: move it forward concretely — acknowledge exactly what they "
                   "offered (formats/prices), agree on a sensible next step, and offer to give them "
                   "full theme access so they can start. If they declined or are travelling, reply "
                   "warmly and leave the door clearly open for later."),
    "ecom": ("They are a Shopify merchant who replied to outreach about UTD themes. Goal: help them to "
             "the next step toward trying the right theme, offer to advise which theme fits their store, "
             "and point them to the official Theme Store, with no pressure."),
    "agency": ("They are an agency replying about the UTD referral program (their client buys the theme on "
               "the official Shopify Theme Store, they keep the client and earn a commission). Goal: answer "
               "their questions and move toward sending the full program memo / agreeing a next step. If they "
               "instead pitched us their own paid service (guest posts, etc.), politely clarify we are "
               "inviting them to OUR referral program, not buying their service, and keep it friendly."),
}

SIG = "\n\nBest regards,\nSergey\nUTD Web | utdweb.team"

SYSTEM = (
    "You are Sergey from UTD Web, a Shopify theme studio with themes on Shopify's official Theme Store. "
    "You are replying to a REAL person who answered your outreach. Read the WHOLE thread and write your next "
    "reply to their LATEST message, responding directly and specifically to what THEY said.\n"
    "VOICE: write like a real person, never like AI or ad copy. Longer, flowing, simple sentences that go "
    "straight to the point and join ideas with 'and', 'so', 'because'. Simple words anyone understands. No em "
    "dash anywhere (use a comma or period). No hype words (exclusive, exciting, seamless, game-changer). No "
    "'I hope this finds you well'.\n"
    "CONTENT: acknowledge the concrete thing they said (a price, a format, a question, a soft no), then move "
    "the conversation ONE clear step forward with ONE ask that can be answered by email. Never invent numbers, "
    "prices or results. The only links allowed: https://utdweb.team and https://themes.shopify.com/themes?q=UTD. "
    "Never suggest a call or meeting.\n"
    "Reply with ONLY the email body: a greeting line, a blank line, then short paragraphs. No 'Subject:' line, "
    "no signature (it is added automatically)."
)

ACCOUNTS = [a for a in (
    {"user": os.environ.get("UTD_MAIL_SERGEY", ""), "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")},
    {"user": os.environ.get("UTD_MAIL_SERGE", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGE", "")},
) if a["user"] and a["password"]]

OWN = {os.environ.get(k, "").lower() for k in
       ("UTD_MAIL_SERGEY", "UTD_MAIL_SERGE", "UTD_MAIL_SERGI", "UTD_MAIL_SERHII", "UTD_MAIL_DENYS")} - {""}

DRY = os.environ.get("DRY_RUN", "true").lower() != "false"
LOOKBACK = int(os.environ.get("SCAN_LOOKBACK_DAYS", "30"))


def _resubject(s):
    s = (s or "").strip()
    return s if s.lower().startswith("re:") else "Re: " + (s or "your message")


def _strip_signoff(b):
    """Remove a sign-off the model appended (Best regards,\\nSergey\\nUTD Web...)
    so our own SIG is not duplicated. Cuts at a sign-off line whose tail to the
    end is short and mentions Sergey/UTD (i.e. a signature, not body text)."""
    b = (b or "").strip()
    for mm in re.finditer(r"(?im)^\s*(best regards|best|cheers|warm regards|kind regards|"
                          r"regards|sincerely|thanks|thank you)\b[,.]?\s*$", b):
        tail = b[mm.start():]
        if len(tail) < 260 and ("sergey" in tail.lower() or "utd" in tail.lower()):
            return b[:mm.start()].strip()
    return b


def main():
    if not ACCOUNTS:
        print("no mailbox creds"); return
    sent = 0
    for acc in ACCOUNTS:
        try:
            msgs = ec.fetch_inbox(acc, since_days=LOOKBACK)
        except Exception as e:
            print(f"IMAP {acc['user']} failed: {e}"); continue
        done_thr = set()
        for m in msgs:  # newest-first
            frm = (m.get("from_email", "") or "").lower()
            if frm not in ALLOW:
                continue
            thr = m.get("gm_thrid") or frm
            if thr in done_thr:
                continue  # only the newest message per thread
            done_thr.add(thr)
            if frm in OWN:
                continue  # newest is ours → already answered
            # Build the reply from the full thread.
            try:
                thread = ec.fetch_thread(acc, str(m.get("gm_thrid", "") or ""), tuple(OWN), max_msgs=8)
                convo = ec.format_thread_readable(thread) if thread else (m.get("body", "") or "")
            except Exception:
                convo = m.get("body", "") or ""
            goal = GOALS[ALLOW[frm]]
            user = (f"GOAL: {goal}\n\nCONVERSATION (oldest to newest):\n{convo}\n\n"
                    "Write Sergey's reply to their latest message now.")
            body = ec.call_claude(SYSTEM, user, model="claude-sonnet-5", max_tokens=900)
            if not body or not body.strip():
                print(f"  [skip {frm}] empty AI reply"); continue
            body = re.sub(r"\s+—\s+", ", ", body).replace("—", "-").strip()
            body = _strip_signoff(body)   # remove any sign-off the model added; we add SIG once
            if not re.match(r"(?i)^\s*(hi|hello|hey|dear)\b", body):
                body = "Hi there,\n\n" + body
            body += SIG
            subj = _resubject(m.get("subject", ""))
            print(f"\n===== {acc['user']} → {frm} [{ALLOW[frm]}] =====\nSUBJECT: {subj}\n{body}\n")
            if not DRY:
                try:
                    ec.send_email(acc, frm, subj, body,
                                  in_reply_to=m.get("message_id") or None,
                                  references=((m.get("references", "") + " " + m.get("message_id", "")).strip() or None),
                                  from_name="Sergey | UTD Web")
                    print(f"  [SENT ✅ {frm}]")
                    sent += 1
                except Exception as e:
                    print(f"  [SEND FAILED {frm}] {e}")
    print(f"\nDONE. {'DRY_RUN — nothing sent' if DRY else f'sent {sent} replies'}")


if __name__ == "__main__":
    main()
