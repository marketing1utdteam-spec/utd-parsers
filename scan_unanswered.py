#!/usr/bin/env python3
"""
scan_unanswered.py — one-off (read-only): list HUMAN inbound messages on both
outreach mailboxes that we have NOT replied to. "Unanswered" = the newest
message in the thread came from THEM (not us), and they are a real person (not
a mailer-daemon / auto-reply / no-reply). Prints per-mailbox counts + a sample
list so we can decide how to answer. Sends nothing.
"""
import os
import re
import email_common as ec

LOOKBACK = int(os.environ.get("SCAN_LOOKBACK_DAYS", "21"))

ACCOUNTS = [a for a in (
    {"user": os.environ.get("UTD_MAIL_SERGEY", ""), "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")},
    {"user": os.environ.get("UTD_MAIL_SERGE", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGE", "")},
) if a["user"] and a["password"]]

OWN = {os.environ.get(k, "").lower() for k in
       ("UTD_MAIL_SERGEY", "UTD_MAIL_SERGE", "UTD_MAIL_SERGI", "UTD_MAIL_SERHII",
        "UTD_MAIL_DENYS")} - {""}

DAEMON = re.compile(r"mailer-daemon|postmaster|no-?reply|donotreply|do-not-reply|"
                    r"notifications?@|@bounce|@reply\.|mailchimp|sendgrid|noreply", re.I)
AUTO = re.compile(r"out of office|autoreply|auto-reply|automatic reply|thank you for contacting|"
                  r"we have received your|vacation|away from", re.I)


def main():
    if not ACCOUNTS:
        print("no mailbox creds"); return
    for acc in ACCOUNTS:
        try:
            msgs = ec.fetch_inbox(acc, since_days=LOOKBACK)
        except Exception as e:
            print(f"IMAP {acc['user']} failed: {e}"); continue
        # newest-first; walk unique threads, keep the newest msg per thread
        seen_thr = set()
        unanswered = []
        for m in msgs:
            thr = m.get("gm_thrid") or m.get("message_id")
            if not thr or thr in seen_thr:
                continue
            seen_thr.add(thr)
            frm = (m.get("from_email", "") or "").lower()
            frm_hdr = m.get("from", "") or ""
            subj = m.get("subject", "") or ""
            body = m.get("body", "") or ""
            # newest message in this thread is from THEM (human) and not answered after
            if not frm or frm in OWN:
                continue  # newest msg is ours → already answered
            if DAEMON.search(frm_hdr) or DAEMON.search(frm):
                continue  # bounce / no-reply
            if AUTO.search(subj) or AUTO.search(body[:400]):
                continue  # auto-reply / OOO
            unanswered.append({
                "from": frm, "name": frm_hdr.split("<")[0].strip().strip('"')[:40],
                "subject": subj[:70], "date": (m.get("date", "") or "")[:16],
                "snippet": re.sub(r"\s+", " ", body)[:120]})
        print(f"\n===== {acc['user']}: {len(unanswered)} НЕОТВЕЧЕННЫХ (человек, новейшее письмо в треде — от них) =====")
        for u in unanswered[:60]:
            print(f"  [{u['date']}] {u['from']} | {u['subject']}")
            print(f"       {u['snippet']}")
        if len(unanswered) > 60:
            print(f"  … ещё {len(unanswered)-60}")


if __name__ == "__main__":
    main()
