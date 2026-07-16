#!/usr/bin/env python3
"""
scan_pricing_diag.py — DIAGNOSTIC (prints only, writes nothing).
Figure out why the strict scan found 0 price emails: scan INBOX *and* All Mail
across every configured mailbox, and for each external-human message report
whether it matches MONEY, RATE_KW, or both — dumping the actual matching lines so
we can eyeball real price quotes and recalibrate the filter.
"""
import os
import re
import email_common as ec

LOOKBACK_DAYS = int(os.environ.get("SCAN_LOOKBACK_DAYS", "180"))

ACCOUNTS = [a for a in (
    {"user": os.environ.get("UTD_MAIL_SERGEY", ""), "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")},
    {"user": os.environ.get("UTD_MAIL_SERGE", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGE", "")},
    {"user": os.environ.get("UTD_MAIL_SERGI", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGI", "")},
    {"user": os.environ.get("UTD_MAIL_SERHII", ""), "password": os.environ.get("GMAIL_APP_PW_SERHII", "")},
) if a["user"] and a["password"]]

OWN = {os.environ.get(k, "").lower() for k in
       ("UTD_MAIL_SERGEY", "UTD_MAIL_SERGE", "UTD_MAIL_SERGI", "UTD_MAIL_SERHII",
        "UTD_MAIL_DENYS")} - {""}

MONEY = re.compile(r"(\$\s?\d[\d,. ]*|\d[\d,. ]*\s?(?:usd|eur|gbp|dollars?|euros?)\b|€\s?\d|£\s?\d)", re.I)
RATE_KW = re.compile(r"\b(per\s+(?:video|post|story|reel|short|month|article|integration|mention)|"
                     r"rate\s*card|my\s+rate|our\s+rate|rates?\b|charge|pricing|price\b|fee\b|"
                     r"sponsor|flat\s*rate|package|cost|collab|paid|budget|quote)\b", re.I)
DAEMON = re.compile(r"mailer-daemon|postmaster|no-?reply|donotreply|notifications?@", re.I)


def money_lines(body):
    out = []
    for ln in (body or "").splitlines():
        ln = ln.strip()
        if ln and MONEY.search(ln) and len(ln) < 220:
            out.append(ln)
    return out[:5]


def main():
    if not ACCOUNTS:
        print("no mailbox credentials"); return
    grand = {"msgs": 0, "human": 0, "money": 0, "kw": 0, "both": 0}
    samples = []
    for acc in ACCOUNTS:
        for box in ("INBOX", "[Gmail]/All Mail"):
            try:
                msgs = ec.fetch_inbox(acc, since_days=LOOKBACK_DAYS, mailbox=box)
            except Exception as e:
                print(f"IMAP {acc['user']} {box} failed: {e}"); continue
            print(f">>> {acc['user']} / {box}: {len(msgs)} messages")
            for m in msgs:
                grand["msgs"] += 1
                frm = (m.get("from_email", "") or "").lower()
                if not frm or frm in OWN or DAEMON.search(m.get("from", "") or ""):
                    continue
                grand["human"] += 1
                body = m.get("body", "") or ""
                has_m = bool(MONEY.search(body))
                has_k = bool(RATE_KW.search(body))
                grand["money"] += has_m
                grand["kw"] += has_k
                if has_m and has_k:
                    grand["both"] += 1
                if has_m and len(samples) < 60:
                    samples.append((frm, m.get("subject", "")[:60], box.split("/")[-1],
                                    " | ".join(money_lines(body))[:160]))
    print("\n==== TOTALS ====")
    print(grand)
    print("\n==== MONEY-matching human emails (samples) ====")
    for frm, subj, box, lines in samples:
        print(f"[{box}] {frm} | {subj}")
        print(f"        {lines}")


if __name__ == "__main__":
    main()
