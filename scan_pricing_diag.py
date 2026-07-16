#!/usr/bin/env python3
"""
scan_pricing_diag.py — DIAGNOSTIC v2 (prints only, writes nothing).
For every external-human INBOX message that contains BOTH a money amount and a
price keyword, print a ±110-char context window around each money hit, so we can
read the actual quote and tell real rate cards from promos/signatures/quoted-
original noise. Line-splitting failed because bodies arrive as one HTML blob.
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
PROMO = re.compile(r"millionverifier|thank you for contacting|welcome to|unsubscribe here", re.I)


def windows(body):
    txt = re.sub(r"\s+", " ", body or "")
    out, seen = [], set()
    for mo in MONEY.finditer(txt):
        a = max(0, mo.start() - 110); b = min(len(txt), mo.end() + 110)
        w = txt[a:b].strip()
        key = w[:40]
        if key in seen:
            continue
        seen.add(key)
        out.append(w)
        if len(out) >= 4:
            break
    return out


def main():
    if not ACCOUNTS:
        print("no mailbox credentials"); return
    seen_contacts = {}
    for acc in ACCOUNTS:
        try:
            msgs = ec.fetch_inbox(acc, since_days=LOOKBACK_DAYS)
        except Exception as e:
            print(f"IMAP {acc['user']} failed: {e}"); continue
        print(f">>> {acc['user']}: {len(msgs)} messages")
        for m in msgs:
            frm = (m.get("from_email", "") or "").lower()
            if not frm or frm in OWN or DAEMON.search(m.get("from", "") or ""):
                continue
            body = m.get("body", "") or ""
            subj = m.get("subject", "") or ""
            if PROMO.search(body) or PROMO.search(subj):
                continue
            if not (MONEY.search(body) and RATE_KW.search(body)):
                continue
            wins = windows(body)
            if not wins:
                continue
            # keep the richest sample per contact
            prev = seen_contacts.get(frm)
            if not prev or len(" ".join(wins)) > len(" ".join(prev[2])):
                seen_contacts[frm] = (acc["user"], subj, wins)

    print(f"\n==== {len(seen_contacts)} contacts with money+keyword context ====\n")
    for frm, (acc, subj, wins) in seen_contacts.items():
        print(f"### {frm}  ({acc})")
        print(f"    subj: {subj[:70]}")
        for w in wins:
            print(f"    …{w}…")
        print()


if __name__ == "__main__":
    main()
