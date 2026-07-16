#!/usr/bin/env python3
"""
scan_pricing.py — one-off recovery: dig through ALL inbox history of both
outreach mailboxes and find every reply where a contact quoted PRICES for their
services (influencer rate cards, sponsorship fees, etc.). Many of these never
reached the Pricing sheet because of the old marker / serge-mailbox / fetch-cap
bugs, so they are missing from the case memory.

For each pricing email found it writes one row to the shared 'Closed' log:
  Date · Chain=Influencer · Contact · Company · Account · Outcome · the price lines
Idempotent: skips a contact already in Closed.
"""
import os
import re

import email_common as ec

creds = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
creds_file = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")
if creds:
    with open(creds_file, "w") as f:
        f.write(creds)
os.environ["GOOGLE_CREDS_FILE"] = creds_file

B2B_SHEET_ID = os.environ.get("B2B_SHEET_ID", "")
NOTABLE_SHEET_ID = os.environ.get("NOTABLE_SHEET_ID") or B2B_SHEET_ID
CLOSED_TAB = os.environ.get("CLOSED_TAB", "Closed")
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

# A real money amount + a services/rate context word nearby.
MONEY = re.compile(r"(\$\s?\d[\d,. ]*|\d[\d,. ]*\s?(?:usd|eur|gbp|dollars?|euros?)\b|€\s?\d|£\s?\d)", re.I)
RATE_KW = re.compile(r"\b(per\s+(?:video|post|story|reel|short|month|article|integration|mention)|"
                     r"rate\s*card|my\s+rate|our\s+rate|rates?\b|charge|pricing|price\b|fee\b|"
                     r"sponsor|flat\s*rate|package|cost)\b", re.I)
DAEMON = re.compile(r"mailer-daemon|postmaster|no-?reply|donotreply|notifications?@", re.I)


def _price_lines(body):
    out = []
    for ln in (body or "").splitlines():
        ln = ln.strip()
        if ln and MONEY.search(ln) and len(ln) < 200:
            out.append(ln)
    return " | ".join(out[:8])


def main():
    if not ACCOUNTS:
        print("no mailbox credentials — nothing to scan"); return
    try:
        existing = {str(r.get("Contact", "")).strip().lower()
                    for r in ec.read_rows(NOTABLE_SHEET_ID, CLOSED_TAB)}
    except Exception:
        existing = set()
    print(f"Closed already has {len(existing)} contacts")

    found = {}   # contact_email -> row (last one wins)
    for acc in ACCOUNTS:
        try:
            msgs = ec.fetch_inbox(acc, since_days=LOOKBACK_DAYS, unseen_only=False)
        except Exception as e:
            print(f"IMAP {acc['user']} failed: {e}"); continue
        print(f">>> {acc['user']}: {len(msgs)} messages")
        for m in msgs:
            frm = (m.get("from_email", "") or "").lower()
            if not frm or frm in OWN or DAEMON.search(m.get("from", "") or ""):
                continue
            body = m.get("body", "") or ""
            if not (MONEY.search(body) and RATE_KW.search(body)):
                continue
            lines = _price_lines(body)
            if not lines:
                continue
            found[frm] = {
                "Date": "(scan)", "Chain": "Influencer", "Contact": frm,
                "Company/Store": (m.get("from", "") or "").split("<")[0].strip().strip('"'),
                "Account": acc["user"],
                "Outcome": "Прайс в письме (найдено сканом почты)",
                "Details / review": lines[:900]}

    added = 0
    for frm, row in found.items():
        if frm in existing:
            continue
        if ec.append_closed(NOTABLE_SHEET_ID, CLOSED_TAB, row):
            added += 1
            print(f"  + {frm}: {row['Details / review'][:80]}")
    print(f"DONE: {len(found)} pricing emails found, {added} new added to '{CLOSED_TAB}'")


if __name__ == "__main__":
    main()
