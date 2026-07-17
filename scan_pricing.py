#!/usr/bin/env python3
"""
scan_pricing.py — one-off recovery: dig through ALL inbox history of both
outreach mailboxes and find every reply where a contact quoted PRICES for their
services (influencer rate cards, guest-post pricing, etc.). Many of these never
reached the Pricing sheet because of the old marker / serge-mailbox / fetch-cap
bugs, so they are missing from the case memory.

Extraction uses CONTEXT WINDOWS around each money amount (bodies arrive as one
HTML blob, so line-splitting missed everything) and keeps only windows that also
contain a real service-offer word — this excludes promos, signatures, and our
own pitch copy ("apps run $15 to $50 a month") quoted back in autoreplies.

For each pricing email found it writes one row to the shared 'Closed' log:
  Date · Chain · Contact · Company · Account · Outcome · the price snippets
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

MONEY = re.compile(r"(\$\s?\d[\d,. ]*|\d[\d,. ]*\s?(?:usd|eur|gbp|dollars?|euros?)\b|€\s?\d|£\s?\d)", re.I)
# A real service offer that a supplier/creator names next to a price.
OFFER = re.compile(r"\b(per\s+(?:video|post|story|reel|short|month|article|integration|mention)|"
                   r"dedicated\s+video|integration\s+video|video\s+integration|walkthrough|"
                   r"shorts?|reel|story|stories|newsletter|guest\s+post|link\s+insertion|"
                   r"insight-?driven|article|segment|rate\s*card|my\s+rate|our\s+rate|"
                   r"sponsorship|sponsor|package|bundle|placement|affiliate|"
                   r"long\s+video|tutorial|collaboration\s+options|pricing\s+options)\b", re.I)
DAEMON = re.compile(r"mailer-daemon|postmaster|no-?reply|donotreply|notifications?@", re.I)
PROMO = re.compile(r"millionverifier|thank you for contacting|welcome to|unsubscribe here", re.I)


def _price_windows(body):
    """Return de-duped ±110-char windows around money amounts that also contain
    a service-offer word inside the window (real quote, not noise)."""
    txt = re.sub(r"\s+", " ", body or "")
    out, seen = [], set()
    for mo in MONEY.finditer(txt):
        a = max(0, mo.start() - 110); b = min(len(txt), mo.end() + 110)
        w = txt[a:b].strip()
        if not OFFER.search(w):
            continue
        key = w[:45]
        if key in seen:
            continue
        seen.add(key)
        out.append(w)
        if len(out) >= 6:
            break
    return " ⋯ ".join(out)


def _chain_for(subject):
    s = (subject or "").lower()
    if "referral" in s or "b2b" in s or "shopify plus" in s:
        return "B2B"
    return "Influencer"


def main():
    if not ACCOUNTS:
        print("no mailbox credentials — nothing to scan"); return
    try:
        existing = {str(r.get("Contact", "")).strip().lower()
                    for r in ec.read_rows(NOTABLE_SHEET_ID, CLOSED_TAB)}
    except Exception:
        existing = set()
    print(f"Closed already has {len(existing)} contacts")

    found = {}   # contact_email -> row (richest wins)
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
            subj = m.get("subject", "") or ""
            if PROMO.search(body) or PROMO.search(subj):
                continue
            snippets = _price_windows(body)
            if not snippets:
                continue
            row = {
                "Date": "(scan)", "Chain": _chain_for(subj), "Contact": frm,
                "Company/Store": (m.get("from", "") or "").split("<")[0].strip().strip('"'),
                "Account": acc["user"],
                "Outcome": "Прайс в письме (найдено сканом почты)",
                "Details / review": snippets[:1200]}
            prev = found.get(frm)
            if not prev or len(row["Details / review"]) > len(prev["Details / review"]):
                found[frm] = row

    added = 0
    for frm, row in found.items():
        if frm in existing:
            print(f"  = {frm} (already in Closed)"); continue
        if ec.append_closed(NOTABLE_SHEET_ID, CLOSED_TAB, row):
            added += 1
            print(f"  + {frm}: {row['Details / review'][:90]}")
    print(f"DONE: {len(found)} pricing emails found, {added} new added to '{CLOSED_TAB}'")


if __name__ == "__main__":
    main()
