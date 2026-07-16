#!/usr/bin/env python3
"""
backfill_closed.py — one-off: populate the shared 'Closed' log from what is
ALREADY marked complete in the CRM sheets, so the case memory is full from the
start (past deals weren't captured because the log began mid-July).

Sources of truth for "completed":
  • Influencer  → Pricing tab, Contact Status == "Data Complete"  (rate card done)
  • B2B         → B2B sheet,   Status == "Agreement Signed"
  • Ecom        → Ecom sheet,  Status == "Deal Closed"

Idempotent: skips a (chain, contact) already present in the Closed tab.
"""
import os
import email_common as ec

creds = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
creds_file = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")
if creds:
    with open(creds_file, "w") as f:
        f.write(creds)
os.environ["GOOGLE_CREDS_FILE"] = creds_file

B2B_SHEET_ID = os.environ.get("B2B_SHEET_ID", "")
B2B_SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")
ECOM_SHEET_ID = os.environ.get("ECOM_SHEET_ID", "")
ECOM_SHEET_TAB = os.environ.get("ECOM_SHEET_TAB", "Ecom Contacts")
INFL_SHEET_ID = os.environ.get("INFL_SHEET_ID", "")
INFL_PRICING_TAB = os.environ.get("INFL_PRICING_TAB", "Pricing")
NOTABLE_SHEET_ID = os.environ.get("NOTABLE_SHEET_ID") or B2B_SHEET_ID
CLOSED_TAB = os.environ.get("CLOSED_TAB", "Closed")

PRICE_COLS = ["Price Article", "Price YouTube Video", "Price Video Mention",
              "Price Shorts/Reels", "Price Social Post", "Price Story Mention",
              "Price Newsletter", "Packages", "Usage Rights", "Affiliate/RevShare",
              "Audience Size", "Audience Geo", "Expected Views"]


def _s(v):
    return str(v if v is not None else "").strip()


def _existing():
    try:
        rows = ec.read_rows(NOTABLE_SHEET_ID, CLOSED_TAB)
    except Exception:
        return set()
    return {(_s(r.get("Chain")).lower(), _s(r.get("Contact")).lower()) for r in rows}


def main():
    seen = _existing()
    print(f"Closed already has {len(seen)} rows")
    added = 0

    def add(chain, contact, company, outcome, details):
        nonlocal added
        contact = _s(contact)
        if not contact or (chain.lower(), contact.lower()) in seen:
            return
        ok = ec.append_closed(NOTABLE_SHEET_ID, CLOSED_TAB, {
            "Date": "(backfill)", "Chain": chain, "Contact": contact,
            "Company/Store": _s(company), "Account": "",
            "Outcome": outcome, "Details / review": details[:900]})
        if ok:
            seen.add((chain.lower(), contact.lower()))
            added += 1
            print(f"  + {chain}: {contact}")

    # Influencer — Data Complete (rate card collected)
    try:
        for r in ec.read_rows(INFL_SHEET_ID, INFL_PRICING_TAB):
            if _s(r.get("Contact Status")) == "Data Complete":
                rate = "; ".join(f"{c}={_s(r.get(c))}" for c in PRICE_COLS if _s(r.get(c)))
                add("Influencer", r.get("Email"), r.get("Name"),
                    "Rate card collected (ready for our content)", rate)
    except Exception as e:
        print(f"influencer pricing read failed: {e}")

    # B2B — Agreement Signed
    try:
        for r in ec.read_rows(B2B_SHEET_ID, B2B_SHEET_TAB):
            if _s(r.get("Status")) == "Agreement Signed":
                add("B2B", r.get("Email"), r.get("Company") or r.get("Company Name"),
                    "Agreement signed", "")
    except Exception as e:
        print(f"b2b read failed: {e}")

    # Ecom — Deal Closed
    try:
        for r in ec.read_rows(ECOM_SHEET_ID, ECOM_SHEET_TAB):
            if _s(r.get("Status")) == "Deal Closed":
                add("Ecom", r.get("Email"), r.get("Store Name"), "Deal closed", "")
    except Exception as e:
        print(f"ecom read failed: {e}")

    print(f"DONE: added {added} closed cases to '{CLOSED_TAB}'")


if __name__ == "__main__":
    main()
