#!/usr/bin/env python3
"""
b2b_signed_notify.py — notify the team the moment a B2B client signs a contract.

TRIGGER: a row in the B2B sheet whose Status is a "signed" value (see
SIGNED_STATUSES). Whoever closes the deal sets the row's Status to "Signed"
(or Won / "Договор подписан"); this script, run every dispatcher tick, detects
newly-signed rows and emails everyone on SIGNED_NOTIFY_TO, exactly once per
contact.

Runs 24/7 (it is in the dispatcher ALWAYS list) and costs nothing until a row
is actually marked signed — it only reads the sheet and sends on a real change.

State: data/signed_notify_state.json {"notified": [<email>...]}.
Recipients: SIGNED_NOTIFY_TO env (comma-separated), else the default below.
"""
import json
import os

import email_common as ec

SHEET_ID = os.environ.get("B2B_SHEET_ID", "")
SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")
STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(STATE_DIR, "signed_notify_state.json")

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")

# Send the notification from the primary UTD mailbox (same creds as b2b sender).
ACCOUNT = {
    "user": os.environ.get("B2B_SENDER_USER", os.environ.get("UTD_MAIL_SERGEY", "")),
    "password": os.environ.get("SENDER_APP_PW") or os.environ.get("GMAIL_APP_PW_SERGEY", ""),
}

# Who gets the "contract signed" alert. Override with the SIGNED_NOTIFY_TO repo
# variable (comma-separated) to add/adjust addresses without a code change.
DEFAULT_TO = ("denvdavydov@gmail.com,marketing@utdweb.team,"
              "sergey.smortkin.utd@gmail.com,george.smortkin@gmail.com")
RECIPIENTS = [x.strip() for x in
              (os.environ.get("SIGNED_NOTIFY_TO") or DEFAULT_TO).split(",") if x.strip()]

# A row counts as "signed" when its Status matches (case-insensitive) any of:
SIGNED_STATUSES = {
    "signed", "contract signed", "won", "closed won", "deal won",
    "договор подписан", "подписан", "подписал",
}


def run_once():
    if not SHEET_ID:
        return {"parser": "b2b_signed_notify", "sent": 0, "error": "no B2B_SHEET_ID"}
    st = ec.load_state(STATE_FILE)
    notified = set(st.get("notified", []))
    try:
        rows = ec.read_rows(SHEET_ID, SHEET_TAB)
    except Exception as e:
        return {"parser": "b2b_signed_notify", "sent": 0, "error": str(e)[:200]}

    sent = 0
    for r in rows:
        status = str(r.get("Status", "")).strip().lower()
        email = str(r.get("Email", "")).strip().lower()
        if status not in SIGNED_STATUSES or not email or email in notified:
            continue
        company = str(r.get("Company Name", "") or r.get("Company", "") or "(unknown)").strip()
        website = str(r.get("Website", "")).strip()
        subject = f"✅ B2B contract signed: {company}"
        body = (
            "A B2B client just signed the referral contract.\n\n"
            f"Company: {company}\n"
            f"Email:   {r.get('Email', '')}\n"
            f"Website: {website}\n"
            f"Status:  {r.get('Status', '')}\n\n"
            "(Automated alert from the UTD outreach system.)"
        )
        if DRY_RUN:
            print(f"[DRY_RUN] would alert {RECIPIENTS} → signed: {company} ({email})")
        else:
            ec.send_email(ACCOUNT, RECIPIENTS, subject, body, from_name="UTD Outreach")
            print(f"alerted {len(RECIPIENTS)} recipients → signed: {company} ({email})")
        notified.add(email)
        sent += 1

    st["notified"] = sorted(notified)
    ec.save_state(STATE_FILE, st)
    return {"parser": "b2b_signed_notify", "sent": sent, "total_signed": len(notified)}


if __name__ == "__main__":
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
