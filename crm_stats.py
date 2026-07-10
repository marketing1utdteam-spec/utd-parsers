#!/usr/bin/env python3
"""
crm_stats.py — UTD outreach "CRM Stats": daily pipeline snapshot.

Port of the n8n workflow CRM_STATS (crdEiDwQ7ahyLxye) to plain Python for
GitHub Actions.

What it does (mirrors the n8n «Статистика Б2Б» / «Статистика Инфл» code nodes):
  • Reads the B2B CRM sheet and the influencer sheet (+ its Pricing tab).
  • Counts inbound/outbound Gmail volume for both outreach mailboxes.
  • Computes the exact B2B and Influencer metric dicts (Russian keys VERBATIM).
  • Appends each dict as a row to its «Stats» tab (as «Запись Stats …» did).
  • Emails a plain-text summary of both snapshots to the manager.
    NOTE: the n8n workflow had NO email node (it only appended to the Stats
    tabs). The summary email is added for the GHA port per the task brief;
    recipient is configurable (CRM_STATS_TO). The metric computation and the
    Stats-tab append reproduce the n8n logic verbatim.

Safety:
  • DRY_RUN=true (default) prints the metrics, the Stats append and the email;
    writes/sends nothing.

Usage:  python crm_stats.py
Env:    GOOGLE_CREDENTIALS_JSON, GMAIL_APP_PW_SERGEY, GMAIL_APP_PW_SERGE,
        DRY_RUN, B2B_SHEET_ID, B2B_SHEET_TAB, INFL_SHEET_ID, INFL_SHEET_TAB,
        INFL_PRICING_TAB, STATS_TAB, CRM_STATS_TO, GMAIL_COUNT_DAYS,
        INFL_GMAIL_1_USER/INFL_GMAIL_1_PW, INFL_GMAIL_2_USER/INFL_GMAIL_2_PW,
        STATE_DIR
"""

import os
import math
from datetime import datetime, timezone

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

# «Читать Б2Б» / «Запись Stats Б2Б»
B2B_SHEET_ID = os.environ.get("B2B_SHEET_ID", "")
B2B_SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")

# «Читать Инфл» / «Читать Pricing» / «Запись Stats Инфл»
INFL_SHEET_ID = os.environ.get("INFL_SHEET_ID", "")
INFL_SHEET_TAB = os.environ.get("INFL_SHEET_TAB", "Sheet1")
INFL_PRICING_TAB = os.environ.get("INFL_PRICING_TAB", "Pricing")

# Both stat snapshots are appended to a «Stats» tab in their own spreadsheet.
STATS_TAB = os.environ.get("STATS_TAB", "Stats")

# Summary email recipient (added for the GHA port — no email node in the n8n src).
CRM_STATS_TO = os.environ.get("CRM_STATS_TO", "")

# Gmail volume counts: the n8n getAll nodes had no time filter (a default page
# limit). We count over a bounded window; override with GMAIL_COUNT_DAYS.
GMAIL_COUNT_DAYS = int(os.environ.get("GMAIL_COUNT_DAYS", "35"))
SENT_MAILBOX = os.environ.get("SENT_MAILBOX", "[Gmail]/Sent Mail")

# B2B outreach mailboxes («…я1» / «…я2»).
B2B_ACCOUNTS = [a for a in (
    {"user": os.environ.get("UTD_MAIL_SERGEY", ""), "password": os.environ.get("GMAIL_APP_PW_SERGEY", "")},
    {"user": os.environ.get("UTD_MAIL_SERGE", ""),  "password": os.environ.get("GMAIL_APP_PW_SERGE", "")},
) if a["user"]]

# Influencer outreach mailboxes — addresses unknown in source; env-configurable.
# When unset the counts stay 0, exactly like glen()'s try/catch → 0 in the n8n.
INFL_ACCOUNTS = [
    {"user": os.environ.get("INFL_GMAIL_1_USER", ""),
     "password": os.environ.get("INFL_GMAIL_1_PW", "")},
    {"user": os.environ.get("INFL_GMAIL_2_USER", ""),
     "password": os.environ.get("INFL_GMAIL_2_PW", "")},
]

_STATE_DIR = os.environ.get("STATE_DIR", ".")

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")


# ═══════════════════════════════════════════════════════════════════
#   Helpers
# ═══════════════════════════════════════════════════════════════════

def _s(v):
    return str(v if v is not None else "").strip()


def _pct(a, b):
    """pct(a,b) = b ? Math.round(a/b*1000)/10 : 0  (JS Math.round = round-half-up)."""
    if not b:
        return 0
    return math.floor(a / b * 1000.0 + 0.5) / 10.0


def _today():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")  # toISOString().slice(0,10)


def _count_mailbox(account, mailbox):
    """glen() equivalent: number of Gmail messages. Any failure → 0."""
    if not account.get("user") or not account.get("password"):
        return 0
    try:
        return len(ec.fetch_inbox(account, since_days=GMAIL_COUNT_DAYS,
                                  mailbox=mailbox, unseen_only=False, limit=100000))
    except Exception:
        return 0


# ═══════════════════════════════════════════════════════════════════
#   Metric computation  (VERBATIM key set from the n8n code nodes)
# ═══════════════════════════════════════════════════════════════════

def compute_b2b(rows, gmail_in, gmail_out):
    """«Статистика Б2Б» — verbatim."""
    cnt = {}
    total = sent_ever = replied = 0
    for r in rows:
        if not _s(r.get("Email")):
            continue
        total += 1
        st = _s(r.get("Status"))
        cnt[st] = cnt.get(st, 0) + 1
        if _s(r.get("Date Sent")):
            sent_ever += 1
        if _s(r.get("Date Replied")):
            replied += 1
    dialog = (cnt.get("Qualifying", 0) + cnt.get("Memo Sent", 0)
              + cnt.get("Agreement Offered", 0) + cnt.get("Agreement Ready", 0)
              + cnt.get("Agreement Sent", 0))
    return {
        "Дата": _today(),
        "Всего контактов": total,
        "В очереди": cnt.get("", 0),
        "Отправлено (всего писем ушло)": sent_ever,
        "Ожидают ответа (Sent)": cnt.get("Sent", 0),
        "Получено ответов": replied,
        "В диалоге (воронка)": dialog,
        "Ответили (Replied)": cnt.get("Replied", 0),
        "Подписали контракт": cnt.get("Agreement Signed", 0),
        "Отказались": cnt.get("Declined", 0),
        "Отбойники (мёртвые адреса)": cnt.get("Bounced", 0),
        "Сбои отправки": cnt.get("Send Failed", 0),
        "Не контактировать": cnt.get("Do Not Contact", 0),
        "Авто-ответы (тикеты/OOO)": cnt.get("Auto Reply", 0),
        "Ответов получено (Gmail, всего)": gmail_in,
        "Наших ответов отправлено (Gmail)": gmail_out,
        "Reply rate %": _pct(gmail_in, sent_ever),
        "Bounce rate %": _pct(cnt.get("Bounced", 0), sent_ever),
    }


def compute_infl(contacts, pricing, gmail_in, gmail_out):
    """«Статистика Инфл» — verbatim."""
    cnt = {}
    total = sent_ever = 0
    for r in contacts:
        if not _s(r.get("Email")):
            continue
        total += 1
        st = _s(r.get("Status"))
        cnt[st] = cnt.get(st, 0) + 1
        if _s(r.get("Date Sent")) or st == "Sent":
            sent_ever += 1
    # pricing already filtered to real rows (Email non-empty and != '(init)')
    pcnt = {}
    for p in pricing:
        st = _s(p.get("Contact Status"))
        pcnt[st] = pcnt.get(st, 0) + 1
    replied = len(pricing)
    return {
        "Дата": _today(),
        "Всего контактов": total,
        "В очереди": cnt.get("", 0),
        "Отправлено (всего писем ушло)": sent_ever,
        "Получено ответов (в Pricing)": replied,
        "В переговорах": pcnt.get("Negotiating", 0),
        "Прайс собран": pcnt.get("Data Complete", 0),
        "Эскалации": pcnt.get("Escalated", 0),
        "Отказались": cnt.get("Declined", 0),
        "Отбойники (мёртвые адреса)": cnt.get("Bounced", 0),
        "Сбои отправки": cnt.get("Send Failed", 0),
        "Авто-ответы (тикеты/OOO)": cnt.get("Auto Reply", 0),
        "Ответов получено (Gmail, всего)": gmail_in,
        "Наших ответов отправлено (Gmail)": gmail_out,
        "Reply rate %": _pct(gmail_in, sent_ever),
        "Bounce rate %": _pct(cnt.get("Bounced", 0), sent_ever),
    }


# ═══════════════════════════════════════════════════════════════════
#   Outputs (guarded by DRY_RUN)
# ═══════════════════════════════════════════════════════════════════

def append_stats(label, sheet_id, tab, stats):
    """«Запись Stats …»: append one row to the Stats tab (autoMap by header)."""
    if DRY_RUN:
        print(f"[{label}] DRY_RUN — would append 1 row to '{tab}':")
        for k, v in stats.items():
            print(f"    {k}: {v}")
        return
    try:
        ws = ec.open_worksheet(sheet_id, tab)
        header = ws.row_values(1)
        if header:
            # autoMapInputData maps by column NAME; blank for unknown headers.
            row = [stats.get(h, "") for h in header]
        else:
            # Fresh Stats tab: write header then the values.
            header = list(stats.keys())
            ec.with_sheets_backoff(lambda: ws.append_row(
                header, value_input_option="USER_ENTERED"))
            row = [stats[h] for h in header]
        ec.with_sheets_backoff(lambda: ws.append_row(
            row, value_input_option="USER_ENTERED"))
        print(f"[{label}] appended 1 row to '{tab}'.")
    except Exception as e:
        print(f"⚠️  [{label}] append to '{tab}' failed: {e}")


def _format_block(title, stats):
    lines = [title, "─" * len(title)]
    for k, v in stats.items():
        lines.append(f"{k}: {v}")
    return "\n".join(lines)


def send_summary(b2b, infl):
    subject = f"📊 UTD CRM Stats — {_today()}"
    body = (
        _format_block("B2B (IT Companies)", b2b)
        + "\n\n"
        + _format_block("Influencers", infl)
        + "\n\n— Automated CRM stats"
    )
    account = next((a for a in B2B_ACCOUNTS if a["password"]), None)
    if DRY_RUN:
        print("\n" + "=" * 70)
        print("[EMAIL]  DRY_RUN — not sent")
        print(f"  from   : {account['user'] if account else '(no account)'}")
        print(f"  to     : {CRM_STATS_TO}")
        print(f"  subject: {subject}")
        print("  body:")
        for line in body.splitlines():
            print("    " + line)
        print("=" * 70)
        return
    if not account:
        print("⚠️  No sending account available — summary email not sent.")
        return
    ec.send_email(account, CRM_STATS_TO, subject, body)
    print(f"[EMAIL] summary sent to {CRM_STATS_TO}.")


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def run_once():
    print(f"=== UTD CRM Stats | DRY_RUN={DRY_RUN} | "
          f"{datetime.now(timezone.utc).isoformat()} ===")

    # ── Reads ──────────────────────────────────────────────────────
    b2b_rows, infl_rows, pricing_rows = [], [], []
    try:
        b2b_rows = ec.read_rows(B2B_SHEET_ID, B2B_SHEET_TAB)
    except Exception as e:
        print(f"⚠️  Could not read B2B sheet: {e}")
    try:
        infl_rows = ec.read_rows(INFL_SHEET_ID, INFL_SHEET_TAB)
    except Exception as e:
        print(f"⚠️  Could not read Influencer sheet: {e}")
    try:
        pricing_raw = ec.read_rows(INFL_SHEET_ID, INFL_PRICING_TAB)
        # filter: Email non-empty and != '(init)'  (verbatim from «Статистика Инфл»)
        pricing_rows = [p for p in pricing_raw
                        if _s(p.get("Email")) and _s(p.get("Email")) != "(init)"]
    except Exception as e:
        print(f"⚠️  Could not read Pricing tab: {e}")

    # ── Gmail volume counts (glen equivalents; any error → 0) ───────
    b2b_in = sum(_count_mailbox(a, "INBOX") for a in B2B_ACCOUNTS)
    b2b_out = sum(_count_mailbox(a, SENT_MAILBOX) for a in B2B_ACCOUNTS)
    infl_in = sum(_count_mailbox(a, "INBOX") for a in INFL_ACCOUNTS)
    infl_out = sum(_count_mailbox(a, SENT_MAILBOX) for a in INFL_ACCOUNTS)

    # ── Compute ────────────────────────────────────────────────────
    b2b = compute_b2b(b2b_rows, b2b_in, b2b_out)
    infl = compute_infl(infl_rows, pricing_rows, infl_in, infl_out)

    print("\n--- B2B stats ---")
    for k, v in b2b.items():
        print(f"  {k}: {v}")
    print("\n--- Influencer stats ---")
    for k, v in infl.items():
        print(f"  {k}: {v}")

    # ── Outputs ────────────────────────────────────────────────────
    print()
    append_stats("B2B", B2B_SHEET_ID, STATS_TAB, b2b)
    append_stats("Influencer", INFL_SHEET_ID, STATS_TAB, infl)
    send_summary(b2b, infl)

    print("\n=== done. ===")
    return {"parser": "crm_stats", "dry_run": DRY_RUN,
            "b2b": b2b, "influencer": infl, "recipient": CRM_STATS_TO}


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
