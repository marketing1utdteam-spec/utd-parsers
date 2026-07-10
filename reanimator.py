#!/usr/bin/env python3
"""
reanimator.py — UTD outreach "Reanimator": re-queue failed sends.

Port of the n8n workflow REANIMATOR (BpXdnVWJub0lzJ5Z) to plain Python for
GitHub Actions.

What it does (VERBATIM from the n8n nodes «Только Send Failed» + «Сброс в очередь»):
  • Reads two CRM sheets (B2B + influencer).
  • Finds every row whose Status is EXACTLY 'Send Failed' (letters trimmed) — i.e.
    a send that failed by OUR fault (Gmail quota/block).
  • Resets that row's Status to '' (empty) so the sender picks it up again.
  • NEVER touches 'Bounced' (dead addresses): re-hitting dead addresses only
    damages the mailbox reputation. Only 'Send Failed' is reanimated.

Safety:
  • DRY_RUN=true (default) prints the intended cell resets, writes nothing.
  • DRY_RUN=false performs the resets in ONE batched Sheets call per sheet.

Usage:  python reanimator.py
Env:    GOOGLE_CREDENTIALS_JSON (service account), DRY_RUN,
        B2B_SHEET_ID, B2B_SHEET_TAB, INFL_SHEET_ID, INFL_SHEET_TAB, STATE_DIR
"""

import os
from datetime import datetime, timezone

import email_common as ec


# ═══════════════════════════════════════════════════════════════════
#   CONFIG
# ═══════════════════════════════════════════════════════════════════

# B2B CRM — «Читать таблицу» / «Сброс в очередь».
B2B_SHEET_ID = os.environ.get(
    "B2B_SHEET_ID", "1ggMS5Hko2jCY5eqcPvasBy3P6hAwbw8rldr4cS3Zeo4")
B2B_SHEET_TAB = os.environ.get("B2B_SHEET_TAB", "IT Companies — Emails")

# Influencer CRM — «Читать инфлюенсеров» / «Сброс в очередь (инфл)».
INFL_SHEET_ID = os.environ.get(
    "INFL_SHEET_ID", "12IiHIsdibJPRGYNyZfrvdmBDY9OjmsokdmL4GgWg4qQ")
INFL_SHEET_TAB = os.environ.get("INFL_SHEET_TAB", "Sheet1")

# The exact status we reanimate, and the value we reset it to. VERBATIM.
FAILED_STATUS = "Send Failed"
RESET_STATUS = ""

# STATE_DIR kept for parity with the other modules (unused here — no dedup state).
_STATE_DIR = os.environ.get("STATE_DIR", ".")

DRY_RUN = os.environ.get("DRY_RUN", "true").strip().lower() in ("1", "true", "yes", "on")


# ═══════════════════════════════════════════════════════════════════
#   Per-sheet reanimation
# ═══════════════════════════════════════════════════════════════════

def reanimate_sheet(label, sheet_id, tab):
    """Reset every 'Send Failed' row in one sheet back to an empty Status.

    Returns {"read": n_rows, "reset": n_reset, "emails": [...]}.
    """
    result = {"label": label, "read": 0, "reset": 0, "emails": []}
    try:
        ws = ec.open_worksheet(sheet_id, tab)
        rows = ec.read_rows_ws(ws)
    except Exception as e:
        print(f"⚠️  [{label}] could not read sheet: {e}")
        result["error"] = str(e)
        return result

    header = list(rows[0].keys()) if rows else ws.row_values(1)
    result["read"] = len(rows)

    if "Status" not in header:
        print(f"⚠️  [{label}] no 'Status' column — nothing to do.")
        return result
    status_col = header.index("Status") + 1  # 1-based

    cell_updates = []
    for i, r in enumerate(rows, start=2):  # data starts at row 2 (header is row 1)
        # Mirror n8n: String(j['Status']||'').trim() === 'Send Failed'
        if str(r.get("Status", "")).strip() == FAILED_STATUS:
            email = str(r.get("Email", "")).strip()
            a1 = ec.gspread_a1(i, status_col)
            cell_updates.append({"range": a1, "values": [[RESET_STATUS]]})
            result["emails"].append(email)
            print(f"  · row {i} | {email or '(no email)'} | Status "
                  f"'{FAILED_STATUS}' → '' (re-queued)")

    result["reset"] = len(cell_updates)
    if not cell_updates:
        print(f"[{label}] no 'Send Failed' rows — nothing to reanimate.")
        return result

    if DRY_RUN:
        print(f"[{label}] DRY_RUN — would reset {len(cell_updates)} cell(s) in ONE call.")
        return result
    try:
        n = ec.batch_update_cells(ws, cell_updates)
        print(f"[{label}] reset {n} 'Send Failed' row(s) → queue in ONE call.")
    except Exception as e:
        print(f"⚠️  [{label}] batch update failed after retries: {e}")
        result["error"] = str(e)
    return result


# ═══════════════════════════════════════════════════════════════════
#   MAIN
# ═══════════════════════════════════════════════════════════════════

def run_once():
    print(f"=== UTD Reanimator | DRY_RUN={DRY_RUN} | "
          f"{datetime.now(timezone.utc).isoformat()} ===")

    b2b = reanimate_sheet("B2B", B2B_SHEET_ID, B2B_SHEET_TAB)
    infl = reanimate_sheet("Influencer", INFL_SHEET_ID, INFL_SHEET_TAB)

    total_reset = b2b["reset"] + infl["reset"]
    print(f"\n=== done. reanimated {total_reset} row(s) "
          f"(B2B {b2b['reset']}, Infl {infl['reset']}) ===")
    return {
        "parser": "reanimator",
        "dry_run": DRY_RUN,
        "b2b": {"read": b2b["read"], "reset": b2b["reset"]},
        "influencer": {"read": infl["read"], "reset": infl["reset"]},
        "total_reset": total_reset,
    }


if __name__ == "__main__":
    import json
    print("SUMMARY:", json.dumps(run_once(), ensure_ascii=False))
