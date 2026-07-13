#!/usr/bin/env python3
"""
backfill.py — one-off recovery of ecom contacts found but never written to the
sheet on 2026-07-12 (Sheets write failed while the old code still marked them
'seen' in dedup). Self-contained: connects with gspread directly so the exact
write error is visible (the harvester's backoff swallows it). Idempotent —
dedups against column A.

Usage: python backfill.py data/backfill_ecom_0712.json
"""
import json
import os
import sys
import time
import traceback

creds = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
creds_file = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")
if creds:
    with open(creds_file, "w") as f:
        f.write(creds)

SHEET_ID = os.environ.get("ECOM_SHEET_ID", "").strip()
TAB = os.environ.get("ECOM_SHEET_TAB", "Ecom Contacts")


def row_values(r):
    return [r["email"], r["store"], r["url"], r["industry"],
            r["theme"], "Shopify", r["validated_by"], r["date"]]


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "data/backfill_ecom_0712.json"
    recs = json.load(open(path, encoding="utf-8"))
    print(f"backfill: {len(recs)} candidate records from {path}", flush=True)

    import gspread
    gc = gspread.service_account(filename=creds_file)
    ss = gc.open_by_key(SHEET_ID)
    try:
        ws = ss.worksheet(TAB)
    except Exception:
        ws = ss.get_worksheet(0)
        print(f"tab '{TAB}' not found — using first tab '{ws.title}'", flush=True)
    print(f"connected: tab='{ws.title}' in '{ss.title}'", flush=True)

    all_vals = ws.get_all_values()
    have = {row[0].strip().lower() for row in all_vals[1:] if row and "@" in row[0]}
    next_row = len(all_vals) + 1
    print(f"sheet has {len(all_vals)} rows, {len(have)} emails; next append row = {next_row}",
          flush=True)

    todo = [r for r in recs if r["email"].strip().lower() not in have]
    print(f"to write: {len(todo)} new rows", flush=True)

    written = 0
    for i in range(0, len(todo), 20):
        batch = todo[i:i + 20]
        rows = [row_values(r) for r in batch]
        for attempt in range(1, 4):
            try:
                ws.update(range_name=f"A{next_row}", values=rows,
                          value_input_option="USER_ENTERED")
                next_row += len(rows)
                written += len(rows)
                print(f"  wrote {len(rows)} rows (total {written})", flush=True)
                break
            except Exception as e:
                print(f"  WRITE ERROR (attempt {attempt}/3): {type(e).__name__}: "
                      f"{str(e)[:300]}", flush=True)
                if attempt == 3:
                    traceback.print_exc()
                time.sleep(5 * attempt)
        time.sleep(1.2)
    print(f"DONE: wrote {written} of {len(todo)} rows", flush=True)


if __name__ == "__main__":
    main()
