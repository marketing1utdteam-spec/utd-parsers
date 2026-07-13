#!/usr/bin/env python3
"""
backfill.py — one-off recovery of ecom contacts that were found but never
written to the sheet on 2026-07-12 (Sheets 429 dropped the batch while the
old code still marked them 'seen' in dedup).

Reads a JSON list of records recovered from the run log and appends the ones
not already present in the sheet. Safe to re-run (dedups against column A).

Usage: python backfill.py data/backfill_ecom_0712.json
"""
import json
import os
import sys
import time

# Reuse the harvester's Sheets writer + credentials bootstrap.
creds = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
creds_file = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")
if creds:
    with open(creds_file, "w") as f:
        f.write(creds)
os.environ["GOOGLE_CREDS_FILE"] = creds_file

import ecom_harvester as h  # noqa: E402


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "data/backfill_ecom_0712.json"
    recs = json.load(open(path, encoding="utf-8"))
    print(f"backfill: {len(recs)} candidate records from {path}")

    up = h.SheetsUploader(h.SHEETS_SPREADSHEET_ID,
                          h.SHEETS_WORKSHEET_NAME, h.SHEETS_CREDS_FILE)
    if not up.connect():
        print("ERROR: could not connect to sheet"); sys.exit(1)

    have = up.existing_emails()
    print(f"sheet already has {len(have)} emails")
    todo = [r for r in recs if r["email"].strip().lower() not in have]
    print(f"to write: {len(todo)} new rows")

    written = 0
    for i in range(0, len(todo), 12):
        batch = todo[i:i + 12]
        if up.append_rows(batch):
            written += len(batch)
        time.sleep(2)  # stay well under the write quota
    print(f"DONE: wrote {written} rows (session _sent={up._sent})")


if __name__ == "__main__":
    main()
