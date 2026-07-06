#!/usr/bin/env python3
"""Runner used by GitHub Actions. Usage: python run.py b2b|influencers|ecom
Runs one parser, then POSTs a summary to NOTIFY_URL (n8n webhook) so a
notification email is sent — on both success and failure."""
import os, sys, json, traceback

try:
    import requests
except Exception:
    requests = None


def notify(payload):
    url = os.environ.get("NOTIFY_URL", "").strip()
    if url and requests:
        try:
            requests.post(url, json=payload, timeout=30)
        except Exception as e:
            print("notify failed:", e)


# Write Google service-account creds from the GH secret into a file
creds = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
creds_file = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")
if creds:
    with open(creds_file, "w") as f:
        f.write(creds)
os.environ["GOOGLE_CREDS_FILE"] = creds_file

which = (sys.argv[1] if len(sys.argv) > 1 else "").lower()
if which == "b2b":
    import b2b_harvester as m
elif which in ("influencers", "influencer"):
    import influencer_scraper as m
elif which == "ecom":
    import ecom_harvester as m
else:
    print("usage: python run.py b2b|influencers|ecom"); sys.exit(2)

try:
    summary = m.run_once()
    summary["status"] = "ok"
    print("SUMMARY:", json.dumps(summary, ensure_ascii=False))
    notify(summary)
except Exception as e:
    traceback.print_exc()
    notify({"parser": which, "status": "error", "added": 0, "error": str(e)[:300]})
    sys.exit(1)
