#!/usr/bin/env python3
"""Runner used by GitHub Actions. Usage: python run.py b2b|influencers"""
import os, sys, json

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
else:
    print("usage: python run.py b2b|influencers"); sys.exit(2)

summary = m.run_once()
print("SUMMARY:", json.dumps(summary, ensure_ascii=False))
