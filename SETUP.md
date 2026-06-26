# UTD Parsers — GitHub Actions setup (one-time, ~5 min)

Two daily automations that run the parsers in GitHub's cloud, append new
contacts to the Google Sheets, and commit the dedup-state back to the repo.
No server, free. After this setup it runs itself.

## Step 1 — create a PRIVATE GitHub repo
On github.com → New repository → name e.g. `utd-parsers` → **Private** → Create.
(Don't add a README — we push our own.)

## Step 2 — add the secrets
Repo → Settings → Secrets and variables → Actions → **New repository secret**.
Add these:

| Secret name | Value |
|---|---|
| `GOOGLE_CREDENTIALS_JSON` | the FULL service-account JSON (paste file contents) |
| `MILLIONVERIFIER_KEYS` | `key1,key2` (comma-separated MillionVerifier keys) |
| `ANTHROPIC_API_KEY` | optional — funded key for B2B niche filtering; can skip |

The service account is `utd-156@utdweb-498410.iam.gserviceaccount.com` and both
sheets are already shared with it.

## Step 3 — push the code
This folder is already a git repo with an initial commit. Just point it at your
new repo and push:
```bash
cd utd-parsers-gha
git remote add origin https://github.com/<your-username>/utd-parsers.git
git branch -M main
git push -u origin main
```

## Step 4 — verify
- GitHub → **Actions** tab. You'll see "B2B Parser — Daily" and
  "Influencer Parser — Daily".
- Test immediately: open one → **Run workflow** (manual button) → watch the log.
  The last line prints `SUMMARY: {...}` with how many contacts were added.
- Check the Google Sheets — new rows appear (append-only, columns unchanged).

## Schedule
- B2B: every day 06:00 UTC. Influencers: every day 07:00 UTC.
- Change times by editing the `cron:` line in
  `.github/workflows/parsers-*.yml`.
- Batch size per run: B2B `MAX_GOOGLE_QUERIES=15`, influencers
  `CSE_BUDGET=20`/`YT_BUDGET=10` (edit in the yml to harvest more/less per day).

## How dedup stays correct
- After each run the workflow commits `data/*.json` (what was already searched)
  back to the repo → next run continues, no repeats.
- Plus the parsers skip any email already in the sheet → no duplicate contacts
  even if state were lost.

## Optional — notify n8n / manager
The parsers already write to the sheets your n8n chains consume, so the daily
sending picks up new contacts automatically. If you also want a "parser ran:
+N today" ping, we can add an HTTP step that posts the SUMMARY to an n8n webhook.
