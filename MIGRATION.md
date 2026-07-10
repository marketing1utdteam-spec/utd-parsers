# UTD Email — full migration to GitHub Actions

Everything that used to run in n8n (cold-send chains, follow-ups, reminders,
inbound auto-responders, reanimator, reply monitor, CRM stats) is now Python in
this repo, driven by GitHub Actions. n8n is no longer needed once this is verified live.

## Runtime map

| Chain | Module | Workflow | Schedule (UTC) | AI |
|---|---|---|---|---|
| B2B cold send | `b2b_sender.py` | outreach-b2b-sender | `:00,:30` 08–17 Mon–Fri | sonnet-5 |
| B2B follow-up | `b2b_followup.py` | outreach-b2b-followup | daily 09:00 | template |
| Influencer send | `influencer_sender.py` | outreach-influencer-sender | ~every 2h 00–22 | template |
| Influencer follow-up | `influencer_followup.py` | outreach-influencer-followup | daily 09:00 | template |
| Influencer reply reminders | `influencer_reminders.py` | outreach-influencer-reminders | every 15 min | template |
| Ecom sequence | `ecom_sender.py` | outreach-ecom-sender | `*/30` 07–21 | sonnet-5 |
| Referral/Agency autoresponder | `agency_autoresponder.py` | outreach-agency-autoresponder | hourly | sonnet-5 |
| Influencer autoresponder | `influencer_autoresponder.py` | outreach-influencer-autoresponder | every 15 min | sonnet-5 |
| Ecom autoresponder | `ecom_autoresponder.py` | outreach-ecom-autoresponder | every 15 min | sonnet-5 |
| Reanimator (requeue failed) | `reanimator.py` | outreach-reanimator | daily 07:37 | — |
| Reply monitor | `reply_monitor.py` | outreach-reply-monitor | every 15 min | — |
| CRM stats | `crm_stats.py` | outreach-crm-stats | daily 08:07 | — |

Plus the existing parsers: `b2b_harvester.py`, `influencer_scraper.py`, `ecom_harvester.py`.

All email/Sheets/Claude plumbing is shared in `email_common.py` (SMTP + IMAP +
Google Sheets + Claude). Model everywhere: `claude-sonnet-5`. Dedup state lives
in `data/*.json` (SHA-256 hashed — no PII) and is committed back each run.

## Safety default

Every workflow reads `DRY_RUN` from repo variable `vars.DRY_RUN`, defaulting to
**`true`**. In DRY_RUN it logs the exact draft it *would* send and the sheet
writes it *would* make — but sends nothing and writes nothing. So it is safe the
moment it lands; nothing goes out until you flip it live.

## To go live (only these steps require you — I can't add GitHub secrets)

1. **Add repository secrets** (Settings → Secrets and variables → Actions → New secret):
   - `GMAIL_APP_PW_SERGEY` — app password for the sergey.utd mailbox
   - `GMAIL_APP_PW_SERGE` — app password for the serge.utd mailbox
   - `GMAIL_APP_PW_SERGI` — app password for the sergi.utd mailbox (autoresponders)
   - `GMAIL_APP_PW_SERHII` — app password for the serhii.smortkin.utd mailbox (notifications)
   - (`GOOGLE_CREDENTIALS_JSON`, `ANTHROPIC_API_KEY` already exist from the parsers.)
2. **Dry-run test:** open each workflow in the Actions tab → **Run workflow** (manual).
   With `DRY_RUN` still true it prints the drafts + intended writes in the log. Review them.
3. **Go live:** Settings → Secrets and variables → Actions → **Variables** →
   add `DRY_RUN` = `false`. All chains now send on their schedules.
4. **Retire n8n:** once a full day of live GHA runs looks good, archive/delete the
   n8n workflows (they are already deactivated).

## Notes / assumptions carried from the n8n port
- Sender display name + HTML emails are now supported in `email_common.send_email`
  (auto-detects HTML, adds a plain-text fallback).
- `ECOM_SHEET_ID`/`ECOM_SHEET_TAB` are set in the ecom workflows to the B2B
  spreadsheet, tab "Ecom Contacts".
- Mailbox → app-password mapping is via `email_common`'s account convention;
  a chain skips any mailbox whose password secret is absent.
