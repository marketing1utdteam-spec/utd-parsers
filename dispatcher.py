#!/usr/bin/env python3
"""
dispatcher.py — single-scheduler orchestrator for all high-frequency chains.

WHY: GitHub Actions heavily throttles scheduled workflows. With 18 scheduled
workflows the whole repo got ~55 run-slots per DAY (each workflow ~5), so the
senders produced ~5 letters/day instead of the designed cadence. One scheduled
workflow receives all the slots; this dispatcher runs every tick and executes
every task that is due, with per-mailbox daily caps and minimum gaps (canon:
never batch, ~1 letter / 20+ min / mailbox).

Tasks:
  senders  — run when: inside window AND under daily cap AND min-gap passed.
             Each invocation sends at most 1 letter (module default limit).
  always   — autoresponders / reply-monitor / reminders: every tick (they are
             inbox-driven and cost nothing when there is no new mail).

State: data/dispatcher_state.json {task: {date, sent_today, last_sent_iso}}
(committed back by the workflow like every other state file).

Each task runs as a SUBPROCESS of its module with its own env (mailbox account
+ app password), so per-mailbox variants reuse the same module files.
"""

import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

STATE_DIR = os.environ.get("STATE_DIR", ".")
STATE_FILE = os.path.join(STATE_DIR, "dispatcher_state.json")

# All cold-outreach gating is in BELGIUM local time (Europe/Brussels, handles
# CET/CEST DST automatically). GitHub Actions runs this in the cloud on a cron,
# so sending happens 24/7 on Google's/GitHub's servers — the owner's computer
# does NOT need to be on. The windows below simply restrict WHEN cold mail goes.
BE = ZoneInfo("Europe/Brussels")
_UTC = datetime.now(timezone.utc)
NOW = _UTC.astimezone(BE)
TODAY = NOW.strftime("%Y-%m-%d")
WEEKDAY = NOW.weekday()  # 0=Mon .. 6=Sun  (Belgium)
HOUR = NOW.hour          # Belgium local hour
UTC_HOUR = _UTC.hour     # UTC hour — used to schedule parser runs 24/7

# Belgium business hours for cold outreach: Mon-Fri, 09:00-17:59.
BE_WORK = (9, 17)

# The GitHub schedule cron is throttled to ~2 firings/day, so the parsers (which
# should collect several times a day) are triggered from THIS reliable dispatcher
# instead. Each parser workflow is fired at the listed UTC hours, once per slot
# per day (tracked in dispatcher_state). Needs DISPATCH_PAT (GITHUB_TOKEN can't
# trigger another workflow). More b2b runs = more contacts, up to its CSE quota.
PARSER_SLOTS = {
    "parsers-b2b.yml":         [1, 6, 10, 14, 18],   # 5×/day (owner request)
    "parsers-ecom.yml":        [3, 15],              # 2×/day (quota-heavy collector)
    "parsers-influencers.yml": [7, 19],              # 2×/day
}
DISPATCH_PAT = os.environ.get("DISPATCH_PAT", "")
REPO = os.environ.get("GITHUB_REPOSITORY", "marketing1utdteam-spec/utd-parsers")

SERGEY = {"user_env": {}, "pw": ""}  # placeholders for readability

# ── task table ──────────────────────────────────────────────────────
# window_hours: inclusive UTC hour range; weekdays_only for B2B.
# cap: letters per day per this task (== per mailbox per chain).
# gap_min: minimum minutes since this task's previous send.
def _env_serge(extra=None):
    e = {"B2B_SENDER_USER": os.environ.get("UTD_MAIL_SERGE", ""),
         "ECOM_SENDER_EMAIL": os.environ.get("UTD_MAIL_SERGE", ""),
         "INFL_GMAIL_USER": os.environ.get("UTD_MAIL_SERGE", ""),
         "SENDER_APP_PW": os.environ.get("GMAIL_APP_PW_SERGE", "")}
    e.update(extra or {})
    return e

# All cold senders: Belgium business hours (BE_WORK), Mon-Fri only.
SENDERS = [
    {"task": "b2b_sergey", "script": "b2b_sender.py", "env": {},
     "window": BE_WORK, "weekdays_only": True, "cap": 20, "gap_min": 22},
    {"task": "b2b_serge", "script": "b2b_sender.py", "env": _env_serge(),
     "window": BE_WORK, "weekdays_only": True, "cap": 20, "gap_min": 22},
    {"task": "ecom_sergey", "script": "ecom_sender.py", "env": {},
     "window": BE_WORK, "weekdays_only": True, "cap": 30, "gap_min": 22},
    {"task": "ecom_serge", "script": "ecom_sender.py", "env": _env_serge(),
     "window": BE_WORK, "weekdays_only": True, "cap": 30, "gap_min": 22},
    {"task": "infl_sergey", "script": "influencer_sender.py", "env": {},
     "window": BE_WORK, "weekdays_only": True, "cap": 5, "gap_min": 75},
    {"task": "infl_serge", "script": "influencer_sender.py", "env": _env_serge(),
     "window": BE_WORK, "weekdays_only": True, "cap": 5, "gap_min": 75},
]

# ALWAYS tasks run every tick, 24/7. Autoresponders/reply-monitor are inbox-
# driven and cost nothing on an empty inbox. `business_only` tasks (proactive
# outbound reminders) are held to Belgium business hours like cold outreach.
ALWAYS = [
    {"task": "agency_auto", "script": "agency_autoresponder.py", "env": {}},
    {"task": "ecom_auto", "script": "ecom_autoresponder.py", "env": {}},
    {"task": "infl_auto", "script": "influencer_autoresponder.py", "env": {}},
    {"task": "reply_monitor", "script": "reply_monitor.py", "env": {}},
    {"task": "infl_reminders", "script": "influencer_reminders.py", "env": {},
     "business_only": True},
]


def load_state():
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(st):
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(STATE_FILE, "w") as f:
        json.dump(st, f, indent=1, sort_keys=True)


def run_script(script, extra_env, label):
    env = dict(os.environ)
    env.update(extra_env)
    print(f"\n──── {label}: {script} ────", flush=True)
    r = subprocess.run([sys.executable, script], env=env,
                       capture_output=True, text=True, timeout=1500)
    out = (r.stdout or "") + ("\n" + r.stderr if r.stderr else "")
    print(out[-4000:], flush=True)
    sent = 0
    for line in out.splitlines():
        if line.startswith("SUMMARY:"):
            try:
                sent = int(json.loads(line[8:]).get("sent", 0))
            except Exception:
                pass
    return sent, r.returncode


def trigger_parsers(st):
    """Fire each parser workflow at its scheduled UTC slots (once per slot/day),
    since GitHub's own cron is throttled. Uses DISPATCH_PAT. A slot only fires
    within a 2-hour catch-up window so a downtime never causes a burst."""
    import urllib.request
    if not DISPATCH_PAT:
        print("· parser triggers skipped (no DISPATCH_PAT)")
        return
    for wf, hours in PARSER_SLOTS.items():
        key = "parser:" + wf
        rec = st.get(key, {})
        if rec.get("date") != TODAY:
            rec = {"date": TODAY, "fired": []}
        due = [h for h in hours if 0 <= (UTC_HOUR - h) <= 1 and h not in rec["fired"]]
        for h in due:
            try:
                req = urllib.request.Request(
                    f"https://api.github.com/repos/{REPO}/actions/workflows/{wf}/dispatches",
                    data=json.dumps({"ref": "main"}).encode(),
                    headers={"Authorization": f"token {DISPATCH_PAT}",
                             "Accept": "application/vnd.github+json",
                             "User-Agent": "utd-dispatcher"},
                    method="POST")
                urllib.request.urlopen(req, timeout=30)
                rec["fired"].append(h)
                print(f"· triggered {wf} (UTC slot {h}:00)")
            except Exception as e:
                print(f"· parser trigger {wf} failed: {e}")
        st[key] = rec


def main():
    st = load_state()
    print(f"=== dispatcher | {NOW.isoformat()} | weekday={WEEKDAY} hour={HOUR}Z ===")

    for t in SENDERS:
        rec = st.get(t["task"], {})
        if rec.get("date") != TODAY:
            rec = {"date": TODAY, "sent_today": 0, "last_sent_iso": ""}
        reason = None
        lo, hi = t["window"]
        if t["weekdays_only"] and WEEKDAY >= 5:
            reason = "weekend"
        elif not (lo <= HOUR <= hi):
            reason = f"outside window {lo}-{hi}Z"
        elif rec["sent_today"] >= t["cap"]:
            reason = f"daily cap {t['cap']} reached"
        elif rec["last_sent_iso"]:
            mins = (NOW - datetime.fromisoformat(rec["last_sent_iso"])).total_seconds() / 60
            if mins < t["gap_min"]:
                reason = f"gap {mins:.0f}m < {t['gap_min']}m"
        if reason:
            print(f"· {t['task']}: skip ({reason})")
            st[t["task"]] = rec
            continue
        sent, rc = run_script(t["script"], t["env"], t["task"])
        if sent > 0:
            rec["sent_today"] += sent
            rec["last_sent_iso"] = NOW.isoformat()
        rec["date"] = TODAY
        st[t["task"]] = rec
        print(f"· {t['task']}: sent={sent} today={rec['sent_today']}/{t['cap']} rc={rc}")

    # Fire the parsers on their own reliable schedule (cloud, no throttling).
    trigger_parsers(st)

    in_business = (WEEKDAY < 5) and (BE_WORK[0] <= HOUR <= BE_WORK[1])
    for t in ALWAYS:
        if t.get("business_only") and not in_business:
            print(f"· {t['task']}: skip (outside Belgium business hours)")
            continue
        run_script(t["script"], t["env"], t["task"])

    save_state(st)
    print("\n=== dispatcher done ===")


if __name__ == "__main__":
    main()
