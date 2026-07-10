#!/usr/bin/env python3
"""
migrate_hash_state.py — ONE-SHOT PII scrubber for the committed data/*.json.

The repo is PUBLIC. The parsers used to persist RAW emails and RAW domains/URLs
into data/*.json for cross-run dedup, exposing the harvested contact lists to
anyone. The parsers have been changed to store SHA256 hashes instead; this
script rewrites the ALREADY-COMMITTED state files the same way, in place, so the
current files no longer contain any raw email or domain.

Dedup stays intact: hash(x) is deterministic, so hashed old state matches the
new hashed lookups. Counters, indices, query hashes and versions are preserved.

Run once:   python3 migrate_hash_state.py
Idempotent: re-running is safe (already-hashed values are detected and kept).
"""
import os, re, json, hashlib

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

HEX64 = re.compile(r"^[0-9a-f]{64}$")


def hash_key(s: str) -> str:
    return hashlib.sha256(str(s).lower().strip().encode()).hexdigest()


def is_hashed(s: str) -> bool:
    return isinstance(s, str) and bool(HEX64.match(s))


def load(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ── b2b / ecom state: {emails:{email->rec}, seen_domains:[...], ...} ──────────
def migrate_harvester_state(path, keep_meta):
    """keep_meta: list of non-PII value keys to retain in the email record."""
    if not os.path.exists(path):
        print(f"  (skip, missing) {path}")
        return
    s = load(path)

    # emails: key by hash, strip value down to non-PII metadata
    new_emails = {}
    for email, rec in s.get("emails", {}).items():
        k = email if is_hashed(email) else hash_key(email)
        if isinstance(rec, dict):
            rec = {kk: vv for kk, vv in rec.items() if kk in keep_meta}
        else:
            rec = {}
        new_emails[k] = rec
    s["emails"] = new_emails

    # seen_domains: list of raw domains -> list of hashes
    s["seen_domains"] = [d if is_hashed(d) else hash_key(d)
                         for d in s.get("seen_domains", [])]

    save(path, s)
    print(f"  ✓ {os.path.basename(path)}: {len(new_emails)} emails hashed, "
          f"{len(s['seen_domains'])} domains hashed")


# ── visited list: [url, url, ...] -> [hash, hash, ...] ───────────────────────
def migrate_visited(path):
    if not os.path.exists(path):
        print(f"  (skip, missing) {path}")
        return
    v = load(path)
    out = [u if is_hashed(u) else hash_key(u) for u in v]
    save(path, out)
    print(f"  ✓ {os.path.basename(path)}: {len(out)} URLs hashed")


# ── influencer registry: {creator_key -> email_or_status} ────────────────────
def migrate_creator_registry(path):
    if not os.path.exists(path):
        print(f"  (skip, missing) {path}")
        return
    reg = load(path)
    out = {}
    hashed_keys = hashed_vals = kept_status = 0
    for key, val in reg.items():
        # hash the creator key (contains raw domain / channel-id / profile url)
        nk = key if is_hashed(key) else hash_key(key)
        hashed_keys += 0 if is_hashed(key) else 1
        # value: keep "skipped:*" statuses as-is; hash raw emails; keep hashes
        if isinstance(val, str) and val.startswith("skipped"):
            nv = val
            kept_status += 1
        elif is_hashed(val):
            nv = val
        elif isinstance(val, str) and "@" in val:
            nv = hash_key(val)
            hashed_vals += 1
        else:
            nv = val  # unknown non-PII scalar — leave as-is
        out[nk] = nv
    save(path, out)
    print(f"  ✓ {os.path.basename(path)}: {len(out)} entries "
          f"({hashed_keys} keys hashed, {hashed_vals} emails hashed, "
          f"{kept_status} statuses kept)")


def main():
    print("Migrating committed state to hashed (PII-free) form …\n")
    print("b2b harvester:")
    migrate_harvester_state(os.path.join(DATA_DIR, "harvester_v4_state.json"),
                            keep_meta={"status", "score", "date"})
    migrate_visited(os.path.join(DATA_DIR, "harvester_v4_visited.json"))

    print("\necom harvester:")
    # 'theme' excluded on purpose: scraped theme strings sometimes embed the
    # store domain (e.g. "paw.com/main"), which would re-leak PII.
    migrate_harvester_state(os.path.join(DATA_DIR, "ecom_v1_state.json"),
                            keep_meta={"status", "score", "industry",
                                       "date", "source"})
    migrate_visited(os.path.join(DATA_DIR, "ecom_v1_visited.json"))

    print("\ninfluencer scraper:")
    migrate_creator_registry(os.path.join(DATA_DIR, "processed_creators_v4.json"))
    # scraper_state_v4.json holds only counters/indices — no PII, left untouched.
    print("  (scraper_state_v4.json left as-is — counters only, no PII)")

    print("\nDone.")


if __name__ == "__main__":
    main()
