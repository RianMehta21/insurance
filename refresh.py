#!/usr/bin/env python3
"""
VHIS data refresh + change detection.

Answers: "how do I know when data.gov.hk has new files?"

Strategy (two independent signals, belt and braces):
  1. Content hash. Download each resource, SHA-256 it, compare to the hash
     stored from last run. Authoritative and immediate.
  2. data.gov.hk Historical Archive API. Lists archived versions of a file
     within a date range. Lags ~1 day but gives an audit trail and tells you
     WHEN a change landed, which the hash alone cannot.

On change, writes a dated snapshot, rebuilds, and emits a human-readable diff
so you know what actually moved before you tell a client anything.

Usage:
    python refresh.py --check          # detect only, exit 1 if changed
    python refresh.py --pull           # detect, snapshot, and diff
    python refresh.py --history 90     # what changed in the last 90 days
"""

import argparse
import hashlib
import json
import os
import io
import shutil
import sqlite3
import sys
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime, timedelta

# ---------------------------------------------------------------------------
# CONFIGURATION
#
# IMPORTANT: fill these in from the "Copy API string" / resource download link
# on each data.gov.hk dataset page. They are stable per resource. Do not guess
# them; open the four dataset pages and copy the real URLs.
#
#   Certified Plans (Standard + Flexi):
#     https://data.gov.hk/en-data/dataset/hk-hhb-hhbvhis-vhis-certified-plan
#   Standard Premiums:
#     https://data.gov.hk/en-data/dataset/hk-hhb-hhbvhis-vhis-standard-premium
#   VHIS Providers:
#     https://data.gov.hk/en-data/dataset/hk-hhb-hhbvhis-vhis-provider
# ---------------------------------------------------------------------------
# Confirmed from the data.gov.hk dataset pages. Note that four FILES come from
# three DATASETS: the "Certified Plan" dataset carries both standard-plans and
# flexi-plans as separate resources.
#
# Note also that the premium resource is a ZIP, not a raw JSON. fetch_resource()
# unwraps it automatically.
RESOURCES = {
    "standard-plans.json":          "https://www.vhis.gov.hk/public/data/standard-plans.json",
    "flexi-plans.json":             "https://www.vhis.gov.hk/public/data/flexi-plans.json",
    "plan-premium.json":            "https://www.vhis.gov.hk/public/data/plan-premium.zip",
    # VERIFY this one against the provider dataset page before first run:
    # https://data.gov.hk/en-data/dataset/hk-hhb-hhbvhis-vhis-provider
    "psi-registered-providers.json": "https://www.vhis.gov.hk/public/data/psi-registered-providers.json",
}

ARCHIVE_API = "https://api.data.gov.hk/v1/historical-archive/list-file-versions"
STATE_DB = "vhis_state.db"
SNAPSHOT_DIR = "snapshots"
CURRENT_DIR = "current"

# Reduced form of the source data, committed so the next run has something to
# diff against. See write_state_file(). Roughly 60 KB versus 5 MB per snapshot.
STATE_FILE = os.path.join("data", "plan_state.json")

# Exit codes. CI distinguishes these, so do not renumber casually.
EXIT_NO_CHANGE = 0
EXIT_CHANGED = 1
EXIT_FETCH_FAILED = 2

UA = {"User-Agent": "vhis-compare/1.0 (personal plan comparison tool)"}


# ---------------------------------------------------------------------------
# state
# ---------------------------------------------------------------------------
def init_state(path=STATE_DB):
    con = sqlite3.connect(path)
    con.execute("""
        CREATE TABLE IF NOT EXISTS file_state (
            filename    TEXT PRIMARY KEY,
            sha256      TEXT NOT NULL,
            bytes       INTEGER,
            checked_at  TEXT NOT NULL,
            changed_at  TEXT NOT NULL
        )""")
    con.execute("""
        CREATE TABLE IF NOT EXISTS refresh_log (
            run_at      TEXT NOT NULL,
            filename    TEXT NOT NULL,
            outcome     TEXT NOT NULL,
            note        TEXT
        )""")
    con.commit()
    return con


def sha256_bytes(b):
    return hashlib.sha256(b).hexdigest()


def fetch(url, timeout=60):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def fetch_resource(url, want_name, timeout=90):
    """Fetch a resource and return raw JSON bytes.

    Some VHIS resources are published zipped (plan-premium is a .zip holding a
    .json) and some are raw. Detect by magic bytes rather than by file
    extension, because the extension on the URL is not a promise.

    Returns (json_bytes, note).
    """
    blob = fetch(url, timeout=timeout)

    if blob[:2] == b"PK":                       # ZIP magic number
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            names = [n for n in z.namelist() if n.lower().endswith(".json")]
            if not names:
                raise ValueError(f"zip contains no .json member: {z.namelist()[:5]}")
            # prefer an exact filename match, else the only/largest member
            exact = [n for n in names if os.path.basename(n) == want_name]
            member = exact[0] if exact else max(
                names, key=lambda n: z.getinfo(n).file_size)
            return z.read(member), f"unzipped {member}"
    return blob, ""


# ---------------------------------------------------------------------------
# signal 1: content hash
#
# probe() is deliberately READ-ONLY with respect to the state DB.
#
# The state must not move until we have actually committed to the new data.
# An earlier version updated the stored hash inside --check, which meant the
# --pull that CI runs immediately afterwards compared the new bytes against the
# hash --check had just written, concluded "unchanged", and returned before
# snapshotting or diffing. The workbook still built, so it looked fine, but the
# change report was empty on every single run.
# ---------------------------------------------------------------------------
class FetchError(Exception):
    """A resource could not be fetched or was not valid JSON."""


def probe(con):
    """Download every resource and classify it against stored state.

    Returns (results, failures) where results is a list of dicts. Writes
    nothing. Raising is left to the caller so that --check and --pull can
    report identically but act differently.
    """
    results, failures = [], []

    for fname, url in RESOURCES.items():
        if url.startswith("PASTE_"):
            print(f"  {fname:<32} SKIPPED (no URL configured)")
            failures.append((fname, "no URL configured"))
            continue
        try:
            blob, note = fetch_resource(url, fname)
        except Exception as e:
            print(f"  {fname:<32} FETCH FAILED: {e}")
            failures.append((fname, f"fetch failed: {e}"))
            continue

        # Reject anything that is not parseable JSON. A 200 response carrying
        # an HTML error page would otherwise be hashed and stored as "new data".
        try:
            json.loads(blob)
        except Exception:
            print(f"  {fname:<32} REJECTED (response is not valid JSON)")
            failures.append((fname, f"not valid JSON ({len(blob)} bytes)"))
            continue

        digest = sha256_bytes(blob)
        row = con.execute("SELECT sha256 FROM file_state WHERE filename=?",
                          (fname,)).fetchone()
        if row is None:
            status = "NEW"
        elif row[0] != digest:
            status = "CHANGED"
        else:
            status = "unchanged"

        print(f"  {fname:<32} {status:<10} {len(blob):>10,} bytes  "
              f"{digest[:12]}  {note}")
        results.append({"filename": fname, "blob": blob, "digest": digest,
                        "status": status, "note": note})
    return results, failures


def persist(con, results):
    """Record hashes. Called only once the new bytes are actually on disk."""
    now = datetime.now().isoformat(timespec="seconds")
    for r in results:
        fname, digest, status = r["filename"], r["digest"], r["status"]
        if status == "unchanged":
            prev = con.execute("SELECT changed_at FROM file_state WHERE filename=?",
                               (fname,)).fetchone()
            changed_at = prev[0] if prev else now
        else:
            changed_at = now
        con.execute(
            "INSERT INTO file_state VALUES (?,?,?,?,?) "
            "ON CONFLICT(filename) DO UPDATE SET "
            "sha256=excluded.sha256, bytes=excluded.bytes, "
            "checked_at=excluded.checked_at, changed_at=excluded.changed_at",
            (fname, digest, len(r["blob"]), now, changed_at))
        con.execute("INSERT INTO refresh_log VALUES (?,?,?,?)",
                    (now, fname, status.lower(), digest[:12]))
    con.commit()


def log_failures(con, failures):
    now = datetime.now().isoformat(timespec="seconds")
    for fname, note in failures:
        con.execute("INSERT INTO refresh_log VALUES (?,?,?,?)",
                    (now, fname, "failed", note))
    con.commit()


def write_current(results):
    """Write every fetched resource into current/, whatever its status.

    Unconditional on purpose: current/ is gitignored, so on a fresh CI checkout
    it is empty even when the hashes say nothing changed.
    """
    os.makedirs(CURRENT_DIR, exist_ok=True)
    for r in results:
        with open(os.path.join(CURRENT_DIR, r["filename"]), "wb") as f:
            f.write(r["blob"])


# ---------------------------------------------------------------------------
# signal 2: data.gov.hk historical archive
# ---------------------------------------------------------------------------
def archive_versions(resource_url, days=90):
    """Ask data.gov.hk when this file was archived. Lags roughly one day."""
    end = datetime.now()
    start = end - timedelta(days=days)
    q = urllib.parse.urlencode({
        "url": resource_url,
        "start": start.strftime("%Y%m%d"),
        "end": end.strftime("%Y%m%d"),
    })
    try:
        return json.loads(fetch(f"{ARCHIVE_API}?{q}", timeout=45))
    except Exception as e:
        return {"error": str(e)}


def show_history(days):
    for fname, url in RESOURCES.items():
        if url.startswith("PASTE_"):
            continue
        res = archive_versions(url, days)
        vers = res.get("timestamps") or res.get("versions") or []
        print(f"\n{fname}: {len(vers)} archived version(s) in {days} days")
        for v in vers[-10:]:
            print("   ", v)
        if "error" in res:
            print("    error:", res["error"])


# ---------------------------------------------------------------------------
# snapshotting + diff
# ---------------------------------------------------------------------------
def snapshot():
    stamp = datetime.now().strftime("%Y%m%d")
    dest = os.path.join(SNAPSHOT_DIR, stamp)
    os.makedirs(dest, exist_ok=True)
    for fname in RESOURCES:
        src = os.path.join(CURRENT_DIR, fname)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(dest, fname))
    print(f"\nSnapshot written to {dest}/")
    return dest


# ---------------------------------------------------------------------------
# compact state: what the diff actually reads
#
# diff_snapshots() only ever looks at five fields per plan variant. Keeping
# full 5 MB copies of the source JSON purely to recompute those is why the repo
# would have grown ~260 MB/year. STATE_FILE holds the reduced form instead:
# ~60 KB, committed on every run, and enough to diff against next week.
# ---------------------------------------------------------------------------
def write_state_file(state, path=STATE_FILE):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "variants": state,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=1, sort_keys=True, ensure_ascii=False)
    print(f"Plan state written to {path} ({len(state)} variants)")


def read_state_file(path=STATE_FILE):
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f).get("variants") or {}
    except Exception as e:
        print(f"  could not read {path}: {e}")
        return None


def load_variants(directory):
    """cert-no -> (plan_name, insurer, level, prem_date, first-adult-premium)."""
    out = {}
    try:
        plans = []
        for f in ("standard-plans.json", "flexi-plans.json"):
            p = os.path.join(directory, f)
            if os.path.exists(p):
                plans += json.load(open(p, encoding="utf-8"))["certified-plans"]
        prem_path = os.path.join(directory, "plan-premium.json")
        prem = {}
        if os.path.exists(prem_path):
            prem = {r["certification-no"]: r for r in
                    json.load(open(prem_path, encoding="utf-8"))["certified-plans"]}
    except Exception as e:
        print("  could not load", directory, e)
        return out

    for p in plans:
        for pi in p["plan-info-certified"]:
            cert = pi["certification-no"]
            r = prem.get(cert, {})
            pm = (r.get("premium") or {})
            series = next((pm[k] for k in ("SMN", "SFN", "RMN", "RFN")
                           if pm.get(k)), {}) or {}
            out[cert] = {
                "plan": p["plan-name"]["en"],
                "insurer": p["company-name"]["en"],
                "level": pi["plan-level"]["en"],
                "prem_date": r.get("prem-date", ""),
                "age40": series.get("40"),
            }
    return out


def diff_states(a, b):
    """The report you read before quoting anyone.

    Takes two reduced state dicts (cert -> fields), not directories, so it
    works against the committed state file as well as against snapshots.
    """
    added = sorted(set(b) - set(a))
    removed = sorted(set(a) - set(b))
    repriced, recertified = [], []

    for c in sorted(set(a) & set(b)):
        if a[c]["prem_date"] != b[c]["prem_date"]:
            recertified.append((c, a[c]["prem_date"], b[c]["prem_date"]))
        oa, ob = a[c]["age40"], b[c]["age40"]
        if oa and ob and abs(oa - ob) > 1e-9:
            repriced.append((c, b[c]["insurer"], b[c]["level"], oa, ob,
                             (ob - oa) / oa * 100))

    print("\n" + "=" * 70)
    print("CHANGE REPORT")
    print("=" * 70)
    print(f"variants before : {len(a)}")
    print(f"variants after  : {len(b)}")
    print(f"new             : {len(added)}")
    print(f"withdrawn       : {len(removed)}")
    print(f"premium revised : {len(repriced)}  (measured at age 40)")
    print(f"new premium date: {len(recertified)}")

    if added:
        print("\nNEW VARIANTS (need a PDF scrape):")
        for c in added[:20]:
            print(f"  {c}  {b[c]['insurer'][:30]:<32} {b[c]['level'][:34]}")
        if len(added) > 20:
            print(f"  ... and {len(added)-20} more")

    if removed:
        print("\nWITHDRAWN (do not quote these):")
        for c in removed[:20]:
            print(f"  {c}  {a[c]['insurer'][:30]:<32} {a[c]['level'][:34]}")

    if repriced:
        print("\nBIGGEST PREMIUM MOVES (age 40):")
        repriced.sort(key=lambda r: -abs(r[5]))
        for c, ins, lvl, oa, ob, pct in repriced[:15]:
            print(f"  {pct:+7.1f}%  {oa:>10,.0f} -> {ob:>10,.0f}  "
                  f"{ins[:26]:<28} {lvl[:28]}")

    return {"added": added, "removed": removed, "repriced": repriced}


def latest_snapshot_before(new_dir):
    if not os.path.isdir(SNAPSHOT_DIR):
        return None
    dirs = sorted(d for d in os.listdir(SNAPSHOT_DIR)
                  if os.path.isdir(os.path.join(SNAPSHOT_DIR, d)))
    dirs = [d for d in dirs if os.path.join(SNAPSHOT_DIR, d) != new_dir]
    return os.path.join(SNAPSHOT_DIR, dirs[-1]) if dirs else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true",
                    help="detect changes only; exit 1 if anything changed. "
                         "Never writes state.")
    ap.add_argument("--pull", action="store_true",
                    help="download, write current/, diff, and record state")
    ap.add_argument("--history", type=int, metavar="DAYS",
                    help="list archived versions from data.gov.hk")
    ap.add_argument("--snapshot", action="store_true",
                    help="also keep a full dated copy of the source JSON in "
                         "snapshots/ (not needed for the diff; see STATE_FILE)")
    args = ap.parse_args()

    con = init_state()

    if args.history:
        show_history(args.history)
        return

    print("Checking data.gov.hk resources...")
    results, failures = probe(con)
    log_failures(con, failures)

    # A failed download is not "no change". Saying so would let the pipeline
    # publish last week's workbook forever while reporting success, which is
    # the one failure mode nobody would notice.
    if failures:
        print(f"\n{len(failures)} of {len(RESOURCES)} resource(s) could not be "
              f"retrieved:")
        for fname, note in failures:
            print(f"  {fname}: {note}")
        print("\nRefusing to report a result from incomplete data.")
        sys.exit(EXIT_FETCH_FAILED)

    changed = [r["filename"] for r in results if r["status"] != "unchanged"]

    if args.check:
        # Read-only by contract. Report and leave the state exactly as found.
        if changed:
            print(f"\n{len(changed)} file(s) changed: {', '.join(changed)}")
            sys.exit(EXIT_CHANGED)
        print("\nNo changes. Nothing to rebuild.")
        sys.exit(EXIT_NO_CHANGE)

    if not args.pull:
        print("\nNothing to do. Pass --check or --pull.")
        sys.exit(EXIT_NO_CHANGE)

    # ---- --pull -----------------------------------------------------------
    # Read the previous state BEFORE overwriting it, or there is nothing left
    # to diff against.
    old_state = read_state_file()
    write_current(results)
    new_state = load_variants(CURRENT_DIR)

    if changed:
        print(f"\n{len(changed)} file(s) changed: {', '.join(changed)}")
    else:
        print("\nNo content change, but current/ has been refreshed.")

    if old_state is None:
        print("\nNo previous state on record, so there is nothing to diff "
              "against. This run establishes the baseline.")
    else:
        diff_states(old_state, new_state)

    if args.snapshot:
        snapshot()

    write_state_file(new_state)
    persist(con, results)

    print("\nNext: python build_vhis.py --data-dir current --out-dir out")
    # --pull exits 0 whenever it completed its work. Whether anything changed
    # is what --check is for; conflating the two here would make the CI step
    # fail on exactly the runs that did the most.
    sys.exit(EXIT_NO_CHANGE)


if __name__ == "__main__":
    main()
