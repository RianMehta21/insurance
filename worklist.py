#!/usr/bin/env python3
"""
Keeps data/manual_overrides.csv up to date as a to-do list.

WHY THIS EXISTS
---------------
review_queue.csv is a scrape diagnostic: it lists plans whose PDF the parser
could not read. It knows nothing about corrections typed by hand, so a plan
that has already been fixed reappears in it on every single run and the list
never gets shorter.

This script produces the list that actually shrinks. For every plan it works
out the value that will really show up in the workbook -

    manual override  >  scraped from the PDF  >  parsed from the plan label

- and a plan only needs attention if a field is still blank after all three.
Fixed plans drop off. New ones appear.

The output IS data/manual_overrides.csv, so there is one file to edit rather
than a worklist to copy into a separate overrides file. Rows already in it are
never changed or removed: the value columns are the user's, this script only
appends new skeleton rows and refreshes the read-only context columns.

Usage:
    python worklist.py
    python worklist.py --catalog out/plans_catalog.csv --benefits data/plan_benefits.csv
"""

import argparse
import csv
import os
import shutil

# Columns the user fills in. These map onto the Benefits sheet in the workbook.
VALUE_COLS = ["ward", "deductible", "geography", "annual_limit",
              "coinsurance", "lifetime"]

# Filled in from the catalog so the row is identifiable and the source document
# is one click away. Rewritten on every run; not meant to be edited.
CONTEXT_COLS = ["insurer", "plan_name", "plan_level", "plan_doc_url"]

# A plan lands on the list when any of these is still blank. The other value
# columns are useful to have but genuinely absent from some plan documents, so
# demanding them would mean a list that can never be finished.
KEY_COLS = ["ward", "deductible", "geography"]

COLUMNS = (["certification_no"] + CONTEXT_COLS + ["shortlist_hits", "still_missing"]
           + VALUE_COLS + ["note"])

# Client profiles used to work out which plans are worth chasing. Nobody is
# going to look up 234 deductibles by hand, and they should not have to: most
# of those plans never surface in a real comparison. Price every plan for a
# spread of ages and both genders, and count how often each one lands in the
# cheapest 20 of its own currency. A plan that never places is one you can
# leave blank indefinitely.
PROFILE_AGES = (25, 30, 35, 40, 45, 50, 55, 60, 65)
SHORTLIST_N = 20


def shortlist_frequency(premium_json, catalog):
    """cert -> how many client profiles put it in the cheapest SHORTLIST_N."""
    import json
    if not os.path.exists(premium_json):
        return {}
    with open(premium_json, encoding="utf-8") as f:
        prem = {r["certification-no"]: r
                for r in json.load(f)["certified-plans"]}

    hits = {}
    for age in PROFILE_AGES:
        for gender in ("M", "F"):
            priced = []
            for cert, cat in catalog.items():
                # Only plans that can actually be sold are worth chasing.
                if cat.get("sellable_new") != "Y":
                    continue
                rec = prem.get(cert)
                if not rec:
                    continue
                pm = rec.get("premium") or {}
                series = pm.get(f"S{gender}N") or pm.get(f"R{gender}N")
                if not series:
                    continue
                idx = age + 1 if cat.get("age_counting_method") == "N" else age
                v = series.get(str(idx))
                if v is not None:
                    priced.append((rec.get("currency", ""), v, cert))
            for cur in {p[0] for p in priced}:
                sub = sorted([p for p in priced if p[0] == cur])[:SHORTLIST_N]
                for _, _, cert in sub:
                    hits[cert] = hits.get(cert, 0) + 1
    return hits

# Where each value column can come from, best source first.
SCRAPED_FROM = {
    "ward": "ward_restriction",
    "deductible": "deductible_amount",
    "geography": "geographical_coverage",
    "annual_limit": "annual_benefit_limit",
    "coinsurance": "coinsurance_pct",
    "lifetime": "lifetime_benefit_limit",
}
CATALOG_FROM = {
    "ward": "ward_type",
    "deductible": "deductible_amount",
    "geography": "geographical_coverage",
}


def clean(v):
    return (v or "").strip()


def read_csv(path):
    if not os.path.exists(path):
        return [], []
    with open(path, encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        return list(rd), list(rd.fieldnames or [])


def effective(cert, col, overrides, scraped, catalog):
    """The value that will actually reach the workbook for this field."""
    o = overrides.get(cert, {})
    if clean(o.get(col)):
        return clean(o[col]), "override"
    s = scraped.get(cert, {})
    src = SCRAPED_FROM.get(col)
    if src and clean(s.get(src)):
        return clean(s[src]), "scraped"
    c = catalog.get(cert, {})
    src = CATALOG_FROM.get(col)
    if src and clean(c.get(src)):
        return clean(c[src]), "label"
    return "", "missing"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--catalog", default="out/plans_catalog.csv")
    ap.add_argument("--benefits", default="data/plan_benefits.csv")
    ap.add_argument("--overrides", default="data/manual_overrides.csv")
    ap.add_argument("--premiums", default="current/plan-premium.json",
                    help="used to rank rows by how often the plan actually "
                         "reaches a client shortlist")
    ap.add_argument("--dry-run", action="store_true",
                    help="report what would change without writing")
    args = ap.parse_args()

    cat_rows, _ = read_csv(args.catalog)
    if not cat_rows:
        raise SystemExit(f"no catalog at {args.catalog}; run build_vhis.py first")
    catalog = {r["certification_no"]: r for r in cat_rows}
    scraped = {r["certification_no"]: r
               for r in read_csv(args.benefits)[0]}
    hits = shortlist_frequency(args.premiums, catalog)
    over_rows, over_fields = read_csv(args.overrides)
    overrides = {r["certification_no"]: r for r in over_rows
                 if clean(r.get("certification_no"))}

    # Preserve any column the user added that we do not know about.
    extra = [f for f in over_fields if f not in COLUMNS]
    columns = COLUMNS + extra

    out, still_needed, newly_added, resolved = [], 0, 0, 0
    for cert, cat in catalog.items():
        missing = [c for c in KEY_COLS
                   if not effective(cert, c, overrides, scraped, catalog)[0]]
        existing = overrides.get(cert)

        if not missing and existing is None:
            continue                      # nothing wrong, nothing recorded
        if not missing and existing is not None:
            resolved += 1                 # user fixed it; keep their row

        row = {c: "" for c in columns}
        if existing:
            row.update({k: v for k, v in existing.items() if k in columns})
        else:
            newly_added += 1

        row["certification_no"] = cert
        row["insurer"] = cat.get("insurer", "")
        row["plan_name"] = cat.get("plan_name", "")
        row["plan_level"] = cat.get("plan_level_raw", "")
        row["plan_doc_url"] = cat.get("plan_doc_url", "")
        row["shortlist_hits"] = hits.get(cert, 0)
        row["still_missing"] = ";".join(missing)
        if missing:
            still_needed += 1
        out.append(row)

    # Outstanding first, then by how often the plan actually gets quoted.
    out.sort(key=lambda r: (not r["still_missing"],
                            -int(r["shortlist_hits"] or 0),
                            r["insurer"], r["plan_level"]))

    print(f"catalog plans        : {len(catalog)}")
    print(f"rows on the list     : {len(out)}")
    print(f"  still need filling : {still_needed}")
    print(f"  already filled in  : {resolved}")
    print(f"  newly added        : {newly_added}")

    if args.dry_run:
        print("\n--dry-run: nothing written")
        return

    if over_rows and os.path.exists(args.overrides):
        shutil.copy2(args.overrides, args.overrides + ".bak")

    os.makedirs(os.path.dirname(args.overrides) or ".", exist_ok=True)
    with open(args.overrides, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        w.writerows(out)
    print(f"\nwrote {args.overrides}")
    if still_needed:
        ranked = [r for r in out if r["still_missing"]]
        hot = [r for r in ranked if int(r["shortlist_hits"] or 0) > 0]
        print(f"\n{still_needed} row(s) still need filling, but they are not equal:")
        print(f"  {len(hot)} reach a client shortlist and are worth looking up")
        print(f"  {still_needed - len(hot)} never place in the cheapest "
              f"{SHORTLIST_N} for any profile - leave them")
        if hot:
            print("\nstart here:")
            for r in hot[:10]:
                print(f"  {r['shortlist_hits']:>3} hits  {r['insurer'][:26]:<28}"
                      f"{r['plan_level'][:26]:<28}{r['still_missing']}")


if __name__ == "__main__":
    main()
