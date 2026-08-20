#!/usr/bin/env python3
"""
VHIS plan comparison pipeline.

Flattens the four Hong Kong VHIS open-data JSON files into CSVs that can be
opened in Excel, filtered, and used to compare plans for a client.

Inputs (data.gov.hk / vhis.gov.hk):
    standard-plans.json          List of Certified Standard Plans
    flexi-plans.json             List of Certified Flexi Plans
    plan-premium.json            Standard Premiums of Certified Plans
    psi-registered-providers.json  Registered VHIS providers

Outputs:
    plans_catalog.csv    one row per certified plan variant (the dimension table)
    premiums_long.csv    tidy premium matrix: cert x basis x gender x smoker x age
    quote.csv            per-client comparison table (age/gender/smoker applied)
    unparsed_levels.csv  plan-level strings we could not parse (PDF scrape targets)

Usage:
    python build_vhis.py --data-dir . --out-dir ./out
    python build_vhis.py --age 35 --gender M --smoker N --basis S
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import Counter, defaultdict

VHIS_BASE = "https://www.vhis.gov.hk"

# ---------------------------------------------------------------------------
# Premium series decoding
# ---------------------------------------------------------------------------
# Each premium record carries up to 8 series keyed <basis><gender><smoker>:
#   basis   S = Standard premium for NEW applications
#           R = Renewal premium (present for 184/579 variants; always <= S)
#   gender  M / F
#   smoker  N = non-smoker, Y = smoker
#           (verified empirically: Y and N diverge first at exactly age 18
#            in 40 of the 46 variants where they differ at all)
BASES = ("S", "R")
GENDERS = ("M", "F")
SMOKERS = ("N", "Y")

# Age-counting method, per plan. Standard actuarial conventions.
AGE_METHOD_LABEL = {
    "A": "Age last birthday (attained)",
    "N": "Age next birthday",
    "R": "Age nearest birthday",
}


def availability(plan, plan_info):
    """Can this variant actually be sold to a new client today?

    The source data carries three separate withdrawal signals and none of them
    is the absence of a premium table, so a de-registered plan still prices
    perfectly and looks identical to a live one in a sorted list. Collapsing
    them into one column keeps a withdrawn plan from being quoted by accident.

      de-reg       certification withdrawn by the Health Bureau
      unavailable  insurer has stopped offering it
      renewal-only 'Y' existing policyholders may renew, no new business
                   'P' partially restricted (some levels closed)
    """
    ro = str(plan_info.get("renewal-only", "") or plan.get("renewal-only", "")).upper()
    if str(plan.get("de-reg", "")).upper() == "Y":
        return "De-registered", "N"
    if str(plan.get("unavailable", "")).upper() == "Y":
        return "Withdrawn", "N"
    if ro == "Y":
        return "Renewal only", "N"
    if ro == "P":
        return "Partly restricted", "P"
    return "Open", "Y"

# ---------------------------------------------------------------------------
# plan-level free-text parsing
# ---------------------------------------------------------------------------
DEDUCTIBLE_RE = re.compile(
    r"(HKD|USD|RMB|CNY)\s*\$?\s*([\d,]+)\s*(?:Deductible|deductible)", re.I
)
DEDUCTIBLE_ALT_RE = re.compile(
    r"(?:Deductible|deductible)\s*[:\-]?\s*(HKD|USD|RMB|CNY)\s*\$?\s*([\d,]+)", re.I
)

WARD_PATTERNS = [
    (r"\bsuite\b", "Suite"),
    (r"\bstandard\s*private\b", "Standard Private"),
    (r"\bsemi[\s\-]?private\b", "Semi-private"),
    (r"\bprivate\b", "Private"),
    (r"\bward\b", "Ward"),
]

GEO_PATTERNS = [
    (r"worldwide\s*(?:excluding|except|excl\.?)\s*(?:the\s*)?(?:usa|us|united states)"
     r"(?:\s*(?:and|&)\s*canada)?", "Worldwide excl. USA"),
    (r"\bworldwide\b", "Worldwide"),
    (r"\bgreater\s*china\b", "Greater China"),
    (r"\basia[\s\-]*pacific\b", "Asia Pacific"),
    (r"\basia\b", "Asia"),
    (r"\bhong\s*kong\b", "Hong Kong"),
    (r"\bmacau\b", "Hong Kong/Macau"),
]


def money_to_int(text):
    try:
        return int(text.replace(",", ""))
    except ValueError:
        return None


def parse_plan_level(level_en, currency):
    """Best-effort extraction of deductible / ward / geography from the
    free-text plan-level label. Returns dict with None where not derivable."""
    out = {
        "deductible_amount": None,
        "deductible_currency": None,
        "ward_type": None,
        "geographical_coverage": None,
        "has_smm": None,
        "parse_confidence": None,
    }
    if not level_en:
        return out

    text = level_en.strip()
    hits = 0

    m = DEDUCTIBLE_RE.search(text) or DEDUCTIBLE_ALT_RE.search(text)
    if m:
        out["deductible_currency"] = m.group(1).upper()
        out["deductible_amount"] = money_to_int(m.group(2))
        hits += 1
    elif re.search(r"\bno\s*deductible\b|\bnil\s*deductible\b", text, re.I):
        out["deductible_currency"] = currency
        out["deductible_amount"] = 0
        hits += 1

    for pat, label in WARD_PATTERNS:
        if re.search(pat, text, re.I):
            out["ward_type"] = label
            hits += 1
            break

    for pat, label in GEO_PATTERNS:
        if re.search(pat, text, re.I):
            out["geographical_coverage"] = label
            hits += 1
            break

    out["has_smm"] = "Y" if re.search(r"\bSMM\b|supplementary\s*major\s*medical", text, re.I) else "N"

    if hits >= 2:
        out["parse_confidence"] = "high"
    elif hits == 1:
        out["parse_confidence"] = "partial"
    else:
        out["parse_confidence"] = "none"
    return out


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------
def en(node, default=""):
    if isinstance(node, dict):
        return node.get("en", default)
    return node if node is not None else default


def abs_url(path):
    """Doc URL fields are sometimes a plain string, sometimes a {en, zh-hk, zh-cn}
    dict, depending on the field and the record. Normalise both."""
    if isinstance(path, dict):
        path = path.get("en", "")
    if not path or not isinstance(path, str):
        return ""
    return VHIS_BASE + path if path.startswith("/") else path


def load(data_dir):
    def rd(name):
        with open(os.path.join(data_dir, name), encoding="utf-8") as f:
            return json.load(f)

    return (
        rd("standard-plans.json")["certified-plans"],
        rd("flexi-plans.json")["certified-plans"],
        rd("plan-premium.json")["certified-plans"],
        rd("psi-registered-providers.json")["providers"],
    )


def build_catalog(std, flexi, prem_by_cert, providers):
    """One row per certified plan variant."""
    alias = {}
    reg = {}
    for p in providers:
        name = en(p["company-name"])
        alias[name] = en(p.get("alias", {}))
        reg[name] = en(p.get("registration-number", {}))

    rows = []
    unparsed = []
    for plan_type, coll in (("Standard", std), ("Flexi", flexi)):
        for plan in coll:
            company = en(plan["company-name"])
            for pi in plan["plan-info-certified"]:
                cert = pi["certification-no"]
                pr = prem_by_cert.get(cert, {})
                level = en(pi["plan-level"])
                currency = pi.get("currency", "")
                parsed = parse_plan_level(level, currency)

                series = {
                    f"{b}{g}{s}": bool((pr.get("premium") or {}).get(f"{b}{g}{s}"))
                    for b in BASES for g in GENDERS for s in SMOKERS
                }
                ages = []
                for k, present in series.items():
                    if present:
                        d = pr["premium"][k]
                        ages += [int(a) for a, v in d.items() if v is not None]

                row = {
                    "certification_no": cert,
                    "product_id": pi.get("product-id", ""),
                    "plan_type": plan_type,
                    "insurer": company,
                    "insurer_alias": alias.get(company, ""),
                    "insurer_reg_no": reg.get(company, ""),
                    "plan_name": en(plan["plan-name"]),
                    "plan_level_raw": level,
                    "ward_type": parsed["ward_type"] or "",
                    "deductible_amount": parsed["deductible_amount"]
                    if parsed["deductible_amount"] is not None else "",
                    "deductible_currency": parsed["deductible_currency"] or "",
                    "geographical_coverage": parsed["geographical_coverage"] or "",
                    "has_smm": parsed["has_smm"] or "",
                    "level_parse_confidence": parsed["parse_confidence"],
                    "currency": currency,
                    "plan_date": en(plan["plan-date"]),
                    "earliest_plan_date": en(plan["earliest-plan-date"]),
                    "premium_effective_date": pr.get("prem-date", ""),
                    "age_counting_method": pr.get("age-counting-method", ""),
                    "age_counting_method_label": AGE_METHOD_LABEL.get(
                        pr.get("age-counting-method", ""), ""),
                    "min_age": min(ages) if ages else "",
                    "max_age": max(ages) if ages else "",
                    "has_standard_basis": "Y" if series["SMN"] or series["SFN"] else "N",
                    "has_renewal_basis": "Y" if series["RMN"] or series["RFN"] else "N",
                    "smoker_rated": "",  # filled below
                    "availability": availability(plan, pi)[0],
                    "sellable_new": availability(plan, pi)[1],
                    "renewal_only": pi.get("renewal-only", "") or plan.get("renewal-only", ""),
                    "de_reg": plan.get("de-reg", ""),
                    "unavailable": plan.get("unavailable", ""),
                    "has_other_info": plan.get("has-other-info", ""),
                    "has_historical_plan": plan.get("has-historical-plan", ""),
                    "remarks": en(plan.get("remarks", "")),
                    "plan_doc_url": abs_url(en(pi.get("plan-doc-url", {}))),
                    "premium_doc_url": abs_url(en(pi.get("premium-doc-url", {}))),
                    "tax_doc_url": abs_url(plan.get("tax-doc-url", "")),
                    "hospital_def_doc_url": abs_url(plan.get("def-doc-url", "")),
                }

                # smoker rating: does Y differ from N anywhere?
                smoker_rated = "N"
                pm = pr.get("premium") or {}
                for b in BASES:
                    for g in GENDERS:
                        a, c = pm.get(f"{b}{g}N"), pm.get(f"{b}{g}Y")
                        if a and c and any(
                            a.get(k) is not None and c.get(k) is not None
                            and abs(a[k] - c[k]) > 1e-9 for k in a):
                            smoker_rated = "Y"
                row["smoker_rated"] = smoker_rated

                rows.append(row)
                if parsed["parse_confidence"] == "none":
                    unparsed.append({
                        "certification_no": cert,
                        "insurer": company,
                        "plan_name": en(plan["plan-name"]),
                        "plan_level_raw": level,
                        "plan_doc_url": row["plan_doc_url"],
                    })
    return rows, unparsed


def build_premiums_long(prem):
    rows = []
    for r in prem:
        cert = r["certification-no"]
        pm = r["premium"] or {}
        for b in BASES:
            for g in GENDERS:
                for s in SMOKERS:
                    d = pm.get(f"{b}{g}{s}")
                    if not d:
                        continue
                    for age, val in d.items():
                        if val is None:
                            continue
                        rows.append({
                            "certification_no": cert,
                            "basis": b,
                            "gender": g,
                            "smoker": s,
                            "age": int(age),
                            "annual_premium": val,
                            "currency": r["currency"],
                            "prem_date": r["prem-date"],
                        })
    rows.sort(key=lambda x: (x["certification_no"], x["basis"], x["gender"],
                             x["smoker"], x["age"]))
    return rows


# ---------------------------------------------------------------------------
# Quoting
# ---------------------------------------------------------------------------
def insurance_age(actual_age, method):
    """Map a client's actual age to the age index each insurer's table uses.

    'A' attained / age last birthday  -> actual age
    'N' age next birthday             -> actual age + 1
    'R' age nearest birthday          -> needs DOB to be exact; approximated
                                         as actual age here.
    """
    if method == "N":
        return actual_age + 1, True
    if method == "R":
        return actual_age, False  # approximation flag
    return actual_age, True


def build_quote(catalog, prem_by_cert, age, gender, smoker, basis, horizon=10):
    out = []
    for row in catalog:
        cert = row["certification_no"]
        pm = (prem_by_cert.get(cert) or {}).get("premium") or {}
        method = row["age_counting_method"]
        idx, exact = insurance_age(age, method)

        chosen_basis = basis
        series = pm.get(f"{basis}{gender}{smoker}")
        if not series:
            other = "R" if basis == "S" else "S"
            series = pm.get(f"{other}{gender}{smoker}")
            chosen_basis = other if series else ""
        if not series:
            continue

        first = series.get(str(idx))
        window = [series.get(str(idx + i)) for i in range(horizon)]
        have = [v for v in window if v is not None]
        truncated = len(have) < horizon

        q = dict(row)
        q.update({
            "quote_actual_age": age,
            "quote_table_age": idx,
            "quote_age_exact": "Y" if exact else "APPROX",
            "quote_gender": gender,
            "quote_smoker": smoker,
            "quote_basis": chosen_basis,
            "first_year_premium": first if first is not None else "",
            "ten_year_total": round(sum(have), 2) if have else "",
            "ten_year_avg": round(sum(have) / len(have), 2) if have else "",
            "ten_year_years_available": len(have),
            "ten_year_truncated": "Y" if truncated else "N",
        })
        out.append(q)

    # rank on first-year premium within currency
    for cur in {r["currency"] for r in out}:
        subset = [r for r in out if r["currency"] == cur and r["first_year_premium"] != ""]
        subset.sort(key=lambda r: r["first_year_premium"])
        for i, r in enumerate(subset, 1):
            r["rank_first_year_in_currency"] = i
    for r in out:
        r.setdefault("rank_first_year_in_currency", "")
    return out


def write_csv(path, rows, fieldnames=None):
    if not rows:
        return
    fieldnames = fieldnames or list(rows[0].keys())
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default="current",
                    help="directory holding the four source JSON files")
    ap.add_argument("--out-dir", default="out")
    ap.add_argument("--age", type=int, default=35)
    ap.add_argument("--gender", choices=["M", "F"], default="M")
    ap.add_argument("--smoker", choices=["N", "Y"], default="N")
    ap.add_argument("--basis", choices=["S", "R"], default="S")
    args = ap.parse_args()

    std, flexi, prem, providers = load(args.data_dir)
    prem_by_cert = {r["certification-no"]: r for r in prem}

    catalog, unparsed = build_catalog(std, flexi, prem_by_cert, providers)
    long_rows = build_premiums_long(prem)
    quote = build_quote(catalog, prem_by_cert, args.age, args.gender,
                        args.smoker, args.basis)

    write_csv(os.path.join(args.out_dir, "plans_catalog.csv"), catalog)
    write_csv(os.path.join(args.out_dir, "premiums_long.csv"), long_rows)
    write_csv(os.path.join(args.out_dir, "quote.csv"), quote)
    write_csv(os.path.join(args.out_dir, "unparsed_levels.csv"), unparsed)

    # ---- diagnostics -----------------------------------------------------
    print("=" * 66)
    print("BUILD REPORT")
    print("=" * 66)
    print(f"plan variants (catalog rows) : {len(catalog)}")
    print(f"premium rows (long/tidy)     : {len(long_rows):,}")
    print(f"quote rows                   : {len(quote)}")
    print(f"insurers                     : {len({r['insurer'] for r in catalog})}")
    print(f"distinct plan names          : {len({r['plan_name'] for r in catalog})}")
    print()
    print("-- join integrity --")
    cat_certs = {r["certification_no"] for r in catalog}
    print(f"catalog certs                : {len(cat_certs)}")
    print(f"premium certs                : {len(prem_by_cert)}")
    print(f"catalog without premium      : {len(cat_certs - set(prem_by_cert))}")
    print(f"premium without catalog      : {len(set(prem_by_cert) - cat_certs)}")
    ins = {r["insurer"] for r in catalog}
    provnames = {en(p["company-name"]) for p in providers}
    print(f"insurers not in provider list: {len(ins - provnames)} {sorted(ins - provnames)[:3]}")
    print()
    print("-- field coverage (what the JSONs give you natively) --")

    def cov(field):
        n = sum(1 for r in catalog if str(r[field]).strip() != "")
        return f"{n:>3}/{len(catalog)}  {n / len(catalog) * 100:5.1f}%"

    for f in ["ward_type", "deductible_amount", "geographical_coverage",
              "currency", "age_counting_method", "plan_doc_url"]:
        print(f"  {f:<24} {cov(f)}")
    print()
    print("-- parse confidence on plan-level free text --")
    for k, v in Counter(r["level_parse_confidence"] for r in catalog).most_common():
        print(f"  {k:<10} {v:>3}  ({v / len(catalog) * 100:.1f}%)")
    print()
    print("-- by plan type --")
    for k, v in Counter(r["plan_type"] for r in catalog).most_common():
        print(f"  {k:<10} {v}")
    print()
    print(f"-- rows needing PDF scrape for deductible/geography: {len(unparsed)} --")
    print()
    print(f"-- sample quote: age {args.age} {args.gender} smoker={args.smoker} "
          f"basis={args.basis} --")
    hkd = sorted([r for r in quote
                  if r["currency"] == "HKD" and r["first_year_premium"] != ""],
                 key=lambda r: r["first_year_premium"])
    print(f"{'insurer':<34}{'level':<30}{'1st yr':>10}{'10y avg':>10}")
    for r in hkd[:8]:
        print(f"{r['insurer_alias'] or r['insurer'][:32]:<34}"
              f"{r['plan_level_raw'][:28]:<30}"
              f"{r['first_year_premium']:>10}{r['ten_year_avg']:>10}")
    print("  ...")
    for r in hkd[-3:]:
        print(f"{r['insurer_alias'] or r['insurer'][:32]:<34}"
              f"{r['plan_level_raw'][:28]:<30}"
              f"{r['first_year_premium']:>10}{r['ten_year_avg']:>10}")


if __name__ == "__main__":
    main()
