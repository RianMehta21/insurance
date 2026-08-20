#!/usr/bin/env python3
"""
VHIS plan document scraper.

Extracts the fields the JSONs do not carry: room type, geographical coverage,
deductible, coinsurance, annual benefit limit, lifetime limit.

WHY THIS WORKS DESPITE 33 DIFFERENT INSURERS
--------------------------------------------
Every Certified Plan document must follow the government's Certified Plan
Policy Template. The section numbering is mandated, and the variable parts are
written as explicit Option A / Option B choices. So we anchor on fixed headings
rather than trying to guess at layout:

  Part 6 Sec 1(a) Territorial scope of cover   -> geographical coverage
  Part 6 Sec 1(b) Lifetime Benefit Limit       -> lifetime limit
  Part 6 Sec 1(c) Choice of providers          -> network restriction
  Part 6 Sec 1(d) Choice of ward class         -> room type restriction
  Part 6 Sec 5    Cost-sharing requirement     -> deductible / coinsurance
  Benefit Schedule table                        -> annual limit, per-item limits

CONFIDENCE TIERS
----------------
  tier1  Binary Option A/B detection. Template-mandated wording, high accuracy.
         Trust these without review.
  tier2  Numbers pulled from the Benefit Schedule table. Layout varies between
         insurers, especially Flexi plans with enhanced benefits. Spot check.
  tier3  Nothing matched. Goes to the review queue for manual entry.

Every row carries its confidence tier and the raw snippet the value came from,
so nothing is a black box when a number looks wrong.

Usage:
    python scrape_plans.py --catalog out/plans_catalog.csv --cache pdf_cache
    python scrape_plans.py --catalog out/plans_catalog.csv --only F00022,F00041
    python scrape_plans.py --selftest
"""

import argparse
import csv
import os
import re
import sys
import time
import urllib.request

UA = {"User-Agent": "vhis-compare/1.0 (personal plan comparison tool)"}


# ---------------------------------------------------------------------------
# text extraction
# ---------------------------------------------------------------------------
def pdf_to_text(path):
    """pdftotext -layout preserves column structure, which the benefit
    schedule table needs. Falls back to pdfplumber."""
    try:
        import subprocess
        r = subprocess.run(["pdftotext", "-layout", path, "-"],
                           capture_output=True, timeout=120)
        if r.returncode == 0 and len(r.stdout) > 200:
            return r.stdout.decode("utf-8", errors="replace")
    except Exception:
        pass
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception as e:
        return f"__EXTRACTION_FAILED__ {e}"


def download(url, cache_dir, cert):
    os.makedirs(cache_dir, exist_ok=True)
    dest = os.path.join(cache_dir, f"{cert}.pdf")
    if os.path.exists(dest) and os.path.getsize(dest) > 1000:
        return dest, "cached"
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=90) as r:
        blob = r.read()
    if not blob.startswith(b"%PDF"):
        raise ValueError("not a PDF")
    with open(dest, "wb") as f:
        f.write(blob)
    return dest, "downloaded"


def squash(t):
    """PDFs break sentences across lines. Collapse whitespace so a regex can
    span what was visually two lines."""
    return re.sub(r"\s+", " ", t)


def window(text, pattern, before=0, after=1200):
    m = re.search(pattern, text, re.I)
    if not m:
        return None
    return text[max(0, m.start() - before): m.end() + after]


# ---------------------------------------------------------------------------
# tier 1: Option A / Option B binaries
# ---------------------------------------------------------------------------
def extract_geography(text):
    sec = window(text, r"Territorial\s+scope\s+of\s+cover", after=1400)
    if not sec:
        return None, None, None
    s = squash(sec)
    if re.search(r"benefits?\s+described[^.]{0,120}applicable\s+worldwide", s, re.I):
        return "Worldwide", "tier1", s[:220]
    if re.search(r"subject\s+to\s+the\s+geographical\s+limitation", s, re.I):
        for pat, label in [
            (r"worldwide\s*(?:excluding|except)[^.]{0,60}(?:USA|United States)",
             "Worldwide excl. USA"),
            (r"Greater\s+China", "Greater China"),
            (r"Asia[\s\-]*Pacific", "Asia Pacific"),
            (r"\bAsia\b", "Asia"),
            (r"Hong\s*Kong\s*(?:and|,)?\s*Macau", "Hong Kong/Macau"),
        ]:
            if re.search(pat, s, re.I):
                return label, "tier2", s[:220]
        return "Limited (see doc)", "tier2", s[:220]
    return None, "tier3", s[:220]


def extract_lifetime_limit(text):
    sec = window(text, r"Lifetime\s+Benefit\s+Limit", after=900)
    if not sec:
        return None, None, None
    s = squash(sec)
    if re.search(r"not\s+subject\s+to\s+any\s+Lifetime\s+Benefit\s+Limit", s, re.I):
        return "None", "tier1", s[:200]
    m = re.search(r"Lifetime\s+Benefit\s+Limit[^.]{0,80}?"
                  r"(?:HKD?|USD?|\$)\s*([\d,]{4,})", s, re.I)
    if m:
        return m.group(1).replace(",", ""), "tier2", s[:200]
    if re.search(r"subject\s+to\s+the\s+Lifetime\s+Benefit\s+Limit", s, re.I):
        return "Capped (see doc)", "tier2", s[:200]
    return None, "tier3", s[:200]


def extract_ward_restriction(text):
    sec = window(text, r"Choice\s+of\s+ward\s+class", after=1000)
    if not sec:
        return None, None, None
    s = squash(sec)
    if re.search(r"not\s+subject\s+to\s+any\s+restriction\s+in\s+the\s+choice\s+of\s+ward",
                 s, re.I):
        return "No restriction", "tier1", s[:200]
    if re.search(r"subject\s+to\s+the\s+restriction\s+in\s+the\s+choice\s+of\s+ward",
                 s, re.I):
        for pat, label in [(r"\bsuite\b", "Suite"),
                           (r"standard\s*private", "Standard Private"),
                           (r"semi[\s\-]?private", "Semi-private"),
                           (r"\bprivate\b", "Private"),
                           (r"\bward\b", "Ward")]:
            if re.search(pat, s, re.I):
                return label, "tier2", s[:200]
        return "Restricted (see doc)", "tier2", s[:200]
    return None, "tier3", s[:200]


def extract_provider_restriction(text):
    sec = window(text, r"Choice\s+of\s+healthcare\s+services?\s+providers?", after=900)
    if not sec:
        return None, None, None
    s = squash(sec)
    if re.search(r"not\s+subject\s+to\s+any\s+restriction", s, re.I):
        return "No restriction", "tier1", s[:200]
    if re.search(r"subject\s+to\s+the\s+restriction", s, re.I):
        return "Network restricted", "tier1", s[:200]
    return None, "tier3", s[:200]


def extract_cost_sharing(text):
    """Deductible and coinsurance. Part 6 Section 5."""
    sec = window(text, r"Cost[\s\-]?sharing\s+requirement", after=1400)
    if not sec:
        return {}, "tier3", ""
    s = squash(sec)
    out = {}

    has_ded = bool(re.search(r"required\s+to\s+pay\s+Coinsurance\s+and\s*/?\s*or\s+Deductible",
                             s, re.I))
    only_coins = bool(re.search(
        r"required\s+to\s+pay\s+for\s+Coinsurance\s+for\s+Prescribed\s+Diagnostic", s, re.I))

    if only_coins and not has_ded:
        out["has_deductible"] = "N"
        out["deductible_amount"] = 0
        tier = "tier1"
    elif has_ded:
        out["has_deductible"] = "Y"
        tier = "tier1"
    else:
        tier = "tier3"

    m = re.search(r"Deductible[^.]{0,120}?(?:HKD?|USD?|\$)\s*([\d,]{3,})", s, re.I)
    if m:
        out["deductible_amount"] = int(m.group(1).replace(",", ""))
        tier = "tier2"

    m = re.search(r"(\d{1,3})\s*%\s*Coinsurance", s, re.I)
    if m:
        out["coinsurance_pct"] = int(m.group(1))
    return out, tier, s[:240]


# ---------------------------------------------------------------------------
# tier 2: benefit schedule table
# ---------------------------------------------------------------------------
# The lookbehind matters: without it a greedy gap ahead of this pattern can
# chew into the number itself and capture only its trailing digits, so
# "$14,000" silently becomes 0. Gaps below are all non-greedy for the same
# reason.
MONEY = r"(?:HKD?|USD?|\$)?\s*(?<![\d,])(\d[\d,]*\d)"


def extract_benefit_schedule(text):
    sec = window(text, r"Benefit\s+Schedule", before=200, after=9000)
    if not sec:
        return {}, "tier3", ""
    out = {}
    s = squash(sec)

    m = re.search(r"Annual\s+Benefit\s+Limit[^%]{0,200}?" + MONEY, s, re.I)
    if m:
        out["annual_benefit_limit"] = int(m.group(1).replace(",", ""))

    m = re.search(r"Room\s+and\s+board[^A-Za-z]{0,40}?" + MONEY + r"\s*per\s*day", s, re.I)
    if m:
        out["room_board_per_day"] = int(m.group(1).replace(",", ""))

    m = re.search(r"Miscellaneous\s+charges[^A-Za-z]{0,40}?" + MONEY, s, re.I)
    if m:
        out["misc_charges_limit"] = int(m.group(1).replace(",", ""))

    m = re.search(r"Surgeon'?s?\s+fee.{0,400}?Complex[^\d]{0,20}?" + MONEY, s, re.I)
    if m:
        out["surgeon_complex_limit"] = int(m.group(1).replace(",", ""))

    m = re.search(r"Psychiatric\s+treatments?[^A-Za-z]{0,40}?" + MONEY, s, re.I)
    if m:
        out["psychiatric_limit"] = int(m.group(1).replace(",", ""))

    if re.search(r"no\s+sub[\s\-]?limit", s, re.I):
        out["no_sublimits"] = "Y"

    return out, ("tier2" if out else "tier3"), s[:240]


# ---------------------------------------------------------------------------
# premium schedule PDFs (validated against a real Manulife schedule)
# ---------------------------------------------------------------------------
def extract_premium_meta(text):
    """The premium PDFs carry things the JSON drops entirely: instalment
    loading factors, the levy disclaimer, and the renewal-only age band."""
    s = squash(text)
    out = {}

    for pat, key in [(r"Age\s+Nearest\s+Birthday", "R"),
                     (r"Age\s+Next\s+Birthday", "N"),
                     (r"Age\s+(?:Last|Attained)\s+Birthday", "A")]:
        if re.search(pat, s, re.I):
            out["age_basis_stated"] = key
            break

    m = re.search(r"Semi[\s\-]?annual:\s*([\d.]+)", s, re.I)
    if m:
        out["factor_semiannual"] = float(m.group(1))
    m = re.search(r"Quarterly:\s*([\d.]+)", s, re.I)
    if m:
        out["factor_quarterly"] = float(m.group(1))
    m = re.search(r"Monthly:\s*([\d.]+)", s, re.I)
    if m:
        out["factor_monthly"] = float(m.group(1))
        out["monthly_annualised_loading_pct"] = round(
            (out["factor_monthly"] * 12 - 1) * 100, 1)

    m = re.search(r"from\s+age\s+(\d+)\s+to\s+(\d+)\s+years\s+at\s+Policy\s+commencement",
                  s, re.I)
    if m:
        out["new_business_age_min"] = int(m.group(1))
        out["new_business_age_max"] = int(m.group(2))

    if re.search(r"For\s+renewal\s+only", s, re.I):
        out["has_renewal_only_ages"] = "Y"
    if re.search(r"does\s+not\s+include\s+levy", s, re.I):
        out["levy_excluded"] = "Y"

    m = re.search(r"Effective\s+Date\s*\([^)]*\)\s*:?\s*(\d{2}/\d{2}/\d{4})", s, re.I)
    if m:
        out["premium_effective_date"] = m.group(1)
    return out


# ---------------------------------------------------------------------------
# orchestration
# ---------------------------------------------------------------------------
FIELDS = ["certification_no", "insurer", "plan_name", "plan_level_raw",
          "geographical_coverage", "geo_tier",
          "ward_restriction", "ward_tier",
          "provider_restriction",
          "has_deductible", "deductible_amount", "coinsurance_pct", "cost_tier",
          "lifetime_benefit_limit", "lifetime_tier",
          "annual_benefit_limit", "room_board_per_day", "misc_charges_limit",
          "surgeon_complex_limit", "psychiatric_limit", "no_sublimits",
          "schedule_tier", "overall_tier", "scrape_status", "source_url"]


def scrape_one(row, cache_dir, sleep=1.0):
    cert = row["certification_no"]
    url = row["plan_doc_url"]
    rec = {f: "" for f in FIELDS}
    rec.update({
        "certification_no": cert,
        "insurer": row.get("insurer", ""),
        "plan_name": row.get("plan_name", ""),
        "plan_level_raw": row.get("plan_level_raw", ""),
        "source_url": url,
    })
    if not url:
        rec["scrape_status"] = "no_url"
        rec["overall_tier"] = "tier3"
        return rec
    try:
        path, how = download(url, cache_dir, cert)
        if how == "downloaded":
            time.sleep(sleep)
    except Exception as e:
        rec["scrape_status"] = f"download_failed: {e}"
        rec["overall_tier"] = "tier3"
        return rec

    text = pdf_to_text(path)
    if text.startswith("__EXTRACTION_FAILED__"):
        rec["scrape_status"] = "no_text_layer_needs_ocr"
        rec["overall_tier"] = "tier3"
        return rec

    geo, gt, _ = extract_geography(text)
    ward, wt, _ = extract_ward_restriction(text)
    prov, _, _ = extract_provider_restriction(text)
    cost, ct, _ = extract_cost_sharing(text)
    life, lt, _ = extract_lifetime_limit(text)
    sched, st, _ = extract_benefit_schedule(text)

    rec.update({
        "geographical_coverage": geo or "", "geo_tier": gt or "tier3",
        "ward_restriction": ward or "", "ward_tier": wt or "tier3",
        "provider_restriction": prov or "",
        "has_deductible": cost.get("has_deductible", ""),
        "deductible_amount": cost.get("deductible_amount", ""),
        "coinsurance_pct": cost.get("coinsurance_pct", ""),
        "cost_tier": ct,
        "lifetime_benefit_limit": life or "", "lifetime_tier": lt or "tier3",
        "annual_benefit_limit": sched.get("annual_benefit_limit", ""),
        "room_board_per_day": sched.get("room_board_per_day", ""),
        "misc_charges_limit": sched.get("misc_charges_limit", ""),
        "surgeon_complex_limit": sched.get("surgeon_complex_limit", ""),
        "psychiatric_limit": sched.get("psychiatric_limit", ""),
        "no_sublimits": sched.get("no_sublimits", ""),
        "schedule_tier": st,
        "scrape_status": "ok",
    })
    tiers = [gt, wt, ct, lt, st]
    rec["overall_tier"] = ("tier3" if all(t == "tier3" for t in tiers)
                           else "tier2" if "tier2" in tiers or "tier3" in tiers
                           else "tier1")
    return rec


def selftest():
    """Validates the regexes against the exact wording the government template
    mandates, plus a real premium schedule. No network needed."""
    tpl_worldwide = """
    (a) Territorial scope of cover
    [Option A - applicable to Standard Plan, and those Flexi Plans without
    geographical limitation for benefit coverage
    Except for the psychiatric treatment as stated in Section 3(l) of this
    Part 6, all benefits described in these Terms and Benefits shall be
    applicable worldwide.]
    (b) Lifetime Benefit Limit
    [Option A - All benefits described in these Terms and Benefits are not
    subject to any Lifetime Benefit Limit.]
    (d) Choice of ward class
    [Option A - All benefits described in these Terms and Benefits are not
    subject to any restriction in the choice of ward class in Hospital.]
    5. Cost-sharing requirement
    [Option A - The Policy Holder is required to pay for Coinsurance for
    Prescribed Diagnostic Imaging Tests as specified in this Part 6.]
    Standard Plan Benefit Schedule
    (a) Room and board $750 per day Maximum 180 days per Policy Year
    (b) Miscellaneous charges $14,000 per Policy Year
    (f) Surgeon's fee Per surgery, subject to surgical category
    Complex $50,000 Major $25,000 Intermediate $12,500 Minor $5,000
    (l) Psychiatric treatments $30,000 per Policy Year
    Annual Benefit Limit for benefit items (a) - (l) $420,000 per Policy Year
    """
    tpl_limited = """
    (a) Territorial scope of cover
    [Option B - Except for the psychiatric treatment, all benefits described in
    these Terms and Benefits are subject to the geographical limitation for
    benefit coverage as stated in Section 12, namely Asia excluding Japan.]
    (d) Choice of ward class
    [Option B - The benefits described in these Terms and Benefits are subject
    to the restriction in the choice of ward class, namely Semi-private room.]
    5. Cost-sharing requirement
    [Option B - The Policy Holder is required to pay Coinsurance and/or
    Deductible as stated in these Terms and Benefits. Deductible: HKD 25,000
    per Policy Year, subject to 20% Coinsurance.]
    """
    prem = """
    Standard Premium Schedule - HKD  Age Nearest Birthday  Annual
    (For Insured Persons from age 0 to 81 years at Policy commencement)
    *For renewal only.
    This Standard Premium schedule does not include levy which is collected by
    the Insurance Authority.
    The above premiums are for annual payment mode. The following adjustment
    factor will be multiplied: Semi-annual: 0.52, Quarterly: 0.265, Monthly: 0.09
    Effective Date (DD/MM/YYYY): 29/12/2025
    """

    checks, fails = [], 0

    def ck(name, got, want):
        nonlocal fails
        ok = got == want
        if not ok:
            fails += 1
        checks.append((("PASS" if ok else "FAIL"), name, got, want))

    g, gt, _ = extract_geography(tpl_worldwide)
    ck("worldwide geography", g, "Worldwide")
    ck("worldwide geo tier", gt, "tier1")
    l, lt, _ = extract_lifetime_limit(tpl_worldwide)
    ck("no lifetime limit", l, "None")
    w, wt, _ = extract_ward_restriction(tpl_worldwide)
    ck("no ward restriction", w, "No restriction")
    c, ct, _ = extract_cost_sharing(tpl_worldwide)
    ck("no deductible", c.get("has_deductible"), "N")
    s, st, _ = extract_benefit_schedule(tpl_worldwide)
    ck("annual benefit limit", s.get("annual_benefit_limit"), 420000)
    ck("room and board", s.get("room_board_per_day"), 750)
    ck("misc charges", s.get("misc_charges_limit"), 14000)
    ck("surgeon complex", s.get("surgeon_complex_limit"), 50000)
    ck("psychiatric", s.get("psychiatric_limit"), 30000)

    g2, _, _ = extract_geography(tpl_limited)
    ck("limited geography", g2, "Asia")
    w2, _, _ = extract_ward_restriction(tpl_limited)
    ck("ward restricted", w2, "Semi-private")
    c2, _, _ = extract_cost_sharing(tpl_limited)
    ck("has deductible", c2.get("has_deductible"), "Y")
    ck("deductible amount", c2.get("deductible_amount"), 25000)
    ck("coinsurance", c2.get("coinsurance_pct"), 20)

    p = extract_premium_meta(prem)
    ck("age basis", p.get("age_basis_stated"), "R")
    ck("monthly factor", p.get("factor_monthly"), 0.09)
    ck("monthly loading %", p.get("monthly_annualised_loading_pct"), 8.0)
    ck("new business max age", p.get("new_business_age_max"), 81)
    ck("renewal-only ages", p.get("has_renewal_only_ages"), "Y")
    ck("levy excluded", p.get("levy_excluded"), "Y")

    print("=" * 72)
    print("SELF TEST (template wording, no network required)")
    print("=" * 72)
    for status, name, got, want in checks:
        mark = "ok " if status == "PASS" else "XX "
        print(f"  {mark} {name:<26} got={str(got):<18} want={want}")
    print(f"\n{len(checks)-fails}/{len(checks)} passed")
    return 1 if fails else 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--catalog", default="out/plans_catalog.csv")
    ap.add_argument("--cache", default="pdf_cache")
    ap.add_argument("--out", default="out/plan_benefits.csv")
    ap.add_argument("--review", default="out/review_queue.csv")
    ap.add_argument("--only", help="comma-separated cert-no prefixes")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--sleep", type=float, default=1.0)
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()

    if args.selftest:
        sys.exit(selftest())

    rows = list(csv.DictReader(open(args.catalog, encoding="utf-8-sig")))
    if args.only:
        pref = tuple(p.strip() for p in args.only.split(","))
        rows = [r for r in rows if r["certification_no"].startswith(pref)]
    if args.limit:
        rows = rows[:args.limit]

    results = []
    for i, row in enumerate(rows, 1):
        rec = scrape_one(row, args.cache, args.sleep)
        results.append(rec)
        print(f"[{i}/{len(rows)}] {rec['certification_no']:<22} "
              f"{rec['overall_tier']:<6} {rec['scrape_status']}")

    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(results)

    review = [r for r in results if r["overall_tier"] == "tier3"
              or r["scrape_status"] != "ok"]
    with open(args.review, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(review)

    from collections import Counter
    print("\nconfidence:", dict(Counter(r["overall_tier"] for r in results)))
    print(f"review queue: {len(review)} rows -> {args.review}")


if __name__ == "__main__":
    main()
