#!/usr/bin/env python3
"""
Builds the VHIS comparison workbook.

The workbook is the application, not a report. Change the age on the Client
sheet and all 579 plans reprice instantly via INDEX/MATCH. No Python, no
macros, no add-ins on the user's machine.

Sheets
  Client    inputs (age, gender, smoker, basis) + legend + disclaimers
  Compare   579 plans, AutoFilter on, live premium columns
  Benefits  scraped/override benefit fields, user-editable
  prem      hidden wide premium matrix, one row per cert|basis|gender|smoker
  Notes     provenance, refresh date, assumptions

Usage:
    python make_workbook.py --data-dir current --out VHIS_Compare.xlsx
"""

import argparse
import json
import os
import re
from collections import OrderedDict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# reuse the parsing/loading logic already written and tested
import build_vhis as bv

MAX_AGE = 100
N_AGE_COLS = MAX_AGE + 1                      # ages 0..100
PREM_FIRST_COL = 2                            # column B holds age 0
PREM_LAST_COL = PREM_FIRST_COL + MAX_AGE      # column CX

ARIAL = "Arial"
BLUE = Font(name=ARIAL, size=10, color="0000FF", bold=True)
BLACK = Font(name=ARIAL, size=10)
HDR = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
TITLE = Font(name=ARIAL, size=14, bold=True)
SUB = Font(name=ARIAL, size=10, italic=True, color="595959")
HDR_FILL = PatternFill("solid", fgColor="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
EDIT_FILL = PatternFill("solid", fgColor="FFF2CC")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def money(fmt_zero_dash=True):
    return '#,##0;(#,##0);-' if fmt_zero_dash else '#,##0'


# ---------------------------------------------------------------------------
def build_premium_rows(prem_by_cert):
    """One row per cert|basis|gender|smoker. Blank where the plan has no
    premium at that age, so Excel's COUNT can distinguish 'no cover' from 0."""
    rows = []
    for cert, rec in prem_by_cert.items():
        pm = rec.get("premium") or {}
        for b in ("S", "R"):
            for g in ("M", "F"):
                for s in ("N", "Y"):
                    series = pm.get(f"{b}{g}{s}")
                    if not series:
                        continue
                    vals = [series.get(str(a)) for a in range(N_AGE_COLS)]
                    if all(v is None for v in vals):
                        continue
                    rows.append([f"{cert}|{b}|{g}|{s}"] + vals)
    rows.sort(key=lambda r: r[0])
    return rows


def sheet_client(wb, catalog_n, prem_date_max):
    ws = wb.create_sheet("Client", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 82

    ws["B2"] = "VHIS Plan Comparison"
    ws["B2"].font = TITLE
    ws["B3"] = (f"{catalog_n} certified plan variants across 33 insurers. "
                "Change the yellow cells and every plan reprices.")
    ws["B3"].font = SUB

    ws["B5"] = "CLIENT PROFILE"
    ws["B5"].font = Font(name=ARIAL, size=11, bold=True)

    spec = [
        ("Age (years)", 35, "B6",
         "Client's actual age. Each insurer's age basis is applied automatically."),
        ("Gender", "M", "B7", "M or F. 220 of the variants are gender-rated; the rest are unisex."),
        ("Smoker", "N", "B8", "N or Y. Only 46 variants price smokers differently."),
        ("Premium basis", "S", "B9",
         "S = standalone policy. R = rider attached to another policy. Falls back automatically."),
    ]
    r = 6
    for label, default, _, note in spec:
        ws.cell(r, 2, label).font = BLACK
        c = ws.cell(r, 3, default)
        c.font = BLUE
        c.fill = INPUT_FILL
        c.border = BOX
        c.alignment = Alignment(horizontal="center")
        ws.cell(r, 4, note).font = SUB
        r += 1

    ws["C6"].number_format = "0"
    for rng, formula in [("C7", '"M,F"'), ("C8", '"N,Y"'), ("C9", '"S,R"')]:
        dv = DataValidation(type="list", formula1=formula, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ws[rng])
    dv_age = DataValidation(type="whole", operator="between",
                            formula1=0, formula2=100, allow_blank=False)
    ws.add_data_validation(dv_age)
    dv_age.add(ws["C6"])

    ws["B11"] = "HOW TO USE"
    ws["B11"].font = Font(name=ARIAL, size=11, bold=True)
    steps = [
        "1. Set the four yellow cells above.",
        "2. Go to the Compare sheet. Every premium column has already recalculated.",
        "3. Use the filter arrows on row 4 to narrow by ward, deductible, geography, insurer.",
        "4. Sort by 10-Year Avg, not First Year. See the warning below.",
        "5. Copy the shortlisted rows into a fresh sheet for the client.",
    ]
    for i, s in enumerate(steps):
        ws.cell(12 + i, 2, s).font = BLACK
        ws.merge_cells(start_row=12 + i, start_column=2, end_row=12 + i, end_column=4)

    ws["B18"] = "COLOUR LEGEND"
    ws["B18"].font = Font(name=ARIAL, size=11, bold=True)
    legend = [("Blue on yellow", INPUT_FILL, "You type here. These are the only cells to change."),
              ("Orange fill", EDIT_FILL, "Benefits sheet. Paste scraper output or type corrections."),
              ("Black text", None, "Formula. Do not overwrite or the sheet stops recalculating.")]
    for i, (lbl, fill, desc) in enumerate(legend):
        c = ws.cell(19 + i, 2, lbl)
        c.font = BLUE if fill is INPUT_FILL else BLACK
        if fill:
            c.fill = fill
        c.border = BOX
        ws.cell(19 + i, 4, desc).font = SUB

    ws["B23"] = "READ BEFORE QUOTING"
    ws["B23"].font = Font(name=ARIAL, size=11, bold=True, color="C00000")
    warns = [
        "Age 72 and above: the source data stops at each insurer's new-application age ceiling, "
        "so the 10-year window is incomplete for most plans. Check the Years Avail column.",
        "Premiums exclude the Insurance Authority levy, underwriting loading, and any discounts. "
        "They are portfolio-basis standard premiums, not quotes.",
        "Paying monthly typically costs about 8% more per year than annual. Not reflected here.",
        "Age-nearest-birthday insurers are approximated from whole-year age. "
        "Those rows show APPROX and need date of birth to be exact.",
        "Never rank across currencies. Filter to HKD or USD first.",
    ]
    for i, w in enumerate(warns):
        c = ws.cell(24 + i, 2, w)
        c.font = Font(name=ARIAL, size=9)
        c.fill = WARN_FILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=24 + i, start_column=2, end_row=24 + i, end_column=4)
        ws.row_dimensions[24 + i].height = 26

    ws.cell(31, 2, "Data refreshed").font = BLACK
    ws.cell(31, 3, datetime.now().strftime("%Y-%m-%d")).font = BLACK
    ws.cell(32, 2, "Latest premium date in data").font = BLACK
    ws.cell(32, 3, prem_date_max).font = BLACK
    ws.cell(33, 2, "Source").font = BLACK
    ws.cell(33, 3, "data.gov.hk / Health Bureau VHIS open data").font = SUB
    return ws


COMPARE_COLS = OrderedDict([
    ("Insurer", 26), ("Plan Name", 34), ("Type", 8), ("Plan Level", 30),
    ("Ward", 14), ("Deductible", 12), ("Geography", 18), ("Ccy", 6),
    ("First Year", 12), ("10-Year Total", 14), ("10-Year Avg", 13),
    ("Years Avail", 11), ("Age Basis", 11), ("Basis Used", 11),
    ("Smoker Rated", 12), ("Annual Limit", 13), ("Coinsurance %", 12),
    ("Cert No", 22), ("_tableage", 10), ("_key", 30), ("_row", 8),
])


def sheet_compare(wb, catalog, n_prem_rows):
    ws = wb.create_sheet("Compare")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    ws["A1"] = "Compare"
    ws["A1"].font = TITLE
    ws["A2"] = ("Live. Driven by the yellow cells on the Client sheet. "
                "Filter with the arrows on row 4. Sort by 10-Year Avg.")
    ws["A2"].font = SUB

    for i, (name, width) in enumerate(COMPARE_COLS.items(), start=1):
        c = ws.cell(4, i, name)
        c.font = HDR
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width

    prem_last_row = n_prem_rows + 1
    keys = f"prem!$A$2:$A${prem_last_row}"
    grid = (f"prem!${get_column_letter(PREM_FIRST_COL)}$2:"
            f"${get_column_letter(PREM_LAST_COL)}${prem_last_row}")

    age = "Client!$C$6"
    gender = "Client!$C$7"
    smoker = "Client!$C$8"
    basis = "Client!$C$9"

    ben_last = len(catalog) + 1
    ben_keys = f"Benefits!$A$2:$A${ben_last}"

    r = 5
    for row in catalog:
        cert = row["certification_no"]
        ws.cell(r, 1, row["insurer"]).font = BLACK
        ws.cell(r, 2, row["plan_name"]).font = BLACK
        ws.cell(r, 3, row["plan_type"]).font = BLACK
        ws.cell(r, 4, row["plan_level_raw"]).font = BLACK

        # Ward / deductible / geography: prefer the Benefits sheet (scraped or
        # hand-entered), fall back to what we parsed from the plan-level label.
        for col, ben_col, fallback in (
                (5, "B", row["ward_type"]),
                (6, "C", row["deductible_amount"]),
                (7, "D", row["geographical_coverage"])):
            fb = f'"{fallback}"' if fallback != "" else '""'
            f = (f'=IFERROR(IF(INDEX(Benefits!${ben_col}$2:${ben_col}${ben_last},'
                 f'MATCH($R{r},{ben_keys},0))="",{fb},'
                 f'INDEX(Benefits!${ben_col}$2:${ben_col}${ben_last},'
                 f'MATCH($R{r},{ben_keys},0))),{fb})')
            c = ws.cell(r, col, f)
            c.font = BLACK
        ws.cell(r, 6).number_format = money()

        ws.cell(r, 8, row["currency"]).font = BLACK
        ws.cell(r, 13, row["age_counting_method"]).font = BLACK
        ws.cell(r, 15, row["smoker_rated"]).font = BLACK
        ws.cell(r, 18, cert).font = Font(name=ARIAL, size=9, color="808080")

        # helper: table age per insurer's own age-counting convention
        ws.cell(r, 19, f'=IF($M{r}="N",{age}+1,{age})').font = BLACK

        # helper: basis actually used, with automatic fallback
        ws.cell(r, 14,
                f'=IF(COUNTIF({keys},$R{r}&"|"&{basis}&"|"&{gender}&"|"&{smoker})>0,'
                f'{basis},IF({basis}="S","R","S"))').font = BLACK

        ws.cell(r, 20,
                f'=$R{r}&"|"&$N{r}&"|"&{gender}&"|"&{smoker}').font = BLACK
        ws.cell(r, 21, f'=IFERROR(MATCH($T{r},{keys},0),"")').font = BLACK

        col_first = f"$S{r}+1"
        col_last = f"MIN($S{r}+10,{N_AGE_COLS})"

        # First year. INDEX returns 0 on a blank cell, so gate on COUNT.
        ws.cell(r, 9,
                f'=IF($U{r}="","",IF(COUNT(INDEX({grid},$U{r},{col_first}))=0,"",'
                f'INDEX({grid},$U{r},{col_first})))').number_format = money()
        ws.cell(r, 10,
                f'=IF($U{r}="","",IF($L{r}=0,"",'
                f'SUM(INDEX({grid},$U{r},{col_first}):'
                f'INDEX({grid},$U{r},{col_last}))))').number_format = money()
        ws.cell(r, 11,
                f'=IF($J{r}="","",IF($L{r}=0,"",$J{r}/$L{r}))').number_format = money()
        ws.cell(r, 12,
                f'=IF($U{r}="",0,COUNT(INDEX({grid},$U{r},{col_first}):'
                f'INDEX({grid},$U{r},{col_last})))')

        for col, ben_col in ((16, "E"), (17, "F")):
            idx = (f'INDEX(Benefits!${ben_col}$2:${ben_col}${ben_last},'
                   f'MATCH($R{r},{ben_keys},0))')
            ws.cell(r, col, f'=IFERROR(IF(COUNT({idx})=0,"",{idx}),"")')
        ws.cell(r, 16).number_format = money()

        for col in range(9, 13):
            ws.cell(r, col).font = BLACK
        r += 1

    last = r - 1
    ws.auto_filter.ref = f"A4:U{last}"

    # Hide the helper columns. They must stay inside the filter range for the
    # formulas to resolve, but nobody needs to look at them.
    for letter in ("S", "T", "U"):
        ws.column_dimensions[letter].hidden = True
    return ws, last


def load_benefits(benefits_csv, overrides_csv):
    """Merge scraper output with hand-entered overrides. Overrides win.

    These live in CSV files, NOT in the workbook, because the workbook is
    regenerated on every refresh. Anything typed into the spreadsheet itself
    would be silently destroyed the next time the pipeline runs.
    """
    import csv as _csv
    merged = {}
    for path, tag in ((benefits_csv, "scraped"), (overrides_csv, "override")):
        if not path or not os.path.exists(path):
            continue
        with open(path, encoding="utf-8-sig") as f:
            for row in _csv.DictReader(f):
                cert = (row.get("certification_no") or "").strip()
                if not cert:
                    continue
                cur = merged.setdefault(cert, {})
                # Accept the scraper's long column names AND plain short ones,
                # so a hand-written overrides file does not silently no-op.
                for src, dst in (("ward_restriction", "ward"),
                                 ("ward_type", "ward"),
                                 ("ward", "ward"),
                                 ("deductible_amount", "deductible"),
                                 ("deductible", "deductible"),
                                 ("geographical_coverage", "geography"),
                                 ("geography", "geography"),
                                 ("annual_benefit_limit", "annual_limit"),
                                 ("annual_limit", "annual_limit"),
                                 ("coinsurance_pct", "coinsurance"),
                                 ("coinsurance", "coinsurance"),
                                 ("lifetime_benefit_limit", "lifetime"),
                                 ("lifetime", "lifetime")):
                    v = (row.get(src) or "").strip()
                    if v:
                        cur[dst] = v
                if any(cur.get(k) for k in ("ward", "deductible", "geography")):
                    cur["tier"] = tag if tag == "override" else row.get("overall_tier", "scraped")
    return merged


def sheet_benefits(wb, catalog, benefits=None):
    ws = wb.create_sheet("Benefits")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    ws["A1"] = ("Benefit fields from scrape_plans.py plus manual overrides. "
                "DO NOT EDIT HERE: this sheet is rebuilt from data/plan_benefits.csv "
                "and data/manual_overrides.csv every refresh, so edits made in Excel "
                "are lost. Put corrections in manual_overrides.csv instead.")
    ws["A1"].font = SUB
    ws.merge_cells("A1:H1")

    cols = [("Cert No", 22), ("Ward", 16), ("Deductible", 13), ("Geography", 20),
            ("Annual Limit", 14), ("Coinsurance %", 13), ("Lifetime Limit", 15),
            ("Source / Tier", 16)]
    for i, (name, w) in enumerate(cols, start=1):
        c = ws.cell(2, i, name)
        c.font = HDR
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = w

    benefits = benefits or {}
    for i, row in enumerate(catalog, start=3):
        cert = row["certification_no"]
        ws.cell(i, 1, cert).font = Font(name=ARIAL, size=9)
        b = benefits.get(cert, {})
        vals = [b.get("ward", ""), b.get("deductible", ""), b.get("geography", ""),
                b.get("annual_limit", ""), b.get("coinsurance", ""),
                b.get("lifetime", ""), b.get("tier", "")]
        for j, col in enumerate(range(2, 9)):
            c = ws.cell(i, col)
            v = vals[j]
            if v not in ("", None):
                try:
                    c.value = int(str(v).replace(",", ""))
                except ValueError:
                    c.value = v
            c.fill = EDIT_FILL
            c.font = BLACK
        ws.cell(i, 3).number_format = money()
        ws.cell(i, 5).number_format = money()
    return ws


def sheet_prem(wb, prem_rows):
    ws = wb.create_sheet("prem")
    ws.cell(1, 1, "key").font = HDR
    for a in range(N_AGE_COLS):
        ws.cell(1, PREM_FIRST_COL + a, a).font = HDR
    for i, row in enumerate(prem_rows, start=2):
        ws.cell(i, 1, row[0])
        for j, v in enumerate(row[1:]):
            if v is not None:
                ws.cell(i, PREM_FIRST_COL + j, v)
    ws.sheet_state = "hidden"
    return ws


def sheet_notes(wb, stats):
    ws = wb.create_sheet("Notes")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 78
    ws["A1"] = "Provenance and assumptions"
    ws["A1"].font = TITLE
    rows = [
        ("Source", "data.gov.hk, Health Bureau VHIS open data (4 JSON resources)"),
        ("Built", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Plan variants", stats["variants"]),
        ("Insurers", stats["insurers"]),
        ("Premium series rows", stats["prem_rows"]),
        ("Latest premium date", stats["prem_date_max"]),
        ("", ""),
        ("Age basis A", "Age last birthday (attained). Table age = client age."),
        ("Age basis N", "Age next birthday. Table age = client age + 1."),
        ("Age basis R", "Age nearest birthday. APPROXIMATED as client age; "
                        "exact resolution needs date of birth."),
        ("", ""),
        ("Basis S", "STANDALONE policy (獨立保單). Confirmed by the Health "
                    "Bureau data dictionary for the Standard Premium dataset. "
                    "Present for 530 of 579 variants."),
        ("Basis R", "RIDER / supplementary benefit attached to another policy "
                    "(附屬保單). Same source. Present for 184 variants, all from "
                    "insurers that also sell VHIS on top of a life policy. Never "
                    "priced above the standalone rate: lower on 87, identical on "
                    "42, higher on 0. Quote R only when the plan is genuinely "
                    "being sold as a rider."),
        ("Basis fallback", "If the chosen basis has no table for a plan, the "
                           "other basis is used. See the Basis Used column."),
        ("", ""),
        ("Excluded from premiums", "Insurance Authority levy; underwriting "
                                   "premium loading; payment-mode loading "
                                   "(monthly is roughly +8%/yr); any discounts."),
        ("Age ceiling", "Source data stops at each insurer's new-application "
                        "age limit. Renewal-only ages are absent, so 10-year "
                        "figures truncate from client age 72."),
        ("Benefit fields", "Ward, deductible and geography are parsed from the "
                           "plan label where possible and otherwise must come "
                           "from the plan PDF. The Benefits sheet overrides."),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        ws.cell(i, 1, k).font = Font(name=ARIAL, size=10, bold=bool(k))
        c = ws.cell(i, 2, v)
        c.font = BLACK
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default="/mnt/user-data/uploads")
    ap.add_argument("--out", default="VHIS_Compare.xlsx")
    ap.add_argument("--benefits", default="data/plan_benefits.csv")
    ap.add_argument("--overrides", default="data/manual_overrides.csv")
    args = ap.parse_args()

    std, flexi, prem, providers = bv.load(args.data_dir)
    prem_by_cert = {r["certification-no"]: r for r in prem}
    catalog, _ = bv.build_catalog(std, flexi, prem_by_cert, providers)
    catalog.sort(key=lambda r: (r["insurer"], r["plan_name"], r["plan_level_raw"]))

    prem_rows = build_premium_rows(prem_by_cert)
    prem_date_max = max(r["prem-date"] for r in prem)
    prem_date_max = (f"{prem_date_max[:4]}-{prem_date_max[4:6]}-{prem_date_max[6:]}")

    wb = Workbook()
    wb.remove(wb.active)

    sheet_client(wb, len(catalog), prem_date_max)
    sheet_compare(wb, catalog, len(prem_rows))
    benefits = load_benefits(args.benefits, args.overrides)
    sheet_benefits(wb, catalog, benefits)
    sheet_prem(wb, prem_rows)
    sheet_notes(wb, {
        "variants": len(catalog),
        "insurers": len({r["insurer"] for r in catalog}),
        "prem_rows": len(prem_rows),
        "prem_date_max": prem_date_max,
    })

    wb.save(args.out)
    print(f"wrote {args.out}")
    print(f"  compare rows : {len(catalog)}")
    print(f"  premium rows : {len(prem_rows)} (vs 234,464 in the long CSV)")
    print(f"  benefit rows : {len(benefits)} merged from scrape + overrides")


if __name__ == "__main__":
    main()
