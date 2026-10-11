# `data/` — the files that must survive a rebuild

Everything else in this project is regenerated from scratch on every run.
These four are not, and losing them costs either 40 minutes or real work.

| File | What it is | If you lose it |
|---|---|---|
| `plan_state.json` | Reduced form of last run's source data: one entry per plan variant with its premium date and its premium at age 40. | The next change report has nothing to diff against and comes out empty. Rebuilds itself the run after. |
| `plan_benefits.csv` | Everything `scrape_plans.py` pulled out of the plan PDFs. Keyed on `certification_no` + `plan_date`. | All 579 PDFs get re-downloaded and re-parsed, roughly 40 minutes. |
| `review_queue.csv` | Technical record of which PDFs the parser struggled with. **Not** the worklist — it does not know about your corrections, so entries stay after you fix them. | Regenerated on the next scrape. |
| `manual_overrides.csv` | The to-do list *and* the corrections file. Maintained by `worklist.py`: a row appears for every plan still missing a key field, and your typed values are never overwritten. **Wins over anything scraped.** | Gone for good. Nothing else knows these values. |
| `schedule_candidates.csv` | Audit trail for the ward, geography and lifetime-limit values `schedule_fields.py` read out of the Benefit Schedules, each with the verbatim schedule text it came from. Regenerable, but only against the same certification versions. | Re-run `schedule_fields.py`. Values already merged into `manual_overrides.csv` are unaffected. |

## manual_overrides.csv

The one file here meant to be edited by a person, and the only to-do list worth
working from. `worklist.py` keeps it current on every run.

Every plan missing a key field gets a row, so there is always something to fill
in and never a certification number to hunt down. The `priority` column ranks
them.

`shortlist_hits` counts how many of 18 test clients (ages 25-65 in fives, male
and female, non-smoker) would see that plan among the 20 cheapest in its
currency, counting only plans that can still be sold. Maximum 18.

  CHECK FIRST   1+ hits - an ordinary client meets this plan, so fill it
  when needed   0 hits  - only reachable by filtering deliberately; fill on demand

Of 193 outstanding plans, 9 are `CHECK FIRST` and 184 are `when needed`. Every
one of the 193 is missing only `deductible` — ward and geography are covered
for every plan, one way or another (see "Ward and geography" below), so they
never land on this list.

The file holds 399 rows, not 193: the other 206 carry no outstanding gap but
do carry a recorded ward/geography correction, and a row with a value in it is
kept forever. `still_missing` is blank on those.

Ranking comes from pricing every plan across 18 client profiles and counting
shortlist appearances among plans that can still be sold. `--min-hits 1` trims
the file to just the `CHECK FIRST` rows if the full list is in the way.

Two rules it never breaks: a row carrying a value you typed is kept forever, even
if that plan would otherwise be filtered out, and no value you entered is ever
altered or removed.

Each row identifies the plan and links to its PDF. `still_missing` says what to
fill; leave any other column blank to leave that field alone.

```csv
certification_no,insurer,plan_name,plan_level,plan_doc_url,still_missing,ward,deductible,...
F00022-01-000-03,AIA International,VHIS Flexi,Semi-private (HKD),https://...,deductible,,25000,...
```

To correct something that is wrong rather than missing, add the certification
number as a new row and type the right values.

Two things worth knowing:

- **Do not type corrections into the Excel file.** The `Benefits` sheet is
  rebuilt from these CSVs on every refresh, so anything typed into the
  workbook is silently overwritten the next time the pipeline runs. That is
  why the sheet carries a warning in row 1.
- **Deductible usually has to come from here.** Certified plan documents state
  the deductible as "*as stated in the Policy Schedule*" — a per-client form,
  not part of the certified terms. On 40 real PDFs the scraper recovered an
  amount zero times. Where the plan-level label in the source JSON does not
  carry it either (193 variants, all named like "Regular", "Premier" or
  "Plan 1" — a marketing tier name with no figure to parse out, checked
  digit-by-digit, not just eyeballed), this file is the only place it can
  come from.

## Ward and geography: why so many say "(see doc)"

`scrape_plans.py` reads two clauses per PDF — *Choice of ward class* and
*Territorial scope of cover*. Insurers write one of two templates: either "not
subject to any restriction" (a clean answer — `No restriction` / `Worldwide`),
or "subject to the restriction ..." which only confirms a restriction exists
without naming it inline; the actual ward class or region usually lives in a
separate benefit-schedule table the scraper doesn't read. That's what
`Restricted (see doc)` and `Limited (see doc)` mean: confirmed-but-unnamed,
not "unknown."

Before falling back to that sentinel, `make_workbook.py` now checks whether
the certified plan-level label already names the answer — `build_vhis.py`
parses it into `ward_type` / `geographical_coverage` on every row. Insurers
often do bake it in, e.g. `"Semi-Private Room"` or `"HKD16,000 Deductible
(Worldwide except USA)"`, and when they do, that specific value is shown
instead of the placeholder. It only stays `(see doc)` when neither the PDF
clause nor the plan's own name says which one — as of October 2026, 10 plans
for ward (AIA SelectWise: the ward class depends on whether the attending
doctors are in AIA's network, so no single value is right) and none for
geography.

### Reading the Benefit Schedule instead

Those counts were 239 and 325 until `schedule_fields.py` was run. The clause
that comes back "restricted but unnamed" almost always names *where* the
answer is, and it is the plan's own Benefit Schedule — which every certified
variant has its own PDF of. So the answer is per-variant and it is there:

```
Entitled ward class                General Ward
Geographical limitation            Except for psychiatric treatments ... –
                                   For non-Emergency Treatment: Greater China
                                   For Emergency Treatment: Worldwide
```

It was run for AXA, Bupa, Blue Cross, FWD, Sun Life, HSBC Life and Manulife
first (209 plans), then in October 2026 for the remaining ten insurers with a
gap — AIA, Chow Tai Fook, Prudential, Chubb, Zurich, China Taiping, BOC, YF
Life, China Life, Liberty, plus geography for Bowtie, CIGNA, Allianz and Asia
Insurance. The same pass also reads the **lifetime benefit limit**, which had
exactly the same gap: 210 plans said only "Capped (see doc)" because Part 6
1(b) defers to the schedule. A schedule that says "Nil" there means no limit,
the same as a clause saying "not subject to any Lifetime Benefit Limit" (Chubb
and Asia iCare print both). HSBC is the case to know: its Bronze level caps at
HKD20,000,000 and every higher level prints "Nil".

Where a plan has no schedule row at all, the geography is taken from its
supplement's prose ("subject to a geographical limitation of Asia") and
labelled `prose (no schedule row)` in the candidates file. Allianz, Asia
Elite, Chubb's VHIS Series Flexi and YF Life are that shape.

Three things worth knowing before re-running it:

- **There is no single format.** Seven insurers used four different labels for
  the ward row (`Entitled` / `Designated` / `Restricted` / `Covered ward
  class`) and three for geography. Bupa lays the ward row out in three
  columns; Manulife centres the label vertically so most of the value sits
  *above* the words naming it; FWD splits a cell across a page break.
- **Values vary by benefit level, not just by product.** HSBC's Flexi Plan
  runs General Ward/Greater China at Bronze up to Standard Private/Worldwide
  excluding the USA at Platinum. Reading one variant per product would be
  wrong for every other level.
- **Some plans restrict by location too**, and both tiers are recorded:
  `Semi-private (HK/Australia/NZ); Standard Private (elsewhere)`. Same for the
  emergency split on geography: `Asia (Worldwide for emergency)`.

`schedule_fields.py --apply` merges the reviewed values in. It only fills
placeholders, never a value typed by hand. `--correct-wrong` additionally
replaces a *scraped* value that disagrees with the schedule, printing each
one — that is how AXA WiseGuard Pro's 24 variants were fixed, having all read
`Hong Kong/Macau` because the clause-level regex matched that phrase inside
the list of benefits **excluded** from the limitation. Their real scopes are
Greater China, Asia, or Worldwide excluding USA depending on level.

**What it deliberately leaves alone.** A clause reading "not subject to any
restriction in the choice of ward class" is a real answer, not a gap, even
when that plan's schedule does print a `Designated Ward Class` — there it
sets the room-and-board benefit limit and names the pricing tier rather than
restricting anything. Manulife First and Bupa MyFlexi are both that shape
(their levels are literally called "(Semi-private Room)", "(Private Room)"),
so 12 plans keep their `No restriction`. `NEVER_REPLACE` in the script
enforces this.

**Standard Plan is a special case.** All 33 insurers' VHIS Standard Plan
variants certify identical, government-mandated terms — Worldwide, no ward
restriction, HKD420,000 annual limit, no deductible, 30% coinsurance, no
lifetime limit — only the premium differs per insurer. `make_workbook.py` and
`worklist.py` both know these constants and use them to fill in whatever a
Standard Plan's own PDF scrape left blank (never to overrule a real scraped
value, and a manual override still wins over even that).
