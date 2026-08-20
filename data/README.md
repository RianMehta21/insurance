# `data/` — the files that must survive a rebuild

Everything else in this project is regenerated from scratch on every run.
These four are not, and losing them costs either 40 minutes or real work.

| File | What it is | If you lose it |
|---|---|---|
| `plan_state.json` | Reduced form of last run's source data: one entry per plan variant with its premium date and its premium at age 40. | The next change report has nothing to diff against and comes out empty. Rebuilds itself the run after. |
| `plan_benefits.csv` | Everything `scrape_plans.py` pulled out of the plan PDFs. Keyed on `certification_no` + `plan_date`. | All 579 PDFs get re-downloaded and re-parsed, roughly 40 minutes. |
| `review_queue.csv` | The subset of the above where the scraper could not find the core fields. This is the worklist. | Regenerated on the next scrape. |
| `manual_overrides.csv` | Corrections typed by hand. **Wins over anything scraped.** | Gone for good. Nothing else knows these values. |

## manual_overrides.csv

The one file here meant to be edited by a person. Add a row per plan you want
to correct; leave any column blank to leave that field alone.

```csv
certification_no,ward,deductible,geography,annual_limit,coinsurance,lifetime,note
F00022-01-000-03,Semi-private,25000,Asia,1000000,20,,checked against brochure 2026-08
```

Two things worth knowing:

- **Do not type corrections into the Excel file.** The `Benefits` sheet is
  rebuilt from these CSVs on every refresh, so anything typed into the
  workbook is silently overwritten the next time the pipeline runs. That is
  why the sheet carries a warning in row 1.
- **Deductible usually has to come from here.** Certified plan documents state
  the deductible as "*as stated in the Policy Schedule*" — a per-client form,
  not part of the certified terms. On 40 real PDFs the scraper recovered an
  amount zero times. Where the plan-level label in the source JSON does not
  carry it either (about 139 variants with names like "Superior" or "Benefit
  Level 3"), this file is the only place it can come from.
