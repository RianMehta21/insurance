# `data/` — the files that must survive a rebuild

Everything else in this project is regenerated from scratch on every run.
These four are not, and losing them costs either 40 minutes or real work.

| File | What it is | If you lose it |
|---|---|---|
| `plan_state.json` | Reduced form of last run's source data: one entry per plan variant with its premium date and its premium at age 40. | The next change report has nothing to diff against and comes out empty. Rebuilds itself the run after. |
| `plan_benefits.csv` | Everything `scrape_plans.py` pulled out of the plan PDFs. Keyed on `certification_no` + `plan_date`. | All 579 PDFs get re-downloaded and re-parsed, roughly 40 minutes. |
| `review_queue.csv` | Technical record of which PDFs the parser struggled with. **Not** the worklist — it does not know about your corrections, so entries stay after you fix them. | Regenerated on the next scrape. |
| `manual_overrides.csv` | The to-do list *and* the corrections file. Maintained by `worklist.py`: a row appears for every plan still missing a key field, and your typed values are never overwritten. **Wins over anything scraped.** | Gone for good. Nothing else knows these values. |

## manual_overrides.csv

The one file here meant to be edited by a person, and the only to-do list worth
working from. `worklist.py` keeps it current on every run.

It is deliberately short. A plan earns a row only if it is still missing a key
field **and** actually reaches a client shortlist — scored by pricing every plan
across 18 client profiles and counting how often it lands in the cheapest 20 of
its currency. 194 plans are missing a deductible; 9 of them ever get quoted, so
9 is what you see. `--min-hits 0` lists the rest.

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
  carry it either (about 139 variants with names like "Superior" or "Benefit
  Level 3"), this file is the only place it can come from.
