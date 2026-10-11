# VHIS Plan Comparison

Compares all 579 certified Voluntary Health Insurance Scheme plans from 33 Hong Kong
insurers in one spreadsheet. Type a client's age and every plan reprices.

**Download the spreadsheet:**
https://github.com/RianMehta21/insurance/releases/latest/download/VHIS_Compare.xlsx

---

## For everyday use

You only need **`VHIS_Compare.xlsx`**. Nothing else.

1. Open the **Client** sheet. Check the build date, then fill in the four yellow
   cells: age, gender, smoker, premium basis.
2. Go to the **Compare** sheet. All 579 plans have already repriced.
3. Use the filter arrows in **row 4**. Set `Available = Open` first, then narrow
   by **Ward Class** (General Ward / Semi-private / Private / No restriction) and
   **Region** (Greater China / Asia / Worldwide excl. USA / Worldwide). The
   *Detail* column next to each keeps the insurer's exact wording.
4. Sort by **10-Year Avg**, not First Year.

   A few cells are words, not numbers, and each means something specific:
   **No limit** (the plan has no annual or lifetime cap), **Per policy** (the
   deductible is chosen per policy, ask the insurer), **Not stated** (the plan
   document gives no single figure). The Notes sheet lists them all.
5. Copy your shortlist into your own file.

The file does not update itself. Download a fresh copy before each client meeting.

Don't save notes into the workbook — it is rebuilt from scratch every refresh and
anything typed into it is lost.

---

## Where the numbers come from

Everything starts from Hong Kong government open data on data.gov.hk, published by
the Health Bureau. Four files: the certified Standard plans, the certified Flexi
plans, the premium tables, and the list of registered insurers.

Nothing is estimated or modelled. Every premium is a figure the insurer published.

### How a premium is chosen

Each insurer publishes a table of annual premiums by age, and they don't all count
age the same way. The **Age Basis** column shows which convention applies:

| Basis | Meaning | Age used |
|---|---|---|
| A | Age last birthday | the client's actual age |
| N | Age next birthday | actual age + 1 |
| R | Age nearest birthday | approximated as actual age |

Rows marked `APPROX` use basis R. Getting those exact needs the client's date of
birth, so treat them as indicative.

**Premium basis S or R** (the yellow cell on the Client sheet) is a different thing:
`S` is a standalone policy, `R` is a rider attached to another policy. The insurers'
own premium schedules print these two columns side by side as 基本計劃 *Basic plan*
and 附加契約 *Policy rider*.

It matters. Of the 579 variants, 444 are priced on only one basis, so the choice is
ignored and the other one is used automatically — the **Basis Used** column tells you
which. But 135 carry both, and there R runs about 10% below S at the median and up to
17.7% below. Quote R only if the plan is genuinely being sold as a rider.

### First Year and 10-Year Avg

- **First Year** — the published premium at the client's table age.
- **10-Year Total** — that year plus the next nine, read straight off the table.
- **10-Year Avg** — the total divided by the number of years actually available.
- **Years Avail** — how many of those ten years the insurer publishes.

Sort by 10-Year Avg. The cheapest plan in year one is often not the cheapest over
ten, because insurers age-band differently.

Watch `Years Avail`. Insurers stop publishing rates past their new-application age
ceiling, so from about age 75 the ten-year window is incomplete for most plans and
the average covers fewer years.

### What the premiums exclude

- the Insurance Authority levy
- any underwriting loading for medical history
- any discount
- payment-mode loading — paying monthly typically costs about 8% more per year

They are the insurers' published standard premiums. **They are not quotes.**

### Available

69 of the 579 plans cannot be sold to a new client, but they still carry full
premium tables and sort exactly like live plans. Right now the single cheapest
plan in the file is one of them.

| Shows | Meaning |
|---|---|
| Open | Can be sold to a new client |
| Renewal only | Existing policyholders may renew; closed to new business |
| De-registered | Certification withdrawn by the Health Bureau |
| Withdrawn | The insurer has stopped offering it |
| Partly restricted | Some benefit levels are closed |

Filter to `Open` before shortlisting. Rows you can't sell are shaded orange.

---

## How the pipeline works

Runs itself every Monday morning, Hong Kong time. Nobody has to do anything.

```
1. Check      Has data.gov.hk published anything new?
              No  -> stop. Four requests, costs nothing.
              Yes -> carry on.

2. Build      Flatten the JSON into one row per plan variant.

3. Scrape     Download the plan PDFs and read out ward class,
              geography, deductible, coinsurance and benefit limits.
              Only plans whose certification date moved are re-read.

4. Schedules  For the plans just scraped, read the Benefit Schedule:
              ward class, geography, annual and lifetime limits,
              coinsurance. Fills gaps only, never changes a value.

5. Worklist   Refresh the list of plans still missing a key field.

6. Workbook   Build VHIS_Compare.xlsx.

7. Verify     Recalculate every formula in LibreOffice. If any cell
              errors, refuse to publish. Otherwise publish that
              calculated copy, so previews that do not calculate
              (Quick Look, phones, email) show prices instead of 0.

8. Publish    Replace the file at the download link above, and email
              the change report if the email secrets are set.

9. Keep-alive Re-enable the schedule, and add an empty commit after 45
              quiet days, so GitHub never switches the job off.
```

Nothing in it needs a person. Wording no rule recognises shows as `Other`
(with the exact text beside it) instead of failing the build, and a schedule
the reader cannot parse leaves that field as unknown rather than stopping the
publish. A failed run emails the repository owner — GitHub does that by
default.

Ward class and geography come from the plan PDF first; when the PDF only confirms
a restriction without naming it, the certified plan-level label is checked too
(insurers often put it right in the name, e.g. `"Semi-Private Room"` or
`"(Worldwide except USA)"`). Only when neither source names it does a plan show
`Restricted (see doc)` / `Limited (see doc)` — an honest "go check the PDF," not
a guess. As of October 2026 that is 10 plans on ward (AIA SelectWise, whose ward
class depends on whether your doctors are in AIA's network) and none on
geography. Lifetime limit works the same way and has no gaps left either: 168
plans show the actual cap, the rest have none.

Those counts used to be 325 ward and 239 geography. The gap closed because the
clause the scraper reads usually does not state the answer — it points at the
plan's **Benefit Schedule**, which does:

```
(d) Choice of ward class
    ... subject to the restriction in the choice of ward class as stated in
    Section 3 of Supplement 5 and the Benefit Schedule of these Terms and
    Benefits.
```

`schedule_fields.py` reads that schedule row — ward class, geography and
lifetime limit. It has been run for every insurer that had a gap. A plan that
is re-certified gets a new certificate number and starts over from the PDF
clause, so after a refresh that adds plans it is worth re-running for that
insurer.

Standard Plan is solved outright: it's one government-mandated product across
all 33 insurers (Worldwide, no ward restriction, HKD420,000 annual limit, no
deductible, 30% coinsurance, no lifetime limit), so those fields are known
constants, not scraped guesses. What's left to chase is almost entirely **Flexi**
plans, where terms genuinely vary by insurer.

Deductible is the hardest of the three: most plan documents say it is "as stated
in the Policy Schedule", meaning it is set per policy and simply isn't in the
document, PDF or plan name. Those have to be filled in by hand — see below.

---

## Filling in what's missing

`data/manual_overrides.csv` is the one file to edit. It's a to-do list that keeps
itself current: a row appears for every plan still missing a key field, and drops
off the "needs filling" count once you fill it. Anything you type is never
overwritten and always beats the scraped value.

Open it in Excel. Each row names the insurer, plan and level, and links to the
plan PDF. The `still_missing` column says what to fill.

```
certification_no,insurer,plan_name,plan_level,plan_doc_url,still_missing,ward,deductible,geography,...
F00022-01-000-03,AIA International,VHIS Flexi,Semi-private (HKD),https://...,deductible,,25000,,
```

Fill the blank cells, save as CSV, commit. The next build picks it up.

Every plan that is missing something gets a row, so you never have to look up a
certification number by hand. The **priority** column tells you where to start.

### What CHECK FIRST and when needed actually mean

The build prices every sellable plan for **18 test clients** — ages 25, 30, 35, 40,
45, 50, 55, 60, 65, each as male and female, non-smoker. For each of those 18, it
takes the **20 cheapest plans** in each currency (HKD and USD ranked separately).

`shortlist_hits` is how many of those 18 clients would have seen that plan in their
cheapest 20. The most any plan can score is 18.

| priority | rule | means |
|---|---|---|
| `CHECK FIRST` | 1 or more hits | at least one ordinary client would see this plan near the top of their list, so the gap will show up in a real comparison |
| `when needed` | 0 hits | no ordinary client sees it in their cheapest 20. It only surfaces if you filter for something specific — a particular insurer, ward class or currency |

`when needed` does **not** mean the plan is bad or that the gap is fake. It means
nobody hits it by accident, so filling it in advance is wasted effort. If a client
does ask about one, fill that row then.

Right now the split is decisive — there is no grey middle:

```
18 hits    3 plans  CHECK FIRST
17 hits    1 plan   CHECK FIRST
16 hits    1 plan   CHECK FIRST
13 hits    1 plan   CHECK FIRST
 6 hits    1 plan   CHECK FIRST
 2 hits    1 plan   CHECK FIRST
 1 hit     1 plan   CHECK FIRST
 0 hits  184 plans  when needed
```

So: **fill the 9 first**

Anything you type is kept forever and always beats the scraped value.

To fix something that's wrong rather than missing, add the certification number as
a new row and type the correct values. It will override whatever was scraped.

> `data/review_queue.csv` is a different thing — a technical record of which PDFs
> the parser struggled with. It does not know about your corrections, so entries
> stay in it even after you fix them. Use `manual_overrides.csv` as the to-do list.

---

## Running it yourself

```bash
pip install -r requirements.txt          # also needs poppler-utils and libreoffice-calc

python refresh.py --pull                 # fetch the latest data
python build_vhis.py                     # flatten to CSVs
python scrape_plans.py --catalog out/plans_catalog.csv   # read the plan PDFs (slow)
python worklist.py                       # refresh the to-do list
python make_workbook.py                  # build the workbook
python verify_workbook.py VHIS_Compare.xlsx              # check every formula
```

To publish without waiting for Monday: **Actions → Refresh VHIS workbook → Run
workflow**, leaving "Rebuild and publish" ticked.

### Update emails

When the government data actually changes, the Monday run emails the download
link and the change report. It is off until three repository secrets exist
(**Settings → Secrets and variables → Actions**):

| Secret | Value |
|---|---|
| `MAIL_USERNAME` | the Gmail address that sends it |
| `MAIL_PASSWORD` | a Gmail [app password](https://myaccount.google.com/apppasswords) for that account — not its normal password |
| `MAIL_TO` | who receives it; separate several with commas |

A manual republish or a quiet week sends nothing.

### Re-reading Benefit Schedules by hand

The weekly run already does this for every newly scraped plan (fill-only).
Run it by hand only to re-check a whole insurer and review the evidence first:

```bash
# 1. propose values, with the schedule text each one came from
python schedule_fields.py --insurers "Prudential,Chubb" \
    --out data/schedule_candidates.csv

# 2. read the *_evidence columns, then merge
python schedule_fields.py --insurers "Prudential,Chubb" \
    --apply data/manual_overrides.csv
```

Step 2 only fills `(see doc)` placeholders and never overwrites a value typed
by hand. Add `--correct-wrong` to also replace a *scraped* value that
disagrees with the schedule; it prints every replacement, so check them.

---

Source: Hong Kong Health Bureau VHIS open data via data.gov.hk. Figures are the
insurers' published standard premiums, not quotations. Always confirm with the
insurer before advising a client.
