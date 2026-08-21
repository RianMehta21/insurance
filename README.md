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
3. Use the filter arrows in **row 4**. Set `Available = Open` first.
4. Sort by **10-Year Avg**, not First Year.
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

4. Worklist   Refresh the list of plans still missing a key field.

5. Workbook   Build VHIS_Compare.xlsx.

6. Verify     Recalculate every formula in LibreOffice. If any cell
              errors, refuse to publish.

7. Publish    Replace the file at the download link above.
```

Ward class and geography come from the plan PDFs and are near-complete. Deductible
is different: most Flexi plan documents say it is "as stated in the Policy
Schedule", meaning it is set per policy and simply isn't in the document. Those
have to be filled in by hand — see below.

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

Currently 234 rows need attention: 189 need only the deductible, 40 only the ward
class, 5 need all three.

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

---

Source: Hong Kong Health Bureau VHIS open data via data.gov.hk. Figures are the
insurers' published standard premiums, not quotations. Always confirm with the
insurer before advising a client.
