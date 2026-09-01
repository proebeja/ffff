# Project Luma — the databook rebuilt as Option B (run 4)

Same four sources and the same numbers as run 3. What changed is the way the
workbook is wired: a third layer between the mastersheet and the leads.

## What Option B actually means here

Option A has two layers. Every line item in Lead NA pulls its total straight
out of the mastersheet with a `SUMIFS` on the line-item ticker, and the eight
collapsed account rows underneath pull the same figures account by account.

Option B puts a schedule in between. The line item now reads

```
Lead NA!F18   ='NA_Sachanlagen'!E20
Lead NA!F40   ='NA_Vorräte'!F22
```

and the schedule itself carries the breakdown, each of its rows a visible
`SUMIFS` over the accounts that belong to it. The collapsed account rows
underneath the line item are untouched — they still hang off the mastersheet.

That is the whole point. Two independent paths onto the same figure: down
through the schedule, and straight from the accounts. If they disagree, the
schedule is missing an account or carrying one twice.

## Which schedules can actually carry a line item

The template has six NA schedules. Only two of them can be inverted.

| Schedule | Carries its line item? | Why |
| --- | --- | --- |
| `NA_Vorräte` | yes, Lead NA 40 | breaks inventories down by component, and components are in the chart of accounts |
| `NA_Sachanlagen` | yes, Lead NA 18 | breaks tangible assets into asset classes at cost and accumulated depreciation — exactly how the chart is built |
| `NA_Ford LuL` | no | an ageing analysis. A trial balance has no due dates. |
| `NA_Verb LuL` | no | same |
| `NA_TWC` | no | a roll-up: every row already pulls from Lead NA |
| `NA_Net Debt` | no | same |

The two ageing schedules are a **data request**, not a gap in the run. Filling
them needs an open-item listing. Until one arrives, trade receivables and trade
payables stay on the Option A path — a line item pulling from an empty schedule
would show a clean, round zero, which is worse than showing the figure.

`NA_TWC` and `NA_Net Debt` cannot be inverted at all. They read the leads and
add normalisation blocks on top; pointing the leads back at them would make
them reference themselves.

## The two schedules as they are now filled

**Inventories.** The template lays this schedule out by product — ten product
slots, an "other products" row and a valuation allowance. The source has no
product dimension, so the ten slots are relabelled as components:

```
  9 Inventory – production                    2,140.3  3,013.9  2,949.5  3,049.5
 13 Work in progress – materials and parts      889.1    883.8  1,133.7    521.9
 14 Work in progress – labour                   286.6    328.1    329.1    218.8
 16 Finished goods                               76.7     31.6    115.2     -5.7
 17 Finished goods – labour                      71.8     71.8     71.8      0.0
 19 Inventory components                      3,464.5  4,329.2  4,599.3  3,784.4
 21 Valuation allowance for inventories            —     -195.5   -317.1   -317.1
 22 Inventories (→ Lead NA 40)                3,464.5  4,133.7  4,282.3  3,467.3
```

The two provision accounts go into row 21, which the template already holds for
exactly that. Six of the fourteen inventory accounts are nil in all four
periods; they keep their slot in the mapping so the schedule stays a complete
picture of the chart, and simply show nothing.

**Tangible assets.** The template holds three asset classes, each with a cost
row and an accumulated-depreciation row. The chart carries every class twice
("– At Cost" / "– Acc Dep"), so this is the one place where source and template
agree on the layout without any interpretation. 51 accounts become three
classes:

```
  9/10 Prototypes, tools and R&D equipment            21 cost, 10 acc. dep.
 12/13 Office equipment, furniture and fittings       11 cost,  4 acc. dep.
 15/16 Leasehold improvements and right-of-use assets  3 cost,  2 acc. dep.
 18    Tangible assets (historical costs)   2,392.6  2,804.8  3,279.3  2,721.5
 19    Accumulated depreciation            -2,168.2 -2,082.4 -2,939.5 -2,022.2
 20    Tangible assets (→ Lead NA 18)         224.4    722.4    339.9    699.3
```

The template's own "degree of wear" ratio (row 23, accumulated depreciation
over cost) now shows −91%, −74%, −90%, −74%. The minus sign is the template's:
the book-value rows need accumulated depreciation carried negative, and the
ratio then inherits that sign. Read it as the absolute value.

## The mandatory control

Option B is only worth the extra layer if the two paths are checked against
each other, so the check is written by the export itself, not by the mandate.
The sheet `Schedule control` carries it, per line item and per period:

| Line item | Schedule | FY2023 | FY2024 | FY2025 | FY2026 |
| --- | --- | ---: | ---: | ---: | ---: |
| Inventories | `NA_Vorräte` | 0.00 | 0.00 | 0.00 | 0.00 |
| Tangible assets | `NA_Sachanlagen` | 0.00 | 0.00 | 0.00 | 0.00 |

It also lists, in its own column, any account of the line item that no schedule
row claims. Both rows are empty — all 14 inventory accounts and all 51 tangible
asset accounts are covered. The schedule side is summed **with multiplicity**,
so an account accidentally sitting in two schedule rows shows up as a
difference rather than cancelling out against a set.

One consequence worth naming: `NA_Sachanlagen` has a check row of its own
(row 30), which compares the schedule total against Lead NA 18. Under Option A
that is a real control. Under Option B it compares the schedule against a cell
that now reads the schedule, so it is 0.00 by construction. That is precisely
why the control above had to be built somewhere else.

## Two details of the wiring

**The opening column stays on the mastersheet.** Lead NA column E is the
opening column; the schedules of the template start at the first reporting
period and have no counterpart for it. Column E therefore keeps the template's
own `SUMIFS`. Same source, one layer fewer, and for this mandate it is empty
anyway.

**Column positions are read, never counted.** `NA_Vorräte` starts its first
period in column F, `NA_Sachanlagen` in column E, and Lead NA in F. Each
schedule's header row is parsed for its `=Cockpit!…39` references and the
column derived from that. Counting would have shifted inventories by a year.

## Acceptance

| Criterion | Result |
| --- | --- |
| Cell-by-cell format identity with the template | 0 deviations across 38 sheets |
| Error count after recalc | 140 against 143 for the empty template — **0 new, 3 fewer** (identical to Option A on the same data) |
| Schedule control, both line items, all periods | 0.00 |
| Accounts of a line item with no schedule row | none |
| Lead NA 18 after recalc | 224.4 / 722.4 / 339.9 / 699.3 T AUD — equals the account sums to the cent |
| Lead NA 40 after recalc | 3,464.5 / 4,133.7 / 4,282.3 / 3,467.3 T AUD — likewise |
| Net assets (Lead NA 223) | 7,106.0 / 7,125.5 / 5,248.4 / 1,724.0 T AUD, unchanged from Option A |
| Balance sheet identity | −0.01 to −0.04 in all periods |
| Lead NA check row 237 | 0.00 in FY2023, −268.38 T AUD from FY2024 (unchanged finding) |
| Test suite | 300 tests, all passing |

## Unchanged from run 3

Everything the numbers say is the same as in run 3 — the reconciliation against
the consolidated financial statements, the 192 nil accounts kept off the
400-row mastersheet, the eight-slot limit on account detail, and the
268,375.20 movement in FY2024 equity that neither capital nor result explains.
Option B changes how the workbook is wired, not what it reports.

One correction to run 3's reconciliation table: FY2026 net assets are
1,723,987 and the difference to the draft accounts −633,997. The table there
carries 1,723,990 and −633,994, from a run made before the last fix of that
session landed. The conclusions are unaffected.
