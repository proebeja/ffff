# Project Luma — databook in English, Option A (run 2)

Run 2 changes the reporting language to English. The numbers are unchanged
from run 1; what changed is which half of the workbook is visible and what
language the working papers speak.

## The P&L is not in the file

The file attached to this request is **byte-identical to the one from run 1**
(MD5 `0376ff27a8566e7809aa9eb1566960db`, 87.538 bytes, four tabs, 221 rows
each). It carries class codes 10 to 50 only — current assets, non-current
assets, current liabilities, non-current liabilities, equity. There is no
class for income or expenses, and `MajorGroup` never takes a value of 4 or 5.

So the P&L did not come through. Nothing was built for it, because there is
nothing to build from: the account groups of a P&L export decide how each
line maps, and guessing them would produce a Lead P&L that looks finished and
is invented.

What stays in place until it arrives:

* **Lead P&L is empty**, and says so. Row 264 carries the result per the
  source, row 265 shows it against an empty P&L. That difference is the
  missing statement, expressed as a control rather than a silence.
* **The equity roll forward does not depend on the P&L.** The result comes
  from account `3-90000 Current Earnings`, so the Lead NA closes without it.
* **The status line for every column** reads "no P&L in the source — classes
  10 to 50 are balance sheet only".

Send the P&L export and the Lead P&L fills from the same mechanism. Expect one
piece of work on top: the account groups of the P&L need the same crosswalk
treatment the balance sheet groups got, because `Freight & Cartage` matches no
German keyword any more than `Trade Debtors` did.

## What "in English" now means

**The template's own mechanism.** Language is a matter of showing and hiding,
never of translating — both versions stay in the file. The German header block
(rows 4 and 5) is hidden, the English one (rows 6 and 7) is shown, column C
(German labels) is hidden and column D (English) is visible. That applies to
all 26 tabs that carry both blocks.

**The working papers of this run** are written in English throughout: sheet
names (`Mapping`, `Review queue`, `Status by column`, `Assumptions`,
`Open items`, `Behaviour check`), their column headers, the QA findings, the
status per column, the control rows and the run report.

**The four line items created from dummy rows** carry proper English labels
taken from the template's own wording plus a qualifier that says how they
differ: `Receivables from affiliated/ participating companies (net debt)`,
`Tax provisions (net debt)`, `Bonds`, `Deferred tax assets`.

**Two things stay German, deliberately.** The two vocabulary columns of the
`Mapping` sheet hold classification terms, not prose: the house vocabulary on
one side and the template's own Ticker 1 wording on the other. The template's
tickers are German, and a ticker must match the mastersheet character for
character or the `SUMIFS` returns zero without saying so — translating that
column would break the workbook, and translating only the other one would
empty the sheet of its purpose. The `Reason` column of the review queue
carries the engine's audit trail per account, generated German prose shared
with every other mandate; translating it would mean translating the engine,
not this databook.

## Acceptance

| Criterion | Result |
| --- | --- |
| Cell-by-cell format identity with the template | 0 deviations across 38 sheets |
| Newly created cells | 8 (tickers of the four dummy rows) |
| Row numbers | Lead NA 237, Lead PL 265 — unchanged |
| No plain `SUM` over account slots | none |
| Error count after recalc | 142 — exactly the level of the empty template, **0 new** |

## Control rows

| Control | FY2023 | FY2024 | FY2025 | FY2026 |
| --- | ---: | ---: | ---: | ---: |
| Balance sheet identity | −0.01 | −0.01 | −0.04 | −0.04 |
| Net assets + equity | −0.01 | −0.01 | −0.04 | −0.04 |
| Accounts with no line item in the lead | 0.00 | −0.01 | 0.00 | 0.00 |
| Equity roll forward incl. residual | 0.01 | 0.00 | 0.03 | −0.00 |
| Unexplained movement in equity | 0.00 | **268,375.20** | −0.00 | −0.00 |
| Accounts without an account slot | 1,430,614 | 1,995,711 | 1,965,742 | 1,866,123 |

Both open findings are unchanged from run 1 and documented there
(`protokoll/luma_lauf1.md`): the FY2024 movement in equity that neither
capital nor result explains, and the eight account slots per line item that do
not stretch to a 198-account MYOB chart.
