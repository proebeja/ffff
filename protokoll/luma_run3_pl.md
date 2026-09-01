# Project Luma — P&L loaded and reconciled to the financial statements (run 3)

Four sources now: the annual trial balance, the annual P&L, the unaudited
accounts for FY2023–FY2025 and the draft statements for FY2026. The Lead P&L
is populated, and the financial statements are the reconciliation target.

## The P&L

237 accounts across classes 60 to 95 — Income, Cost of Goods Sold,
Manufacturing, Expenses, Other Income, Other Expenses. It ties to the trial
balance exactly: the P&L sums to the same figure the balance sheet carries on
`3-90000 Current Earnings`, in all four periods, to the cent.

Three things had to be settled before it could be loaded.

**The result would otherwise be counted twice.** `3-90000 Current Earnings` is
not a posted account; it is the P&L expressed as one number so the balance
sheet balances. With the P&L loaded it is left out of the mastersheet, and the
sum of all accounts goes back to zero (−0.01 to −0.04 rounding, from the
export itself). It is listed on the `Open items` sheet, not silently dropped.

**The periods are out of order in the source.** The P&L tabs run FY2023,
FY2026, FY2024, FY2025. Taking the sheet order would have put FY2026 in the
second column of every lead. Periods are now sorted by reporting date.

**The P&L follows a cost-type layout, and so does the template.** The house
§ 275 vocabulary would have collapsed nine tenths of this P&L into "other
operating expenses". Instead the account groups map onto the template's own
cost-type line items, which the template already carries: premise costs,
repairs, sales costs, delivery costs, insurance, advertising and travel. The
mapping is 65 balance-sheet groups plus 47 P&L groups, with 42 account-level
exceptions where a group is too coarse.

The exceptions matter more than they look. `Administration` holds 63 accounts
and contains, among ordinary overheads, three depreciation accounts, three
interest accounts and six occupancy accounts. Left in the group, EBITDA would
have been wrong by the depreciation — 425 to 656 TAUD a year — and the
financial result would have been empty.

## Lead P&L (T AUD)

```
  9 Revenues                       8,095.3   11,966.1    7,258.5    3,366.8
 39 Total output                   8,095.3   11,966.1    7,258.5    3,366.8
 40 Cost of materials             -3,306.2   -4,536.1   -2,987.2   -1,713.4
 52 Gross profit                   4,789.1    7,430.0    4,271.3    1,653.4
 53 Personnel expenses            -2,428.2   -2,543.8   -3,068.6   -2,688.5
 62 Premise costs                    -53.2      -96.4     -111.8      -93.1
 80 Sales costs                     -852.4     -950.9     -717.7     -286.3
 89 Delivery costs                  -137.3     -237.6     -137.3      -79.2
107 Insurance and levies             -77.1      -84.8      -47.9      -47.8
116 Advertising and travel          -125.5     -344.5     -254.8      -99.9
125 Other operating income            -4.6      235.4      893.1      781.5
134 Other operating expenses      -1,001.9   -2,533.2   -2,448.7   -1,806.3
146 EBITDA                           105.3      870.1   -1,628.7   -2,666.2
147 Depreciation                    -425.4     -486.6     -655.8     -564.3
159 EBIT                            -320.0      383.5   -2,284.4   -3,230.4
190 Financial result                 -25.3       46.4       77.8      -14.3
191 EBT                             -345.4      429.8   -2,206.6   -3,244.7
192 Income taxes                      57.5      160.0      329.5     -279.7
210 Profit / Loss                   -287.8      589.8   -1,877.1   -3,524.4
```

Row 210 matches the trial balance's own result to the cent in every period.

## Reconciliation to the financial statements

The statements are **consolidated** ("and its controlled entities"); the trial
balance is one division. The difference is therefore expected, and it is shown
rather than defined away. In AUD:

| | FY2023 | FY2024 | FY2025 | FY2026 |
| --- | ---: | ---: | ---: | ---: |
| **Net assets** databook | 7,106,030 | 7,125,510 | 5,248,433 | 1,723,990 |
| per accounts | 6,761,557 | 7,391,527 | 5,466,698 | 2,357,984 |
| difference | 344,473 | −266,017 | −218,265 | −633,994 |
| **Revenue** databook | 8,095,332 | 11,966,133 | 7,258,524 | 3,366,846 |
| per accounts | 11,087,006 | 12,307,888 | 6,953,560 | 3,388,866 |
| difference | −2,991,674 | −341,755 | 304,964 | −22,020 |
| **Result** databook | −287,823 | 589,771 | −1,877,076 | −3,524,446 |
| per accounts | 318,663 | 619,069 | −1,900,061 | −3,024,663 |
| difference | −606,486 | −29,298 | 22,985 | −499,783 |

Read it period by period:

* **FY2024 and FY2025 are close** — 29 and 23 TAUD on a result of 0.6 and
  −1.9 million, which is what a consolidation of small subsidiaries looks like.
* **FY2023 is not comparable.** The statements carry a footnote: *"The Company
  was acquired in August 2022; the results of operations are based on average
  operational results over twelve months."* Our column is the nine months from
  1 July 2022 to 31 March 2023. Revenue differs by 3.0 million and the result
  by 606 TAUD, and the period length explains the bulk of it.
* **FY2026 differs by 634 TAUD on net assets and 500 TAUD on the result**, and
  the draft statements name part of the reason themselves: they carry goodwill
  of 963,337 and restricted investments of 50,000, neither of which exists in
  a single-division trial balance. The draft is also internally inconsistent —
  its own YTD column does not equal the sum of its four quarters for foreign
  exchange or gross profit.

Row 264 of the Lead P&L now carries the result per the financial statements
and row 265 shows the difference, so the bridge sits in the workbook and not
only in this note.

## Acceptance

| Criterion | Result |
| --- | --- |
| Cell-by-cell format identity with the template | 0 deviations across 38 sheets |
| Error count after recalc | 139 against 142 for the empty template — **0 new, 3 fewer** |
| Balance sheet identity | −0.01 to −0.04 in all periods |
| Net assets + equity + result | −0.01 to −0.04 in all periods |
| Equity roll forward | closes; row 230 pulls from Lead P&L again |
| Lead NA check row | 0.00 in FY2023, −268.38 T AUD from FY2024 (unchanged finding) |
| Every P&L line item finds a template row | yes, all 20 |

The three `#DIV/0!` in `PL_Top customer` have disappeared: they needed revenue
as a denominator, and now there is revenue.

## Two limits worth naming

**The mastersheet holds 400 rows.** 434 accounts came in; 242 carry a balance
in at least one period and 192 do not. The nil accounts are listed on their own
sheet and left out of the mastersheet rather than extending the template — the
house rule is that the template is extended centrally, not per mandate. If
those 192 are wanted in the workbook, that is a template change.

**Eight account slots per line item are not enough for this chart.** `Other
operating expenses` alone gathers 36 accounts, `Personnel expenses` 18,
`Cost of materials` 17. The slots take the largest accounts, so 1.29 million of
FY2026 sits without visible account detail — most of it in the four line items
created from dummy rows, which have no slots at all by construction. Both are
findings for the template, unchanged from run 1.

## Still open

The FY2024 movement in equity of 268,375.20 that neither capital nor result
explains — a prior-year adjustment booked straight against retained earnings.
It stays in the check row until someone explains it, and it is not distributed
into a line called "profit distribution".
