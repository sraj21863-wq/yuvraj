# How the workbook feeds the dashboard

Everything on the dashboard comes from two sheets. The rest of the workbook
exists to keep those two correct.

```
  raw export (.xls from AX)
        │  paste into A2:AA
        ▼
  ASO_Raw ──────────────────────────────┐
   A:AA  raw export, the only paste zone │
   AB    = XLOOKUP(C, Sheet5!A:A, B:B)   │  region, normalised
   AC    = XLOOKUP(D, Sheet5!D:D, E:E)   │  segment, normalised
        │                                 │
        │  a 500-row live block reads     │
        │  ASO_Raw rows 2..501            │
        ▼                                 │
  YTD Merg                                │
   rows 3..2316    frozen history, values │
   rows 2317..2816 live formulas ─────────┘
     A..F   straight from ASO_Raw
     G      project name, parsed out of SOP column R
     H..N   customer, dates, status from ASO_Raw
     O      not-invoiced amount   INDEX/MATCH into SOP
     P      invoiced amount       SUMIFS over Invoice data_Raw
     Q      = O + P               order value at release
     R      = Q / R1              the same value in Rs Crore
     S      = month start of I    the booking month
     T      sales person          INDEX/MATCH into CPS by PS number
        │
        ▼
  ob_pipeline.py  ── reads YTD Merg + Monthly Booking Target (2)
        │
        ▼
  data.json / data.js  ── validator.js checks it at both ends
        │
        ▼
  Order_Booking_Dashboard.html
```

Supporting sheets, and what each is for:

| Sheet | Role |
|---|---|
| `Sheet5` | the mapping table. Raw commission group and sales group in, canonical region and segment out. This is configuration, not data. |
| `SOP` | open sales orders. Supplies the not-invoiced amount and the project name. |
| `Invoice data_Raw` | invoice lines. Summed per sales order into column P. |
| `CPS` | the project register. Supplies the sales person for a PS number. |
| `MTD Merg` | the current month's export, with its own lookups. Not read by the dashboard. |
| `Monthly Booking Target (2)` | one row per person × region × territory × segment, twelve monthly quotas in Rs Crore. Both the grid target and the per-person target come from here. |
| `Weekly Pivot`, `Pivot`, `Pivot Invoice`, `Sheet4` | pivot output. Their caches are set to refresh when the file opens. |

## What the Stage 2 build changed, and what it did not

The build started from the automated workbook part for part, so anything not
listed below is byte-identical to it — pivot definitions, external links,
conditional formatting, data validation, print settings, web extensions, custom
properties, the lot.

**Data cleared and reloaded from the sample.** `MTD Merg`, `CPS`, `SOP`,
`Invoice data_Raw`, `Sheet2`, `GK (2)`, `Monthly Booking Target (2)`, `PH`,
`FOC`, `CP`, `RT`, `RP`, `Sheet3`. Each sheet's cell data was replaced wholesale
with the sample's, and three things were renumbered on the way in so nothing
points at the wrong place:

- shared-string indexes, re-interned into the template's own string table;
- cell-format indexes, with the sample's formats appended to the template's
  rather than replacing them;
- external-workbook `[n]` indexes. The two files number the same three linked
  workbooks differently — the sample's `[1]` is the template's `[3]` — so every
  formula copied across was renumbered. Left alone this is silent and
  catastrophic: the formula still calculates, against the wrong workbook.

**`ASO_Raw`.** Only the A:AA paste zone was touched: cleared, then the sample's
486-row export written into rows 2–487. The AB/AC lookups stayed formulas for
all 88,462 rows.

**`YTD Merg`.** The sample's frozen history (rows 3–2316) was loaded as values.
The template's live block was then re-anchored to sit directly under it, at rows
2317–2816, still reading `ASO_Raw` rows 2–501. Re-anchoring is not a plain
row-shift: the block's references to itself had to move with it while its
references into `ASO_Raw` stayed pinned, so those two kinds of reference are
shifted separately.

**One repair.** In the template, 68 of the live block's 500 rows had been pasted
over with values and another 14 had lost their Q/R/S formulas — the block was
four-fifths live and carried the template's own leftover orders. Every row was
rebuilt from the one row that still had all twenty formulas. That both cleared
the leftover data and made the whole block live again.

**Left as the template had it.** `Weekly Pivot`, `Pivot`, `Pivot Invoice`,
`Sheet4` (pivot output) and `Sheet5` (the mapping table, which is configuration).

**Housekeeping.** `calcChain.xml` was dropped — it is an internal ordering hint
Excel rebuilds on open, and a stale one after formulas move is what makes Excel
report "unreadable content". All seven pivot caches are set to refresh on load,
`fullCalcOnLoad` is set, and the autofilter ranges and the two table ranges were
moved to the new data extents.

**Cached values.** Formula cells carry cached results, exactly as Excel saves
them, so the pipeline can read the file before anyone has opened it in Excel.
For the live block those cached values are the sample's own tail rows — which
`verify_stage2.py` proves are what these formulas produce from this raw export.

## One thing the live block improves on the sample

The sample's tail rows have a blank sales person: whoever froze that snapshot
pasted the newest orders in before column T had been filled. The live block
computes that name from `CPS`, so 449 rows that were blank in the sample now
carry an owner. `verify_stage2.py` reports these separately rather than as
matches, so the difference is visible rather than assumed.

## The payload

`ob_pipeline.py` writes one JSON object. Its keys, and the shape of each:

| Key | Shape | Meaning |
|---|---|---|
| `actual`, `target`, `count` | `"month\|region\|segment"` → number | booking, target, order count |
| `soip`, `invoicedAmount` | same key | not-invoiced and invoiced money |
| `status`, `statusCount` | `"month\|region\|segment\|flag"` | booked value and orders per status flag |
| `pactual`, `pcount` | `"month\|region\|segment\|person"` | per-person booking and orders |
| `ptargetByMonth` | `"month\|region\|segment\|person"` | the quota as the register states it |
| `ptarget` | `"region\|segment\|person"` | flat monthly average, kept for older readers |
| `excluded`, `excludedByMonth`, `excludedTotal` | | booking outside the grid |
| `warnings`, `unmatchedPeople`, `fuzzyMatchedPeople`, `peopleWithoutBooking` | | what the run could not reconcile |
| `generatedAt`, `sourceFile`, `sourceModified`, `sourceVia`, `checksum` | | provenance, shown on screen |

Months are 1–12, and every amount is in Rs Crore.

If a column moves in the workbook, change `pipeline/config.json` — the column
letters, the sheet names, the region and segment spellings, the status flags
and the person aliases all live there, not in the code.
