# OB Stage 2

A new live order-booking dashboard, and the workbook it reads.

The Stage 2 workbook is the automated master with its data cleared and the
sample data loaded in its place. Every formula, pivot, external link, table and
piece of formatting that made the old file "live" is still there and still
wired to the same places — the workbook builder never touched them.

```
OB Stage 2/
├── Order_Booking_Master_STAGE2.xlsx   the workbook the dashboard reads
├── dashboard/
│   ├── Order_Booking_Dashboard.html   the dashboard
│   ├── dashboard-config.js            where it looks for data, and when data stops being "live"
│   ├── validator.js                   the payload schema check, run at both ends
│   └── data.js                        offline snapshot, rebuilt by the pipeline
├── pipeline/
│   ├── config.json                    every rule about the workbook's shape lives here
│   ├── ob_pipeline.py                 workbook  ->  dashboard payload
│   ├── serve.py                       the live data service
│   ├── run_service.cmd / .sh          start the service and open the dashboard
│   ├── refresh.cmd / .sh              rebuild the offline snapshot only
│   ├── roll_forward.cmd / .sh         freeze the live block, lay a fresh one under it
│   ├── build_share_page.py            fold the dashboard into one mailable file
│   ├── build_stage2_workbook.py       how the Stage 2 workbook was built (re-runnable)
│   ├── verify_stage2.py               proves the workbook's live block still works
│   └── xlsx_surgery.py                the OOXML helpers those two use
└── docs/
    ├── DAILY_USE.md                   pasting new data and refreshing — start here
    ├── PIPELINE.md                    how the workbook feeds the dashboard, end to end
    ├── build_report.txt               what the workbook build did, sheet by sheet
    └── reconciliation.txt             what the last data build could not match — read this
```

## Running it

Double-click **`pipeline\run_service.cmd`**. It installs `openpyxl` if it is
missing, starts the data service, and opens the dashboard at
<http://127.0.0.1:8787/>.

The service watches the workbook. Save the workbook in Excel and the next page
refresh shows the new numbers — nothing else to run. Leave the window open while
you use the dashboard; close it when you are done.

If you only want to refresh the offline copy (to mail the HTML to someone, or to
open it without the service running), run **`pipeline\refresh.cmd`** instead.
The page then says **SNAPSHOT** rather than **LIVE**, with the age of the data,
because it is not live and should not claim to be.

- `http://127.0.0.1:8787/health` — what the service can see right now.
- `http://127.0.0.1:8787/refresh` — force a rebuild.

macOS and Linux: use `run_service.sh` and `refresh.sh`.

**Entering new data:** paste the export at `ASO_Raw!A2`, save, refresh the page.
Full steps, including the 500-row window and when to run `roll_forward.cmd`, are
in **[docs/DAILY_USE.md](docs/DAILY_USE.md)**.

## What the dashboard will not do

It will not show a number it cannot stand behind. The payload is schema-checked
by `validator.js` when the pipeline writes it and again when the page reads it.
If the check fails, or there is no data at all, the page renders an explanation
of what to run instead of a screen of zeroes. If the service is unreachable it
falls back to the snapshot, and labels every such render as a snapshot with its
age. If the pipeline had to read a copy of the workbook because Excel had the
original locked, the page says "via working copy".

Three things it states rather than hides:

- **Booking outside the reporting grid.** EXPORT, INTERNAL, FOC, Others and
  Quality do not fit the 3 × 4 grid. Their booking is shown on its own line in
  the pivot, recalculated for whatever month and filter is selected — 15.63 Cr
  across FY 2026 at the time of writing.
- **Booked value against money.** Booked value grouped by status flag and money
  invoiced are different measures that happen to be similar sizes. The band
  under the KPIs prints both, filtered identically, so neither can be read as
  the other. In this data they point opposite ways: 253.3 Cr of booking carries
  an open status while 250.8 Cr has actually been invoiced.
- **Sales people with no target.** Names are matched between the register and
  the target sheet by word set, then by unambiguous word overlap. Whatever is
  left is reported, never guessed at.

## Read `docs/reconciliation.txt`

The pipeline writes it on every run. Two things in it need a human:

1. **Ten names book against no row in the target register** (44.5 Cr, ~12% of
   YTD booking). Their booking counts in every regional and segment figure; only
   their personal target column is blank. Two of them look like spelling
   variants rather than genuinely untargeted people — `Shahbaz, Ahmed` against
   the target sheet's `Shahbaaz`, and `J Bharadwaj, Ankitha` against `Ankita`.
   They were deliberately **not** matched automatically, because a wrong guess
   credits one person's booking to another. If they are the same person, pin
   them in `pipeline/config.json`:

   ```json
   "personAliases": {
     "Shahbaz, Ahmed": "Shahbaaz",
     "J Bharadwaj, Ankitha": "Ankita"
   }
   ```

2. **Thirteen names were matched across a spelling difference** — `P, Sharath
   Kumar` to `Kumar,Sharath`, `Uthayan B, Uthayan` to `B, Uthayan`, and so on.
   Each is listed with how it was matched. Check them once; pin any that are
   wrong with the same `personAliases` block, which always wins over the
   automatic match.

Also worth knowing: 17 orders carry no sales person at all and appear as
`(unassigned)`, and 2.69 Cr carries a status that is neither open nor invoiced
(short close, cancelled), shown separately in the order-book KPI rather than
folded into either bucket.

## Rebuilding the workbook

`Order_Booking_Master_STAGE2.xlsx` is already built. To rebuild it from the two
source files:

```
python build_stage2_workbook.py ^
  --automated "Order Booking Master  AUTOMATED.xlsx" ^
  --sample    "order booking Sample  Copy.xlsx" ^
  --out       "..\Order_Booking_Master_STAGE2.xlsx" ^
  --report    "..\docs\build_report.txt"

python verify_stage2.py ^
  --workbook "..\Order_Booking_Master_STAGE2.xlsx" ^
  --expected "order booking Sample  Copy.xlsx"
```

`verify_stage2.py` re-implements the twenty formulas of the live block and runs
them against the built workbook's own raw sheets. If the automation survived,
those formulas fed by the sample's raw export must reproduce the sample's own
register rows exactly. It reports 486 rows and 4,860 cell comparisons with no
mismatch.

What the build did to the workbook is in `docs/PIPELINE.md` and, sheet by sheet,
in `docs/build_report.txt`.
