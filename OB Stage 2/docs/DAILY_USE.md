# Entering new data and refreshing the dashboard

---

## First: getting the folder onto your computer

The `OB Stage 2` folder does not exist on your PC yet. It was built on a cloud
machine and pushed to your GitHub repository, so it lives there until you
download it. That is why you cannot find it in Explorer.

**The quick way — download a ZIP**

1. Open <https://github.com/sraj21863-wq/yuvraj>
2. Click the **branch** button (it says `main`) and pick
   **`claude/live-dashboard-order-booking-nddjvm`**
3. Click the green **Code** button, then **Download ZIP**
4. Right-click the ZIP in your Downloads folder and choose **Extract All**
5. Inside the extracted folder you will find **`OB Stage 2`**. Move it wherever
   you keep your work — your Documents folder is fine.

> Extract the ZIP before opening anything. Windows lets you look inside a ZIP as
> if it were a folder, but Excel and Python cannot write to files in there, so
> the dashboard will fail in confusing ways if you skip this step.

**Or, if you have Git installed**

```
git clone https://github.com/sraj21863-wq/yuvraj.git
cd yuvraj
git checkout claude/live-dashboard-order-booking-nddjvm
```

Either way you end up with the same thing:

```
OB Stage 2\
├── Order_Booking_Master_STAGE2.xlsx    <- the workbook you paste into
├── dashboard\                          <- the page
├── pipeline\                           <- run_service.cmd lives here
└── docs\                               <- this file
```

**You also need Python**, once. Get it from <https://www.python.org/downloads/>
and — this matters — tick **"Add python.exe to PATH"** on the first screen of
the installer. Everything else installs itself the first time you run
`run_service.cmd`.

---

## Once, at the start of the day

Double-click **`OB Stage 2\pipeline\run_service.cmd`** and leave the black
window open. It opens the dashboard at <http://127.0.0.1:8787/>.

That window *is* the connection between Excel and the dashboard. Close it and
the page goes to **SNAPSHOT — not live** on the next refresh. That is the page
telling the truth, not a fault.

You do not need VS Code Live Server, and you should not use it here. The service
serves the page itself, so there is only one address to remember.

---

## The short loop — every time new orders come in

### 1. Open the workbook

`OB Stage 2\Order_Booking_Master_STAGE2.xlsx`

### 2. Paste the export into `ASO_Raw`

Go to the **`ASO_Raw`** sheet and click cell **`A2`**.

Paste the export **without its header row** — row 1 already holds the headings.
Paste **values only** (Ctrl+Alt+V → Values) so you don't bring the export's
formatting in with it.

The columns must land in this order, which is the order the export already
comes in:

| | | | |
|---|---|---|---|
| A Release picking | B Customer account | C Commission group | D Sales group |
| E PS number | F Quotation | G Sales order | H Name |
| I Estimated amount | J Created date and time | K Delivery release | L Release date |
| M Requested ship date | N Confirmed ship date | O Delivery release by | P Released by |
| Q Order type | R Invoice account | S Object | T Commission group2 |
| U Order type2 | V Status | W Currency | X Project ID |
| Y Release status | Z Quotation2 | AA Do not process | |

**Stop at column AA.** Columns **AB** and **AC** hold the formulas that turn the
raw commission group and sales group into the region and segment everything else
uses. Pasting over them is the one action that breaks the workbook. If you
select whole rows and paste, you will paste over them — click `A2` and paste
there.

### 3. Refresh the other feeds if you have new ones

Same idea, same rule — paste values, don't touch formula columns:

| Sheet | Paste at | Why it matters |
|---|---|---|
| `SOP` | `A5` | the not-invoiced amount (column O of the register) and the project name |
| `Invoice data_Raw` | `A2` | the invoiced amount (column P of the register) |
| `CPS` | `A2` | which sales person owns which PS number |

If you skip these, the new orders still appear, but with **0.00** value — the
money comes from `SOP` and `Invoice data_Raw`, not from the export.

### 4. Save the workbook

**Ctrl+S.** This is the step that matters. The dashboard reads the numbers Excel
calculated and *saved*; an unsaved change exists only in Excel's memory.

### 5. Refresh the browser

**F5** on <http://127.0.0.1:8787/>.

The service notices the workbook was saved, rebuilds, and the page comes back
with the new numbers. Nothing else to run.

### 6. Check the top of the page

- The green band should say **LIVE — from Order_Booking_Master_STAGE2.xlsx** and
  a workbook-saved time that matches the save you just did.
- The amber band underneath lists anything that could not be reconciled. If a
  new name appears there, see `reconciliation.txt` and the note in the README
  about `personAliases`.

That is the whole short loop: **paste at `A2` → save → F5.**

---

## The long loop — before the next export replaces the last one

### Why this exists

The register's live rows are a **500-row window** onto `ASO_Raw`. Row 2317 of
`YTD Merg` reads `ASO_Raw` row 2, row 2318 reads row 3, and so on for 500 rows.
That window is what makes the workbook live, and it has two consequences:

- **An export longer than 500 rows will not fit.** Rows past the 500th are in
  `ASO_Raw` but nothing reads them, so they are missing from every figure on the
  dashboard without any error appearing.
- **The window only ever shows the current export.** Paste a new export over the
  old one and the previous orders stop existing — they were never anything but a
  view of `ASO_Raw`.

So before the next export goes in, the rows the last one produced have to become
ordinary values — history — and a fresh window has to be laid out underneath
them.

### Doing it

1. Save and **close** the workbook in Excel.
2. Double-click **`pipeline\roll_forward.cmd`**.
3. It shows you what it is about to do and waits:

   ```
   YTD Merg: history ends at row 2316, live block at rows 2317-2816
     486 of its 500 rows carry an order -- these become history (rows 2317-2802)
     a fresh 500-row live block goes in at rows 2803-3302, reading ASO_Raw rows 2-501
     the ASO_Raw paste zone A2:AA is cleared, ready for the next export

   Go ahead? (y/N)
   ```

4. Type `y`. It keeps a timestamped copy of the workbook in
   `OB Stage 2\backups\` before changing anything.

Now paste the next export at `ASO_Raw!A2` and carry on with the short loop.

The numbers on the dashboard do not move when you do this — freezing changes how
the rows are stored, not what they say. That is checked: every aggregate the
pipeline produces is identical before and after.

### When to run it

Whenever you are about to paste an export that does not include the orders
already in the register — in practice, at the end of each month's cycle, and any
time an export is approaching 500 rows.

You can run `roll_forward.cmd` and answer `n` at the prompt any time you just
want to see where the window currently sits.

---

## When something looks wrong

**The band is red: "SNAPSHOT — STALE, not live"**
The service is not running. Start `run_service.cmd` and refresh. The page is
showing you an old file and saying so, which is the intended behaviour, not a
bug.

**The band says LIVE but the numbers did not change**
The workbook was not saved. Check the "workbook saved" time in the band against
the clock.

**"The page won't load at all"**
Check the address. It is `http://127.0.0.1:8787/` — not port 5500, and not a
file path. If you have an older dashboard open from another folder, close that
tab; the two look similar and only this one reads the Stage 2 workbook.

**New orders appear with 0.00 value**
`SOP` and `Invoice data_Raw` were not refreshed. The export carries the order;
those two sheets carry the money.

**New orders do not appear at all**
Either the export went in below row 501 of `ASO_Raw` (the window is full — run
the long loop), or it was pasted somewhere other than `A2`.

**A sales person shows "—" instead of a percentage**
They have no row in `Monthly Booking Target (2)` under that spelling. Their
booking still counts everywhere else on the page. `reconciliation.txt` names
them; `personAliases` in `config.json` fixes it.

**Excel had the file open and the refresh failed**
It does not fail — the service reads a copy when Excel holds the lock, and the
band then says "via working copy" so you know it happened.

---

## Running it from VS Code

You do not need VS Code for this — `run_service.cmd` does the same job with a
double-click. But if you would rather work in VS Code:

1. **File → Open Folder** and pick the `OB Stage 2` folder (the folder itself,
   not the one above it). The Explorer panel on the left should show
   `dashboard`, `docs` and `pipeline`.
2. **Terminal → New Terminal** (or Ctrl+`). It opens at the folder you just
   opened.
3. First time only, install the one library the pipeline needs:

   ```
   pip install -r "pipeline/requirements.txt"
   ```

4. Start the service:

   ```
   python "pipeline/serve.py"
   ```

5. The terminal prints:

   ```
   Order Booking Stage 2 service
     dashboard  http://127.0.0.1:8787/
     payload    http://127.0.0.1:8787/data.json
     health     http://127.0.0.1:8787/health
     watching   ...\Order_Booking_Master_STAGE2.xlsx
     Ctrl-C to stop.
   ```

   Ctrl-click that first address, or type it into your browser.
6. Leave the terminal running while you use the dashboard. **Ctrl+C** stops it.

**Do not use the Live Server extension for this page.** Live Server serves files
on port 5500 and knows nothing about the workbook, so the dashboard cannot reach
its data and falls back to the snapshot — that is exactly the "SNAPSHOT — STALE,
not live" banner. The service on port 8787 serves the page *and* the data, so
there is only one address and it is always live.

If port 8787 is already taken, change `service.port` in `pipeline\config.json`
and use the new number.

---

## Sending the dashboard to someone

Run `refresh.cmd`, then `python build_share_page.py`. It writes
`dashboard\Order_Booking_Dashboard_standalone.html` — one file, no service, no
other files needed. Mail that. It renders exactly as the live page and labels
itself **SNAPSHOT** with the age of the data, so nobody mistakes it for live.
