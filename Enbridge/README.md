# Financial Deal Spreadsheet — Distribution Macro

A macro-enabled Excel workbook (`Financial Deal Spreadsheet MACRO ENABLE.xlsm`)
for logging financial trades on one sheet and automatically **distributing each
trade into a per-month tab** for every month its term spans.

You type trades into a single **Data Entry** sheet, run one macro, and the macro
fans each trade out to the right monthly sheets (`Jan 26`, `Feb 26`, …), creating
and formatting those sheets as needed, then clears Data Entry for the next batch.

> The macro source is `DistributeDataToMonths` plus three helpers
> (`UpdateDataValidation`, `GetMonthIndex`, `SheetExists`), stored in the
> workbook's VBA project. See also **How to Use and Maintain the Distribution
> Macro** (`.docx` / `.pdf`) in this folder.

---

## The columns (A–G)

Trades are entered on the **Data Entry** sheet under these headers (the macro
re-writes and bolds them every run, so they can't drift):

| Col | Header | Notes |
|---|---|---|
| A | Trade Date | formatted `d/m/yyyy` |
| B | Deal number | |
| C | Volume | formatted `#,##0` |
| D | Differential | currency, negatives in parens: `$#,##0.00#;$(#,##0.00#)` |
| E | Term | **drives distribution** — see below |
| F | Trade Type | drop-down (self-learning list) |
| G | Book Strategy | drop-down (self-learning list) |

### The **Term** column drives everything

The macro reads column **E (Term)** to decide which monthly sheets a trade lands
on. Accepted formats:

- **Single month** — `Jan26` / `Jan 26` → the trade goes to the `Jan 26` sheet.
- **Range** — `Jan26 - Apr26` → the trade is copied to **every** month in the
  span: `Jan 26`, `Feb 26`, `Mar 26`, `Apr 26`.

Format rules the parser expects:
- Month = first 3 letters (`Jan`, `Feb`, … case-insensitive).
- Year = last 2 digits (`26` = 2026).
- A range is two such tokens separated by a hyphen `-`.
- Anything under 5 characters, or with an unrecognized month, is skipped.

---

## What the macro does (`DistributeDataToMonths`)

1. **Ensures the lists database.** Creates a hidden **System_Lists** sheet (if
   missing) seeded with starter Trade Types (`ARV`, `HTT`, `WTT`) and Book
   Strategies (`FSP Initiative Sales`, `Midcon - Sales`, `TM - Sale`,
   `TM - Time Trade`).
2. **Enforces headers** on Data Entry (row 1, bold).
3. **Processes every trade row** (rows 2 → last used row in column E):
   - **Self-learning drop-downs** — any new Trade Type or Book Strategy it sees
     is appended to System_Lists so it shows up in the drop-down next time.
   - **Parses the Term**, expands ranges into individual months.
   - For each month: **creates the monthly sheet** if needed (with headers),
     appends the trade row, (re)builds an Excel **table** (`TableStyleMedium2`),
     applies number formats, sets column widths (A:F = 20, G = 25), and **freezes
     the header row**.
4. **Refreshes the drop-down validation** on Data Entry columns F and G.
5. **Clears Data Entry** (rows 2 down) *only if* at least one trade was
   distributed, leaving formatting in place for the next batch.

If no rows parsed correctly, **nothing is cleared** and you get a warning — your
data is safe to fix and re-run.

---

## Sheets in the workbook

| Sheet | Role |
|---|---|
| **Data Entry** | Where you type trades. Cleared after a successful run. |
| **System_Lists** | Hidden. Self-growing database of Trade Types (col A) and Book Strategies (col B) that feed the drop-downs. |
| **`Mmm YY`** (e.g. `Jan 26`) | Auto-created per-month output tabs. One row per trade-month, formatted as an Excel table with a frozen header. |

---

## How to use

1. Open the `.xlsm` and **enable macros / content** when prompted.
2. On **Data Entry**, fill in one row per trade (A–G). Use the **Term** column
   (E) in `Mmm YY` or `Mmm YY - Mmm YY` form.
3. Run the macro: **Developer ▸ Macros ▸ `DistributeDataToMonths` ▸ Run**
   (or assign it to a button).
4. On success you'll see *"Data distributed, columns sized, and headers pinned"*
   and Data Entry will be cleared. Check the monthly tabs for the new rows.

### Tips & gotchas

- **Term format is strict** — `Jan26` or `Jan 26` works; `January 2026`,
  `01/26`, or `1-26` will be skipped. A skipped row triggers the "no matching
  dates" warning and **no clearing happens**.
- **Two-digit years.** `26` → 2026. Years are stored absolutely
  (`year*12 + month`), so cross-year ranges like `Nov 25 - Feb 26` work.
- **Drop-downs grow automatically.** To pre-seed a Trade Type or Book Strategy,
  unhide **System_Lists** and add it to column A or B. To remove a stale option,
  delete it there.
- **Don't rename Data Entry or System_Lists** — the macro looks them up by exact
  name.
- **Validation covers rows 2–5000** on columns F and G.

---

## Files in this folder

- `Financial Deal Spreadsheet MACRO ENABLE.xlsm` — the workbook + macro.
- `How to Use and Maintain the Distribution Macro.docx` / `.pdf` — the longer
  user/maintenance guide.
