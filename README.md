# Channel Report Automation

Fills the 5 brand channel reports from the daily POS dumps.

## Folder layout

```
ChannelReports/
  run_reports.py
  cfpatch.py
  chartpatch.py
  rename_dumps.py
  config.json
  config_mbr.json
  dumps/      <- put the 8 downloaded dumps here (fixed names, overwrite daily)
  reports/    <- the 5 report workbooks
  logs/       <- run logs + unmatched rows
  backups/    <- automatic dated backups
```

## Dump filenames (exact)

| File | Used by |
|---|---|
| sales_current_month.xlsx | all 5 |
| sales_previous_month.xlsx | all 5 |
| sales_last_year.xlsx | all 5 |
| returns_current_month.xlsx | all 5 |
| returns_previous_month.xlsx | all 5 |
| returns_last_year.xlsx | all 5 |
| sales_two_months_ago.xlsx | Safety Sure only |
| returns_two_months_ago.xlsx | Safety Sure only |

## Report filenames (exact)

Giggles_ChannelWise_Report.xlsx, Protecto_ChannelWise_Report.xlsx,
SafetySure_ChannelWise_Report.xlsx, BeDapper_ChannelWise_Report.xlsx,
BeDesi_ChannelWise_Report.xlsx, BU_Daily_Revenue_Report.xlsx,
NonFashion_Tracker.xlsx, Fashion_Tracker.xlsx

BU_Daily_Revenue_Report is the cross-check report. It is configured with
`"all_rows": true`, so it takes the WHOLE dump with no brand filtering, and
only uses the two current-month dumps. Its totals therefore include Mad Deals,
Mega Deals and the Main Warehouse rows that the five channel reports exclude -
that difference is expected, not a fault.

## Daily workflow

1. Download the 6 (or 8) exports from GPOS - names do not matter
2. `python rename_dumps.py --source "C:\Users\<you>\Downloads"`
3. `python run_reports.py`
4. Open the reports in Excel (NOT WPS - see Notes)

`rename_dumps.py` opens each file, works out whether it is sales or returns,
reads the date column, and renames it for the period it covers - this month,
last month, the month before, or this month last year. Add `--dry-run` to see
what it would do without moving anything.

## Running only some reports

Every report has a `group` in config.json.

    daily       (8)  the 5 channel reports, both trackers, BU_DailyRevenue
    on_request  (5)  BU_Giggles, BU_Protecto, BU_SafetySure,
                     BU_BeDapper, BU_BeDesi

    python run_reports.py --group daily        the normal daily run
    python run_reports.py --group daily --no-truncate     FIRST RUN OF A MONTH
    python run_reports.py --group on_request   when management asks for them
    python run_reports.py                      everything, all 13

Only the dumps the selected reports actually read are loaded. The daily run
does not touch the stock dumps at all, so the two stock exports are only
needed for the on_request run - no need to download them every morning.

Or name them directly, comma-separated, no spaces:

    python run_reports.py --brand BU_Giggles
    python run_reports.py --brand Giggles,BU_Giggles,FashionTracker

    python run_reports.py --list               show the names and groups

A wrong name stops the run and tells you, rather than silently doing nothing.
Note Giggles and BU_Giggles are different reports.

The dumps needed depend on what you run:

    daily       8 sales/returns dumps, no stock
    on_request  6 sales/returns dumps plus the 2 stock exports

## Two config files

    config.json       the 13 daily and on-request reports
    config_mbr.json   the 5 MBR reports

Same engine, separate settings - MBR changes often, and keeping it in its own
file means an edit there cannot affect the daily reports. Point the script at
whichever one you need:

    python run_reports.py --group daily
    python run_reports.py --config config_mbr.json --group mbr

Or double-click run_daily.bat / run_mbr.bat.

## Period naming follows YESTERDAY

The reports always cover yesterday, so rename_dumps names the periods from
yesterday's month, not today's. On most days there is no difference. On the
1st there is: run on 1 October, the data is 30 September, so September is the
current month, August the previous one and July the one before that.

MBR is separate again - it covers the month that has just finished, which is
the previous calendar month whether it is run on the 1st or the 5th.

## MBR reports (monthly)

The five MBR workbooks live in reports\mbr\ rather than alongside the rest.
That is set in config.json as part of each file name, e.g.
"file": "mbr/Giggles MBR Final Rep.xlsx" - no code change is involved, so any
other group can be foldered the same way. Backups mirror the structure.


Five Monthly Business Review reports, one per brand, run after month end -
usually the first working day, before the MBR meeting.

They cover the month that has just FINISHED, so they cannot share the daily
dumps: on 3 October the daily reports need October data while MBR still needs
September. They have their own dump set, named mbr_*.

Drop the MBR downloads into incoming\mbr\ (the stock export into
incoming\mbr\stock\), then:

    python rename_dumps.py --source incoming\mbr --mbr
    python run_reports.py --group mbr

`--mbr` shifts the reference month back by one, so on 1-5 October a September
dump is named mbr_sales_month, August becomes mbr_sales_prev, July
mbr_sales_prev2 and September 2025 mbr_sales_ly. It prints the month it has
assumed, so check that line before letting it move anything.

What goes where, per report:

    Data            sales:   month, prev, prev2, last year   (stacked)
    Returns         returns: month, prev, prev2, last year   (stacked)
    Summary         summary: month, prev                     (stacked)
    Stock           stock:   as at today

"Summary" is called "Sales Summary" in the Be Dapper and Be Desi files. The
summary export is the bill-wise one (bid, Items, Tracking No), not the
bill-wise DETAIL export the other reports use; rename_dumps tells them apart
by their headers.

## Commands

```
pip install openpyxl

python run_reports.py --audit      verify layout only, changes nothing
python run_reports.py --dry-run    full run, nothing saved
python run_reports.py              normal daily run
python run_reports.py --brand Giggles
python run_reports.py --brand Giggles,BU_Giggles
python run_reports.py --group daily
python run_reports.py --group mbr
python run_reports.py --list
python run_reports.py --no-backup

python rename_dumps.py --source "<downloads folder>"
python rename_dumps.py --source "<downloads folder>" --dry-run
python rename_dumps.py --source "<downloads folder>" --keep-source
python rename_dumps.py --source "<folder>" --today 2026-10-01
```

Run `--audit` after anyone edits a report workbook.

## Brand rules

Sales are matched on the `Location` column, returns on `Warehouse`.

| Brand | Keywords |
|---|---|
| Giggles | giggles |
| Protecto | protecto |
| Safety Sure | safety sure, safetysure |
| Be Dapper | be dapper, bedapper, sultan |
| Be Desi | be desi, bedesi, lux desi, wedding studio, raja rani, raja & rani |
| Mad Deals | mad deals, maddeals (trackers only) |
| Mega Deals | mega deals, megadeals (trackers only) |
| Deal House | deal house, dealhouse (Fashion tracker only - spelling not yet seen in real data) |

Matching is case-insensitive substring. A row matching no brand, or more than
one, is written to `logs/unmatched_<date>.csv` instead of being dropped.
To add a new outlet, edit `brand_keywords` in config.json — no code change.

## If a report's layout changes

Edit that sheet's entry in config.json:

- `sheet` - sheet name
- `header_row` - row holding the dump's column headers
- `data_start_row` - first row of data
- `paste_start_col` - column number where the paste zone begins (A=1, W=23)

Then run `--audit` to confirm.

## How the dumps are used

Downloaded ONCE at the start of each month (full calendar month):

    sales_previous_month / returns_previous_month     all of last month
    sales_last_year      / returns_last_year          all of the same month last year
    sales_two_months_ago / returns_two_months_ago     Safety Sure only

Downloaded EVERY day:

    sales_current_month  / returns_current_month

Stock, downloaded on request (daily for the BU revenue reports). The two stock
exports are identical in every field - same rows, same headers, same dates -
so nothing in the file says which period it covers. They are told apart by the
folder you drop them into:

    incoming\stock_this_month\     -> stock_this_month.xlsx
    incoming\stock_last_month\     -> stock_last_month.xlsx

Exactly one file per folder. Two files, or none, and rename_dumps reports it
instead of guessing. Stock is matched on Warehouse and is NEVER truncated -
it is a snapshot, not a list of transactions.

The five channel reports need the comparison periods cut to the same point in
the month as the current data. The script does that itself: it reads the
latest date in the current-month sales dump and keeps only rows on or before
that day of the month. So if the current dump runs to the 20th, last month and
last year are both cut to their 20th. Run it on the 2nd after a holiday and it
still lines up, because the cut-off comes from the data, not from the calendar.

`"truncate_to_day": true` in config.json controls this, per sheet.

AT MONTH END this is wrong. On the first working day of a month you run the
finished month, which is complete, so there is nothing to cut back to - and
truncating would drop the 31st from a 31-day comparison month. Use
`--no-truncate` for that run only:

    1 Oct   python run_reports.py --group daily --no-truncate    (September month-end)
    2 Oct   python run_reports.py --group daily                  (back to normal)

The log says which mode ran, so it can be checked afterwards. run_daily.bat
does NOT use the flag, because it is the exception.

NonFashion_Tracker and Fashion_Tracker are different: they want FULL months,
so they have no truncation. Their DataSet and Returns sheets are rebuilt from
scratch on every run, stacking last month, then last year, then the current
month, via the `"stack"` list in config.json.

Which brands each tracker takes:

    NonFashion_Tracker   Giggles, Protecto, Safety Sure, Mad Deals, Mega Deals
    Fashion_Tracker      Be Desi (incl. Lux Desi, Wedding Studio, Raja & Rani),
                         Be Dapper (incl. Sultan, Be Dapper Wellawatte), Deal House

Together they cover every classified row exactly once. Rows from
"WH - Main Warehouse Dehiwala - *Do Not Use*" belong to neither and are listed
in unmatched_<date>.csv. Nothing has to be re-pasted by hand at month start and no row
numbers need maintaining.

If a monthly dump is missing, the affected sheet is SKIPPED with an error
rather than written with a partial block.

The pivot caches in this file were emptied and marked refresh-on-load, so
Excel rebuilds them from DataSet when the file opens. The first open after a
run takes a few seconds longer because of that.

## The five BU brand revenue reports

BU_Giggles, BU_Protecto, BU_SafetySure, BU_BeDapper, BU_BeDesi.

Giggles, Protecto and Safety Sure use this month + previous month.
Be Dapper and Be Desi add a last-year period, and their sheets are named
"Off Season"/"Off_season" for the previous month and SalesLY/SalesRepItem LY
for last year - different names, same dumps.

Returns paste 19 columns (the dump's # through Created On); the final
Created By column is empty and is not pasted. That is the `max_cols` setting.

These reports contain charts. openpyxl rewrites charts from a simplified model
and throws away their style and colour parts, so `chartpatch.py` snapshots the
original chart parts before each save and puts them back afterwards. The log
reports "charts restored (N parts)" on every run. It must stay in the same
folder as run_reports.py.

## When GPOS adds a column to a dump

Symptom: #VALUE! errors in a report, and values sitting under the wrong
headers from some point onwards.

Cause: the dump has a column the report sheet was never built for, so every
column after it lands one place to the left.

Fix: find the new column's position in the dump (1-based) and add it to that
sheet's `skip_cols` list in config.json. Already set:

    BU revenue reports, sales sheets : skip_cols [8]   - "Note"
    BU revenue reports, stock sheets : skip_cols [2]   - "Outlet"

If the column IS wanted in the report, insert it in the sheet in Excel at the
matching position - Excel moves the formulas correctly - then remove the entry
from skip_cols.

## Notes

- Array formulas (TEXTJOIN and similar, e.g. the Channels column) are handled:
  each row gets its own anchor. If a Channels column ever comes out filled in
  only the first row, that is this handling failing - check run_reports.py has
  `re` imported and the ArrayFormula branch in fill_sheet.

- ALWAYS open these reports in Excel, never WPS Office. WPS silently strips
  the extended conditional formatting (coloured data bars, icon sets) when it
  saves, and the damage is permanent. This has already happened once.

- Report files must be CLOSED when the script runs, or that brand is skipped.
- Formulas recalculate when you open the workbook, not during the run.
- Helper formulas are stretched to match each day's row count, so files stay small.
- All conditional formatting is preserved, including coloured data bars and
  icon sets. `cfpatch.py` snapshots the extended rules that openpyxl cannot
  read and splices them back after each save. It must stay in the same folder
  as run_reports.py. The log reports "conditional formatting restored on N
  sheet(s)" on every run - if you ever see a restore FAILED line, the data is
  still correct but the formatting on that report needs checking.
