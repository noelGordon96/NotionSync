# Notion Excel Sync

A standalone AutoHotkey **v2** GUI utility for reconciling an open local Excel worksheet against
a Notion database — pull data down from Notion into the sheet, or push a column up to Notion.

This utility was split out of the [AHK Utility Library](https://github.com/noelGordon96/AHK_Utility_Lib)
and bundles local copies of that library's Notion/Excel layers in `lib\` (see **Dependencies**
below).

## What it does

Pick a Notion database from the list, point it at the Excel window to work on, then use:

- **Pull Database** — download the latest copy of the Notion database and save it as the local
  copy. This is the only button that talks to Notion when reading.
- **Pull Page ID** — match sheet rows to Notion pages and write each page's ID into a column.
- **Pull Property** — copy one Notion property down a column, using the page IDs in the sheet.
- **Push Column** — push one Excel column **up** to a Notion property (text, URL/email/phone,
  multi-select, relation, date-only), with overwrite protection and a Replace/Append choice. This **writes
  to Notion**, and can optionally create new pages.

Pull Page ID and Pull Property only ever read from the **local copy** (click Pull Database to
refresh it); the bottom status bar shows when that copy was last updated.

### Pushing dates (date-only)

Choose a Notion **date** property in Push Column and select **Replace**. Use normal Excel
calendar dates (including date formulas) or text in exactly `YYYY-MM-DD` format. Numeric
cells are interpreted as Excel date serials using the workbook's 1900 or 1904 date system.
Text such as `10/23/2026` must first be converted to an Excel date or ISO text; the app does
not guess locale-specific formats. Values containing times, fractional-day serials, date
ranges, invalid dates, and Excel errors are rejected and reported as failed rows, not written.
A midnight Excel date has no fractional day and is accepted; the output never contains time.

Blank source cells are skipped and never clear an existing Notion date. Existing dates stay
unchanged unless **Overwrite existing data** is checked. Overwriting replaces the entire
Notion date value, including any previous time or range end, with one date. Dates do not support **Append**;
the dialog requires Replace. New pages still require `NEW_PAGE` in the ID cell and **Create
missing pages** checked; their date is validated before creating the page and writing its ID
back. Blank IDs are still skipped, and the push still ends at the last populated ID cell.
This conversion applies to the GUI's column-push sweep, not the separate single-row helper.

## Running it

1. Open the Excel workbook you want to work on in **local desktop Excel** (online Microsoft 365
   sessions won't work — this reads Excel's COM object directly).
2. Run `NotionSync_main.ahk` with AutoHotkey v2.
3. The first time **Pull Database** needs it, a dialog asks for your Notion integration secret
   (starts with `ntn_` or `secret_`). It is saved in `NotionSync.ini` next to the script, in
   **plain text**.

## Files

- `NotionSync_main.ahk` — the GUI application (run this one).
- `NotionExcelSync.ahk` — the Excel⇄Notion bridge library it includes.
- `lib\` — local (vendored) copies of the AHK Utility Library layers this project depends on
  (`ExcelCom.ahk`, `NotionAPI.ahk`, `NotionCom.ahk`, `UtilityWindows.ahk`). Committed so the
  project is self-contained; **treat as read-only here** — update them upstream in the library,
  not in this repo (see **Dependencies**).
- `NotionSync.ini` — *(created at runtime)* settings + saved API key + known databases.
  **Git-ignored — contains the API key in plain text.**
- `databases\` — *(created at runtime)* the saved local copy of each pulled database (real
  data). **Git-ignored.**

## Dependencies

The Notion and Excel layers originate in the
[AHK Utility Library](https://github.com/noelGordon96/AHK_Utility_Lib), but a **local copy** of
each now lives in this repo's `lib\` folder, so the utility carries all of its dependencies with
it and runs wherever it is cloned:

- `lib\NotionCom.ahk` (friendly Notion layer) → `lib\NotionAPI.ahk` (raw API) + `lib\UtilityWindows.ahk`
- `lib\ExcelCom.ahk` (Excel COM layer)

`NotionExcelSync.ahk` includes these by a **repo-relative path** (`#Include lib\...`) — there is
no longer a hard-coded absolute path into the library's install location.

**These four files are vendored copies of stable library files — do not edit them here.** If one
needs a change, an addition, or a bug fix, make it **upstream in the AHK Utility Library** and
refresh the local copy; this repo is not the place to diverge from the library. (A future
enhancement may pull library updates in automatically instead of keeping a manual local copy.)

## Requirements

- Windows with [AutoHotkey v2](https://www.autohotkey.com/)
- A local installation of Microsoft Excel
- A Notion integration secret with access to the target database(s)