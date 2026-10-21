# Notion Excel Sync

A standalone AutoHotkey **v2** GUI utility for reconciling an open local Excel worksheet against
a Notion database — pull data down from Notion into the sheet, or push a column up to Notion.

This utility was split out of the [AHK Utility Library](https://github.com/noelGordon96/AHK_Utility_Lib)
and still depends on that library's Notion/Excel layers (see **Dependencies** below).

## What it does

Pick a Notion database from the list, point it at the Excel window to work on, then use:

- **Pull Database** — download the latest copy of the Notion database and save it as the local
  copy. This is the only button that talks to Notion when reading.
- **Pull Page ID** — match sheet rows to Notion pages and write each page's ID into a column.
- **Pull Property** — copy one Notion property down a column, using the page IDs in the sheet.
- **Push Column** — push one Excel column **up** to a Notion property (text, URL/email/phone,
  multi-select, relation), with overwrite protection and a Replace/Append choice. This **writes
  to Notion**, and can optionally create new pages.

Pull Page ID and Pull Property only ever read from the **local copy** (click Pull Database to
refresh it); the bottom status bar shows when that copy was last updated.

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
- `NotionSync.ini` — *(created at runtime)* settings + saved API key + known databases.
  **Git-ignored — contains the API key in plain text.**
- `databases\` — *(created at runtime)* the saved local copy of each pulled database (real
  data). **Git-ignored.**

## Dependencies

The Notion and Excel layers live in the **AHK Utility Library**, not in this repo:

- `NotionCom.ahk` (friendly Notion layer) → `NotionAPI.ahk` (raw API) + `UtilityWindows.ahk`
- `ExcelCom.ahk` (Excel COM layer)

`NotionExcelSync.ahk` currently references these by a **hard-coded absolute path** into the
library's install location. This is a temporary stopgap: making the utility portable/installable
on another machine (mapping those dependencies without a hard-coded path) is tracked as a
separate task and not yet done.

## Requirements

- Windows with [AutoHotkey v2](https://www.autohotkey.com/)
- A local installation of Microsoft Excel
- A Notion integration secret with access to the target database(s)
