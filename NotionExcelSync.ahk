;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	NotionExcelSync (Function Library)
; DESCRIPTION:	Excel BRIDGE: reconciles rows in a local Excel worksheet with
;				entries in a Notion database. Knows about both Excel (via
;				ExcelCom.ahk) and Notion, but talks to Notion ONLY through the
;				friendly verbs in NotionCom.ahk -- never the raw NotionAPI.ahk
;				layer. Match a sheet row to a Notion entry, then push/pull data
;				both ways.

; VERSION:		2.10.23.26
; AUTHOR:		Noel Gordon (noel.gordon@ceterapartners.com)
; SOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; CONCEPT
;   The front-end functions all act on the CURRENTLY SELECTED ROW in Excel, so
;   they are easy to drive from hotkeys while testing. There is NO pre-configured
;   column map / ID column / match column -- every call names the sheet column(s)
;   and Notion property it needs right there in the argument list. This keeps
;   each call a single, obvious unit of work, and none of them assume where (or
;   whether) you keep a page-ID column: you pass a page ID in, or you get one
;   back, and you decide what to do with it.
;
;   Excel columns are always a PLAIN COLUMN NUMBER (1 = A, 2 = B, ...), never a
;   header name -- this file does no header lookups of its own. If you think in
;   header names, resolve the number yourself once (e.g. scan row 1, or use
;   Excel's own Range.Find) and pass the result in.
;
; SETUP (in your calling script)
;   #Include NotionExcelSync.ahk      ; pulls in NotionCom.ahk (+ NotionAPI.ahk + UtilityWindows.ahk) and ExcelCom.ahk
;
;   NotionExcelSync_init("ntn_your_secret", "<database_id>")
;
;   ; --- Match the selected row on column B, get back a page ID (writes nothing) ---
;   id := NotionExcelSync_matchSelectedRow(2, "Work Email")
;
;   ; --- Pull ONE Notion property into column D of the selected row ---
;   NotionExcelSync_pullSelectedRow(id, "Department", 4)
;
;   ; --- Push column D of the selected row to ONE Notion property (existing page) ---
;   NotionExcelSync_pushSelectedRow(id, 4, "Department")
;
;   ; --- Sync several columns: just call the primitive once per column ---
;   for col, notionProp in Map(1, "Name", 4, "Department")
;       NotionExcelSync_pullSelectedRow(id, notionProp, col)
;
;   ; --- Whole-column versions: start at the selected row, work down the sheet ---
;   ; (the database must already be in memory: NotionCom_pull / NotionCom_load)
;   NotionExcelSync_setExcelWindow("Employees")           ; optional: pick the Excel window
;   r := NotionExcelSync_stampPageIds(2, "Work Email", 1)  ; match col B, write IDs to col A
;   MsgBox(r.matched . " matched, " . r.unmatched . " not found")
;   NotionExcelSync_pullProperty(1, "Department", 4)       ; IDs in col A -> Department into col D
;
; NOTES
;   - Matching, the database schema, and the row cache are all owned by
;     NotionCom.ahk (case-insensitive matching, cached until you call
;     NotionExcelSync_refresh() / NotionCom_refresh()). This file just reads
;     Excel and calls those verbs -- it keeps no Notion cache of its own.
;   - NotionExcelSync_matchSelectedRow does NOT write the ID anywhere -- it just
;     returns it. If more than one entry matches, the first is used and a
;     warning is shown.
;   - NotionExcelSync_pushSelectedRow only UPDATES an existing page (it requires
;     a page ID). To create a brand-new Notion entry from a row, build a
;     Map("<Notion property>" -> value) from whichever cells you need and call
;     NotionCom_createEntry(...) directly -- that's a pure-Notion operation with
;     no Excel-column bookkeeping, so it doesn't need a bridge wrapper.
;   - Read-only Notion property types (formula, rollup, created/last-edited time
;     and by, unique_id, files) are silently skipped by a push.


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2

; Dependencies (auto-included; #Include is only processed once even if your
; calling script also includes them). This utility was split out of the AHK
; Utility Library; local (vendored) copies of NotionCom.ahk and ExcelCom.ahk
; (plus NotionAPI.ahk and UtilityWindows.ahk) live in this repo's lib\ folder
; and are included by repo-relative path. Do not edit those copies here --
; change them upstream in the library and refresh the local copy.
#Include lib\ExcelCom.ahk
#Include lib\NotionCom.ahk




; ##########################################################
; 	PUBLIC CONFIGURATION FUNCTIONS
; ##########################################################


; Configure the sync for a given database. Call once before the other
; functions (typically in your script's auto-execute section). A thin
; pass-through to NotionCom_pull -- it pulls the database live and saves a copy
; to disk (see NotionCom.ahk). To start from the saved copy instead, call
; NotionCom_load(databaseId) yourself.
;
; token      : Notion integration secret.
; databaseId : the target Notion database ID (hyphens optional).
NotionExcelSync_init(token, databaseId) {
	NotionCom_pull(token, databaseId)
}


; Re-pull the database live and re-save its file (a pass-through to
; NotionCom_refresh). Call after you change data in Notion so the next
; match/pull sees it.
NotionExcelSync_refresh() {
	NotionCom_refresh()
}


; Choose which Excel window this bridge works against. Every function here acts
; on the active workbook/sheet of whichever Excel window this names. Leave
; windowName blank (the default) to use the top-most Excel window.
;
; windowName : the start of the Excel window's title, typically the workbook's
;              name (e.g. "Employees" matches "Employees.xlsx - Excel"). AHK's
;              usual title matching applies, so a calling script that used
;              SetTitleMatchMode(2) can give any part of the title instead.
;              "ahk_class XLMAIN" is added for you so only Excel main windows
;              can match.
NotionExcelSync_setExcelWindow(windowName := "") {
	cfg := _NotionExcelSync_cfg()
	name := Trim(windowName)
	cfg.winTitle := (name = "" ? "" : name . " ") . "ahk_class XLMAIN"
}


; The row number currently selected in the chosen Excel window. Doubles as a
; quick "can we reach Excel?" check -- it throws a clear error if the window
; can't be found.
NotionExcelSync_getSelectedRow() {
	return _NotionExcelSync_ctx().row
}




; ##########################################################
; 	PUBLIC SYNC FUNCTIONS (act on the selected Excel row)
; ##########################################################


; Match the selected row to a Notion entry. Reads column <excelMatchCol> from
; the selected row and looks for a Notion page whose <notionMatchProp> equals it
; (case-insensitively). Does NOT write anywhere -- the result is yours to do
; with as you like (stamp it into a cell, pass it straight to
; NotionExcelSync_pullSelectedRow/_pushSelectedRow, etc.).
;
; excelMatchCol  : the sheet column NUMBER (1 = A, 2 = B, ...) to read the match
;                  value from.
; notionMatchProp: the Notion property to match that value against.
;
; Returns the matched page ID (String), or "" if no match / on error.
NotionExcelSync_matchSelectedRow(excelMatchCol, notionMatchProp) {
	try {
		_NotionExcelSync_status("Reading the selected row...")

		ctx := _NotionExcelSync_ctx()
		matchValue := ctx.sheet.Cells(ctx.row, excelMatchCol).Value

		if (Trim(matchValue . "") = "") {
			MsgBox("The match cell (column " . excelMatchCol . ") on the selected row is empty.", "NotionExcelSync", "Icon!")
			return ""
		}

		; First call of the session pulls + caches the whole database inside
		; NotionCom (which shows its own "Pulling..." / "Pull Complete" windows
		; on that first, slower call). Every call after that is instant and
		; silent on NotionCom's side.
		_NotionExcelSync_status("Matching against Notion...")
		ids := NotionCom_matchIds(notionMatchProp, matchValue)

		if (ids.Length = 0) {
			; MsgBox("No Notion entry found where " . notionMatchProp . " = '" . matchValue . "'.", "NotionExcelSync", "Icon!")
			return ""
		}
		if (ids.Length > 1)
			MsgBox("Found " . ids.Length . " entries where " . notionMatchProp . " = '" . matchValue . "'. Using the first one.", "NotionExcelSync", "Icon!")

		return ids[1]
	} catch as err {
		_NotionExcelSync_showError("NotionExcelSync_matchSelectedRow()", err)
		return ""
	} finally {
		_NotionExcelSync_statusClear()
	}
}


; Pull ONE property FROM Notion INTO ONE cell of the selected row.
;
; notionPageID   : the Notion page to read (e.g. from NotionExcelSync_matchSelectedRow()).
; notionPullProp : the Notion property to read.
; excelWriteCol  : the sheet column NUMBER (1 = A, 2 = B, ...) to write the value into.
;
; Returns the value that was written, or "" on error.
NotionExcelSync_pullPropertyForSelectedRow(notionPageID, notionPullProp, excelWriteCol) {
	try {
		_NotionExcelSync_status("Reading the selected row...")
		ctx := _NotionExcelSync_ctx()

		_NotionExcelSync_status("Pulling this property from Notion...")
		val := _NotionExcelSync_toCell(NotionCom_getProperty(notionPageID, notionPullProp))
		ctx.sheet.Cells(ctx.row, excelWriteCol).Value := val
		return val
	} catch as err {
		_NotionExcelSync_showError("NotionExcelSync_pullPropertyForSelectedRow()", err)
		return ""
	} finally {
		_NotionExcelSync_statusClear()
	}
}


; Push ONE cell of the selected row TO ONE Notion property on an EXISTING page.
; Never creates a page -- see the NOTES above for creating a new entry.
;
; notionPageID   : the Notion page to update (e.g. from NotionExcelSync_matchSelectedRow()).
; excelReadCol   : the sheet column NUMBER (1 = A, 2 = B, ...) to read the value from.
; notionPushProp : the Notion property to write.
;
; Returns notionPageID on success, or "" on error / if the property isn't writable
; (not in the schema, or a read-only type).
NotionExcelSync_pushSelectedRow(notionPageID, excelReadCol, notionPushProp) {
	try {
		_NotionExcelSync_status("Reading the selected row...")
		ctx := _NotionExcelSync_ctx()
		cellVal := ctx.sheet.Cells(ctx.row, excelReadCol).Value

		_NotionExcelSync_status("Updating the Notion entry...")
		if (NotionCom_setProperty(notionPageID, notionPushProp, cellVal) = "") {
			MsgBox("'" . notionPushProp . "' isn't a writable Notion property on this database (not in the schema, or read-only).", "NotionExcelSync", "Icon!")
			return ""
		}
		return notionPageID
	} catch as err {
		_NotionExcelSync_showError("NotionExcelSync_pushSelectedRow()", err)
		return ""
	} finally {
		_NotionExcelSync_statusClear()
	}
}




; ##########################################################
; 	PUBLIC SWEEP FUNCTIONS (start at a row, work down the sheet)
; ##########################################################


; These do the selected-row jobs above for MANY rows in one call: they start at
; the currently selected row (or the row you pass) and work down to the last row
; that has data in the column that drives them. They never stop to show a
; message for an individual row -- a row that can't be handled is simply
; counted and left alone -- and they hand back one summary object at the end.
; Unlike the selected-row functions they THROW on a real problem (Excel not
; found, unknown Notion property, nothing pulled/loaded) so the caller decides
; how to report it.
;
; The database must already be in memory (NotionCom_pull / NotionCom_load).


; Fill in Notion page IDs down a column. For each row, reads <excelMatchCol>,
; finds the Notion page whose <notionMatchProp> equals it (case-insensitively),
; and writes that page's ID into <excelIdCol>. A row with no match is left
; completely unchanged; a row with a blank match cell is skipped. If more than
; one page matches, the first is used and the row is counted as a duplicate.
;
; excelMatchCol  : sheet column NUMBER (1 = A, 2 = B, ...) holding the match values.
;                  The sweep stops at the last row with data in THIS column.
; notionMatchProp: the Notion property to match those values against.
; excelIdCol     : sheet column NUMBER to write the page IDs into.
; startRow       : (optional) first row to process. Default 0 = the selected row.
; endRow         : (optional) last row to process. Default 0 = the last row with
;                  data in the driving column. A given endRow is capped at that.
;
; Returns { firstRow, lastRow, matched, unmatched, blank, duplicates }.
NotionExcelSync_stampPageIds(excelMatchCol, notionMatchProp, excelIdCol, startRow := 0, endRow := 0) {
	_NotionExcelSync_requireProperty(notionMatchProp)
	ctx := _NotionExcelSync_ctx()
	firstRow := (startRow > 0) ? startRow : ctx.row
	lastRow := _NotionExcelSync_lastRowInRange(ctx.sheet, excelMatchCol, endRow)

	result := { firstRow: firstRow, lastRow: lastRow, matched: 0, unmatched: 0, blank: 0, duplicates: 0 }
	loop (lastRow - firstRow + 1) {
		row := firstRow + A_Index - 1
		matchValue := _NotionExcelSync_cellText(ctx.sheet.Cells(row, excelMatchCol))
		if (matchValue = "") {
			result.blank++
			continue
		}
		ids := NotionCom_matchIds(notionMatchProp, matchValue)
		if (ids.Length = 0) {
			result.unmatched++
			continue
		}
		_NotionExcelSync_writeCell(ctx.sheet, row, excelIdCol, ids[1])
		result.matched++
		if (ids.Length > 1)
			result.duplicates++
	}
	return result
}


; Copy one Notion property down a column. For each row, reads a Notion page ID
; from <excelIdCol> (e.g. filled in by NotionExcelSync_stampPageIds), reads
; <notionPullProp> from that page, and writes it into <excelWriteCol>. A row
; with a blank ID cell is skipped. A row whose page can't be read (say, an ID
; that isn't in the loaded database and can't be fetched live) is counted as
; failed and left unchanged. Notion is treated as the source of truth: a
; property that is empty in Notion writes an empty cell.
;
; Text that Excel would misread -- a phone number starting with "+", an ID
; with leading zeros, anything starting with "=" -- is written as text.
;
; excelIdCol     : sheet column NUMBER holding the page IDs. The sweep stops at
;                  the last row with data in THIS column.
; notionPullProp : the Notion property to read.
; excelWriteCol  : sheet column NUMBER to write the values into.
; startRow       : (optional) first row to process. Default 0 = the selected row.
; endRow         : (optional) last row to process. Default 0 = the last row with
;                  data in the driving column. A given endRow is capped at that.
;
; Returns { firstRow, lastRow, pulled, blank, failed, firstError } -- firstError
; is the message from the first failed row ("" if none failed).
NotionExcelSync_pullProperty(excelIdCol, notionPullProp, excelWriteCol, startRow := 0, endRow := 0) {
	_NotionExcelSync_requireProperty(notionPullProp)
	ctx := _NotionExcelSync_ctx()
	firstRow := (startRow > 0) ? startRow : ctx.row
	lastRow := _NotionExcelSync_lastRowInRange(ctx.sheet, excelIdCol, endRow)

	result := { firstRow: firstRow, lastRow: lastRow, pulled: 0, blank: 0, failed: 0, firstError: "" }
	loop (lastRow - firstRow + 1) {
		row := firstRow + A_Index - 1
		pageId := _NotionExcelSync_cellText(ctx.sheet.Cells(row, excelIdCol))
		if (pageId = "") {
			result.blank++
			continue
		}
		try {
			value := _NotionExcelSync_toCell(NotionCom_getProperty(pageId, notionPullProp))
			_NotionExcelSync_writeCell(ctx.sheet, row, excelWriteCol, value)
			result.pulled++
		} catch as err {
			result.failed++
			if (result.firstError = "")
				result.firstError := err.Message
		}
	}
	return result
}


; Push one Excel column UP to a Notion property, down the sheet. For each row,
; reads a Notion page ID from <excelIdCol> and the value from <excelSourceCol>,
; then writes that value to <notionPushProp> on the page. A row with a blank ID
; is skipped; a row with a BLANK SOURCE CELL is skipped too (an empty cell never
; clears a Notion value). A row whose page can't be read/written is counted as
; failed and left unchanged.
;
; DATE TARGETS: this sweep accepts whole-day Excel date serials (using the
; workbook's 1900/1904 date system) or valid YYYY-MM-DD text, and writes a
; single date without time or range. Times/fractional serials, ambiguous text,
; invalid dates, and Excel errors count as failed rows. Blank sources are
; skipped. Date targets require strategy "replace" and retain overwrite protection.
;
; CREATING PAGES: a row whose ID cell reads <newIdToken> (default "NEW_PAGE",
; matched case-insensitively) is a request to CREATE a new page. It's only acted
; on when <createPages> is true (otherwise the row is counted in newSkipped and
; left alone). The new page is created in the configured database with
; <notionPushProp> set to the source value (plus an empty title, when the target
; isn't itself the title, so Notion accepts it), and its new ID is written back
; into the ID cell. New pages ignore <strategy>/<overwrite> -- there's nothing to
; append to or overwrite.
;
; How an EXISTING page's value is written depends on <strategy>:
;   "replace"    : set the property to the source value. A property that ALREADY
;                  has a value is only replaced when <overwrite> is true;
;                  otherwise that row is left unchanged (counted in skippedFilled).
;   "appendText" : for text properties -- append the source to the existing text
;                  ("existing; source"), or just the source if it was empty. Always
;                  applies (append never removes data), so <overwrite> is ignored.
;   "appendList" : for multi_select -- the case-insensitive union of the existing
;                  options and the ";"-split source. Always applies, so <overwrite>
;                  is ignored.
;   "appendRelation" : for relation -- the union of the existing related page IDs
;                  and the comma-split source page IDs (deduped by ID, ignoring
;                  hyphens/case). Always applies, so <overwrite> is ignored.
;
; excelIdCol     : sheet column NUMBER holding the page IDs (the driving column).
; excelSourceCol : sheet column NUMBER to read the value to push from.
; notionPushProp : the Notion property to write.
; strategy       : "replace", "appendText", "appendList" or "appendRelation".
; overwrite      : true to let "replace" overwrite a property that already has a
;                  value. Ignored by the append strategies and by new pages.
; startRow       : (optional) first row to process. Default 0 = the selected row.
; endRow         : (optional) last row to process. Default 0 = the last row with
;                  data in the driving column. A given endRow is capped at that.
; createPages    : (optional) true to act on <newIdToken> cells by creating pages.
; newIdToken     : (optional) the ID-cell sentinel meaning "create a new page"
;                  (default "NEW_PAGE", matched case-insensitively).
;
; NOTE: writing needs a LIVE Notion connection (NotionCom_pull configures one;
; NotionCom_load alone does not). Blank-detection reads the in-memory copy.
;
; Returns { firstRow, lastRow, pushed, created, blankId, blankSource,
; skippedFilled, newSkipped, failed, firstError } -- firstError is the message
; from the first failed row.
NotionExcelSync_pushProperty(excelIdCol, excelSourceCol, notionPushProp, strategy, overwrite, startRow := 0, endRow := 0, createPages := false, newIdToken := "NEW_PAGE") {
	_NotionExcelSync_requireProperty(notionPushProp)
    isDateProperty := (NotionCom_getSchema()[notionPushProp] = "date")
    if (isDateProperty && strategy != "replace")
        throw Error("Date properties support Replace only; Append is not supported.", -1)
	ctx := _NotionExcelSync_ctx()
	firstRow := (startRow > 0) ? startRow : ctx.row
	lastRow := _NotionExcelSync_lastRowInRange(ctx.sheet, excelIdCol, endRow)

	result := { firstRow: firstRow, lastRow: lastRow, pushed: 0, created: 0, blankId: 0, blankSource: 0, skippedFilled: 0, newSkipped: 0, failed: 0, firstError: "" }
	loop (lastRow - firstRow + 1) {
		row := firstRow + A_Index - 1
		pageId := _NotionExcelSync_cellText(ctx.sheet.Cells(row, excelIdCol))
		if (pageId = "") {
			result.blankId++
			continue
		}
		isNew := (pageId = Trim(newIdToken))   ; "=" is case-insensitive in v2
		if (isNew && !createPages) {
			result.newSkipped++
			continue
		}
        ; Date targets read Value2 inside the row's error handler, not locale text.
        sourceText := isDateProperty ? "" : _NotionExcelSync_cellText(ctx.sheet.Cells(row, excelSourceCol))
        if (!isDateProperty && sourceText = "") {
            result.blankSource++
            continue
        }
        try {
            if (isDateProperty) {
                sourceText := _NotionExcelSync_dateOnlyCell(ctx.sheet.Cells(row, excelSourceCol), ctx.sheet.Parent.Date1904)
                if (sourceText = "") {
                    result.blankSource++
                    continue
                }
            }
			if (isNew) {
				newId := _NotionExcelSync_createPageForPush(notionPushProp, sourceText)
				_NotionExcelSync_writeCell(ctx.sheet, row, excelIdCol, newId)
				result.created++
			} else {
				current := NotionCom_getProperty(pageId, notionPushProp)
				isBlank := (current is Array) ? (current.Length = 0) : (Trim(current . "") = "")

				; "replace" onto a property that already has data needs the opt-in.
				if (strategy = "replace" && !isBlank && !overwrite) {
					result.skippedFilled++
					continue
				}

				if (strategy = "appendText")
					value := isBlank ? sourceText : (current . "; " . sourceText)
				else if (strategy = "appendList")
					value := _NotionExcelSync_unionOptions((current is Array) ? current : [], sourceText)
				else if (strategy = "appendRelation")
					value := _NotionExcelSync_unionRelationIds((current is Array) ? current : [], sourceText)
				else
					value := sourceText

				NotionCom_setProperty(pageId, notionPushProp, value)
				result.pushed++
			}
		} catch as err {
			result.failed++
			if (result.firstError = "")
				result.firstError := err.Message
		}
	}
	return result
}




; ##########################################################
; 	PRIVATE UTILITY FUNCTIONS
; ##########################################################


; Internal: shared settings (persist across calls). Just which Excel window to use.
_NotionExcelSync_cfg() {
	static cfg := { winTitle: "ahk_class XLMAIN" }
	return cfg
}


; Internal: connect to live Excel and return { app, sheet, row } for the active
; workbook and the currently selected row, in the window chosen with
; NotionExcelSync_setExcelWindow (default: the top-most Excel window).
_NotionExcelSync_ctx() {
	app := _Excel_Get(_NotionExcelSync_cfg().winTitle)
	sheet := app.ActiveWorkbook.ActiveSheet
	row := app.ActiveWindow.RangeSelection.Row
	return { app: app, sheet: sheet, row: row }
}


; Internal: the last row with data in a column (walks up from the bottom of the
; sheet to the last populated cell). An empty column reports row 1.
_NotionExcelSync_lastRow(sheet, col) {
	; xlUp = -4162
	return sheet.Cells(sheet.Rows.Count, col).End(-4162).Row
}


; Internal: the last row to process in a column -- the last row with data in it,
; or <endRow> capped to that when a positive endRow is given (0 = no limit).
_NotionExcelSync_lastRowInRange(sheet, col, endRow) {
	lastData := _NotionExcelSync_lastRow(sheet, col)
	return (endRow > 0) ? Min(endRow, lastData) : lastData
}


; Internal: a cell's value as trimmed plain text, ready to compare. Blank for an
; empty cell or an error cell (#N/A etc. come back as COM objects), and a whole
; number Excel hands back as a float (12345.0) is turned back into 12345.
_NotionExcelSync_cellText(cell) {
	v := cell.Value
	if IsObject(v)
		return ""
	if (v is Float && v = Floor(v))
		v := Integer(v)
	return Trim(v . "")
}


; Internal: convert one Excel cell to a validated date-only ISO string.
; Value2 avoids Date/Currency coercion and display-width/locale dependence.
; Numeric values are Excel serials in the workbook's own date system; text
; must be YYYY-MM-DD. Reject times instead of silently discarding them.
; Sources: https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2
; https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487
; https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/wrongly-assumes-1900-is-leap-year
_NotionExcelSync_dateOnlyCell(cell, uses1904DateSystem) {
    static EXCEL_1900_BASE_TIMESTAMP := "18991231000000" ; Serial 1 is 1900-01-01.
    static EXCEL_1904_BASE_TIMESTAMP := "19040101000000" ; Serial 0 is 1904-01-01.
    static EXCEL_1900_MAXIMUM_SERIAL := 2958465          ; Last supported date: 9999-12-31.
    static EXCEL_1904_MAXIMUM_SERIAL := 2957003          ; Same last date in the 1904 system.
    static EXCEL_FICTITIOUS_LEAP_DAY := 60              ; 1900-02-29 never existed.
    static ISO_DATE_PATTERN         := "^[0-9]{4}-[0-9]{2}-[0-9]{2}$"
    static MONTH_LENGTHS            := [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    ; 1. Read the underlying value, leaving blank cells blank and rejecting errors.
    sourceValue := cell.Value2
    if IsObject(sourceValue)
        throw ValueError("Date source is an Excel error or unsupported value; correct the cell before pushing.", -1)

    ; 2. Convert a whole-day serial without any timezone or locale interpretation.
    if (sourceValue is Integer || sourceValue is Float) {
        if (sourceValue != Floor(sourceValue))
            throw ValueError("Date-only push does not accept times. Remove the time portion from the Excel value.", -1)
        minimumSerial := uses1904DateSystem ? 0 : 1
        maximumSerial := uses1904DateSystem ? EXCEL_1904_MAXIMUM_SERIAL : EXCEL_1900_MAXIMUM_SERIAL
        if (sourceValue < minimumSerial || sourceValue > maximumSerial)
            throw ValueError("Excel date serial is outside the supported workbook date range.", -1)
        if (!uses1904DateSystem && sourceValue = EXCEL_FICTITIOUS_LEAP_DAY)
            throw ValueError("Excel serial 60 represents the nonexistent date 1900-02-29; correct this cell.", -1)
        dayOffset := Integer(sourceValue)
        if (!uses1904DateSystem && dayOffset > EXCEL_FICTITIOUS_LEAP_DAY)
            dayOffset -= 1                            ; Remove Excel's fictitious leap day.
        baseTimestamp := uses1904DateSystem ? EXCEL_1904_BASE_TIMESTAMP : EXCEL_1900_BASE_TIMESTAMP
        return FormatTime(DateAdd(baseTimestamp, dayOffset, "Days"), "yyyy-MM-dd")
    }

    ; 3. Accept only explicit ISO date text; never guess month/day ordering.
    sourceText := Trim(sourceValue . "")
    if (sourceText = "")
        return ""
    if !RegExMatch(sourceText, ISO_DATE_PATTERN)
        throw ValueError("Use a date-only Excel value or YYYY-MM-DD text (no time or date range).", -1)
    yearNumber  := Integer(SubStr(sourceText, 1, 4))
    monthNumber := Integer(SubStr(sourceText, 6, 2))
    dayNumber   := Integer(SubStr(sourceText, 9, 2))
    if (yearNumber < 1 || monthNumber < 1 || monthNumber > 12)
        throw ValueError("Invalid calendar date: " . sourceText, -1)
    maximumDay := MONTH_LENGTHS[monthNumber]
    if (monthNumber = 2 && Mod(yearNumber, 4) = 0 && (Mod(yearNumber, 100) != 0 || Mod(yearNumber, 400) = 0))
        maximumDay := 29
    if (dayNumber < 1 || dayNumber > maximumDay)
        throw ValueError("Invalid calendar date: " . sourceText, -1)
    return sourceText
}


; Internal: write a value into a cell without letting Excel reinterpret text. A
; string that starts with = + - @ (Excel would treat it as a formula) or that
; is numeric (Excel would drop leading zeros, turn 1e5 into 100000, ...) gets a
; leading apostrophe, which Excel hides and stores as "this is text". Numbers
; are written as numbers, and ordinary text is written as-is.
_NotionExcelSync_writeCell(sheet, row, col, value) {
	if (value is String && value != "" && (InStr("=+-@", SubStr(value, 1, 1)) || IsNumber(value)))
		value := "'" . value
	sheet.Cells(row, col).Value := value
}


; Internal: throw a clear error unless <propName> is a property of the loaded
; database (catches a typo up front instead of failing on every row).
_NotionExcelSync_requireProperty(propName) {
	if !NotionCom_getSchema().Has(propName)
		throw Error("'" . propName . "' is not a property of this Notion database.", -1)
}


; Internal: the case-insensitive union of an Array of existing option names and
; the ";"-split parts of <sourceText>, existing options first, blanks dropped.
; Used by NotionExcelSync_pushProperty's "appendList" strategy for multi_select.
_NotionExcelSync_unionOptions(existing, sourceText) {
	out := []
	seen := Map()
	seen.CaseSense := false
	for name in existing {
		s := Trim(name . "")
		if (s != "" && !seen.Has(s)) {
			seen[s] := true
			out.Push(s)
		}
	}
	for part in StrSplit(sourceText, ";") {
		s := Trim(part)
		if (s != "" && !seen.Has(s)) {
			seen[s] := true
			out.Push(s)
		}
	}
	return out
}


; Internal: the union of an Array of existing related page IDs and the
; comma-split parts of <sourceText>, existing IDs first, blanks dropped. Deduped
; by normalized ID (hyphens removed, case-insensitive) so the same page written
; with or without hyphens isn't added twice. Used by the "appendRelation"
; strategy of NotionExcelSync_pushProperty.
_NotionExcelSync_unionRelationIds(existing, sourceText) {
	out := []
	seen := Map()
	seen.CaseSense := false
	for id in existing {
		s := Trim(id . "")
		if (s = "")
			continue
		key := StrReplace(s, "-", "")
		if (!seen.Has(key)) {
			seen[key] := true
			out.Push(s)
		}
	}
	for part in StrSplit(sourceText, ",") {
		s := Trim(part)
		if (s = "")
			continue
		key := StrReplace(s, "-", "")
		if (!seen.Has(key)) {
			seen[key] := true
			out.Push(s)
		}
	}
	return out
}


; Internal: the name of the database's title property (every database has exactly
; one), or "" if the schema somehow has none.
_NotionExcelSync_titleProperty() {
	for name, type in NotionCom_getSchema()
		if (type = "title")
			return name
	return ""
}


; Internal: create a new page in the configured database for a push, with
; <pushProp> set to <sourceText>. When <pushProp> isn't the title property, an
; empty title is included so Notion accepts the page. Returns the new page's ID.
_NotionExcelSync_createPageForPush(pushProp, sourceText) {
	props := Map()
	titleProp := _NotionExcelSync_titleProperty()
	if (titleProp != "" && titleProp != pushProp)
		props[titleProp] := ""
	props[pushProp] := sourceText
	return NotionCom_createEntry(props)
}


; Internal: convert a value returned by NotionCom_getProperty into a single
; cell-friendly value (Arrays joined with "; ", nested objects blanked).
_NotionExcelSync_toCell(val) {
	if (val is Array) {
		out := ""
		for i, e in val
			out .= (i > 1 ? "; " : "") . (IsObject(e) ? "" : e)
		return out
	}
	if IsObject(val)
		return ""
	return val
}


; Internal: show a "working" status tooltip while a slow action runs. The short
; Sleep lets the tooltip actually paint before the script blocks on the network
; (AHK is single-threaded, so nothing repaints while a request is in flight).
_NotionExcelSync_status(text) {
	ToolTip("NotionExcelSync: " . text)
	Sleep(50)
}


; Internal: clear the status tooltip. Call from a "finally" so it always clears,
; even when the action fails.
_NotionExcelSync_statusClear() {
	ToolTip()
}


; Internal: show a detailed, traceable error message. A caught error normally
; only shows its text; this also shows the exact file and line the error came
; from and the full call stack, so you can trace where it actually happened.
;
; functionName : the public function that was running (for the title line).
; err          : the caught error object.
_NotionExcelSync_showError(functionName, err) {
	; Guard: if something other than an Error object was thrown, just show it.
	if !(err is Error) {
		MsgBox(functionName . " failed:`n`n" . err, "NotionExcelSync - Error", "Iconx")
		return
	}

	message := functionName . " failed:`n`n" . err.Message

	; "Extra" holds the offending value/name on many built-in errors -- for the
	; "local variable has not been assigned a value" error it is the variable name.
	if (err.Extra != "")
		message .= "`n`nDetails: " . err.Extra

	message .= "`n`nSource: " . err.File . "  (line " . err.Line . ")"

	if (err.What != "")
		message .= "`nThrown by: " . err.What

	if (err.Stack != "")
		message .= "`n`nCall stack:`n" . err.Stack

	MsgBox(message, "NotionExcelSync - Error", "Iconx")
}