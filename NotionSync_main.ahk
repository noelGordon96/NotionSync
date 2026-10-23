;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	NotionSync (GUI Application)
; DESCRIPTION:	Small GUI for reconciling the open Excel worksheet against a
;				Notion database. Pick a database from the list, tell it which
;				Excel window to work on, then use:
;				  Pull Database - (beside the database list) download the
;				                  latest copy of the Notion database and save
;				                  it as the LOCAL COPY.
;				  Pull Page ID  - match rows to Notion pages (starting at the
;				                  selected row, working down) and write each
;				                  page's ID into a column. Rows with no match
;				                  are left alone.
;				  Pull Property - copy one Notion property down a column, using
;				                  the page IDs already in the sheet.
;				  Push Column   - push one Excel column UP to a Notion property (text,
;				                  URL/email/phone, or multi-select), using the page
;				                  IDs in the sheet. Has overwrite protection and a
;				                  replace/append choice. This DOES write to Notion.
;				Pull Page ID and Pull Property only ever READ from the local
;				copy, never from Notion directly -- click Pull Database
;				whenever you want the local copy refreshed. The bottom of the
;				window shows when it was last updated.
;
;				The Notion API key is asked for (in a small dialog) the first
;				time it's needed and then kept in NotionSync.ini, along with
;				the list of known databases (name + ID), the selected database
;				and the Excel window name. On startup the currently open Excel window is detected and filled into that field automatically. Every window this app opens stays on
;				top.
;
;				This is an APPLICATION, not a library: it only calls the
;				friendly layers (NotionCom.ahk and NotionExcelSync.ahk), never
;				NotionAPI.ahk directly.

; VERSION:		2.10.23.26
; AUTHOR:		Noel Gordon (noel.gordon@ceterapartners.com)
; SOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; RUNNING IT
;   Run this file with AutoHotkey v2 (it includes NotionExcelSync.ahk from this
;   folder, which in turn pulls in the Notion/Excel library layers by absolute
;   path). The Excel workbook you want to work on must be open in local desktop
;   Excel. Select the row to start from, then click a button.
;
; FILES IT CREATES (next to this script)
;   NotionSync.ini   see below. NOTE: the API key is stored here in PLAIN TEXT.
;   databases\       the local copy of each Notion database that has been
;                    pulled (real data -- keep it out of source control).
;   The .gitignore in this folder already lists both.
;
; THE INI FILE
;   [Settings]
;   ApiKey=ntn_...               asked for the first time Pull Database needs it
;                                (delete this line to be asked again)
;   ExcelWindowName=Employees    any part of the Excel window's title
;   Database=Employee Directory  the database selected in the list
;   [Databases]                  every known database: Name=ID (edit by hand to
;   Employee Directory=2687...   rename or remove one; new ones are added from
;                                the last option in the database list)
;
; WORKFLOW
;   1. Pull Database (asks for the API key the first time). This is the only
;      button that talks to Notion. It saves the local copy and loads it.
;   2. Pull Page ID / Pull Property read the local copy. The first one used in
;      a session (or after the database selection changes) loads it from disk.
;   3. Each asks for its columns/property in a small dialog, then runs. A
;      summary says how many rows were matched, skipped, or left alone.
;   The page ID header you pick in either dialog is remembered (until the app
;   closes) as the default for both.
;
; COLUMNS / HEADERS
;   The dialogs pick columns by their Excel header: choose one from the drop-down
;   (each shows its column letter, e.g. "Work Email  (col C)"). To target a column
;   that has no header yet, pick "Add new header..." at the bottom of the list --
;   it asks for a name and a column (a letter like C or a number like 3), writes
;   that header into row 1 of the sheet, and selects it.


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2
#SingleInstance Force

; Let "Excel Window Name" be any part of the window's title.
SetTitleMatchMode(2)

; These must be set BEFORE the library is included: NotionCom.ahk checks for
; databaseDir the moment it is loaded.
iniFile := A_ScriptDir . "\NotionSync.ini"
databaseDir := A_ScriptDir . "\databases"

#Include "NotionExcelSync.ahk"


; ##########################################################
; 	APPLICATION STATE
; ##########################################################


; The last option in the database list; picking it adds a new database.
addDatabaseLabel := "Add new database..."

; The last option in each header pick-list; picking it adds a new Excel header.
addHeaderLabel := "Add new header..."

; How far each spawned dialog steps down-and-right from its parent window's
; CURRENT position (screen pixels), so a child cascades off wherever its parent
; actually is at spawn time rather than re-centring on the monitor.
cascadeOffsetX := 30
cascadeOffsetY := 30

; busy      : true while an action is running (ignores extra clicks)
; databases : the known databases, in list order: an Array of { name, id }
; dbIndex   : which of them is selected (1-based), or 0 for none
; loadedKey : the database ID (hyphens removed, lowercase) whose local copy is
;             currently in memory, or "" if none
; connected : true once Pull Database has set up a live Notion connection this
;             session -- Push Column needs it (NotionCom_load alone doesn't connect)
; pageIdCol : the page ID header last chosen in EITHER dialog, offered as the
;             default in both (Pull Page ID sets it, Pull Property reads and
;             sets it)
; the rest  : the last other answers given in the dialogs, offered again next
;             time (this session only -- only the settings in the ini are
;             saved to disk)
appState := { busy: false, databases: [], dbIndex: 0, loadedKey: "", connected: false, pageIdCol: "", matchCol: "", matchProp: "", pullProp: "", writeCol: "", pushSourceCol: "", pushProp: "", pushMode: "Replace" }

; The main window's controls, filled in by NotionSync_showMainWindow().
ui := {}


; Start the app -- but only when this file is the one being run, so a test
; script can #Include it to use the functions below without opening a window.
if (A_LineFile = A_ScriptFullPath)
	NotionSync_showMainWindow()




; ##########################################################
; 	MAIN WINDOW
; ##########################################################


; Build and show the main window, filled in from the ini file.
NotionSync_showMainWindow() {
	try NotionSync_migrateOldSettings()
	appState.databases := NotionSync_loadDatabases()
	saved := NotionSync_loadSettings()

	ui.win := Gui("+AlwaysOnTop", "Notion Sync")
	ui.win.SetFont("s9", "Segoe UI")
	ui.win.MarginX := 14
	ui.win.MarginY := 12

	ui.win.Add("Text", "xm w120 h24 +0x200", "Database")
	ui.dbList := ui.win.Add("DropDownList", "x+8 yp w226 R12")
	ui.btnDatabase := ui.win.Add("Button", "x+8 yp-1 w106 h26", "Pull Database")

	ui.win.Add("Text", "xm y+8 w120 h24 +0x200", "Excel Window Name")
	; Detected open Excel window wins over the saved name; fall back to the saved
	; name when no Excel window is open.
	detectedWin := NotionSync_detectExcelWindow()
	ui.winName := ui.win.Add("Edit", "x+8 yp w340", detectedWin != "" ? detectedWin : saved.winName)

	ui.btnPageId := ui.win.Add("Button", "xm y+18 w148 h32", "Pull Page ID")
	ui.btnProperty := ui.win.Add("Button", "x+10 yp w148 h32", "Pull Property")
	ui.btnPush := ui.win.Add("Button", "x+10 yp w148 h32", "Push Column")

	ui.status := ui.win.Add("StatusBar",, "")

	; Fill the database list and select the one used last time.
	startIndex := 0
	for i, db in appState.databases
		if (db.name = saved.database)
			startIndex := i
	NotionSync_fillDatabaseList(startIndex)
	NotionSync_showIdleStatus()

	ui.dbList.OnEvent("Change", (*) => NotionSync_onDatabaseChosen())
	ui.btnDatabase.OnEvent("Click", (*) => NotionSync_run("database"))
	ui.btnPageId.OnEvent("Click", (*) => NotionSync_run("pageId"))
	ui.btnProperty.OnEvent("Click", (*) => NotionSync_run("property"))
	ui.btnPush.OnEvent("Click", (*) => NotionSync_run("push"))
	ui.win.OnEvent("Close", (*) => ExitApp())

	ui.win.Show()
}


; Show a short message in the status bar at the bottom of the main window.
NotionSync_status(text) {
	ui.status.SetText("  " . text)
}


; The status bar's resting text: when the local copy of the given database was
; last updated (or why we can't say).
NotionSync_idleStatusText(databaseId) {
	databaseId := Trim(databaseId)
	if !NotionSync_isDatabaseId(databaseId)
		return "Choose a database to see when its local copy was last updated"
	savedTime := NotionCom_savedFileTime(databaseId)
	if (savedTime = "")
		return "No local copy of this database yet -- click Pull Database"
	return "Local database last updated: " . FormatTime(savedTime, "MMM d, yyyy  h:mm tt")
}


; Put the resting text back in the status bar.
NotionSync_showIdleStatus() {
	db := NotionSync_selectedDatabase()
	NotionSync_status(NotionSync_idleStatusText(IsObject(db) ? db.id : ""))
}


; Grey out the fields and the working buttons while an action runs.
NotionSync_setBusy(isBusy) {
	appState.busy := isBusy
	for ctl in [ui.dbList, ui.winName, ui.btnDatabase, ui.btnPageId, ui.btnProperty, ui.btnPush]
		ctl.Enabled := !isBusy
}




; ##########################################################
; 	THE DATABASE LIST
; ##########################################################


; The selected database as { name, id }, or "" if none is selected.
NotionSync_selectedDatabase() {
	return (appState.dbIndex > 0) ? appState.databases[appState.dbIndex] : ""
}


; Refill the drop-down from the known databases (names, then the "Add new
; database..." option last) and select one by its position (0 = none).
NotionSync_fillDatabaseList(selectIndex) {
	items := []
	for db in appState.databases
		items.Push(db.name)
	items.Push(addDatabaseLabel)
	ui.dbList.Delete()
	ui.dbList.Add(items)
	appState.dbIndex := selectIndex
	if (selectIndex > 0)
		ui.dbList.Choose(selectIndex)
}


; The drop-down's selection changed: either a database was picked, or the last
; option ("Add new database...") was.
NotionSync_onDatabaseChosen() {
	index := ui.dbList.Value
	if (index = 0)
		return
	if (index = appState.databases.Length + 1) {
		NotionSync_addDatabase()
		return
	}
	appState.dbIndex := index
	NotionSync_showIdleStatus()
}


; Ask for a new database's name and ID, add it to the list and the ini, and
; select it. Cancelling puts the previous selection back.
NotionSync_addDatabase() {
	previous := appState.dbIndex
	answer := NotionSync_prompt("Add Database",
		"Give the database a name you'll recognize, and paste its ID (32 letters/numbers -- the part of the database's URL before the '?v=').",
		[ { key: "name", kind: "text", label: "Name", def: "" },
		  { key: "id", kind: "text", label: "Database ID", def: "" } ],
		NotionSync_checkNewDatabase)
	if !IsObject(answer) {
		NotionSync_fillDatabaseList(previous)
		return
	}
	NotionSync_saveDatabase(answer.values["name"], answer.values["id"])
	appState.databases.Push({ name: answer.values["name"], id: answer.values["id"] })
	NotionSync_fillDatabaseList(appState.databases.Length)
	NotionSync_showIdleStatus()
}


; Returns "" if a new database's name and ID (a Map of key -> value) can be
; added, otherwise a message saying why not.
NotionSync_checkNewDatabase(values) {
	name := values["name"]
	if (RegExMatch(name, "[=\[\]]") || SubStr(name, 1, 1) = ";")
		return "The name can't contain = [ ] or start with a semicolon (it's used as a key in the ini file)."
	if (name = addDatabaseLabel)
		return "Pick a different name -- that one is the list's 'add' option."
	if !NotionSync_isDatabaseId(values["id"])
		return "That doesn't look like a database ID.`n`nIt should be 32 letters/numbers (hyphens are fine) -- the part of the database's URL before the '?v='."
	for db in appState.databases {
		if (db.name = name)
			return "There's already a database called '" . name . "'."
		if (NotionSync_idKey(db.id) = NotionSync_idKey(values["id"]))
			return "That database is already in the list as '" . db.name . "'."
	}
	return ""
}


; True if the text looks like a Notion database ID: 32 letters/numbers, with or
; without hyphens.
NotionSync_isDatabaseId(text) {
	return RegExMatch(StrReplace(Trim(text), "-", ""), "i)^[0-9a-f]{32}$") ? true : false
}


; The database ID in one canonical form (no hyphens, lowercase), so the same
; database compares equal however it was typed.
NotionSync_idKey(databaseId) {
	return StrLower(StrReplace(Trim(databaseId), "-", ""))
}




; ##########################################################
; 	SETTINGS (the ini file)
; ##########################################################


; Read the saved Excel window name and the name of the database that was
; selected (blank if there's no ini yet).
NotionSync_loadSettings() {
	return {
		winName: IniRead(iniFile, "Settings", "ExcelWindowName", ""),
		database: IniRead(iniFile, "Settings", "Database", "")
	}
}


; Write the Excel window name and the selected database's name to the ini.
NotionSync_saveSettings(winName, databaseName) {
	IniWrite(winName, iniFile, "Settings", "ExcelWindowName")
	IniWrite(databaseName, iniFile, "Settings", "Database")
}


; The saved Notion API key, or "" if there isn't one.
NotionSync_apiKey() {
	return Trim(IniRead(iniFile, "Settings", "ApiKey", ""))
}


; Save the Notion API key to the ini.
NotionSync_saveApiKey(apiKey) {
	IniWrite(apiKey, iniFile, "Settings", "ApiKey")
}


; Remove the saved API key, so the next Pull Database asks for one again.
NotionSync_clearApiKey() {
	try IniDelete(iniFile, "Settings", "ApiKey")
}


; Read the known databases from the [Databases] section: an Array of { name, id }
; in file order (empty if there are none yet).
NotionSync_loadDatabases() {
	found := []
	try section := IniRead(iniFile, "Databases")
	catch
		return found
	for line in StrSplit(section, "`n", "`r") {
		equals := InStr(line, "=")
		if (equals < 2)
			continue
		name := Trim(SubStr(line, 1, equals - 1))
		id := Trim(SubStr(line, equals + 1))
		if (name != "" && id != "")
			found.Push({ name: name, id: id })
	}
	return found
}


; Add (or update) a database in the [Databases] section.
NotionSync_saveDatabase(name, id) {
	IniWrite(id, iniFile, "Databases", name)
}


; Earlier versions kept a single "DatabaseId" under [Settings]. If one is there,
; move it into the [Databases] list under a placeholder name (rename it in the
; ini if you like), select it, and remove the old key.
NotionSync_migrateOldSettings() {
	oldId := Trim(IniRead(iniFile, "Settings", "DatabaseId", ""))
	if (oldId = "")
		return
	name := ""
	known := NotionSync_loadDatabases()
	for db in known
		if (NotionSync_idKey(db.id) = NotionSync_idKey(oldId))
			name := db.name
	if (name = "") {
		base := "Database " . SubStr(NotionSync_idKey(oldId), 1, 8)
		name := base
		taken := Map()
		taken.CaseSense := false
		for db in known
			taken[db.name] := true
		n := 1
		while taken.Has(name)
			name := base . " (" . ++n . ")"
		NotionSync_saveDatabase(name, oldId)
	}
	IniWrite(name, iniFile, "Settings", "Database")
	IniDelete(iniFile, "Settings", "DatabaseId")
}


; Returns "" if the selected database (an { name, id } object, or "" for none)
; can be used, otherwise a message saying what's wrong.
NotionSync_validate(db) {
	if !IsObject(db)
		return "Choose a database first (pick 'Add new database...' at the bottom of the list to add one)."
	if !NotionSync_isDatabaseId(db.id)
		return "The ID saved for '" . db.name . "' doesn't look like a database ID.`n`nIt should be 32 letters/numbers (hyphens are fine). Fix it in the [Databases] section of NotionSync.ini."
	return ""
}




; ##########################################################
; 	ACTIONS
; ##########################################################


; Run one of the button actions ("database", "pageId", "property" or "push"). Ignores
; clicks while another action is still running, always reports a problem in a
; message box rather than letting it escape, and puts the status bar back to
; its resting text afterwards.
NotionSync_run(action) {
	if (appState.busy)
		return
	NotionSync_setBusy(true)
	try {
		NotionSync_doAction(action)
	} catch as err {
		NotionSync_reportError(err)
	} finally {
		NotionSync_setBusy(false)
		NotionSync_showIdleStatus()
	}
}


; The shared front half of every action: check a database is chosen and save
; the settings -- then either pull the database, or (for the read/push actions)
; make sure Excel can be reached and the local copy is in memory before handing
; off to the action's own dialog + sweep.
NotionSync_doAction(action) {
	db := NotionSync_selectedDatabase()
	problem := NotionSync_validate(db)
	if (problem != "") {
		NotionSync_msg(problem, "Notion Sync", "Icon!")
		return
	}
	winName := Trim(ui.winName.Value)

	try {
		NotionSync_saveSettings(winName, db.name)
	} catch {
		NotionSync_status("Couldn't save the settings to " . iniFile)
	}

	if (action = "database") {
		NotionSync_pullDatabase(db)
		return
	}

	; Push writes to Notion, which needs a live connection only Pull Database sets
	; up this session. Fail fast (before Excel/local-copy work) with a clear steer.
	if (action = "push" && !appState.connected) {
		NotionSync_msg("Push needs a live connection to Notion, which only Pull Database sets up this session.`n`nClick Pull Database first (it also refreshes the local copy so blank-detection is accurate), then try Push again.", "Notion Sync - Push Column", "Icon!")
		return
	}

	; Fail fast on an Excel problem, before touching the local copy.
	NotionSync_status("Looking for Excel...")
	NotionExcelSync_setExcelWindow(winName)
	selectedRow := NotionExcelSync_getSelectedRow()

	if !NotionSync_useLocalCopy(db)
		return

	if (action = "pageId")
		NotionSync_pullPageIds(selectedRow)
	else if (action = "property")
		NotionSync_pullProperty(selectedRow)
	else if (action = "push")
		NotionSync_pushColumn(selectedRow)
}


; "Pull Database": download the latest copy of the database from Notion. This
; saves the local copy (NotionCom does that as part of a pull) and leaves it
; loaded, so the read actions can use it straight away. The API key comes from
; the ini; if there isn't one yet a small dialog asks for it (and it's saved).
; If Notion rejects the saved key it's removed again, so the next try asks for
; a new one.
NotionSync_pullDatabase(db) {
	apiKey := NotionSync_apiKey()
	if (apiKey = "") {
		answer := NotionSync_prompt("API Key",
			"Paste your Notion integration secret (it starts with ntn_ or secret_). It's saved in NotionSync.ini next to this script, in plain text.",
			[ { key: "apiKey", kind: "secret", label: "Notion API key", def: "" } ])
		if !IsObject(answer)
			return
		apiKey := answer.values["apiKey"]
		NotionSync_saveApiKey(apiKey)
	}

	; If this fails part-way, what's in memory can't be trusted: forget it, so
	; the next read action reloads the local copy from disk.
	appState.loadedKey := ""
	NotionSync_status("Pulling the database from Notion...")
	NotionSync_keepLibraryWindowsOnTop(true)
	try {
		NotionCom_pull(apiKey, db.id)
	} catch as err {
		if InStr(err.Message, "Notion API error 401") {
			NotionSync_clearApiKey()
			err.Message := "Notion rejected the saved API key, so it has been removed -- click Pull Database again to enter a new one.`n`n" . err.Message
		}
		throw err
	} finally {
		NotionSync_keepLibraryWindowsOnTop(false)
	}
	appState.loadedKey := NotionSync_idKey(db.id)
	appState.connected := true
}


; Make sure this database's LOCAL COPY is in memory. The first read action of a
; session (or after the selected database changes) loads it from disk. There is
; never a live pull here -- if there's no local copy yet, say so and stop.
; Returns false if it couldn't be loaded.
NotionSync_useLocalCopy(db) {
	key := NotionSync_idKey(db.id)
	if (appState.loadedKey = key)
		return true

	if (NotionCom_savedFileTime(db.id) = "") {
		NotionSync_msg("There's no local copy of '" . db.name . "' yet.`n`nClick Pull Database first to download one. Pull Page ID and Pull Property read from that local copy.", "Notion Sync", "Icon!")
		return false
	}

	appState.loadedKey := ""
	NotionSync_status("Loading the local copy of the database...")
	NotionSync_keepLibraryWindowsOnTop(true)
	try {
		NotionCom_load(db.id)
	} finally {
		NotionSync_keepLibraryWindowsOnTop(false)
	}
	appState.loadedKey := key
	return true
}


; "Pull Page ID": ask which columns/property, then stamp page IDs down the sheet.
NotionSync_pullPageIds(selectedRow) {
	answer := NotionSync_prompt("Pull Page ID",
		"By default, matches the whole match column (row 2 down to the last row with a value). Tick 'Start at selected row' to instead begin at row " . selectedRow . " (the selected row) and process the number of rows you set. Rows with no match are left unchanged.",
		[ { key: "matchCol", kind: "header", label: "Excel header to match on", def: appState.matchCol },
		  { key: "matchProp", kind: "property", label: "Notion property to match on", def: appState.matchProp },
		  { key: "pageIdCol", kind: "header", label: "Header to put the page IDs in", def: appState.pageIdCol },
		  { key: "startAtSelected", kind: "checkbox", label: "Start at selected row", def: false },
		  { key: "rowCount", kind: "number", label: "Rows to process", def: "", enabledWhen: { key: "startAtSelected", checked: true } } ],
		(v) => (v["pageIdCol"] = v["matchCol"]) ? "The page ID column can't be the same as the match column -- it would overwrite the values being matched." : "")
	if !IsObject(answer)
		return

	appState.matchCol := answer.texts["matchCol"]
	appState.matchProp := answer.texts["matchProp"]
	appState.pageIdCol := answer.texts["pageIdCol"]

	NotionSync_status("Matching rows against the local copy...")
	range := NotionSync_rowRange(answer, selectedRow)
	r := NotionExcelSync_stampPageIds(answer.values["matchCol"], answer.values["matchProp"], answer.values["pageIdCol"], range.start, range.end)

	if (r.lastRow < r.firstRow) {
		NotionSync_msg("Nothing to do: no rows to process -- the first row to check (" . r.firstRow . ") is past the last row with data in the match column (" . r.lastRow . ").", "Notion Sync - Pull Page ID", "Icon!")
		return
	}

	summary := "Checked rows " . r.firstRow . " to " . r.lastRow . ".`n`n"
		. r.matched . "  matched -- page ID written`n"
		. r.unmatched . "  not matched -- left unchanged`n"
		. r.blank . "  blank match cell -- skipped"
	if (r.duplicates)
		summary .= "`n`n" . r.duplicates . " of the matched rows had more than one Notion page with that value (the first was used)."
	NotionSync_msg(summary, "Notion Sync - Pull Page ID", "Iconi")
}


; "Pull Property": ask which property/columns, then copy it down the sheet.
NotionSync_pullProperty(selectedRow) {
	answer := NotionSync_prompt("Pull Property",
		"By default, reads the whole page-ID column (row 2 down to the last ID) and copies the chosen Notion property beside each. Tick 'Start at selected row' to instead begin at row " . selectedRow . " (the selected row) and process the number of rows you set. Rows with a blank page ID are left unchanged.",
		[ { key: "pageIdCol", kind: "header", label: "Header holding the page IDs", def: appState.pageIdCol },
		  { key: "pullProp", kind: "property", label: "Notion property to pull", def: appState.pullProp },
		  { key: "writeCol", kind: "header", label: "Header to put the property in", def: appState.writeCol },
		  { key: "startAtSelected", kind: "checkbox", label: "Start at selected row", def: false },
		  { key: "rowCount", kind: "number", label: "Rows to process", def: "", enabledWhen: { key: "startAtSelected", checked: true } } ],
		(v) => (v["writeCol"] = v["pageIdCol"]) ? "The destination column can't be the page ID column -- it would overwrite the page IDs." : "")
	if !IsObject(answer)
		return

	appState.pageIdCol := answer.texts["pageIdCol"]
	appState.pullProp := answer.texts["pullProp"]
	appState.writeCol := answer.texts["writeCol"]

	NotionSync_status("Pulling '" . answer.values["pullProp"] . "' from the local copy...")
	range := NotionSync_rowRange(answer, selectedRow)
	r := NotionExcelSync_pullProperty(answer.values["pageIdCol"], answer.values["pullProp"], answer.values["writeCol"], range.start, range.end)

	if (r.lastRow < r.firstRow) {
		NotionSync_msg("Nothing to do: no rows to process -- the first row to check (" . r.firstRow . ") is past the last row with a page ID (" . r.lastRow . ").", "Notion Sync - Pull Property", "Icon!")
		return
	}

	summary := "Checked rows " . r.firstRow . " to " . r.lastRow . ".`n`n"
		. r.pulled . "  pulled`n"
		. r.blank . "  blank page ID -- skipped`n"
		. r.failed . "  couldn't be read -- left unchanged"
	if (r.failed)
		summary .= "`n`nFirst problem: " . r.firstError . "`n`nA page ID that isn't in the local copy can't be read. If those pages are new in Notion, click Pull Database to refresh the local copy and try again."
	NotionSync_msg(summary, "Notion Sync - Pull Property", "Iconi")
}


; "Push Column": ask which columns/property and options, then push a column UP
; to Notion down the sheet. Needs a live connection (checked in NotionSync_doAction).
NotionSync_pushColumn(selectedRow) {
	; The Notion property types this can push to. Others are hidden from the list.
	pushableTypes := ["title", "rich_text", "url", "email", "phone_number", "multi_select", "relation", "date"]

	answer := NotionSync_prompt("Push Column",
		"Pushes a column UP to Notion. By default it processes the whole page-ID column (row 2 down); tick 'Start at selected row' to begin at row " . selectedRow . " and limit to the rows you set. Blank source cells are skipped. 'Overwrite' lets Replace change a property that already has data; Append always adds without removing. 'Create missing pages' makes a new Notion page for any ID cell reading NEW_PAGE (writing the new ID back). The note under the property shows how to fill the source cell for the type you pick.",
		[ { key: "idCol", kind: "header", label: "Header holding the page IDs", def: appState.pageIdCol },
		  { key: "sourceCol", kind: "header", label: "Header to push from", def: appState.pushSourceCol },
		  { key: "pushProp", kind: "property", label: "Notion property to push to", def: appState.pushProp, types: pushableTypes },
		  { key: "help", kind: "help", watch: "pushProp", textFor: NotionSync_pushHelpText },
		  { key: "mode", kind: "choice", label: "When writing", def: appState.pushMode, options: ["Replace", "Append"] },
		  { key: "overwrite", kind: "checkbox", label: "Overwrite existing data", def: false, enabledWhen: { key: "mode", equals: "Replace" } },
		  { key: "createPages", kind: "checkbox", label: "Create missing pages", def: false },
		  { key: "startAtSelected", kind: "checkbox", label: "Start at selected row", def: false },
		  { key: "rowCount", kind: "number", label: "Rows to process", def: "", enabledWhen: { key: "startAtSelected", checked: true } } ],
		(v) => (v["sourceCol"] = v["idCol"]) ? "The source column can't be the same as the page ID column."
			: (NotionCom_getSchema()[v["pushProp"]] = "date" && v["mode"] = "Append") ? "Date properties support Replace only. Choose Replace; use Overwrite if the existing date should change." : "")
	if !IsObject(answer)
		return

	appState.pageIdCol := answer.texts["idCol"]
	appState.pushSourceCol := answer.texts["sourceCol"]
	appState.pushProp := answer.texts["pushProp"]
	appState.pushMode := answer.texts["mode"]

	strategy := NotionSync_pushStrategy(answer.values["pushProp"], answer.values["mode"])

	NotionSync_status("Pushing '" . answer.values["pushProp"] . "' to Notion...")
	range := NotionSync_rowRange(answer, selectedRow)
	r := NotionExcelSync_pushProperty(answer.values["idCol"], answer.values["sourceCol"], answer.values["pushProp"], strategy, answer.values["overwrite"], range.start, range.end, answer.values["createPages"])

	if (r.lastRow < r.firstRow) {
		NotionSync_msg("Nothing to do: no rows to process -- the first row to check (" . r.firstRow . ") is past the last row with a page ID (" . r.lastRow . ").", "Notion Sync - Push Column", "Icon!")
		return
	}

	summary := "Checked rows " . r.firstRow . " to " . r.lastRow . ".`n`n"
		. r.pushed . "  pushed to existing pages`n"
		. r.created . "  new pages created`n"
		. r.blankId . "  blank page ID -- skipped`n"
		. r.blankSource . "  blank source cell -- skipped`n"
		. r.skippedFilled . "  already had data -- left unchanged (Overwrite off)`n"
		. r.failed . "  couldn't be written -- left unchanged"
	if (r.newSkipped)
		summary .= "`n`n" . r.newSkipped . " row(s) marked NEW_PAGE were left alone (tick 'Create missing pages' to create them)."
	if (r.failed)
		summary .= "`n`nFirst problem: " . r.firstError
	NotionSync_msg(summary, "Notion Sync - Push Column", "Iconi")
}


; Map the target property's Notion type and the chosen Replace/Append mode to the
; push sweep's strategy. Append only applies to text (append) and multi_select
; (union of options); the single-line string types (url/email/phone) always
; replace, since appending to them would just make malformed values.
NotionSync_pushStrategy(propName, mode) {
	if (mode != "Append")
		return "replace"
	schema := NotionCom_getSchema()
	type := schema.Has(propName) ? schema[propName] : ""
	if (type = "title" || type = "rich_text")
		return "appendText"
	if (type = "multi_select")
		return "appendList"
	if (type = "relation")
		return "appendRelation"
	return "replace"
}


; The help text under the target-property dropdown in the Push Column dialog,
; keyed off the selected property's Notion type -- explains how to fill the
; source cell (and how Append/Replace behave) for that type. "" for no selection.
NotionSync_pushHelpText(propName) {
	schema := NotionCom_getSchema()
	type := schema.Has(propName) ? schema[propName] : ""
	switch type {
		case "relation":
			return "Relation: the source cell holds Notion page IDs, separated by commas (spaces are trimmed). Append adds to the pages already related; Replace sets them."
		case "multi_select":
			return "Multi-select: the source cell holds option names, separated by semicolons. Append adds to the options already set; Replace sets them."
		case "title", "rich_text":
			return "Text: the source cell is pushed as text. Append adds it after the existing text (existing; new); Replace overwrites it."
		case "date":
			return "Date only: use Excel dates or YYYY-MM-DD text. Times, ranges, and other text formats are rejected. Choose Replace; Overwrite is required to change an existing date. Blank source cells are skipped."
		case "url", "email", "phone_number":
			return "Single value: the source cell is pushed as-is. Append isn't supported for this type, so it always replaces."
		default:
			return ""
	}
}


; The row range for a sweep, from the "Start at selected row" checkbox + count.
; Unchecked (default): the whole column -- row 2 (below the header row) down to
; the last row with data (start 2, end 0 = no limit). Checked: from the selected
; row for `rowCount` rows (end 0 when the count is blank = down to the last row).
; Returns { start, end } to pass as a sweep's startRow/endRow.
NotionSync_rowRange(answer, selectedRow) {
	if (!answer.values["startAtSelected"])
		return { start: 2, end: 0 }
	count := answer.values["rowCount"]
	return { start: selectedRow, end: (count > 0) ? selectedRow + count - 1 : 0 }
}




; ##########################################################
; 	DIALOGS
; ##########################################################


; Show a small modal dialog that asks for a few values, and wait for it.
;
; title     : dialog title
; info      : explanatory text shown at the top
; fields    : an Array of { key, kind, label, def }. kind says what the field is:
;               "column"   a text box taking a column letter or number
;               "header"   a pick-list of the target Excel sheet's row-1 headers,
;                          each shown with its column letter -- "Work Email  (col
;                          C)" -- plus an "Add new header..." option that writes a
;                          new header into row 1 of the sheet and selects it
;               "property" a pick-list of the loaded database's Notion properties,
;                          each shown with its type -- "Work Email (rollup)". An
;                          optional `types` Array on the field limits it to those
;                          Notion types.
;               "choice"   a fixed pick-list; the field supplies an `options`
;                          Array of strings (def pre-selects one)
;               "checkbox" a checkbox; def (true/false) sets its initial state
;               "number"   a small integer text box (blank allowed, reads as 0)
;               "help"     a full-width, read-only note (no left label); pair it
;                          with `watch` (another field's key) and `textFor` (a
;                          function of that field's value -> the note text) to
;                          update it live as that field changes
;               "text"     a text box for any non-blank text
;               "secret"   like "text", but masked as you type
;             def is the starting value (for "header"/"property"/"choice", the
;             name to pre-select; for "checkbox", true/false). A field may also
;             carry `enabledWhen: { key, equals }` or `{ key, checked }` to grey
;             it out unless another field equals a value / a checkbox is (un)checked.
; check     : (optional) a function taking the Map of answers (key -> column
;             number, property name or text) and returning "" if fine, or a
;             message to show.
; ownerHwnd : (optional) the window to own and disable while this dialog is up.
;             Default 0 uses the main window; a nested dialog ("Add new
;             header...") passes the parent dialog's Hwnd so it's modal to it.
;
; Returns { values, texts } -- two Maps by key: values holds column numbers,
; property/header/choice names, text, and a checkbox's 1/0; texts holds what was
; chosen/typed (a name, not its "(...)" label) -- or "" if it was cancelled.
NotionSync_prompt(title, info, fields, check := "", ownerHwnd := 0) {
	oh := ownerHwnd ? ownerHwnd : ui.win.Hwnd
	dlg := Gui("+AlwaysOnTop +Owner" . oh, "Notion Sync - " . title)
	dlg.SetFont("s9", "Segoe UI")
	dlg.MarginX := 14
	dlg.MarginY := 12
	dlg.Add("Text", "w520", info)

	; The header pick-list is built once, up front, if any field needs it (there's
	; no Excel sheet to read for the API key / add-database dialogs). Property
	; pick-lists are built per field below, since each may filter to its own types.
	headerNames := []
	headerCols := []
	headerItems := []
	headerBuilt := false
	hasColumn := false
	for f in fields {
		if (f.kind = "header" && !headerBuilt) {
			NotionSync_headerChoices(&headerNames, &headerCols, &headerItems)
			headerBuilt := true
		}
		if (f.kind = "column")
			hasColumn := true
	}

	controls := Map()
	propNamesByKey := Map()
	headerFieldKeys := []
	headerPrev := Map()
	for f in fields {
		; A "help" field is a full-width note with no left label; every other kind
		; gets the usual "label on the left, control on the right" row.
		if (f.kind != "help")
			dlg.Add("Text", "xm y+12 w200 h24 +0x200", f.label)
		if (f.kind = "property") {
			names := []
			items := []
			NotionSync_propertyChoices(&names, &items, f.HasOwnProp("types") ? f.types : "")
			propNamesByKey[f.key] := names
			ctl := dlg.Add("DropDownList", "x+8 yp w310", items)
			; the list can be wider than the box, so long "name (type)" labels fit
			SendMessage(0x160, Round(470 * A_ScreenDPI / 96), 0, ctl.Hwnd)   ; CB_SETDROPPEDWIDTH
			for i, n in names
				if (n = f.def)
					ctl.Choose(i)
		} else if (f.kind = "header") {
			ctl := dlg.Add("DropDownList", "x+8 yp w310", NotionSync_headerItems(headerItems))
			SendMessage(0x160, Round(470 * A_ScreenDPI / 96), 0, ctl.Hwnd)   ; CB_SETDROPPEDWIDTH
			for i, n in headerNames
				if (n = f.def)
					ctl.Choose(i)
			ctl.OnEvent("Change", onHeaderChange.Bind(f.key))
			headerFieldKeys.Push(f.key)
			headerPrev[f.key] := ctl.Value
		} else if (f.kind = "choice") {
			ctl := dlg.Add("DropDownList", "x+8 yp w310", f.options)
			for i, opt in f.options
				if (opt = f.def)
					ctl.Choose(i)
			if (ctl.Value = 0)
				ctl.Choose(1)
		} else if (f.kind = "checkbox") {
			ctl := dlg.Add("CheckBox", "x+8 yp h24", "")
			ctl.Value := f.def ? 1 : 0
		} else if (f.kind = "number") {
			ctl := dlg.Add("Edit", "x+8 yp w80 Number", f.HasOwnProp("def") ? f.def : "")
		} else if (f.kind = "help") {
			; full-width, read-only note; content is set by the help wiring below
			ctl := dlg.Add("Text", "xm y+8 w520 h40 cGray", "")
		} else if (f.kind = "secret") {
			ctl := dlg.Add("Edit", "x+8 yp w310 Password", f.def)
		} else {
			ctl := dlg.Add("Edit", "x+8 yp w310", f.def)
		}
		controls[f.key] := ctl
	}

	; Wire up any "enabled only when another field's value/checked state holds"
	; links (a field greys out unless its condition is met -- e.g. Overwrite only
	; while Replace is chosen, the row count only while "start at selected row" is
	; ticked).
	for f in fields {
		if (f.HasOwnProp("enabledWhen") && controls.Has(f.enabledWhen.key))
			NotionSync_wireEnabledWhen(controls[f.key], controls[f.enabledWhen.key], f.enabledWhen)
	}

	; Wire up any "help" fields so their text updates from the field they watch.
	for f in fields {
		if (f.kind = "help" && f.HasOwnProp("watch") && f.HasOwnProp("textFor") && controls.Has(f.watch))
			NotionSync_wireHelp(controls[f.key], controls[f.watch], propNamesByKey.Has(f.watch) ? propNamesByKey[f.watch] : "", f.textFor)
	}

	if (hasColumn)
		dlg.Add("Text", "xm y+10 w520 cGray", "Columns can be a letter (C) or a number (3).")

	okBtn := dlg.Add("Button", "xm y+14 w100 Default", "OK")
	cancelBtn := dlg.Add("Button", "x+8 yp w100", "Cancel")
	okBtn.OnEvent("Click", submit)
	cancelBtn.OnEvent("Click", cancelIt)
	dlg.OnEvent("Close", cancelIt)
	dlg.OnEvent("Escape", cancelIt)

	state := { result: "" }
	ownerGui := GuiFromHwnd(oh)
	dialogWindowHandle := dlg.Hwnd                 ; Cache before OK/Cancel can destroy the GUI.
	try {
		; 1. Keep the owner modal while positioning and waiting for the dialog.
		if (ownerGui)
			ownerGui.Opt("+Disabled")
		NotionSync_cascadeDialog(dlg, oh)
		WinWaitClose("ahk_id " . dialogWindowHandle)
	} finally {
		; 2. Restore the owner even if positioning fails, before closing any child.
		if (ownerGui && WinExist(oh))
			ownerGui.Opt("-Disabled")
		; 3. Clean up an unfinished dialog; direct handles also find hidden GUIs.
		if WinExist(dialogWindowHandle)
			dlg.Destroy()
		; 4. Return focus only if the owner still exists.
		if WinExist(oh)
			WinActivate(oh)
	}
	return state.result


	; gui internal methods
	submit(*) {
		values := Map()
		texts := Map()
		for f in fields {
			ctl := controls[f.key]
			if (f.kind = "property") {
				if (ctl.Value = 0) {
					NotionSync_msg("Pick a Notion property for: " . f.label, "Notion Sync", "Icon! Owner" . dlg.Hwnd)
					ctl.Focus()
					return
				}
				values[f.key] := propNamesByKey[f.key][ctl.Value]
				texts[f.key] := propNamesByKey[f.key][ctl.Value]
			} else if (f.kind = "header") {
				if (ctl.Value = 0 || ctl.Value = headerNames.Length + 1) {
					NotionSync_msg("Pick an Excel header for:  " . f.label . "`n`n(Or choose '" . addHeaderLabel . "' to create one.)", "Notion Sync", "Icon! Owner" . dlg.Hwnd)
					ctl.Focus()
					return
				}
				values[f.key] := headerCols[ctl.Value]
				texts[f.key] := headerNames[ctl.Value]
			} else if (f.kind = "choice") {
				values[f.key] := f.options[ctl.Value]
				texts[f.key] := f.options[ctl.Value]
			} else if (f.kind = "checkbox") {
				values[f.key] := ctl.Value
				texts[f.key] := ctl.Value ? "1" : "0"
			} else if (f.kind = "number") {
				entered := Trim(ctl.Value)
				values[f.key] := (entered = "") ? 0 : Integer(entered)
				texts[f.key] := entered
			} else if (f.kind = "help") {
				; display-only -- nothing to collect or validate
			} else if (f.kind = "column") {
				col := NotionSync_parseColumn(ctl.Value)
				if (col = 0) {
					NotionSync_msg("'" . ctl.Value . "' isn't a valid column for:  " . f.label . "`n`nUse a letter (like C) or a number (like 3).", "Notion Sync", "Icon! Owner" . dlg.Hwnd)
					ctl.Focus()
					return
				}
				values[f.key] := col
				texts[f.key] := Trim(ctl.Value)
			} else {
				entered := Trim(ctl.Value)
				if (entered = "") {
					NotionSync_msg("Enter a value for:  " . f.label, "Notion Sync", "Icon! Owner" . dlg.Hwnd)
					ctl.Focus()
					return
				}
				values[f.key] := entered
				texts[f.key] := entered
			}
		}
		if (check != "") {
			problem := check(values)
			if (problem != "") {
				NotionSync_msg(problem, "Notion Sync", "Icon! Owner" . dlg.Hwnd)
				return
			}
		}
		state.result := { values: values, texts: texts }
		dlg.Destroy()
	}

	cancelIt(*) {
		dlg.Destroy()
	}

	; A header drop-down's selection changed. Picking the "Add new header..."
	; option runs the add flow; any real header is remembered so the selection
	; can be restored if a later add is cancelled.
	onHeaderChange(key, *) {
		ctl := controls[key]
		index := ctl.Value
		if (index = 0)
			return
		if (index = headerNames.Length + 1) {
			addHeaderFlow(key)
			return
		}
		headerPrev[key] := index
	}

	; Ask for a new header's name and column, write it into row 1 of the sheet,
	; then rebuild the header drop-downs and select the new header. Owned by this
	; dialog so it's modal to it. Cancelling restores the previous selection.
	addHeaderFlow(key) {
		ctl := controls[key]
		answer := NotionSync_prompt("Add Header",
			"Type a name for the new header and the column to put it in. It's written into row 1 of the Excel sheet and then selected here.",
			[ { key: "name", kind: "text", label: "New header name", def: "" },
			  { key: "col", kind: "column", label: "Column to put it in", def: "" } ],
			checkNewHeader, dlg.Hwnd)
		if !IsObject(answer) {
			NotionSync_chooseIndex(ctl, headerPrev.Has(key) ? headerPrev[key] : 0)
			return
		}
		try {
			NotionSync_writeHeader(answer.values["col"], answer.texts["name"])
		} catch as err {
			NotionSync_reportError(err)
			NotionSync_chooseIndex(ctl, headerPrev.Has(key) ? headerPrev[key] : 0)
			return
		}
		rebuildHeaderControls(key, answer.texts["name"])
	}

	; Returns "" if a new header's name/column can be added, else a message.
	checkNewHeader(values) {
		for name in headerNames
			if (name = values["name"])
				return "There's already a header called '" . values["name"] . "'. Pick it from the list instead of adding it again."
		for i, col in headerCols
			if (col = values["col"])
				return "Column " . NotionSync_colLetter(values["col"]) . " already has the header '" . headerNames[i] . "'. Choose a column that doesn't have a header yet."
		return ""
	}

	; Re-read the sheet's headers and refill every header drop-down in this dialog,
	; keeping each one's current selection (by name). selectKey / selectName force
	; one drop-down onto a specific header (the one just added).
	rebuildHeaderControls(selectKey := "", selectName := "") {
		want := Map()
		for k in headerFieldKeys {
			index := controls[k].Value
			want[k] := (index >= 1 && index <= headerNames.Length) ? headerNames[index] : ""
		}
		if (selectKey != "")
			want[selectKey] := selectName

		; Re-read into fresh locals, then point the shared lists at them (avoids
		; taking a VarRef of a closure-captured variable).
		newNames := [], newCols := [], newItems := []
		NotionSync_headerChoices(&newNames, &newCols, &newItems)
		headerNames := newNames
		headerCols := newCols
		headerItems := newItems
		items := NotionSync_headerItems(headerItems)

		for k in headerFieldKeys {
			ctl := controls[k]
			ctl.Delete()
			ctl.Add(items)
			chosen := 0
			for i, n in headerNames
				if (n = want[k]) {
					chosen := i
					break
				}
			NotionSync_chooseIndex(ctl, chosen)
			headerPrev[k] := chosen
		}
	}
}


; The loaded database's Notion properties for a pick-list: fills `names` with the
; property names and `items` with the labels to show ("Work Email (rollup)"),
; alphabetized, in matching order. If `types` is an Array of Notion type strings,
; only properties of those types are included (all types when "" or omitted).
NotionSync_propertyChoices(&names, &items, types := "") {
	names := []
	items := []
	for propName, propType in NotionCom_getSchema() {
		if (types is Array) {
			allowed := false
			for t in types
				if (t = propType) {
					allowed := true
					break
				}
			if (!allowed)
				continue
		}
		names.Push(propName)
		items.Push(propName . " (" . propType . ")")
	}
}


; The target Excel sheet's headers for a pick-list: fills `names` with the header
; text, `cols` with each header's column NUMBER, and `items` with the labels to
; show ("Work Email  (col C)"), left to right across the sheet, in matching order.
NotionSync_headerChoices(&names, &cols, &items) {
	names := []
	cols := []
	items := []
	for h in NotionSync_readHeaders() {
		names.Push(h.name)
		cols.Push(h.col)
		items.Push(h.name . "  (col " . NotionSync_colLetter(h.col) . ")")
	}
}


; A header pick-list's drop-down items: the header labels with the "Add new
; header..." option appended last (mirrors the database list's "add" option).
NotionSync_headerItems(headerItems) {
	items := headerItems.Clone()
	items.Push(addHeaderLabel)
	return items
}




; ##########################################################
; 	PRIVATE UTILITY FUNCTIONS
; ##########################################################


; A MsgBox that stays on top, like every other window this app opens. The
; options are the usual MsgBox ones ("Icon!", "Owner<hwnd>", ...); returns the
; button that was clicked.
NotionSync_msg(text, title := "Notion Sync", options := "") {
	; 262144 (0x40000) is MsgBox's "always on top" flag
	return MsgBox(text, title, Trim(options . " 262144"))
}


; While NotionCom pulls or loads it opens its own windows, including "Pull
; Complete" / "Load Complete" message boxes that we can't change (they're in
; the library). This keeps any window with a "NotionCom - " title on top too, by
; checking every 100 ms while the pull/load runs. (Timers still run while one of
; those message boxes is showing, so the box is caught right after it appears.)
NotionSync_keepLibraryWindowsOnTop(turnOn) {
	SetTimer(NotionSync_raiseLibraryWindows, turnOn ? 100 : 0)
}


; The timer's job: make every "NotionCom - ..." window topmost.
NotionSync_raiseLibraryWindows() {
	for hwnd in WinGetList("NotionCom - ")
		if !(WinGetExStyle("ahk_id " . hwnd) & 0x8)
			WinSetAlwaysOnTop(true, "ahk_id " . hwnd)
}


; Show <dlg> stepped down-and-right from its owner window's CURRENT position, so a
; spawned dialog cascades off wherever its parent actually is at spawn time (not
; the monitor centre). The window is sized while hidden, then moved with Win32
; SCREEN coordinates -- which WinGetPos and WinMove share -- before being revealed,
; so Gui.Show's DPI-scaled coordinates are never mixed in. Nested dialogs pass
; their parent dialog's Hwnd as the owner, so each level steps off the one above it.
NotionSync_cascadeDialog(dlg, ownerHwnd) {
	dlg.Show("Hide")
	WinGetPos(&ownerX, &ownerY, , , "ahk_id " . ownerHwnd)
	; A direct HWND finds the hidden GUI without changing DetectHiddenWindows.
	; Source: https://www.autohotkey.com/docs/v2/lib/DetectHiddenWindows.htm
	WinMove(ownerX + cascadeOffsetX, ownerY + cascadeOffsetY, , , dlg.Hwnd)
	WinShow("ahk_id " . dlg.Hwnd)
	WinActivate("ahk_id " . dlg.Hwnd)
}


; Turn what someone typed for a column into a column number: a number (3) is
; used as-is, a letter (C, AA) is converted. Returns 0 if it isn't a valid
; column (Excel has 16384 columns, A to XFD).
NotionSync_parseColumn(text) {
	text := Trim(text)
	if RegExMatch(text, "^\d+$") {
		n := Integer(text)
		return (n >= 1 && n <= 16384) ? n : 0
	}
	if RegExMatch(text, "i)^[A-Z]{1,3}$") {
		n := 0
		loop parse StrUpper(text)
			n := n * 26 + (Ord(A_LoopField) - 64)
		return (n <= 16384) ? n : 0
	}
	return 0
}


; A column number as its Excel letter(s) (1 -> A, 27 -> AA) -- the inverse of the
; letter branch of NotionSync_parseColumn. Used for the "(col C)" pick-list labels.
NotionSync_colLetter(colNumber) {
	n := colNumber
	letters := ""
	while (n > 0) {
		n -= 1
		letters := Chr(65 + Mod(n, 26)) . letters
		n := n // 26
	}
	return letters
}


; Select item <index> in a drop-down, or clear the selection when index < 1.
; (DropDownList.Choose needs a valid 1-based index; clearing goes through .Value.)
NotionSync_chooseIndex(ctl, index) {
	if (index >= 1)
		ctl.Choose(index)
	else
		ctl.Value := 0
}


; Keep <target> control enabled only while <source> satisfies <cond>, greying it
; out (and blocking input) otherwise. <cond> is either { key, equals } (source's
; selected text must equal a value) or { key, checked } (a checkbox source must
; be checked / unchecked). Uses the source's own event ("Click" for a checkbox,
; "Change" otherwise). Sets the initial state and updates on every source change.
NotionSync_wireEnabledWhen(target, source, cond) {
	toggle(*) {
		if (cond.HasOwnProp("checked"))
			on := ((source.Value != 0) = cond.checked)
		else
			on := (source.Text = cond.equals)
		target.Enabled := on
	}
	source.OnEvent((source.Type = "CheckBox") ? "Click" : "Change", toggle)
	toggle()
}


; Keep the read-only <helpCtl> text in sync with the <source> control's current
; selection: on every change it calls textFor(value) and shows the result. When
; <sourceNames> is an Array (a property field's names, by index) the selected
; NAME is passed; otherwise the control's displayed text is. Sets the initial text.
NotionSync_wireHelp(helpCtl, source, sourceNames, textFor) {
	update(*) {
		if (sourceNames is Array) {
			idx := source.Value
			value := (idx >= 1 && idx <= sourceNames.Length) ? sourceNames[idx] : ""
		} else {
			value := source.Text
		}
		helpCtl.Text := textFor(value)
	}
	source.OnEvent("Change", update)
	update()
}


; The currently open Excel window as a match fragment for the Excel Window Name
; field, or "" if no Excel window is open. When several are open the most recently
; active one wins (ExcelCom returns them in Z-order). A trailing " - Excel" is
; stripped so the field holds a clean workbook fragment.
NotionSync_detectExcelWindow() {
	detected := ExcelCom_getExcelWinName()   ; String | Array (several) | false (none)
	if (detected is Array)
		detected := detected.Length ? detected[1] : ""
	if !(detected is String)
		return ""
	detected := Trim(detected)
	if RegExMatch(detected, "i)^(.+?)\s+-\s+Excel$", &m)
		detected := Trim(m[1])
	return detected
}


; Connect to the SAME Excel window the sweep functions use (the one named in the
; Excel Window Name field) and return its active sheet. The app normally reaches
; Excel only through NotionExcelSync, but reading/writing headers needs a window-
; targeted connection the bridge doesn't expose; rather than add one to the
; bridge, the app calls ExcelCom's connection helper directly here (a deliberate,
; agreed exception to the "friendly layers only" rule). The window title is built
; exactly as NotionExcelSync_setExcelWindow builds it, so both target one window.
NotionSync_excelSheet() {
	winName := Trim(ui.winName.Value)
	winTitle := (winName = "" ? "" : winName . " ") . "ahk_class XLMAIN"
	return _Excel_Get(winTitle).ActiveWorkbook.ActiveSheet
}


; The target sheet's header row (row 1) as an Array of { name, col } for every
; non-blank header cell, left to right. Blank cells are skipped, so a gap between
; headers doesn't cut the scan short. Reads the whole row in one COM call.
NotionSync_readHeaders() {
	sheet := NotionSync_excelSheet()
	; xlToLeft = -4159: walk left from the last column to the used width of row 1.
	lastCol := sheet.Cells(1, sheet.Columns.Count).End(-4159).Column
	headers := []
	if (lastCol = 1) {
		; A single-cell range returns a scalar from .Value (not a 2D array).
		v := sheet.Cells(1, 1).Value
		if !IsObject(v) && Trim(v . "") != ""
			headers.Push({ name: Trim(v . ""), col: 1 })
		return headers
	}
	vals := sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, lastCol)).Value   ; 1-based 2D array
	loop lastCol {
		v := vals[1, A_Index]
		if IsObject(v)
			continue
		name := Trim(v . "")
		if (name != "")
			headers.Push({ name: name, col: A_Index })
	}
	return headers
}


; Write a header name into row 1 of the target sheet at the given column number
; (used by the "Add new header..." option in the header pick-lists).
NotionSync_writeHeader(col, name) {
	NotionSync_excelSheet().Cells(1, col).Value := name
}


; Show a problem in a message box: a plain-English hint when there is one, then
; what went wrong, where, and the call stack (so it's traceable), like the
; libraries' own error boxes.
NotionSync_reportError(err) {
	NotionSync_status("Something went wrong")
	if !(err is Error) {
		NotionSync_msg("Something went wrong:`n`n" . err, "Notion Sync - Error", "Iconx")
		return
	}
	message := err.Message
	if (err.Extra != "")
		message .= "`n`nDetails: " . err.Extra
	message .= "`n`nSource: " . err.File . "  (line " . err.Line . ")"
	if (err.Stack != "")
		message .= "`n`nCall stack:`n" . err.Stack
	hint := NotionSync_hintFor(err)
	if (hint != "")
		message := hint . "`n`n----------`n`n" . message
	NotionSync_msg(message, "Notion Sync - Error", "Iconx")
}


; A plain-English hint for errors people are likely to run into, or "" if there
; isn't one. Excel's COM errors are cryptic: "call rejected" means Excel is busy
; (most often a cell is mid-edit), and "RPC server unavailable" means Excel has
; gone away.
NotionSync_hintFor(err) {
	msg := (err is Error) ? err.Message : err
	if (InStr(msg, "0x80010001") || InStr(msg, "0x800AC472"))
		return "Excel is busy -- usually a cell is being edited or a dialog is open. Finish or cancel that (press Enter or Esc in Excel) and try again."
	if (InStr(msg, "0x800706BA") || InStr(msg, "0x800706BE"))
		return "Excel stopped responding or was closed. Make sure the workbook is still open and try again."
	return ""
}