;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	NotionCom (Function Library)
; DESCRIPTION:	FRIENDLY layer: ergonomic, one-line Notion verbs for a single
;				configured database. Pure Notion -- knows nothing about Excel or
;				any other spreadsheet (it uses UtilityWindows.ahk only for a
;				generic progress window, not for anything Excel-specific). Built
;				on top of NotionAPI.ahk (the raw API/JSON layer), which it hides
;				completely: you never build a NotionAPI_prop* value or read a
;				raw property Map by hand here. This is the stable contract other
;				tools (e.g. NotionExcelSync.ahk) should script against instead of
;				reaching into NotionAPI.ahk.

; VERSION:		2.9.28.26
; AUTHOR:		Noel Gordon (noel.gordon@ceterapartners.com)
; SOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; CONCEPT
;   You configure NotionCom for a single database ID, and every other verb then
;   operates against that one database, served entirely from an IN-MEMORY cache
;   (rows + schema). There is NO automatic/lazy live pull anymore -- reworked
;   2026-09-24 because a live pull is slow, and doing it silently on whatever
;   call happened to need it first was surprising and, for a database of any
;   size, too slow to eat on every script run. Getting data into memory is now
;   always an explicit choice between two verbs:
;     NotionCom_pull(token, databaseId, version)  -- hits the live Notion API,
;       caches the result in memory, AND saves it to a file in the "databaseDir"
;       directory (see verifyDatabaseDir() below) so it doesn't have to be
;       pulled again. Slow (network); this might not run every time your
;       script does -- call it the first time, or whenever you actually want
;       fresh data.
;     NotionCom_load(databaseId)  -- reads that saved file straight into
;       memory. No network call. This is what an ordinary script run should
;       call every time, once NotionCom_pull() has saved the file at least
;       once. Shows the same kind of progress windows a pull does, since
;       reading + parsing a large file still takes a moment.
;   Every matching/reading/writing verb below reads from whatever's currently
;   in memory and THROWS a clear error if nothing has been pulled or loaded
;   yet -- there's no silent fallback to the network.
;
; DATABASE FILE STORAGE
;   NotionCom_pull() and NotionCom_refresh() write a JSON file per database to
;   the directory named by the GLOBAL variable "databaseDir" (same convention
;   as UtilityWindows.ahk's "settingsFile" -- define it in your script before
;   this file is included). See verifyDatabaseDir() for the exact requirement.
;
; QUICK START
;   #Include NotionCom.ahk
;
;   ; First run (or whenever you want fresh data) -- hits the API, saves a file:
;   NotionCom_pull("ntn_your_secret_here", "<database_id>")
;
;   ; Ordinary runs after that -- fast, no network, reads the saved file:
;   NotionCom_load("<database_id>")
;
;   ; --- Match on one property, get back just the page ID ---
;   id := NotionCom_matchId("Email", "jane@example.com")
;
;   ; --- Match and get the whole page in one call ---
;   page := NotionCom_matchPage("Email", "jane@example.com")
;
;   ; --- Match and read a single other property in one call (no second lookup) ---
;   dept := NotionCom_matchProperty("Email", "jane@example.com", "Department")
;
;   ; --- Already have a page ID? Read/write one property at a time ---
;   name := NotionCom_getProperty(id, "Name")
;   NotionCom_setProperty(id, "Department", "Engineering")
;
;   ; --- Write several properties in one request ---
;   NotionCom_setProperties(id, Map("Department", "Engineering", "Active", true))
;
;   ; --- Create a brand-new entry (never automatic -- you always call this) ---
;   newId := NotionCom_createEntry(Map("Name", "New Hire", "Email", "new@example.com"))
;
;   ; NOTE: writes need a live, authenticated connection -- NotionCom_load()
;   ; alone does not set one up (it never touches the network). Only
;   ; NotionCom_pull() does, so a script that only ever loads from file and
;   ; then tries to write will get a clear error from the raw layer.
;
; VERB FAMILIES
;   Getting data into memory:
;     NotionCom_pull(token, databaseId, version := "2022-06-28", showInfo := false)
;       -- live pull + cache + save to file. Configures the connection too
;          (needed for any write). Silent by default; pass showInfo := true for
;          the "Pulling..."/"Available Properties"/"Pull Complete" windows.
;          Returns the row count.
;     NotionCom_refresh()
;       -- re-pulls live and re-saves the file, using the connection
;          NotionCom_pull() already configured this session. Shows the progress
;          windows. Returns the row count.
;     NotionCom_load(databaseId)
;       -- reads the saved file into memory. No network, no connection setup.
;          Shows "Loading..." / "Available Properties" / "Load Complete"
;          windows like pull does. Throws if no file has been saved for this
;          database yet. Returns the row count.
;   Matching (search the in-memory rows by any one property):
;     NotionCom_matchId(matchProp, matchValue, caseSensitive := false)
;       -- the BASE verb. Returns the matching page ID, or "" if nothing matched.
;     NotionCom_matchPage(matchProp, matchValue, caseSensitive := false)
;       -- wraps matchId; returns the whole page object (a Map), or "".
;     NotionCom_matchProperty(matchProp, matchValue, readProp, caseSensitive := false)
;       -- wraps matchId; returns one OTHER property's value straight away, or "".
;     NotionCom_matchIds(matchProp, matchValue, caseSensitive := false)
;       -- returns an Array of ALL matching page IDs (use to detect duplicates
;          before trusting matchId's "first match").
;   Reading a page you already have the ID for:
;     NotionCom_getPage(pageId)       -- the whole page object (a Map).
;     NotionCom_getProperty(pageId, propName) -- one property's value.
;   Looking at the database itself:
;     NotionCom_getSchema()           -- Map of property name -> type (for pick-lists).
;     NotionCom_savedFileTime(databaseId)
;       -- when the saved file was last written ("" if none) -- cheap way to
;          choose between pull() and load().
;   Writing (never automatic -- only when you call these):
;     NotionCom_setProperty(pageId, propName, value)
;     NotionCom_setProperties(pageId, propValueMap)   ; one PATCH for several props
;     NotionCom_createEntry(propValueMap)             ; new page in the configured DB
;   All value/property parameters here are PLAIN values (strings, numbers,
;   true/false, arrays of strings for multi-select) -- this layer looks up each
;   property's Notion type from the database schema and builds the right raw
;   structure for you. You never touch a NotionAPI_prop* builder.
;
; NOTES
;   - Matching here defaults to CASE-INSENSITIVE (unlike the raw layer's
;     matchRow/matchRows, which default to case-sensitive) since the values
;     you're reconciling against -- emails, names -- are rarely typed
;     consistently.
;   - A property name that isn't in the database schema is silently skipped on
;     writes (the schema is a moving target by design; one stale mapped column
;     shouldn't fail an entire push).
;   - A write (setProperty/setProperties/createEntry) updates that ONE page
;     in-place in the in-memory rows cache (using the fresh page object Notion's
;     write response already returns) -- so a later match/get in the same
;     script sees the change without needing NotionCom_refresh(). It does NOT
;     touch the saved file, though -- that still only updates via
;     NotionCom_pull()/NotionCom_refresh(), so the file can drift stale
;     relative to memory across a session of writes.


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2

#Include NotionAPI.ahk
#Include UtilityWindows.ahk   ; async progress windows shown while a database is pulled or loaded


; Verify "databaseDir" variable is defined for error checking
verifyDatabaseDir()


; ##########################################################
; 	PUBLIC CONFIGURATION FUNCTIONS
; ##########################################################


; Configure NotionCom for a single database AND pull it live right now: rows +
; schema, cached in memory and saved to a file in databaseDir (see
; verifyDatabaseDir()) so a later NotionCom_load() can skip the network
; entirely. This might not run every time your script does -- call it the first
; time, or whenever you actually want fresh data; use NotionCom_load() for fast,
; ordinary runs.
;
; token      : Notion integration secret (passed through to NotionAPI_init).
; databaseId : the database every other NotionCom_ verb operates against.
; version    : (optional) Notion-Version header; passed through to NotionAPI_init.
; showInfo   : (optional) when true, show the progress/reference windows the pull
;              used to always show -- the async "Pulling..." start window, the
;              "Available Properties" window, and the "Pull Complete" summary
;              (rows + elapsed time). Defaults to false, so the pull is silent
;              (it still saves the file and throws on error, as always).
;
; Returns the number of rows pulled.
NotionCom_pull(token, databaseId, version := "2022-06-28", showInfo := false) {
	cfg := _NotionCom_cfg()
	cfg.databaseId := databaseId
	NotionAPI_init(token, version)
	return _NotionCom_livePullAndSave(showInfo)
}


; Re-pull the database live and re-save its file. Uses the connection
; NotionCom_pull() already configured this session -- call after you've
; changed data in Notion (or made several writes via NotionCom itself) and
; want the in-memory cache and the saved file both fully back in sync. Shows the
; progress/reference windows (unlike NotionCom_pull, which is silent by default).
;
; Returns the number of rows pulled.
NotionCom_refresh() {
	return _NotionCom_livePullAndSave(true)
}


; Load a previously pulled database straight from its saved file in
; databaseDir into memory -- no network call, no connection setup (so writes
; won't work until NotionCom_pull() has configured one, this session or a
; prior one). Use this for ordinary script runs once NotionCom_pull() has
; saved the file at least once. Reading a big file still takes a moment, so it
; shows the same progress windows NotionCom_pull() does: an async "Loading..."
; window while it reads, an "Available Properties" window, then a "Load
; Complete" box (rows loaded, elapsed time, and when the file was last saved).
;
; databaseId : the database whose saved file to load.
; Throws if no saved file exists yet for this database.
; Returns the number of rows loaded.
NotionCom_load(databaseId) {
	cfg := _NotionCom_cfg()
	cfg.databaseId := databaseId

	; Announce the start. ASYNC, so it shows while the (blocking) file read and
	; parse run; we keep the window object so we can close it when done.
	startTime := FormatTime(A_Now, "h:mm:ss tt")
	startBox := UtilityWindows_parallelMessageBox("NotionCom - Loading Database",
		"Loading the Notion database from file...`n`nStarted at " . startTime)
	Sleep(50)   ; let the async window paint before we block on the file read

	startTick := A_TickCount
	try {
		saved := _NotionCom_loadDatabaseFile(databaseId)
	} catch as err {
		; Don't leave the "Loading..." window up behind the error (e.g. no file yet).
		try startBox.Destroy()
		throw err
	}
	cfg.rows := saved.Has("rows") ? saved["rows"] : []
	cfg.schema := saved.Has("schema") ? saved["schema"] : Map()
	elapsedSeconds := Round((A_TickCount - startTick) / 1000, 1)

	; Close the "starting" window (guarded, in case the user already closed it).
	try startBox.Destroy()

	; Same non-blocking reference window a live pull shows.
	UtilityWindows_parallelMessageBox("NotionCom - Available Properties",
		_NotionCom_propertyListText())

	; Announce completion (blocking, like the pull's). Includes when the file
	; was saved so it's obvious how stale the loaded data might be.
	savedAt := saved.Has("savedAt") ? saved["savedAt"] : "unknown"
	MsgBox("Loaded " . cfg.rows.Length . " row(s) from file in " . elapsedSeconds . " seconds.`n`nFile last saved: " . savedAt,
		"NotionCom - Load Complete", "Iconi")

	return cfg.rows.Length
}




; ##########################################################
; 	PUBLIC MATCHING VERBS (search the cached database)
; ##########################################################


; Find the page whose <matchProp> equals <matchValue> and return its page ID.
; This is the BASE matching verb -- NotionCom_matchPage and
; NotionCom_matchProperty are convenience wrappers around this same lookup.
;
; Searches whatever's currently in memory -- call NotionCom_pull() or
; NotionCom_load() first. Throws if nothing has been pulled/loaded yet.
;
; Returns the page ID (String), or "" if nothing matched.
NotionCom_matchId(matchProp, matchValue, caseSensitive := false) {
	row := NotionAPI_matchRow(_NotionCom_rows(), matchProp, matchValue, caseSensitive)
	return (row = "") ? "" : row["id"]
}


; Like NotionCom_matchId, but returns the whole matching page object (a Map)
; instead of just its ID -- use this when you need more than one property off
; the match and don't want a second lookup.
;
; Returns the matching page (a Map), or "" if nothing matched.
NotionCom_matchPage(matchProp, matchValue, caseSensitive := false) {
	id := NotionCom_matchId(matchProp, matchValue, caseSensitive)
	return (id = "") ? "" : NotionCom_getPage(id)
}


; Like NotionCom_matchId, but immediately reads a different property off the
; matched page instead of returning the ID -- the shortest path from "an
; external value" to "the Notion fact you actually wanted."
;
; matchProp/matchValue : what to search for.
; readProp              : the property to read off the page that matched.
; Returns the value of <readProp> on the matched page, or "" if nothing matched.
NotionCom_matchProperty(matchProp, matchValue, readProp, caseSensitive := false) {
	id := NotionCom_matchId(matchProp, matchValue, caseSensitive)
	return (id = "") ? "" : NotionCom_getProperty(id, readProp)
}


; Like NotionCom_matchId, but returns ALL matching page IDs as an Array (empty
; if none). Use this to detect duplicate keys before trusting matchId's
; "first match wins" behavior.
NotionCom_matchIds(matchProp, matchValue, caseSensitive := false) {
	out := []
	for row in NotionAPI_matchRows(_NotionCom_rows(), matchProp, matchValue, caseSensitive)
		out.Push(row["id"])
	return out
}




; ##########################################################
; 	PUBLIC READING VERBS (you already know the page ID)
; ##########################################################


; Return the whole page object for a known page ID. Served from the cached
; database pull when the page is already cached; otherwise fetched live (e.g.
; a page created after your last match/refresh).
;
; Returns the page (a Map). Throws if the page doesn't exist / isn't accessible.
NotionCom_getPage(pageId) {
	clean := StrReplace(pageId, "-", "")
	for row in _NotionCom_rows()
		if (StrReplace(row["id"], "-", "") = clean)
			return row
	return NotionAPI_retrievePage(pageId)
}


; The atom of reading: one property off a known page ID, as a plain value.
; Returns the property's value (see NotionAPI_getPropValue for the type-by-type
; return shapes this passes through unchanged).
NotionCom_getProperty(pageId, propName) {
	return NotionAPI_getPropValue(NotionCom_getPage(pageId), propName)
}


; The shape of the loaded database: a Map of "<property name>" -> "<type>"
; (e.g. "Work Email" -> "email"), alphabetized by name. Handy for building a
; pick-list of valid property names instead of typing them. Treat it as
; read-only -- it's the live in-memory copy, not a duplicate.
;
; Throws if nothing has been pulled/loaded yet.
NotionCom_getSchema() {
	return _NotionCom_schema()
}


; When this database's saved file was last written, as a YYYYMMDDHH24MISS
; timestamp (pass it to FormatTime), or "" if no file has been saved for it yet.
; Cheap -- no network and no parsing -- so a script can use it to choose between
; NotionCom_pull() and NotionCom_load().
NotionCom_savedFileTime(databaseId) {
	path := _NotionCom_dbFilePath(databaseId)
	return FileExist(path) ? FileGetTime(path) : ""
}




; ##########################################################
; 	PUBLIC WRITING VERBS (never automatic -- only when called)
; ##########################################################


; The atom of writing: set one property on a known page ID. Convenience
; wrapper around NotionCom_setProperties for a single property.
;
; value : a PLAIN value (string, number, true/false, or an Array of strings for
;         a multi-select) -- NOT a NotionAPI_prop* builder result.
; Returns the page ID on success, or "" if <propName> isn't a writable property
; on this database.
NotionCom_setProperty(pageId, propName, value) {
	return NotionCom_setProperties(pageId, Map(propName, value))
}


; Set several properties on a known page ID in a single Notion request.
;
; propValueMap : Map of "<Notion property name>" -> a PLAIN value. Each value's
;                Notion type is looked up from the database schema so the right
;                raw structure is built for you. A property not present in the
;                schema, or a read-only type (formula, rollup, created/last-
;                edited *, unique_id, files), is silently skipped.
;
; Returns the page ID on success, or "" if nothing ended up writable.
NotionCom_setProperties(pageId, propValueMap) {
	schema := _NotionCom_schema()
	props := _NotionCom_buildProps(schema, propValueMap)
	if (props.Count = 0)
		return ""
	updated := NotionAPI_updatePage(pageId, props)
	_NotionCom_replaceRow(updated)
	return pageId
}


; Create a new page (row) in the configured database. NEVER called
; automatically by any other NotionCom_ verb -- only when your script calls it.
;
; propValueMap : Map of "<Notion property name>" -> a PLAIN value, same as
;                NotionCom_setProperties. Must include the database's title
;                property (Notion rejects a page with no title).
;
; Returns the new page's ID.
NotionCom_createEntry(propValueMap) {
	cfg := _NotionCom_cfg()
	schema := _NotionCom_schema()
	props := _NotionCom_buildProps(schema, propValueMap)
	page := NotionAPI_createPage(cfg.databaseId, props)
	_NotionCom_replaceRow(page)
	return page["id"]
}




; ##########################################################
; 	PRIVATE UTILITY FUNCTIONS
; ##########################################################


; Internal: shared configuration/cache store (persists across calls).
_NotionCom_cfg() {
	static cfg := { databaseId: "", rows: "", schema: "" }
	return cfg
}


; Internal: throw-on-miss accessor for the in-memory rows cache. Populated by
; NotionCom_pull()/NotionCom_refresh() (live) or NotionCom_load() (from the
; saved file) -- never fetched automatically from here. See NotionCom_pull().
_NotionCom_rows() {
	cfg := _NotionCom_cfg()
	if !(cfg.rows is Array)
		throw Error("NotionCom: no database loaded. Call NotionCom_pull(token, databaseId) or NotionCom_load(databaseId) first.", -1)
	return cfg.rows
}


; Internal: live-pull rows + schema from the Notion API, cache them in memory,
; and persist the result to this database's file in databaseDir. Shared by
; NotionCom_pull() and NotionCom_refresh() -- the only two places that ever
; trigger a live network pull. When showInfo is true it also shows the
; progress/reference windows (the async start window, "Available Properties",
; and "Pull Complete"); otherwise it runs silently. The pull + save happen either
; way, and errors always throw.
_NotionCom_livePullAndSave(showInfo := false) {
	cfg := _NotionCom_cfg()

	; Announce the start -- only when info windows are wanted. This is the ASYNC
	; message box, so it shows while the (blocking) pull runs instead of pausing
	; it. We keep its window object so we can close it once the pull is done.
	startBox := ""
	if (showInfo) {
		startTime := FormatTime(A_Now, "h:mm:ss tt")
		startBox := UtilityWindows_parallelMessageBox("NotionCom - Pulling Database",
			"Pulling the Notion database...`n`nStarted at " . startTime)
		Sleep(50)   ; let the async window paint before we block on the network
	}

	; Do the (potentially slow) full-database pull, timing it.
	startTick := A_TickCount
	try {
		cfg.rows := NotionAPI_queryDatabase(cfg.databaseId)
		cfg.schema := _NotionCom_fetchSchemaLive(cfg.databaseId)
	} catch as err {
		; Don't leave the "Pulling..." window up behind the error (bad API key,
		; no network, unknown database, ...).
		if IsObject(startBox)
			try startBox.Destroy()
		throw err
	}
	elapsedSeconds := Round((A_TickCount - startTick) / 1000, 1)

	; Close the "starting" window (guarded, in case the user already closed it).
	if IsObject(startBox)
		try startBox.Destroy()

	; Show the database's available properties as a second, non-blocking
	; reference window -- lets you see valid property names/types without
	; cross-referencing the real Notion database while writing a script. Left
	; open (not auto-closed) so it can sit on screen for reference.
	if (showInfo)
		UtilityWindows_parallelMessageBox("NotionCom - Available Properties",
			_NotionCom_propertyListText())

	_NotionCom_saveDatabaseFile(cfg.databaseId, cfg.rows, cfg.schema)

	; Announce completion. This is a NORMAL (blocking) message box, since the
	; work is already done and nothing needs to run behind it.
	if (showInfo)
		MsgBox("Pulled " . cfg.rows.Length . " row(s) from Notion in " . elapsedSeconds . " seconds, and saved to file.",
			"NotionCom - Pull Complete", "Iconi")

	return cfg.rows.Length
}


; Internal: insert/replace a single page in the in-memory rows cache (matched
; by ID), so a write's effect is visible to a later match/get in the same
; script without needing NotionCom_refresh(). No-op if nothing is loaded yet
; (setProperties/createEntry would already have thrown via _NotionCom_schema()
; before reaching this, so that shouldn't happen in practice). Does NOT touch
; the saved file -- see the NOTES at the top of this file.
_NotionCom_replaceRow(page) {
	cfg := _NotionCom_cfg()
	if !(cfg.rows is Array)
		return
	clean := StrReplace(page["id"], "-", "")
	for i, row in cfg.rows {
		if (StrReplace(row["id"], "-", "") = clean) {
			cfg.rows[i] := page
			return
		}
	}
	cfg.rows.Push(page)
}


; Internal: build a "<name> (<type>)" list, one per line, of every property in
; the configured database's schema -- the text shown in the "Available
; Properties" window above. AHK Maps enumerate keys alphabetically, so the
; list comes out alphabetized with no extra sorting needed.
_NotionCom_propertyListText() {
	schema := _NotionCom_schema()
	text := "This database has " . schema.Count . " propert" . (schema.Count = 1 ? "y" : "ies") . ":`n`n"
	for name, type in schema
		text .= name . " (" . type . ")`n"
	return text
}


; Internal: throw-on-miss accessor for the in-memory schema cache. Populated
; the same way rows are -- see _NotionCom_rows() above.
_NotionCom_schema() {
	cfg := _NotionCom_cfg()
	if !(cfg.schema is Map)
		throw Error("NotionCom: no database loaded. Call NotionCom_pull(token, databaseId) or NotionCom_load(databaseId) first.", -1)
	return cfg.schema
}


; Internal: fetch the schema live from the API as a Map of "<property name>"
; -> "<type>". Used only by _NotionCom_livePullAndSave() -- never called
; lazily from a reader (see _NotionCom_schema() above).
_NotionCom_fetchSchemaLive(databaseId) {
	db := NotionAPI_retrieveDatabase(databaseId)
	sch := Map()
	if (db.Has("properties"))
		for name, def in db["properties"]
			sch[name] := def.Has("type") ? def["type"] : ""
	return sch
}


; Internal: the path this database's saved file lives at inside databaseDir.
_NotionCom_dbFilePath(databaseId) {
	global databaseDir
	return databaseDir . "\" . StrReplace(Trim(databaseId), "-", "") . ".json"
}


; Internal: write rows + schema to this database's file (JSON, overwriting any
; existing file). Reuses NotionAPI's own JSON serializer (a private helper
; there) rather than duplicating one here -- the one place NotionCom reaches
; past NotionAPI's public surface, justified by keeping file storage entirely
; inside this file rather than touching NotionAPI.ahk.
_NotionCom_saveDatabaseFile(databaseId, rows, schema) {
	path := _NotionCom_dbFilePath(databaseId)
	payload := Map(
		"databaseId", databaseId,
		"savedAt", FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss"),
		"rows", rows,
		"schema", schema)
	text := _NotionAPI_jsonStringify(payload)
	if FileExist(path)
		FileDelete(path)
	FileAppend(text, path, "UTF-8")
}


; Internal: read + parse this database's saved file. Throws a clear,
; actionable error if it doesn't exist yet.
_NotionCom_loadDatabaseFile(databaseId) {
	path := _NotionCom_dbFilePath(databaseId)
	if !FileExist(path)
		throw Error("NotionCom: no saved database file found (expected " . path . "). Call NotionCom_pull(token, databaseId) first.", -1)
	fileContents := FileRead(path, "UTF-8")
	return _NotionAPI_jsonParse(fileContents)
}


; Internal: turn a Map of "<property name>" -> plain value into a Map of
; "<property name>" -> raw NotionAPI_prop* structure, ready to send. Properties
; not in the schema, or of a read-only type, are left out entirely.
_NotionCom_buildProps(schema, propValueMap) {
	props := Map()
	for propName, val in propValueMap {
		if !schema.Has(propName)
			continue
		built := _NotionCom_buildPropValue(schema[propName], val)
		if (built != "")
			props[propName] := built
	}
	return props
}


; Internal: build the correct NotionAPI_prop* value for a write, from a plain
; value and the property's Notion type. Returns "" for read-only/unsupported
; types (the caller then skips that property).
_NotionCom_buildPropValue(type, value) {
	; An Array value (multi_select / relation append) has no string form, so only
	; stringify non-array values; the array-taking cases below read `value` directly.
	s := (value is Array) ? "" : Trim(value . "")

	switch type {
		case "title":
			return NotionAPI_propTitle(s)
		case "rich_text":
			return NotionAPI_propRichText(s)
		case "number":
			return (s = "") ? Map("number", NotionAPI_null()) : NotionAPI_propNumber(Number(s))
		case "checkbox":
			return NotionAPI_propCheckbox(_NotionCom_truthy(value))
		case "select":
			return (s = "") ? Map("select", NotionAPI_null()) : NotionAPI_propSelect(s)
		case "status":
			return (s = "") ? "" : NotionAPI_propStatus(s)   ; status cannot be cleared to null
		case "multi_select":
			names := []
			if (value is Array) {
				for n in value
					if (Trim(n . "") != "")
						names.Push(Trim(n . ""))
			} else if (s != "") {
				for part in StrSplit(s, ";")
					if (Trim(part) != "")
						names.Push(Trim(part))
			}
			return NotionAPI_propMultiSelect(names)
		case "date":
			return (s = "") ? Map("date", NotionAPI_null()) : NotionAPI_propDate(s)
		case "url":
			return NotionAPI_propUrl(s)
		case "email":
			return NotionAPI_propEmail(s)
		case "phone_number":
			return NotionAPI_propPhone(s)
		case "relation":
			ids := []
			if (value is Array) {
				for id in value
					if (Trim(id . "") != "")
						ids.Push(Trim(id . ""))
			} else if (s != "") {
				for part in StrSplit(s, ",")
					if (Trim(part) != "")
						ids.Push(Trim(part))
			}
			return NotionAPI_propRelation(ids)
		default:
			return ""   ; formula, rollup, created/last-edited *, unique_id, files, etc.
	}
}


; Internal: interpret a plain value as a boolean for a checkbox property.
_NotionCom_truthy(value) {
	if (value is Integer || value is Float)
		return value != 0
	v := StrLower(Trim(value . ""))
	return (v = "true" || v = "yes" || v = "y" || v = "x" || v = "1" || v = "checked" || v = Chr(0x2713))
}




; Verify "databaseDir" variable is defined for error checking
; Also create the required directory if it does not exist
verifyDatabaseDir(){
	global databaseDir

	varSet := IsSet(databaseDir)
	if (varSet){
		if InStr(FileExist(databaseDir), "D"){
			return
		}
		else{
			DirCreate databaseDir
		}
	}
	else{
		msg := "ERROR: NotionCom_verifyDatabaseDir(): databaseDir variable is not set. This variable should be defined in the main script."
		MsgBox(msg)
		Exit()
	}
}
