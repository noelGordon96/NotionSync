;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	NotionAPI (Function Library)
; DESCRIPTION:	RAW layer: talks to the Notion API and parses its JSON. Thin,
;				unopinionated, maps closely to the HTTP API. Complex by design --
;				this layer is not held to the repo's usual readability bar (see
;				NotionCom.ahk's header). Most scripts should use NotionCom.ahk
;				(the friendly verb layer built on top of this one) instead of
;				calling NotionAPI_* directly.

; VERSION:		3.9.24.26
; AUTHOR:		Noel Gordon (noel.gordon@ceterapartners.com)
; SOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; SETUP
;   1. Create an internal integration at https://www.notion.com/my-integrations
;      and copy its "Internal Integration Secret" (starts with "ntn_" or the
;      older "secret_").
;   2. Open the target database in Notion, click the "..." menu -> Connections,
;      and connect your integration so it has access.
;   3. The database ID is the 32-char hex string in the database URL:
;        https://www.notion.so/<workspace>/<DATABASE_ID>?v=...
;      A page ID is the 32-char hex string in a page URL. Hyphens are optional.
;
; QUICK START (raw layer -- prefer NotionCom.ahk's friendly verbs for scripting)
;   #Include NotionAPI.ahk
;
;   NotionAPI_init("ntn_your_secret_here")
;
;   ; --- Read every row of a database ---
;   rows := NotionAPI_queryDatabase("<database_id>")
;   for row in rows
;       MsgBox(NotionAPI_getPropValue(row, "Name"))
;
;   ; --- Add a new row ---
;   props := Map(
;       "Name",     NotionAPI_propTitle("New task"),
;       "Status",   NotionAPI_propSelect("In progress"),
;       "Priority", NotionAPI_propNumber(3),
;       "Done",     NotionAPI_propCheckbox(false))
;   page := NotionAPI_createPage("<database_id>", props)
;
;   ; --- Update an existing row ---
;   NotionAPI_updatePage(page["id"], Map("Done", NotionAPI_propCheckbox(true)))
;
;   ; --- Export a whole database to CSV ---
;   NotionAPI_databaseToCsv("<database_id>", A_Desktop . "\export.csv")
;
; NOTES
;   - Every request is synchronous (blocks until Notion responds).
;   - On any HTTP error (4xx/5xx) the request functions throw an Error whose
;     .Message contains Notion's own error code and message. Wrap calls in
;     try/catch for graceful handling.
;   - JSON parsing maps a JSON object to an AHK Map (case-sensitive keys, as
;     Notion requires), a JSON array to an AHK Array, true/false to 1/0, and
;     null to "" (an empty string). To WRITE an explicit JSON null in a request,
;     use NotionAPI_null().


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2

; Dependencies: none. HTTP is handled via the OS "WinHttp.WinHttpRequest.5.1"
; COM object and JSON is handled by the self-contained parser/serializer in the
; PRIVATE UTILITY FUNCTIONS section below.




; ##########################################################
; 	PUBLIC CONNECTION / CONFIGURATION FUNCTIONS
; ##########################################################


; Initialize the library with your integration secret. Must be called once
; before any other NotionAPI_ request function.
;
; token   : the integration secret (e.g. "ntn_..." or "secret_...").
; version : (optional) the Notion-Version header to send. Defaults to the
;           long-stable "2022-06-28"; override only if you need newer features.
NotionAPI_init(token, version := "2022-06-28") {
	cfg := _NotionAPI_cfg()
	cfg.token := Trim(token)
	cfg.version := version
}




; ##########################################################
; 	PUBLIC DATABASE / PAGE FUNCTIONS
; ##########################################################


; Retrieve a database object, including its schema (the "properties" map that
; describes each column's name and type). Useful for discovering property names
; and select/multi-select option names.
;
; databaseId : the database ID (hyphens optional).
; Returns the parsed database object (a Map).
NotionAPI_retrieveDatabase(databaseId) {
	return _NotionAPI_request("GET", "/databases/" . _NotionAPI_cleanId(databaseId))
}


; Query a database and return ALL matching rows, transparently following
; Notion's pagination until every result has been collected.
;
; databaseId : the database ID (hyphens optional).
; filter     : (optional) a Notion filter object (Map). Pass "" for no filter.
;              See https://developers.notion.com/reference/post-database-query-filter
;              e.g. Map("property","Done","checkbox",Map("equals",NotionAPI_bool(true)))
; sorts      : (optional) a Notion sorts array (Array of Maps). Pass "" for none.
;              e.g. [ Map("property","Priority","direction","descending") ]
; pageSize   : (optional) rows fetched per request (max 100, the default).
;
; Returns an Array of page objects (each a Map). Read a property from a row with
; NotionAPI_getPropValue(row, "<Property Name>").
NotionAPI_queryDatabase(databaseId, filter := "", sorts := "", pageSize := 100) {
	dbId := _NotionAPI_cleanId(databaseId)
	results := []
	cursor := ""

	loop {
		body := Map("page_size", pageSize)
		if (filter != "")
			body["filter"] := filter
		if (sorts != "")
			body["sorts"] := sorts
		if (cursor != "")
			body["start_cursor"] := cursor

		resp := _NotionAPI_request("POST", "/databases/" . dbId . "/query", body)

		if (resp.Has("results"))
			for row in resp["results"]
				results.Push(row)

		; Stop when Notion reports no further pages
		if !(resp.Has("has_more") && resp["has_more"] && resp.Has("next_cursor") && resp["next_cursor"] != "")
			break
		cursor := resp["next_cursor"]
	}

	return results
}


; Retrieve a single page (a database row is a page).
;
; pageId : the page ID (hyphens optional).
; Returns the parsed page object (a Map).
NotionAPI_retrievePage(pageId) {
	return _NotionAPI_request("GET", "/pages/" . _NotionAPI_cleanId(pageId))
}


; Create a new page (row) inside a database.
;
; databaseId : the parent database ID (hyphens optional).
; properties : a Map of "<Property Name>" -> a value built with one of the
;              NotionAPI_prop* helpers below. The database's Title property is
;              required (build it with NotionAPI_propTitle).
;
; Returns the newly created page object (a Map); read its "id" for later edits.
NotionAPI_createPage(databaseId, properties) {
	body := Map(
		"parent", Map("database_id", _NotionAPI_cleanId(databaseId)),
		"properties", properties)
	return _NotionAPI_request("POST", "/pages", body)
}


; Update one or more properties of an existing page (row). Only the properties
; you include are changed; all others are left untouched.
;
; pageId     : the page ID (hyphens optional).
; properties : a Map of "<Property Name>" -> a NotionAPI_prop* value.
;
; Returns the updated page object (a Map).
NotionAPI_updatePage(pageId, properties) {
	return _NotionAPI_request("PATCH", "/pages/" . _NotionAPI_cleanId(pageId),
		Map("properties", properties))
}


; Archive a page (row). Notion has no hard-delete via the API; archiving is the
; standard "delete" and removes the row from the database view. Restore with
; NotionAPI_restorePage.
;
; pageId : the page ID (hyphens optional).
; Returns the updated (archived) page object (a Map).
NotionAPI_archivePage(pageId) {
	return _NotionAPI_request("PATCH", "/pages/" . _NotionAPI_cleanId(pageId),
		Map("archived", NotionAPI_bool(true)))
}


; Restore (un-archive) a previously archived page.
;
; pageId : the page ID (hyphens optional).
; Returns the updated page object (a Map).
NotionAPI_restorePage(pageId) {
	return _NotionAPI_request("PATCH", "/pages/" . _NotionAPI_cleanId(pageId),
		Map("archived", NotionAPI_bool(false)))
}




; ##########################################################
; 	PUBLIC MATCHING / LOOKUP FUNCTIONS
; ##########################################################


; These help reconcile external data (e.g. a spreadsheet row) against Notion
; entries by matching on a chosen property value. The CALLING SCRIPT is
; responsible for choosing a property that uniquely identifies a row; if more
; than one entry matches, the "first match" functions return only the first one
; found (use NotionAPI_matchRows to detect duplicates).
;
; Matching is done client-side against already-fetched rows via
; NotionAPI_getPropValue, so it works uniformly for any property type (title,
; rich_text, number, select, status, email, date, formula, etc.). Values are
; compared as text. An Array-valued property (rollup, multi_select, people,
; relation) matches if ANY entry in it equals the value -- so matching on a
; rollup that shows a related page's email finds the row that email belongs to.
; (Before 2026-09-24 array-valued properties were skipped as match keys, which
; made such a match silently find nothing.)


; Return the first row in an already-fetched rows Array whose <propName> value
; equals <value>. Pull the rows once with NotionAPI_queryDatabase, then match
; many external records against them in memory (no extra API calls).
;
; rows          : an Array of page objects (from NotionAPI_queryDatabase).
; propName      : the property (column) name to match on (case-sensitive name).
; value         : the value to compare against (compared as text).
; caseSensitive : (optional) compare values case-sensitively. Default true.
;
; Returns the matching row (a Map), or "" if nothing matched.
NotionAPI_matchRow(rows, propName, value, caseSensitive := true) {
	target := value . ""
	for row in rows
		if _NotionAPI_valueMatches(NotionAPI_getPropValue(row, propName), target, caseSensitive)
			return row
	return ""
}


; Like NotionAPI_matchRow, but return ALL matching rows as an Array (empty if
; none). Useful for detecting duplicate keys before trusting a match.
NotionAPI_matchRows(rows, propName, value, caseSensitive := true) {
	out := []
	target := value . ""
	for row in rows
		if _NotionAPI_valueMatches(NotionAPI_getPropValue(row, propName), target, caseSensitive)
			out.Push(row)
	return out
}


; Convenience: query a database and return the first entry whose <propName>
; equals <value>. Fetches every row on each call, so when matching many records
; prefer NotionAPI_queryDatabase once + NotionAPI_matchRow in a loop.
;
; Returns the matching page (a Map), or "" if nothing matched. Read the matched
; entry's ID with the "id" key (e.g. match["id"]).
NotionAPI_findPage(databaseId, propName, value, caseSensitive := true) {
	rows := NotionAPI_queryDatabase(databaseId)
	return NotionAPI_matchRow(rows, propName, value, caseSensitive)
}




; ##########################################################
; 	PUBLIC PROPERTY BUILDER FUNCTIONS (for create / update requests)
; ##########################################################


; Each helper returns a Map shaped as Notion expects for that property type.
; Use them as the values of the "properties" Map you pass to NotionAPI_createPage
; or NotionAPI_updatePage.


; Title property (every database has exactly one).
NotionAPI_propTitle(text) {
	return Map("title", [ Map("text", Map("content", text)) ])
}

; Rich text (a plain-text-only value; formatting is not applied).
NotionAPI_propRichText(text) {
	return Map("rich_text", [ Map("text", Map("content", text)) ])
}

; Number property. Pass an integer or float.
NotionAPI_propNumber(number) {
	return Map("number", number)
}

; Checkbox property. Pass a truthy/falsy value.
NotionAPI_propCheckbox(value) {
	return Map("checkbox", NotionAPI_bool(value))
}

; Single select. optionName must match (or create, if allowed) an option name.
NotionAPI_propSelect(optionName) {
	return Map("select", Map("name", optionName))
}

; Status property (behaves like select but is a distinct Notion type).
NotionAPI_propStatus(optionName) {
	return Map("status", Map("name", optionName))
}

; Multi-select. Pass an Array of option-name strings.
NotionAPI_propMultiSelect(names) {
	arr := []
	for n in names
		arr.Push(Map("name", n))
	return Map("multi_select", arr)
}

; Date property. Dates are ISO-8601 strings, e.g. "2026-09-22" or
; "2026-09-22T15:30:00-07:00". Pass endISO := "" for a single (non-range) date.
NotionAPI_propDate(startISO, endISO := "") {
	d := Map("start", startISO)
	d["end"] := (endISO != "") ? endISO : NotionAPI_null()
	return Map("date", d)
}

; URL property.
NotionAPI_propUrl(url) {
	return Map("url", url)
}

; Email property.
NotionAPI_propEmail(email) {
	return Map("email", email)
}

; Phone number property.
NotionAPI_propPhone(phone) {
	return Map("phone_number", phone)
}

; Relation property. Pass an Array of related page-ID strings.
NotionAPI_propRelation(pageIds) {
	arr := []
	for id in pageIds
		arr.Push(Map("id", _NotionAPI_cleanId(id)))
	return Map("relation", arr)
}


; --- JSON value sentinels (for booleans / null inside request bodies) ---

; Return a sentinel the serializer emits as JSON true/false. Needed because AHK
; cannot distinguish a real boolean from the integers 1/0.
NotionAPI_bool(value) {
	static t := { isBool: true, val: true }
	static f := { isBool: true, val: false }
	return value ? t : f
}

; Return a sentinel the serializer emits as JSON null.
NotionAPI_null() {
	static n := { isNull: true }
	return n
}




; ##########################################################
; 	PUBLIC PROPERTY READER FUNCTIONS (for query / retrieve responses)
; ##########################################################


; Extract a plain, easy-to-use AHK value from a property of a returned page.
;
; page     : a page object (a row from NotionAPI_queryDatabase, or the result of
;            NotionAPI_retrievePage / create / update).
; propName : the exact property (column) name, case-sensitive.
;
; Return value depends on the property type:
;   title, rich_text      -> the concatenated plain text (String)
;   number                -> the number, or "" if empty
;   checkbox              -> 1 or 0
;   select, status        -> the selected option name, or "" if none
;   multi_select          -> an Array of selected option names
;   date                  -> the ISO start date string, or "" if empty
;   url, email, phone     -> the string value, or "" if empty
;   people                -> an Array of person names (or ids if unnamed)
;   relation              -> an Array of related page-ID strings
;   formula               -> the resolved value (a date formula gives its ISO start)
;   rollup                -> number / date: the value (a date gives its ISO start).
;                            "Show original" style (an array): an Array of the
;                            PLAIN values of each rolled-up item, exactly as the
;                            item's own property type would read -- e.g. a rollup
;                            of an email property gives ["a@b.com"], of a
;                            multi-select gives every option name, flattened.
;                            Empty items are left out.
;   (anything else)       -> the raw parsed value for that type
;
; Throws if the property name is not present on the page.
NotionAPI_getPropValue(page, propName) {
	if !(page is Map) || !page.Has("properties")
		throw Error("NotionAPI_getPropValue: page has no 'properties'.", -1)

	props := page["properties"]
	if !props.Has(propName)
		throw Error("NotionAPI_getPropValue: property not found: " . propName, -1)

	return _NotionAPI_propObjectValue(props[propName])
}




; ##########################################################
; 	PUBLIC CSV EXPORT FUNCTIONS
; ##########################################################


; Export an entire database to a CSV file (UTF-8 with BOM, RFC-4180 quoting,
; CRLF line endings -- opens cleanly in Excel, including unicode text).
;
; databaseId : the database ID (hyphens optional).
; filePath   : full path of the .csv file to write. An existing file is
;              overwritten.
; columns    : (optional) an Array of property (column) names to export, in the
;              order you want them. Pass "" (default) to auto-detect every
;              column from the database schema.
; filter     : (optional) a Notion filter object, passed through to the query.
; sorts      : (optional) a Notion sorts array, passed through to the query.
; delimiter  : (optional) field separator. Defaults to "," ; use "`t" for TSV.
;
; Array-valued properties (multi_select, people, relation) are flattened into a
; single cell, their entries joined with "; ". Returns the number of data rows
; written (excludes the header row).
NotionAPI_databaseToCsv(databaseId, filePath, columns := "", filter := "", sorts := "", delimiter := ",") {
	; Determine the column list from the schema when none was supplied
	if !(columns is Array) {
		columns := []
		db := NotionAPI_retrieveDatabase(databaseId)
		if (db.Has("properties"))
			for name, def in db["properties"]
				columns.Push(name)
	}

	rows := NotionAPI_queryDatabase(databaseId, filter, sorts)

	; Header row
	headerFields := []
	for name in columns
		headerFields.Push(_NotionAPI_csvField(name, delimiter))
	content := _NotionAPI_join(headerFields, delimiter) . "`r`n"

	; Data rows
	written := 0
	for row in rows {
		fields := []
		for name in columns {
			val := NotionAPI_getPropValue(row, name)

			; Flatten arrays (multi_select / people / relation) into one cell
			if (val is Array) {
				flat := ""
				for i, e in val
					flat .= (i > 1 ? "; " : "") . (IsObject(e) ? "" : e)
				val := flat
			} else if IsObject(val) {
				val := ""
			}

			fields.Push(_NotionAPI_csvField(val, delimiter))
		}
		content .= _NotionAPI_join(fields, delimiter) . "`r`n"
		written++
	}

	if FileExist(filePath)
		FileDelete(filePath)
	FileAppend(content, filePath, "UTF-8")   ; "UTF-8" writes a BOM

	return written
}




; ##########################################################
; 	PRIVATE UTILITY FUNCTIONS
; ##########################################################


; Internal: shared configuration store. Held in a static object so its values
; persist across calls and can be mutated in place (objects are references).
_NotionAPI_cfg() {
	static cfg := { token: "", version: "2022-06-28", baseUrl: "https://api.notion.com/v1" }
	return cfg
}


; Internal: perform a single Notion API request and return the parsed response.
;
; method   : "GET", "POST", "PATCH", "DELETE", etc.
; endpoint : path beginning with "/" (e.g. "/pages", "/databases/<id>/query").
; bodyObj  : (optional) a Map/Array to serialize as the JSON request body. Pass
;            "" (default) to send no body.
;
; Returns the parsed response (a Map). Throws an Error on transport failure or
; on any HTTP status >= 400, including Notion's error code and message.
_NotionAPI_request(method, endpoint, bodyObj := "") {
	cfg := _NotionAPI_cfg()
	if (cfg.token = "")
		throw Error("NotionCom is not initialized. Call NotionAPI_init(token) first.", -1)

	http := ComObject("WinHttp.WinHttpRequest.5.1")

	; Timeouts in ms: resolve, connect, send, receive
	http.SetTimeouts(10000, 10000, 30000, 60000)

	http.Open(method, cfg.baseUrl . endpoint, false)
	http.SetRequestHeader("Authorization", "Bearer " . cfg.token)
	http.SetRequestHeader("Notion-Version", cfg.version)
	http.SetRequestHeader("Content-Type", "application/json")

	try {
		if (bodyObj = "")
			http.Send()
		else
			; Send the body as raw UTF-8 bytes so non-ASCII text (accents,
			; emoji, etc.) is transmitted correctly.
			http.Send(_NotionAPI_utf8Array(_NotionAPI_jsonStringify(bodyObj)))
	} catch as err {
		throw Error("NotionCom network error: " . err.Message, -1)
	}

	status := http.Status
	text := http.ResponseText

	; Empty body (e.g. some 204s) -> return an empty Map
	if (Trim(text) = "")
		result := Map()
	else
		result := _NotionAPI_jsonParse(text)

	if (status >= 400) {
		code := (result is Map && result.Has("code")) ? result["code"] : "unknown"
		msg := (result is Map && result.Has("message")) ? result["message"] : text
		throw Error("Notion API error " . status . " (" . code . "): " . msg, -1)
	}

	return result
}


; Internal: strip hyphens and whitespace from a Notion ID so both the hyphenated
; and bare 32-char hex forms are accepted everywhere.
_NotionAPI_cleanId(id) {
	return StrReplace(Trim(id), "-", "")
}


; Internal: build a UTF-8 byte SAFEARRAY from a string, for sending as an HTTP
; body.
_NotionAPI_utf8Array(str) {
	n := StrPut(str, "UTF-8") - 1        ; byte count, excluding null terminator
	if (n < 0)
		n := 0
	buf := Buffer(n + 1)
	StrPut(str, buf, "UTF-8")
	arr := ComObjArray(0x11, n)          ; VT_UI1 (unsigned byte) array
	i := 0
	while (i < n) {
		arr[i] := NumGet(buf, i, "UChar")
		i++
	}
	return arr
}


; Internal: concatenate the plain_text pieces of a rich-text / title array.
_NotionAPI_richToText(arr) {
	out := ""
	if (arr is Array)
		for piece in arr
			out .= (piece is Map && piece.Has("plain_text")) ? piece["plain_text"] : ""
	return out
}


; Internal: turn a null-sentinel / "" into "", otherwise pass the value through.
_NotionAPI_orBlank(v) {
	return _NotionAPI_isNull(v) ? "" : v
}


; Internal: true if a value is the JSON-null sentinel.
_NotionAPI_isNull(v) {
	return IsObject(v) && v.HasOwnProp("isNull")
}


; Internal: does a value read with NotionAPI_getPropValue equal <target> (text)?
; A plain value is compared as text. An Array matches if any entry matches. Any
; other object (a property type with no plain value, like a button) never matches.
_NotionAPI_valueMatches(pv, target, caseSensitive) {
	if (pv is Array) {
		for item in pv
			if (!IsObject(item) && _NotionAPI_textEquals(item . "", target, caseSensitive))
				return true
		return false
	}
	if IsObject(pv)
		return false
	return _NotionAPI_textEquals(pv . "", target, caseSensitive)
}


; Internal: compare two strings, optionally ignoring case.
_NotionAPI_textEquals(a, b, caseSensitive) {
	return caseSensitive ? (a == b) : (StrLower(a) = StrLower(b))
}


; Internal: turn ONE property object -- a Map with a "type" key and its value
; under that key, as found on a page OR as an element inside a rollup's array --
; into a plain AHK value. See NotionAPI_getPropValue for what each type gives.
_NotionAPI_propObjectValue(p) {
	if !(p is Map)
		return ""
	propType := p.Has("type") ? p["type"] : ""

	switch propType {
		case "title", "rich_text":
			return _NotionAPI_richToText(p[propType])

		case "number":
			return _NotionAPI_orBlank(p["number"])

		case "checkbox":
			return p["checkbox"] ? 1 : 0

		case "select", "status":
			v := p[propType]
			return (v is Map && v.Has("name")) ? v["name"] : ""

		case "multi_select":
			out := []
			for opt in p["multi_select"]
				out.Push(opt["name"])
			return out

		case "date":
			v := p["date"]
			return (v is Map && v.Has("start")) ? v["start"] : ""

		case "url", "email", "phone_number":
			return _NotionAPI_orBlank(p[propType])

		case "people":
			out := []
			for person in p["people"]
				out.Push(person.Has("name") ? person["name"] : person["id"])
			return out

		case "relation":
			out := []
			for rel in p["relation"]
				out.Push(rel["id"])
			return out

		case "formula":
			f := p["formula"]
			ft := f.Has("type") ? f["type"] : ""
			if (ft = "date")
				return _NotionAPI_propObjectValue(Map("type", "date", "date", f.Has("date") ? f["date"] : ""))
			return (ft != "" && f.Has(ft)) ? _NotionAPI_orBlank(f[ft]) : ""

		case "rollup":
			r := p["rollup"]
			rt := r.Has("type") ? r["type"] : ""
			if (rt = "array") {
				; Each element looks like a page property ({type, <type>: value}), so
				; read it the same way; flatten any element that itself gives an
				; Array (multi-select, relation, ...) and drop empty ones.
				out := []
				for el in r["array"] {
					v := _NotionAPI_propObjectValue(el)
					if (v is Array) {
						for item in v
							out.Push(item)
					} else if (!IsObject(v) && Trim(v . "") != "") {
						out.Push(v)
					}
				}
				return out
			}
			if (rt = "date")
				return _NotionAPI_propObjectValue(Map("type", "date", "date", r.Has("date") ? r["date"] : ""))
			return (rt != "" && r.Has(rt)) ? _NotionAPI_orBlank(r[rt]) : ""

		default:
			return (propType != "" && p.Has(propType)) ? p[propType] : ""
	}
}


; Internal: quote/escape a single CSV field per RFC 4180. The field is wrapped
; in double quotes only when it contains the delimiter, a double quote, or a
; line break; embedded double quotes are doubled.
_NotionAPI_csvField(value, delimiter) {
	s := value . ""   ; coerce numbers to string
	if (InStr(s, delimiter) || InStr(s, '"') || InStr(s, "`n") || InStr(s, "`r"))
		return '"' . StrReplace(s, '"', '""') . '"'
	return s
}


; Internal: serialize an AHK value to a compact JSON string.
;   Map        -> object       Array      -> array
;   String     -> quoted str   Integer/Float -> number
;   NotionAPI_bool() -> true/false   NotionAPI_null() -> null
; A plain object's own properties are serialized as a JSON object.
_NotionAPI_jsonStringify(v) {
	if !IsObject(v) {
		if (v is Integer || v is Float)
			return v . ""
		return '"' . _NotionAPI_jsonEscape(v) . '"'
	}

	if (v.HasOwnProp("isBool"))
		return v.val ? "true" : "false"
	if (v.HasOwnProp("isNull"))
		return "null"

	if (v is Map) {
		parts := []
		for k, val in v
			parts.Push('"' . _NotionAPI_jsonEscape(k) . '":' . _NotionAPI_jsonStringify(val))
		return "{" . _NotionAPI_join(parts, ",") . "}"
	}

	if (v is Array) {
		parts := []
		for val in v
			parts.Push(_NotionAPI_jsonStringify(val))
		return "[" . _NotionAPI_join(parts, ",") . "]"
	}

	; Fallback: treat any other object's own properties as a JSON object
	parts := []
	for k, val in v.OwnProps()
		parts.Push('"' . _NotionAPI_jsonEscape(k) . '":' . _NotionAPI_jsonStringify(val))
	return "{" . _NotionAPI_join(parts, ",") . "}"
}


; Internal: escape a string for inclusion in JSON. Backslash must be replaced
; first.
_NotionAPI_jsonEscape(s) {
	s := StrReplace(s, "\", "\\")
	s := StrReplace(s, '"', '\"')
	s := StrReplace(s, "`n", "\n")
	s := StrReplace(s, "`r", "\r")
	s := StrReplace(s, "`t", "\t")
	s := StrReplace(s, Chr(8), "\b")
	s := StrReplace(s, Chr(12), "\f")
	return s
}


; Internal: join an Array of strings with a separator (AHK v2 has no built-in
; join).
_NotionAPI_join(arr, sep) {
	out := ""
	for i, v in arr
		out .= (i > 1 ? sep : "") . v
	return out
}


; Internal: parse a JSON string into AHK values (object -> Map, array -> Array,
; true/false -> 1/0, null -> ""). Throws an Error on malformed input.
;
; PERFORMANCE NOTE (do not "simplify" this away): every _NotionAPI_json* helper
; takes the JSON text BY REFERENCE (&s) and keeps only the read position in
; the small ctx object. An earlier version stored the text on ctx (ctx.s) and
; read characters with SubStr(ctx.s, ctx.i, 1) -- but in AutoHotkey v2, reading
; a string out of an object property copies the ENTIRE string every time, so
; each character read cost time proportional to the whole document and total
; parse time grew with the SQUARE of its size (a multi-MB document took hours;
; the same data in 100-row API pages, parsed separately, was tolerable).
; Passing the text by reference costs nothing, making the parse linear.
_NotionAPI_jsonParse(str) {
	ctx := { i: 1 }
	val := _NotionAPI_jsonValue(&str, ctx)
	_NotionAPI_jsonWs(&str, ctx)
	return val
}


; Internal: read the next JSON value at the current position.
_NotionAPI_jsonValue(&s, ctx) {
	_NotionAPI_jsonWs(&s, ctx)
	c := SubStr(s, ctx.i, 1)

	switch c {
		case "{":
			return _NotionAPI_jsonObject(&s, ctx)
		case "[":
			return _NotionAPI_jsonArray(&s, ctx)
		case '"':
			return _NotionAPI_jsonString(&s, ctx)
	}

	if (SubStr(s, ctx.i, 4) == "true") {
		ctx.i += 4
		return 1
	}
	if (SubStr(s, ctx.i, 5) == "false") {
		ctx.i += 5
		return 0
	}
	if (SubStr(s, ctx.i, 4) == "null") {
		ctx.i += 4
		return ""
	}

	return _NotionAPI_jsonNumber(&s, ctx)
}


; Internal: parse a JSON object into a (case-sensitive) Map.
_NotionAPI_jsonObject(&s, ctx) {
	obj := Map()
	ctx.i++                       ; skip "{"
	_NotionAPI_jsonWs(&s, ctx)
	if (SubStr(s, ctx.i, 1) == "}") {
		ctx.i++
		return obj
	}
	loop {
		_NotionAPI_jsonWs(&s, ctx)
		key := _NotionAPI_jsonString(&s, ctx)
		_NotionAPI_jsonWs(&s, ctx)
		if (SubStr(s, ctx.i, 1) != ":")
			throw Error("JSON parse error: expected ':' at position " . ctx.i, -1)
		ctx.i++                   ; skip ":"
		obj[key] := _NotionAPI_jsonValue(&s, ctx)
		_NotionAPI_jsonWs(&s, ctx)
		c := SubStr(s, ctx.i, 1)
		ctx.i++                   ; skip "," or "}"
		if (c == "}")
			break
		if (c != ",")
			throw Error("JSON parse error: expected ',' or '}' at position " . (ctx.i - 1), -1)
	}
	return obj
}


; Internal: parse a JSON array into an Array.
_NotionAPI_jsonArray(&s, ctx) {
	arr := []
	ctx.i++                       ; skip "["
	_NotionAPI_jsonWs(&s, ctx)
	if (SubStr(s, ctx.i, 1) == "]") {
		ctx.i++
		return arr
	}
	loop {
		arr.Push(_NotionAPI_jsonValue(&s, ctx))
		_NotionAPI_jsonWs(&s, ctx)
		c := SubStr(s, ctx.i, 1)
		ctx.i++                   ; skip "," or "]"
		if (c == "]")
			break
		if (c != ",")
			throw Error("JSON parse error: expected ',' or ']' at position " . (ctx.i - 1), -1)
	}
	return arr
}


; Internal: parse a JSON string literal (assumes current char is the opening ").
_NotionAPI_jsonString(&s, ctx) {
	if (SubStr(s, ctx.i, 1) != '"')
		throw Error("JSON parse error: expected string at position " . ctx.i, -1)
	ctx.i++                       ; skip opening quote

	; Fast path: if the text up to the next quote contains no backslash, that
	; run IS the whole string (no escapes to decode), so take it in one slice
	; instead of appending it one character at a time. (A quote preceded by a
	; backslash means the run contains a backslash, so it correctly falls
	; through to the character-by-character path below.)
	close := InStr(s, '"', true, ctx.i)
	if (close) {
		run := SubStr(s, ctx.i, close - ctx.i)
		if !InStr(run, "\") {
			ctx.i := close + 1
			return run
		}
	}

	out := ""
	loop {
		c := SubStr(s, ctx.i, 1)
		if (c = "")
			throw Error("JSON parse error: unterminated string.", -1)
		if (c == '"') {
			ctx.i++
			break
		}
		if (c == "\") {
			ctx.i++
			e := SubStr(s, ctx.i, 1)
			switch e {
				case '"':
					out .= '"'
				case "\":
					out .= "\"
				case "/":
					out .= "/"
				case "b":
					out .= Chr(8)
				case "f":
					out .= Chr(12)
				case "n":
					out .= "`n"
				case "r":
					out .= "`r"
				case "t":
					out .= "`t"
				case "u":
					hi := Integer("0x" . SubStr(s, ctx.i + 1, 4))
					ctx.i += 4
					; Combine a UTF-16 surrogate pair into one code point
					if (hi >= 0xD800 && hi <= 0xDBFF && SubStr(s, ctx.i + 1, 2) == "\u") {
						lo := Integer("0x" . SubStr(s, ctx.i + 3, 4))
						ctx.i += 6
						out .= Chr(0x10000 + ((hi - 0xD800) << 10) + (lo - 0xDC00))
					} else {
						out .= Chr(hi)
					}
				default:
					throw Error("JSON parse error: invalid escape \" . e, -1)
			}
			ctx.i++
		} else {
			out .= c
			ctx.i++
		}
	}
	return out
}


; Internal: parse a JSON number (returns an AHK Integer or Float).
_NotionAPI_jsonNumber(&s, ctx) {
	start := ctx.i
	loop {
		c := SubStr(s, ctx.i, 1)
		if (c = "" || !InStr("+-0123456789.eE", c))
			break
		ctx.i++
	}
	token := SubStr(s, start, ctx.i - start)
	if (token = "")
		throw Error("JSON parse error: unexpected character at position " . start, -1)
	return Number(token)
}


; Internal: advance past JSON whitespace.
_NotionAPI_jsonWs(&s, ctx) {
	loop {
		c := SubStr(s, ctx.i, 1)
		if (c = " " || c = "`t" || c = "`n" || c = "`r")
			ctx.i++
		else
			break
	}
}
