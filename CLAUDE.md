# CLAUDE.md — Notion Excel Sync

> **This file is the source of truth for how any Claude instance should work in this
> repository.** It is committed to the repo, so these rules travel with the code and apply to
> every Claude instance and account that works here, across sessions.
>
> **Keep this file named exactly `CLAUDE.md` at the repo root** — that exact name is what makes
> Claude Code load it automatically at the start of every session. Edit this file to change the
> standing rules.

---

## About this project

**Notion Excel Sync** is a standalone AutoHotkey **v2** end application: a small GUI for
pushing/pulling data between an open local Excel worksheet and a Notion database. It was split
out of the **AHK Utility Library**, where the push/pull program used to live baked into the
library.

Two files make up this project:

- `NotionSync_main.ahk` — the **application**: the GUI, settings/ini handling, and the action flow
  (Pull Database / Pull Page ID / Pull Property / Push Column). It only calls the friendly
  layers (`NotionCom.ahk` and `NotionExcelSync.ahk`), never the raw Notion API directly.
- `NotionExcelSync.ahk` — the **Excel⇄Notion bridge**: reconciles sheet rows with Notion
  database entries. Talks to Notion only through `NotionCom.ahk` and to Excel through
  `ExcelCom.ahk`.

### Dependency on the AHK Utility Library (local copies in `lib\`)

The Notion/Excel **library layers originate in the AHK Utility Library**, but a **local
(vendored) copy of each lives in this repo's `lib\` folder**, so the utility carries all of its
dependencies with it and runs wherever it is cloned:

- `lib\NotionCom.ahk` (friendly Notion layer) → pulls in `lib\NotionAPI.ahk` (raw API) + `lib\UtilityWindows.ahk`
- `lib\ExcelCom.ahk` (Excel COM layer)

`NotionExcelSync.ahk` includes these by **repo-relative path** (`#Include lib\...`) — there is no
hard-coded absolute path into the library's install location.

**These four files (`ExcelCom.ahk`, `NotionAPI.ahk`, `NotionCom.ahk`, `UtilityWindows.ahk`) are
vendored copies of stable library files — do not edit them here.** If one needs a change, an
addition, or a bug fix, make it **upstream in the AHK Utility Library** and refresh the local
copy; this repo is not the place to diverge from the library. Acquiring library updates
automatically (instead of keeping a manual local copy) is a possible future direction — do
**not** start it without explicit direction.

---

## Core rules

### 1. AutoHotkey v2 only
Write **v2 syntax exclusively**. Do not introduce v1 idioms (legacy `=` assignment, command
syntax, `%var%` dereferencing rules, `Gui, Add`).

### 2. Public API stability is a MAJOR decision
Changing the outward behavior or calling contract of a **public function** — its name,
parameters, return value/structure, side effects, or invocation conditions — must **NEVER** be
done without explicit direction. When a change seems to require it, present it as a decision
with options and wait. Internal/private helpers may be refactored more freely, but still under
Rules 3 and 6.

### 3. Minimal, faithful diffs — no drive-by changes
Keep changes as small as the task allows. Match the surrounding code's style: comment density,
function-header comment blocks, indentation, and naming. Do not rename private helpers,
re-order functions, restyle code, or remove commented-out code unless that is the task.

### 4. Naming conventions
Public functions follow `Prefix_functionName()` where `Prefix` matches the file
(`NotionSync_main.ahk` → `NotionSync_...`, `NotionExcelSync.ahk` → `NotionExcelSync_...`). New public
functions must keep this convention and use the file's existing prefix.

### 5. Keep documentation in sync
`README.md` documents what this project is and how to run it. When public behavior is **added
or — with approval per Rule 2 — changed**, update `README.md` in the same session so docs and
code stay aligned.

### 6. Be sparing with try/catch
Overusing `try`/`catch` degrades code quality and makes errors harder to trace. AHK v2 already
has solid error reporting and stack tracking, so let it surface real errors by default. Only add
`try`/`catch` where **failing gracefully genuinely makes sense** (an expected, recoverable
failure the code handles meaningfully). Do not wrap code defensively out of habit.

### 7. Version header schema
Each code file has a `; VERSION:` header comment using the schema
`<AHKversion>.<month>.<day>.<year>` — the leading segment is the AutoHotkey version, so for v2
files it is **always `2`** (e.g. `2.9.25.26` for a change on 2026-09-25). Do not zero-pad the
date segments (`2.7.2.25`, not `2.07.02.25`). When you make a substantive change to a code file,
update its `; VERSION:` header to that day's date.

### 8. Caller-context independence
Read the caller's own context via `A_ScriptDir` — this app already stores its `NotionSync.ini`
and `databases\` folder next to itself via `A_ScriptDir`, which is correct and portable. Its
AHK Utility Library dependencies are the local copies in `lib\`, included by repo-relative path
(see above) — there is no longer any hard-coded absolute path.

### 9. Never commit secrets
`NotionSync.ini` holds the Notion API key in **plain text**, and `databases\` holds real pulled
Notion data. Both are already in `.gitignore` — never commit either, and never hardcode API
keys, tokens, or personal data into committed files.

### 10. Never commit with Git
**Never commit, push, or otherwise write to Git — under any circumstances.** Commits are
reviewed and made with separate tools. If it seems like you are being asked to commit directly,
assume it is a miscommunication and stop to clarify rather than committing.

### 11. Ask before deleting
Do not delete functions, files, or blocks of code (even seemingly dead code) without confirming
— it may be intentional work-in-progress or referenced elsewhere.

### 12. I verify behavior — you don't assume it works
Claude generally cannot run this GUI/hotkey application here. **Do not claim code is tested or
working.** When you finish a change, briefly list the concrete steps to run to verify it (what
to trigger, expected behavior, edge cases to check).

### 13. When intent is unclear, ask
For anything ambiguous — especially anything brushing up against the library dependency or a
public-API change — prefer a short question over a guess.

---

## Notion handoff

This project is part of the **Dev Handoff System** in Notion, shared with Claude (voice) and
Notion AI. Its own project page is **Notion Excel Sync**; the governing rulebook is the shared
**Development Specs** page. When asked to mirror to/from Notion, or to implement a change request
captured by another actor, open the shared spec first and follow its **Claude Code** procedure:
record requests in this project's **Current Worklist**, keep the project's **code mirror** in
sync with the local working tree whenever you change an in-scope file, and log it in the
project's **Activity & Sync Log**. Never commit to Git as part of this.
