# Changelog

## 0.7.36 — 2026-05-25

Five new capabilities. Net tool count 62 → 65 (3 new tools; macros and the Office-version autodetect refactor are non-additive). All changes are additive — no existing schema or function signature was modified.

### Added
- **`access_search_data`** — search any text string across Text/Memo fields of any local table in a single call. Skips system tables (`MSys*`, `~*`) and linked tables (querying remote SQL servers with `LIKE` per column is rarely what the caller wants). Per-table and total caps, `match_case`, optional `tables` whitelist. Returns matches grouped by table with an `_excerpt` around each hit. See `sql.py:ac_search_data`.
- **`access_clone_object`** — duplicate a form, report, module, class_module, query or macro to a new name. Internally `SaveAsText` → `LoadFromText` with binary sections (PrtMip / PrtDevMode / NameMap / GUID) preserved by using a raw read path that bypasses `strip_binary_sections`. VBA code-behind comes along via the existing `ac_set_code` injection. Refuses to overwrite unless `overwrite=true`. See `code.py:ac_clone_object`.
- **`access_manage_tab_order`** — get / set / auto_renumber the `TabIndex` of controls on a form or report. `get` returns controls grouped by section, sorted by TabIndex. `set` assigns 0..N-1 in the order of `tab_order` (two-phase write to avoid the unique-index collision Access enforces per section). `auto_renumber` re-sequences existing TabIndex values per section. Skips controls that don't support TabIndex (Label, Line, Rectangle, Image, PageBreak, Page). Optional `section` filter. See `controls.py:ac_manage_tab_order`.

### Changed
- **Macros: docs and tips upgrade only — no new tool**. Macros were already fully supported via `access_get_code` / `access_set_code` (UTF-16 encoding correctly applied for SaveAsText/LoadFromText), `access_list_objects`, `access_run_macro`, `access_delete_object`. Tool descriptions now call macros out explicitly, and `access_tips('macros')` documents the read → edit → write workflow.
- **Office version autodetect**. The hardcoded `16.0` / `Office16` references in `_Session._suppress_recovery_dialog`, `_Session._decompile` and `maintenance.ac_decompile_compact` are now driven by a one-shot registry probe (`_Session._detect_office_install`). Detection enumerates `Software\Microsoft\Office\<ver>\Access\InstallRoot\Path` under HKLM, HKLM\\WOW6432Node and HKCU (per-user Click-to-Run) and picks the highest version with a working `MSACCESS.EXE`. Falls back to `App Paths\MSACCESS.EXE\(Default)` and finally to the previous hardcoded defaults — so machines with a normal Office install keep working unchanged, and machines with a different major version (15.0 / 14.0) or a non-default install root start working without manual edits. Schema of `access_decompile_compact` is unchanged.

### Notes
- Test passing the `access_clone_object` overwrite path on a class_module — the `_ensure_class_module_header` re-injection runs after the raw export.
- `_detect_office_install` never raises; the worst case logs `Could not detect Office install via registry — using hardcoded defaults` and behaviour is identical to v0.7.35.

## 0.7.35 — 2026-05-25

Preventive bug sweep across the codebase. No reported regressions, but several latent issues fixed.

### Fixed (critical)
- **`_Session.quit()`** captured the Access PID by calling `app.hWndAccessApp()` from the `atexit` thread, but the COM proxy was created on the COM worker — the cross-thread call returned silently and `pid` was `None`, disabling the `taskkill` fallback added recently. PID is now captured in `_launch()` on the COM worker and stored as `cls._pid`.
- **`coerce_arguments`** only widened scalar argument types — clients that serialize arrays/objects as JSON strings (common with some MCP transports) failed on every batch tool. The coerce now JSON-parses string-typed arrays/objects, and the schema fixup also widens `number`, `array`, `object`. Boolean coercion also accepts `on`/`y`/`si`/`sí`.
- **`restore_binary_sections`** matched the first `End` after `Begin Form` as the injection point. Forms with embedded subforms (their own nested `Begin Form ... End`) had `PrtMip` / `PrtDevMode` / `NameMap` injected inside the subform, corrupting the result on `ac_set_code`. Now tracks full block depth and injects at the outermost End.
- **`write_tmp`** used `errors="replace"` for ANSI codepage writes (`.bas` modules) — a non-cp1252 character (emoji, asian, `✓`) was silently replaced with `?`. Now `errors="strict"` and the resulting `UnicodeEncodeError` carries a concrete snippet of the offending text.
- **`access_compile_vba`** accepted a `timeout` parameter but never used it. Now controls the watchdog grace window (default 2s, clamped to 1–30).

### Fixed (medium)
- VBE read operations (`ac_vbe_get_lines`, `get_proc`, `module_info`, `find`, `search_all`) now close the form/report Design view before reading the CodeModule. Skipping this could surface as `Catastrophic failure` (-2147418113) when the same object was open in design mode.
- `ac_vbe_replace_lines` no longer calls `cm.DeleteLines(start, 0)` when the count clamps to zero (raised in VBE); error message now lists separate upper bounds for replace/delete vs pure insert.
- `ac_vbe_patch_proc` normalizes `find_text`/`replace_text` line endings to CRLF before the exact-match check (callers commonly send LF and were always falling through to the ws-normalized fallback), and warns when `find_text` appears more than once (only the first occurrence is replaced).
- `_proc_kind` no longer silently picks the first matching kind when a procedure name resolves to multiple kinds (a class with both `Property Get Foo` and `Property Let Foo` is normal VBA). It raises a descriptive error so the caller can disambiguate.
- `set_db_property` / `set_field_property` infer `dbDouble`, `dbDate`, `dbSingle`, `dbMemo` for float/datetime/long-string values. Previously these fell to `dbText` (stored as string).
- `_eval_via_temp_module` pre-binds `temp_name` to avoid `UnboundLocalError` in the cleanup `finally` if creation fails before `comp.Name` is read.
- Compile watchdog captures up to 3 dialog screenshots / texts and the caller picks the last one — the first dialog is often a benign "Save changes?" and the real compile error came last (and used to be discarded).
- `ac_create_relationship` validates that local/foreign fields exist on the referenced tables before `Append`, so the error names the missing field instead of a cryptic DAO message.
- `_check_module_health` and `module_info` regex now recognize `Public Static Sub`/`Function`/`Property`.
- `decompile_compact` resets `_Session._pid` and `_attached` when killing the spawned process, keeping the `quit()` fallback consistent.

### Fixed (low / hardening)
- `read_tmp` tries UTF-8 before cp1252 (cp1252 single-byte never raises and was masking real UTF-8 files as mojibake).
- `_invoke_app_run` validates `len(args) <= 30` explicitly instead of producing a confusing InvokeTypes failure via negative-multiplier padding.
- `_split_code_behind` matches `CodeBehindForm` only at the start of a line to avoid false positives from property values containing that literal.
- `_SQL_LINE_COMMENT` / `_SQL_BLOCK_COMMENT` removed (dead code) — the destructive guard already uses `_sql_effective_prefix`. `SELECT … INTO` (make-table) now flagged as destructive.
- `ac_create_database` rejects paths without `.accdb` / `.mdb` extension.
- `relink_table` UID/PWD detection uses a parameter-boundary regex (`(^|;)UID=`), not a substring check.
- Linked-table count query escapes `]` in table names by doubling.
- `compact_repair` cleans up the orphaned `_compact_tmp.accdb` if rollback succeeds.
- Safe-args logging in `server.py` / `dispatcher.py` guards against non-string `code`.
- `tools.py` docstring now reports 62 tools (was 58).

## 0.7.34 — 2026-05-05

### Fixed
- **`access_list_controls`** silently lost controls inside `Page` / `OptionGroup` containers when any earlier control in the same Page had a multi-line property block (`GUID = Begin … End`, `NameMap = Begin … End`, `ConditionalFormat = Begin … End`, etc.). The depth counter inside a control's body matched plain `Begin <Type>` but not `Property = Begin`, so the property's closing `End` was decremented without ever being incremented — the enclosing control was closed prematurely, and every control that came after it inside the Page was never enumerated. The form-level loop already handled this; the per-control loop now mirrors it.

  Visible symptom: `access_list_controls` reported a TabControl Page as a 15-line empty stub even though the Page actually contained dozens of controls. Fixed in `mcp_access/controls.py:_parse_controls`.
