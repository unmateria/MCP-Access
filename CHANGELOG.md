# Changelog

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
