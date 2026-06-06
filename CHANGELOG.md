# Changelog

## 0.7.42 — 2026-06-06

VBE procedure-editing fixes from field reports (thanks to
[@TvanStiphout-Home](https://github.com/TvanStiphout-Home)). No new tools — tool
count stays **66**.

### Fixed

- **`access_vbe_replace_proc` no longer eats the blank separator line above a
  procedure.** `ProcStartLine` is the previous proc's `End` + 1, so it *includes*
  the blank line VBE attributes to the proc; the old code deleted from there and
  re-inserted code with no leading blank, consuming the separator on every
  replace. Replaces now preserve the leading blank line(s) — delete/insert happen
  below them. A pure delete (`new_code=''`) still removes the whole range incl.
  the leading blank (so deleting a proc closes its gap cleanly).
- **No more spurious "Option … expected in first 5 lines" warning on modules
  with a long comment header** (e.g. a banner block pushing `Option Compare` past
  line 5). The structural health check replaced its fixed line-number threshold
  with a rule that flags an Option statement only when real (non-comment,
  non-blank) code already appeared above it — still catching genuinely misplaced
  Option statements.

### Added

- **`access_vbe_replace_lines` accepts `new_lines`** (a list of strings, joined
  with `\n`; `''` entries become blank lines) as an alias for `new_code`. This
  closes a destructive footgun: a call that passed the code under a wrong key
  left `new_code` empty and silently degraded into a pure delete. A replace that
  deletes lines but inserts nothing now also appends an explicit note to its
  result, so a destructive no-op is never silent.

### Changed

- `access_vbe_get_proc` / `access_vbe_module_info` descriptions and docstrings
  now spell out `start_line` (VBE proc start — includes preceding blank/comment
  lines) vs `body_line` (the `Sub`/`Function`/`Property` declaration line).
  `access_tips('vbe')` documents both, the separator-preserving replace, and
  `new_code=''` deletion.

### Tests

- `tests/test_vbe_fixes.py` — 11 COM-free tests for the Option-placement check
  and the `new_lines` alias normalisation. Blank-separator preservation and the
  end-to-end `new_lines` path were verified with a live COM integration run.

## 0.7.41 — 2026-05-29

Adds **deterministic UI design validation** so the assistant stops accepting
objectively broken form/report layouts. **65 → 66 tools.**

### Added

- **`access_lint_form`** — pure-Python, rules-based lint of a form/report.
  Returns structured JSON violations with `summary.verdict` (PASS / REVIEW /
  FAIL) and per-violation `suggested_fix`. Rules: `contrast` (WCAG 2.1
  ratio — catches white-on-white & low-contrast text), `overlap`,
  `out_of_bounds`, `truncation`, `sibling_inconsistency`, `misalignment`,
  `invisible_or_zero_size`. Static — one `SaveAsText` export, never opens
  Design view. `measure="auto"|"wizhook"|"heuristic"` (WizHook gives exact
  rendered text width when the VBA project is compiled; otherwise a
  conservative heuristic).
- **Embedded enforcement.** `access_set_control_props`,
  `access_set_multiple_controls` and `access_create_control` now attach a
  compact `lint` block (errors + warnings) to their result **automatically**.
  The validation is deterministic and lives entirely inside the MCP — it
  cannot be skipped or "talked past" by the model. `skip_lint=true` opts out
  for bulk programmatic edits. A lint failure never breaks the mutation.
- `access_tips('lint')` documents the rules, thresholds and colour encoding.
- Unit tests: `tests/test_lint.py` (29 COM-free tests for colour decoding,
  WCAG contrast, geometry parsing, and every rule incl. false-positive guards).

### Notes

False-positive guards, hardened against a real 85-control ERP form (findings
dropped 62 → 10, remainder genuine):

- **Conditional formatting** overrides colours at runtime (binary in the export)
  → contrast skips + notes those controls.
- **Captions wrap** — Labels *and* CommandButtons; line breaks (`\015\012`) are
  split, truncation compares wrapped-line count vs lines that fit the height.
- **`sibling_inconsistency` clusters** values: two legitimate sizes (main vs
  inline buttons) are both accepted; only a lone outlier flags.
- **Transparent buttons** stacked on styled labels (the custom-button pattern)
  are not flagged as overlaps.
- **Heuristic width is calibrated** for narrow UI fonts (≈0.46×) and only flags
  overflow past 1.25× — bold header labels that fit are no longer flagged.
- Absent dimensions inherit form defaults (not zero); attached labels,
  cross-tab-page controls, container Pages and transparent layering never count
  as overlaps; icon buttons aren't measured for caption truncation.
- Access auto-grows form Width / section Height to fit controls, so horizontal
  `out_of_bounds` rarely fires for forms (still effective for reports and
  negative coordinates).

## 0.7.40 — 2026-05-29

Fixes an indefinite hang on databases whose **startup form raises a blocking
modal** during open / VBE access.

### Fixed

- **Global dialog watchdog.** Until now the dialog-dismiss watchdog only ran
  during `open` / `compile` / `run_vba`. Operations that access the VBE
  (`vbe_get_proc`, `find_definition`, `module_info`, ...) had **no** watchdog,
  so a modal raised by a DB's startup form — e.g.
  `"Error accessing file. Network connection may have been lost."` on a DB with
  `StartupForm` set — would hang the COM call forever (observed: a ~1-hour hang
  on one such database). `_Session` now starts a background watchdog in
  `_launch()` that dismisses Access-owned `#32770` / wizard dialogs which
  persist past a 3 s grace period, for the whole lifetime of the spawned Access
  process. The grace period lets operation-specific watchdogs handle (and
  screenshot) their own dialogs first; the global thread only backstops
  un-watched operations. It is **not** started when attached to an existing
  interactive user session.
- Factored the dialog enumeration into `_find_dialog_hwnds_by_pid()` (shared by
  the dismisser and the new watchdog).

## 0.7.39 — 2026-05-28

Hardening of the v0.7.38 `_looks_like_vba_only` detector. No behaviour change
for the well-formed cases that v0.7.38 already handled.

### Fixed

- **VBA comments containing "Begin Form" or "Version =" no longer
  misclassify pure VBA as a form export.** The previous detector ran
  `_FORM_EXPORT_RE.search(code)` over the whole text, so a comment like
  `' Begin Form: this sub opens it` made `_looks_like_vba_only` return
  False — sending the file through `LoadFromText` instead of VBE
  injection, which would then fail with "errors while importing". The
  new detector only inspects the first non-blank line for `Version =NN`
  and the first 20 lines for `Option Compare` / Sub/Function/etc.,
  matching how Access actually emits form text exports (the `Version`
  declaration is the very first line of any SaveAsText output).
- `_VBA_HINT_RE` now also matches `Public Static Sub` / `Public Static
  Function` (a real-world VBA pattern used by counters and singletons).

## 0.7.38 — 2026-05-28

DX fixes for `access_set_code` on freshly-created forms, `access_create_control`
with TabControl-Page parents, and the cryptic VBE error you got when the form
had no code module yet. All real-world tripping points hit while building
`frmSugerirPedido` for an ERP — see notes below for the actual reproductions.

### Fixed

- **`ac_set_code(form|report)`** failed on forms recently created via
  `ac_create_form` because `LoadFromText` was always invoked even when the
  caller passed pure VBA (`Option Compare Database` + `Private Sub …`). Pure
  VBA isn't a valid form text export, and the binary-section restoration only
  works against an already-exported baseline — so the import raised `errors
  while importing` and rolled the form back. The same code now detects
  VBA-only input and routes it through `_inject_vba_after_import` (open in
  Design view → activate `HasModule` → write via VBE), preserving the form
  layout and never touching `LoadFromText`. A full form export (containing
  `Version =` / `Begin Form`) still takes the original `LoadFromText` path.
  See `code.py:_looks_like_vba_only`.

- **`ac_vbe_module_info` / any VBE read on a brand-new form** raised
  `Subscript out of range` with a misleading error mentioning *"Trust access to
  the VBA project object model"*. The actual cause: `HasModule=False` on a
  form just made by `ac_create_form`, so `VBComponents("Form_xxx")` had nothing
  to return. `_force_vbe_init` now flips `HasModule` on when the form/report is
  opened in Design view before retrying. The fallback error message is also
  rewritten so it no longer blames Trust Center first when the obvious cause
  is a missing module. See `vbe.py:_force_vbe_init`.

- **`ac_create_control` rejected `Parent`** (capital P) with `Property
  'CreateControl.Parent' can not be set` because special keys were popped from
  `props` case-sensitively. Hand a control to a TabControl Page using
  `{"parent": "myTab", ...}` or `{"Parent": "myTab", ...}` — both work now.
  Same case-insensitive treatment for `section`, `column_name`, `left`, `top`,
  `width`, `height`. See `controls.py:_pop_ci`.

- **`ac_create_control` lost properties Access exposes only via the
  `Properties` collection.** `setattr(ctrl, "ScrollBars", 2)` raises for some
  control types even when `ctrl.Properties("ScrollBars").Value = 2` succeeds.
  The loop now retries via the Properties collection before recording an
  entry in `property_errors`. Properties that don't exist at all on a given
  control type (e.g. the `MultiLine` UserForm property on an Access TextBox)
  still fail loudly — those are legitimate user errors.

### Added

- **`ac_create_control` accepts `control_name` at the top level** (in addition
  to `{"Name": "..."}` inside `props`). Without this you had to discover that
  the control was auto-named `Command1` / `Label2` and rename it via
  `set_control_props` in a second round-trip. `props["Name"]` still wins if
  both are provided, so existing callers don't change behaviour.

### Notes

- No schema changes to existing required fields — `control_name` is optional,
  the `Parent` fix is silent, and `set_code` only takes the new path when the
  caller passes VBA-only code to an existing form/report. Old fixtures that
  pass complete form exports continue using `LoadFromText`.
- The CLAUDE.md *Recipes* section gained two entries documenting the new flows
  for callers who want to build forms from scratch in one MCP session.

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
