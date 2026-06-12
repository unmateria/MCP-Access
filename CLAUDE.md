# CLAUDE.md — mcp-access MCP Server

## Overview

MCP server for reading and editing Microsoft Access databases (`.accdb`/`.mdb`) via COM automation (pywin32). Runs as stdio MCP server. Entry point: `access_mcp_server.py`. Implementation: `mcp_access/` package (~7500 lines across 20 modules).

## Architecture

- **Singleton COM session** (`_Session`): one `Access.Application` instance shared across all tool calls. Opening a different `.accdb` closes the previous one.
- **Dedicated COM thread** (`_com_executor`): All tool calls run in a single-threaded `ThreadPoolExecutor` with `CoInitialize()`. This keeps COM in one STA thread while the asyncio event loop stays free to read/write stdio.
- **Caches**: `_parsed_controls_cache` (control parsing) and `_Session._cm_cache` (CodeModule COM objects — live COM proxies). Both invalidated on DB switch, object modification, and design operations. There is **no** Python-side cache of VBE text: `_cm_all_code()` always reads via `cm.Lines(1, total)` so external edits (manual VBE edits, Ctrl+Z, add-ins) are picked up immediately. See issue #26 for the reason this cache was removed.
- **Binary section handling**: `ac_get_code` strips PrtMip/PrtDevMode from form/report exports; `ac_set_code` restores them automatically before import.

## Tools (66 total)

| Category | Tools |
|----------|-------|
| **Database** | `access_create_database`, `access_close` |
| **Objects** | `access_list_objects`, `access_get_code`, `access_set_code`, `access_export_structure`, `access_delete_object`, `access_create_form`, `access_clone_object` |
| **SQL/Tables** | `access_execute_sql`, `access_execute_batch`, `access_table_info`, `access_search_queries`, `access_search_data`, `access_create_table`, `access_alter_table` |
| **VBE line-level** | `access_vbe_get_lines`, `access_vbe_get_proc`, `access_vbe_module_info`, `access_vbe_replace_lines`, `access_vbe_find`, `access_vbe_search_all`, `access_vbe_replace_proc`, `access_vbe_patch_proc`, `access_vbe_append` |
| **Controls** | `access_list_controls`, `access_get_control`, `access_create_control`, `access_delete_control`, `access_set_control_props`, `access_set_multiple_controls`, `access_manage_tab_order` |
| **UI lint** | `access_lint_form` |
| **DB Properties** | `access_get_db_property`, `access_set_db_property`, `access_get_form_property`, `access_set_form_property` |
| **Text Export/Import** | `access_export_text`, `access_import_text` |
| **Linked Tables** | `access_list_linked_tables`, `access_relink_table` |
| **Relationships** | `access_list_relationships`, `access_create_relationship`, `access_delete_relationship` |
| **VBA References** | `access_list_references`, `access_manage_reference` |
| **Maintenance** | `access_compact_repair`, `access_decompile_compact` |
| **Screenshot & UI** | `access_screenshot`, `access_ui_click`, `access_ui_type` |
| **Queries** | `access_manage_query` |
| **Indexes** | `access_list_indexes`, `access_manage_index` |
| **VBA Compilation** | `access_compile_vba` |
| **VBA Execution** | `access_run_macro`, `access_run_vba`, `access_eval_vba` |
| **Export** | `access_output_report` |
| **Data Transfer** | `access_transfer_data` |
| **Field Properties** | `access_get_field_properties`, `access_set_field_property` |
| **Startup Options** | `access_list_startup_options` |
| **Cross-reference** | `access_find_usages`, `access_find_definition` |
| **Knowledge base** | `access_tips` |

## Key Implementation Details

### Encoding in ac_set_code
- **Modules** (`.bas`): written using the system ANSI codepage (`locale.getpreferredencoding()`, typically `cp1252` on Western Windows) — no BOM.
- **Forms, reports, queries, macros**: written as `utf-16` (UTF-16LE with BOM) — Access LoadFromText expects this.

### Control parsing (_parse_controls)
The Access export format nests controls inside sections:
```
Begin Form
    Begin                    <- defaults block (NOT controls)
    End
    Begin Section            <- section (Detail, FormHeader, FormFooter)
        Begin                <- container
            Begin Label      <- REAL CONTROL
            End
            Begin Page       <- CONTAINER -- children re-scanned
                Begin        <- anonymous wrapper
                    Begin ComboBox  <- child control (parent = Page)
                    End
                End
            End
        End
    End
End Form
```
**Container types** (`_CONTAINER_TYPES = {"Page", "OptionGroup"}`): re-scanned for child controls. Children get a `"parent"` field. `container_stack` tracks nesting.

**Depth counter inside a control block must include `Property = Begin`** (e.g. `GUID = Begin`, `NameMap = Begin`, `ConditionalFormat = Begin`). These open multi-line blocks closed by their own `End`. If the parser only counts plain `Begin <Type>` it decrements depth on the closing `End` of the property block without ever incrementing — the control closes prematurely at the first such `End`, and any controls that follow inside a `Page` / `OptionGroup` are silently lost. Fixed in v0.7.34 (was: `re.match(r"^Begin\b", bl_s)` — now also matches `r"^\w+\s*=\s*Begin\s*$"`, mirroring the form-level loop).

### VBE + Design view conflict
After design operations (`ac_set_control_props`, `ac_create_control`, `ac_delete_control`), the form may remain open in Design view. All VBE write functions close the form first (DoCmd.Close with acSaveYes), invalidate `_cm_cache`, then access VBE. Without this: `"Catastrophic failure" (-2147418113)`. All design operations invalidate all three caches in their `finally` block.

### VBE procedure kinds
VBE `ProcStartLine`/`ProcBodyLine`/`ProcCountLines`/`ProcOfLine` require a `kind` argument (`_ALL_PROC_KINDS = (0, 1, 2, 3)`): 0=Sub/Function, 1=Property Let, 2=Property Set, 3=Property Get. `_proc_kind()` iterates all four. `ac_vbe_module_info` deduplicates by `(name.lower(), keyword.lower())` so paired Get/Let/Set appear separately.

### Class module format (LoadFromText vs VBE)
Two **different** export/import formats exist:
- **VBE** (`VBComponent.Export/Import`): `VERSION 1.0 CLASS` header + `Attribute VB_Name`. This is `.cls` file format.
- **Access** (`SaveAsText/LoadFromText`): just the 4 `Attribute VB_*` lines (GlobalNameSpace, Creatable, PredeclaredId, Exposed), NO version header.

Passing VBE-style headers to `LoadFromText` creates a corrupt standard module. `_ensure_class_module_header()` strips VBE headers and injects the correct LoadFromText-style attributes.

### Dialog watchdog system
Blocking COM calls (`OpenCurrentDatabase`, `CompactRepair`, `RunCommand`, `Application.Run`) are protected by polling watchdog threads that dismiss Access dialogs via `_dismiss_access_dialogs()` / `_dismiss_dialogs_by_pid()`. Button priority: Cancel > End > OK (Cancel-first avoids advancing wizards).

**Attached-instance policy (v0.7.43)**: the global watchdog also runs when we attached to the user's Access, but it only dismisses dialogs while one of OUR tool calls has been in flight longer than the grace period (5 s attached vs 3 s spawned). `_Session._tool_started` (monotonic timestamp, set/cleared by `server.call_tool` around `run_in_executor`) is the in-flight signal. A modal with no tool call in flight belongs to the interactive user — never touched. Do NOT "simplify" this back to disabling the watchdog on attach: that re-introduces the 1-hour VBE hang when a broken-reference VBA project pops "Error accessing file..." during one of our calls.

**Dismissal note (v0.7.44, issue #31)**: `_dismiss_dialogs_by_pid` (the funnel every watchdog goes through) records `(monotonic, title)` in `_Session._last_dismissed`; `server.call_tool` appends a "a modal dialog (X) was auto-dismissed during this call" note to the tool result when the timestamp falls inside the call. Cancel-first dismissal can alter outcomes (e.g. cancel a save prompt) — the note makes that traceable instead of silent.

**Eval/delete hardening (v0.7.44, issue #31)**:
- `ac_eval_vba` accepts optional `timeout` with the same `_dialog_watchdog` treatment as `ac_run_vba`, wrapping BOTH `Application.Eval` and the temp-module fallback.
- `_eval_via_temp_module` sweeps orphan `_mcp_eval_wrapper` temp modules (`_sweep_orphan_eval_modules`, marker in the first 10 lines of std modules only) before creating a new one — a failed `Remove` used to wedge every later call with "cannot find the procedure 'Module1._mcp_eval_wrapper'". Deliberately NOT in `connect()`: that would cost VBE access (Trust Center + Visible) on every tool call.
- `ac_delete_object` calls `_save_all_modules(app)` (best-effort: `RunCommand 280` = acCmdSaveAllModules, fallback per-module `DoCmd.Save`) before `DoCmd.DeleteObject` to prevent the "save changes to the design of module X?" prompt — dirty state can come from user code run via eval, so per-tool bookkeeping isn't enough. Do NOT extend this to close/quit paths: on attached instances it would silently persist the interactive user's half-finished VBE edits.

### Wedged-session detection (v0.7.43, from PR #30 by @CaptainStormfield)
A DB whose startup code closes it during the open (startup-form error path + `AllowBypassKey=False`) used to leave `_db_open` pointing at a dead database — every later call died at `CurrentDb` and reconnects re-attached to the same broken instance. Now: `_switch()` validates `CurrentDb() is not None` post-open (raises an actionable RuntimeError after `quit()`-resetting the session), `connect()` health-checks `CurrentDb()` whenever `_db_open` is set (auto-reconnect via `_force_cleanup()`), and `ac_create_database` validates its reopen. Cost: one extra `CurrentDb()` COM round-trip per tool call — accepted trade-off.

### Multi-object scans must not lie with "0 matches" (v0.7.43)
`ac_vbe_search_all` / `ac_find_usages` / `ac_find_definition` collect per-object failures into `errors` (capped at `_SEARCH_ERROR_CAP = 20`) + `objects_skipped` + a `warning`, instead of `except: continue`. A VBA project that fails to load (broken reference, Trust Center) makes EVERY object fail — a clean `total: 0` was a false "doesn't exist". Same idea in `ac_list_references`: each reference property is read defensively (broken references raise `com_error` on `FullPath`), never kill the listing.

### Application.Run via InvokeTypes
`Application.Run` has 31 params (1 required + 30 optional). pywin32's late-bound `Dispatch` can't handle this. `_invoke_app_run()` calls `_oleobj_.InvokeTypes()` directly with `pythoncom.Missing` padding. Same approach for `Application.Eval` via `_invoke_app_eval()`.

## Adding a new tool

1. Write the implementation function (e.g. `ac_new_tool()`)
2. Add a `types.Tool(...)` entry to the `TOOLS` list
3. Add an `elif name == "access_new_tool":` branch in `call_tool()`
4. Update the tool count in this CLAUDE.md and README.md

## Office version autodetect (v0.7.36+)

`_Session._office_version` / `_Session._office_msaccess` are populated by `_Session._detect_office_install()` — a one-shot probe that enumerates `Software\Microsoft\Office\<ver>\Access\InstallRoot\Path` under HKLM, HKLM\\WOW6432Node and HKCU (per-user C2R), picks the highest matching version with a working `MSACCESS.EXE`, falls back to `App Paths\MSACCESS.EXE\(Default)`, and finally to the previous hardcoded `16.0` / `Office16` defaults. Used by:
- `_Session._suppress_recovery_dialog` — Resiliency registry key path
- `_Session._decompile` — MSACCESS.EXE path for `/decompile` subprocess
- `maintenance.ac_decompile_compact` — same

Detection is idempotent (`_office_detected` flag) and never raises. Schema of `access_decompile_compact` is unchanged.

## Macros (v0.7.36 docs upgrade)

Macros have always been fully supported via the regular code tools — no dedicated tool exists. The workflow is:
- List: `access_list_objects(object_type="macro")`
- Read: `access_get_code(object_type="macro", name=...)`
- Write: `access_set_code(object_type="macro", name=..., code=...)` — UTF-16 encoded
- Run: `access_run_macro(macro_name=...)`
- Delete: `access_delete_object(object_type="macro", object_name=..., confirm=true)`

`restore_binary_sections` does NOT apply to macros (they have no PrtMip/PrtDevMode). `access_tips('macros')` shows the workflow.

## Clone object (v0.7.36)

`access_clone_object` duplicates an object by raw `SaveAsText` → `LoadFromText`. Critical detail: it does its own `app.SaveAsText` + `read_tmp` directly (≈10 lines duplicated from `code.py:ac_get_code`) — explicitly NOT going through `strip_binary_sections`, so PrtMip / PrtDevMode / NameMap / GUID ride along inside the text. `ac_set_code` then sees the binaries are present and skips restoration (`code.py:348` — only restores when absent). For `class_module`, `_ensure_class_module_header(text, target_name)` re-runs so the implicit VB_Name stays consistent.

## Tab order (v0.7.36)

`access_manage_tab_order` uses **single-pass assignment** in target order — Access enforces TabIndex to be in `0..(N-1)` per section and auto-renumbers the rest to preserve uniqueness when you set one. Do NOT try to "park" controls at indices >= N (Access rejects with "The value you used for the TabIndex property isn't valid. The correct values are from 0 through N-1."). Skips non-tabbable types (100=Label, 101=Rectangle, 102=Line, 103=Image, 114=PageBreak, 118=Page). Optional `section` filter; defaults to all sections.

## UI design lint (v0.7.41)

`mcp_access/lint.py` is a **deterministic, pure-Python** design validator (no
LLM, no external service). `access_lint_form` returns structured JSON
violations; the same engine runs **automatically** on every design mutation.

### Why it exists
The LLM sets control coordinates/colours blind and used to accept objectively
broken layouts (white-on-white, overlap, truncation, inconsistent siblings,
out-of-bounds). The fix the user asked for: validation that lives *inside* the
MCP and **cannot be skipped or "talked past"** by the model. So the rules are
numeric and the result is attached to mutations whether the model asks or not.

### Architecture
- One SaveAsText export (via `ac_get_code`, binary sections already stripped),
  never opens Design view. `_build_model` layers a style dict (`_extract_style`,
  reading the control's `raw_block`) and section assignment onto the cached
  `_parse_controls` result, plus `_parse_geometry` (form Width/BackColor +
  per-section Height/BackColor/line-range).
- Rules: `contrast` (WCAG 2.1, `_decode_bgr`+`_contrast_ratio`), `overlap`
  (AABB, same section+parent only), `out_of_bounds`, `truncation`,
  `sibling_inconsistency`, `misalignment`, `invisible_or_zero_size`.
- `lint_compact()` (errors+warnings, heuristic measure, capped) is attached by
  `_attach_lint` to the result of `ac_set_control_props`,
  `ac_set_multiple_controls`, `ac_create_control`. Wrapped in try/except so a
  lint failure NEVER breaks the mutation. `skip_lint=true` opts out for bulk ops.

### Hard-won gotchas baked into the rules (do NOT "simplify" these away)
- **Absent dimension ≠ 0.** Access omits `Left/Top/Width/Height` (and `BackColor`)
  when they equal the form default. `_twips_opt` returns None for absent; rules
  use `_has_full_geom` and skip None. Treating absent as 0 caused false
  "zero-size" / bounds violations on inherited-default controls.
- **Opaque text control with no `BackColor` renders on white.** Access omits the
  default white BackColor — this is exactly how white-on-white slips through, so
  `_effective_background` defaults Label/TextBox/ComboBox/ListBox to white.
- **`ControlType =` is often absent** in modern exports; type comes from the
  `Begin <Type>` keyword. Rules key off `type_name`, not the int.
- **Attached labels are nested inside their control's block** → `_parse_controls`
  never enumerates them, so no overlap false positives there for free.
- **Access auto-grows form Width (and section Height) to fit controls**, so
  horizontal `out_of_bounds` rarely fires for forms (still useful for reports +
  negative coords). Not a bug — documented limitation.
- **`ConditionalFormat = Begin … End`** holds its own colours; `_extract_style`
  tracks block depth so only the control's own (depth-1) props are read.
- **System/theme colours** have the high bit `0x80000000` (e.g. `-2147483633`) —
  `_decode_bgr` flags them; contrast emits an `info` note instead of a number.
- **Conditional formatting** (`format_conditions`) overrides ForeColor/BackColor
  at runtime and is BINARY in the export — `_rule_contrast` skips those controls
  and notes them (can't verify the runtime colour statically).
- **Captions wrap.** Both Labels AND CommandButtons wrap their caption across
  lines; SaveAsText encodes the breaks as literal `\015\012`. `_caption_lines`
  splits them; truncation counts how many display lines the text needs
  (`ceil(line_width/avail)`) vs how many fit the height (`round(height/lineH)`,
  `lineH ≈ fontPt*20*1.2`). A 540-twip button shows 2 lines of 11pt — use
  `round`, not `floor`, or you under-count and false-flag.
- **Heuristic width is approximate.** Narrow UI fonts (Calibri/Tahoma) average
  ~0.46× the point size per glyph; the heuristic only flags a line as
  overflowing past **1.25×** the available width (absorbs metric error). WizHook
  uses 1.02×. Without these, bold header labels that fit get false-flagged.
- **Transparent buttons are a click layer, not an overlap.** A `Transparent=True`
  CommandButton stacked on a styled Label/Rectangle is the standard Access
  custom-button pattern (the label shows the colour, the invisible button takes
  the click) — `_rule_overlap` skips any pair where one side is a transparent
  button. (Classic command buttons ignore `BackColor` even with `UseTheme=No`
  on Win11/Office16, so this label+transparent-button trick is how you get
  coloured tiles.)
- **`sibling_inconsistency` clusters, not modes.** A form legitimately uses two
  sizes (tall main buttons + a row of short inline buttons). `_accepted_clusters`
  treats any value ≥2 controls share as a norm; only a lone outlier (and not a
  >2× different class like a memo box) is flagged. Needs ≥4 controls in the group.

### WizHook text measurement
`measure="auto"|"wizhook"|"heuristic"`. WizHook (`_measure_text_batch`) measures
exact rendered width in ONE COM round-trip via a temp std module +
`_invoke_app_run`. It REQUIRES a compiled VBA project (`Application.IsCompiled`);
during active development the ERP project is usually uncompiled, so it fails and
falls back to the conservative heuristic (a `note` is added when `measure` was
explicitly `wizhook`). The embedded lint always uses `heuristic` (fast, no Run
dependency). Default everywhere leans on the heuristic for reliability.

## Build-a-form-from-scratch recipes (v0.7.38)

### Add VBA to a form you just created with ac_create_form

Before v0.7.38, calling `ac_set_code(form, "Option Compare Database\n...")`
on a freshly-created form failed with `errors while importing` — `LoadFromText`
was always invoked and `restore_binary_sections` had nothing to restore from.
Now `ac_set_code` detects VBA-only input (`_looks_like_vba_only`: no
`Version =` / `Begin Form`, but Option/Sub/Function/etc.) and routes through
`_inject_vba_after_import` (Design view → `HasModule=True` → VBE write). No
`LoadFromText` round-trip, layout preserved.

```
ac_create_form(db, "frmFoo")
ac_create_control(db, "form", "frmFoo", "CommandButton",
                  {"left": 100, "top": 100, "width": 1500, "height": 400},
                  control_name="btCerrar")  # NEW: top-level control_name
ac_set_code(db, "form", "frmFoo",
            "Option Compare Database\nOption Explicit\n"
            "Private Sub btCerrar_Click()\n"
            "    DoCmd.Close acForm, Me.Name\n"
            "End Sub\n")  # routes via VBE, not LoadFromText
```

If you need to write a full form export (e.g. cloning the binary sections of
another form), include `Version =NN` / `Begin Form` and the original
`LoadFromText` path runs. The two paths are mutually exclusive — the
detection in `_looks_like_vba_only` is the discriminator.

### Drop a control inside a TabControl Page

`ac_create_control` accepts `parent` (or `Parent` — case-insensitive since
v0.7.38) as a special key that maps to the 4th positional arg of
`CreateControl(form, type, section, parent, column, l, t, w, h)`. Passing
`Parent` with capital P used to fall through to `setattr(ctrl, "Parent", ...)`
which Access rejects with `"Property 'CreateControl.Parent' can not be set"`
— misleading because Parent IS available, just not via setattr.

```
ac_create_control(db, "form", "frmFoo", "CommandButton",
                  {"Parent": "tabGestion",    # case-insensitive special key
                   "Left": 100, "Top": 100, "Width": 2000, "Height": 500,
                   "Caption": "Acción", "OnClick": "[Event Procedure]"},
                  control_name="btMiAccion")
```

If `Parent` doesn't refer to an existing TabControl Page (or other container
like OptionGroup), the control lands in Detail at the requested coordinates
and `Parent` is silently ignored by CreateControl — same behaviour as VBA.

### Read VBE from a brand-new form

Before v0.7.38: `ac_vbe_module_info(form, "frmFoo")` on a form just made by
`ac_create_form` raised `Subscript out of range`. The error message blamed
the Trust Center, but the actual cause was `HasModule=False` — VBComponents
had nothing to return because the code module had not been created yet.
Now `_force_vbe_init` activates `HasModule` when opening the form in Design
view during the retry, so this works out of the box.

If you want to *be explicit*, the original workaround is still valid:

```
ac_set_form_property(db, "form", "frmFoo", {"HasModule": True})
ac_vbe_module_info(db, "form", "frmFoo")  # then this works too
```

## VBE procedure editing (v0.7.42)

Field-report fixes for `vbe.py`. Three behaviours to keep in mind:

- **`ProcStartLine` owns the blank separator above a proc** (it equals the
  previous proc's `End` + 1, so it includes the blank/comment lines VBE attributes
  to the proc). `ac_vbe_replace_proc` therefore, *when replacing*, counts the run
  of leading whitespace-only lines (`lead`) and deletes/inserts at `start + lead`
  over `count - lead` — preserving the separator. A pure delete (`new_code==""`)
  still deletes the whole `[start, count]` range (separator included) so a deleted
  proc doesn't leave an orphan blank. Do NOT "simplify" this back to
  `DeleteLines(start, count)` for the replace path — that re-introduces the
  blank-eating bug Tom reported.
- **The Option-placement health check is comment-header-aware**, not
  line-number-thresholded. `_check_module_health` flags an `Option …` line only
  when real code (non-blank, non-comment `'`/`Rem`, non-`Option`) already appeared
  above it. A banner comment header of any length is fine. Do NOT restore the old
  `i >= 5` threshold — it false-positived on long headers (e.g. `_modTest`).
- **`new_lines` is an alias for `new_code` in `access_vbe_replace_lines`.** The
  dispatcher (`_new_lines_to_code`) joins a list with `\n` (so `""` entries are
  blank lines) and tolerates a JSON-encoded string from string-serialising
  clients. A single-mode replace that deletes lines but inserts nothing appends a
  note — the silent destructive-delete footgun (wrong arg name → empty `new_code`
  → pure delete) is now surfaced, not hidden.

`start_line` vs `body_line` (get_proc / module_info): `start_line` is the VBE proc
start (includes the blank/comment lines above); `body_line` is the
`Sub`/`Function`/`Property` declaration line. Use `start_line` for whole-proc ops,
`body_line` for body line-range edits.

## Common Gotchas

- VBE line numbers are **1-based**
- `ProcCountLines` can inflate the last proc's count past end of module — always clamp with `min(count, total - start + 1)`
- Access must be `Visible = True` for VBE COM access to work
- *"Trust access to the VBA project object model"* must be enabled in Access Trust Center

### CreateForm via COM shows "Save As" MsgBox
- **Do NOT** call `CreateForm()` directly followed by `_save_and_close()`.
- Use `access_create_form` tool: `CreateForm()` -> `DoCmd.Save(acForm, autoName)` -> `DoCmd.Close(acForm, autoName, acSaveNo)` -> `DoCmd.Rename(desired, acForm, autoName)`.
- Pass `record_source` to bind the form to a table/query and `default_view` (0=Single, 1=Continuous, 2=Datasheet, ...) to set the initial view — both are applied on the live `CreateForm()` object before `DoCmd.Save`. Without `record_source`, every bound `ControlSource` on the form will render as `#Name?`.
- Alternative: export an existing form with `ac_get_code`, modify the text, reimport with `ac_set_code`.

### AutoExec / startup forms block OpenCurrentDatabase
- `_switch()` holds Shift key during `OpenCurrentDatabase` (standard Access bypass). Auto-opened forms are closed as safety net.
- `AutomationSecurity = 3` is set as defence-in-depth but does NOT suppress AutoExec macro objects (tested).
- `_Session.reopen(path)` always applies SHIFT bypass.

### Linked tables and dbAttachSavePWD
- `dbAttachSavePWD` = **131072** (0x20000), NOT 65536.
- Setting `TableDef.Attributes` from Python COM before Append does not work reliably. Use `DoCmd.TransferDatabase(acLink, ..., StoreLogin:=True)` instead.

### ac_execute_sql / ac_execute_batch
- Both use try/except retry with `dbSeeChanges` for ODBC linked tables with IDENTITY columns.
- DELETE/DROP/TRUNCATE/ALTER require `confirm_destructive=true`.

### MCP schema type coercion
- Some MCP clients serialize ALL arguments as strings. `_fixup_schema()` widens schemas to accept both native types and strings. `_coerce_arguments()` converts back before dispatch.
- Do NOT change schemas back to strict `"type": "integer"` — clients can't be trusted.

### Jet SQL DDL Gotchas
- `YESNO` is not valid in DDL — use `BIT`, or better use `access_create_table`
- `DEFAULT` is not supported in `CREATE TABLE` — use `access_set_field_property` or `access_create_table`
- Multiple JOINs need nested parens: `FROM (A INNER JOIN B ON ...) INNER JOIN C ON ...`
- `AUTOINCREMENT` works as a type in DDL
- Use `SHORT` not `SMALLINT`, `LONG` not `INT`
- Prefer `access_create_table` over `CREATE TABLE` for full type + default + description support

### VBA Language Gotchas
- **`Private Type` without `End Type`**: All code after the block remains "inside" the type. If you get "Statement invalid inside Type block" on a correct-looking line, check for missing `End Type` above.
- **`SysCmd acSysCmdInitMeter`/`acSysCmdUpdateMeter`**: Cause intermittent "Illegal function call". Use `SysCmd acSysCmdSetStatus, "..."` instead.

### ActiveX controls
- Type 119 (`acCustomControl`): pass `class_name` with ProgID (e.g. `Shell.Explorer.2`).
- Type 128 (`acWebBrowser`): native control, no ActiveX needed.

## Critical DO NOTs

- **Do NOT remove the `DispatchEx` fallback** in `_Session._launch()`. `_launch()` tries `GetActiveObject("Access.Application")` first to attach to a user's running Access (avoids spawning a second process); on failure it falls back to `DispatchEx`, which is required after `/decompile` kills to bypass stale ROT entries. Do NOT swap `DispatchEx` for `Dispatch` in the fallback — `Dispatch` latches onto the stale ROT entry.
- **Do NOT call `cls._app.Quit()` unconditionally in `_decompile()` / `ac_decompile_compact()`**. Check `_Session._attached` first — when True we attached to the user's Access and must only `CloseCurrentDatabase()`, keeping the instance alive. Only when `_attached=False` (we spawned via `DispatchEx`) is `Quit(1)` safe. Same applies to the `atexit` handler `_Session.quit()`.
- **Do NOT use `EnsureDispatch`** — it changes binding for all 61 tools and adds `gen_py` cache dependency.
- **Do NOT run `OpenCurrentDatabase` in a separate thread** — COM STA objects can only be used from the thread that created them.
- **Do NOT call `CreateForm()` directly** — use `access_create_form` tool to avoid the "Save As" MsgBox.
- **Do NOT change schemas to strict `"type": "integer"`** — MCP clients can't be trusted to send correct types.
- **Do NOT auto-decompile on DB open** — only on first compile. Auto-decompile on open caused SHIFT key stuck issues and process accumulation on MCP reconnect.

## MCP SDK Patch (local to this machine)

The MCP Python SDK (`mcp/shared/session.py`) swallows all exceptions with a generic `-32602` error. A local patch at `c:\program files\python310\lib\site-packages\mcp\shared\session.py` adds full traceback to `ErrorData.message` and `ErrorData.data`. Re-apply after `pip install --upgrade mcp`.
