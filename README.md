# mcp-access

<!-- mcp-name: io.github.unmateria/msaccess-database -->

**Give any AI assistant full control over Microsoft Access databases.**

Create forms, write VBA, design tables, manage controls, run queries, build relationships, and edit every corner of an `.accdb` — all through natural language. 65 tools that turn Access into something you can *talk to*.

No Access expertise required. Just describe what you want.

```
"Create a form called Invoices with a ListBox, two date filters, and a search button"
"Add a VBA click handler that filters the recordsource by date range"
"Create a table called audit_log with timestamp, user, and action fields"
"List all controls inside the Payment tab and change the combo's row source"
```

The AI handles the COM automation, design view, VBA modules, binary sections, cache invalidation, and all the ugly parts. You get the result.

### What it can do

- **Forms & Reports** — create, clone, export, import, screenshot, click, type. Full UI automation loop
- **VBA** — read, write, replace, compile, and *run* procedures. Line-level or full-proc editing
- **Controls** — create, delete, modify, list, set tab order. Finds controls nested inside TabControl pages
- **Tables & SQL** — create via DAO, alter, query, batch execute, full-text search across every Text/Memo field. Linked ODBC tables supported
- **Relationships, indexes, references, queries, macros** — full CRUD. Clone any object (form / report / module / class / query / macro) preserving VBA and binary sections
- **Maintenance** — compact & repair, decompile bloated databases, export structure docs. Office install autodetected (no more hardcoded Office 16 paths)

Works with Claude Code, Cursor, Windsurf, Continue, or any MCP-compatible client.

---

## Requirements

- Windows (COM automation is Windows-only)
- Microsoft Access installed (any version that supports VBE, 2010+)
- Python 3.9+
- *"Trust access to the VBA project object model"* enabled in Access Trust Center

## Installation

```bash
pip install mcp pywin32
```

### Enable VBA object model access

`File → Options → Trust Center → Trust Center Settings → Macro Settings`
→ check **Trust access to the VBA project object model**

Or run the included PowerShell script:

```powershell
.\enable_vba_trust.ps1
```

## Register with Claude Code

**Global** (available in all projects):
```bash
claude mcp add access -- python C:\path\to\access_mcp_server.py
```

**Project-only** (creates `.mcp.json` in current directory):
```bash
claude mcp add --scope project access -- python C:\path\to\access_mcp_server.py
```

## Register with other MCP clients

Add to your MCP config file (`.mcp.json`, `mcp.json`, or client-specific settings):

```json
{
  "mcpServers": {
    "access": {
      "type": "stdio",
      "command": "python",
      "args": ["C:\\path\\to\\access_mcp_server.py"]
    }
  }
}
```

Compatible with any MCP-compliant client (Cursor, Windsurf, Continue, etc.).

## Tools (61)

### Database

| Tool | Description |
|------|-------------|
| `access_create_database` | Create a new empty `.accdb` database file |
| `access_close` | Close the COM session and release the `.accdb` file |

### Database objects

| Tool | Description |
|------|-------------|
| `access_list_objects` | List objects by type (`table`, `module`, `form`, `report`, `query`, `macro`, `all`). System tables filtered |
| `access_get_code` | Export an object's full definition as text |
| `access_set_code` | Import modified text back (creates or overwrites) |
| `access_export_structure` | Generate a Markdown index of all modules, forms, reports, queries |
| `access_delete_object` | Delete a module, form, report, query, or macro. Requires `confirm=true` |
| `access_create_form` | Create a new form without triggering the "Save As" MsgBox that blocks COM. Optional `has_header` for header/footer section, `record_source` (bind to table/query), `default_view` (0=Single, 1=Continuous, 2=Datasheet, ...) |

### SQL & tables

| Tool | Description |
|------|-------------|
| `access_execute_sql` | Run SQL via DAO — SELECT returns rows as JSON (`limit` default 500). DELETE/DROP/ALTER require `confirm_destructive=true` |
| `access_execute_batch` | Execute multiple SQL statements in one call. Supports mixed SELECT/INSERT/UPDATE/DELETE with per-statement results, `stop_on_error`, and `confirm_destructive` |
| `access_table_info` | Show table structure via DAO (fields, types, sizes, required, linked status) |
| `access_search_queries` | Search text in the SQL of ALL queries at once (find which queries reference a table, field, or keyword) |
| `access_create_table` | Create a table via DAO with full type, default, description and primary key support in one call. More robust than `CREATE TABLE` SQL |
| `access_alter_table` | Modify table structure via DAO: add field, delete field (requires `confirm=true`), rename field |

### VBE line-level editing

| Tool | Description |
|------|-------------|
| `access_vbe_get_lines` | Read a line range from a VBA module without exporting the whole file |
| `access_vbe_get_proc` | Get a procedure's code and position by name |
| `access_vbe_module_info` | List all procedures with their line numbers |
| `access_vbe_replace_lines` | Replace/insert/delete lines in a VBA module directly via VBE |
| `access_vbe_find` | Search text in ONE specific module. To search all modules at once, use `access_vbe_search_all` |
| `access_vbe_search_all` | Search text across ALL modules/forms/reports in the database at once |
| `access_vbe_replace_proc` | Replace a full procedure by name (auto-calculates line bounds). Strips misplaced `Option` lines, runs structural health check |
| `access_vbe_patch_proc` | Surgical find/replace within a procedure. Whitespace-tolerant fallback matching + contextual error messages when patches fail |
| `access_vbe_append` | Append code at the end of a module. Auto-strips `Option Explicit`/`Option Compare` to prevent misplacement |

### Form & report controls

| Tool | Description |
|------|-------------|
| `access_list_controls` | List all controls of a form/report with key properties. Controls inside Pages/OptionGroups include a `parent` field |
| `access_get_control` | Get the full definition block of a specific control (finds controls inside Pages/OptionGroups) |
| `access_create_control` | Create a new control via COM in design view. Supports `class_name` for ActiveX (type 119) ProgID initialization. Use type 128 (`acWebBrowser`) for native WebBrowser |
| `access_delete_control` | Delete a control via COM |
| `access_set_control_props` | Modify control properties via COM in design view |
| `access_set_multiple_controls` | Modify properties of multiple controls in a single design-view session |

### Text export/import

| Tool | Description |
|------|-------------|
| `access_export_text` | Export form/report/module as text via SaveAsText. Does NOT open Design view. UTF-16 LE output |
| `access_import_text` | Import form/report/module from text via LoadFromText. Replaces if exists. Auto-splits CodeBehindForm VBA |

### Database properties

| Tool | Description |
|------|-------------|
| `access_get_db_property` | Read a DB property (`CurrentDb.Properties`) or Access option (`GetOption`) |
| `access_set_db_property` | Set a DB property or Access option — creates the property if it doesn't exist |
| `access_get_form_property` | Read form or report properties (RecordSource, Caption, DefaultView, etc.). `object_type` required (`form` or `report`). Omit `property_names` for all |
| `access_set_form_property` | Set form/report properties (RecordSource, Caption, DefaultView, HasModule, etc.) via COM in Design view |

### Linked tables

| Tool | Description |
|------|-------------|
| `access_list_linked_tables` | List all linked tables with source table, connection string, ODBC flag |
| `access_relink_table` | Change connection string and refresh link — auto-saves credentials (`dbAttachSavePWD`) when UID/PWD detected. `relink_all=true` updates all tables with the same original connection |

### Relationships

| Tool | Description |
|------|-------------|
| `access_list_relationships` | List table relationships with field mappings and cascade flags |
| `access_create_relationship` | Create a relationship between two tables (supports cascade update/delete) |
| `access_delete_relationship` | Delete a relationship by name |

### VBA References

| Tool | Description |
|------|-------------|
| `access_list_references` | List VBA project references with GUID, path, broken/built-in status |
| `access_manage_reference` | Add (by GUID or file path) or remove a VBA reference — guards against removing built-in refs |

### Maintenance

| Tool | Description |
|------|-------------|
| `access_compact_repair` | Compact & repair the database — closes, compacts to temp, swaps atomically, reopens |
| `access_decompile_compact` | Remove orphaned VBA p-code via `/decompile`, recompile, then compact. Typical reduction: 60-70% on heavily-edited front-end databases. Use when a data-free `.accdb` exceeds 30-40 MB |

### Query management

| Tool | Description |
|------|-------------|
| `access_manage_query` | Create, modify, delete, rename, or read SQL of a QueryDef. Delete requires `confirm=true` |

### Indexes

| Tool | Description |
|------|-------------|
| `access_list_indexes` | List indexes of a table with fields, primary, unique, foreign flags |
| `access_manage_index` | Create or delete an index. Create requires fields list with optional sort order |

### VBA Compilation

| Tool | Description |
|------|-------------|
| `access_compile_vba` | Compile and save all VBA modules. Optional `timeout` to auto-dismiss error MsgBox |

### VBA & macro execution

| Tool | Description |
|------|-------------|
| `access_run_macro` | Execute an Access macro by name |
| `access_run_vba` | Execute a VBA Sub/Function. Standard modules via `Application.Run`, form modules via `Forms.FormName.Method` syntax (COM). Optional `timeout` auto-dismisses MsgBox/InputBox |
| `access_eval_vba` | Evaluate a VBA expression via `Application.Eval`. Domain functions, VBA built-ins, open form properties, standard module functions. Auto-fallback via temp module for class instances and other expressions Eval cannot resolve |

### Export

| Tool | Description |
|------|-------------|
| `access_output_report` | Export a report to PDF, XLSX, RTF, or TXT via `DoCmd.OutputTo` |

### Data transfer

| Tool | Description |
|------|-------------|
| `access_transfer_data` | Import/export data between Access and Excel (`.xlsx`) or CSV. Supports range (Excel) and spec_name (CSV) |

### Field properties

| Tool | Description |
|------|-------------|
| `access_get_field_properties` | Read all properties of a table field (DefaultValue, ValidationRule, Description, Format, etc.) |
| `access_set_field_property` | Set a field property — creates the property if it doesn't exist |

### Startup options

| Tool | Description |
|------|-------------|
| `access_list_startup_options` | List 14 common startup options (AppTitle, StartupForm, AllowBypassKey, etc.) with current values |

### Screenshot & UI automation

| Tool | Description |
|------|-------------|
| `access_screenshot` | Capture the Access window as PNG. Optionally opens a form/report first. Returns path, dimensions (original + image), and metadata. Configurable `max_width` (default 1920), `wait_ms` (pumps Windows messages — Timer events fire, ActiveX initializes), and `open_timeout_sec` (default 30 — sends ESC to cancel if `Form_Load` hangs on a slow query) |
| `access_ui_click` | Click at image coordinates on the Access window. Coordinates are relative to a previous screenshot (`image_width` required for scaling). Supports `left`, `double`, and `right` click |
| `access_ui_type` | Type text or send keyboard shortcuts. `text` for normal characters (WM_CHAR), `key` for special keys (enter, tab, escape, f1-f12, arrows, etc.), `modifiers` for combos (ctrl, shift, alt) |

### Cross-reference

| Tool | Description |
|------|-------------|
| `access_find_usages` | Search a name across VBA code, query SQL, and control properties (ControlSource, RecordSource, RowSource, SourceObject, DefaultValue, ValidationRule, LinkChildFields, LinkMasterFields) in one call |

### Knowledge base

| Tool | Description |
|------|-------------|
| `access_tips` | On-demand tips and gotchas. Topics: `eval`, `controls`, `gotchas`, `sql`, `vbe`, `compile`, `design`. Zero tokens until called |

## Typical workflows

### Targeted VBA editing (recommended)

```
1. access_list_objects      → find the module or form name
2. access_vbe_module_info   → get procedure list and line numbers
3. access_vbe_get_proc      → read the specific procedure
4. access_vbe_replace_lines → apply targeted line-level changes
5. access_close             → release the file when done
```

### Full object replacement (forms, reports, modules)

```
1. access_get_code   → export to text
2. (edit the text)
3. access_set_code   → reimport — binary sections are restored automatically
```

### Creating a new form

```
1. access_create_form(db, "myForm", has_header=true)  → creates empty form
2. access_create_control(db, "form", "myForm", "CommandButton", {Name: "btn1", ...})
3. access_vbe_append(db, "form", "myForm", code)  → add VBA event handlers
4. access_set_form_property(db, "form", "myForm", {HasModule: true, OnCurrent: "[Event Procedure]"})
```

### Screenshot & UI interaction

```
1. access_screenshot(db, "form", "myForm")  → capture form as PNG
2. (LLM reads the image and identifies UI elements)
3. access_ui_click(db, x=850, y=120, image_width=1920)  → click a button
4. access_ui_type(db, text="search term")  → type in a field
5. access_ui_type(db, key="enter")  → press Enter
6. access_screenshot(db)  → verify the result
```

## Notes

- Access runs visible (`Visible = True`) so VBE COM access works correctly.
- One Access instance is shared across all tool calls (singleton session). Opening a different `.accdb` closes the previous one.
- **COM thread isolation**: All COM calls run in a dedicated single-thread executor (`_com_executor`) with `CoInitialize()`. This keeps COM in one STA thread while the asyncio event loop stays free for stdio I/O, preventing `-32602` errors from message corruption.
- **Auto-reconnect**: if the COM session becomes stale (Access crashed, closed manually, or COM corruption), the server detects it via a health check and reconnects automatically on the next tool call.
- `access_get_code` strips binary sections (`PrtMip`, `PrtDevMode`, etc.) from form/report exports — `access_set_code` restores them automatically before importing.
- All VBE line numbers are 1-based.

## Known limitations

- **ActiveX controls** (type 119 = `acCustomControl`): `access_create_control` now accepts a `class_name` parameter with the ProgID (e.g. `Shell.Explorer.2`) to initialize the OLE control. For WebBrowser specifically, use type 128 (`acWebBrowser`) which creates a native control without OLE complexity. Setting `ctrl.Class` from COM may not work for all ActiveX controls — manual insertion from the ribbon remains the most reliable method.
- **`access_run_vba`**: Now supports form module procedures via `Forms.FormName.Method` syntax (direct COM access, form must be open). Also supports `timeout` parameter — if exceeded, auto-dismisses MsgBox/InputBox dialogs. For more flexible form interaction, use `access_eval_vba`.
- **Timer events** (`Form_Timer`): Now fire during `access_screenshot` when `wait_ms > 0` — the wait loop pumps Windows messages via `pythoncom.PumpWaitingMessages()`. Other tools still block the message pump.
- **`access_vbe_append`** previously HTML-encoded `&` as `&amp;` due to MCP transport escaping. Fixed in v0.7.3 with explicit `html.unescape()` decoding.

## Troubleshooting

### Intermittent `-32602 Invalid request parameters` errors

The MCP Python SDK (v1.26.0) has a catch-all `except Exception` in `mcp/shared/session.py` that swallows real errors and returns a generic `-32602` code with no detail. A local patch is applied to this machine that includes the actual exception and traceback in the error response. If you upgrade the `mcp` package, re-apply the patch — see `CLAUDE.md` for details.

## Changelog

### v0.7.36 — 2026-05-25

Five new capabilities — three new tools, one docs-only upgrade for macros, one transparent refactor for Office-version autodetect. **62 → 65 tools.** All changes additive; no existing schema or signature changed.

**Added**:
- **`access_search_data`** — search a text string across every Text/Memo field of every local table. Skips system tables (`MSys*`, `~*`) and linked tables (querying remote SQL servers per text column rarely ends well). Per-table and total caps; optional `tables` whitelist; `match_case`. Returns matches grouped by table with an `_excerpt` around each hit and `_matched_fields` per row.
- **`access_clone_object`** — duplicate a `form` / `report` / `module` / `class_module` / `query` / `macro` to a new name. Uses `SaveAsText` → `LoadFromText` with binary sections (PrtMip / PrtDevMode / NameMap / GUID) preserved by reading the raw export (bypasses `strip_binary_sections`). VBA code-behind rides along via the existing `ac_set_code` injection. Refuses to overwrite unless `overwrite=true`.
- **`access_manage_tab_order`** — `get` / `set` / `auto_renumber` the `TabIndex` of controls on a form/report. `set` does a two-phase reassignment (park at high indices, then assign 0..N-1) so it doesn't trip Access's per-section uniqueness constraint mid-loop. Skips controls that don't support TabIndex (Label, Line, Rectangle, Image, PageBreak, Page). Optional `section` filter.

**Changed**:
- **Macros docs**. Macros were already fully usable via `access_get_code` / `access_set_code` (UTF-16 encoding) plus `access_list_objects`, `access_run_macro`, `access_delete_object`. Tool descriptions now name macros explicitly and `access_tips('macros')` documents the read → edit → write workflow.
- **Office version autodetect**. The hardcoded `16.0` / `Office16` strings used by `_Session._suppress_recovery_dialog`, `_Session._decompile` and `ac_decompile_compact` now come from a one-shot registry probe (`_Session._detect_office_install`). Detection walks `Software\Microsoft\Office\<ver>\Access\InstallRoot\Path` under HKLM, HKLM\\WOW6432Node and HKCU (per-user C2R), picks the highest matching version with a working `MSACCESS.EXE`, falls back to `App Paths\MSACCESS.EXE\(Default)`, and finally to the previous hardcoded paths. Schema of `access_decompile_compact` is unchanged; behaviour on machines with a normal Office install is identical to v0.7.35.

### v0.7.35 — 2026-05-25

Preventive bug sweep across the codebase. No reported regressions; the issues below were caught by a full-package review while everything in production was working.

**Critical**:
- **`_Session.quit()` PID capture was a cross-thread COM call**: the new `taskkill` fallback added in the previous diff captured the Access PID by calling `app.hWndAccessApp()` inside `quit()`, but `quit()` runs from the `atexit` main thread while the COM proxy was created on the dedicated COM worker. Cross-thread STA calls return `RPC_E_WRONG_THREAD` silently, the result was `pid=None`, and the `taskkill` fallback was effectively a no-op exactly when it was needed (Access hung on Quit). PID is now captured at the end of `_launch()` — same thread that created the proxy — and stored as `cls._pid`.
- **`coerce_arguments` did not handle arrays/objects sent as JSON strings**: some MCP clients serialize every argument as a string, including arrays and objects. The schema fixup widened scalar types but not `array` / `object`, so batch tools (`execute_batch`, `create_table`, `manage_index`, `set_multiple_controls`, `vbe_patch_proc`, `vbe_replace_lines`, `find_definition`, `run_vba`, etc.) failed for those clients. Now both schemas and the coercer accept JSON-encoded strings for arrays and objects, plus `number` widening and additional truthy literals (`on`, `y`, `si`, `sí`).
- **`restore_binary_sections` injected binary blocks at the wrong `End` on forms with subforms**: the re-injector matched the first `End` after `Begin Form`, but a subform has its own nested `Begin Form ... End` — the injection landed inside the subform and corrupted the result on `ac_set_code`. Now tracks full block depth (same approach as the v0.7.34 control parser fix) and injects only at the outermost `End`.
- **`write_tmp` silently lost characters outside cp1252**: ANSI-encoded module writes used `errors="replace"`, which substituted `?` for any character the codepage couldn't represent (emoji, asian characters, `✓` in a comment). Switched to `errors="strict"` and the resulting `UnicodeEncodeError` now carries a snippet of the offending substring with an actionable hint.
- **`access_compile_vba` accepted `timeout` but never used it**: parameter was in the schema and signature, the body had a hardcoded `time.sleep(2)`. Now the parameter controls the watchdog grace window (default 2, clamped to 1-30).

**Medium**:
- VBE read tools (`get_lines`, `get_proc`, `module_info`, `find`, `search_all`) now close the form/report Design view before touching the CodeModule — same protection write tools already had. Skipping this could surface as `Catastrophic failure` (-2147418113) when the same object was open in design mode.
- `ac_vbe_replace_lines` no longer calls `cm.DeleteLines(line, 0)` (which raises in VBE) when count clamps to zero at end of module. Error message now lists separate upper bounds for replace/delete vs pure insert (was: misleading single range).
- `ac_vbe_patch_proc` normalizes `find_text` / `replace_text` line endings to CRLF before the exact-match check — callers commonly send LF and the exact match always fell through to the ws-normalized fallback. Also warns when `find_text` matches more than once (only the first occurrence is replaced via `.replace(..., 1)`).
- `_proc_kind` raises a descriptive error when a procedure name resolves to multiple kinds (a class with both `Property Get Foo` and `Property Let Foo` — normal VBA). Previously the first kind found was silently picked and write operations acted on it regardless of which the caller meant.
- `set_db_property` / `set_field_property` infer `dbDouble` (7), `dbDate` (8), `dbSingle` (6), `dbMemo` (12) when creating properties from `float` / `datetime` / `long-string` values. Previously these fell to `dbText` (10) and were stored as string.
- `_eval_via_temp_module` pre-binds `temp_name` so the cleanup `finally` can log a sensible name if `proj.VBComponents.Add(1)` succeeded but the subsequent `comp.Name` access failed (was: `UnboundLocalError` in the cleanup masking the real failure).
- Compile dialog watchdog captures up to 3 dialog screenshots/texts and picks the last one — the first dialog is often a benign "Save changes?" while the real compile error came last and used to be discarded.
- `ac_create_relationship` validates that the named tables exist and that each `local` / `foreign` field exists on its respective table before `Append`. Replaces a cryptic DAO error ("Item not found in this collection") with a message that names the missing item.
- `_check_module_health` / `ac_vbe_module_info` regexes recognize `Public Static Sub | Function | Property` (was: at most one modifier — VBA accepts both scope and `Static`).
- `decompile_compact` resets `_Session._pid` and `_attached` after killing the spawned Access, keeping the new `quit()` fallback consistent.

**Hardening**:
- `read_tmp` tries `utf-8` before `cp1252` — cp1252 is single-byte and almost never raises, so genuine UTF-8 files were being mis-decoded as mojibake.
- `_invoke_app_run` validates `len(args) <= 30` explicitly instead of producing a confusing `InvokeTypes` failure via a negative-multiplier padding (`[Missing] * -1` is just `[]`).
- `_split_code_behind` matches `CodeBehindForm` / `CodeBehindReport` only at start of a line, so a property value that happens to contain the literal can't trick the splitter.
- `SELECT … INTO` (make-table query) is flagged as destructive in `execute_sql` / `execute_batch`. Dead `_SQL_LINE_COMMENT` / `_SQL_BLOCK_COMMENT` regexes removed (the destructive guard already used `_sql_effective_prefix`).
- `ac_create_database` requires the path to end in `.accdb` or `.mdb`.
- `relink_table` UID/PWD detection uses a parameter-boundary regex (`(?:^|;)\s*(UID|PWD)\s*=`) instead of a substring check that would false-positive on values like `APP=My UID Manager`.
- Linked-table count query in `ac_table_info` escapes `]` in table names by doubling.
- `compact_repair` cleans up the orphaned `_compact_tmp.accdb` on the rolled-back swap path.
- Safe-args logging in `server.py` / `dispatcher.py` guards against non-string `code` to avoid `TypeError` inside the error handler.
- `tools.py` module docstring updated 58 → 62 tools.

### v0.7.34 — 2026-05-05

**`access_list_controls` silently lost controls inside `Page` / `OptionGroup`**: when any earlier control in the same Page had a multi-line property block (`GUID = Begin … End`, `NameMap = Begin … End`, `ConditionalFormat = Begin … End`, etc.) the depth counter inside the control's body matched plain `Begin <Type>` but not `Property = Begin`. The property's closing `End` was decremented without ever being incremented — the enclosing control closed prematurely, and every control that came after it inside the Page was never enumerated. The form-level loop in `_parse_controls` already handled this; the per-control loop now mirrors it (`r"^\w+\s*=\s*Begin\s*$"`).

Visible symptom: `access_list_controls` reported a TabControl Page as a 15-line empty stub even though the Page actually contained dozens of controls. Fixed in `mcp_access/controls.py:_parse_controls`. Tool count unchanged (62).

### v0.7.33 — 2026-05-04

**`access_create_form` silently dropped `record_source` and `default_view`**:

- **The bug**: callers passing `record_source="myTable"` or `default_view=2` got back a form that was bound to nothing and rendered in Single view. The arguments were not in the input schema, not in the dispatcher, and not in `ac_create_form`'s signature, so the MCP transport accepted them and discarded them without warning. Symptom downstream: every bound TextBox on the form (or on a continuous subform built on top of it) showed `#Name?`, because the form had no recordset to resolve `ControlSource` against.
- **The fix**: `ac_create_form(db_path, form_name, has_header=False, record_source=None, default_view=None)` now applies `RecordSource` and `DefaultView` on the live `CreateForm()` object before `DoCmd.Save`. Schema in `tools.py` and the `dispatcher.py` branch updated to forward the new arguments. Both fields are optional and back-compat (existing `access_create_form` calls without them keep working unchanged). Result dict now echoes back the applied values when set.
- **Why it matters for subforms**: the typical use is `access_create_form("subf_lines", record_source="my_table", default_view=1)` so the new form is immediately a continuous subform bound to the table. Without this fix, you had to follow up with `access_set_form_property` to bind it manually — and the missing `RecordSource` was easy to miss until the parent form rendered and showed `#Name?` everywhere.

Tool count unchanged (62).

### v0.7.31 — 2026-05-03

**Fix `_vbe_code_cache` returning stale text after external edits** — thanks to [@TvanStiphout-Home](https://github.com/TvanStiphout-Home) for reporting in [issue #26](https://github.com/unmateria/MCP-Access/issues/26):

- **The bug**: `access_vbe_get_proc` (and other VBE read tools) could return a cached snapshot of a procedure that no longer matched what was in the VBE module. The cache was only invalidated when the MCP itself wrote to the module — manual edits in the VBE (including Ctrl+Z), add-ins, or any change made outside the MCP left the cache stale. Worse, the WRITE tools (`access_vbe_replace_proc` / `patch_proc` / `replace_lines` / `append`) also read through the same cache before writing, so they could overwrite the wrong baseline and corrupt code.
- **The fix**: `_cm_all_code()` no longer caches; it reads directly from COM via `cm.Lines(1, total)` on every call. The `_vbe_code_cache` dictionary and all its `.pop` / `.clear` invalidation sites have been removed (`core.py`, `vbe.py`, `code.py`, `controls.py`, `compile.py`, `maintenance.py`, `relations.py`, plus stale imports in `database.py` and `helpers.py`). The `_Session._cm_cache` (CodeModule COM proxies) is kept — proxies are live, not snapshots, so they don't suffer the same problem and they save 2 COM calls per VBE tool. Tom's case (Claude wrote a buggy version with `replace_proc`, Tom reverted manually in VBE, the next `get_proc` still served the buggy cached version) now reads fresh from COM and matches the real module state.

**`access_relink_table` no longer hangs the COM session on a bad ODBC connect string**:

- **The bug**: passing an unreachable SQL Server (wrong host, firewalled named-instance, or simply not running) caused the ODBC driver to open a modal *"SQL Server Login"* dialog inside Access. There is no human at the keyboard in an MCP session, so the dialog was never dismissed and the entire COM session hung indefinitely.
- **The fix**: `_ensure_login_timeout()` injects `LoginTimeout=8` into the connect string when missing, so DAO bails out instead of opening a dialog. Before touching DAO at all, `_odbc_preflight()` test-opens the connect string via ADODB; if that fails, `ac_relink_table` raises a clean `RuntimeError` with the underlying ODBC error. `_detect_named_instance()` adds a hint for the common case of a `SERVER=host\instance` named instance whose UDP 1434 (SQL Browser) is firewalled — the suggested fix is to switch to explicit TCP `SERVER=host,port`.

Tool count unchanged (62).

### v0.7.30 — 2026-05-03

**Bug fix** — thanks to [@CaptainStormfield](https://github.com/CaptainStormfield) ([#27](https://github.com/unmateria/MCP-Access/pull/27)):

- **`inserted` line count was off by one when `new_code` ended with `\r\n`**: `len(decoded.splitlines())` treats a trailing newline as a *terminator*, so it does not count the extra blank line that VBE's `InsertLines` does add when the input ends in `\r\n`. In `access_vbe_replace_lines` batch mode this triggered spurious `WARNING: Expected N lines after edit, but module has N+k` health-check warnings — one count off per operation whose `new_code` had a trailing newline (intentional blank-line separators between procedures). Fix: replace the `splitlines()` estimate with an exact measurement — read `cm.CountOfLines` after `InsertLines` and subtract the pre-insert total. Same off-by-N pattern was also present in `ac_vbe_replace_proc` and `ac_vbe_append`; both corrected the same way for consistency.
- **`.gitignore`**: also added `.mcp.json` and `.claude/` so local Claude Code / MCP config doesn't leak into commits.

Tool count unchanged (62).

### v0.7.29 — 2026-04-24

**Audit pass** — bugs and UX rough edges found during a full-package review:

- **`access_vbe_patch_proc` could silently drop patches on form/report code-behind**: the function invalidated the VBE text cache but never called `DoCmd.Save` after writing, unlike `access_vbe_replace_proc` / `access_vbe_replace_lines` / `access_vbe_append`. On forms/reports this meant the object's dirty flag was never raised, so the edits could be discarded on close. Fix: added the same `app.DoCmd.Save(obj_type_code, object_name)` block that the other three writers use.
- **`access_delete_relationship` now requires `confirm=true`**: deleting a relationship is irreversible and was the only destructive operation in the package without a confirmation guard. Aligned with `access_execute_sql`'s `confirm_destructive` pattern.
- **Destructive-SQL detection ignored leading comments**: `-- note\nDELETE FROM t` and `/* prefix */ DROP TABLE t` passed through the `startswith("DELETE"|"DROP"|...)` check. New `_sql_effective_prefix()` walks past leading `--` line comments and `/* ... */` block comments before prefix-matching, in both `ac_execute_sql` and `ac_execute_batch`.
- **`access_output_report` now refuses to clobber existing files unless `overwrite=true`**: Access `DoCmd.OutputTo` silently overwrites, which can lose an earlier export if the same `output_path` is reused.
- **`access_ui_type` now uses `WM_UNICHAR` for non-ASCII characters**: accented Spanish (á, é, ñ), Cyrillic, CJK, and other code points >127 were being sent via `WM_CHAR` with `ord(ch)`, which routes through the window's ANSI code page and produces mojibake on English locales. Characters ≤127 still use `WM_CHAR` so keyboard shortcuts work the same.
- **`access_vbe_find` with `proc_name=""` no longer errors**: some MCP clients send `""` instead of omitting the argument when a procedure-scope filter is not wanted. Now treated as "search the whole module".
- **Hardened `GetActiveObject` attach**: a round-trip property read now validates that the returned COM reference points at a live Access process before `_attached=True` is set. Zombie marshalled references from a dying process now fall through to `DispatchEx` cleanly.
- **Deduplicated `_split_code_behind`**: was byte-identical in `code.py` and `controls.py`, now lives in `helpers.split_code_behind` with backwards-compatible re-exports at both old names.
- **Minor**: `"Error en %s"` log message in dispatcher translated to English; removed an unused `DB_SEE_CHANGES` import in `database.py`.

Tool count unchanged (62). Documented false positives from the review (type-name case mismatch in `ac_table_info`, silent-typo in `ac_set_control_props`) were verified against the code and turned out to be non-issues.

### v0.7.28 — 2026-04-24

**Polish of `access_find_definition`** (v0.7.27 follow-up):

- **Const values were contaminated by trailing comments**: `Public Const MAX_ROWS = 100  ' default page size` returned `value = "100  ' default page size"` instead of `"100"`. The naive `split("'", 1)[0]` stripper also broke on string literals with apostrophes like `Public Const Name = "O'Brien"`, truncating the value at the inner `'`. Fix: new `_strip_trailing_vba_comment()` helper that respects `"..."` string literals (same state machine as `_split_top_level_commas`), applied at all value-extraction sites.
- **`Default` property modifier not detected**: `Public Default Property Get Item(...)` — valid VBA on class modules with a default member — was invisible to the tool. `_FD_PROC_RE` only allowed `Static`. Fix: widened to `(?:(?:Static|Default)\s+)?`.
- **Line continuations not joined**: `Public Const FOO As Long = _\n    &H1000` returned `value = "_"` (only the first physical line was parsed). `Public Declare PtrSafe Function SendMessage _\n    Lib "user32" …` had the same problem. Fix: new `_join_continuations()` helper folds trailing-`_` continuations into single logical statements before matching. The reported `line` still points at the first physical line of the statement.
- **`scan_types` parameter**: optionally restrict the scan to any subset of `["module", "form", "report"]`. Passing `scan_types=["module"]` skips form/report code-behind and is **≈7× faster** on cold scans, because form/report scanning triggers a Design-view open/close round-trip per object when the VBComponent cache is empty. Useful when you know the target is a public declaration in a standard module — as Tom pointed out in his original request: *"All public enums are defined in modGlobal"*.
- **`first_only` parameter**: stop at the first match. Good for unique names in big databases.
- **`subkind` cleanup**: `subkind` is now only emitted when it carries information beyond `kind` — i.e. on `property` (Get/Let/Set) and `declare` (Sub vs Function). For `sub` and `function` it was redundant with `kind` and has been removed.

### v0.7.27 — 2026-04-24

**New tool** — thanks to [@TvanStiphout-Home](https://github.com/TvanStiphout-Home):

- **`access_find_definition`** — "Go To Definition" for VBA symbols, the mirror of `access_find_usages`. Scans every standard module, form code-behind and report code-behind for DECLARATIONS of a given name and returns where each one lives (object, line, full declaration, scope, and for constants the literal value). Detects `Const` (including multi-const lines like `Const A = 1, B = 2`), `Enum` + enum members, `Type` + type fields, `Sub`, `Function`, `Property Get/Let/Set`, `Declare` (Win32 API), and module-level `Dim/Public/Private/Global` variables. Locals inside a `Sub`/`Function`/`Property` are correctly skipped — they are not definitions in the "go to" sense. Case-insensitive by default (VBA is), with an optional `kinds` whitelist to narrow results (e.g. `["const","enum_member"]` when resolving a symbol used as a numeric literal like `dbAccess`). Previously the agent had to spawn throw-away VBA helpers with `MsgBox` just to discover the value of a named constant — this tool makes that unnecessary.

### v0.7.26 — 2026-04-23

**Bug fix** — thanks to [@TvanStiphout-Home](https://github.com/TvanStiphout-Home) ([#25](https://github.com/unmateria/MCP-Access/issues/25)):

- **`access_compile_vba` spawned a second Access instance when user had Access open**: Follow-up to v0.7.25. The `GetActiveObject` attach fix applied to `_Session._launch()`, but the auto-decompile path in `_Session._decompile()` (triggered on first compile per session) and `ac_decompile_compact()` still unconditionally called `cls._app.Quit(1)` and re-launched — killing the user's attached Access and spawning a fresh one via `DispatchEx`. Fix: new `_Session._attached` flag tracks whether we attached (`GetActiveObject`) or launched (`DispatchEx`). When attached, `/decompile` now releases the file lock via `CloseCurrentDatabase()` only, never calls `Quit()`, and reuses the same COM reference after the subprocess finishes. `atexit` handler also skips `Quit()` when attached so MCP shutdown doesn't kill the user's Access. Defence-in-depth PID-diff cleanup (`_list_msaccess_pids()` before vs. after the subprocess) kills any forked `msaccess.exe` children that escape `taskkill /T`. Refuses to decompile a path when the user has a **different** DB open, instead of silently closing their unsaved work.

### v0.7.25 — 2026-04-14

**Bug fix** — thanks to [@TvanStiphout-Home](https://github.com/TvanStiphout-Home) ([#24](https://github.com/unmateria/MCP-Access/issues/24)):

- **`_Session._launch()` always spawned a second Access process**: When the user already had Access open (common during interactive debugging), `win32com.client.DispatchEx("Access.Application")` unconditionally created a new COM instance instead of attaching to the running one. Fix: `_launch()` now tries `win32com.client.GetActiveObject("Access.Application")` first, syncs `cls._db_open` from the attached instance's `CurrentDb().Name` so `connect()` can skip an unnecessary `_switch()` when the target DB is already open, and falls back to `DispatchEx` only when no running instance exists. The `DispatchEx` fallback is still required after `/decompile` kills (stale ROT entries) — `GetActiveObject` will fail cleanly in that path since the process was taskkill'd.

### v0.7.24 — 2026-04-13

**Enhancement** — thanks to [@AccessWizard](https://github.com/AccessWizard) ([#23](https://github.com/unmateria/MCP-Access/pull/23)):

- **`AutomationSecurity = 3` defence-in-depth in `_switch()`**: Sets `msoAutomationSecurityForceDisable` before `OpenCurrentDatabase()` and restores to `msoAutomationSecurityLow` (1) in the `finally` block. Does not replace the Shift key bypass (Access ignores `AutomationSecurity` for AutoExec macro objects), but provides an extra safety layer for VBA auto-run code in edge cases where the Shift key doesn't register (remote desktop sessions, key events eaten by other processes).

**Bug fixes**:

- **`_exec_single_replace` counted inserted lines before HTML unescape**: `inserted = len(new_code.splitlines())` used the pre-unescape string; now uses `len(decoded.splitlines())` so the count matches what VBE actually received.
- **`ProcCountLines` could crash after patch**: `ac_vbe_patch_proc` called `cm.ProcCountLines()` without protection after patching — if the patch corrupted the proc structure, this threw an unhandled COM error. Now wrapped in try/except.
- **Registry key leak in `_suppress_recovery_dialog`**: If the second `SetValueEx` call failed, `CloseKey` was never reached. Now uses try/finally to guarantee key closure.
- **Atomic swap in `ac_compact_repair` could lose DB on double failure**: If both the swap rename and the rollback rename failed, the error message didn't indicate where the files ended up. Now reports the exact paths of both the `.bak` original and the compacted temp file.
- **`_eval_via_temp_module` warning could crash**: The error handler accessed `comp.Name` on a potentially dead COM object. Now uses the pre-captured `temp_name` variable.
- **`_inject_vba_after_import` Option check too narrow**: Only searched the first 5 lines for existing `Option Compare`/`Option Explicit`, so code with these directives at line 6+ got duplicates. Now searches the entire code.
- **Module encoding hardcoded to cp1252**: `ac_set_code` now uses `locale.getpreferredencoding()` (the system ANSI codepage) instead of hardcoded `cp1252`, so non-Western Windows systems (Greek, Cyrillic, etc.) work correctly.

### v0.7.23 — 2026-04-11

**Bug fixes** — thanks to [@CaptainStormfield](https://github.com/CaptainStormfield):

- **Property Let/Set procedures invisible to all VBE tools**: `_proc_kind()` only tried `kind=0` (`vbext_pk_Proc`) and `kind=3` (`vbext_pk_Get`), completely missing `kind=1` (`vbext_pk_Let`) and `kind=2` (`vbext_pk_Set`). Any Let-only or Set-only property (e.g. `Property Let ItemPrefix`) would fail with "Sub or Function not defined". Fix: new `_ALL_PROC_KINDS = (0, 1, 2, 3)` tuple — `_proc_kind()`, `_proc_of_line()`, and all callers now iterate all four VBE proc kinds. The old constant `_VBEXT_PK_PROPERTY = 3` was misleadingly named (3 is specifically `vbext_pk_Get`, not a generic "property" kind) and has been replaced with explicit `_VBEXT_PK_LET = 1`, `_VBEXT_PK_SET = 2`, `_VBEXT_PK_GET = 3`.
- **`access_vbe_module_info` silently dropped Property Let/Set entries**: The `seen` set deduplicated by procedure name alone, so when `Property Get Foo` was encountered first, `Property Let Foo` with the same name was skipped entirely. Fix: deduplicate by `(name.lower(), keyword.lower())` so Get, Let, and Set variants of the same property are listed as separate entries. Each entry now includes a `"keyword"` field (e.g. `"Property Get"`, `"Property Let"`). A `_KEYWORD_TO_KIND` mapping lets `module_info` pass the correct VBE kind directly to `_proc_bounds()` instead of relying on the blind iteration in `_proc_kind()`.
- **Fallback for VBE kind-specific lookup failures**: When VBE's `ProcStartLine` fails for a specific kind (Access quirk with certain Let-only or Set-only properties), the old fallback emitted an entry with no `body_line` or `count`. Fix: scans forward in the source text from the declaration line to the matching `End Property`/`End Sub`/`End Function` keyword to derive an accurate count.
- **Zombie COM object after `ac_decompile_compact`**: After `taskkill /F /T` kills the `/decompile` subprocess, Access doesn't run cleanup code and can leave a stale entry in the Windows Running Object Table (ROT). The subsequent `Dispatch("Access.Application")` in `_Session._launch()` latched onto this dead ROT entry, yielding a zombie COM object that passed the `_app.Visible` health check but failed on any database operation. Fix: replaced `win32com.client.Dispatch` with `win32com.client.DispatchEx` — always creates a fresh instance, bypassing the ROT entirely. Added a 1-second sleep after `taskkill` in both `_Session._decompile()` and `ac_decompile_compact()` as belt-and-suspenders to allow Windows time to evict the dead entry.

### v0.7.22 — 2026-04-08

**Bug fixes** — thanks to [@CaptainStormfield](https://github.com/CaptainStormfield) and [@unmateria](https://github.com/unmateria) (wizard-during-compact report), and [@TvanStiphout-Home](https://github.com/TvanStiphout-Home) (class module request):

- **`access_decompile_compact` silently launched Report Wizard on every call** (root cause of the "wizard hang" report): `maintenance.py` had `app2.RunCommand(137)  # acCmdCompileAllModules = 137` — but **137 is `acCmdNewObjectReport`**, not `acCmdCompileAllModules` (the correct value is 125 for compile-only or 126 for compile-and-save). Every `ac_decompile_compact` invocation was silently opening the Report Wizard and blocking the COM thread indefinitely until a human clicked Cancel. The "intermittent" symptoms in the original report were actually 100% reproducible — the wizard was always there, just sometimes hidden behind other windows. Fix: changed `RunCommand(137)` → `RunCommand(126)` (`acCmdCompileAndSaveAllModules`).
- **`access_compact_repair` / `access_decompile_compact` hang on unexpected dialogs** (defence-in-depth for any future wizard, ODBC credential prompts, recovery dialogs): Neither `CompactRepair` nor the `/decompile` subprocess had dialog protection — only `OpenCurrentDatabase` did. Fix: new `_call_with_dialog_watchdog(app, label, callable_fn)` generic helper wraps any blocking COM call with a polling daemon thread that dismisses any Access-owned dialog every 0.5s via `_dismiss_access_dialogs`. `_compact_with_watchdog` is a thin wrapper around it. The `RunCommand(126)` call in `ac_decompile_compact` is also wrapped in this helper. `_Session._decompile()` and `ac_decompile_compact` replace their fixed `time.sleep(3) + sleep(5)` sequence with a polling loop that calls `_dismiss_dialogs_by_pid(proc.pid)` on the standalone MSACCESS subprocess.
- **OpenCurrentDatabase watchdog no longer sends `VK_RETURN`** (v0.7.19 regression): The old one-shot watchdog sent Enter (`VK_RETURN`) to dismiss blocking dialogs — **dangerous on wizards**, as Enter clicks "Next >" and advances the wizard, creating stray `Report1`/`Form1` objects. Rewritten as a polling loop that delegates to `_dismiss_access_dialogs` with the new Cancel-first button priority.
- **`_try_click_button()` button priority fix**: Previously used a `set` of target button labels with undefined iteration order, so it could click any of End/OK/Cancel depending on `set` hash ordering. Now uses an explicit priority tuple `("cancel", "cancelar", "end", "finalizar", "ok", "aceptar")` — Cancel first so wizards cancel cleanly, End second to preserve existing `ac_run_vba` behaviour on VBA runtime-error dialogs (which have no Cancel button).
- **Wizard title detection**: `_dismiss_dialogs_by_pid` now matches windows where `class == '#32770'` **OR** the title contains `"wizard"` / `"asistente"` (case-insensitive). Catches non-standard wizard windows that don't use the `#32770` class.

**New feature**:

- **`access_set_code(object_type="class_module", ...)`**: Creates a VBA class module by injecting the four `Attribute VB_*` lines at the top of the text. Previously, `object_type="module"` always created a standard module. Tested on production (Access 2016): `Application.LoadFromText(acModule=5)` distinguishes class from standard modules by the presence of `Attribute VB_GlobalNameSpace`, `Attribute VB_Creatable`, `Attribute VB_PredeclaredId`, `Attribute VB_Exposed` at the top of the file — **NOT** by a `VERSION 1.0 CLASS` header (that header is for `VBComponent.Export`/`Import`, a different mechanism; passing it to `LoadFromText` makes Access interpret the header lines as literal VBA code and creates a corrupt Type=1 standard module). New `_ensure_class_module_header(code, name)` strips any BOM, strips any legacy `VERSION 1.0 CLASS` / `BEGIN` / `END` / `Attribute VB_Name` block the user may have pasted from a `VBComponent.Export` file, detects existing `Attribute VB_GlobalNameSpace` (round-trip safe — feeding `access_get_code` output back does not duplicate), and prepends the 4 attribute lines with CRLF endings. `class_module` re-uses `acModule=5` under the hood — no changes needed in `access_get_code` or `access_delete_object`. Verified on production DB round-trip: create → read → re-import → overwrite → delete, all with `VBComponent.Type == 2`.

### v0.7.21 — 2026-04-06

**Bug fixes** — thanks to [@CaptainStormfield](https://github.com/CaptainStormfield) ([PR #17](https://github.com/unmateria/MCP-Access/pull/17)):
- **Property Get/Let/Set procedures invisible to VBE tools**: All VBE call sites (`get_proc`, `module_info`, `replace_proc`, `patch_proc`, `find`) hardcoded `kind=0` (`vbext_pk_Proc`). Property procedures require `kind=3` (`vbext_pk_Property`). New helpers `_proc_kind()`, `_proc_bounds()`, `_proc_of_line()` try kind=0 first and fall back to kind=3
- **Wrong VBProject after decompile+compact**: `app.VBE.VBProjects(1)` could return `acwzmain` (wizard library) instead of the user's database project. New `_get_vb_project(app)` enumerates all VBProjects and matches by `.FileName` against the open database path
- **"Subscript out of range" after decompile**: `VBComponents(name)` could fail even though the component exists. `_get_code_module()` now retries once after forcing VBE initialisation (opens form/report briefly in Design view, or toggles `VBE.MainWindow.Visible` for modules)

### v0.7.20 — 2026-04-05

**Bug fixes** — thanks to [@CaptainStormfield](https://github.com/CaptainStormfield) ([PR #11](https://github.com/unmateria/MCP-Access/pull/11)):
- **`access_get_form_property` crash on binary GUID properties**: `_serialize_value` handled `bytes` but not `memoryview`, causing JSON serialization crash. Fix: extended type check to `(bytes, memoryview)`
- **`access_find_usages` missed subform link references**: `LinkChildFields` and `LinkMasterFields` were not in `CONTROL_SEARCH_PROPS`, so stale subform link references after table/field renames went undetected
- **`access_list_controls` / `access_get_control` missed Subform controls**: Access exports `Begin Subform` (lowercase f) but the constant was `"SubForm"` (capital F), causing case-sensitive matching to skip subforms entirely. Also fixed wrong control number in tips (was 114, correct is 112)
- **Spurious duplicate-label warnings in VBE health check**: The check scanned modules flat, flagging `ErrHandler:` as duplicate when it appeared in separate procedures. VBA labels are procedure-scoped — fix tracks Sub/Function/Property boundaries and only flags duplicates within the same procedure
- **`access_vbe_append` / `access_vbe_replace_lines` "Catastrophic failure" after design operations** ([#12](https://github.com/unmateria/MCP-Access/issues/12) — thanks [@CaptainStormfield](https://github.com/CaptainStormfield)): These VBE tools did not close the form/report in Design view before accessing the VBE CodeModule, causing `com_error(-2147418113)`. Fix: all VBE write functions now close the object and invalidate `_cm_cache` before accessing VBE, matching the pattern already used by `access_vbe_replace_proc` and `access_vbe_patch_proc`

### v0.7.19 — 2026-04-05

**OpenCurrentDatabase watchdog — auto-dismiss blocking dialogs:**
- `OpenCurrentDatabase` now runs in a background thread with a 10-second watchdog. If the open blocks (recovery dialog, save changes prompt, or any unexpected modal dialog), the watchdog:
  1. Captures a screenshot of the Access window and saves to `%TEMP%\access_blocked_*.png`
  2. Detects the blocking dialog via `GetForegroundWindow`
  3. Sends Enter (`VK_RETURN`) via `PostMessageW` to dismiss it (accepts default button)
  4. Logs the screenshot path for debugging
- **Recovery dialog suppression**: Before every `OpenCurrentDatabase`, writes `DisableAllCallersWarning=1` and `DoNotShowUI=1` to `HKCU\Software\Microsoft\Office\16.0\Access\Resiliency` registry key. This prevents the "last time you opened this file it caused a serious error" dialog that appears after a crash or `Stop-Process`

### v0.7.18 — 2026-04-05

**`access_compile_vba` — complete rewrite for reliable error detection:**
- **VBE CommandBars compile**: Now compiles via `VBE.CommandBars("Menu Bar").Controls("Debug").Execute()` instead of `RunCommand(126)`. This is equivalent to clicking Debug > Compile in VBE and reliably compiles ALL modules including form/report (RunCommand silently skipped them)
- **Error dialog text capture**: Watchdog reads the Microsoft error dialog text via Win32 `GetWindowText` before dismissing it. Returns the exact error message (e.g. "Compile error: Block If without End If") + module name + line number via `VBE.ActiveCodePane.GetSelection()`
- **Watchdog always active**: Polls every 0.5s regardless of timeout parameter, preventing hangs from unexpected dialogs
- **Block mismatch parser (fallback)**: When `IsCompiled=False` but no dialog was caught, `_find_block_mismatches()` statically analyzes VBA code for unclosed blocks. Handles single-line If, comments after Then, `#If`/`#End If`, and colon-separated statements
- **Lint removed from compile**: `_lint_form_modules()` no longer runs during compilation — it opened every form in Design view causing dozens of "Save changes?" dialogs and broken reference errors

**Auto-decompile moved to compile only:**
- `/decompile` + SHIFT now runs automatically on first compile per session (not on every DB open). Previous approach caused SHIFT key stuck issues and MSACCESS.EXE process accumulation on MCP reconnect

### v0.7.17 — 2026-04-01

**VBE robustness — 3 layers of protection for VBA editing:**

- **Whitespace-tolerant matching in `access_vbe_patch_proc`**: When exact match fails, a whitespace-normalized fallback strips leading indentation and retries. If both fail, `difflib.SequenceMatcher` finds the closest line and returns contextual error with 3 surrounding lines — no more opaque "not found" errors
- **Structural health check on all write operations**: After every VBE write (`replace_lines`, `replace_proc`, `patch_proc`, `append`), checks for: (1) `Option Explicit`/`Option Compare` misplaced below line 5, (2) duplicate labels, (3) unexpected line count changes (batch mode). Warnings in return string, never fails the operation
- **Option Explicit/Compare protection**: `access_vbe_append` auto-strips `Option` lines (they'd end up at the bottom of the module). `replace_proc` and `patch_proc` strip them when replacing non-top procedures. `access_set_code` and `access_import_text` auto-prepend `Option Compare Database` + `Option Explicit` when missing from injected VBA

### v0.7.16 — 2026-03-30

**Bug fix:**
- **`access_find_usages` missed SubForm SourceObject**: Control property search didn't include `SourceObject`, so SubForm/SubReport references to other forms were invisible. Deleting a form that was a SubForm's SourceObject would break the parent form silently. Fix: added `SourceObject` to `CONTROL_SEARCH_PROPS`

**Improvements:**
- **`access_get_code` description** now recommends `access_vbe_get_proc` for reading specific VBA procedures (faster, smaller output — avoids 90KB+ exports for large forms)
- **`access_tips("vbe")`** expanded with guidance: use `access_vbe_module_info` → `access_vbe_get_proc` flow instead of `access_get_code` for VBA investigation
- **`access_tips("gotchas")`** documents that SHIFT bypass on database open is automatic (no manual intervention needed)

### v0.7.15 — 2026-03-30

**Usability improvements (reduce LLM hallucination of parameter names):**
- **`access_vbe_append`**: renamed `new_code` parameter to `code` for consistency with `access_set_code` and other tools
- **`access_vbe_find`**: description now clarifies it searches ONE module and suggests `access_vbe_search_all` for cross-module search
- **`access_get_form_property`**: description now explicitly states `object_type` is required (`form` or `report`)

### v0.7.14 — 2026-03-29

**Improvement:**
- **`access_eval_vba` auto-fallback**: `Application.Eval` cannot resolve class default instances (`VB_PredeclaredId`), variables, or project-level symbols. Now when Eval fails, the tool automatically creates a temp standard module with a wrapper function, calls it via `Application.Run`, and cleans up. If both fail, the error includes both messages and suggests `access_run_vba`. Tool description updated to clarify supported/unsupported patterns

### v0.7.13 — 2026-03-29

**Bug fix:**
- **`access_compile_vba` false positive — reported "compiled" on broken code**: Two root causes: (1) VBE edits via COM don't always invalidate `IsCompiled`, so `RunCommand(126)` on an already-compiled project was a no-op. (2) Without the VBE window visible, `RunCommand(126)` silently skips form/report module compilation. Fix: dirty trick (insert+remove dummy comment) forces `IsCompiled=False`, then VBE MainWindow is opened before compiling so `RunCommand` behaves like clicking Debug > Compile. Additionally, new `_verify_module_structure()` scans all modules (standard + form/report) for executable code outside Sub/Function blocks as a safety net — catches the specific pattern of accidentally deleted Sub headers leaving orphan code

### v0.7.12 — 2026-03-27

**Bug fix:**
- **Compact/decompile reopen could trigger AutoExec/startup forms/wizards**: `ac_compact_repair` and `ac_decompile_compact` reopened the database after compacting via direct `app.OpenCurrentDatabase()`, bypassing the SHIFT key handling in `_switch()`. If the database had AutoExec macros or startup forms (e.g. Report Wizard), the reopen would block COM indefinitely. Fix: new `_Session.reopen()` method forces `_switch()` (SHIFT held + auto-close forms) for all reopens in maintenance operations

### v0.7.11 — 2026-03-26

**Bug fixes:**
- **AutoExec / startup forms block `OpenCurrentDatabase` indefinitely**: Databases with `AutoExec` macros that open modal forms (e.g. Northwind Developer Edition's welcome/login dialog via `acDialog`) block the COM call until the user manually closes the form. Fix: `_switch()` now holds the Shift key via `win32api.keybd_event(VK_SHIFT)` during `OpenCurrentDatabase` — the standard Access bypass for AutoExec and startup forms. After opening, any auto-opened forms are closed as a safety net. `AutomationSecurity = 3` does NOT work (Access ignores it for database-level AutoExec macros). VK_ESCAPE is unreliable (doesn't reach modal forms)
- **MCP clients sending integer/boolean arguments as strings**: Some MCP clients (e.g. Claude Desktop) serialize ALL tool arguments as strings. The MCP SDK validates against the JSON Schema before `call_tool()` runs, so `start_line: "1"` fails with `'1' is not of type 'integer'`. Fix: `_fixup_schema()` runs at module load and widens all 58 tool schemas to accept `["integer", "string"]` and `["boolean", "string"]`. `_coerce_arguments()` in `call_tool()` converts string args to the expected type before dispatch

### v0.7.10 — 2026-03-25

**Bug fix:**
- **`access_list_controls` / `access_get_control` didn't find controls inside Pages or OptionGroups**: The text parser (`_parse_controls`) consumed the entire `Begin Page ... End` block as one control and skipped all children inside it. Fix: new `_CONTAINER_TYPES` set (`{"Page", "OptionGroup"}`) and a `container_stack` mechanism — when the parser finds a container type, it records the container and re-scans inside the block instead of jumping past it. Child controls now include a `"parent"` field with the container's name. `delete_control` and `set_control_props` were unaffected (they use COM directly)

### v0.7.9 — 2026-03-22

**Bug fix:**
- **Intermittent `-32602 Invalid request parameters` errors**: Root cause — the `call_tool` handler was `async` but ran blocking COM calls synchronously, stalling the asyncio event loop. When the event loop couldn't read stdin fast enough, the MCP SDK's message parser received truncated or merged JSON-RPC frames, causing Pydantic `model_validate` to fail with a catch-all `-32602` error. Fix: all COM work now runs in a dedicated single-thread `ThreadPoolExecutor` (`_com_executor`) with `CoInitialize()`, keeping the event loop free for stdio I/O. The `call_tool` async wrapper uses `loop.run_in_executor()` to delegate to the COM thread

### v0.7.8 — 2026-03-20

**Bug fix:**
- **`access_screenshot` OpenForm hang**: `DoCmd.OpenForm` could block indefinitely when a form's `Form_Load` event ran a slow or blocked `OpenRecordset` (ODBC query to SQL Server). New `open_timeout_sec` parameter (default 30 s) starts a daemon thread before opening the form. If `OpenForm` does not return within the timeout, the thread sends `PostMessage(WM_KEYDOWN/WM_KEYUP, VK_ESCAPE)` to the Access window to cancel the pending DAO operation, then raises `TimeoutError` with a descriptive message. The hwnd is captured before `OpenForm` blocks so the cancel thread does not need COM access (STA restriction). `open_timeout_sec` can be increased for forms that are legitimately slow to load

### v0.7.7 — 2026-03-17

**New tool:**
- `access_decompile_compact` — removes orphaned VBA p-code by launching `MSACCESS.EXE /decompile`, recompiling all modules, then running Compact & Repair. Compact alone cannot reclaim p-code space; this combination achieves 60-70% size reduction on heavily-edited front-end databases (tested: 69 MB → 26 MB). Launches Access as a subprocess, waits 8 s for decompile to complete, kills the process, reconnects via COM for recompile, then compact

### v0.7.6 — 2026-03-17

**New tool:**
- `access_create_form` — create a new form safely via COM without the "Save As" MsgBox that blocks the session. Uses `CreateForm()` → `DoCmd.Save(acForm, autoName)` → `DoCmd.Close(acForm, autoName, acSaveNo)` → `DoCmd.Rename(desired, acForm, autoName)`. Optional `has_header=true` to create with header/footer section via `RunCommand(36)`

### v0.7.5 — 2026-03-17

**Known limitations reduced:**
- **Timer events fixed**: `access_screenshot` now uses `pythoncom.PumpWaitingMessages()` loop during `wait_ms` instead of `time.sleep()`. `Form_Timer` events fire, ActiveX controls initialize, WebBrowser navigates
- **MsgBox/InputBox timeout**: `access_run_vba` and `access_compile_vba` now accept optional `timeout` (seconds). If exceeded, `_dismiss_access_dialogs()` finds Access modal dialogs (class `#32770`) via `win32gui.EnumWindows` and sends `WM_CLOSE` to dismiss them
- **Form module support**: New `access_eval_vba` tool — evaluates expressions via `Application.Eval` (form properties, form module functions, domain functions, VBA built-ins). `access_run_vba` now supports `Forms.FormName.Method` syntax for direct COM access to open forms
- **ActiveX `class_name`**: `access_create_control` now accepts `class_name` parameter for ActiveX (type 119) — sets `ctrl.Class` with ProgID to initialize OLE. Type 128 (`acWebBrowser`) documented as native WebBrowser alternative. New AcControlType constants added (128-134)

**Improvements:**
- **`access_compile_vba` timeout + error diagnostics**: Optional `timeout` to auto-dismiss MsgBox on compilation error. Reports exact module, line, code context via `VBE.ActiveCodePane`, and captures dialog screenshot
- **`access_tips`** (new tool): On-demand knowledge base with tips and gotchas (eval, controls, sql, vbe, compile, design, gotchas). Zero tokens until called
- **`access_list_controls` / `access_get_control`**: Now detect conditional formatting — show `format_conditions` count when a control has `ConditionalFormat` entries

**Bug fix:**
- **"You already have the database open"** after MCP reconnect: `_switch()` now catches this error and syncs internal state instead of failing. Happens when the MCP server restarts but Access.exe keeps running with the DB open from the previous session

### v0.7.4 — 2026-03-16

**Bug fix:**
- **`access_run_vba` was completely broken** — every call failed with `DISP_E_BADPARAMCOUNT` (-2147352562). Root cause: pywin32's late-bound `Dispatch` uses `IDispatch.Invoke()` which only passes provided arguments. Access's `Application.Run` has 31 parameters (1 required + 30 optional) and its COM server rejects calls missing `VT_ERROR/DISP_E_PARAMNOTFOUND` markers for the 30 optional params. Fix: new `_invoke_app_run()` helper calls `_oleobj_.InvokeTypes()` directly with full argument types and `pythoncom.Missing` padding — the same COM protocol that early-bound (`EnsureDispatch`) wrappers generate, but without changing the binding model for the other 53 tools

### v0.7.3 — 2026-03-14

**Reliability improvements:**
- **Auto-reconnect COM**: `_Session.connect()` now performs a health check (`app.Visible`) before every tool call. If the COM session is stale (Access crashed, closed manually, or corrupted), it automatically reconnects instead of failing with cryptic COM errors
- **`access_vbe_append` / `access_vbe_replace_lines`**: fixed HTML entity encoding bug where `&` was silently converted to `&amp;` by MCP transport. Now applies `html.unescape()` to decode entities before inserting code
- **VBE cache invalidation**: `_get_code_module` now evicts stale cache entries on failure, preventing cascading "Subscript out of range" errors after `access_set_code` or COM reconnection
- **Tool descriptions updated** with known limitations:
  - `access_run_vba`: documents that only standard module procedures work (not form/report modules) and that MsgBox/InputBox blocks indefinitely
  - `access_create_control`: documents that ActiveX (type 126) creates empty containers without OLE initialization
  - `access_screenshot`: documents that Timer events do not fire during capture (no message pump)

### v0.7.2 — 2026-03-13

**Robustness improvements:**
- `access_relink_table`: added rollback — if `TransferDatabase` fails after deleting the old link, the original link is restored automatically. Previously the table would be left deleted with no replacement
- `access_execute_sql` / `access_execute_batch`: fixed silent retry swallowing errors. The `dbSeeChanges` retry pattern now preserves the original error message when both attempts fail, instead of showing only the retry error
- `access_set_code`: backup before import now includes modules (previously only forms/reports). If a module import fails, the original is restored via `LoadFromText`
- `access_run_vba`: tool description now warns that `MsgBox`/`InputBox` in VBA will block the call indefinitely. Recommends using `access_ui_click`/`access_ui_type` for UI interaction

### v0.7.1 — 2026-03-13

**Bug fix:**
- Fixed `access_relink_table` not persisting ODBC credentials: `_DB_ATTACH_SAVE_PWD` constant was **65536** (wrong — that's `dbAttachExclusive`) instead of **131072** (`dbAttachSavePWD`). Tables relinked with UID/PWD would lose credentials on next database open, causing login prompts
- Replaced DAO `CreateTableDef` + `Attributes` approach with `DoCmd.TransferDatabase(acLink, ..., StoreLogin=True)` which works reliably from Python COM (setting `Attributes` before `Append` works in native VBA but fails via pywin32 with Type Mismatch)

### v0.7.0 — 2026-03-12

**New tools (3):**
- `access_screenshot` — capture the Access window as PNG using `PrintWindow` API with DPI awareness. Optionally opens a form/report, captures, then closes it. Resizes to configurable `max_width` for token efficiency
- `access_ui_click` — click at image coordinates on the Access window. Scales from screenshot space to screen space automatically. Supports left, double, and right click
- `access_ui_type` — type text via `WM_CHAR` or send keyboard shortcuts via `keybd_event`. Supports special keys (enter, tab, escape, F1-F12, arrows) and modifier combos (ctrl, shift, alt)

**Infrastructure:**
- DPI awareness (`SetProcessDpiAwareness(2)`) set at module load for accurate window dimensions
- COM `hWndAccessApp` handled for both property and method variants

### v0.6.0 — 2026-03-10

**New tools (3):**
- `access_execute_batch` — execute multiple SQL statements in a single call with per-statement results, `stop_on_error` flag, and batch destructive guard
- `access_get_form_property` — read form/report properties (RecordSource, Caption, DefaultView, HasModule, etc.) via COM in design view
- `access_set_multiple_controls` — modify properties on multiple controls in a single design-view open/close cycle

### v0.5.0 — 2026-03-07

**New tools (5):**
- `access_create_database` — create a new empty `.accdb` database via `NewCurrentDatabase`
- `access_delete_object` — delete modules, forms, reports, queries, or macros via `DoCmd.DeleteObject` (requires `confirm=true`)
- `access_run_vba` — execute VBA Sub/Function via `Application.Run` with optional arguments and return value capture
- `access_delete_relationship` — delete a table relationship by name via DAO
- `access_find_usages` — cross-reference search across VBA code, query SQL, and control properties in a single call

**Enhancements:**
- `access_list_objects` now supports `object_type="table"` via `AllTables` (system/temp tables filtered)

### v0.4.0 — 2026-03-07

**New tools (10):**
- `access_manage_query` — create, modify, delete, rename, or read SQL of QueryDefs via DAO
- `access_list_indexes` / `access_manage_index` — list table indexes; create or delete indexes with field order and primary/unique flags
- `access_compile_vba` — compile and save all VBA modules (acCmdCompileAndSaveAllModules)
- `access_run_macro` — execute an Access macro by name
- `access_output_report` — export reports to PDF, XLSX, RTF, or TXT via DoCmd.OutputTo
- `access_transfer_data` — import/export data between Access and Excel (.xlsx) or CSV via DoCmd.TransferSpreadsheet/TransferText
- `access_get_field_properties` / `access_set_field_property` — read all field properties; set or create field-level properties (DefaultValue, ValidationRule, Description, etc.)
- `access_list_startup_options` — list 14 common startup options with current values

### v0.3.0 — 2026-03-07

**New tools (9):**
- `access_get_db_property` / `access_set_db_property` — read/write database properties (AppTitle, StartupForm, etc.) and Access application options
- `access_list_linked_tables` / `access_relink_table` — list linked tables with connection info; change connection strings with bulk relink support
- `access_list_relationships` / `access_create_relationship` — list and create table relationships with cascade flags
- `access_list_references` / `access_manage_reference` — list VBA references (with broken/built-in detection); add by GUID or path, remove by name
- `access_compact_repair` — compact & repair with atomic file swap and automatic reopen

### v0.2.1 — 2026-03-07

**New tools:**
- `access_search_queries` — search text in the SQL of all queries at once (equivalent to iterating `QueryDefs` with `InStr`)

**Improvements:**
- `access_execute_sql`: added `limit` parameter (default 500, max 10000) to cap SELECT results and prevent token explosions
- `access_execute_sql`: added `confirm_destructive` flag — DELETE/DROP/TRUNCATE/ALTER now require explicit confirmation
- `access_vbe_search_all` and `access_search_queries`: added `max_results` parameter (default 100) with `truncated` indicator
- `access_export_structure`: now returns the Markdown content directly (no extra Read needed)
- All tool descriptions compacted ~60% to reduce token overhead per MCP session

### v0.2.0 — 2026-03-05

**New tools:**
- `access_vbe_search_all` — search text across all modules, forms, and reports in a single call
- `access_table_info` — inspect table structure via DAO (field names, types, sizes, required flags, record count, linked status)
- `access_vbe_replace_proc` — replace or delete a full procedure by name without manual line arithmetic
- `access_vbe_append` — append code to the end of a module safely

**Bug fixes:**
- Fixed `access_set_code` corrupting VBA modules by writing UTF-16 BOM; modules now use `cp1252` (ANSI) encoding as Access expects
- Fixed `access_list_controls` returning empty results; control parser rewritten to correctly find `Begin <TypeName>` blocks at any nesting depth
- Fixed `access_vbe_replace_proc` failing with catastrophic COM error after design-view operations; now closes the form in Design view and invalidates cache before accessing VBE
- Fixed `access_vbe_module_info` reporting inconsistent `start_line`/`count` values; now uses COM `ProcStartLine` consistently and clamps count to module bounds
- Added boundary validation to `access_vbe_replace_lines` — checks `start_line` range and clamps `count` to prevent overflows

**Improvements:**
- All design-view operations (`access_create_control`, `access_delete_control`, `access_set_control_props`) now invalidate all internal caches in their `finally` block
