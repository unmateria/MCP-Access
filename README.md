# mcp-access

**Give any AI assistant full control over Microsoft Access databases.**

Create forms, write VBA, design tables, manage controls, run queries, build relationships, and edit every corner of an `.accdb` — all through natural language. 61 tools that turn Access into something you can *talk to*.

No Access expertise required. Just describe what you want.

```
"Create a form called Invoices with a ListBox, two date filters, and a search button"
"Add a VBA click handler that filters the recordsource by date range"
"Create a table called audit_log with timestamp, user, and action fields"
"List all controls inside the Payment tab and change the combo's row source"
```

The AI handles the COM automation, design view, VBA modules, binary sections, cache invalidation, and all the ugly parts. You get the result.

### What it can do

- **Forms & Reports** — create, export, import, screenshot, click, type. Full UI automation loop
- **VBA** — read, write, replace, compile, and *run* procedures. Line-level or full-proc editing
- **Controls** — create, delete, modify, list. Finds controls nested inside TabControl pages
- **Tables & SQL** — create via DAO, alter, query, batch execute. Linked ODBC tables supported
- **Relationships, indexes, references, queries, macros** — full CRUD
- **Maintenance** — compact & repair, decompile bloated databases, export structure docs

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
| `access_create_form` | Create a new form without triggering the "Save As" MsgBox that blocks COM. Optional `has_header` for header/footer section |

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
| `access_find_usages` | Search a name across VBA code, query SQL, and control properties (ControlSource, RecordSource, RowSource, SourceObject, DefaultValue, ValidationRule) in one call |

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
