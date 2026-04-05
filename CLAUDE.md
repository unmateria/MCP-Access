# CLAUDE.md — mcp-access MCP Server

## Overview

MCP server for reading and editing Microsoft Access databases (`.accdb`/`.mdb`) via COM automation (pywin32). Runs as stdio MCP server. Entry point: `access_mcp_server.py`. Implementation: `mcp_access/` package (~7000 lines across 20 modules).

## Architecture

- **Singleton COM session** (`_Session`): one `Access.Application` instance shared across all tool calls. Opening a different `.accdb` closes the previous one.
- **Dedicated COM thread** (`_com_executor`): All tool calls run in a single-threaded `ThreadPoolExecutor` with `CoInitialize()`. This keeps COM in one STA thread while the asyncio event loop stays free to read/write stdio. Without this, blocking COM calls would stall the event loop and cause the MCP SDK to produce `-32602 Invalid request parameters` errors from message corruption.
- **Caches**: `_vbe_code_cache` (VBE text), `_parsed_controls_cache` (control parsing), `_Session._cm_cache` (CodeModule COM objects). All invalidated on DB switch, object modification, and design operations.
- **Binary section handling**: `ac_get_code` strips PrtMip/PrtDevMode from form/report exports; `ac_set_code` restores them automatically before import.

## Tools (61 total)

| Category | Tools |
|----------|-------|
| **Database** | `access_create_database`, `access_close` |
| **Objects** | `access_list_objects`, `access_get_code`, `access_set_code`, `access_export_structure`, `access_delete_object`, `access_create_form` |
| **SQL/Tables** | `access_execute_sql`, `access_execute_batch`, `access_table_info`, `access_search_queries`, `access_create_table`, `access_alter_table` |
| **VBE line-level** | `access_vbe_get_lines`, `access_vbe_get_proc`, `access_vbe_module_info`, `access_vbe_replace_lines`, `access_vbe_find`, `access_vbe_search_all`, `access_vbe_replace_proc`, `access_vbe_patch_proc`, `access_vbe_append` |
| **Controls** | `access_list_controls`, `access_get_control`, `access_create_control`, `access_delete_control`, `access_set_control_props`, `access_set_multiple_controls` |
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
| **Cross-reference** | `access_find_usages` |
| **Knowledge base** | `access_tips` |

## Key Implementation Details

### Encoding in ac_set_code
- **Modules** (`.bas`): written as `cp1252` (ANSI) — Access expects no BOM for VBA modules
- **Forms, reports, queries, macros**: written as `utf-16` (UTF-16LE with BOM) — Access LoadFromText expects this

### Control parsing (_parse_controls)
The Access export format nests controls inside sections:
```
Begin Form
    Begin                    ← defaults block (NOT controls)
    End
    Begin Section            ← section (Detail, FormHeader, FormFooter)
        Begin                ← container
            Begin Label      ← REAL CONTROL
            End
            Begin Page       ← CONTAINER — children re-scanned
                Begin        ← anonymous wrapper
                    Begin ComboBox  ← child control (parent = Page)
                    End
                End
            End
        End
    End
    Begin ClassModule        ← VBA code
    End
End Form
```
The parser scans for `Begin <TypeName>` where TypeName matches known control types (`_CTRL_TYPE` dict) at any depth, ignoring the defaults block.

**Container types** (`_CONTAINER_TYPES = {"Page", "OptionGroup"}`): When the parser finds a container control, it records the container, then re-scans inside the block instead of skipping it. Child controls get a `"parent"` field with the container's name. A `container_stack` tracks nesting so deeply nested containers (Page inside OptionGroup, etc.) are handled correctly.

### VBE + Design view conflict
After design operations (`ac_set_control_props`, `ac_create_control`, `ac_delete_control`), the form may remain open in Design view. `ac_vbe_replace_proc` now:
1. Closes the form in Design view (DoCmd.Close with acSaveYes)
2. Invalidates `_cm_cache` for the object
3. Then accesses the VBE CodeModule

All design operations invalidate all three caches in their `finally` block.

### VBE robustness (v0.7.17)
Three layers of protection added to all VBE write operations:

**1. Whitespace-tolerant matching in `ac_vbe_patch_proc`:**
- When exact `str.replace` fails, `_ws_normalized_match()` strips leading whitespace from each line and does a sliding-window search. This catches patches where indentation differs (4 spaces vs 8, tabs vs spaces).
- If both exact and ws-normalized match fail, `_closest_match_context()` uses `difflib.SequenceMatcher` to find the most similar line and returns a contextual snippet (3 lines around the best candidate), making errors actionable instead of just "not found".
- Fallback matches are reported in the result as `ws_fallback_notes`.

**2. Structural health check (`_check_module_health`):**
Called after every write operation (`replace_lines` single+batch, `replace_proc`, `patch_proc`, `append`). Three checks:
- **Option placement**: Detects `Option Explicit`/`Option Compare` on lines > 5 (should always be at the top).
- **Duplicate labels**: Regex scan for `label:` patterns that appear more than once (common after copy-paste errors).
- **Line count sanity** (batch mode only): Compares expected total (`original - deleted + inserted`) with actual `cm.CountOfLines`.

Warnings are appended to the return string, never fail the operation.

**3. Option Explicit/Compare protection:**
- `_strip_option_lines()` removes `Option Explicit`/`Option Compare` from code being written to wrong positions.
- `ac_vbe_append`: Always strips Option lines (append goes to end of module — Option there is always wrong). Returns NOOP if code was only Option lines.
- `ac_vbe_replace_proc` / `ac_vbe_patch_proc`: Strips Option lines only when `start > 5` (proc is not at the top of the module).
- `_inject_vba_after_import` (code.py + controls.py): Auto-prepends `Option Compare Database` and `Option Explicit` if not present in the first 5 lines of injected VBA.

### ac_execute_sql safety
- SELECT results are limited by `limit` parameter (default 500, max 10000). If truncated, response includes `truncated: true`.
- DELETE/DROP/TRUNCATE/ALTER require `confirm_destructive=true` — without it the server returns an error.
- `_DESTRUCTIVE_PREFIXES` tuple defines the guarded prefixes.

### ac_execute_batch (batch SQL)
Executes multiple SQL statements in one call. Accepts `statements: [{sql, label?}, ...]`.
- `stop_on_error` (default true): stops at first error, returns partial results with `stopped_at` index.
- `confirm_destructive`: applies to entire batch — pre-scans all statements for destructive prefixes.
- SELECT statements return `{rows, count}` (limit 100 per SELECT). Others return `{affected_rows}`.
- Response: `{total, succeeded, failed, results: [{index, label?, status, ...}]}`.

### ac_get_form_property
Reads properties of a form/report (RecordSource, Caption, DefaultView, HasModule, etc.).
- If `property_names` is provided, reads only those. Otherwise reads all readable properties.
- Opens in Design view, reads, closes. Uses `_serialize_value` for COM value conversion.

### ac_set_multiple_controls
Modifies properties on multiple controls in a single design-view session.
- Opens form/report once, iterates controls, applies props, saves and closes once.
- Each control reports `{name, applied, errors?}`. Invalidates all 3 caches on completion.

### Search tools — regex and limits
- All search tools (`ac_vbe_find`, `ac_vbe_search_all`, `ac_search_queries`, `ac_find_usages`) support `use_regex=true` for regex patterns via `_text_matches()` helper.
- `ac_vbe_search_all` and `ac_search_queries` accept `max_results` (default 100). When exceeded, response includes `truncated: true`.
- `ac_find_usages` delegates to `ac_vbe_search_all` and `ac_search_queries` internally (DRY). Only control property scanning is inline.

### Compact & Repair (ac_compact_repair) / Decompile (ac_decompile_compact)
Closes the DB, compacts to temp file in same directory, does atomic swap (original→.bak, tmp→original), then reopens. Clears all 3 caches. Rollback on failure.

**Reopen with SHIFT**: Both `ac_compact_repair` and `ac_decompile_compact` reopen the database after compacting via `_Session.reopen()`, which forces `_switch()` (SHIFT held + auto-close forms). Previously they called `app.OpenCurrentDatabase()` directly, which could trigger AutoExec/startup forms/wizards and block COM indefinitely.

### Relationship attributes (_REL_ATTR)
`_REL_ATTR` maps DAO Relation.Attributes bitmask: 1=Unique, 2=DontEnforce, 256=UpdateCascade, 4096=DeleteCascade.

### VBA References (ac_manage_reference)
After add/remove, invalidates `_vbe_code_cache` and `_Session._cm_cache` since references affect VBA compilation. Guards against removing built-in references (VBA, Access, DAO).

### Query management (ac_manage_query)
CRUD for QueryDefs via DAO. `delete` requires `confirm=true`. `_QUERYDEF_TYPE` maps DAO QueryDef.Type to readable names (0=Select, 32=Delete, 48=Update, etc.).

### Auto-decompile on first open (v0.7.18)
`_Session._switch()` now runs `/decompile` + SHIFT automatically the first time each `.accdb` is opened in a session. This strips orphaned p-code so compile errors are real, not phantom. Tracked in `_decompiled_dbs` set (cleared on `_force_cleanup` / `quit`). Adds ~10s latency on first open only. If `MSACCESS.EXE` is not found, logs a warning and continues.

### Compile VBA (ac_compile_vba)
Uses `app.RunCommand(126)` (`acCmdCompileAndSaveAllModules`). Invalidates VBE caches after compilation. Optional `timeout` parameter — if compilation shows a MsgBox (error dialog), the watchdog dismisses it automatically (same pattern as `ac_run_vba`). After the error, `_get_vbe_error_location()` reads `VBE.ActiveCodePane.GetSelection()` to report the exact module, line number, and surrounding code where the error occurred.

**Reliable compilation (v0.7.13+)**: Multiple layers to avoid false "compiled" results:
1. VBE edits via COM don't always invalidate `IsCompiled`, so `RunCommand(126)` on an already-compiled project is a no-op. Fix: insert+remove a dummy comment in a standard module before compiling to force `IsCompiled=False`.
2. `RunCommand(126)` without the VBE window open silently skips form/report modules. Fix: open `VBE.MainWindow.Visible=True` before compiling, restore afterwards.
3. As a safety net, `_verify_module_structure()` scans ALL modules (standard + form/report) for executable code outside Sub/Function/Property/Type/Enum blocks. This catches the specific pattern of accidentally deleted Sub headers leaving orphan code that VBA absorbs into the previous procedure.

**Block mismatch detection (v0.7.18)**: After `RunCommand(126)`, checks `app.IsCompiled`. If False, `_find_block_mismatches()` parses ALL VBA modules for mismatched block structures: `If/End If`, `For/Next`, `Do/Loop`, `While/Wend`, `Select Case/End Select`, `With/End With`. Handles: single-line `If x Then action` (not pushed to stack), comments after `Then`, conditional compilation directives (`#If`/`#End If`), single-line `For Each...: Next` and `Do While...: Loop` with colons. Returns module name, line number, and error description for each mismatch.

### Output report (ac_output_report)
Uses `DoCmd.OutputTo(acOutputReport=3, ...)`. `_OUTPUT_FORMATS` maps format names to Access format strings. Auto-generates output_path if omitted.

### Transfer data (ac_transfer_data)
Consolidated import/export for Excel and CSV. Excel uses `DoCmd.TransferSpreadsheet` with `acSpreadsheetTypeExcel12Xml=10`. CSV uses `DoCmd.TransferText`.

### Field properties (ac_get_field_properties / ac_set_field_property)
Reads all `Field.Properties` (skips COM errors on unreadable ones). Set uses `_coerce_prop()` with fallback to `CreateProperty`.

### Startup options (ac_list_startup_options)
`_STARTUP_PROPS` lists 14 common startup properties. Reads each via DB Properties fallback to GetOption.

### DAO field types (ac_table_info)
`_DAO_FIELD_TYPE` maps DAO Type integers to readable names. AutoNumber is detected as Long (type 4) with `dbAutoIncrField` attribute (bit 16).

### Create form (ac_create_form)
Creates a new form without triggering the "Save As" MsgBox that blocks COM. Uses `CreateForm()` → `DoCmd.Save(acForm, autoName)` → `DoCmd.Close(acForm, autoName, acSaveNo)` → `DoCmd.Rename(desired, acForm, autoName)`. Optional `has_header=true` toggles header/footer section via `RunCommand(36)` before saving. Invalidates all 3 caches.

### Create database (ac_create_database)
Uses `app.NewCurrentDatabase()` then closes and reopens with `OpenCurrentDatabase()` to ensure `CurrentDb()` works reliably. Refuses to overwrite existing files. Bypasses `_Session._switch()` (which requires file to exist) and manages Access lifecycle directly.

### Create table via DAO (ac_create_table)
Creates tables using DAO `CreateTableDef` + `CreateField` instead of DDL SQL. Supports all field types via `_FIELD_TYPE_MAP`, defaults, descriptions, and primary keys in a single call. More robust than `CREATE TABLE` via `access_execute_sql` which has Jet SQL limitations (no DEFAULT, no YESNO type). Uses `_set_field_prop()` helper for post-creation property assignment.

### Alter table via DAO (ac_alter_table)
Modifies table structure: `add_field` (with type, size, default, description), `delete_field` (requires `confirm=true`), `rename_field`. Uses DAO `TableDef.CreateField/Fields.Delete/Fields.Name` directly.

### List objects with tables (ac_list_objects)
`access_list_objects` now supports `object_type="table"` via `app.CurrentData.AllTables`. System tables (`MSys*`) and temp tables (`~*`) are filtered out.

### Delete object (ac_delete_object)
Uses `DoCmd.DeleteObject(AC_TYPE[object_type], object_name)`. Requires `confirm=true` (destructive). Invalidates all 3 caches in `finally`.

### Run VBA (ac_run_vba)
Uses `Application.Run` via direct `InvokeTypes` call (bypasses pywin32's late-bound `__getattr__`). Max 30 arguments (Access limit). Result from Functions is captured; non-serializable COM types converted to `str`.

The helper `_invoke_app_run()` builds the full 31-param call with `pythoncom.Missing` for unused optional args, converted to `VT_ERROR/DISP_E_PARAMNOTFOUND` by `InvokeTypes`. This is necessary because Access COM rejects `Invoke()` calls missing the 30 optional params with `DISP_E_BADPARAMCOUNT`.

**Form module support** (`Forms.FormName.Method` syntax): When `procedure` starts with `Forms.`, uses direct COM `app.Forms(name).Method()` instead of `Application.Run`. The form must be open.

**Timeout parameter**: Optional `timeout` (seconds). If exceeded, `_dismiss_access_dialogs(hwnd)` finds Access modal dialogs (class `#32770`) via `win32gui.EnumWindows` and sends `WM_CLOSE` to dismiss them. The hwnd is captured on the main thread before starting the Timer (COM is apartment-threaded — accessing `app.hWndAccessApp` from the Timer thread fails silently). Without `timeout`, blocks indefinitely on MsgBox (backward compatible).

### Eval VBA (ac_eval_vba)
Uses `Application.Eval` via `InvokeTypes` (same pattern as `_invoke_app_run`). Evaluates a string expression in Access context. Can call form module functions (`Eval("Forms!frmX.MiFuncion()")`), read form properties, use domain functions (DLookup, DCount), and built-in VBA functions. Only Functions (not Subs). Form must be open.

**Auto-fallback for unsupported expressions**: `Application.Eval` cannot resolve class default instances (`VB_PredeclaredId = True`), variables, or other VBA project-level symbols. When Eval fails, `ac_eval_vba` automatically creates a temp standard module with a wrapper function (`_mcp_eval_wrapper`) that evaluates the expression in the full VBA project namespace, calls it via `Application.Run`, and cleans up the module in `finally`. If both Eval and fallback fail, the error message includes both errors and suggests using `access_run_vba` directly.

### Screenshot (ac_screenshot) — message pump + OpenForm timeout
`wait_ms` uses `pythoncom.PumpWaitingMessages()` loop (~60 Hz) instead of `time.sleep()`. This pumps Windows messages so `Form_Timer` events fire, ActiveX controls initialize, and WebBrowser navigates during the wait.

`open_timeout_sec` (default 30): before calling `DoCmd.OpenForm`, a daemon thread is started. If `OpenForm` does not return within the timeout, the thread sends `PostMessage(WM_KEYDOWN, VK_ESCAPE)` to the Access hwnd to cancel any pending `OpenRecordset` in the form's Load event, then `TimeoutError` is raised. Without this, a slow ODBC query in `Form_Load` can block `DoCmd.OpenForm` indefinitely (observed: 40+ minutes). The hwnd is captured **before** `OpenForm` blocks (COM is STA — the cancel thread cannot access `app.hWndAccessApp` directly).

### Delete relationship (ac_delete_relationship)
Uses `db.Relations.Delete(name)` via DAO.

### Find usages (ac_find_usages)
Cross-reference search in 3 locations: VBA code (all modules/forms/reports), SQL of all queries, and control properties (ControlSource, RecordSource, RowSource, SourceObject, DefaultValue, ValidationRule) via SaveAsText exports. `max_results` default 200.

## Adding a new tool

1. Write the implementation function (e.g. `ac_new_tool()`)
2. Add a `types.Tool(...)` entry to the `TOOLS` list
3. Add an `elif name == "access_new_tool":` branch in `call_tool()`
4. Update the tool count in this CLAUDE.md and README.md

## MCP SDK Patch: -32602 error detail (mcp 1.26.0)

The MCP Python SDK (`mcp/shared/session.py`, line ~380) catches **all** exceptions during request handling and returns a generic `-32602 Invalid request parameters` error with an empty `data` field. This makes debugging impossible — the actual exception (Pydantic validation, COM error, etc.) is swallowed.

**Patch applied** to `c:\program files\python310\lib\site-packages\mcp\shared\session.py`:
- `logging.warning` now includes full traceback
- `logging.debug` changed to `logging.warning` so the failing message is always visible
- `ErrorData.message` now includes the exception string (e.g. `"Invalid request parameters: 'NoneType' object..."`)
- `ErrorData.data` now includes the full traceback instead of empty string

This patch is local to this machine and will be lost on `pip install --upgrade mcp`. Re-apply if needed. The upstream issue is that the catch-all `except Exception` at line 380 swallows errors from `model_validate`, `_received_request`, and `_handle_incoming` indiscriminately.

## Common Gotchas

- VBE line numbers are **1-based**
- `ProcCountLines` can inflate the last proc's count past end of module — always clamp with `min(count, total - start + 1)`
- Access must be `Visible = True` for VBE COM access to work
- *"Trust access to the VBA project object model"* must be enabled in Access Trust Center

### CreateForm via COM shows "Save As" MsgBox
- `app.CreateForm()` opens a new form in Design view. `DoCmd.Close(acForm, name, acSaveYes)` triggers a "Save As" dialog that blocks the COM session.
- **Fix**: `access_create_form` tool uses the sequence: `CreateForm()` → `DoCmd.Save(acForm, autoName)` (saves with auto-name, no dialog) → `DoCmd.Close(acForm, autoName, acSaveNo)` (already saved) → `DoCmd.Rename(desired, acForm, autoName)`. No dialogs at any step.
- **Do NOT** call `CreateForm()` directly followed by `_save_and_close()` — always use `access_create_form` tool instead.
- Alternative: export an existing form with `ac_get_code`, modify the text, and reimport with `ac_set_code` (avoids CreateForm entirely).

### AutoExec / startup forms block OpenCurrentDatabase
- Databases with `AutoExec` macros or startup forms (especially modal `acDialog` forms like login/welcome screens) block the `OpenCurrentDatabase` COM call indefinitely. The call doesn't return until the user manually closes the form.
- Fix: `_switch()` holds the Shift key via `keybd_event(VK_SHIFT)` during `OpenCurrentDatabase`. This is the standard Access trick to bypass AutoExec and startup forms. After opening, any auto-opened forms are closed as a safety net.
- `_Session.reopen(path)` — convenience method that clears `_db_open` and calls `_switch()`, for use after `CloseCurrentDatabase` + `CompactRepair` sequences. All reopens in `maintenance.py` use this method to ensure SHIFT bypass is always applied.
- `AutomationSecurity = 3` (msoAutomationSecurityForceDisable) does NOT work — Access ignores it for database-level AutoExec macros.
- VK_ESCAPE to dismiss modal forms is unreliable (doesn't always reach the right window).

### Auto-decompile on first database open
- `_switch()` calls `_decompile(path)` if the DB has not been decompiled yet in this session (`_decompiled_dbs` set).
- `_decompile()` closes COM completely (`CloseCurrentDatabase` + `Quit`), spawns `MSACCESS.EXE /decompile` with SHIFT held, waits ~8s, kills the process, then re-launches COM via `_launch()`.
- Without this, stale p-code can cause phantom compile errors (e.g. "Block If without End If" that only surfaces after manual Debug > Compile in VBE) while `RunCommand(126)` silently reports "compiled".

### "You already have the database open" after MCP reconnect
- After `/mcp` reconnect, the MCP server process restarts (`_Session._app = None`) but Access.exe keeps running with the DB open. New `Dispatch("Access.Application")` connects to the existing instance, and `OpenCurrentDatabase` fails with "already have the database open".
- Fix: `_switch()` catches this specific error and syncs internal state (`_db_open = path`) without re-opening.

### dbAttachSavePWD and linked tables
- `dbAttachSavePWD` = **131072** (0x20000). NOT 65536 (that's `dbAttachExclusive`).
- Setting `TableDef.Attributes` from Python COM before Append **does not work reliably** (Type Mismatch errors). It works in native VBA but fails via pywin32.
- `ac_relink_table` uses `DoCmd.TransferDatabase(acLink, ..., StoreLogin:=True)` instead of DAO `CreateTableDef` + `Attributes` for reliable `dbAttachSavePWD` handling.
- `DoCmd.DeleteObject(acTable, name)` is used to remove the old link before recreating. This works from Python COM, unlike `db.TableDefs.Delete()` which can leave stale references.
- If `TransferDatabase` fails after deleting the old link, `ac_relink_table` attempts rollback by recreating the original link.

### ac_execute_sql / ac_execute_batch retry pattern
- Both use try/except retry with `dbSeeChanges` for ODBC linked tables with IDENTITY columns.
- If the first attempt fails and the retry also fails, the **original** error is raised (not the retry error).

### ac_set_code backup
- Forms, reports, **and modules** are backed up via `SaveAsText` before `LoadFromText`. If import fails, the backup is restored automatically.

### MCP schema type coercion (integer/boolean as strings)
- Some MCP clients serialize ALL tool arguments as strings. The MCP SDK validates against JSON Schema before `call_tool()` runs, so `"1"` fails for `{"type": "integer"}`.
- Fix: `_fixup_schema()` runs at module load and widens all schemas to accept `["integer", "string"]` and `["boolean", "string"]`. `_coerce_arguments()` in `call_tool()` converts string args to the expected type before dispatch.
- Do NOT change schemas back to strict `"type": "integer"` — clients can't be trusted to send correct types.

### Application.Run and late-bound COM (DISP_E_BADPARAMCOUNT)
- `Application.Run` has 31 params (1 required + 30 optional). pywin32's late-bound `Dispatch` uses `IDispatch.Invoke()` which only passes provided args — Access COM rejects this with `DISP_E_BADPARAMCOUNT` because the 30 optional params lack `VT_ERROR/DISP_E_PARAMNOTFOUND` markers.
- Fix: `_invoke_app_run()` calls `_oleobj_.InvokeTypes()` directly with full arg types + `pythoncom.Missing` padding. Same protocol as `EnsureDispatch`-generated wrappers, but without changing the binding model for all other tools.
- Do NOT switch `_Session._launch()` to `EnsureDispatch` — it would change binding for all 56 tools and add `gen_py` cache dependency.

### ac_run_vba and modal dialogs
- Without `timeout`: `Application.Run` blocks indefinitely if VBA shows `MsgBox`/`InputBox`.
- With `timeout`: `_dismiss_access_dialogs()` fires via `threading.Timer`, finds `#32770` dialogs owned by Access via `win32gui.EnumWindows`, sends `WM_CLOSE`. The blocked `InvokeTypes` then returns and the tool reports a timeout error.

### ac_create_control and ActiveX
- Type 119 (`acCustomControl`): pass `class_name` with the ProgID (e.g. `Shell.Explorer.2`) to initialize the OLE control via `ctrl.Class = class_name`.
- Type 128 (`acWebBrowser`): **native** WebBrowser control, no ActiveX/OLE needed.
- `_CTRL_TYPE` maps SaveAsText type numbers (for parsing). `_AC_CONTROL_TYPE_NAMES` adds AcControlType enum names (128=WebBrowser, 129=NavigationControl, etc.) for `CreateControl`.

### Jet SQL DDL Gotchas (access_execute_sql)
- `YESNO` is not valid in DDL — use `BIT` for Yes/No fields, or better yet use `access_create_table` which accepts `yesno`/`boolean`
- `DEFAULT` is not supported in `CREATE TABLE` Jet SQL — use `access_set_field_property` afterwards, or `access_create_table` which handles defaults automatically
- Multiple JOINs require nested parentheses: `FROM (A INNER JOIN B ON ...) INNER JOIN C ON ...`
- `AUTOINCREMENT` works as a type in DDL (no need for `IDENTITY` like SQL Server)
- Use `SHORT` instead of `SMALLINT`, `LONG` instead of `INT` in DDL
- Prefer `access_create_table` over `CREATE TABLE` via SQL for full type + default + description support in one call

### VBA Language Gotchas

- **`Private Type` without `End Type`**: All code after the block remains "inside" the type → error "Statement invalid inside Type block" on any subsequent `Declare`/`Function`/`Sub`. If the compiler gives this error on a line that looks correct, check that all `Private Type` blocks have their `End Type`.
- **`SysCmd acSysCmdInitMeter`/`acSysCmdUpdateMeter`**: Cause "Illegal function call" intermittently (especially with value=maxValue, or without calling `acSysCmdRemoveMeter` between sequences). Always use `SysCmd acSysCmdSetStatus, "..."` instead — never fails.
