# CLAUDE.md — mcp-access MCP Server

## Overview

MCP server for reading and editing Microsoft Access databases (`.accdb`/`.mdb`) via COM automation (pywin32). Runs as stdio MCP server. Entry point: `access_mcp_server.py`. Implementation: `mcp_access/` package (~7000 lines across 20 modules).

## Architecture

- **Singleton COM session** (`_Session`): one `Access.Application` instance shared across all tool calls. Opening a different `.accdb` closes the previous one.
- **Dedicated COM thread** (`_com_executor`): All tool calls run in a single-threaded `ThreadPoolExecutor` with `CoInitialize()`. This keeps COM in one STA thread while the asyncio event loop stays free to read/write stdio.
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

### Application.Run via InvokeTypes
`Application.Run` has 31 params (1 required + 30 optional). pywin32's late-bound `Dispatch` can't handle this. `_invoke_app_run()` calls `_oleobj_.InvokeTypes()` directly with `pythoncom.Missing` padding. Same approach for `Application.Eval` via `_invoke_app_eval()`.

## Adding a new tool

1. Write the implementation function (e.g. `ac_new_tool()`)
2. Add a `types.Tool(...)` entry to the `TOOLS` list
3. Add an `elif name == "access_new_tool":` branch in `call_tool()`
4. Update the tool count in this CLAUDE.md and README.md

## Common Gotchas

- VBE line numbers are **1-based**
- `ProcCountLines` can inflate the last proc's count past end of module — always clamp with `min(count, total - start + 1)`
- Access must be `Visible = True` for VBE COM access to work
- *"Trust access to the VBA project object model"* must be enabled in Access Trust Center

### CreateForm via COM shows "Save As" MsgBox
- **Do NOT** call `CreateForm()` directly followed by `_save_and_close()`.
- Use `access_create_form` tool: `CreateForm()` -> `DoCmd.Save(acForm, autoName)` -> `DoCmd.Close(acForm, autoName, acSaveNo)` -> `DoCmd.Rename(desired, acForm, autoName)`.
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
- **Do NOT use `EnsureDispatch`** — it changes binding for all 61 tools and adds `gen_py` cache dependency.
- **Do NOT run `OpenCurrentDatabase` in a separate thread** — COM STA objects can only be used from the thread that created them.
- **Do NOT call `CreateForm()` directly** — use `access_create_form` tool to avoid the "Save As" MsgBox.
- **Do NOT change schemas to strict `"type": "integer"`** — MCP clients can't be trusted to send correct types.
- **Do NOT auto-decompile on DB open** — only on first compile. Auto-decompile on open caused SHIFT key stuck issues and process accumulation on MCP reconnect.

## MCP SDK Patch (local to this machine)

The MCP Python SDK (`mcp/shared/session.py`) swallows all exceptions with a generic `-32602` error. A local patch at `c:\program files\python310\lib\site-packages\mcp\shared\session.py` adds full traceback to `ErrorData.message` and `ErrorData.data`. Re-apply after `pip install --upgrade mcp`.
