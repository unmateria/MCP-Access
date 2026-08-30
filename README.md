# mcp-access

<!-- mcp-name: io.github.unmateria/msaccess-database -->

**Give any AI assistant full control over Microsoft Access databases.**

Create forms, write VBA, design tables, manage controls, run queries, build relationships, and edit every corner of an `.accdb` — all through natural language. 68 tools that turn Access into something you can *talk to*.

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
- **UI lint** — `access_lint_form` flags *objectively* broken layouts (white-on-white text, overlaps, truncation, off-canvas controls). A checker, **not** a designer — see the note below

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

## Tools (68)

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
| `access_vbe_patch_proc` | Surgical find/replace within a procedure. **Atomic by default** (a failed patch writes nothing), case-insensitive anchors, optional `require_unique`, whitespace-tolerant fallback matching + contextual error messages when patches fail. `proc_name='(Declarations)'` targets the declarations section |
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
| `access_lint_form` | Deterministic check for objectively-broken layout: contrast (WCAG), overlap, out-of-bounds, truncation, sibling inconsistency, zero-size/invisible. Returns `verdict` PASS/REVIEW/FAIL. Also runs automatically on every control edit |

> ### ⚠️ A note on `access_lint_form` — manage your expectations
>
> **This is NOT a designer and there is zero super-design here.** Don't expect it to make a form *look good*, suggest a nice palette, or have any taste — it has none and never will.
>
> It is a **dumb, deterministic verifier of the obvious, easy-to-check stuff**: is the text the same colour as its background? do two controls physically overlap? does a caption not fit its box? is something off the edge of the form, or zero pixels tall? That's it. Plain math — WCAG contrast ratios and rectangle intersection — with a pile of false-positive guards so it doesn't cry wolf.
>
> Think **seatbelt, not stylist**: it won't make the car pretty, it just stops you shipping a form with white text on a white background without noticing. It runs automatically on every control edit so those obvious mistakes surface on their own. If you were hoping for a UI-design AI, this isn't it (honest PRs to make it smarter are very welcome 😄).

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
| `access_list_linked_tables` | List linked tables with source table, connection string, ODBC flag. `name='X'` returns one table; `names_only=true` is a light listing (no connect strings — use it when hundreds of links overflow the result); `mask_password=true` masks `PWD=` |
| `access_relink_table` | Change connection string and refresh link — auto-saves credentials (`dbAttachSavePWD`) when UID/PWD detected. `relink_all=true` updates all tables with the same original connection. `refresh=true` re-reads the schema using the table's own connect string (no `new_connect`, password never dumped) |

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
| `access_vbe_check_syntax` | Static structural check of the already-open VBA project — no decompile, nothing discarded. The safe post-edit check. Not a compiler: it does not resolve identifiers, types or references |
| `access_compile_vba` | Compile and save all VBA modules. Optional `timeout` to auto-dismiss error MsgBox. **Decompiles first and can discard unsaved VBA** — for a quick check use `access_vbe_check_syntax` |

### VBA & macro execution

> ⚠️ **Disabled by default (v0.7.51).** These three tools run arbitrary code and
> are gated behind the `MCP_ACCESS_ALLOW_CODE_EXEC` environment variable. See
> [Security](#security) to enable them.

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

## Security

This is a **local stdio server** with no network surface, so there is no login
by design — see [SECURITY.md](SECURITY.md) for the full threat model. The main
risk is **prompt injection**: an agent tricked (via `db_path` or content it reads
out of the database) into calling a code-execution tool.

**Code execution is disabled by default (v0.7.51).** The three tools that run
arbitrary VBA/Shell — `access_run_vba`, `access_eval_vba`, `access_run_macro` —
are hidden and rejected unless you opt in with an environment variable. To
re-enable, add it to this server's `env` in your MCP client config and **restart**:

```json
"env": { "MCP_ACCESS_ALLOW_CODE_EXEC": "1" }
```

Enabling grants arbitrary OS command execution; only point the server at trusted
databases. See [SECURITY.md](SECURITY.md) for details and how to report issues.

## Environment variables

All three are read from the server process, so they go in the `env` block of
your MCP client config and take effect on **restart**.

| Variable | Default | Effect |
|----------|---------|--------|
| `MCP_ACCESS_ALLOW_CODE_EXEC` | *off* | Set to `1`/`true`/`yes`/`on` to enable `access_run_vba`, `access_eval_vba` and `access_run_macro`. Fails **closed**: anything else keeps them disabled. |
| `MCP_ACCESS_SHIFT_BYPASS` | *on* | Set to `0`/`false`/`no`/`off` to stop holding SHIFT during `OpenCurrentDatabase` and `/decompile`. Fails **open**: anything else keeps the bypass. |
| `MCP_ACCESS_EXCLUSIVE` | *off* | Set to `1`/`true`/`yes`/`on` to open the session's database exclusively. Fails **closed**: anything else stays shared. |

**About the SHIFT bypass (v0.7.53).** Holding SHIFT is how the server skips a
database's AutoExec macro and startup form, but the key-down is a *global* OS
event — it is not scoped to Access, so anything you type anywhere on the machine
while it is held arrives capitalised (~0.3 s on every database switch, ~3 s per
decompile). Turn it off if that bothers you and your databases guard their own
startup, which is the cleaner fix and belongs in the database:

```vba
If Not Application.UserControl Then
  Exit Function
End If
```

`Application.UserControl` is False when Access was started via COM, so the
database opts itself out under automation and needs no bypass at all. With the
bypass off, `AutomationSecurity` and the dialog watchdog still apply — but an
unguarded AutoExec *macro object* will run.

**About exclusive opens (v0.7.55).** Shared opens — the default, and what every
version before this one did — cannot take a **design lock**. Attaching Data
Macros via `SaveAsText`/`LoadFromText`, or anything routing through
`DoCmd.OpenTable acViewDesign`, is refused for any table another Access session
has open, one table at a time and without stopping the run: it finishes,
reports success, and nothing changed. Turn the switch on and that becomes a
single visible failure when the database is opened.

It is off by default because it is genuinely exclusive: the session holds the
database between tool calls, so while the server is connected **nobody else can
open it**. That is right for a build or a development machine and wrong for a
shared front-end. `access_close` releases the database without stopping the
server.

The switch also verifies the mode instead of trusting it. Access treats
`Exclusive:=True` as a request: with the file already in use it opens it
*shared* without a word, and when someone else holds it exclusively it raises
nothing and leaves the session with no database at all. So the server checks
the lock file before opening (refusing without touching the session), checks it
again afterwards (closing the database if Access downgraded it), and names the
sessions holding it in the error. A stale lock file left by a crashed Access is
not mistaken for a live one. With the switch on the server also stops attaching
to an already-running Access instance, which would hold the file shared.

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

See [CHANGELOG.md](CHANGELOG.md) for the full release history.
