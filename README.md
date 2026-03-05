![mcpAccess](https://github.com/user-attachments/assets/9d37af7d-b829-4133-8518-44097c8ad8c9)


# mcp-access

MCP server for reading and editing Microsoft Access databases (`.accdb` / `.mdb`) via COM automation.

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

## Tools (20)

### Database objects

| Tool | Description |
|------|-------------|
| `access_list_objects` | List objects by type (`module`, `form`, `report`, `query`, `macro`, `all`) |
| `access_get_code` | Export an object's full definition as text |
| `access_set_code` | Import modified text back (creates or overwrites) |
| `access_export_structure` | Generate a Markdown index of all modules, forms, reports, queries |
| `access_close` | Close the COM session and release the `.accdb` file |

### SQL & tables

| Tool | Description |
|------|-------------|
| `access_execute_sql` | Run SQL via DAO — SELECT returns rows as JSON, others return affected count |
| `access_table_info` | Show table structure via DAO (fields, types, sizes, required, linked status) |

### VBE line-level editing

| Tool | Description |
|------|-------------|
| `access_vbe_get_lines` | Read a line range from a VBA module without exporting the whole file |
| `access_vbe_get_proc` | Get a procedure's code and position by name |
| `access_vbe_module_info` | List all procedures with their line numbers |
| `access_vbe_replace_lines` | Replace/insert/delete lines in a VBA module directly via VBE |
| `access_vbe_find` | Search text in a module and return matching lines with numbers |
| `access_vbe_search_all` | Search text across ALL modules/forms/reports in the database at once |
| `access_vbe_replace_proc` | Replace a full procedure by name (auto-calculates line bounds) |
| `access_vbe_append` | Append code at the end of a module |

### Form & report controls

| Tool | Description |
|------|-------------|
| `access_list_controls` | List direct controls of a form/report with key properties |
| `access_get_control` | Get the full definition block of a specific control |
| `access_create_control` | Create a new control via COM in design view |
| `access_delete_control` | Delete a control via COM |
| `access_set_control_props` | Modify control properties via COM in design view |

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

## Notes

- Access runs visible (`Visible = True`) so VBE COM access works correctly.
- One Access instance is shared across all tool calls (singleton session). Opening a different `.accdb` closes the previous one.
- `access_get_code` strips binary sections (`PrtMip`, `PrtDevMode`, etc.) from form/report exports — `access_set_code` restores them automatically before importing.
- All VBE line numbers are 1-based.

## Changelog

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
