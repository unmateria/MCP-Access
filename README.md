![mcpAccess](https://github.com/user-attachments/assets/9fad6eb6-2806-43dd-8fe6-73ea645d56aa)


# MCP-Access

MCP server for reading and editing Microsoft Access databases (`.accdb` / `.mdb`) via COM automation.

## Requirements

- Windows (COM automation is Windows-only)
- Microsoft Access installed (any version that supports VBE, 2010+)
- Python 3.9+
- *"Trust access to the VBA project object model"* enabled in Access Trust Center or use enable_vba_trust.ps1 powershell script

## Installation

```bash
pip install mcp pywin32
```

### Enable VBA object model access in Access

`File → Options → Trust Center → Trust Center Settings → Macro Settings`
→ check **Trust access to the VBA project object model** 
or use enable_vba_trust.ps1 powershell script

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

## Tools

| Tool | Description |
|------|-------------|
| `access_list_objects` | List objects by type (`module`, `form`, `report`, `query`, `macro`, `all`) |
| `access_get_code` | Export an object's full definition as text |
| `access_set_code` | Import modified text back (creates or overwrites) |
| `access_execute_sql` | Run SQL via DAO — SELECT returns rows as JSON, others return affected count |
| `access_export_structure` | Generate a Markdown index of all modules, forms, reports, queries |
| `access_close` | Close the COM session and release the `.accdb` file |
| `access_vbe_get_lines` | Read a line range from a VBA module without exporting the whole file |
| `access_vbe_get_proc` | Get a procedure's code and position by name |
| `access_vbe_module_info` | List all procedures with their line numbers |
| `access_vbe_replace_lines` | Replace/insert/delete lines in a VBA module directly via VBE |
| `access_vbe_find` | Search text in a module and return matching lines with numbers |
| `access_list_controls` | List direct controls of a form/report with key properties |
| `access_get_control` | Get the full definition block of a specific control |
| `access_create_control` | Create a new control via COM in design view |
| `access_delete_control` | Delete a control via COM |
| `access_set_control_props` | Modify control properties via COM in design view |

## Typical VBA editing workflow

```
1. access_list_objects      → find the module or form name
2. access_vbe_module_info   → get procedure list and line numbers
3. access_vbe_get_proc      → read the specific procedure
4. access_vbe_replace_lines → apply targeted line-level changes
5. access_close             → release the file when done
```

For full object replacement (forms, reports, modules):
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
