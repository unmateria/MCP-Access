"""
MCP Tool definitions (58 tools) and schema utilities.
"""

import mcp.types as types

TOOLS = [
    types.Tool(
        name="access_list_objects",
        description="Lists database objects by type (table, module, form, report, query, macro, all). System tables (MSys*, ~*) are filtered out.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {
                    "type": "string",
                    "enum": ["all", "table", "module", "form", "report", "query", "macro"],
                    "default": "all",
                },
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_get_code",
        description=(
            "Reads code/definition of an Access object. "
            "Modules: .bas code. Forms/reports: internal format (props + VBA)."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report", "query", "macro"]},
                "object_name": {"type": "string", "description": "Object name"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_set_code",
        description=(
            "Imports code into the database. Overwrites if exists, creates if not. "
            "Call access_get_code first to read the original. "
            "For forms/reports: supports CodeBehindForm/CodeBehindReport (VBA is injected via VBE). "
            "Automatic backup and restore on import failure."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report", "query", "macro"]},
                "object_name": {"type": "string", "description": "Object name"},
                "code": {"type": "string", "description": "Full object content"},
            },
            "required": ["db_path", "object_type", "object_name", "code"],
        },
    ),
    types.Tool(
        name="access_execute_sql",
        description=(
            "Executes SQL via DAO. SELECT returns JSON rows (default limit: 500). "
            "INSERT/UPDATE return affected_rows. "
            "DELETE/DROP/ALTER require confirm_destructive=true."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "sql": {"type": "string", "description": "SQL statement"},
                "limit": {"type": "integer", "default": 500,
                          "description": "Max rows for SELECT (default: 500, max: 10000)"},
                "confirm_destructive": {
                    "type": "boolean", "default": False,
                    "description": "Required for DELETE/DROP/TRUNCATE/ALTER",
                },
            },
            "required": ["db_path", "sql"],
        },
    ),
    types.Tool(
        name="access_table_info",
        description="Table structure via DAO: fields, types, size, required, record_count, is_linked.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "table_name": {"type": "string", "description": "Table name"},
            },
            "required": ["db_path", "table_name"],
        },
    ),
    types.Tool(
        name="access_export_structure",
        description=(
            "Generates Markdown with database structure: modules with signatures, forms, reports, queries, macros. "
            "Writes to disk and returns the content."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "output_path": {"type": "string", "description": "Output .md path (default: db_structure.md next to the DB)"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_close",
        description="Closes the COM session and releases the .accdb/.mdb file.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    # -- VBE line-level tools ------------------------------------------------
    types.Tool(
        name="access_vbe_get_lines",
        description="Reads a range of lines from a VBA module via VBE COM. Provide either count or end_line (not both).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Object name"},
                "start_line":  {"type": "integer", "description": "First line (1-based)"},
                "count":       {"type": "integer", "description": "Number of lines to read"},
                "end_line":    {"type": "integer", "description": "Last line (1-based, alternative to count)"},
            },
            "required": ["db_path", "object_type", "object_name", "start_line"],
        },
    ),
    types.Tool(
        name="access_vbe_get_proc",
        description="Code of a VBA procedure by name. Returns start_line, body_line, count, code.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Object name"},
                "proc_name":   {"type": "string", "description": "Sub/Function/Property name"},
            },
            "required": ["db_path", "object_type", "object_name", "proc_name"],
        },
    ),
    types.Tool(
        name="access_vbe_module_info",
        description="Index of procedures in a VBA module: total_lines, procs [{name, start_line, body_line, count}].",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Object name"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_vbe_replace_lines",
        description=(
            "Replaces lines in a VBA module via VBE. "
            "count=0: insertion. new_code='': deletion. Validates bounds automatically. "
            "Batch mode: pass 'operations' (list of {start_line, count, new_code}) "
            "to execute multiple operations in 1 call (auto-sorted bottom-to-top)."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Object name"},
                "start_line":  {"type": "integer", "description": "First line (1-based). Ignored if operations present."},
                "count":       {"type": "integer", "description": "Lines to delete (0 = insert). Ignored if operations present."},
                "new_code":    {"type": "string",  "description": "New code ('' = delete). Ignored if operations present."},
                "operations":  {
                    "type": "array",
                    "description": "Batch mode: list of operations. Each: {start_line, count, new_code}. Auto-sorted bottom-to-top.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "start_line": {"type": "integer"},
                            "count": {"type": "integer"},
                            "new_code": {"type": "string"},
                        },
                        "required": ["start_line", "count"],
                    },
                },
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_vbe_find",
        description=(
            "Searches for text or regex in a VBA module. Returns matches [{line, content, proc}]. "
            "Each match includes 'proc' (procedure name). "
            "Optional: proc_name to limit search to a single procedure."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Object name"},
                "search_text": {"type": "string", "description": "Text or regex pattern to search"},
                "match_case":  {"type": "boolean", "default": False},
                "use_regex": {"type": "boolean", "default": False,
                              "description": "true = interpret search_text as regex"},
                "proc_name": {"type": "string",
                              "description": "Optional: limit search to this procedure"},
            },
            "required": ["db_path", "object_type", "object_name", "search_text"],
        },
    ),
    types.Tool(
        name="access_vbe_search_all",
        description="Searches for text or regex in ALL VBA modules (modules, forms, reports) in the database.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "search_text": {"type": "string", "description": "Text or regex pattern to search"},
                "match_case":  {"type": "boolean", "default": False},
                "max_results": {"type": "integer", "default": 100,
                                "description": "Max total matches (default: 100)"},
                "use_regex": {"type": "boolean", "default": False,
                              "description": "true = interpret search_text as regex"},
            },
            "required": ["db_path", "search_text"],
        },
    ),
    types.Tool(
        name="access_search_queries",
        description="Searches for text or regex in the SQL of ALL queries. Returns [{query_name, sql}].",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "search_text": {"type": "string", "description": "Text or regex pattern to search in SQL"},
                "match_case": {"type": "boolean", "default": False},
                "max_results": {"type": "integer", "default": 100,
                                "description": "Max queries to return (default: 100)"},
                "use_regex": {"type": "boolean", "default": False,
                              "description": "true = interpret search_text as regex"},
            },
            "required": ["db_path", "search_text"],
        },
    ),
    types.Tool(
        name="access_vbe_replace_proc",
        description="Replaces an entire VBA procedure by name. new_code='' deletes it.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Object name"},
                "proc_name":   {"type": "string", "description": "Sub/Function/Property name"},
                "new_code":    {"type": "string", "description": "New code ('' = delete)"},
            },
            "required": ["db_path", "object_type", "object_name", "proc_name", "new_code"],
        },
    ),
    types.Tool(
        name="access_vbe_patch_proc",
        description=(
            "Surgical find/replace WITHIN a VBA procedure. "
            "More efficient than replace_proc when only a few lines change "
            "in a large proc. patches: [{find, replace}]."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Object name"},
                "proc_name": {"type": "string", "description": "Sub/Function/Property name"},
                "patches": {
                    "type": "array",
                    "description": "List of find/replace to apply within the proc",
                    "items": {
                        "type": "object",
                        "properties": {
                            "find": {"type": "string", "description": "Text to find (literal, first occurrence)"},
                            "replace": {"type": "string", "description": "Replacement text ('' = delete)"},
                        },
                        "required": ["find"],
                    },
                },
            },
            "required": ["db_path", "object_type", "object_name", "proc_name", "patches"],
        },
    ),
    types.Tool(
        name="access_vbe_append",
        description="Appends code to the end of a VBA module.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report"]},
                "object_name": {"type": "string", "description": "Object name"},
                "new_code":    {"type": "string", "description": "Code to append"},
            },
            "required": ["db_path", "object_type", "object_name", "new_code"],
        },
    ),
    # -- Control-level tools -------------------------------------------------
    types.Tool(
        name="access_list_controls",
        description="Lists controls of a form/report with name, type, caption, control_source, position.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Form/report name"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    types.Tool(
        name="access_get_control",
        description="Full definition (Begin...End) of a control by name.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Form/report name"},
                "control_name": {"type": "string", "description": "Control name"},
            },
            "required": ["db_path", "object_type", "object_name", "control_name"],
        },
    ),
    types.Tool(
        name="access_create_control",
        description=(
            "Creates a control on a form/report via COM. "
            "control_type: name or number (e.g.: 119=acCustomControl for ActiveX, 128=acWebBrowser native). "
            "Special props: section (0=Detail,1=Header,2=Footer,3=PageHeader,4=PageFooter "
            "or name: 'detail','header','footer','reportheader','pageheader'...), "
            "parent, column_name, left, top, width, height. "
            "For ActiveX (type 119), use class_name with the ProgID (e.g.: 'Shell.Explorer.2')."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Form/report name"},
                "control_type": {"type": "string", "description": "'CommandButton', 'TextBox', 'Label', 'CustomControl'(119), 'WebBrowser'(128)... or number"},
                "props": {
                    "type": "object",
                    "description": "Properties: section, parent, column_name, left, top, width, height, Name, Caption, etc.",
                    "additionalProperties": True,
                },
                "class_name": {
                    "type": "string",
                    "description": "ProgID for ActiveX (type 119). E.g.: 'Shell.Explorer.2', 'MSCAL.Calendar.7'. Initializes the OLE control.",
                },
            },
            "required": ["db_path", "object_type", "object_name", "control_type", "props"],
        },
    ),
    types.Tool(
        name="access_delete_control",
        description="Deletes a control from a form/report via COM.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Form/report name"},
                "control_name": {"type": "string", "description": "Control name"},
            },
            "required": ["db_path", "object_type", "object_name", "control_name"],
        },
    ),
    types.Tool(
        name="access_export_text",
        description="Exports form/report/module as text (SaveAsText). Does NOT open in Design view — does not recalculate positions. File is UTF-16 LE.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report", "module", "query", "macro"]},
                "object_name": {"type": "string", "description": "Object name"},
                "output_path": {"type": "string", "description": "Output file path (.txt)"},
            },
            "required": ["db_path", "object_type", "object_name", "output_path"],
        },
    ),
    types.Tool(
        name="access_import_text",
        description="Imports form/report/module from text (LoadFromText). REPLACES if exists. Does NOT open in Design view — does not recalculate positions. For forms/reports with CodeBehindForm: separates VBA automatically and injects via VBE (same as access_set_code).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report", "module", "query", "macro"]},
                "object_name": {"type": "string", "description": "Object name"},
                "input_path": {"type": "string", "description": "Input text file path (.txt) in UTF-16 LE format"},
            },
            "required": ["db_path", "object_type", "object_name", "input_path"],
        },
    ),
    types.Tool(
        name="access_set_control_props",
        description="Modifies properties of a control via COM. Numeric/boolean values are converted automatically.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Form/report name"},
                "control_name": {"type": "string", "description": "Control name"},
                "props": {
                    "type": "object",
                    "description": "Properties to modify: {Caption: 'X', Left: 1000, Visible: true, ...}",
                    "additionalProperties": True,
                },
            },
            "required": ["db_path", "object_type", "object_name", "control_name", "props"],
        },
    ),
    types.Tool(
        name="access_set_form_property",
        description="Sets form/report-level properties (RecordSource, Caption, DefaultView, HasModule, etc.) via COM in Design view.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Form/report name"},
                "props": {
                    "type": "object",
                    "description": "Properties to modify: {RecordSource: 'Table', Caption: 'Title', HasModule: true, ...}",
                    "additionalProperties": True,
                },
            },
            "required": ["db_path", "object_type", "object_name", "props"],
        },
    ),
    # -- Database properties -------------------------------------------------
    types.Tool(
        name="access_get_db_property",
        description="Reads a database property (CurrentDb.Properties) or Access option (GetOption).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "name": {"type": "string", "description": "Property name (e.g.: AppTitle, StartupForm, AllowBypassKey)"},
            },
            "required": ["db_path", "name"],
        },
    ),
    types.Tool(
        name="access_set_db_property",
        description="Sets a database property or Access option. Creates the property if it does not exist.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "name": {"type": "string", "description": "Property name"},
                "value": {"description": "Value (string, number or boolean)"},
                "prop_type": {"type": "integer", "description": "DAO type for CreateProperty (1=Boolean, 4=Long, 10=Text). Auto-detected if omitted"},
            },
            "required": ["db_path", "name", "value"],
        },
    ),
    # -- Linked tables -------------------------------------------------------
    types.Tool(
        name="access_list_linked_tables",
        description="Lists linked tables with source_table, connect_string, is_odbc.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_relink_table",
        description="Changes the connect string of a linked table and refreshes. relink_all=true updates all tables with the same original connection.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "table_name": {"type": "string", "description": "Linked table name"},
                "new_connect": {"type": "string", "description": "New connection string"},
                "relink_all": {"type": "boolean", "default": False, "description": "true = relink all tables with the same original connection"},
            },
            "required": ["db_path", "table_name", "new_connect"],
        },
    ),
    # -- Relationships -------------------------------------------------------
    types.Tool(
        name="access_list_relationships",
        description="Lists relationships between tables: name, tables, fields, cascade flags.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_create_relationship",
        description="Creates a relationship between two tables. attributes: 256=cascade update, 4096=cascade delete (combinable with OR).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "name": {"type": "string", "description": "Relationship name"},
                "table": {"type": "string", "description": "Primary table (one side)"},
                "foreign_table": {"type": "string", "description": "Foreign table (many side)"},
                "fields": {
                    "type": "array",
                    "description": "[{local: 'ID', foreign: 'FK_ID'}, ...]",
                    "items": {
                        "type": "object",
                        "properties": {"local": {"type": "string"}, "foreign": {"type": "string"}},
                        "required": ["local", "foreign"],
                    },
                },
                "attributes": {"type": "integer", "default": 0, "description": "Bitmask: 256=cascade update, 4096=cascade delete"},
            },
            "required": ["db_path", "name", "table", "foreign_table", "fields"],
        },
    ),
    # -- VBA References ------------------------------------------------------
    types.Tool(
        name="access_list_references",
        description="Lists VBA references: name, GUID, path, is_broken, built_in.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_manage_reference",
        description="Adds (add) or removes (remove) a VBA reference. add: requires guid or path. remove: requires name.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "action": {"type": "string", "enum": ["add", "remove"]},
                "name": {"type": "string", "description": "[remove] Reference name"},
                "path": {"type": "string", "description": "[add] Path to .dll/.tlb/.olb"},
                "guid": {"type": "string", "description": "[add] Type library GUID"},
                "major": {"type": "integer", "default": 0, "description": "[add+guid] Major version"},
                "minor": {"type": "integer", "default": 0, "description": "[add+guid] Minor version"},
            },
            "required": ["db_path", "action"],
        },
    ),
    # -- Compact & Repair ----------------------------------------------------
    types.Tool(
        name="access_compact_repair",
        description="Compacts and repairs the database. Closes, compacts to temp, replaces original and reopens. Returns sizes.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_decompile_compact",
        description=(
            "Removes orphan VBA p-code via /decompile, recompiles and compacts. "
            "Typical reduction 60-70% on heavily edited front-end files. "
            "Use when the .accdb exceeds 30-40 MB without having local tables with data."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
            },
            "required": ["db_path"],
        },
    ),
    # -- Query management ----------------------------------------------------
    types.Tool(
        name="access_manage_query",
        description=(
            "Manages QueryDefs: create, modify, delete (requires confirm=true), rename, get_sql. "
            "create/modify require sql."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "action": {"type": "string", "enum": ["create", "modify", "delete", "rename", "get_sql"]},
                "query_name": {"type": "string", "description": "Query name"},
                "sql": {"type": "string", "description": "[create/modify] Query SQL"},
                "new_name": {"type": "string", "description": "[rename] New name"},
                "confirm": {"type": "boolean", "default": False, "description": "[delete] Confirm deletion"},
            },
            "required": ["db_path", "action", "query_name"],
        },
    ),
    # -- Indexes -------------------------------------------------------------
    types.Tool(
        name="access_list_indexes",
        description="Lists indexes of a table: name, fields, primary, unique, foreign.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "table_name": {"type": "string", "description": "Table name"},
            },
            "required": ["db_path", "table_name"],
        },
    ),
    types.Tool(
        name="access_manage_index",
        description="Creates or deletes an index. create requires fields [{name, order?}]. primary/unique optional.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "table_name": {"type": "string", "description": "Table name"},
                "action": {"type": "string", "enum": ["create", "delete"]},
                "index_name": {"type": "string", "description": "Index name"},
                "fields": {
                    "type": "array", "description": "[create] [{name: 'Field', order: 'asc'|'desc'}]",
                    "items": {
                        "type": "object",
                        "properties": {"name": {"type": "string"}, "order": {"type": "string", "default": "asc"}},
                        "required": ["name"],
                    },
                },
                "primary": {"type": "boolean", "default": False},
                "unique": {"type": "boolean", "default": False},
            },
            "required": ["db_path", "table_name", "action", "index_name"],
        },
    ),
    # -- Compile VBA ---------------------------------------------------------
    types.Tool(
        name="access_compile_vba",
        description="Compiles and saves all VBA modules. Returns status or compilation error. With timeout, automatically dismisses compilation error MsgBox.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "timeout": {"type": "integer", "description": "Timeout in seconds. If exceeded, automatically dismisses compilation error MsgBox"},
            },
            "required": ["db_path"],
        },
    ),
    # -- Run macro -----------------------------------------------------------
    types.Tool(
        name="access_run_macro",
        description="Executes an Access macro by name.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "macro_name": {"type": "string", "description": "Macro name"},
            },
            "required": ["db_path", "macro_name"],
        },
    ),
    # -- Output report -------------------------------------------------------
    types.Tool(
        name="access_output_report",
        description="Exports a report to PDF, XLSX, RTF or TXT. output_path auto-generated if omitted.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "report_name": {"type": "string", "description": "Report name"},
                "output_path": {"type": "string", "description": "Output path (auto if omitted)"},
                "format": {"type": "string", "default": "pdf", "description": "pdf, xlsx, rtf, txt"},
            },
            "required": ["db_path", "report_name"],
        },
    ),
    # -- Transfer data -------------------------------------------------------
    types.Tool(
        name="access_transfer_data",
        description=(
            "Import/export data between Access and Excel/CSV. "
            "file_type: xlsx or csv. range only for Excel, spec_name only for CSV."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "action": {"type": "string", "enum": ["import", "export"]},
                "file_path": {"type": "string", "description": "Path to Excel/CSV file"},
                "table_name": {"type": "string", "description": "Access table name"},
                "has_headers": {"type": "boolean", "default": True},
                "file_type": {"type": "string", "default": "xlsx", "description": "xlsx or csv"},
                "range": {"type": "string", "description": "[xlsx] Range e.g.: Sheet1!A1:D100"},
                "spec_name": {"type": "string", "description": "[csv] Import/Export spec saved in Access"},
            },
            "required": ["db_path", "action", "file_path", "table_name"],
        },
    ),
    # -- Field properties ----------------------------------------------------
    types.Tool(
        name="access_get_field_properties",
        description="Reads all properties of a field: DefaultValue, ValidationRule, Description, Format, etc.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "table_name": {"type": "string", "description": "Table name"},
                "field_name": {"type": "string", "description": "Field name"},
            },
            "required": ["db_path", "table_name", "field_name"],
        },
    ),
    types.Tool(
        name="access_set_field_property",
        description="Sets a field property. Creates the property if it does not exist.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "table_name": {"type": "string", "description": "Table name"},
                "field_name": {"type": "string", "description": "Field name"},
                "property_name": {"type": "string", "description": "Property name (e.g.: Description, DefaultValue)"},
                "value": {"description": "Value (string, number or boolean)"},
            },
            "required": ["db_path", "table_name", "field_name", "property_name", "value"],
        },
    ),
    # -- Startup options -----------------------------------------------------
    types.Tool(
        name="access_list_startup_options",
        description="Lists the 14 common startup options with their current values.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
            },
            "required": ["db_path"],
        },
    ),
    # -- Create database -----------------------------------------------------
    types.Tool(
        name="access_create_database",
        description="Creates an empty Access database (.accdb). Error if the file already exists.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path for the new .accdb"},
            },
            "required": ["db_path"],
        },
    ),
    # -- Create table via DAO ------------------------------------------------
    types.Tool(
        name="access_create_table",
        description=(
            "Creates an Access table via DAO with full support: types, defaults, "
            "descriptions and primary key — all in a single call. "
            "More robust than CREATE TABLE via SQL, which does not support DEFAULT or YESNO in Jet DDL. "
            "Each field accepts: name, type, size, required, primary_key, default, description. "
            "Valid types: autonumber, long, integer, short, byte, text, memo, currency, "
            "double, single, datetime, boolean/yesno/bit, guid, ole, bigint."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "table_name": {"type": "string", "description": "Table name"},
                "fields": {
                    "type": "array",
                    "description": "List of fields [{name, type, size?, required?, primary_key?, default?, description?}]",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string"},
                            "type": {"type": "string", "default": "text"},
                            "size": {"type": "integer"},
                            "required": {"type": "boolean", "default": False},
                            "primary_key": {"type": "boolean", "default": False},
                            "default": {"description": "Default value (string, number or boolean)"},
                            "description": {"type": "string"},
                        },
                        "required": ["name"],
                    },
                },
            },
            "required": ["db_path", "table_name", "fields"],
        },
    ),
    # -- Alter table via DAO -------------------------------------------------
    types.Tool(
        name="access_alter_table",
        description=(
            "Modifies the structure of an Access table via DAO. "
            "Actions: add_field (with type, size, default, description), "
            "delete_field (requires confirm=true), rename_field. "
            "More robust than ALTER TABLE via SQL in Jet."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "table_name": {"type": "string", "description": "Table name"},
                "action": {"type": "string", "enum": ["add_field", "delete_field", "rename_field"]},
                "field_name": {"type": "string", "description": "Field name"},
                "new_name": {"type": "string", "description": "[rename_field] New name"},
                "field_type": {"type": "string", "default": "text", "description": "[add_field] Field type"},
                "size": {"type": "integer", "description": "[add_field] Size for Text"},
                "required": {"type": "boolean", "default": False},
                "default": {"description": "[add_field] Default value"},
                "description": {"type": "string", "description": "[add_field] Field description"},
                "confirm": {"type": "boolean", "default": False, "description": "[delete_field] Confirm deletion"},
            },
            "required": ["db_path", "table_name", "action", "field_name"],
        },
    ),
    # -- Create form ---------------------------------------------------------
    types.Tool(
        name="access_create_form",
        description=(
            "Creates a new form in the database. Avoids the 'Save As' MsgBox that blocks COM "
            "when using CreateForm() directly. Option has_header to create with "
            "header/footer section."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "form_name": {"type": "string", "description": "Name of the form to create"},
                "has_header": {"type": "boolean", "default": False, "description": "Create with header/footer section"},
            },
            "required": ["db_path", "form_name"],
        },
    ),
    # -- Delete object -------------------------------------------------------
    types.Tool(
        name="access_delete_object",
        description="Deletes an Access object (module, form, report, query, macro). Requires confirm=true.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["module", "form", "report", "query", "macro"]},
                "object_name": {"type": "string", "description": "Name of the object to delete"},
                "confirm": {"type": "boolean", "default": False, "description": "Required true to confirm deletion"},
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    # -- Run VBA -------------------------------------------------------------
    types.Tool(
        name="access_run_vba",
        description="Executes a VBA Sub/Function. Supports 3 syntaxes: 'Module.MySub' (standard module via Application.Run), 'MySub' (same), 'Forms.FormName.Method' (form module via COM — form must be open). Returns result if Function. With timeout, automatically dismisses MsgBox/InputBox if exceeded.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "procedure": {"type": "string", "description": "Procedure name: 'MyModule.MySub', 'MySub', or 'Forms.FormName.Method'"},
                "args": {
                    "type": "array",
                    "description": "Optional arguments (max 30)",
                    "items": {},
                },
                "timeout": {
                    "type": "integer",
                    "description": "Timeout in seconds. If exceeded, automatically dismisses MsgBox/InputBox dialogs and returns error",
                },
            },
            "required": ["db_path", "procedure"],
        },
    ),
    # -- Eval VBA ------------------------------------------------------------
    types.Tool(
        name="access_eval_vba",
        description="Evaluates a VBA/Access expression via Application.Eval. Works with: domain functions (DLookup, DCount...), built-in VBA functions (Date(), Len()...), form properties of OPEN forms (Forms!frmX.Prop), public functions in standard modules. Does NOT work with: class instances (even default/predeclared), variables, Subs. For class instances use access_run_vba with a wrapper function, or this tool will attempt an automatic fallback via a temp VBA module.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "expression": {
                    "type": "string",
                    "description": "Expression to evaluate (e.g.: 'Forms!frmX.MARGEN_SEG', 'Date()', 'DLookup(\"Empresa\",\"Ventas\",\"numc=1\")')",
                },
            },
            "required": ["db_path", "expression"],
        },
    ),
    # -- Delete relationship -------------------------------------------------
    types.Tool(
        name="access_delete_relationship",
        description="Deletes a relationship between tables by name.",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "name": {"type": "string", "description": "Name of the relationship to delete"},
            },
            "required": ["db_path", "name"],
        },
    ),
    # -- Find usages ---------------------------------------------------------
    types.Tool(
        name="access_find_usages",
        description=(
            "Searches for text or regex in VBA, query SQL and control properties "
            "(ControlSource, RecordSource, RowSource, DefaultValue, ValidationRule). "
            "Returns results grouped: vba_matches, query_matches, control_matches."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "search_text": {"type": "string", "description": "Text or regex pattern to search"},
                "match_case": {"type": "boolean", "default": False},
                "max_results": {"type": "integer", "default": 200,
                                "description": "Max total matches (default: 200)"},
                "use_regex": {"type": "boolean", "default": False,
                              "description": "true = interpret search_text as regex"},
            },
            "required": ["db_path", "search_text"],
        },
    ),
    # -- Batch SQL -----------------------------------------------------------
    types.Tool(
        name="access_execute_batch",
        description=(
            "Executes multiple SQL statements in a single call. "
            "Each statement can be SELECT (returns rows, limit 100) or "
            "INSERT/UPDATE/DELETE (returns affected_rows). "
            "stop_on_error=true stops at first error. "
            "DELETE/DROP/TRUNCATE/ALTER require confirm_destructive=true."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "statements": {
                    "type": "array",
                    "description": "List of SQL statements [{sql: str, label?: str}, ...]",
                    "items": {
                        "type": "object",
                        "properties": {
                            "sql": {"type": "string", "description": "SQL statement"},
                            "label": {"type": "string",
                                      "description": "Optional label to identify the statement"},
                        },
                        "required": ["sql"],
                    },
                },
                "stop_on_error": {
                    "type": "boolean", "default": True,
                    "description": "true = stop at first error (default: true)",
                },
                "confirm_destructive": {
                    "type": "boolean", "default": False,
                    "description": "Required for DELETE/DROP/TRUNCATE/ALTER",
                },
            },
            "required": ["db_path", "statements"],
        },
    ),
    # -- Get form/report property --------------------------------------------
    types.Tool(
        name="access_get_form_property",
        description=(
            "Reads properties of a form/report (RecordSource, Caption, DefaultView, "
            "HasModule, etc.). If property_names is omitted, returns all readable properties."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Form/report name"},
                "property_names": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of properties to read. Omit to read all.",
                },
            },
            "required": ["db_path", "object_type", "object_name"],
        },
    ),
    # -- Set multiple controls -----------------------------------------------
    types.Tool(
        name="access_set_multiple_controls",
        description=(
            "Modifies properties of multiple controls on a form/report in a single "
            "operation. Opens in Design view once, applies changes, saves and closes."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {"type": "string", "enum": ["form", "report"]},
                "object_name": {"type": "string", "description": "Form/report name"},
                "controls": {
                    "type": "array",
                    "description": "List of controls [{name: str, props: {prop: val}}, ...]",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "description": "Control name"},
                            "props": {
                                "type": "object",
                                "additionalProperties": True,
                                "description": "Properties to modify {Caption: 'X', Left: 1000, ...}",
                            },
                        },
                        "required": ["name", "props"],
                    },
                },
            },
            "required": ["db_path", "object_type", "object_name", "controls"],
        },
    ),
    # -- Tips / knowledge base -----------------------------------------------
    types.Tool(
        name="access_tips",
        description="Tips and gotchas for working with Access via MCP. Topics: eval, controls, gotchas, sql, vbe, compile, design. Without topic returns the list.",
        inputSchema={
            "type": "object",
            "properties": {
                "topic": {"type": "string", "description": "Topic: eval, controls, gotchas, sql, vbe, compile, design (empty = list of topics)"},
            },
        },
    ),
    # -- Screenshot + UI Automation ------------------------------------------
    types.Tool(
        name="access_screenshot",
        description="Captures the Access window as PNG. Optionally opens a form/report before capturing. Returns path, dimensions and metadata. wait_ms pumps Windows messages (Timer events, ActiveX init).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "object_type": {
                    "type": "string",
                    "enum": ["form", "report"],
                    "description": "Object type to open before capturing (optional)",
                },
                "object_name": {
                    "type": "string",
                    "description": "Form/report name to open (requires object_type)",
                },
                "output_path": {
                    "type": "string",
                    "description": "Output PNG path (auto if omitted)",
                },
                "wait_ms": {
                    "type": "integer",
                    "default": 300,
                    "description": "Wait in ms after opening object (0 = instant)",
                },
                "max_width": {
                    "type": "integer",
                    "default": 1920,
                    "description": "Max image width in px",
                },
                "open_timeout_sec": {
                    "type": "integer",
                    "default": 30,
                    "description": "Max seconds waiting for OpenForm/OpenReport. If Form_Load takes longer (slow OpenRecordset), ESC is sent and error is raised. Default 30.",
                },
            },
            "required": ["db_path"],
        },
    ),
    types.Tool(
        name="access_ui_click",
        description="Click at image coordinates on the Access window. Coordinates are relative to a previous screenshot (image_width required for scaling).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "x": {"type": "integer", "description": "X coordinate in image space"},
                "y": {"type": "integer", "description": "Y coordinate in image space"},
                "image_width": {
                    "type": "integer",
                    "description": "Screenshot width used for the coordinates",
                },
                "click_type": {
                    "type": "string",
                    "enum": ["left", "double", "right"],
                    "default": "left",
                    "description": "Click type: left, double, right",
                },
                "wait_after_ms": {
                    "type": "integer",
                    "default": 200,
                    "description": "Wait in ms after click",
                },
            },
            "required": ["db_path", "x", "y", "image_width"],
        },
    ),
    types.Tool(
        name="access_ui_type",
        description="Types text or sends keyboard shortcuts to the Access window. Use 'text' for normal text, 'key' for special keys (enter, tab, escape, f1-f12, etc.), 'modifiers' for combinations (ctrl, shift, alt).",
        inputSchema={
            "type": "object",
            "properties": {
                "db_path": {"type": "string", "description": "Path to .accdb/.mdb file"},
                "text": {
                    "type": "string",
                    "description": "Text to type (normal characters)",
                },
                "key": {
                    "type": "string",
                    "description": "Special key: enter, tab, escape, backspace, delete, up, down, left, right, home, end, f1-f12, space, pageup, pagedown",
                },
                "modifiers": {
                    "type": "string",
                    "description": "Modifiers: ctrl, shift, alt, ctrl+shift — combined with key",
                },
                "wait_after_ms": {
                    "type": "integer",
                    "default": 100,
                    "description": "Wait in ms after typing",
                },
            },
            "required": ["db_path"],
        },
    ),
]

# ---------------------------------------------------------------------------
# Schema fixup: accept "integer" and "boolean" as strings too
# ---------------------------------------------------------------------------

def _fixup_schema(schema: dict) -> None:
    """Recursively change {"type":"integer"} -> {"type":["integer","string"]}
    and {"type":"boolean"} -> {"type":["boolean","string"]} in a JSON Schema."""
    if not isinstance(schema, dict):
        return
    t = schema.get("type")
    if t == "integer":
        schema["type"] = ["integer", "string"]
    elif t == "boolean":
        schema["type"] = ["boolean", "string"]
    for key in ("properties", "patternProperties"):
        block = schema.get(key)
        if isinstance(block, dict):
            for v in block.values():
                _fixup_schema(v)
    for key in ("items", "additionalProperties"):
        sub = schema.get(key)
        if isinstance(sub, dict):
            _fixup_schema(sub)

for _tool in TOOLS:
    _fixup_schema(_tool.inputSchema)

# Build schema index for argument coercion at call time
_TOOL_SCHEMA_INDEX: dict[str, dict] = {t.name: t.inputSchema for t in TOOLS}


def coerce_arguments(name: str, arguments: dict) -> dict:
    """Convert string arguments to the expected type based on the tool schema."""
    schema = _TOOL_SCHEMA_INDEX.get(name)
    if not schema:
        return arguments
    props = schema.get("properties", {})
    for key, val in list(arguments.items()):
        if not isinstance(val, str):
            continue
        pdef = props.get(key, {})
        ptypes = pdef.get("type")
        if isinstance(ptypes, list):
            if "integer" in ptypes:
                try:
                    arguments[key] = int(val)
                except (ValueError, TypeError):
                    pass
            elif "boolean" in ptypes:
                arguments[key] = val.lower() in ("true", "1", "yes")
    return arguments
