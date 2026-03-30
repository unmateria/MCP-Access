"""Tips and gotchas for working with Access via MCP."""

_TIPS: dict[str, str] = {
    "eval": (
        "ac_eval_vba can query the Access Object Model without new tools:\n"
        "  Application.IsCompiled — check if VBA is compiled (no compile triggered)\n"
        "  SysCmd(10, 2, \"formName\") — check if form is open (acSysCmdGetObjectState=10, acForm=2)\n"
        "    Returns: 0=closed, 1=open, 2=new, 4=dirty, 8=has new record\n"
        "  Application.BrokenReference — True if any VBA reference is broken\n"
        "  Screen.ActiveForm.Name / Screen.ActiveControl.Name — active form/control\n"
        "  Forms.Count — number of open forms\n"
        "  TempVars(\"x\") — session-persistent variables (read/write across tools)\n"
        "  DLookup/DCount/DSum — domain aggregate functions\n"
        "  TypeName(expr) — inspect type of any expression\n"
        "  Eval only works for expressions/functions, NOT statements/Subs."
    ),
    "controls": (
        "FormatConditions: access_list_controls / access_get_control show\n"
        "  'format_conditions' count when a control has conditional formatting.\n"
        "  ConditionalFormat data in SaveAsText is binary (hex blobs, not readable).\n"
        "  To read details: use VBA via access_run_vba (FormatConditions collection in Design view).\n"
        "  To modify: write a temp VBA function with access_vbe_append, call with access_run_vba.\n"
        "  BE CAREFUL: modifying control properties may break existing conditional formatting.\n\n"
        "Control types for access_create_control:\n"
        "  119 = acCustomControl (ActiveX) — use class_name param for ProgID (e.g. 'Shell.Explorer.2')\n"
        "  128 = acWebBrowser (native, NOT ActiveX — no OLE needed)\n"
        "  Common: 100=Label, 109=TextBox, 106=ComboBox, 105=ListBox, 104=CommandButton,\n"
        "          110=CheckBox, 114=SubForm, 122=Image, 101=Rectangle"
    ),
    "gotchas": (
        "COM & ODBC:\n"
        "  dbSeeChanges (512) — REQUIRED in CurrentDb.Execute for DELETE/UPDATE on ODBC linked tables\n"
        "  LIKE wildcards — use % for ODBC (not *)\n"
        "  COM from threads — MUST call pythoncom.CoInitialize() before and CoUninitialize() in finally\n"
        "  COM apartment threading — can't access COM objects from Timer threads (capture hwnd as int first)\n"
        "  ListBox.Value — use .Column(0) explicitly, .Value may return wrong column\n"
        "  ComboBox filtering — never use Change event (infinite loops), use TextBox + LostFocus\n"
        "  dbAttachSavePWD = 131072 (NOT 65536) — use DoCmd.TransferDatabase, not DAO Attributes\n"
        "  Multiple JOINs — Access requires nested parentheses: FROM (A JOIN B ON ...) JOIN C ON ...\n\n"
        "VBA:\n"
        "  Str() adds leading space to positive numbers — use CStr() instead\n"
        "  Chr(128) truncates MsgBox text — use ChrW(8364) or \"EUR\" for euro symbol\n"
        "  ListBox AddItem — column separator is \";\", never use Format \"#,##0.00\" (comma breaks columns)\n"
        "  GetClipboardFilePath() can throw — always wrap in On Error Resume Next\n\n"
        "Startup:\n"
        "  SHIFT bypass is automatic — OpenCurrentDatabase always holds SHIFT to skip AutoExec/startup forms.\n"
        "  Any auto-opened forms are closed automatically after opening the database."
    ),
    "sql": (
        "Jet SQL DDL (access_execute_sql):\n"
        "  YESNO is not valid — use BIT, or better use access_create_table (accepts 'yesno')\n"
        "  DEFAULT not supported in CREATE TABLE — use access_set_field_property or access_create_table\n"
        "  AUTOINCREMENT works as a type (no IDENTITY needed)\n"
        "  Use SHORT instead of SMALLINT, LONG instead of INT\n"
        "  Prefer access_create_table over CREATE TABLE SQL for full type+default+description support\n\n"
        "ODBC pass-through:\n"
        "  QueryDef.Connect limit 255 chars — hardcode minimal connect string:\n"
        "  \"ODBC;DRIVER={ODBC Driver 17 for SQL Server};SERVER=myserver\\sqlexpress;"
        "DATABASE=mydb;UID=myuser;PWD=mypassword\""
    ),
    "vbe": (
        "VBE line numbers are 1-based.\n"
        "ProcCountLines can inflate the last proc's count past end of module — always clamp.\n"
        "Access must be Visible=True for VBE COM access to work.\n"
        "'Trust access to the VBA project object model' must be enabled in Trust Center.\n"
        "After design operations (set_control_props, create_control, delete_control),\n"
        "  close the form in Design view before accessing VBE CodeModule.\n"
        "access_vbe_append: was encoding & as &amp; (fixed in v0.7.3 with html.unescape).\n\n"
        "Reading VBA code:\n"
        "  For specific procedures: use access_vbe_get_proc (fast, precise).\n"
        "  For procedure index: use access_vbe_module_info first.\n"
        "  access_get_code exports the ENTIRE form (controls + VBA, can be 90KB+) — avoid for VBA investigation.\n"
        "  Recommended flow: access_vbe_module_info → access_vbe_get_proc for each relevant proc."
    ),
    "compile": (
        "access_compile_vba tips:\n"
        "  Use timeout param — RunCommand(126) shows MsgBox on error, blocks without it.\n"
        "  With timeout: polls every 2s, auto-clicks End/OK button + reports module/line/code via VBE.ActiveCodePane.\n"
        "  Also captures screenshot of the error MsgBox (dialog_screenshot in response).\n"
        "  Before compiling, check: Eval('Application.BrokenReference') — broken refs cause mysterious failures.\n"
        "  After compile error: use access_vbe_get_lines to read the problematic code, fix with access_vbe_replace_lines."
    ),
    "design": (
        "Design view + VBE conflict:\n"
        "  After design operations, the form may remain open in Design view.\n"
        "  access_vbe_replace_proc closes the form (acSaveYes) before accessing VBE.\n"
        "  All design operations invalidate all 3 caches (_vbe_code_cache, _parsed_controls_cache, _cm_cache).\n\n"
        "SaveAsText encoding:\n"
        "  Modules (.bas) — cp1252 (ANSI, no BOM)\n"
        "  Forms/reports/queries/macros — utf-16 (UTF-16LE with BOM)\n"
        "  access_set_code handles this automatically."
    ),
    "subform_tabcontrol": (
        "SubForm inside TabControl Page — BROKEN LAYOUT workaround:\n"
        "  Access recalculates TabControl positions when a SubForm exists inside a Page,\n"
        "  even with no SourceObject, both via CreateControl and LoadFromText.\n"
        "  Opening in Design view also triggers recalculation.\n\n"
        "SOLUTION: SubForm as SIBLING of TabControl (not child of Page).\n"
        "  1. access_export_text the form\n"
        "  2. In the UTF-16 text, add an EMPTY Page inside the TabControl\n"
        "  3. Add SubForm OUTSIDE the TabControl (same indent level, sibling in Detail section)\n"
        "     Set Visible = NotDefault (hidden by default)\n"
        "     Position it at the same coords as the Page content area (Left=75, Top=465)\n"
        "  4. Add OnChange =\"[Event Procedure]\" to the TabControl (Begin Tab)\n"
        "  5. Add OnOpen =\"[Event Procedure]\" to the form properties\n"
        "  6. Add CodeBehindForm section at end with VBA:\n"
        "     Private Sub Form_Open(Cancel As Integer)\n"
        "         sfName.Visible = False\n"
        "     End Sub\n"
        "     Private Sub tabName_Change()\n"
        "         If tabName.Pages(tabName.Value).Name = \"pagX\" Then\n"
        "             If sfName.SourceObject = \"\" Then\n"
        "                 sfName.SourceObject = \"Form.subf_xxx\"\n"
        "             End If\n"
        "             sfName.Visible = True\n"
        "         Else\n"
        "             sfName.Visible = False\n"
        "         End If\n"
        "     End Sub\n"
        "  7. access_import_text — single operation, NEVER open Design view after\n\n"
        "CRITICAL: Do NOT use set_form_property, set_control_props, or vbe_append\n"
        "  after import — they open Design view which recalculates positions.\n"
        "  All changes must go in the text file BEFORE import."
    ),
}


def ac_tips(topic: str = "") -> dict:
    """Return tips and gotchas for working with Access via MCP."""
    if not topic:
        return {
            "available_topics": list(_TIPS.keys()),
            "usage": "Call access_tips with a topic to get relevant tips.",
        }
    key = topic.lower().strip()
    if key in _TIPS:
        return {"topic": key, "tips": _TIPS[key]}
    # Fuzzy match — return any topic containing the search term
    matches = {k: v for k, v in _TIPS.items() if key in k or key in v.lower()}
    if matches:
        return {"matched_topics": {k: v for k, v in matches.items()}}
    return {
        "error": f"Topic '{topic}' not found.",
        "available_topics": list(_TIPS.keys()),
    }
