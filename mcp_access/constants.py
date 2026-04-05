"""
Constants used across multiple modules.
"""

# ---------------------------------------------------------------------------
# Binary section names in Access form/report exports
# ---------------------------------------------------------------------------
BINARY_SECTIONS: frozenset[str] = frozenset({
    "PrtMip", "PrtDevMode", "PrtDevModeW",
    "PrtDevNames", "PrtDevNamesW",
    "RecSrcDt", "GUID", "NameMap",
})

# ---------------------------------------------------------------------------
# VBE CodeModule prefixes
# ---------------------------------------------------------------------------
VBE_PREFIX: dict[str, str] = {
    "module": "",
    "form":   "Form_",
    "report": "Report_",
}

# ---------------------------------------------------------------------------
# Control types (SaveAsText type numbers -> names)
# ---------------------------------------------------------------------------
CTRL_TYPE: dict[int, str] = {
    100: "Label",
    101: "Rectangle",
    102: "Line",
    103: "Image",
    104: "CommandButton",
    105: "OptionButton",
    106: "CheckBox",
    107: "OptionGroup",
    108: "BoundObjectFrame",
    109: "TextBox",
    110: "ListBox",
    111: "ComboBox",
    112: "Subform",
    113: "ObjectFrame",
    114: "PageBreak",
    118: "Page",
    119: "CustomControl",  # ActiveX in SaveAsText
    122: "Attachment",
    124: "NavigationButton",
    125: "NavigationControl",
    126: "WebBrowser",
}

CONTAINER_TYPES = {"Page", "OptionGroup"}

# AcControlType enum values (used by CreateControl/CreateReportControl)
AC_CONTROL_TYPE_NAMES: dict[str, int] = {
    "customcontrol": 119,
    "webbrowser": 128,
    "navigationcontrol": 129,
    "navigationbutton": 130,
    "chart": 133,
    "edgebrowser": 134,
}

# ---------------------------------------------------------------------------
# Design view constants
# ---------------------------------------------------------------------------
AC_DESIGN   = 1   # acDesign / acViewDesign
AC_FORM     = 2   # acForm (para DoCmd.Close/Save)
AC_REPORT   = 3   # acReport (para DoCmd.Close/Save)
AC_SAVE_YES = 1   # acSaveYes
AC_SAVE_NO  = 2   # acSaveNo

# Reverse map: control name -> type number
CTRL_TYPE_BY_NAME: dict[str, int] = {v.lower(): k for k, v in CTRL_TYPE.items()}
CTRL_TYPE_BY_NAME.update(AC_CONTROL_TYPE_NAMES)

# Section map (name -> enum value)
SECTION_MAP: dict[str, int] = {
    "detail": 0,
    "header": 1, "formheader": 1, "reportheader": 1,
    "footer": 2, "formfooter": 2, "reportfooter": 2,
    "pageheader": 3,
    "pagefooter": 4,
    "grouplevel1header": 5, "group1header": 5,
    "grouplevel1footer": 6, "group1footer": 6,
    "grouplevel2header": 7, "group2header": 7,
    "grouplevel2footer": 8, "group2footer": 8,
}

# ---------------------------------------------------------------------------
# DAO field types and constants
# ---------------------------------------------------------------------------
FIELD_TYPE_MAP: dict[str, int] = {
    "autonumber": 4, "autoincrement": 4,
    "long": 4, "integer": 3, "short": 3, "byte": 2,
    "text": 10, "memo": 12, "currency": 5,
    "double": 7, "single": 6, "float": 7,
    "datetime": 8, "date": 8,
    "boolean": 1, "yesno": 1, "bit": 1,
    "guid": 15, "ole": 11, "bigint": 16,
}

DB_AUTO_INCR_FIELD = 16       # dbAutoIncrField attribute flag
DB_ATTACH_SAVE_PWD = 131072   # dbAttachSavePWD (0x20000)
DB_SEE_CHANGES = 512          # dbSeeChanges

DAO_FIELD_TYPE: dict[int, str] = {
    1: "Boolean", 2: "Byte", 3: "Integer", 4: "Long", 5: "Currency",
    6: "Single", 7: "Double", 8: "Date/Time", 10: "Text",
    11: "OLE Object", 12: "Memo", 15: "GUID", 16: "BigInt",
    20: "Decimal",
}

# ---------------------------------------------------------------------------
# SQL safety
# ---------------------------------------------------------------------------
DESTRUCTIVE_PREFIXES = ("DELETE", "DROP", "TRUNCATE", "ALTER")

# ---------------------------------------------------------------------------
# Relationship attributes
# ---------------------------------------------------------------------------
REL_ATTR: dict[int, str] = {
    1: "Unique", 2: "DontEnforce", 256: "UpdateCascade", 4096: "DeleteCascade",
}

# ---------------------------------------------------------------------------
# Access output/transfer constants
# ---------------------------------------------------------------------------
AC_OUTPUT_REPORT = 3       # acOutputReport
AC_IMPORT = 0              # acImport
AC_EXPORT = 1              # acExport
AC_EXPORT_DELIM = 2        # acExportDelim (CSV export)
AC_SPREADSHEET_XLSX = 10   # acSpreadsheetTypeExcel12Xml
AC_CMD_COMPILE = 126       # acCmdCompileAndSaveAllModules

# ---------------------------------------------------------------------------
# QueryDef type constants
# ---------------------------------------------------------------------------
QUERYDEF_TYPE: dict[int, str] = {
    0: "Select", 16: "Crosstab", 32: "Delete", 48: "Update",
    64: "Append", 80: "MakeTable", 96: "DDL", 112: "SQLPassThrough",
    128: "Union", 240: "Action",
}

# ---------------------------------------------------------------------------
# Startup properties
# ---------------------------------------------------------------------------
STARTUP_PROPS = [
    "AppTitle", "AppIcon", "StartupForm", "StartupShowDBWindow",
    "StartupShowStatusBar", "StartupShortcutMenuBar",
    "AllowShortcutMenus", "AllowFullMenus", "AllowBuiltInToolbars",
    "AllowToolbarChanges", "AllowBreakIntoCode", "AllowSpecialKeys",
    "AllowBypassKey", "AllowDatasheetSchema",
]

# ---------------------------------------------------------------------------
# Output formats
# ---------------------------------------------------------------------------
OUTPUT_FORMATS: dict[str, str] = {
    "pdf": "PDF Format (*.pdf)",
    "xlsx": "Microsoft Excel (*.xlsx)",
    "rtf": "Rich Text Format (*.rtf)",
    "txt": "MS-DOS Text (*.txt)",
}

# ---------------------------------------------------------------------------
# Control property search (for find_usages)
# ---------------------------------------------------------------------------
CONTROL_SEARCH_PROPS = frozenset({
    "ControlSource", "RecordSource", "RowSource", "DefaultValue", "ValidationRule",
    "SourceObject", "LinkChildFields", "LinkMasterFields",
})
