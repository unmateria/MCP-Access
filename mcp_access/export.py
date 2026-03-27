"""
Report output and data transfer (import/export Excel/CSV).
"""

import os
from pathlib import Path
from typing import Optional

from .core import _Session, log
from .constants import (
    OUTPUT_FORMATS,
    AC_OUTPUT_REPORT,
    AC_IMPORT,
    AC_EXPORT,
    AC_EXPORT_DELIM,
    AC_SPREADSHEET_XLSX,
)


# ---------------------------------------------------------------------------
# Output report (PDF, XLSX, RTF, TXT)
# ---------------------------------------------------------------------------

def ac_output_report(
    db_path: str, report_name: str,
    output_path: Optional[str] = None, fmt: str = "pdf",
) -> dict:
    """Exports a report to PDF, XLSX, RTF or TXT."""
    app = _Session.connect(db_path)
    fmt_lower = fmt.lower()
    format_string = OUTPUT_FORMATS.get(fmt_lower)
    if not format_string:
        raise ValueError(f"Format '{fmt}' not supported. Use: {list(OUTPUT_FORMATS.keys())}")

    ext_map = {"pdf": ".pdf", "xlsx": ".xlsx", "rtf": ".rtf", "txt": ".txt"}
    if output_path is None:
        resolved = str(Path(db_path).resolve())
        db_dir = os.path.dirname(resolved)
        output_path = os.path.join(db_dir, f"{report_name}{ext_map[fmt_lower]}")

    output_path = str(Path(output_path).resolve())
    try:
        app.DoCmd.OutputTo(AC_OUTPUT_REPORT, report_name, format_string, output_path)
    except Exception as exc:
        raise RuntimeError(f"Error exporting report '{report_name}': {exc}")

    size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
    return {
        "report_name": report_name, "output_path": output_path,
        "format": fmt_lower, "size_bytes": size,
    }


# ---------------------------------------------------------------------------
# Transfer data (import/export Excel/CSV)
# ---------------------------------------------------------------------------

def ac_transfer_data(
    db_path: str, action: str, file_path: str, table_name: str,
    has_headers: bool = True, file_type: str = "xlsx",
    range_: Optional[str] = None, spec_name: Optional[str] = None,
) -> dict:
    """Imports or exports data between Access and Excel/CSV."""
    app = _Session.connect(db_path)
    file_path = str(Path(file_path).resolve())
    ft = file_type.lower()

    if action == "import":
        transfer_type_spreadsheet = AC_IMPORT       # 0
        transfer_type_text = 0                       # acImportDelim
    elif action == "export":
        transfer_type_spreadsheet = AC_EXPORT        # 1
        transfer_type_text = AC_EXPORT_DELIM         # 2
    else:
        raise ValueError(f"action must be 'import' or 'export', received: '{action}'")

    try:
        if ft in ("xlsx", "xls", "excel"):
            app.DoCmd.TransferSpreadsheet(
                transfer_type_spreadsheet,
                AC_SPREADSHEET_XLSX,
                table_name,
                file_path,
                has_headers,
                range_ or "",
            )
        elif ft in ("csv", "txt", "text"):
            app.DoCmd.TransferText(
                transfer_type_text,
                spec_name or "",
                table_name,
                file_path,
                has_headers,
            )
        else:
            raise ValueError(f"file_type '{file_type}' not supported. Use: xlsx, csv")
    except ValueError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Error en TransferData ({action} {ft}): {exc}")

    return {"action": action, "file_type": ft, "table_name": table_name, "file_path": file_path}
