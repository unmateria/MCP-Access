"""
Database and table operations: create database, create/alter/info tables.
"""

import os
from pathlib import Path
from typing import Any, Optional

from .core import (
    AC_TYPE, _Session, _vbe_code_cache, _parsed_controls_cache, log,
    invalidate_all_caches,
)
from .constants import FIELD_TYPE_MAP, DB_AUTO_INCR_FIELD, DAO_FIELD_TYPE, DB_SEE_CHANGES


# ---------------------------------------------------------------------------
# Field property helper
# ---------------------------------------------------------------------------

def _set_field_prop(db: Any, table_name: str, field_name: str,
                    prop_name: str, value: Any) -> None:
    """Internal helper to set field property with fallback to CreateProperty."""
    fld = db.TableDefs(table_name).Fields(field_name)
    try:
        fld.Properties(prop_name).Value = value
    except Exception:
        prop = fld.CreateProperty(prop_name, 10, value)  # 10 = dbText
        fld.Properties.Append(prop)


# ---------------------------------------------------------------------------
# Create database
# ---------------------------------------------------------------------------

def ac_create_database(db_path: str) -> dict:
    """Creates an empty Access database (.accdb). Error if it already exists."""
    resolved = str(Path(db_path).resolve())
    if os.path.exists(resolved):
        raise FileExistsError(
            f"'{resolved}' already exists. Use access_execute_sql to modify it."
        )
    # Ensure Access is running
    if _Session._app is None:
        _Session._launch()
    app = _Session._app
    # Close any previously open DB
    if _Session._db_open is not None:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        _Session._db_open = None
    try:
        app.NewCurrentDatabase(resolved)
    except Exception as exc:
        raise RuntimeError(f"Error creating database: {exc}")
    # FIX: Close and reopen to ensure CurrentDb() works reliably
    try:
        app.CloseCurrentDatabase()
        app.OpenCurrentDatabase(resolved)
    except Exception:
        pass  # If reopen fails, at least the file was created
    _Session._db_open = resolved
    invalidate_all_caches()
    size = os.path.getsize(resolved) if os.path.exists(resolved) else 0
    return {"db_path": resolved, "status": "created", "size_bytes": size}


# ---------------------------------------------------------------------------
# Create table via DAO
# ---------------------------------------------------------------------------

def ac_create_table(db_path: str, table_name: str, fields: list[dict]) -> dict:
    """
    Creates an Access table via DAO with full support for types, defaults,
    descriptions and properties — all in a single call.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()

    # Verify it doesn't exist
    existing = [db.TableDefs(i).Name for i in range(db.TableDefs.Count)]
    if table_name in existing:
        raise ValueError(f"Table '{table_name}' already exists.")

    td = db.CreateTableDef(table_name)
    pk_fields: list[str] = []
    created_fields: list[dict] = []

    for fdef in fields:
        name = fdef["name"]
        ftype = fdef.get("type", "text").lower()
        size = fdef.get("size", 0)
        required = fdef.get("required", False)
        pk = fdef.get("primary_key", False)

        dao_type = FIELD_TYPE_MAP.get(ftype)
        if dao_type is None:
            raise ValueError(
                f"Unknown type: '{ftype}'. "
                f"Valid: {sorted(set(FIELD_TYPE_MAP.keys()))}"
            )

        is_auto = ftype in ("autonumber", "autoincrement")

        # Text needs size
        if dao_type == 10 and size == 0:
            size = 255

        if size > 0:
            fld = td.CreateField(name, dao_type, size)
        else:
            fld = td.CreateField(name, dao_type)

        if is_auto:
            fld.Attributes = fld.Attributes | DB_AUTO_INCR_FIELD

        fld.Required = required or pk

        td.Fields.Append(fld)

        if pk:
            pk_fields.append(name)

        created_fields.append({
            "name": name,
            "type": ftype,
            "size": size if size > 0 else None,
        })

    # Create primary key index
    if pk_fields:
        idx = td.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Unique = True
        for pk_name in pk_fields:
            idx_fld = idx.CreateField(pk_name)
            idx.Fields.Append(idx_fld)
        td.Indexes.Append(idx)

    db.TableDefs.Append(td)
    db.TableDefs.Refresh()

    # Set defaults and descriptions via field properties (post-creation)
    for fdef in fields:
        name = fdef["name"]
        default = fdef.get("default")
        description = fdef.get("description")
        if default is not None:
            try:
                _set_field_prop(db, table_name, name, "DefaultValue", str(default))
            except Exception as e:
                log.warning("Error setting default for %s.%s: %s", table_name, name, e)
        if description is not None:
            try:
                _set_field_prop(db, table_name, name, "Description", description)
            except Exception as e:
                log.warning("Error setting description for %s.%s: %s", table_name, name, e)

    return {
        "table_name": table_name,
        "fields": created_fields,
        "primary_key": pk_fields,
        "status": "created",
    }


# ---------------------------------------------------------------------------
# Alter table via DAO
# ---------------------------------------------------------------------------

def ac_alter_table(
    db_path: str, table_name: str, action: str,
    field_name: str, new_name: str | None = None,
    field_type: str = "text", size: int = 0,
    required: bool = False, default: Any = None,
    description: str | None = None, confirm: bool = False,
) -> dict:
    """
    Modifies the structure of an Access table via DAO.
    Actions: add_field, delete_field (requires confirm=true), rename_field.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    td = db.TableDefs(table_name)

    if action == "add_field":
        ftype = field_type.lower()
        dao_type = FIELD_TYPE_MAP.get(ftype)
        if dao_type is None:
            raise ValueError(
                f"Unknown type: '{ftype}'. "
                f"Valid: {sorted(set(FIELD_TYPE_MAP.keys()))}"
            )
        is_auto = ftype in ("autonumber", "autoincrement")
        if dao_type == 10 and size == 0:
            size = 255
        if size > 0:
            fld = td.CreateField(field_name, dao_type, size)
        else:
            fld = td.CreateField(field_name, dao_type)
        if is_auto:
            fld.Attributes = fld.Attributes | DB_AUTO_INCR_FIELD
        fld.Required = required
        td.Fields.Append(fld)
        td.Fields.Refresh()
        if default is not None:
            try:
                _set_field_prop(db, table_name, field_name, "DefaultValue", str(default))
            except Exception as e:
                log.warning("Error setting default for %s.%s: %s", table_name, field_name, e)
        if description is not None:
            try:
                _set_field_prop(db, table_name, field_name, "Description", description)
            except Exception as e:
                log.warning("Error setting description for %s.%s: %s", table_name, field_name, e)
        return {"action": "field_added", "table": table_name, "field": field_name, "type": ftype}

    elif action == "delete_field":
        if not confirm:
            return {
                "error": (
                    f"Deleting field '{field_name}' from '{table_name}' is destructive. "
                    "Use confirm=true to confirm."
                )
            }
        td.Fields.Delete(field_name)
        return {"action": "field_deleted", "table": table_name, "field": field_name}

    elif action == "rename_field":
        if not new_name:
            raise ValueError("rename_field requires new_name")
        fld = td.Fields(field_name)
        fld.Name = new_name
        return {"action": "field_renamed", "table": table_name,
                "old_name": field_name, "new_name": new_name}

    else:
        raise ValueError(
            f"Unknown action: '{action}'. "
            "Valid: add_field, delete_field, rename_field"
        )


# ---------------------------------------------------------------------------
# Table info
# ---------------------------------------------------------------------------

def ac_table_info(db_path: str, table_name: str) -> dict:
    """
    Returns the structure of a local or linked Access table:
    fields with name, type, size, required; record_count; is_linked.
    Uses DAO TableDef.Fields.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Table '{table_name}' not found: {exc}")

    is_linked = bool(td.Connect)
    fields: list[dict] = []
    for i in range(td.Fields.Count):
        fld = td.Fields(i)
        ftype = fld.Type
        # AutoNumber detection: Long (4) + dbAutoIncrField attribute (16)
        type_name = DAO_FIELD_TYPE.get(ftype, f"Type{ftype}")
        if ftype == 4 and (fld.Attributes & 16):
            type_name = "AutoNumber"
        fields.append({
            "name": fld.Name,
            "type": type_name,
            "size": fld.Size,
            "required": bool(fld.Required),
        })

    # Record count (may fail on linked tables)
    try:
        record_count = td.RecordCount
        if record_count == -1:
            # For linked tables, open recordset to count
            rs = db.OpenRecordset(f"SELECT COUNT(*) AS cnt FROM [{table_name}]")
            record_count = rs.Fields(0).Value
            rs.Close()
    except Exception:
        record_count = -1

    return {
        "table_name": table_name,
        "fields": fields,
        "record_count": record_count,
        "is_linked": is_linked,
        "source_table": td.SourceTableName if is_linked else "",
        "connect": td.Connect if is_linked else "",
    }
