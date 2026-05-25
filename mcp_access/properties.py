"""
Database, field, form properties and startup options.
"""

from datetime import date, datetime
from typing import Any, Optional

from .core import _Session, log
from .constants import STARTUP_PROPS
from .helpers import coerce_prop, serialize_value


# DAO property type constants.  bool must be tested BEFORE int because
# isinstance(True, int) is True in Python.
_DB_BOOLEAN = 1
_DB_LONG = 4
_DB_SINGLE = 6
_DB_DOUBLE = 7
_DB_DATE = 8
_DB_TEXT = 10
_DB_MEMO = 12


def _infer_db_type(value: Any) -> int:
    """Pick a DAO property type code matching the Python value.
    Used when CreateProperty is asked to create a property that doesn't
    yet exist on the DB / field — wrong type was the cause of float and
    datetime properties being stored as text."""
    if isinstance(value, bool):
        return _DB_BOOLEAN
    if isinstance(value, int):
        return _DB_LONG
    if isinstance(value, float):
        return _DB_DOUBLE
    if isinstance(value, (datetime, date)):
        return _DB_DATE
    if isinstance(value, str) and len(value) > 255:
        return _DB_MEMO
    return _DB_TEXT


def ac_get_db_property(db_path: str, name: str) -> dict:
    """Reads a database property or an Access application option."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        val = db.Properties(name).Value
        return {"name": name, "value": val, "source": "database"}
    except Exception:
        pass
    try:
        val = app.GetOption(name)
        return {"name": name, "value": val, "source": "application"}
    except Exception as exc:
        raise ValueError(
            f"Property '{name}' not found in CurrentDb().Properties "
            f"or Application.GetOption. Error: {exc}"
        )


def ac_set_db_property(
    db_path: str, name: str, value: Any,
    prop_type: Optional[int] = None,
) -> dict:
    """Sets a database property or an Access application option."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    coerced = coerce_prop(value)

    # Try DB-level property
    try:
        db.Properties(name).Value = coerced
        return {"name": name, "value": coerced, "source": "database", "action": "updated"}
    except Exception:
        pass

    # Try Application option
    try:
        app.SetOption(name, coerced)
        return {"name": name, "value": coerced, "source": "application", "action": "updated"}
    except Exception:
        pass

    # Property doesn't exist -- create it
    if prop_type is None:
        prop_type = _infer_db_type(coerced)
    try:
        prop = db.CreateProperty(name, prop_type, coerced)
        db.Properties.Append(prop)
        return {"name": name, "value": coerced, "source": "database", "action": "created"}
    except Exception as exc:
        raise RuntimeError(
            f"Could not create property '{name}'. "
            f"prop_type: 1=Boolean, 4=Long, 6=Single, 7=Double, 8=Date, "
            f"10=Text, 12=Memo. Error: {exc}"
        )


def ac_get_field_properties(db_path: str, table_name: str, field_name: str) -> dict:
    """Reads all properties of a field."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Table '{table_name}' not found: {exc}")
    try:
        fld = td.Fields(field_name)
    except Exception as exc:
        raise ValueError(f"Field '{field_name}' not found in '{table_name}': {exc}")

    props = {}
    for i in range(fld.Properties.Count):
        try:
            p = fld.Properties(i)
            val = p.Value
            # Skip binary/complex values
            if isinstance(val, (str, int, float, bool)) or val is None:
                props[p.Name] = val
        except Exception:
            pass  # Some properties throw COM errors when read
    return {"table_name": table_name, "field_name": field_name, "properties": props}


def ac_set_field_property(
    db_path: str, table_name: str, field_name: str,
    property_name: str, value: Any,
) -> dict:
    """Sets a field property. Creates the property if it doesn't exist."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Table '{table_name}' not found: {exc}")
    try:
        fld = td.Fields(field_name)
    except Exception as exc:
        raise ValueError(f"Field '{field_name}' not found in '{table_name}': {exc}")

    coerced = coerce_prop(value)

    # Try to set existing property
    try:
        fld.Properties(property_name).Value = coerced
        return {
            "table_name": table_name, "field_name": field_name,
            "property_name": property_name, "value": coerced, "action": "updated",
        }
    except Exception:
        pass

    # Create property
    prop_type = _infer_db_type(coerced)
    try:
        prop = fld.CreateProperty(property_name, prop_type, coerced)
        fld.Properties.Append(prop)
        return {
            "table_name": table_name, "field_name": field_name,
            "property_name": property_name, "value": coerced, "action": "created",
        }
    except Exception as exc:
        raise RuntimeError(
            f"Could not set '{property_name}' on {table_name}.{field_name}: {exc}"
        )


def ac_list_startup_options(db_path: str) -> dict:
    """Lists common startup options with their current values."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    options = []
    for name in STARTUP_PROPS:
        val = None
        source = "<not set>"
        try:
            val = db.Properties(name).Value
            source = "database"
        except Exception:
            try:
                val = app.GetOption(name)
                source = "application"
            except Exception:
                pass
        options.append({"name": name, "value": val, "source": source})
    return {"count": len(options), "options": options}
