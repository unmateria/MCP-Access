"""
Linked tables, relationships, VBA references, indexes.
"""

from typing import Any, Optional

from .core import _Session, _vbe_code_cache, log
from .constants import REL_ATTR, DB_ATTACH_SAVE_PWD


def ac_list_linked_tables(db_path: str) -> dict:
    """Lists all linked tables with connection information."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    linked: list[dict] = []
    for i in range(db.TableDefs.Count):
        td = db.TableDefs(i)
        conn = td.Connect
        if not conn:
            continue
        name = td.Name
        if name.startswith("~") or name.startswith("MSys"):
            continue
        linked.append({
            "name": name,
            "source_table": td.SourceTableName,
            "connect_string": conn,
            "is_odbc": conn.upper().startswith("ODBC;"),
        })
    return {"count": len(linked), "linked_tables": linked}


def ac_relink_table(
    db_path: str, table_name: str, new_connect: str,
    relink_all: bool = False,
) -> dict:
    """Changes the connection string of a linked table and refreshes."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    relinked: list[dict] = []

    try:
        ref_td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Table '{table_name}' not found: {exc}")
    if not ref_td.Connect:
        raise ValueError(f"'{table_name}' is not a linked table")

    # Auto-detect if connect string has UID/PWD -> set dbAttachSavePWD
    _needs_save_pwd = ("UID=" in new_connect.upper() or "PWD=" in new_connect.upper())

    def _relink_one(td_name: str, old_conn: str):
        """Relink a single table. If dbAttachSavePWD needed, use TransferDatabase."""
        if _needs_save_pwd:
            # DAO Attributes can't be set reliably from Python COM.
            # Use DoCmd.TransferDatabase with StoreLogin=True instead.
            src_table = db.TableDefs(td_name).SourceTableName
            old_connect_backup = db.TableDefs(td_name).Connect
            try:
                app.DoCmd.DeleteObject(0, td_name)  # acTable = 0
            except Exception:
                pass  # already gone
            # acLink=2, acTable=0
            try:
                app.DoCmd.TransferDatabase(
                    2, "ODBC Database", new_connect,
                    0, src_table, td_name, False, True,  # StoreLogin=True
                )
            except Exception as exc:
                # ROLLBACK: try to restore the old link
                try:
                    app.DoCmd.TransferDatabase(
                        2, "ODBC Database", old_connect_backup,
                        0, src_table, td_name, False, True,
                    )
                    log.warning("ac_relink_table: rollback ok for '%s'", td_name)
                except Exception:
                    log.error("ac_relink_table: rollback FAILED for '%s'", td_name)
                raise RuntimeError(
                    f"Error relinking '{td_name}': {exc}. "
                    "Attempted to restore the original link."
                )
        else:
            td = db.TableDefs(td_name)
            td.Connect = new_connect
            td.RefreshLink()
        relinked.append({"name": td_name, "old_connect": old_conn, "new_connect": new_connect})

    if relink_all:
        old_connect = ref_td.Connect
        names_to_relink = []
        for i in range(db.TableDefs.Count):
            td = db.TableDefs(i)
            if td.Connect == old_connect:
                names_to_relink.append((td.Name, td.Connect))
        for name, old in names_to_relink:
            _relink_one(name, old)
    else:
        old = ref_td.Connect
        _relink_one(table_name, old)

    return {"relinked_count": len(relinked), "tables": relinked}


def ac_list_relationships(db_path: str) -> dict:
    """Lists all relationships between tables in the database."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    rels: list[dict] = []
    for i in range(db.Relations.Count):
        rel = db.Relations(i)
        name = rel.Name
        if name.startswith("MSys"):
            continue
        fields: list[dict] = []
        for j in range(rel.Fields.Count):
            fld = rel.Fields(j)
            fields.append({"local": fld.Name, "foreign": fld.ForeignName})
        attrs = rel.Attributes
        attr_flags = [label for bit, label in REL_ATTR.items() if attrs & bit]
        rels.append({
            "name": name,
            "table": rel.Table,
            "foreign_table": rel.ForeignTable,
            "fields": fields,
            "attributes": attrs,
            "attribute_flags": attr_flags,
        })
    return {"count": len(rels), "relationships": rels}


def ac_create_relationship(
    db_path: str, name: str, table: str, foreign_table: str,
    fields: list[dict], attributes: int = 0,
) -> dict:
    """Creates a relationship between two tables via DAO."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    rel = db.CreateRelation(name, table, foreign_table, attributes)
    for fmap in fields:
        local_name = fmap.get("local")
        foreign_name = fmap.get("foreign")
        if not local_name or not foreign_name:
            raise ValueError(f"Each field must have 'local' and 'foreign'. Received: {fmap}")
        fld = rel.CreateField(local_name)
        fld.ForeignName = foreign_name
        rel.Fields.Append(fld)
    try:
        db.Relations.Append(rel)
    except Exception as exc:
        raise RuntimeError(
            f"Error creating relationship '{name}' between '{table}' and '{foreign_table}': {exc}"
        )
    attr_flags = [label for bit, label in REL_ATTR.items() if attributes & bit]
    return {
        "name": name, "table": table, "foreign_table": foreign_table,
        "fields": fields, "attributes": attributes,
        "attribute_flags": attr_flags, "status": "created",
    }


def ac_delete_relationship(db_path: str, name: str) -> dict:
    """Deletes a relationship between tables by name."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        db.Relations.Delete(name)
    except Exception as exc:
        raise RuntimeError(f"Error deleting relationship '{name}': {exc}")
    return {"action": "deleted", "name": name}


def ac_list_references(db_path: str) -> dict:
    """Lists all VBA references in the project."""
    app = _Session.connect(db_path)
    try:
        refs_col = app.VBE.ActiveVBProject.References
    except Exception as exc:
        raise RuntimeError(f"Could not access VBE. Error: {exc}")
    refs: list[dict] = []
    for i in range(1, refs_col.Count + 1):  # VBA collections are 1-based
        ref = refs_col(i)
        try:
            is_broken = bool(ref.IsBroken)
        except Exception:
            is_broken = True
        try:
            built_in = bool(ref.BuiltIn)
        except Exception:
            built_in = False
        refs.append({
            "name": ref.Name,
            "description": ref.Description,
            "full_path": ref.FullPath,
            "guid": ref.GUID if ref.GUID else "",
            "major": ref.Major,
            "minor": ref.Minor,
            "is_broken": is_broken,
            "built_in": built_in,
        })
    return {"count": len(refs), "references": refs}


def ac_manage_reference(
    db_path: str, action: str,
    name: Optional[str] = None,
    path: Optional[str] = None,
    guid: Optional[str] = None,
    major: int = 0, minor: int = 0,
) -> dict:
    """Adds or removes a VBA reference from the project."""
    app = _Session.connect(db_path)
    try:
        refs = app.VBE.ActiveVBProject.References
    except Exception as exc:
        raise RuntimeError(f"Could not access VBE. Error: {exc}")

    if action == "add":
        if guid:
            try:
                ref = refs.AddFromGuid(guid, major, minor)
                result = {"action": "added", "name": ref.Name, "guid": guid, "major": major, "minor": minor}
            except Exception as exc:
                raise RuntimeError(f"Error adding reference by GUID '{guid}': {exc}")
        elif path:
            try:
                ref = refs.AddFromFile(path)
                result = {"action": "added", "name": ref.Name, "full_path": path}
            except Exception as exc:
                raise RuntimeError(f"Error adding reference from '{path}': {exc}")
        else:
            raise ValueError("action='add' requires 'guid' or 'path'")
    elif action == "remove":
        if not name:
            raise ValueError("action='remove' requires 'name'")
        found = None
        for i in range(1, refs.Count + 1):
            ref = refs(i)
            if ref.Name.lower() == name.lower():
                found = ref
                break
        if found is None:
            raise ValueError(f"Reference '{name}' not found")
        try:
            if found.BuiltIn:
                raise ValueError(f"'{name}' is built-in and cannot be removed")
        except AttributeError:
            pass  # BuiltIn property not available in old Access versions
        try:
            refs.Remove(found)
            result = {"action": "removed", "name": name}
        except Exception as exc:
            raise RuntimeError(f"Error removing reference '{name}': {exc}")
    else:
        raise ValueError(f"action must be 'add' or 'remove', received: '{action}'")

    # References affect VBE compilation -- clear code caches
    _vbe_code_cache.clear()
    _Session._cm_cache.clear()
    return result


def ac_list_indexes(db_path: str, table_name: str) -> dict:
    """Lists the indexes of a table."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Table '{table_name}' not found: {exc}")

    indexes = []
    for i in range(td.Indexes.Count):
        idx = td.Indexes(i)
        fields = []
        for j in range(idx.Fields.Count):
            f = idx.Fields(j)
            fields.append({
                "name": f.Name,
                "order": "desc" if f.Attributes & 1 else "asc",
            })
        indexes.append({
            "name": idx.Name,
            "fields": fields,
            "primary": bool(idx.Primary),
            "unique": bool(idx.Unique),
            "foreign": bool(idx.Foreign),
        })
    return {"table_name": table_name, "count": len(indexes), "indexes": indexes}


def ac_manage_index(
    db_path: str, table_name: str, action: str, index_name: str,
    fields: Optional[list] = None,
    primary: bool = False, unique: bool = False,
) -> dict:
    """Creates or deletes an index on a table."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    try:
        td = db.TableDefs(table_name)
    except Exception as exc:
        raise ValueError(f"Table '{table_name}' not found: {exc}")

    if action == "create":
        if not fields:
            raise ValueError("action='create' requires 'fields' [{name, order?}]")
        idx = td.CreateIndex(index_name)
        idx.Primary = primary
        idx.Unique = unique
        for fdef in fields:
            fname = fdef if isinstance(fdef, str) else fdef["name"]
            fld = idx.CreateField(fname)
            if isinstance(fdef, dict) and fdef.get("order", "asc").lower() == "desc":
                fld.Attributes = 1  # dbDescending
            idx.Fields.Append(fld)
        td.Indexes.Append(idx)
        return {
            "action": "created", "table_name": table_name,
            "index_name": index_name, "fields": fields,
            "primary": primary, "unique": unique,
        }

    elif action == "delete":
        try:
            td.Indexes(index_name)  # verify exists
        except Exception as exc:
            raise ValueError(f"Index '{index_name}' not found in '{table_name}': {exc}")
        td.Indexes.Delete(index_name)
        return {"action": "deleted", "table_name": table_name, "index_name": index_name}

    else:
        raise ValueError(f"action must be create/delete, received: '{action}'")
