"""
SQL execution tools: ac_execute_sql, ac_execute_batch, ac_manage_query.
"""

from typing import Any, Optional

from .core import _Session, log
from .constants import DESTRUCTIVE_PREFIXES, DB_SEE_CHANGES, QUERYDEF_TYPE
from .helpers import serialize_value


def ac_execute_sql(
    db_path: str, sql: str, limit: int = 500,
    confirm_destructive: bool = False,
) -> dict:
    """
    Executes SQL in the database via DAO.
    SELECT  -> returns {rows: [...], count: N, truncated?: bool}
    Others  -> returns {affected_rows: N}
    DELETE/DROP/TRUNCATE/ALTER require confirm_destructive=True.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    normalized = sql.strip().upper()

    if normalized.startswith("SELECT"):
        limit = max(1, min(limit, 10000))
        try:
            rs = db.OpenRecordset(sql)
        except Exception as first_err:
            # Retry with dbSeeChanges for ODBC linked tables with IDENTITY columns
            try:
                rs = db.OpenRecordset(sql, 2, DB_SEE_CHANGES)  # 2 = dbOpenDynaset
            except Exception:
                raise RuntimeError(str(first_err)) from first_err
        fields = [rs.Fields(i).Name for i in range(rs.Fields.Count)]
        rows: list[dict] = []
        if not rs.EOF:
            rs.MoveFirst()
            while not rs.EOF and len(rows) < limit:
                rows.append(
                    {fields[i]: serialize_value(rs.Fields(i).Value)
                     for i in range(len(fields))}
                )
                rs.MoveNext()
        truncated = not rs.EOF
        rs.Close()
        result: dict = {"rows": rows, "count": len(rows)}
        if truncated:
            result["truncated"] = True
        return result
    else:
        if any(normalized.startswith(p) for p in DESTRUCTIVE_PREFIXES):
            if not confirm_destructive:
                return {
                    "error": (
                        "Destructive SQL detected. "
                        "Use confirm_destructive=true to execute: "
                        + sql[:100]
                    )
                }
        try:
            db.Execute(sql)
        except Exception as first_err:
            # Retry with dbSeeChanges for ODBC linked tables with IDENTITY columns
            try:
                db.Execute(sql, DB_SEE_CHANGES)
            except Exception:
                raise RuntimeError(str(first_err)) from first_err
        return {"affected_rows": db.RecordsAffected}


def ac_execute_batch(
    db_path: str, statements: list[dict], stop_on_error: bool = True,
    confirm_destructive: bool = False,
) -> dict:
    """
    Executes multiple SQL statements in a single call.
    statements: [{sql: str, label?: str}, ...]
    SELECT returns rows (limit 100 per statement).
    INSERT/UPDATE/DELETE returns affected_rows.
    stop_on_error=True stops at first error; False continues and reports all.
    confirm_destructive applies to entire batch.
    """
    if not statements:
        return {"error": "No SQL statements provided."}

    app = _Session.connect(db_path)
    db = app.CurrentDb()

    # Pre-scan: check destructive
    if not confirm_destructive:
        for i, stmt in enumerate(statements):
            sql_upper = stmt["sql"].strip().upper()
            if any(sql_upper.startswith(p) for p in DESTRUCTIVE_PREFIXES):
                label = stmt.get("label", f"statement[{i}]")
                return {
                    "error": (
                        f"Destructive SQL in '{label}'. "
                        "Use confirm_destructive=true to execute."
                    )
                }

    results: list[dict] = []
    succeeded = 0
    failed = 0

    for i, stmt in enumerate(statements):
        sql = stmt["sql"].strip()
        label = stmt.get("label")
        entry: dict = {"index": i}
        if label:
            entry["label"] = label

        try:
            sql_upper = sql.upper()
            if sql_upper.startswith("SELECT"):
                try:
                    rs = db.OpenRecordset(sql)
                except Exception as first_err:
                    try:
                        rs = db.OpenRecordset(sql, 2, DB_SEE_CHANGES)
                    except Exception:
                        raise RuntimeError(str(first_err)) from first_err
                fields = [rs.Fields(j).Name for j in range(rs.Fields.Count)]
                rows: list[dict] = []
                select_limit = 100
                if not rs.EOF:
                    rs.MoveFirst()
                    while not rs.EOF and len(rows) < select_limit:
                        rows.append(
                            {f: serialize_value(rs.Fields(f).Value) for f in fields}
                        )
                        rs.MoveNext()
                truncated = not rs.EOF
                rs.Close()
                entry["status"] = "ok"
                entry["rows"] = rows
                entry["count"] = len(rows)
                if truncated:
                    entry["truncated"] = True
            else:
                try:
                    db.Execute(sql)
                except Exception as first_err:
                    try:
                        db.Execute(sql, DB_SEE_CHANGES)
                    except Exception:
                        raise RuntimeError(str(first_err)) from first_err
                entry["status"] = "ok"
                entry["affected_rows"] = db.RecordsAffected
            succeeded += 1

        except Exception as exc:
            entry["status"] = "error"
            entry["error"] = str(exc)
            failed += 1
            if stop_on_error:
                results.append(entry)
                return {
                    "total": len(statements),
                    "succeeded": succeeded,
                    "failed": failed,
                    "stopped_at": i,
                    "results": results,
                }

        results.append(entry)

    return {
        "total": len(statements),
        "succeeded": succeeded,
        "failed": failed,
        "results": results,
    }


def ac_manage_query(
    db_path: str, action: str, query_name: str,
    sql: Optional[str] = None, new_name: Optional[str] = None,
    confirm: bool = False,
) -> dict:
    """Creates, modifies, renames, deletes or reads a QueryDef."""
    app = _Session.connect(db_path)
    db = app.CurrentDb()

    if action == "create":
        if not sql:
            raise ValueError("action='create' requires 'sql'")
        qd = db.CreateQueryDef(query_name, sql)
        return {"action": "created", "query_name": query_name, "sql": sql}

    elif action == "modify":
        if not sql:
            raise ValueError("action='modify' requires 'sql'")
        try:
            qd = db.QueryDefs(query_name)
        except Exception as exc:
            raise ValueError(f"Query '{query_name}' not found: {exc}")
        qd.SQL = sql
        return {"action": "modified", "query_name": query_name, "sql": sql}

    elif action == "delete":
        if not confirm:
            return {"error": f"Deleting query '{query_name}' requires confirm=true"}
        try:
            db.QueryDefs(query_name)  # verify exists
        except Exception as exc:
            raise ValueError(f"Query '{query_name}' not found: {exc}")
        db.QueryDefs.Delete(query_name)
        return {"action": "deleted", "query_name": query_name}

    elif action == "rename":
        if not new_name:
            raise ValueError("action='rename' requires 'new_name'")
        try:
            qd = db.QueryDefs(query_name)
        except Exception as exc:
            raise ValueError(f"Query '{query_name}' not found: {exc}")
        qd.Name = new_name
        return {"action": "renamed", "old_name": query_name, "new_name": new_name}

    elif action == "get_sql":
        try:
            qd = db.QueryDefs(query_name)
        except Exception as exc:
            raise ValueError(f"Query '{query_name}' not found: {exc}")
        qd_type = QUERYDEF_TYPE.get(qd.Type, f"Unknown({qd.Type})")
        return {"query_name": query_name, "sql": qd.SQL, "type": qd_type}

    else:
        raise ValueError(f"action must be create/modify/delete/rename/get_sql, received: '{action}'")
