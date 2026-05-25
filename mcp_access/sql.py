"""
SQL execution tools: ac_execute_sql, ac_execute_batch, ac_manage_query.
"""

import re
from typing import Any, Optional

from .core import _Session, log
from .constants import DESTRUCTIVE_PREFIXES, DB_SEE_CHANGES, QUERYDEF_TYPE
from .helpers import serialize_value


def _sql_effective_prefix(sql: str) -> str:
    """Return the leading keyword of a SQL statement after stripping leading
    whitespace and any leading `--` / `/* ... */` comments.

    Used by the destructive-statement guard so that
    ``-- note\\nDELETE FROM t`` cannot sneak past a ``startswith("DELETE")``
    check that only saw ``.strip().upper()`` applied.
    """
    s = sql
    changed = True
    while changed:
        changed = False
        stripped = s.lstrip()
        if stripped != s:
            s = stripped
            changed = True
        # Leading "-- ..." single-line comment
        if s.startswith("--"):
            newline = s.find("\n")
            if newline == -1:
                return ""
            s = s[newline + 1:]
            changed = True
            continue
        # Leading "/* ... */" block comment
        if s.startswith("/*"):
            end = s.find("*/")
            if end == -1:
                return ""
            s = s[end + 2:]
            changed = True
            continue
    return s.upper()


# Pattern for SELECT ... INTO (make-table query). Treated as destructive
# because it can overwrite an existing table without warning.
_SELECT_INTO_RE = re.compile(r'^\s*SELECT\b.*?\bINTO\b', re.IGNORECASE | re.DOTALL)


def _is_destructive(sql: str) -> bool:
    """Return True if *sql* is destructive (DELETE/DROP/TRUNCATE/ALTER)
    or a SELECT ... INTO make-table that could overwrite an existing table."""
    prefix = _sql_effective_prefix(sql)
    if any(prefix.startswith(p) for p in DESTRUCTIVE_PREFIXES):
        return True
    # SELECT ... INTO target FROM source — make-table query.
    return bool(_SELECT_INTO_RE.match(prefix))


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
    # Use an effective prefix that ignores leading comments — prevents a
    # "--\nDELETE FROM t" attempt from slipping past the destructive guard.
    normalized = _sql_effective_prefix(sql)

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
        if _is_destructive(sql) and not confirm_destructive:
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

    # Pre-scan: check destructive (ignore leading SQL comments)
    if not confirm_destructive:
        for i, stmt in enumerate(statements):
            if _is_destructive(stmt["sql"]):
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
            sql_upper = _sql_effective_prefix(sql)
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


# ---------------------------------------------------------------------------
# ac_search_data — search a text string across Text/Memo fields of all tables
# ---------------------------------------------------------------------------

# DAO field types considered text-searchable.  10=Text (dbText), 12=Memo (dbMemo).
# We intentionally exclude Hyperlink (also stored as Memo type 12 but exposed as
# Hyperlink via attributes) because LIKE on those is unreliable; tested 2026-05-25.
_SEARCHABLE_TEXT_TYPES = (10, 12)


def _escape_table_name(name: str) -> str:
    """Mirror of the bracket-escaping used by database.ac_table_info — a `]`
    inside a bracket-quoted identifier ends the identifier early, so double it.
    """
    return name.replace("]", "]]")


def _excerpt_around(text: str, needle: str, ctx: int = 40) -> str:
    """Return a small excerpt around the first case-insensitive match of *needle*."""
    if not text or not needle:
        return text or ""
    idx = text.lower().find(needle.lower())
    if idx < 0:
        return text[: ctx * 2]
    start = max(0, idx - ctx)
    end = min(len(text), idx + len(needle) + ctx)
    prefix = "…" if start > 0 else ""
    suffix = "…" if end < len(text) else ""
    return f"{prefix}{text[start:end]}{suffix}"


def ac_search_data(
    db_path: str,
    search_text: str,
    tables: Optional[list[str]] = None,
    max_results_per_table: int = 50,
    max_results_total: int = 500,
    match_case: bool = False,
) -> dict:
    """Search a text string in any Text/Memo field of any local table.

    Returns matches grouped by table.  Skips system tables (MSys*, ~temp*) and
    linked tables (which could be slow / surprising — querying a remote SQL
    server with `LIKE` over every text column on every linked table is rarely
    what the caller wants).  Pass `tables=[...]` to restrict to specific local
    tables.

    Note: Jet `LIKE` is case-insensitive by default. When `match_case=True`
    we filter Python-side after the SQL match, since Jet has no portable
    way to flip collation at query time on a per-query basis.
    """
    if not search_text:
        raise ValueError("search_text must be a non-empty string")

    app = _Session.connect(db_path)
    db = app.CurrentDb()

    needle = str(search_text)
    pattern = f"*{needle}*"  # Jet/Access LIKE uses * as wildcard (not %)

    requested = None
    if tables:
        # Case-insensitive comparison; preserve original case for the response
        requested = {str(t).lower() for t in tables}

    results: list[dict] = []
    total_matches = 0
    truncated = False

    # Iterate TableDefs — same pattern as ac_table_info / ac_list_linked_tables
    tabledefs = db.TableDefs
    for i in range(tabledefs.Count):
        if total_matches >= max_results_total:
            truncated = True
            break

        td = tabledefs(i)
        tname = td.Name
        if tname.startswith("MSys") or tname.startswith("~"):
            continue
        # Skip linked tables — Connect is non-empty for them
        try:
            if td.Connect:
                continue
        except Exception:
            pass

        if requested is not None and tname.lower() not in requested:
            continue

        # Collect Text/Memo fields
        text_fields: list[str] = []
        for j in range(td.Fields.Count):
            try:
                fld = td.Fields(j)
                if int(fld.Type) in _SEARCHABLE_TEXT_TYPES:
                    text_fields.append(fld.Name)
            except Exception:
                continue
        if not text_fields:
            continue

        # Build the WHERE: one LIKE per text field, OR'd together.
        # Field names quoted with brackets and `]` escaped by doubling.
        where_parts = []
        for fn in text_fields:
            safe_fn = fn.replace("]", "]]")
            where_parts.append(f"[{safe_fn}] LIKE ?")
        where_clause = " OR ".join(where_parts)

        safe_table = _escape_table_name(tname)
        # TOP N+1 lets us mark `table_truncated` accurately
        per_table_cap = max(1, int(max_results_per_table))
        sql = (
            f"SELECT TOP {per_table_cap + 1} * FROM [{safe_table}] "
            f"WHERE {where_clause}"
        )

        # DAO with positional parameters: create a QueryDef and bind via Parameters
        qd = None
        rs = None
        try:
            qd = db.CreateQueryDef("", sql)
            # Bind one parameter per OR-ed LIKE; all share the same pattern.
            for p_idx in range(len(text_fields)):
                qd.Parameters(p_idx).Value = pattern
            rs = qd.OpenRecordset(2)  # dbOpenDynaset = 2
        except Exception:
            # Fallback for drivers that misbehave with positional `?` in DAO:
            # inline the search pattern with quote-doubling (safe — pattern is
            # bounded by `*needle*` and we double single-quotes).
            safe_needle = needle.replace("'", "''")
            inline_pattern = f"'*{safe_needle}*'"
            where_inline = " OR ".join(
                f"[{fn.replace(']', ']]')}] LIKE {inline_pattern}" for fn in text_fields
            )
            sql_inline = (
                f"SELECT TOP {per_table_cap + 1} * FROM [{safe_table}] "
                f"WHERE {where_inline}"
            )
            try:
                rs = db.OpenRecordset(sql_inline)
            except Exception as exc:
                log.warning(
                    "ac_search_data: skip table '%s' — query failed (%s)",
                    tname, exc,
                )
                continue
        finally:
            # QueryDef without a name is implicit/temporary; nothing else to clean
            qd = None

        try:
            field_names = [rs.Fields(k).Name for k in range(rs.Fields.Count)]
        except Exception:
            try:
                rs.Close()
            except Exception:
                pass
            continue

        # Walk the recordset
        rows_for_table: list[dict] = []
        fields_matched: set = set()
        table_truncated = False
        try:
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    if len(rows_for_table) >= per_table_cap:
                        table_truncated = True
                        break
                    if total_matches >= max_results_total:
                        truncated = True
                        break

                    # Read this row into a plain dict
                    row: dict = {}
                    matched_in_this_row: list[str] = []
                    for k, fn in enumerate(field_names):
                        try:
                            v = rs.Fields(k).Value
                        except Exception:
                            v = None
                        row[fn] = serialize_value(v)
                        # Record which Text/Memo fields actually contain the needle
                        if fn in text_fields and isinstance(row[fn], str):
                            haystack = row[fn]
                            hit = (
                                needle in haystack
                                if match_case
                                else needle.lower() in haystack.lower()
                            )
                            if hit:
                                matched_in_this_row.append(fn)
                                fields_matched.add(fn)

                    if match_case and not matched_in_this_row:
                        # Jet LIKE matched case-insensitively but the caller
                        # asked for case-sensitive — drop this row.
                        rs.MoveNext()
                        continue

                    # Excerpt around the first matched field for readability
                    if matched_in_this_row:
                        first_match_field = matched_in_this_row[0]
                        excerpt_text = row.get(first_match_field, "") or ""
                        row["_excerpt"] = _excerpt_around(excerpt_text, needle)
                        row["_matched_fields"] = matched_in_this_row

                    rows_for_table.append(row)
                    total_matches += 1
                    rs.MoveNext()
        finally:
            try:
                rs.Close()
            except Exception:
                pass

        if rows_for_table:
            results.append({
                "table": tname,
                "match_count": len(rows_for_table),
                "fields_matched": sorted(fields_matched),
                "rows": rows_for_table,
                "truncated": table_truncated,
            })

    return {
        "search_text": needle,
        "match_case": bool(match_case),
        "total_matches": total_matches,
        "tables_with_hits": len(results),
        "truncated": truncated,
        "results": results,
    }
