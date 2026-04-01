"""
VBE (Visual Basic Editor) line-level operations.

Extracted from the monolithic access_mcp_server.py — same logic,
only imports updated to point at the refactored package structure.
"""

import difflib
import html as html_mod
import os
import re
import tempfile
from typing import Any

from .core import (
    AC_TYPE, _Session, _vbe_code_cache, _parsed_controls_cache, log,
    invalidate_object_caches,
)
from .constants import (
    VBE_PREFIX, AC_FORM, AC_REPORT, AC_SAVE_YES,
    CONTROL_SEARCH_PROPS,
)
from .helpers import text_matches, read_tmp


# ---------------------------------------------------------------------------
# CodeModule helpers
# ---------------------------------------------------------------------------

def _get_code_module(app: Any, object_type: str, object_name: str) -> Any:
    """
    Returns the VBE CodeModule for the given component.
    Caches the COM object to avoid 3 chained calls per VBE tool.
    Requires 'Trust access to the VBA project object model'
    enabled in Access Trust Center settings.
    """
    if object_type not in VBE_PREFIX:
        raise ValueError(
            f"object_type '{object_type}' does not support VBE. Use 'module', 'form' or 'report'."
        )
    cache_key = f"{object_type}:{object_name}"
    cm = _Session._cm_cache.get(cache_key)
    if cm is not None:
        return cm
    component_name = VBE_PREFIX[object_type] + object_name
    try:
        project = app.VBE.VBProjects(1)
        component = project.VBComponents(component_name)
        cm = component.CodeModule
        _Session._cm_cache[cache_key] = cm
        return cm
    except Exception as exc:
        _Session._cm_cache.pop(cache_key, None)  # invalidate stale cache entry
        raise RuntimeError(
            f"Could not access CodeModule '{component_name}'. "
            f"Is 'Trust access to the VBA project object model' enabled "
            f"in Access Trust Center settings?\nError: {exc}"
        )


def _cm_all_code(cm: Any, cache_key: str) -> str:
    """
    Returns the full text of a CodeModule using _vbe_code_cache.
    In a session with multiple tools on the same module, the full COM read
    (cm.Lines) is done once; subsequent calls use the cache.
    """
    if cache_key not in _vbe_code_cache:
        total = cm.CountOfLines
        _vbe_code_cache[cache_key] = cm.Lines(1, total) if total > 0 else ""
    return _vbe_code_cache[cache_key]


# ---------------------------------------------------------------------------
# Structural helpers — Option protection, health check, ws-matching
# ---------------------------------------------------------------------------

_OPTION_RE = re.compile(r'^\s*Option\s+(Explicit|Compare\s+\w+)\s*$', re.IGNORECASE)


def _strip_option_lines(code: str) -> tuple[str, list[str]]:
    """
    Removes Option Explicit / Option Compare lines from code.
    Returns (cleaned_code, list[str] warnings).
    """
    warnings: list[str] = []
    out_lines: list[str] = []
    for line in code.splitlines(keepends=True):
        if _OPTION_RE.match(line.rstrip('\r\n')):
            warnings.append(f"Stripped misplaced Option line: {line.strip()!r}")
        else:
            out_lines.append(line)
    return "".join(out_lines), warnings


def _check_module_health(cm: Any, cache_key: str, expected_total: int = 0) -> list[str]:
    """
    Structural health check after a write operation.
    Returns list of WARNING strings (empty = OK).
    """
    warnings: list[str] = []
    # Force fresh read (cache was just invalidated)
    total = cm.CountOfLines
    if total == 0:
        return warnings
    all_code = cm.Lines(1, total)
    lines = all_code.splitlines()

    # Check 1 — Option placement: should be in first 5 lines
    for i, line in enumerate(lines):
        if _OPTION_RE.match(line.rstrip('\r\n')) and i >= 5:
            warnings.append(
                f"WARNING: '{line.strip()}' found at line {i + 1} (expected in first 5 lines)"
            )

    # Check 2 — Duplicate labels
    label_re = re.compile(r'^(\w+):\s*$')
    label_positions: dict[str, list[int]] = {}
    for i, line in enumerate(lines):
        stripped = line.strip()
        # Skip comments, Case statements, pure numbers
        if stripped.startswith("'") or stripped.startswith("Case "):
            continue
        m = label_re.match(stripped)
        if m:
            label = m.group(1)
            # Exclude numeric labels and common non-label patterns
            if label.isdigit():
                continue
            label_positions.setdefault(label, []).append(i + 1)
    for label, positions in label_positions.items():
        if len(positions) > 1:
            warnings.append(
                f"WARNING: Duplicate label '{label}:' at lines {positions}"
            )

    # Check 3 — Count sanity
    if expected_total > 0 and total != expected_total:
        warnings.append(
            f"WARNING: Expected {expected_total} lines after edit, but module has {total}"
        )

    return warnings


def _ws_normalized_match(proc_code: str, find_text: str) -> tuple[int, int] | None:
    """
    Whitespace-tolerant matching: strips leading whitespace from each line
    and does a sliding window search.
    Returns (start_idx, end_idx) 0-based line indices into proc_code lines, or None.
    """
    proc_lines = proc_code.splitlines()
    find_lines = find_text.splitlines()
    # Remove empty trailing lines from find_text
    while find_lines and not find_lines[-1].strip():
        find_lines.pop()
    if not find_lines:
        return None

    proc_stripped = [l.lstrip() for l in proc_lines]
    find_stripped = [l.lstrip() for l in find_lines]
    window = len(find_stripped)

    for i in range(len(proc_stripped) - window + 1):
        if proc_stripped[i : i + window] == find_stripped:
            return (i, i + window - 1)
    return None


def _closest_match_context(proc_code: str, find_text: str, proc_name: str) -> str:
    """
    When both exact and ws-normalized match fail, finds the most similar line
    using difflib and returns a contextual snippet for a descriptive error.
    """
    proc_lines = proc_code.splitlines()
    find_lines = [l.strip() for l in find_text.splitlines() if l.strip()]
    if not find_lines:
        return f"Empty find text in proc '{proc_name}'"

    # Use the first non-empty find line as the reference
    ref = find_lines[0]
    best_ratio = 0.0
    best_idx = 0
    sm = difflib.SequenceMatcher(None, ref, "")
    for i, line in enumerate(proc_lines):
        sm.set_seq2(line.strip())
        ratio = sm.ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_idx = i

    # Build context: 3 lines around best candidate
    start = max(0, best_idx - 1)
    end = min(len(proc_lines), best_idx + 2)
    context_lines = []
    for j in range(start, end):
        marker = ">>>" if j == best_idx else "   "
        context_lines.append(f"  {marker} L{j + 1}: {proc_lines[j].rstrip()}")

    return (
        f"Best match ({best_ratio:.0%} similar) near line {best_idx + 1} "
        f"of '{proc_name}':\n" + "\n".join(context_lines) +
        f"\n  Looking for: {ref[:80]!r}"
    )


# ---------------------------------------------------------------------------
# VBE get operations
# ---------------------------------------------------------------------------

def ac_vbe_get_lines(
    db_path: str, object_type: str, object_name: str,
    start_line: int, count: int = None, end_line: int = None
) -> str:
    """Reads a range of lines without exporting the entire module."""
    if end_line is not None and count is None:
        count = end_line - start_line + 1
    if count is None:
        raise ValueError("Either count or end_line must be provided")
    if count < 1:
        raise ValueError(f"count must be >= 1 (got {count})")
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    all_lines = all_code.splitlines()
    total = len(all_lines)
    if start_line < 1 or start_line > total:
        raise ValueError(f"start_line {start_line} out of range (1-{total})")
    actual = min(count, total - start_line + 1)
    return "\n".join(all_lines[start_line - 1 : start_line - 1 + actual])


def ac_vbe_get_proc(
    db_path: str, object_type: str, object_name: str, proc_name: str
) -> dict:
    """
    Returns information and code for a specific procedure.
    Much more efficient than ac_get_code when only one function is needed.
    Returns: start_line, body_line, count, code.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    try:
        start = cm.ProcStartLine(proc_name, 0)   # 0 = vbext_pk_Proc (COM call — fast)
        body  = cm.ProcBodyLine(proc_name, 0)
        count = cm.ProcCountLines(proc_name, 0)
    except Exception as exc:
        raise RuntimeError(
            f"Procedure '{proc_name}' not found in '{object_name}': {exc}"
        )
    # Extract text from cache instead of an extra cm.Lines call
    cache_key = f"{object_type}:{object_name}"
    all_lines = _cm_all_code(cm, cache_key).splitlines()
    code = "\n".join(all_lines[start - 1 : start - 1 + count])
    return {
        "proc_name":  proc_name,
        "start_line": start,
        "body_line":  body,
        "count":      count,
        "code":       code,
    }


def ac_vbe_module_info(
    db_path: str, object_type: str, object_name: str
) -> dict:
    """
    Returns the total lines and the list of procedures with their positions.
    Useful as a quick index before editing, without downloading the full code.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    all_lines = all_code.splitlines()
    total = len(all_lines)
    procs: list[dict] = []
    if total > 0:
        seen: set[str] = set()
        for i, raw_line in enumerate(all_lines, start=1):
            m = re.match(
                r'^(?:Public\s+|Private\s+|Friend\s+)?'
                r'(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)',
                raw_line.strip(), re.IGNORECASE,
            )
            if m:
                pname = m.group(1)
                if pname in seen:
                    continue
                seen.add(pname)
                try:
                    pstart = cm.ProcStartLine(pname, 0)
                    body   = cm.ProcBodyLine(pname, 0)
                    pcount = cm.ProcCountLines(pname, 0)
                    # Clamp count to not exceed total_lines
                    pcount = min(pcount, total - pstart + 1)
                    procs.append({"name": pname, "start_line": pstart,
                                  "body_line": body, "count": pcount})
                except Exception:
                    procs.append({"name": pname, "start_line": i})
    return {"total_lines": total, "procs": procs}


# ---------------------------------------------------------------------------
# VBE replace / edit operations
# ---------------------------------------------------------------------------

def _exec_single_replace(cm, app, object_type, object_name, start_line, count, new_code):
    """Executes a single replace_lines operation. Returns dict with result."""
    total = cm.CountOfLines
    if start_line < 1 or start_line > total + 1:
        raise ValueError(
            f"start_line {start_line} out of range (1–{total})"
        )
    clamped = False
    if count > 0:
        max_count = total - start_line + 1
        if count > max_count:
            count = max_count
            clamped = True
        cm.DeleteLines(start_line, count)
    inserted = 0
    if new_code:
        decoded = html_mod.unescape(new_code)
        normalized = decoded.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        cm.InsertLines(start_line, normalized)
        inserted = len(new_code.splitlines())
    end = start_line + count - 1 if count > 0 else start_line
    clamp_note = " (count adjusted)" if clamped else ""
    return {
        "start_line": start_line, "deleted": count, "inserted": inserted,
        "clamp_note": clamp_note, "end": end,
    }


def ac_vbe_replace_lines(
    db_path: str, object_type: str, object_name: str,
    start_line: int = 0, count: int = 0, new_code: str = "",
    operations: list = None,
) -> str:
    """
    Replaces 'count' lines starting at 'start_line' with 'new_code'.
    - count=0 → pure insertion (deletes nothing).
    - new_code='' → pure deletion (inserts nothing).
    new_code can be multiline (\\n or \\r\\n).

    Batch mode: if 'operations' is passed (list of {start_line, count, new_code}),
    all are executed in 1 call, automatically sorted bottom-to-top.
    In batch mode, individual start_line/count/new_code are ignored.

    Returns the status + preview of inserted code to avoid an extra get_proc call.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)

    if operations:
        # ── Batch mode: sort bottom-to-top and execute sequentially ──
        original_total = cm.CountOfLines
        sorted_ops = sorted(operations, key=lambda op: op["start_line"], reverse=True)
        results = []
        for op in sorted_ops:
            r = _exec_single_replace(
                cm, app, object_type, object_name,
                int(op["start_line"]), int(op["count"]), op.get("new_code", ""),
            )
            results.append(r)
        # Invalidar cache + persist
        cache_key = f"{object_type}:{object_name}"
        _vbe_code_cache.pop(cache_key, None)
        try:
            obj_type_code = AC_TYPE.get(object_type, 5)
            app.DoCmd.Save(obj_type_code, object_name)
        except Exception:
            pass
        new_total = cm.CountOfLines
        total_deleted = sum(r["deleted"] for r in results)
        total_inserted = sum(r["inserted"] for r in results)
        lines_summary = ", ".join(
            f"L{r['start_line']}" for r in results
        )
        # Health check with expected count
        expected = original_total - total_deleted + total_inserted
        health = _check_module_health(cm, cache_key, expected_total=expected)
        status = (
            f"OK batch: {len(results)} operations executed (bottom→top: {lines_summary}). "
            f"Total: {total_deleted} deleted, {total_inserted} inserted "
            f"→ module now has {new_total} lines"
        )
        if health:
            status += f"\n" + "\n".join(health)
        return status

    # ── Single mode (backward compatible) ──
    r = _exec_single_replace(cm, app, object_type, object_name, start_line, count, new_code)
    # Invalidate text cache (module changed)
    cache_key = f"{object_type}:{object_name}"
    _vbe_code_cache.pop(cache_key, None)
    # Persist VBE changes to .accdb — without this, changes are only in memory
    try:
        obj_type_code = AC_TYPE.get(object_type, 5)  # acModule=5 default
        app.DoCmd.Save(obj_type_code, object_name)
    except Exception:
        pass  # save is best-effort; compact/close will also persist
    new_total = cm.CountOfLines
    # Health check
    health = _check_module_health(cm, cache_key)
    status = (
        f"OK: lines {r['start_line']}–{r['end']} replaced "
        f"({r['deleted']} deleted, {r['inserted']} inserted){r['clamp_note']} "
        f"→ module now has {new_total} lines"
    )
    if health:
        status += f"\n" + "\n".join(health)
    if new_code:
        lines = new_code.splitlines()
        preview = (
            new_code if len(lines) <= 60
            else "\n".join(lines[:60]) + f"\n[... +{len(lines) - 60} lines]"
        )
        return f"{status}\n\n{preview}"
    return status


# ---------------------------------------------------------------------------
# VBE search operations
# ---------------------------------------------------------------------------

def ac_vbe_find(
    db_path: str, object_type: str, object_name: str,
    search_text: str, match_case: bool = False, use_regex: bool = False,
    proc_name: str = "",
) -> dict:
    """
    Searches text (or regex) in a module and returns all matching lines.
    Uses _vbe_code_cache to avoid re-reading the module if already read.

    If proc_name is passed, limits the search to that procedure's range.
    Always enriches each match with 'proc' (name of the owning procedure).
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    if not all_code:
        return {"found": False, "match_count": 0, "matches": []}

    # Determine search range
    search_start = 1
    search_end = len(all_code.splitlines())
    if proc_name:
        try:
            search_start = cm.ProcStartLine(proc_name, 0)
            proc_count = cm.ProcCountLines(proc_name, 0)
            search_end = min(search_start + proc_count - 1, search_end)
        except Exception as exc:
            raise RuntimeError(
                f"Procedure '{proc_name}' not found in '{object_name}': {exc}"
            )

    matches: list[dict] = []
    lines = all_code.splitlines()
    for i, raw_line in enumerate(lines, start=1):
        if i < search_start or i > search_end:
            continue
        if text_matches(search_text, raw_line, match_case, use_regex):
            # Enrich with owning procedure name
            owning_proc = ""
            try:
                owning_proc = cm.ProcOfLine(i, 0)
            except Exception:
                pass  # line outside a proc (Declarations, etc.)
            matches.append({
                "line": i, "content": raw_line.rstrip("\r"), "proc": owning_proc,
            })
    return {"found": bool(matches), "match_count": len(matches), "matches": matches}


def ac_vbe_search_all(
    db_path: str, search_text: str, match_case: bool = False,
    max_results: int = 100, use_regex: bool = False,
) -> dict:
    """
    Searches text (or regex) in ALL VBA modules (modules, forms, reports) of the database.
    Returns {total_matches, results: [...], truncated?: bool}.
    """
    # Lazy import to avoid circular dependency (code.py may import from vbe.py)
    from .code import ac_list_objects

    app = _Session.connect(db_path)
    objects = ac_list_objects(db_path, "all")
    results: list[dict] = []
    total = 0
    truncated = False

    for obj_type in ("module", "form", "report"):
        if truncated:
            break
        for obj_name in objects.get(obj_type, []):
            if truncated:
                break
            try:
                cm = _get_code_module(app, obj_type, obj_name)
                cache_key = f"{obj_type}:{obj_name}"
                all_code = _cm_all_code(cm, cache_key)
                if not all_code:
                    continue
                obj_matches: list[dict] = []
                for i, raw_line in enumerate(all_code.splitlines(), start=1):
                    if text_matches(search_text, raw_line, match_case, use_regex):
                        obj_matches.append({"line": i, "content": raw_line.rstrip("\r")})
                        total += 1
                        if total >= max_results:
                            truncated = True
                            break
                if obj_matches:
                    results.append({
                        "object_type": obj_type,
                        "object_name": obj_name,
                        "matches": obj_matches,
                    })
            except Exception:
                continue  # skip objects without accessible CodeModule

    out: dict = {"total_matches": total, "results": results}
    if truncated:
        out["truncated"] = True
    return out


def ac_search_queries(
    db_path: str, search_text: str, match_case: bool = False,
    max_results: int = 100, use_regex: bool = False,
) -> dict:
    """
    Searches text (or regex) in the SQL of ALL queries in the database.
    Returns {total_matches, results: [{query_name, sql}], truncated?: bool}.
    """
    app = _Session.connect(db_path)
    db = app.CurrentDb()
    results: list[dict] = []
    total = 0
    for qd in db.QueryDefs:
        name = qd.Name
        if name.startswith("~"):  # skip internal/temp queries
            continue
        sql = qd.SQL
        if text_matches(search_text, sql, match_case, use_regex):
            results.append({"query_name": name, "sql": sql.strip()})
            total += 1
            if total >= max_results:
                break
    out: dict = {"total_matches": total, "results": results}
    if total >= max_results:
        out["truncated"] = True
    return out


# ---------------------------------------------------------------------------
# Find usages — cross-reference search
# ---------------------------------------------------------------------------

def ac_find_usages(
    db_path: str, search_text: str, match_case: bool = False,
    max_results: int = 200, use_regex: bool = False,
) -> dict:
    """
    Searches for a name (function, table, field, variable) in VBA, queries and
    control properties of forms/reports. Returns grouped results.
    Reuses ac_vbe_search_all and ac_search_queries for VBA and queries.
    """
    # Lazy import to avoid circular dependency
    from .code import ac_list_objects

    # 1. VBA matches — delegates to ac_vbe_search_all
    vba_result = ac_vbe_search_all(
        db_path, search_text, match_case, max_results, use_regex,
    )
    # Flatten: from [{object_type, object_name, matches: [{line, content}]}] to flat list
    vba_matches: list[dict] = []
    for group in vba_result["results"]:
        for m in group["matches"]:
            vba_matches.append({
                "object_type": group["object_type"],
                "object_name": group["object_name"],
                "line": m["line"],
                "content": m["content"],
            })
    total = len(vba_matches)
    truncated = vba_result.get("truncated", False)

    # 2. Query matches — delegates to ac_search_queries
    query_matches: list[dict] = []
    if not truncated:
        remaining = max_results - total
        qry_result = ac_search_queries(
            db_path, search_text, match_case, remaining, use_regex,
        )
        query_matches = qry_result["results"]
        total += qry_result["total_matches"]
        truncated = qry_result.get("truncated", False)

    # 3. Control property matches — search in exports of forms/reports
    control_matches: list[dict] = []
    if not truncated:
        app = _Session.connect(db_path)
        objects = ac_list_objects(db_path, "all")
        for obj_type in ("form", "report"):
            if truncated:
                break
            for obj_name in objects.get(obj_type, []):
                if truncated:
                    break
                try:
                    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_")
                    os.close(fd)
                    try:
                        app.SaveAsText(AC_TYPE[obj_type], obj_name, tmp)
                        raw_text, _enc = read_tmp(tmp)
                    finally:
                        try:
                            os.unlink(tmp)
                        except OSError:
                            pass
                    for line in raw_text.splitlines():
                        stripped = line.strip()
                        for prop in CONTROL_SEARCH_PROPS:
                            if stripped.startswith(prop + " ="):
                                value_part = stripped[len(prop) + 2:].strip()
                                if text_matches(search_text, value_part, match_case, use_regex):
                                    control_matches.append({
                                        "object_type": obj_type,
                                        "object_name": obj_name,
                                        "property": prop,
                                        "value": value_part,
                                    })
                                    total += 1
                                    if total >= max_results:
                                        truncated = True
                                    break
                except Exception:
                    continue

    out: dict = {
        "search_text": search_text,
        "vba_matches": vba_matches,
        "query_matches": query_matches,
        "control_matches": control_matches,
        "total_matches": total,
    }
    if truncated:
        out["truncated"] = True
    return out


# ---------------------------------------------------------------------------
# VBE replace proc / patch / append
# ---------------------------------------------------------------------------

def ac_vbe_replace_proc(
    db_path: str, object_type: str, object_name: str,
    proc_name: str, new_code: str
) -> str:
    """
    Replaces a complete procedure (Sub/Function/Property) by name.
    Calculates boundaries automatically via COM (ProcStartLine/ProcCountLines),
    eliminating calculation errors from the caller.
    If new_code is empty, deletes the procedure.
    """
    app = _Session.connect(db_path)

    # If the form/report is open in Design view (after ac_set_control_props etc.),
    # close it first to avoid COM conflicts with the VBE ("Catastrophic error")
    if object_type in ("form", "report"):
        ac_obj_type = AC_FORM if object_type == "form" else AC_REPORT
        try:
            app.DoCmd.Close(ac_obj_type, object_name, AC_SAVE_YES)
            log.info("ac_vbe_replace_proc: closed '%s' in Design view before accessing VBE", object_name)
        except Exception:
            pass  # was not open — OK

    # Invalidate cm_cache in case CodeModule went stale after design operation
    cache_key = f"{object_type}:{object_name}"
    _Session._cm_cache.pop(cache_key, None)

    cm = _get_code_module(app, object_type, object_name)
    try:
        start = cm.ProcStartLine(proc_name, 0)
        count = cm.ProcCountLines(proc_name, 0)
    except Exception as exc:
        raise RuntimeError(
            f"Procedure '{proc_name}' not found in '{object_name}': {exc}"
        )
    # Clamp count to actual module total (ProcCountLines can inflate the last proc)
    total = cm.CountOfLines
    count = min(count, total - start + 1)
    # Backup original proc in RAM for rollback if it fails
    backup_code = cm.Lines(start, count)
    # Strip Option lines if proc is NOT at the top of the module
    option_warnings = []
    if new_code and start > 5:
        new_code, option_warnings = _strip_option_lines(new_code)
    # Delete old procedure and insert new one with automatic rollback
    try:
        cm.DeleteLines(start, count)
        inserted = 0
        if new_code:
            normalized = new_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
            cm.InsertLines(start, normalized)
            inserted = len(new_code.splitlines())
    except Exception:
        # Restore original code
        try:
            cm.InsertLines(start, backup_code)
        except Exception:
            pass  # best-effort restore
        raise
    # Invalidate cache
    cache_key = f"{object_type}:{object_name}"
    _vbe_code_cache.pop(cache_key, None)
    new_total = cm.CountOfLines
    # Health check
    health = _check_module_health(cm, cache_key)
    action = "replaced" if new_code else "deleted"
    status = (
        f"OK: proc '{proc_name}' {action} "
        f"({count} deleted, {inserted} inserted) "
        f"→ module now has {new_total} lines"
    )
    if option_warnings:
        status += f"\n" + "\n".join(option_warnings)
    if health:
        status += f"\n" + "\n".join(health)
    if new_code:
        lines = new_code.splitlines()
        preview = (
            new_code if len(lines) <= 60
            else "\n".join(lines[:60]) + f"\n[... +{len(lines) - 60} lines]"
        )
        return f"{status}\n\n{preview}"
    return status


def ac_vbe_patch_proc(
    db_path: str, object_type: str, object_name: str,
    proc_name: str, patches: list,
) -> str:
    """
    Applies surgical find/replace WITHIN a procedure without rewriting everything.
    patches: list of {find: str, replace: str}.
    More efficient than vbe_replace_proc when only a few lines change
    within a large proc (e.g.: 174 lines, only 15 change).
    """
    app = _Session.connect(db_path)

    # Close form/report in Design view if open
    if object_type in ("form", "report"):
        ac_obj_type = AC_FORM if object_type == "form" else AC_REPORT
        try:
            app.DoCmd.Close(ac_obj_type, object_name, AC_SAVE_YES)
        except Exception:
            pass

    cache_key = f"{object_type}:{object_name}"
    _Session._cm_cache.pop(cache_key, None)

    cm = _get_code_module(app, object_type, object_name)
    try:
        start = cm.ProcStartLine(proc_name, 0)
        count = cm.ProcCountLines(proc_name, 0)
    except Exception as exc:
        raise RuntimeError(
            f"Procedure '{proc_name}' not found in '{object_name}': {exc}"
        )
    total = cm.CountOfLines
    count = min(count, total - start + 1)

    # Get current proc code
    proc_code = cm.Lines(start, count)
    backup_code = proc_code

    # Apply patches sequentially
    applied = 0
    not_found = []
    ws_fallback_notes = []
    for i, patch in enumerate(patches):
        find_text = patch["find"]
        replace_text = patch.get("replace", "")
        # Decode HTML entities
        find_text = html_mod.unescape(find_text)
        replace_text = html_mod.unescape(replace_text)
        if find_text in proc_code:
            proc_code = proc_code.replace(find_text, replace_text, 1)
            applied += 1
        else:
            # Fallback: whitespace-normalized match
            ws_match = _ws_normalized_match(proc_code, find_text)
            if ws_match is not None:
                s_idx, e_idx = ws_match
                code_lines = proc_code.splitlines(keepends=True)
                # Replace matched lines with replace_text as-is
                replace_normalized = replace_text
                if not replace_normalized.endswith(("\r\n", "\n")) and replace_normalized:
                    replace_normalized += "\r\n"
                code_lines[s_idx : e_idx + 1] = [replace_normalized] if replace_normalized else []
                proc_code = "".join(code_lines)
                applied += 1
                ws_fallback_notes.append(f"patch[{i}]: matched via ws-normalized fallback")
            else:
                ctx = _closest_match_context(proc_code, find_text, proc_name)
                not_found.append(f"patch[{i}]: not found. {ctx}")

    if applied == 0:
        return f"NOOP: no patches matched in '{proc_name}'. Errors:\n" + "\n".join(not_found)

    # Strip Option lines if proc is NOT at the top of the module
    option_warnings = []
    if start > 5:
        proc_code, option_warnings = _strip_option_lines(proc_code)

    # Replace entire proc with patched code
    try:
        cm.DeleteLines(start, count)
        if proc_code.strip():
            normalized = proc_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
            cm.InsertLines(start, normalized)
    except Exception:
        try:
            cm.InsertLines(start, backup_code)
        except Exception:
            pass
        raise

    # Invalidate cache
    _vbe_code_cache.pop(cache_key, None)
    new_total = cm.CountOfLines
    new_count = cm.ProcCountLines(proc_name, 0) if applied > 0 else 0
    # Health check
    health = _check_module_health(cm, cache_key)
    result = (
        f"OK: {applied}/{len(patches)} patches applied in '{proc_name}' "
        f"({count} → {new_count} lines) → module now has {new_total} lines"
    )
    if ws_fallback_notes:
        result += f"\nWS-fallback: {'; '.join(ws_fallback_notes)}"
    if option_warnings:
        result += f"\n" + "\n".join(option_warnings)
    if health:
        result += f"\n" + "\n".join(health)
    if not_found:
        result += f"\nNot found:\n" + "\n".join(not_found)
    return result


def ac_vbe_append(
    db_path: str, object_type: str, object_name: str,
    code: str
) -> str:
    """
    Appends code to the end of a VBA module.
    Safer than replace_lines for inserting new functions
    without needing to calculate line numbers.
    """
    app = _Session.connect(db_path)
    cm = _get_code_module(app, object_type, object_name)
    total = cm.CountOfLines
    # Decode HTML entities that MCP transport may have encoded (& → &amp; etc.)
    decoded = html_mod.unescape(code)
    # Strip Option lines that would end up misplaced at the end of the module
    decoded, option_warnings = _strip_option_lines(decoded)
    if not decoded.strip():
        return "NOOP: code contained only Option lines (stripped to prevent misplacement)"
    normalized = decoded.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    cm.InsertLines(total + 1, normalized)
    inserted = len(decoded.splitlines())
    cache_key = f"{object_type}:{object_name}"
    _vbe_code_cache.pop(cache_key, None)
    # Persist VBE changes to .accdb
    try:
        obj_type_code = AC_TYPE.get(object_type, 5)
        app.DoCmd.Save(obj_type_code, object_name)
    except Exception:
        pass
    new_total = cm.CountOfLines
    # Health check
    health = _check_module_health(cm, cache_key)
    result = f"OK: {inserted} lines appended → module now has {new_total} lines"
    if option_warnings:
        result += f"\n" + "\n".join(option_warnings)
    if health:
        result += f"\n" + "\n".join(health)
    return result
