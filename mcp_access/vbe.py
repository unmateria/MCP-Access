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
    AC_TYPE, _Session, _parsed_controls_cache, log,
    invalidate_object_caches, _get_vb_project,
)
from .constants import (
    VBE_PREFIX, AC_FORM, AC_REPORT, AC_SAVE_YES, AC_DESIGN,
    CONTROL_SEARCH_PROPS,
)
from .helpers import text_matches, read_tmp


# ---------------------------------------------------------------------------
# Property procedure helpers (v0.7.23 — all 4 VBE proc kinds)
# ---------------------------------------------------------------------------
# VBE ProcStartLine/ProcBodyLine/ProcCountLines/ProcOfLine require a ``kind``
# argument.  The VBE enum ``vbext_ProcKind`` has four values:
#   0 = vbext_pk_Proc     (Sub / Function)
#   1 = vbext_pk_Let      (Property Let)
#   2 = vbext_pk_Set      (Property Set)
#   3 = vbext_pk_Get      (Property Get)
# Prior code only tried kind=0 and kind=3, so Property Let and Property Set
# were invisible.  We now iterate all four kinds.

_VBEXT_PK_PROC = 0
_VBEXT_PK_LET = 1
_VBEXT_PK_SET = 2
_VBEXT_PK_GET = 3
_ALL_PROC_KINDS = (0, 1, 2, 3)

# Maps regex-captured keyword → VBE kind for ac_vbe_module_info
_KEYWORD_TO_KIND: dict[str, int] = {
    "sub": 0,
    "function": 0,
    "property get": 3,
    "property let": 1,
    "property set": 2,
}


_KIND_LABEL = {0: "Sub/Function", 1: "Property Let", 2: "Property Set", 3: "Property Get"}


def _proc_kind(cm, name: str) -> int:
    """Return the VBE ``kind`` constant (0–3) for which *name* exists.

    Raises if *name* matches MULTIPLE kinds — a class with both
    ``Property Get Foo`` and ``Property Let Foo`` is normal VBA, and the
    caller has to disambiguate (e.g. by using ``ac_vbe_module_info`` first
    and then editing via line numbers with ``ac_vbe_replace_lines``)."""
    found = []
    for kind in _ALL_PROC_KINDS:
        try:
            cm.ProcStartLine(name, kind)
            found.append(kind)
        except Exception:
            continue
    if not found:
        raise RuntimeError(f"Procedure '{name}' not found with any VBE kind (0-3)")
    if len(found) > 1:
        labels = ", ".join(f"{_KIND_LABEL[k]} (kind={k})" for k in found)
        raise RuntimeError(
            f"Procedure '{name}' is ambiguous — exists as: {labels}. "
            f"Use ac_vbe_module_info to inspect them and edit via "
            f"ac_vbe_replace_lines with explicit line numbers."
        )
    return found[0]


def _proc_bounds(cm, name: str, kind: int = None):
    """Return ``(start, body, count, kind)`` for procedure *name*.

    If *kind* is given, uses it directly; otherwise discovers via ``_proc_kind``.
    """
    if kind is None:
        kind = _proc_kind(cm, name)
    start = cm.ProcStartLine(name, kind)
    body = cm.ProcBodyLine(name, kind)
    count = cm.ProcCountLines(name, kind)
    return start, body, count, kind


def _proc_of_line(cm, line: int) -> str:
    """Return the procedure name that owns *line*, or ``""``."""
    for kind in _ALL_PROC_KINDS:
        try:
            return cm.ProcOfLine(line, kind)
        except Exception:
            continue
    return ""


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
        project = _get_vb_project(app)
        component = project.VBComponents(component_name)
        cm = component.CodeModule
        _Session._cm_cache[cache_key] = cm
        return cm
    except Exception as exc:
        # After decompile+compact, VBComponents may be uninitialised.
        # Force VBE to recognise the component and retry once.
        log.info("_get_code_module: first attempt failed for '%s', forcing VBE init: %s",
                 component_name, exc)
        try:
            _force_vbe_init(app, object_type, object_name)
            project = _get_vb_project(app)
            component = project.VBComponents(component_name)
            cm = component.CodeModule
            _Session._cm_cache[cache_key] = cm
            log.info("_get_code_module: retry succeeded for '%s'", component_name)
            return cm
        except Exception:
            pass  # fall through to original error
        _Session._cm_cache.pop(cache_key, None)
        hint = (
            "Is 'Trust access to the VBA project object model' enabled "
            "in Access Trust Center settings?"
        )
        if object_type in ("form", "report"):
            hint = (
                f"For forms/reports: the component is only created when "
                f"HasModule=True. _force_vbe_init already tried to activate "
                f"it but failed. Either: (1) call access_set_form_property "
                f"with {{'HasModule': true}} explicitly, or (2) check that "
                f"'Trust access to the VBA project object model' is enabled "
                f"in Access Trust Center settings."
            )
        raise RuntimeError(
            f"Could not access CodeModule '{component_name}'. {hint}\n"
            f"Error: {exc}"
        )


def _force_vbe_init(app, object_type: str, object_name: str):
    """Force VBE to recognise a component after decompile+compact OR after
    a brand-new form/report was created without a VBA code module.

    For forms/reports: open in Design view, *flip HasModule to True if it
    is False* (a freshly-created form has no module — VBComponents won't
    find it until HasModule=True), then close. This makes Access load the
    VBA code-behind so VBComponents can find it.

    For modules: toggle VBE.MainWindow.Visible to trigger enumeration.
    """
    if object_type in ("form", "report"):
        ac_obj = AC_FORM if object_type == "form" else AC_REPORT
        try:
            if object_type == "form":
                app.DoCmd.OpenForm(object_name, AC_DESIGN)
                obj = app.Forms(object_name)
            else:
                app.DoCmd.OpenReport(object_name, AC_DESIGN)
                obj = app.Reports(object_name)
            # Activate code module if absent — this is the common case for
            # forms created via ac_create_form which start with HasModule=False.
            try:
                if not bool(obj.HasModule):
                    obj.HasModule = True
                    log.info(
                        "_force_vbe_init: activated HasModule on '%s' "
                        "(form had no code module yet)",
                        object_name,
                    )
            except Exception as e:
                log.warning(
                    "_force_vbe_init: HasModule check/set failed for '%s': %s",
                    object_name, e,
                )
            app.DoCmd.Close(ac_obj, object_name, AC_SAVE_YES)
            log.info("_force_vbe_init: opened/closed '%s' in Design view", object_name)
        except Exception as e:
            log.warning("_force_vbe_init: open/close failed for '%s': %s", object_name, e)
    else:
        try:
            vbe = app.VBE
            was_visible = vbe.MainWindow.Visible
            vbe.MainWindow.Visible = True
            if not was_visible:
                vbe.MainWindow.Visible = False
            log.info("_force_vbe_init: toggled VBE.MainWindow.Visible")
        except Exception as e:
            log.warning("_force_vbe_init: VBE toggle failed: %s", e)


def _close_form_design_view(app: Any, object_type: str, object_name: str) -> None:
    """If the form/report is open in Design view, close it (saving changes).

    Required before ANY VBE CodeModule access — including reads — because
    Access can raise "Catastrophic failure" (-2147418113) when the Design
    view holds the same object the VBE proxy is being queried for.
    No-op for object_type='module' (standard modules have no Design view).
    """
    if object_type not in ("form", "report"):
        return
    ac_obj_type = AC_FORM if object_type == "form" else AC_REPORT
    try:
        app.DoCmd.Close(ac_obj_type, object_name, AC_SAVE_YES)
    except Exception:
        pass  # not open in Design view — that's the common case


def _cm_all_code(cm: Any, cache_key: str) -> str:
    """
    Returns the full text of a CodeModule by reading directly from COM.
    Previously cached in _vbe_code_cache, but the cache could not detect
    edits made outside the MCP (manual VBE edits, Ctrl+Z, add-ins) and
    served stale text. See GitHub issue #26.

    The ``cache_key`` parameter is kept for call-site compatibility and is
    unused.
    """
    total = cm.CountOfLines
    return cm.Lines(1, total) if total > 0 else ""


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

    # Check 2 — Duplicate labels (scoped per procedure).
    # VBA accepts combinations like "Public Static Sub Foo" — allow scope
    # modifier AND optional Static.
    label_re = re.compile(r'^(\w+):\s*$')
    proc_re = re.compile(r'^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+\w+)\s+', re.IGNORECASE)
    end_proc_re = re.compile(r'^End\s+(?:Sub|Function|Property)\b', re.IGNORECASE)
    label_positions: dict[tuple[str, str], list[int]] = {}
    current_proc = ""
    for i, line in enumerate(lines):
        stripped = line.strip()
        if proc_re.match(stripped):
            current_proc = stripped
        elif end_proc_re.match(stripped):
            current_proc = ""
        # Skip comments, Case statements, pure numbers
        if stripped.startswith("'") or stripped.startswith("Case "):
            continue
        m = label_re.match(stripped)
        if m:
            label = m.group(1)
            # Exclude numeric labels and common non-label patterns
            if label.isdigit():
                continue
            label_positions.setdefault((current_proc, label), []).append(i + 1)
    for (proc, label), positions in label_positions.items():
        if len(positions) > 1:
            warnings.append(
                f"WARNING: Duplicate label '{label}:' at lines {positions}"
                + (f" in '{proc}'" if proc else "")
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
    _close_form_design_view(app, object_type, object_name)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    all_lines = all_code.splitlines()
    total = len(all_lines)
    if start_line < 1 or start_line > total:
        raise ValueError(f"start_line {start_line} out of range (1-{total})")
    actual = min(count, total - start_line + 1)
    if actual < count:
        log.info(
            "ac_vbe_get_lines: requested %d but only %d available from line %d",
            count, actual, start_line,
        )
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
    _close_form_design_view(app, object_type, object_name)
    cm = _get_code_module(app, object_type, object_name)
    try:
        start, body, count, _kind = _proc_bounds(cm, proc_name)
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
    _close_form_design_view(app, object_type, object_name)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    all_lines = all_code.splitlines()
    total = len(all_lines)
    procs: list[dict] = []
    if total > 0:
        seen: set[tuple[str, str]] = set()  # (name_lower, keyword_lower)
        for i, raw_line in enumerate(all_lines, start=1):
            m = re.match(
                r'^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
                r'(Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)',
                raw_line.strip(), re.IGNORECASE,
            )
            if m:
                keyword = m.group(1)   # e.g. "Property Let"
                pname = m.group(2)
                dedup_key = (pname.lower(), keyword.lower())
                if dedup_key in seen:
                    continue
                seen.add(dedup_key)
                kind = _KEYWORD_TO_KIND.get(keyword.lower())
                try:
                    pstart, body, pcount, _kind = _proc_bounds(cm, pname, kind=kind)
                    # Clamp count to not exceed total_lines
                    pcount = min(pcount, total - pstart + 1)
                    procs.append({"name": pname, "keyword": keyword,
                                  "start_line": pstart, "body_line": body,
                                  "count": pcount})
                except Exception:
                    # VBE failed to locate this variant — scan forward
                    # in the source text for the matching End keyword.
                    end_kw = ("end property" if keyword.lower().startswith("property")
                              else f"end {keyword}".lower())
                    count = 1
                    for j in range(i - 1, total):  # 0-based scan from declaration
                        if all_lines[j].strip().lower() == end_kw:
                            count = (j + 1) - i + 1  # both 1-based, inclusive
                            break
                    procs.append({"name": pname, "keyword": keyword,
                                  "start_line": i, "body_line": i,
                                  "count": count})
    return {"total_lines": total, "procs": procs}


# ---------------------------------------------------------------------------
# VBE replace / edit operations
# ---------------------------------------------------------------------------

def _exec_single_replace(cm, object_type, object_name, start_line, count, new_code):
    """Executes a single replace_lines operation. Returns dict with result."""
    total = cm.CountOfLines
    # Allow start_line == total + 1 for "append at end" semantics, but make
    # the error message reflect that the inclusive upper bound is total + 1
    # for inserts (count == 0) and total for deletes / replaces.
    if start_line < 1 or start_line > total + 1:
        raise ValueError(
            f"start_line {start_line} out of range "
            f"(1-{total} for replace/delete, 1-{total + 1} for pure insert)"
        )
    clamped = False
    if count > 0:
        max_count = total - start_line + 1
        if count > max_count:
            count = max_count
            clamped = True
        # After clamp, count may become 0 when start_line == total + 1.
        # DeleteLines(line, 0) raises in VBE, so only call it when we
        # actually have lines to delete.
        if count > 0:
            cm.DeleteLines(start_line, count)
    inserted = 0
    if new_code:
        decoded = html_mod.unescape(new_code)
        normalized = decoded.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        pre_insert_total = total - count if count > 0 else total
        cm.InsertLines(start_line, normalized)
        # Ask VBE directly: splitlines() drops a trailing blank line that
        # InsertLines does count (when new_code ends with \r\n).
        inserted = cm.CountOfLines - pre_insert_total
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
    _close_form_design_view(app, object_type, object_name)

    cache_key_pre = f"{object_type}:{object_name}"
    _Session._cm_cache.pop(cache_key_pre, None)

    cm = _get_code_module(app, object_type, object_name)

    if operations:
        # ── Batch mode: sort bottom-to-top and execute sequentially ──
        original_total = cm.CountOfLines
        sorted_ops = sorted(operations, key=lambda op: op["start_line"], reverse=True)
        results = []
        for op in sorted_ops:
            r = _exec_single_replace(
                cm, object_type, object_name,
                int(op["start_line"]), int(op["count"]), op.get("new_code", ""),
            )
            results.append(r)
        # Persist
        cache_key = f"{object_type}:{object_name}"
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
    r = _exec_single_replace(cm, object_type, object_name, start_line, count, new_code)
    cache_key = f"{object_type}:{object_name}"
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

    If proc_name is passed, limits the search to that procedure's range.
    Always enriches each match with 'proc' (name of the owning procedure).
    """
    app = _Session.connect(db_path)
    _close_form_design_view(app, object_type, object_name)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    all_code = _cm_all_code(cm, cache_key)
    if not all_code:
        return {"found": False, "match_count": 0, "matches": []}

    # Determine search range
    search_start = 1
    search_end = len(all_code.splitlines())
    # Treat whitespace-only / empty proc_name as "search the whole module"
    # (callers that omit the arg send "" rather than None via MCP schema).
    if proc_name and proc_name.strip():
        try:
            p_start, _p_body, p_count, _p_kind = _proc_bounds(cm, proc_name)
            search_start = p_start
            search_end = min(p_start + p_count - 1, search_end)
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
            owning_proc = _proc_of_line(cm, i)
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
                _close_form_design_view(app, obj_type, obj_name)
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
    _close_form_design_view(app, object_type, object_name)

    # Invalidate cm_cache in case CodeModule went stale after design operation
    cache_key = f"{object_type}:{object_name}"
    _Session._cm_cache.pop(cache_key, None)

    cm = _get_code_module(app, object_type, object_name)
    try:
        start, _body, count, kind = _proc_bounds(cm, proc_name)
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
            pre_insert_total = total - count
            cm.InsertLines(start, normalized)
            inserted = cm.CountOfLines - pre_insert_total
    except Exception:
        # Restore original code
        try:
            cm.InsertLines(start, backup_code)
        except Exception:
            pass  # best-effort restore
        raise
    cache_key = f"{object_type}:{object_name}"
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

    _close_form_design_view(app, object_type, object_name)

    cache_key = f"{object_type}:{object_name}"
    _Session._cm_cache.pop(cache_key, None)

    cm = _get_code_module(app, object_type, object_name)
    try:
        start, _body, count, kind = _proc_bounds(cm, proc_name)
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
    ambiguous_notes = []
    for i, patch in enumerate(patches):
        find_text = patch["find"]
        replace_text = patch.get("replace", "")
        # Decode HTML entities
        find_text = html_mod.unescape(find_text)
        replace_text = html_mod.unescape(replace_text)
        # Normalize line endings to CRLF (proc_code from VBE is always CRLF;
        # callers commonly send LF — without this the exact match below
        # always falls through to the ws-normalized fallback).
        if "\n" in find_text:
            find_text = find_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        if "\n" in replace_text:
            replace_text = replace_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        if find_text in proc_code:
            # Warn if the find string is ambiguous — only the FIRST occurrence
            # is replaced (replace(count=1)), so multiple matches silently
            # leaving the others alone is a common source of confusion.
            occurrences = proc_code.count(find_text)
            if occurrences > 1:
                ambiguous_notes.append(
                    f"patch[{i}]: find_text matched {occurrences} times — only first occurrence replaced"
                )
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

    # Persist VBE changes to .accdb — without this, patches to form/report
    # code-behind can be lost because the object's dirty flag is not set.
    try:
        obj_type_code = AC_TYPE.get(object_type, 5)
        app.DoCmd.Save(obj_type_code, object_name)
    except Exception:
        pass
    new_total = cm.CountOfLines
    try:
        new_count = cm.ProcCountLines(proc_name, kind) if applied > 0 else 0
    except Exception:
        new_count = 0
    # Health check
    health = _check_module_health(cm, cache_key)
    result = (
        f"OK: {applied}/{len(patches)} patches applied in '{proc_name}' "
        f"({count} → {new_count} lines) → module now has {new_total} lines"
    )
    if ws_fallback_notes:
        result += f"\nWS-fallback: {'; '.join(ws_fallback_notes)}"
    if ambiguous_notes:
        result += f"\nAmbiguous matches: {'; '.join(ambiguous_notes)}"
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
    _close_form_design_view(app, object_type, object_name)

    cache_key_pre = f"{object_type}:{object_name}"
    _Session._cm_cache.pop(cache_key_pre, None)

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
    inserted = cm.CountOfLines - total
    cache_key = f"{object_type}:{object_name}"
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


# ---------------------------------------------------------------------------
# Find definition — "Go To Definition" for VBA symbols
# (requested by Tom — @TvanStiphout-Home, thanks!)
# ---------------------------------------------------------------------------

_FD_PROC_RE = re.compile(
    r'^\s*(?:(Public|Private|Friend|Global|Static)\s+)?'
    r'(?:(?:Static|Default)\s+)?'
    r'(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+(\w+)',
    re.IGNORECASE,
)
_FD_END_PROC_RE = re.compile(r'^\s*End\s+(Sub|Function|Property)\b', re.IGNORECASE)
_FD_CONST_RE = re.compile(r'^\s*(?:(?:Public|Private|Global)\s+)?Const\s+', re.IGNORECASE)
_FD_ENUM_RE = re.compile(r'^\s*(?:(Public|Private)\s+)?Enum\s+(\w+)', re.IGNORECASE)
_FD_END_ENUM_RE = re.compile(r'^\s*End\s+Enum\b', re.IGNORECASE)
_FD_TYPE_RE = re.compile(r'^\s*(?:(Public|Private)\s+)?Type\s+(\w+)', re.IGNORECASE)
_FD_END_TYPE_RE = re.compile(r'^\s*End\s+Type\b', re.IGNORECASE)
_FD_DECLARE_RE = re.compile(
    r'^\s*(?:(Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(Sub|Function)\s+(\w+)',
    re.IGNORECASE,
)
# Variable decl: starts with Public/Private/Global/Dim, followed by something
# that is NOT Const/Enum/Type/Sub/Function/Property/Declare.
_FD_VAR_RE = re.compile(
    r'^\s*(Public|Private|Global|Dim)\s+'
    r'(?!Const\b|Enum\b|Type\b|Sub\b|Function\b|Property\b|Declare\b)',
    re.IGNORECASE,
)
_FD_ENUM_MEMBER_RE = re.compile(r'^\s*(\w+)(?:\s*=\s*([^\']+?))?\s*(?:\'.*)?$')
_FD_TYPE_FIELD_RE = re.compile(
    r'^\s*(\w+)(?:\([^)]*\))?\s+As\s+(.+?)(?:\s*\'.*)?$', re.IGNORECASE,
)


def _split_top_level_commas(s: str) -> list[str]:
    """Split string by commas that are not inside parens or double quotes.

    Note on VBA's "" escape: an embedded double-quote inside a VBA string is
    written as "". This naive in_quote toggle flip-flops twice on each "",
    but the net state at the end of a well-formed string is correct, and
    real commas only appear outside strings — so splits land in the right
    place for any valid VBA source.
    """
    parts: list[str] = []
    buf: list[str] = []
    depth = 0
    in_quote = False
    for ch in s:
        if ch == '"':
            in_quote = not in_quote
            buf.append(ch)
        elif ch == '(' and not in_quote:
            depth += 1
            buf.append(ch)
        elif ch == ')' and not in_quote:
            depth -= 1
            buf.append(ch)
        elif ch == ',' and not in_quote and depth == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    if buf:
        parts.append("".join(buf))
    return parts


def _strip_trailing_vba_comment(line: str) -> str:
    """Strip a trailing VBA comment (' ...) from a line, respecting string
    literals. Returns the line without the comment and without trailing
    whitespace.

    VBA comments start with an apostrophe that is OUTSIDE any "..." string.
    Same state machine as _split_top_level_commas.
    """
    in_quote = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_quote = not in_quote
        elif ch == "'" and not in_quote:
            return line[:i].rstrip()
    return line.rstrip()


def _join_continuations(lines: list[str]) -> list[tuple[int, str]]:
    """Join VBA line continuations (` _` at end of a line) into single
    logical lines.

    Returns a list of ``(first_line_number, joined_text)`` tuples, where
    ``first_line_number`` is the 1-based line number of the FIRST physical
    line of the logical statement — so downstream reporting still points at
    where the declaration starts.

    A continuation is a line that, after stripping trailing whitespace and
    ignoring a trailing VBA comment, ends with ``_`` preceded by whitespace
    (or is exactly ``_``). Continuations can chain.
    """
    result: list[tuple[int, str]] = []
    i = 0
    n = len(lines)
    while i < n:
        first_idx = i
        # Build the logical line. We keep going while the CURRENT physical
        # line (after comment-strip) ends with whitespace + '_'.
        accumulated_parts: list[str] = []
        while True:
            raw = lines[i].rstrip("\r")
            no_comment = _strip_trailing_vba_comment(raw)
            cont_match = re.search(r'(?:^|\s)_\s*$', no_comment)
            if cont_match and i + 1 < n:
                accumulated_parts.append(no_comment[:cont_match.start()].rstrip())
                i += 1
                continue
            accumulated_parts.append(no_comment)
            break
        # First part keeps its leading indentation (regex patterns use ^\s*).
        # Continuation parts get their leading whitespace trimmed — the join
        # space takes its place — so "= _\n   &H1000" becomes "= &H1000".
        pieces = [
            p if idx == 0 else p.lstrip()
            for idx, p in enumerate(accumulated_parts)
        ]
        joined = " ".join(p for p in pieces if p)
        result.append((first_idx + 1, joined))
        i += 1
    return result


def ac_find_definition(
    db_path: str, symbol: str, kinds: list | None = None,
    match_case: bool = False,
    scan_types: list | None = None,
    first_only: bool = False,
) -> dict:
    """
    "Go To Definition" for VBA symbols — the mirror of ac_find_usages.

    Scans every VBA code module (standard modules, form code-behind, report
    code-behind) for DECLARATIONS of the given symbol and returns where each
    one lives (object, line, declaration text, scope).

    Detects:
      - const          Public/Private/Global Const FOO = ...  (multi on one line OK)
      - enum           Public/Private Enum MyEnum
      - enum_member    lines inside an Enum ... End Enum block
      - type           Public/Private Type MyStruct
      - type_field     lines inside a Type ... End Type block
      - sub            [Public|Private] Sub Name(...)
      - function       [Public|Private] Function Name(...) [As Type]
      - property       Property Get/Let/Set Name(...)   (incl. Default Property)
      - declare        [Public|Private] Declare [PtrSafe] Sub/Function Name Lib "..."
      - variable       module-level Dim/Public/Private/Global NAME [As ...]
                       (vars inside Sub/Function/Property are NOT reported —
                       those are locals, not definitions in the "go to" sense).

    Line continuations (` _` at end of line) are joined into a single
    logical statement before matching, so multi-line declarations resolve
    correctly. ``line`` always points at the FIRST physical line.

    Args:
        db_path: path to .accdb/.mdb
        symbol:  name to resolve (e.g. "dbAccess", "modGlobal", "ccRed")
        kinds:   optional whitelist, any subset of the 10 kinds above.
                 Default: all kinds.
        match_case: VBA is case-insensitive, so default False.
        scan_types: which object types to scan. Default ["module", "form",
                    "report"]. Pass ["module"] to skip forms/reports — much
                    faster on large DBs, since form/report code-behind needs
                    a Design-view open/close round-trip per object when the
                    VBComponent cache is cold.
        first_only: stop after the first match. Good for unique names.

    Returns:
        {"symbol", "total", "definitions": [ ... ]}
    """
    # Lazy import to avoid circular dependency
    from .code import ac_list_objects

    VALID_KINDS = {
        "const", "enum", "enum_member", "type", "type_field",
        "sub", "function", "property", "declare", "variable",
    }
    if kinds:
        bad = [k for k in kinds if k not in VALID_KINDS]
        if bad:
            raise ValueError(
                f"Invalid kind(s) {bad}. Valid: {sorted(VALID_KINDS)}"
            )
        kinds_filter = set(kinds)
    else:
        kinds_filter = VALID_KINDS

    VALID_SCAN_TYPES = ("module", "form", "report")
    if scan_types:
        bad_st = [t for t in scan_types if t not in VALID_SCAN_TYPES]
        if bad_st:
            raise ValueError(
                f"Invalid scan_types {bad_st}. Valid: {list(VALID_SCAN_TYPES)}"
            )
        scan_order = tuple(t for t in VALID_SCAN_TYPES if t in scan_types)
    else:
        scan_order = VALID_SCAN_TYPES

    if match_case:
        def name_matches(n: str) -> bool:
            return n == symbol
    else:
        symbol_lower = symbol.lower()
        def name_matches(n: str) -> bool:
            return n.lower() == symbol_lower

    app = _Session.connect(db_path)
    objects = ac_list_objects(db_path, "all")
    definitions: list[dict] = []

    def _stop() -> bool:
        return first_only and bool(definitions)

    for obj_type in scan_order:
        if _stop():
            break
        for obj_name in objects.get(obj_type, []):
            if _stop():
                break
            try:
                cm = _get_code_module(app, obj_type, obj_name)
                cache_key = f"{obj_type}:{obj_name}"
                all_code = _cm_all_code(cm, cache_key)
            except Exception:
                continue  # skip modules we cannot access
            if not all_code:
                continue

            # Fold line continuations; each tuple = (first_physical_line, clean_text).
            # clean_text already has trailing VBA comments stripped, respecting
            # "..." string literals — so value-extraction regex can be greedy-safe.
            logical = _join_continuations(all_code.splitlines())
            in_proc = False
            in_enum = False
            in_type = False
            current_enum = ""
            current_type = ""

            for (i, stripped) in logical:
                if _stop():
                    break
                # Inside proc — only watch for End, ignore everything else
                if in_proc:
                    if _FD_END_PROC_RE.match(stripped):
                        in_proc = False
                    continue

                # Inside enum — every non-empty line is a member
                if in_enum:
                    if _FD_END_ENUM_RE.match(stripped):
                        in_enum = False
                        current_enum = ""
                        continue
                    if not stripped or "enum_member" not in kinds_filter:
                        continue
                    m = _FD_ENUM_MEMBER_RE.match(stripped)
                    if m and name_matches(m.group(1)):
                        value = (m.group(2) or "").strip()
                        definitions.append({
                            "kind": "enum_member",
                            "object_type": obj_type,
                            "object_name": obj_name,
                            "line": i,
                            "declaration": stripped.strip(),
                            "parent_enum": current_enum,
                            "value": value or None,
                        })
                        if _stop():
                            break
                    continue

                # Inside type — every non-empty "Name As Type" line is a field
                if in_type:
                    if _FD_END_TYPE_RE.match(stripped):
                        in_type = False
                        current_type = ""
                        continue
                    if not stripped or "type_field" not in kinds_filter:
                        continue
                    m = _FD_TYPE_FIELD_RE.match(stripped)
                    if m and name_matches(m.group(1)):
                        definitions.append({
                            "kind": "type_field",
                            "object_type": obj_type,
                            "object_name": obj_name,
                            "line": i,
                            "declaration": stripped.strip(),
                            "parent_type": current_type,
                            "as_type": m.group(2).strip(),
                        })
                        if _stop():
                            break
                    continue

                # Module level — try patterns in order of specificity

                # Enum decl
                m = _FD_ENUM_RE.match(stripped)
                if m:
                    enum_name = m.group(2)
                    scope = (m.group(1) or "").strip() or None
                    in_enum = True
                    current_enum = enum_name
                    if "enum" in kinds_filter and name_matches(enum_name):
                        definitions.append({
                            "kind": "enum",
                            "object_type": obj_type,
                            "object_name": obj_name,
                            "line": i,
                            "declaration": stripped.strip(),
                            "scope": scope,
                        })
                        if _stop():
                            break
                    continue

                # Type decl
                m = _FD_TYPE_RE.match(stripped)
                if m:
                    type_name = m.group(2)
                    scope = (m.group(1) or "").strip() or None
                    in_type = True
                    current_type = type_name
                    if "type" in kinds_filter and name_matches(type_name):
                        definitions.append({
                            "kind": "type",
                            "object_type": obj_type,
                            "object_name": obj_name,
                            "line": i,
                            "declaration": stripped.strip(),
                            "scope": scope,
                        })
                        if _stop():
                            break
                    continue

                # Const decl (possibly multi: Const A = 1, B = 2)
                if _FD_CONST_RE.match(stripped):
                    if "const" in kinds_filter:
                        scope_m = re.match(
                            r'^\s*(Public|Private|Global)\s+Const',
                            stripped, re.IGNORECASE,
                        )
                        scope = scope_m.group(1) if scope_m else None
                        rest_m = re.match(
                            r'^\s*(?:(?:Public|Private|Global)\s+)?Const\s+(.+)$',
                            stripped, re.IGNORECASE,
                        )
                        if rest_m:
                            for part in _split_top_level_commas(rest_m.group(1)):
                                sub_m = re.match(
                                    r'^\s*(\w+)\s*(?:As\s+[\w.]+)?\s*=\s*(.+?)\s*$',
                                    part, re.IGNORECASE,
                                )
                                if sub_m and name_matches(sub_m.group(1)):
                                    definitions.append({
                                        "kind": "const",
                                        "object_type": obj_type,
                                        "object_name": obj_name,
                                        "line": i,
                                        "declaration": stripped.strip(),
                                        "scope": scope,
                                        "value": sub_m.group(2).strip(),
                                    })
                                    if _stop():
                                        break
                    continue

                # Declare decl (Win32 API)
                m = _FD_DECLARE_RE.match(stripped)
                if m:
                    if "declare" in kinds_filter:
                        scope = (m.group(1) or "").strip() or None
                        decl_kind = m.group(2)  # Sub or Function
                        decl_name = m.group(3)
                        if name_matches(decl_name):
                            definitions.append({
                                "kind": "declare",
                                "subkind": decl_kind,
                                "object_type": obj_type,
                                "object_name": obj_name,
                                "line": i,
                                "declaration": stripped.strip(),
                                "scope": scope,
                            })
                            if _stop():
                                break
                    continue

                # Sub/Function/Property decl
                m = _FD_PROC_RE.match(stripped)
                if m:
                    scope = (m.group(1) or "").strip() or None
                    proc_kw = re.sub(r'\s+', ' ', m.group(2).strip())
                    proc_name = m.group(3)
                    in_proc = True
                    if proc_kw.lower().startswith("property"):
                        kind_cat = "property"
                    elif proc_kw.lower() == "sub":
                        kind_cat = "sub"
                    else:
                        kind_cat = "function"
                    if kind_cat in kinds_filter and name_matches(proc_name):
                        entry: dict = {
                            "kind": kind_cat,
                            "object_type": obj_type,
                            "object_name": obj_name,
                            "line": i,
                            "declaration": stripped.strip(),
                            "scope": scope,
                        }
                        # subkind only carries extra information for property
                        # (Get/Let/Set) — for sub/function it's redundant with kind.
                        if kind_cat == "property":
                            entry["subkind"] = proc_kw
                        definitions.append(entry)
                        if _stop():
                            break
                    continue

                # Module-level variable decl
                if "variable" in kinds_filter and _FD_VAR_RE.match(stripped):
                    scope_m = re.match(
                        r'^\s*(Public|Private|Global|Dim)\s+(?:WithEvents\s+)?(.+)$',
                        stripped, re.IGNORECASE,
                    )
                    if scope_m:
                        scope = scope_m.group(1)
                        for part in _split_top_level_commas(scope_m.group(2)):
                            name_m = re.match(r'^\s*(\w+)', part)
                            if name_m and name_matches(name_m.group(1)):
                                definitions.append({
                                    "kind": "variable",
                                    "object_type": obj_type,
                                    "object_name": obj_name,
                                    "line": i,
                                    "declaration": stripped.strip(),
                                    "scope": scope,
                                })
                                if _stop():
                                    break

    return {
        "symbol": symbol,
        "total": len(definitions),
        "definitions": definitions,
    }
