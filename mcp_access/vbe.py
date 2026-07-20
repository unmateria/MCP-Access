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
import threading
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
# DoCmd.Save watchdog helper (v0.7.49 — issue #33)
# ---------------------------------------------------------------------------
# DoCmd.Save can pop an error dialog ("Save isn't available now") when the
# target module/form is open in the VBE.  Access shows the dialog and waits
# for a click before raising the COM exception — the bare except already
# swallows the error, but nothing dismissed the dialog, so the user had to
# click "OK" after every VBE write call.
#
# Fix: wrap DoCmd.Save in a lightweight watchdog thread (same pattern as
# _call_with_dialog_watchdog in maintenance.py) with a short 0.3 s grace
# period so the dialog is dismissed automatically and never reaches the user.

def _save_vbe_module(app, obj_type_code: int, object_name: str) -> None:
    """Call DoCmd.Save with a dialog-dismiss watchdog (best-effort)."""
    from .vba_exec import _dismiss_access_dialogs
    try:
        _h = app.hWndAccessApp
        hwnd = int(_h() if callable(_h) else _h)
    except Exception:
        hwnd = 0

    stop_event = threading.Event()

    def _watchdog():
        if stop_event.wait(0.3):
            return
        while not stop_event.is_set():
            if hwnd:
                try:
                    _dismiss_access_dialogs(hwnd)
                except Exception:
                    pass
            stop_event.wait(0.3)

    t = threading.Thread(target=_watchdog, daemon=True)
    t.start()
    try:
        app.DoCmd.Save(obj_type_code, object_name)
    except Exception:
        pass  # best-effort; compact/close will also persist
    finally:
        stop_event.set()


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

# Max per-object errors reported by the multi-object scans (search_all /
# find_usages / find_definition) before collapsing into the skipped count.
_SEARCH_ERROR_CAP = 20

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

    # Check 1 — Option placement. Option statements must precede all executable
    # code, but a comment/blank header of any length is perfectly legal (e.g. a
    # banner comment block). So flag an Option line only when real code already
    # appeared above it — NOT by a fixed line-number threshold, which false-
    # positives on long headers (a 6-line header pushes Option to line 7).
    seen_code = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        low = stripped.lower()
        if _OPTION_RE.match(line.rstrip('\r\n')):
            if seen_code:
                warnings.append(
                    f"WARNING: '{stripped}' found at line {i + 1} after executable "
                    f"code (Option statements must precede all code)"
                )
            continue
        # Blanks, comments and other Option-family lines are not "code".
        if not stripped or stripped.startswith("'") or low.startswith("rem ") \
                or low.startswith("option "):
            continue
        seen_code = True

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


def _ws_normalized_matches(
    proc_code: str, find_text: str, match_case: bool = True
) -> list[tuple[int, int]]:
    """
    Whitespace-tolerant matching: strips leading whitespace from each line
    and does a sliding window search.
    Returns ALL matches as (start_idx, end_idx) 0-based line index pairs into
    proc_code lines.  Overlapping windows are not reported — the scan skips
    past a match, mirroring how the replacement consumes it.
    """
    proc_lines = proc_code.splitlines()
    find_lines = find_text.splitlines()
    # Remove empty trailing lines from find_text
    while find_lines and not find_lines[-1].strip():
        find_lines.pop()
    if not find_lines:
        return []

    proc_stripped = [l.lstrip() for l in proc_lines]
    find_stripped = [l.lstrip() for l in find_lines]
    if not match_case:
        proc_stripped = [l.lower() for l in proc_stripped]
        find_stripped = [l.lower() for l in find_stripped]
    window = len(find_stripped)

    out: list[tuple[int, int]] = []
    i = 0
    while i <= len(proc_stripped) - window:
        if proc_stripped[i : i + window] == find_stripped:
            out.append((i, i + window - 1))
            i += window
        else:
            i += 1
    return out


def _ws_normalized_match(
    proc_code: str, find_text: str, match_case: bool = True
) -> tuple[int, int] | None:
    """First whitespace-normalized match, or None. See _ws_normalized_matches."""
    matches = _ws_normalized_matches(proc_code, find_text, match_case)
    return matches[0] if matches else None


def _case_insensitive_safe(text: str) -> bool:
    """
    True when ``text.lower()`` preserves length, so index positions computed on
    the lowered copy still address the original string.

    ``'İ'.lower()`` (U+0130) expands to TWO characters — a single such char in a
    VBA comment shifts every later index and would splice the replacement into
    the middle of a line.  When this returns False the case-insensitive tiers are
    skipped for that text rather than risking silent corruption.
    """
    return len(text.lower()) == len(text)


def _find_literal(hay: str, needle: str, match_case: bool) -> int | None:
    """Literal find. Returns the 0-based index of the first match, or None."""
    if match_case:
        idx = hay.find(needle)
    else:
        idx = hay.lower().find(needle.lower())
    return idx if idx >= 0 else None


def _count_literal(hay: str, needle: str, match_case: bool) -> int:
    """Number of non-overlapping literal occurrences of *needle* in *hay*."""
    if match_case:
        return hay.count(needle)
    return hay.lower().count(needle.lower())


_DECLARATIONS_TOKEN = "(declarations)"


def _is_declarations(proc_name: str) -> bool:
    """
    True when *proc_name* addresses the module's ``(Declarations)`` section —
    the lines above the first procedure, where Option/Const/Dim/Type live.

    Deliberately NOT triggered by an empty string: ``ac_vbe_find`` already reads
    ``""`` as "the whole module", and two contradictory meanings for the same
    value is worse than a slightly longer token.
    """
    return (proc_name or "").strip().lower() == _DECLARATIONS_TOKEN


def _vbe_line_count(text: str) -> int:
    """
    Number of lines VBE's ``InsertLines`` will create for *text*.

    VBE counts a trailing CRLF as opening a further (empty) line, which is
    exactly why ``splitlines()`` disagrees with ``CountOfLines`` — see the
    comment in ``_exec_single_replace``.  ``"a\\r\\nb"`` → 2, ``"a\\r\\nb\\r\\n"`` → 3.
    """
    if not text:
        return 0
    return text.count("\r\n") + 1


def _cm_lines_list(cm: Any, cache_key: str) -> list[str]:
    """
    Module text as a line list whose length ALWAYS equals ``cm.CountOfLines``.

    ``splitlines()`` drops the final empty line when the module ends in a blank
    one (VBE emits no trailing terminator), so it under-reports by 1 against the
    number VBE itself uses for InsertLines/DeleteLines addressing.  Padding to
    ``CountOfLines`` makes ``len(lines)`` authoritative, so every existing slice
    and bounds check keeps working while the reported total stops disagreeing
    with ``ac_vbe_patch_proc``.
    """
    total = cm.CountOfLines
    lines = _cm_all_code(cm, cache_key).splitlines()
    if len(lines) < total:
        lines.extend([""] * (total - len(lines)))
    return lines


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
# Patch engine (pure — no COM)
# ---------------------------------------------------------------------------
# ac_vbe_patch_proc's matching loop never touched the CodeModule, so it is
# extracted verbatim here and then extended.  Keeping it pure is what makes the
# ``atomic`` guarantee structural rather than a promise: the simulation and the
# commit are the SAME pass, so they cannot diverge.  A pre-pass that validated
# every anchor against the ORIGINAL text would be wrong in both directions —
# patch 0 can destroy (or create) the anchor patch 3 cites.

def _apply_patches(
    proc_code: str,
    patches: list,
    match_case: bool = False,
    require_unique: bool = False,
    proc_name: str = "",
    base_line: int = 1,
) -> dict:
    """
    Apply find/replace patches to *proc_code* sequentially.  No COM.

    Matching is a fixed 4-tier ladder, stopping at the first tier that hits:

      1. literal, case-sensitive
      2. whitespace-normalized, case-sensitive
      3. literal, case-insensitive          (only when ``match_case`` is False)
      4. whitespace-normalized, case-insensitive (idem)

    ALL case-sensitive tiers run before ANY case-insensitive one.  That order is
    what guarantees byte-for-byte compatibility: any call that succeeds today
    lands on tier 1 or 2 exactly as it did before, so relaxing the casing can
    only rescue calls that used to fail outright.

    ``base_line`` is the 1-based module line of ``proc_code``'s first line, so
    reported line numbers are absolute.

    Returns a dict with the patched ``code`` plus the report lists.  The caller
    decides whether to write — ``atomic`` lives there, not here.
    """
    applied = 0
    not_found: list[str] = []
    unique_violations: list[str] = []
    fallback_notes: list[str] = []
    ambiguous_notes: list[str] = []
    case_notes: list[str] = []

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

        # Unicode guard: the case-insensitive tiers index the ORIGINAL string
        # using offsets computed on a lowered copy, which only holds when
        # lowering preserves length.  Skip them otherwise (see
        # _case_insensitive_safe) rather than splice at a shifted offset.
        allow_ci = not match_case
        if allow_ci and not (
            _case_insensitive_safe(proc_code) and _case_insensitive_safe(find_text)
        ):
            allow_ci = False
            case_notes.append(
                f"patch[{i}]: case-insensitive matching skipped — text contains "
                "characters whose lowercase form changes length"
            )

        hit = False
        for tier, tier_case in (("exact", True), ("ws", True),
                                ("exact", False), ("ws", False)):
            if not tier_case and not allow_ci:
                continue

            if tier == "exact":
                idx = _find_literal(proc_code, find_text, tier_case)
                if idx is None:
                    continue
                occurrences = _count_literal(proc_code, find_text, tier_case)
                line_no = base_line + proc_code[:idx].count("\n")
                lines_hit = [line_no]
                if occurrences > 1:
                    # Collect the remaining occurrence line numbers for the report
                    scan_from = idx + len(find_text)
                    while True:
                        nxt = _find_literal(proc_code[scan_from:], find_text, tier_case)
                        if nxt is None:
                            break
                        abs_idx = scan_from + nxt
                        lines_hit.append(base_line + proc_code[:abs_idx].count("\n"))
                        scan_from = abs_idx + len(find_text)
                if require_unique and occurrences > 1:
                    unique_violations.append(
                        f"patch[{i}]: require_unique — find_text matched "
                        f"{occurrences} times at lines {lines_hit} "
                        f"({'case-sensitive' if tier_case else 'case-insensitive'} "
                        "comparison); patch NOT applied"
                    )
                    hit = True
                    break
                stored = proc_code[idx : idx + len(find_text)]
                proc_code = proc_code[:idx] + replace_text + proc_code[idx + len(find_text):]
            else:
                matches = _ws_normalized_matches(proc_code, find_text, tier_case)
                if not matches:
                    continue
                occurrences = len(matches)
                lines_hit = [base_line + m[0] for m in matches]
                if require_unique and occurrences > 1:
                    unique_violations.append(
                        f"patch[{i}]: require_unique — find_text matched "
                        f"{occurrences} times at lines {lines_hit} "
                        f"(whitespace-normalized, "
                        f"{'case-sensitive' if tier_case else 'case-insensitive'}"
                        "); patch NOT applied"
                    )
                    hit = True
                    break
                s_idx, e_idx = matches[0]
                code_lines = proc_code.splitlines(keepends=True)
                stored = "".join(code_lines[s_idx : e_idx + 1]).rstrip("\r\n")
                # Replace matched lines with replace_text as-is
                replace_normalized = replace_text
                if not replace_normalized.endswith(("\r\n", "\n")) and replace_normalized:
                    replace_normalized += "\r\n"
                code_lines[s_idx : e_idx + 1] = [replace_normalized] if replace_normalized else []
                proc_code = "".join(code_lines)

            applied += 1
            hit = True
            if tier == "ws":
                # Keep this exact wording — it predates the tier ladder.
                fallback_notes.append(f"patch[{i}]: matched via ws-normalized fallback")
            if not tier_case:
                # Echo what is actually stored: the caller's mental copy has the
                # wrong casing (the VBE rewrites it), so showing the real text is
                # what lets them fix it.
                fallback_notes.append(
                    f"patch[{i}]: matched case-insensitively (stored: {stored[:80]!r})"
                )
            if occurrences > 1 and not require_unique:
                ambiguous_notes.append(
                    f"patch[{i}]: find_text matched {occurrences} times at lines "
                    f"{lines_hit} — only first occurrence replaced"
                )
            break

        if not hit:
            ctx = _closest_match_context(proc_code, find_text, proc_name)
            not_found.append(f"patch[{i}]: not found. {ctx}")

    return {
        "code": proc_code,
        "applied": applied,
        "not_found": not_found,
        "unique_violations": unique_violations,
        "fallback_notes": fallback_notes,
        "ambiguous_notes": ambiguous_notes,
        "case_notes": case_notes,
    }


# ---------------------------------------------------------------------------
# VBE get operations
# ---------------------------------------------------------------------------

def ac_vbe_get_lines(
    db_path: str, object_type: str, object_name: str,
    start_line: int, count: int = None, end_line: int = None
) -> str:
    """Reads a range of lines without exporting the entire module."""
    if end_line is not None and count is None:
        if end_line < start_line:
            raise ValueError(
                f"end_line ({end_line}) must be >= start_line ({start_line})"
            )
        count = end_line - start_line + 1
    if count is None:
        raise ValueError("Either count or end_line must be provided")
    if count < 1:
        raise ValueError(f"count must be >= 1 (got {count})")
    app = _Session.connect(db_path)
    _close_form_design_view(app, object_type, object_name)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    # len(all_lines) == cm.CountOfLines by construction — a trailing blank line
    # is a real, addressable line in the editor and must stay readable.
    all_lines = _cm_lines_list(cm, cache_key)
    total = len(all_lines)
    if total == 0:
        raise ValueError(
            f"Module '{object_name}' is empty (0 lines) — nothing to read."
        )
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
      - start_line: VBE proc start — INCLUDES the blank/comment lines above the
        proc (use for whole-proc operations).
      - body_line: the Sub/Function/Property declaration line (use for body
        line-range edits).
    Pass proc_name="(Declarations)" to read the module's declarations section.
    """
    app = _Session.connect(db_path)
    _close_form_design_view(app, object_type, object_name)
    cm = _get_code_module(app, object_type, object_name)
    if _is_declarations(proc_name):
        start, body = 1, 1
        count = cm.CountOfDeclarationLines
        if count == 0:
            raise RuntimeError(
                f"Module '{object_name}' has no declarations section (0 lines)."
            )
    else:
        try:
            start, body, count, _kind = _proc_bounds(cm, proc_name)
        except Exception as exc:
            raise RuntimeError(
                f"Procedure '{proc_name}' not found in '{object_name}': {exc}"
            )
    # Extract text from cache instead of an extra cm.Lines call
    cache_key = f"{object_type}:{object_name}"
    all_lines = _cm_lines_list(cm, cache_key)
    # ProcCountLines can inflate the last proc past end of module (see
    # ac_vbe_replace_proc) — clamp so `count` matches the text returned.
    count = min(count, len(all_lines) - start + 1)
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
    Per proc: start_line (VBE proc start — includes preceding blank/comment
    lines) and body_line (the Sub/Function/Property declaration line).
    """
    app = _Session.connect(db_path)
    _close_form_design_view(app, object_type, object_name)
    cm = _get_code_module(app, object_type, object_name)
    cache_key = f"{object_type}:{object_name}"
    # Authoritative line count: len() == cm.CountOfLines by construction, so the
    # number reported here matches what ac_vbe_patch_proc reports (it always
    # used cm.CountOfLines) instead of being 1 short on blank-terminated modules.
    all_lines = _cm_lines_list(cm, cache_key)
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
                    # \b + optional trailing comment: "End Sub ' done" is
                    # valid VBA and must still close the proc.
                    end_re = re.compile(
                        r"^\s*" + re.escape(end_kw) + r"\s*(?:'.*)?$",
                        re.IGNORECASE,
                    )
                    count = 1
                    for j in range(i - 1, total):  # 0-based scan from declaration
                        if end_re.match(all_lines[j]):
                            count = (j + 1) - i + 1  # both 1-based, inclusive
                            break
                    procs.append({"name": pname, "keyword": keyword,
                                  "start_line": i, "body_line": i,
                                  "count": count})
    # Declarations section — addressable via proc_name="(Declarations)" in
    # ac_vbe_get_proc / ac_vbe_patch_proc, so the caller never has to infer its
    # boundary from the first procedure's start_line.
    try:
        decl_count = int(cm.CountOfDeclarationLines)
    except Exception:
        decl_count = 0
    return {
        "total_lines": total,
        "declarations": {"start_line": 1, "count": decl_count},
        "procs": procs,
    }


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
    if not operations and start_line < 1:
        # 0 is the "not provided" sentinel from the dispatcher — turn the
        # cryptic "start_line 0 out of range (1-N)" into an actionable error.
        raise ValueError(
            "start_line is required (1-based). Pass start_line/count/new_code "
            "for a single edit, or operations=[{start_line, count, new_code}, "
            "...] for batch mode."
        )
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
        _save_vbe_module(app, AC_TYPE.get(object_type, 5), object_name)
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
        # Same destructive no-op note as single mode: an operation that
        # deleted lines but inserted nothing usually means new_code arrived
        # empty (misnamed argument) — surface it instead of hiding it.
        destructive = [r for r in results if r["deleted"] > 0 and r["inserted"] == 0]
        if destructive:
            ops_desc = ", ".join(
                f"L{r['start_line']} ({r['deleted']} deleted)" for r in destructive
            )
            status += (
                f"\nnote: {len(destructive)} operation(s) deleted lines and "
                f"inserted nothing: {ops_desc}. If you meant to REPLACE, pass "
                f"the new text in new_code."
            )
        if health:
            status += f"\n" + "\n".join(health)
        return status

    # ── Single mode (backward compatible) ──
    r = _exec_single_replace(cm, object_type, object_name, start_line, count, new_code)
    cache_key = f"{object_type}:{object_name}"
    # Persist VBE changes to .accdb — without this, changes are only in memory
    _save_vbe_module(app, AC_TYPE.get(object_type, 5), object_name)
    new_total = cm.CountOfLines
    # Health check
    health = _check_module_health(cm, cache_key)
    status = (
        f"OK: lines {r['start_line']}–{r['end']} replaced "
        f"({r['deleted']} deleted, {r['inserted']} inserted){r['clamp_note']} "
        f"→ module now has {new_total} lines"
    )
    # Surface a destructive no-op: lines were deleted but nothing was inserted.
    # This is the footgun where new_code/new_lines arrives empty (e.g. a misnamed
    # argument) and a replace silently degrades into a pure delete.
    if r["deleted"] > 0 and r["inserted"] == 0:
        status += (
            f"\nnote: {r['deleted']} line(s) deleted and nothing inserted "
            f"(new_code/new_lines was empty). If you meant to REPLACE, pass the "
            f"new text in new_code or new_lines."
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
    errors: list[dict] = []
    skipped = 0
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
            except Exception as exc:
                # Never swallow this silently: if the whole VBA project fails
                # to load (broken reference, Trust Center...) every object
                # lands here and a clean "0 matches" would be a lie.
                skipped += 1
                if len(errors) < _SEARCH_ERROR_CAP:
                    errors.append({
                        "object": f"{obj_type}:{obj_name}",
                        "error": str(exc).splitlines()[0] if str(exc) else repr(exc),
                    })
                continue

    out: dict = {"total_matches": total, "results": results}
    if truncated:
        out["truncated"] = True
    if skipped:
        out["objects_skipped"] = skipped
        out["errors"] = errors
        out["warning"] = (
            f"{skipped} object(s) had no accessible CodeModule — results may "
            "be incomplete. If ALL objects failed, the VBA project is likely "
            "not loading (broken reference / Trust Center)."
        )
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
    errors: list[dict] = list(vba_result.get("errors", []))
    skipped = vba_result.get("objects_skipped", 0)

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
                except Exception as exc:
                    skipped += 1
                    if len(errors) < _SEARCH_ERROR_CAP:
                        errors.append({
                            "object": f"{obj_type}:{obj_name}",
                            "error": str(exc).splitlines()[0] if str(exc) else repr(exc),
                        })
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
    if skipped:
        out["objects_skipped"] = skipped
        out["errors"] = errors
        out["warning"] = (
            f"{skipped} object(s) could not be scanned — results may be "
            "incomplete. If ALL objects failed, the VBA project is likely "
            "not loading (broken reference / Trust Center)."
        )
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
    Preserves the blank separator line above the proc when replacing (delete/
    insert happen below the leading blanks).
    If new_code is empty, deletes the procedure AND its leading blank separator.
    """
    if _is_declarations(proc_name):
        # Refused on purpose: new_code="" would wipe Option Explicit and every
        # module-level Const in one unconfirmed call, and the leading-blank
        # (`lead`) logic below is meaningless at start=1.
        raise ValueError(
            "'(Declarations)' cannot be replaced wholesale. Use "
            "access_vbe_patch_proc(proc_name='(Declarations)') for surgical "
            "edits, or access_vbe_replace_lines for an explicit line range."
        )
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
    # Strip Option lines if proc is NOT at the top of the module
    option_warnings = []
    if new_code and start > 5:
        new_code, option_warnings = _strip_option_lines(new_code)
    # ProcStartLine = previous proc's End + 1, so it INCLUDES the blank
    # separator line(s) above this proc. When REPLACING, preserve that
    # separator (delete/insert below it) so we don't eat the blank line
    # between procs on every replace. A pure delete (new_code == '') removes
    # the whole range, separator included — that correctly closes the gap
    # (the following proc still owns its own leading blank).
    del_start, del_count = start, count
    if new_code:
        lead = 0
        for ln in cm.Lines(start, count).splitlines():
            if ln.strip() == "":
                lead += 1
            else:
                break
        if 0 < lead < count:
            del_start, del_count = start + lead, count - lead
    # Backup the portion we delete, for rollback on failure
    backup_code = cm.Lines(del_start, del_count)
    # Delete old procedure and insert new one with automatic rollback
    try:
        cm.DeleteLines(del_start, del_count)
        inserted = 0
        if new_code:
            normalized = new_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
            pre_insert_total = total - del_count
            cm.InsertLines(del_start, normalized)
            inserted = cm.CountOfLines - pre_insert_total
    except Exception:
        # Restore original code
        try:
            cm.InsertLines(del_start, backup_code)
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
        f"({del_count} deleted, {inserted} inserted) "
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
    atomic: bool = True, require_unique: bool = False,
    match_case: bool = False,
) -> str:
    """
    Applies surgical find/replace WITHIN a procedure without rewriting everything.
    patches: list of {find: str, replace: str}.
    More efficient than vbe_replace_proc when only a few lines change
    within a large proc (e.g.: 174 lines, only 15 change).

    proc_name="(Declarations)" targets the module's declarations section.

    atomic (default TRUE): if ANY patch fails, nothing is written at all.
    require_unique: reject a patch whose find text matches more than once.
    match_case (default FALSE): VBA is case-insensitive and the VBE rewrites
    identifier casing on its own, so anchors are matched case-insensitively
    unless this is set.
    """
    app = _Session.connect(db_path)

    _close_form_design_view(app, object_type, object_name)

    cache_key = f"{object_type}:{object_name}"
    _Session._cm_cache.pop(cache_key, None)

    cm = _get_code_module(app, object_type, object_name)

    is_declarations = _is_declarations(proc_name)
    kind = None
    if is_declarations:
        start = 1
        count = int(cm.CountOfDeclarationLines)
        if count == 0:
            raise RuntimeError(
                f"Module '{object_name}' has no declarations section (0 lines) — "
                "there is nothing to anchor a patch to. To CREATE one, use "
                "access_vbe_replace_lines with start_line=1, count=0 and the "
                "new declaration lines as new_code."
            )
    else:
        try:
            start, _body, count, kind = _proc_bounds(cm, proc_name)
        except Exception as exc:
            raise RuntimeError(
                f"Procedure '{proc_name}' not found in '{object_name}': {exc}"
            )
    total = cm.CountOfLines
    count = min(count, total - start + 1)

    # Get current proc code (cm.Lines(n, 0) raises in VBE)
    proc_code = cm.Lines(start, count) if count > 0 else ""
    backup_code = proc_code

    report = _apply_patches(
        proc_code, patches,
        match_case=match_case, require_unique=require_unique,
        proc_name=proc_name, base_line=start,
    )
    applied = report["applied"]
    blocking = report["not_found"] + report["unique_violations"]

    # Atomic gate — decided BEFORE any DeleteLines/InsertLines, so a rejected
    # batch leaves the module byte-for-byte identical.
    if atomic and blocking:
        msg = (
            f"ABORTED: atomic patch of '{proc_name}' — NOTHING WAS WRITTEN. "
            f"{len(blocking)} of {len(patches)} patch(es) failed and the module "
            f"is byte-for-byte unchanged.\n"
            f"Re-read the current code and re-send the ENTIRE batch — the "
            f"{applied} patch(es) that DID match were discarded too, so "
            f"re-sending only the failures would silently drop them.\n"
            + "\n".join(blocking)
        )
        if applied:
            msg += (
                "\nNote: failure context above reflects the in-memory text AFTER "
                "the matching patches in this batch were applied, so it may not "
                "line up with the module as it is on disk."
            )
        if report["case_notes"]:
            msg += "\n" + "\n".join(report["case_notes"])
        msg += "\nSet atomic=false to keep the old best-effort behaviour."
        return msg

    if applied == 0:
        return f"NOOP: no patches matched in '{proc_name}'. Errors:\n" + "\n".join(blocking)

    proc_code = report["code"]

    # Strip Option lines if the target is NOT at the top of the module.
    # The (Declarations) section is EXACTLY where Option lines belong, so never
    # strip there — do not rely on `start == 1` failing the `> 5` test.
    option_warnings = []
    if not is_declarations and start > 5:
        proc_code, option_warnings = _strip_option_lines(proc_code)

    # Replace entire proc with patched code
    inserted_text = ""
    try:
        if count > 0:
            cm.DeleteLines(start, count)
        if proc_code.strip():
            inserted_text = proc_code.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
            cm.InsertLines(start, inserted_text)
    except Exception:
        try:
            if backup_code:
                cm.InsertLines(start, backup_code)
        except Exception:
            pass
        raise

    # Persist VBE changes to .accdb — without this, patches to form/report
    # code-behind can be lost because the object's dirty flag is not set.
    _save_vbe_module(app, AC_TYPE.get(object_type, 5), object_name)
    new_total = cm.CountOfLines
    if is_declarations:
        # No proc to measure — ProcCountLines would raise and report a bogus 0.
        try:
            new_count = int(cm.CountOfDeclarationLines)
        except Exception:
            new_count = 0
    else:
        try:
            # Clamp like replace_proc does — ProcCountLines can inflate the last
            # proc's count past the end of the module.
            new_count = min(cm.ProcCountLines(proc_name, kind), new_total - start + 1)
        except Exception:
            new_count = 0
    # Health check — expected_total activates Check 3 (count sanity)
    expected_total = total - count + _vbe_line_count(inserted_text)
    health = _check_module_health(cm, cache_key, expected_total=expected_total)
    result = (
        f"OK: {applied}/{len(patches)} patches applied in '{proc_name}' "
        f"({count} → {new_count} lines) → module now has {new_total} lines"
    )
    if is_declarations and new_count < count:
        result += (
            "\nNote: the declarations section shrank — if the patched text "
            "introduced a Sub/Function line, VBE moved the boundary."
        )
    if report["fallback_notes"]:
        # Not "WS-fallback:" — this list now also carries case-insensitive
        # match notes, which have nothing to do with whitespace.
        result += f"\nMatch notes: {'; '.join(report['fallback_notes'])}"
    if report["ambiguous_notes"]:
        result += f"\nAmbiguous matches: {'; '.join(report['ambiguous_notes'])}"
    if report["case_notes"]:
        result += f"\n" + "\n".join(report["case_notes"])
    if option_warnings:
        result += f"\n" + "\n".join(option_warnings)
    if health:
        result += f"\n" + "\n".join(health)
    if blocking:
        result += f"\nNot found:\n" + "\n".join(blocking)
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
    _save_vbe_module(app, AC_TYPE.get(object_type, 5), object_name)
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
# Static syntax check — the SAFE alternative to access_compile_vba
# (requested by @TvanStiphout-Home, thanks!)
# ---------------------------------------------------------------------------
# access_compile_vba is unusable as a post-edit check: its step 0 shells out to
# MSACCESS.EXE /decompile and then either Quit(1) (= acQuitSaveNone) or closes
# the database, discarding unsaved VBA.  This tool touches the ALREADY OPEN
# project only: no decompile, no RunCommand, no Design view, no second Access
# instance.  It is a structural validator, not a compiler.

_CHECK_SYNTAX_NOTE = (
    "Structural validation only — unbalanced If/For/Do/While/Select/With/"
    "Type/Enum blocks, code outside a procedure, misplaced Option statements. "
    "It does NOT resolve identifiers, types or references, so ok=true does not "
    "mean the project compiles. Use access_compile_vba for a real compile "
    "(warning: it decompiles first and can discard unsaved VBA)."
)


def _check_one_module_syntax(module_name: str, lines: list) -> list[dict]:
    """Run both pure checkers over one module's lines. Returns error dicts."""
    from .compile import _check_blocks_in_module, _check_structure_in_module

    block_errors: list = []
    _check_blocks_in_module(module_name, lines, block_errors)
    out = [
        {"module": e["module"], "line": e["line"], "message": e["error"]}
        for e in block_errors
    ]
    struct_errors: list = []
    _check_structure_in_module(module_name, lines, struct_errors)
    # _check_structure_in_module emits "<module> line N: <text>" strings
    for msg in struct_errors:
        line_no = 0
        m = re.search(r"\bline (\d+):", msg)
        if m:
            line_no = int(m.group(1))
        out.append({"module": module_name, "line": line_no,
                    "message": msg.split(": ", 1)[-1]})
    return out


def ac_vbe_check_syntax(
    db_path: str, object_type: str = None, object_name: str = None
) -> dict:
    """
    Static structural check of the VBA project that is already open.

    Scope: one module when object_type/object_name are given, otherwise every
    standard module and form/report code-behind (VBComponent types 1 and 100).
    """
    app = _Session.connect(db_path)
    errors: list[dict] = []
    modules_checked: list[str] = []
    skipped: list[str] = []

    if object_type and object_name:
        _close_form_design_view(app, object_type, object_name)
        cm = _get_code_module(app, object_type, object_name)
        total = cm.CountOfLines
        # split("\n"), not splitlines() — the pure checkers were written against
        # the raw VBE text and their " _" continuation test sees the stray \r.
        code = cm.Lines(1, total) if total > 0 else ""
        errors.extend(_check_one_module_syntax(object_name, code.split("\n")))
        modules_checked.append(object_name)
    else:
        # _get_vb_project, NOT VBE.ActiveVBProject — the active project can be
        # acwzmain (the wizard library) after a decompile/compact.
        try:
            proj = _get_vb_project(app)
            components = list(proj.VBComponents)
        except Exception as exc:
            raise RuntimeError(
                f"Could not enumerate the VBA project: {exc}. The project may "
                "fail to load (broken reference) or VBA object-model access may "
                "be blocked in the Trust Center."
            )
        for comp in components:
            name = "<unknown>"
            try:
                name = comp.Name
                if comp.Type not in (1, 100):
                    continue
                cm = comp.CodeModule
                total = cm.CountOfLines
                if total == 0:
                    modules_checked.append(name)
                    continue
                code = cm.Lines(1, total)
                errors.extend(_check_one_module_syntax(name, code.split("\n")))
                modules_checked.append(name)
            except Exception as exc:
                # Never report a clean "0 errors" for a module we could not read
                skipped.append(f"{name}: {exc}")

    result = {
        "ok": not errors and not skipped,
        "errors": errors,
        "modules_checked": len(modules_checked),
        "note": _CHECK_SYNTAX_NOTE,
    }
    if skipped:
        result["skipped"] = skipped
        result["warning"] = (
            f"{len(skipped)} module(s) could not be read — the result is "
            "incomplete and ok=false reflects that, not a syntax error."
        )
    return result


# ---------------------------------------------------------------------------
# Find definition — "Go To Definition" for VBA symbols
# (requested by @TvanStiphout-Home, thanks!)
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
    errors: list[dict] = []
    skipped = 0

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
            except Exception as exc:
                # Surface inaccessible modules instead of silently reporting
                # "0 definitions" when the whole VBA project fails to load.
                skipped += 1
                if len(errors) < _SEARCH_ERROR_CAP:
                    errors.append({
                        "object": f"{obj_type}:{obj_name}",
                        "error": str(exc).splitlines()[0] if str(exc) else repr(exc),
                    })
                continue
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

    out: dict = {
        "symbol": symbol,
        "total": len(definitions),
        "definitions": definitions,
    }
    if skipped:
        out["objects_skipped"] = skipped
        out["errors"] = errors
        out["warning"] = (
            f"{skipped} object(s) had no accessible CodeModule — results may "
            "be incomplete. If ALL objects failed, the VBA project is likely "
            "not loading (broken reference / Trust Center)."
        )
    return out
