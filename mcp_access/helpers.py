"""
Shared helpers: temp file I/O, binary section handling, value coercion, text matching.
"""

import os
import re
import tempfile
from datetime import datetime
from typing import Any

from .core import AC_TYPE, _Session, _parsed_controls_cache, log
from .constants import BINARY_SECTIONS


# ---------------------------------------------------------------------------
# Temp file I/O
# ---------------------------------------------------------------------------

def read_tmp(path: str) -> tuple[str, str]:
    """
    Reads a file exported by Access.
    Returns (content, encoding_used).
    Detects UTF-16 with BOM before trying utf-8/cp1252.

    UTF-8 is tried before cp1252 because cp1252 is single-byte and
    almost never raises UnicodeDecodeError, which means files that
    are actually UTF-8 would be silently mis-decoded as cp1252 with
    garbled multi-byte sequences (mojibake).
    """
    with open(path, "rb") as f:
        bom = f.read(2)
    if bom in (b"\xff\xfe", b"\xfe\xff"):
        with open(path, encoding="utf-16") as f:
            return f.read(), "utf-16"
    for enc in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, encoding=enc) as f:
                return f.read(), enc
        except UnicodeDecodeError:
            continue
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read(), "utf-8"


def write_tmp(path: str, content: str, encoding: str = "utf-16") -> None:
    """
    Writes content for Access to read with LoadFromText.
    Default utf-16 (Access .accdb expects UTF-16LE with BOM).

    Uses errors="strict" so that non-encodable characters (e.g. emoji or
    Asian characters in a cp1252-targeted .bas module) surface as a clear
    UnicodeEncodeError instead of being silently replaced with `?` —
    silent corruption used to be possible because we wrote with
    errors="replace".
    """
    try:
        with open(path, "w", encoding=encoding, errors="strict") as f:
            f.write(content)
    except UnicodeEncodeError as e:
        # Surface a single, actionable message: what failed, where, and
        # the offending substring.  The dispatcher wraps this into the
        # tool error and the caller can decide to recode in UTF-16.
        snippet = content[max(0, e.start - 20):e.start] + \
                  "[" + content[e.start:e.end] + "]" + \
                  content[e.end:min(len(content), e.end + 20)]
        raise UnicodeEncodeError(
            e.encoding, e.object, e.start, e.end,
            f"cannot encode character(s) at offset {e.start}-{e.end} "
            f"with {encoding!r}: ...{snippet!r}... "
            f"(VBA modules must be in the system ANSI codepage; "
            f"use only ASCII or codepage-compatible characters)"
        ) from None


# ---------------------------------------------------------------------------
# Binary section handling (forms/reports)
# ---------------------------------------------------------------------------

def strip_binary_sections(text: str) -> str:
    """
    Strips binary sections from an Access form/report export.
    Reduces size ~20x without affecting VBA or controls.
    """
    lines = text.splitlines(keepends=True)
    result: list[str] = []
    skip_depth = 0
    skip_indent = ""

    for line in lines:
        rstripped = line.rstrip("\r\n")
        stripped = rstripped.lstrip()
        indent = rstripped[: len(rstripped) - len(stripped)]

        if skip_depth > 0:
            if stripped == "End" and indent == skip_indent:
                skip_depth -= 1
            continue

        if re.match(r"^Checksum\s*=\s*", rstripped):
            continue

        m = re.match(r"^(\s*)(\w+)\s*=\s*Begin\s*$", rstripped)
        if m and m.group(2) in BINARY_SECTIONS:
            skip_indent = m.group(1)
            skip_depth = 1
            continue

        result.append(line)

    return "".join(result)


def extract_binary_blocks(text: str) -> dict[str, str]:
    """
    Extracts binary Begin...End blocks from the original export.
    Returns {section_name: full_block_text}.
    """
    blocks: dict[str, str] = {}
    lines = text.splitlines(keepends=True)
    i = 0
    while i < len(lines):
        line = lines[i]
        rstripped = line.rstrip("\r\n")
        stripped = rstripped.lstrip()
        indent = rstripped[: len(rstripped) - len(stripped)]

        m = re.match(r"^(\s*)(\w+)\s*=\s*Begin\s*$", rstripped)
        if m and m.group(2) in BINARY_SECTIONS:
            section = m.group(2)
            block_lines = [line]
            j = i + 1
            while j < len(lines):
                bl = lines[j]
                bl_r = bl.rstrip("\r\n")
                bl_s = bl_r.lstrip()
                bl_indent = bl_r[: len(bl_r) - len(bl_s)]
                block_lines.append(bl)
                if bl_s == "End" and bl_indent == indent:
                    break
                j += 1
            blocks[section] = "".join(block_lines)
            i = j + 1
            continue

        i += 1

    return blocks


def restore_binary_sections(app: Any, object_type: str, name: str, new_code: str) -> str:
    """
    Re-injects binary sections from the current export of the object.
    """
    fd, tmp = tempfile.mkstemp(suffix=".txt", prefix="access_mcp_orig_")
    os.close(fd)
    try:
        try:
            app.SaveAsText(AC_TYPE[object_type], name, tmp)
        except Exception:
            log.info("restore_binary_sections: '%s' does not exist yet", name)
            return new_code
        original, _enc = read_tmp(tmp)
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass

    blocks = extract_binary_blocks(original)
    if not blocks:
        return new_code

    # Inject binary blocks before the OUTER closing `End` of the top-level
    # Form/Report.  Subforms have their own nested Begin Form...End that
    # would confuse a simple "first End after Begin Form" approach, so we
    # track full block depth (Begin <Type> AND `prop = Begin` multi-line
    # values like `NameMap = Begin`).  When depth returns to the level the
    # outer Form was opened at, that End is the one we want.
    _top_begin_re = re.compile(r"^Begin\s+(?:Form|Report)\s*$", re.IGNORECASE)
    _begin_any_re = re.compile(r"^Begin\b")
    _begin_prop_re = re.compile(r"^\w+\s*=\s*Begin\s*$")
    lines = new_code.splitlines(keepends=True)
    result: list[str] = []
    depth = 0
    top_depth = -1   # depth at which the outermost Form/Report was opened
    injected = False

    for line in lines:
        stripped = line.strip()

        if not injected and top_depth == -1 and _top_begin_re.match(stripped):
            top_depth = depth
            depth += 1
            result.append(line)
            continue
        if _begin_any_re.match(stripped) or _begin_prop_re.match(stripped):
            depth += 1
            result.append(line)
            continue
        if stripped == "End":
            depth -= 1
            # The End that brings us back to top_depth closes the outermost
            # Form/Report — inject the binary blocks right before it.
            if not injected and top_depth != -1 and depth == top_depth:
                for block_text in blocks.values():
                    result.append(block_text)
                    if not block_text.endswith("\n"):
                        result.append("\n")
                injected = True
            result.append(line)
            continue

        result.append(line)

    return "".join(result)


# ---------------------------------------------------------------------------
# Value coercion and text matching
# ---------------------------------------------------------------------------

def coerce_prop(value: Any) -> Any:
    """Converts strings to int/bool as appropriate for COM properties."""
    if isinstance(value, (int, float, bool)):
        return value
    if isinstance(value, str):
        low = value.lower()
        if low in ("true", "yes", "-1"):
            return True
        if low in ("false", "no", "0"):
            return False
        try:
            return int(value)
        except ValueError:
            pass
        try:
            return float(value)
        except ValueError:
            pass
    return value


def text_matches(needle: str, haystack: str, match_case: bool, use_regex: bool) -> bool:
    """Matches needle against haystack: plain substring or regex."""
    if use_regex:
        flags = 0 if match_case else re.IGNORECASE
        return re.search(needle, haystack, flags) is not None
    if not match_case:
        return needle.lower() in haystack.lower()
    return needle in haystack


def serialize_value(val: Any) -> Any:
    """Converts non-serializable COM types to JSON-safe values."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.isoformat()
    try:
        from decimal import Decimal
        if isinstance(val, Decimal):
            return float(val)
    except ImportError:
        pass
    if isinstance(val, (bytes, memoryview)):
        return f"<binary {len(val)} bytes>"
    return val


# ---------------------------------------------------------------------------
# Form/report code-behind splitter (shared by code.py + controls.py)
# ---------------------------------------------------------------------------

def split_code_behind(code: str) -> tuple[str, str]:
    """
    Splits a form/report text into (form_text, vba_code).
    If the code contains 'CodeBehindForm' or 'CodeBehindReport', it splits it.
    Returns (form_text_without_vba, vba_code) where vba_code may be empty.
    The form_text is cleaned of HasModule if there is VBA (it will be injected later).

    Matches only on its own line — Access export emits the marker as a
    bare header, so this guards against false positives where the literal
    string happens to appear inside a property value (e.g. a Caption).
    """
    for marker in ("CodeBehindForm", "CodeBehindReport"):
        m = re.search(r"(?m)^\s*" + re.escape(marker) + r"\s*$", code)
        if m:
            idx = m.start()
            form_part = code[:idx].rstrip() + "\n"
            vba_part = code[idx:].split("\n", 1)
            vba_code = vba_part[1] if len(vba_part) > 1 else ""
            vba_lines = []
            for line in vba_code.splitlines():
                stripped = line.strip()
                if stripped.startswith("Attribute VB_"):
                    continue
                vba_lines.append(line)
            vba_code = "\n".join(vba_lines).strip()
            return form_part, vba_code
    return code, ""
