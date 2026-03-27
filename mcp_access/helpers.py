"""
Shared helpers: temp file I/O, binary section handling, value coercion, text matching.
"""

import os
import re
import tempfile
from datetime import datetime
from typing import Any

from .core import AC_TYPE, _Session, _vbe_code_cache, _parsed_controls_cache, log
from .constants import BINARY_SECTIONS


# ---------------------------------------------------------------------------
# Temp file I/O
# ---------------------------------------------------------------------------

def read_tmp(path: str) -> tuple[str, str]:
    """
    Reads a file exported by Access.
    Returns (content, encoding_used).
    Detects UTF-16 with BOM before trying cp1252.
    """
    with open(path, "rb") as f:
        bom = f.read(2)
    if bom in (b"\xff\xfe", b"\xfe\xff"):
        with open(path, encoding="utf-16") as f:
            return f.read(), "utf-16"
    for enc in ("utf-8-sig", "cp1252", "utf-8"):
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
    """
    with open(path, "w", encoding=encoding, errors="replace") as f:
        f.write(content)


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

    _end_re = re.compile(r"^\s*End\s+(?:Form|Report)\s*$")
    _begin_re = re.compile(r"^\s*Begin\s+(?:Form|Report)\s*$")
    lines = new_code.splitlines(keepends=True)
    result: list[str] = []
    in_top_form = False
    injected = False

    for line in lines:
        stripped = line.strip()

        if _begin_re.match(stripped):
            in_top_form = True

        if in_top_form and not injected and _end_re.match(stripped):
            for block_text in blocks.values():
                result.append(block_text)
                if not block_text.endswith("\n"):
                    result.append("\n")
            injected = True
            in_top_form = False

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
    if isinstance(val, bytes):
        return f"<binary {len(val)} bytes>"
    return val
