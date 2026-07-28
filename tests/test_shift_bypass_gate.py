"""
Pure-Python tests for the opt-in SHIFT AutoExec bypass.

Holding SHIFT across OpenCurrentDatabase / MSACCESS /decompile is how the server
skips a target database's AutoExec macro and startup form. `keybd_event` is a
GLOBAL OS-level key-down, though - it is not scoped to Access, so every keystroke
the human types anywhere on the machine during the hold arrives shifted. The open
path holds it ~0.3s, the decompile path ~3s.

It stays ON by default - turning it off silently would change behaviour for every
existing user whose databases rely on it, with no error to explain why AutoExec
suddenly runs. `MCP_ACCESS_SHIFT_BYPASS=0` opts out. Unlike
MCP_ACCESS_ALLOW_CODE_EXEC this is an ergonomics switch, not a security gate, so
it fails OPEN rather than closed - which is exactly why the name has no `ALLOW_`
prefix.

No COM / no Access needed.

Run with:
    python -m pytest tests/test_shift_bypass_gate.py
"""

import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from mcp_access.security import shift_bypass_enabled  # noqa: E402

ENV = "MCP_ACCESS_SHIFT_BYPASS"


def test_on_when_unset(monkeypatch):
    """Back-compat is the whole point of the default: an existing user who
    upgrades and sets nothing must see no change in behaviour."""
    monkeypatch.delenv(ENV, raising=False)
    assert shift_bypass_enabled() is True


@pytest.mark.parametrize("value", ["0", "false", "FALSE", "no", "NO", "off", " off "])
def test_falsy_values_disable(monkeypatch, value):
    monkeypatch.setenv(ENV, value)
    assert shift_bypass_enabled() is False


@pytest.mark.parametrize("value", ["", "   ", "1", "true", "yes", "on", "maybe", "2"])
def test_everything_else_stays_on(monkeypatch, value):
    """Fails OPEN: only an explicit falsy word disables it. A typo or an empty
    value must not quietly drop the bypass and let AutoExec run."""
    monkeypatch.setenv(ENV, value)
    assert shift_bypass_enabled() is True


def test_read_per_call_not_at_import(monkeypatch):
    """Import order must not decide the answer - the code-exec gate has the same
    property and the tests rely on it."""
    monkeypatch.delenv(ENV, raising=False)
    assert shift_bypass_enabled() is True
    monkeypatch.setenv(ENV, "0")
    assert shift_bypass_enabled() is False
    monkeypatch.delenv(ENV, raising=False)
    assert shift_bypass_enabled() is True


def test_switch_is_independent_of_the_code_exec_gate(monkeypatch):
    """Two separate concerns, opposite defaults: turning the keyboard bypass off
    must not affect the security gate, or vice versa."""
    monkeypatch.setenv("MCP_ACCESS_ALLOW_CODE_EXEC", "1")
    monkeypatch.setenv(ENV, "0")
    assert shift_bypass_enabled() is False
    monkeypatch.delenv("MCP_ACCESS_ALLOW_CODE_EXEC", raising=False)
    assert shift_bypass_enabled() is False


def test_old_allow_prefixed_name_is_not_honoured(monkeypatch):
    """The switch was briefly named MCP_ACCESS_ALLOW_SHIFT_BYPASS with the
    opposite default. If someone carries that stale var over from notes or an
    old config, it must not appear to work."""
    monkeypatch.delenv(ENV, raising=False)
    monkeypatch.setenv("MCP_ACCESS_ALLOW_SHIFT_BYPASS", "0")
    assert shift_bypass_enabled() is True


@pytest.mark.parametrize("module", ["core", "maintenance"])
def test_every_keybd_event_site_is_gated(module):
    """Guards against a future edit reintroducing an ungated press.

    Every `keybd_event(VK_SHIFT, 0, 0, 0)` key-DOWN must sit under a
    shift_bypass_enabled() check. Key-UP calls (KEYEVENTF_KEYUP) are exempt -
    releasing a key that was never pressed is harmless and the safety net on
    exit depends on staying unconditional.
    """
    src = (Path(__file__).resolve().parent.parent / "mcp_access" / f"{module}.py").read_text(
        encoding="utf-8"
    )
    assert "shift_bypass_enabled" in src, f"{module}.py does not consult the gate at all"

    lines = src.splitlines()
    for i, line in enumerate(lines):
        stripped = line.strip()
        if "_kbd(VK_SHIFT, 0, 0, 0)" not in stripped:
            continue
        # Walk back to the nearest enclosing gate; the press sits inside
        # `if shift_bypass_enabled():` plus a try block, so allow a few levels.
        window = "\n".join(lines[max(0, i - 8):i])
        assert "shift_bypass_enabled()" in window, (
            f"{module}.py line {i + 1}: SHIFT key-down is not gated by "
            f"shift_bypass_enabled()"
        )
