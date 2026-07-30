"""
Pure-Python tests for the opt-out SHIFT AutoExec bypass.

Based on the tests contributed by @CustomDataNZ in PR #34.

Holding SHIFT across OpenCurrentDatabase / MSACCESS /decompile is how the server
skips a target database's AutoExec macro and startup form. `keybd_event` is a
GLOBAL OS-level key-down, though — it is not scoped to Access, so every keystroke
the human types anywhere on the machine during the hold arrives shifted. The open
path holds it ~0.3s, the decompile path ~3s.

It stays ON by default — turning it off silently would change behaviour for every
existing user whose databases rely on it, with no error to explain why AutoExec
suddenly runs. `MCP_ACCESS_SHIFT_BYPASS=0` opts out. Unlike
MCP_ACCESS_ALLOW_CODE_EXEC this is an ergonomics switch, not a security gate, so
it fails OPEN rather than closed — which is exactly why the name has no `ALLOW_`
prefix.

No COM / no Access needed.

Run with:
    python -m pytest tests/test_shift_bypass_gate.py
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from mcp_access.security import shift_bypass_enabled  # noqa: E402

ENV = "MCP_ACCESS_SHIFT_BYPASS"
PKG = Path(__file__).resolve().parent.parent / "mcp_access"


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
    """Import order must not decide the answer — the code-exec gate has the same
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


def test_no_allow_prefixed_alias_is_honoured(monkeypatch):
    """`ALLOW_` would imply default-off, which this switch is not. A stale
    MCP_ACCESS_ALLOW_SHIFT_BYPASS carried over from notes must not appear to
    work — silently ignoring it is better than silently obeying it."""
    monkeypatch.delenv(ENV, raising=False)
    monkeypatch.setenv("MCP_ACCESS_ALLOW_SHIFT_BYPASS", "0")
    assert shift_bypass_enabled() is True


def test_press_returns_false_when_disabled(monkeypatch):
    """The press helper is the single gate point: with the switch off it must
    report 'not held' WITHOUT touching the keyboard, so callers skip the
    matching release too."""
    import mcp_access.core as core

    pressed = []
    monkeypatch.setenv(ENV, "0")
    monkeypatch.setattr(core.ctypes.windll.user32, "keybd_event",
                        lambda *a: pressed.append(a), raising=False)
    assert core._press_shift_bypass("test") is False
    assert pressed == []


def test_shift_synthesis_lives_only_in_core(monkeypatch):
    """Structural guard against the gate being half-applied again.

    Before v0.7.53 the press/sleep/release sequence was copy-pasted into three
    call sites, so gating it meant remembering all three. Now `core.py` owns the
    only two functions that touch the SHIFT key — a future edit that reopens a
    second synthesis site fails here rather than silently shifting the user's
    typing with the switch turned off.
    """
    for module in ("maintenance.py", "vbe.py", "vba_exec.py", "compile.py"):
        src = (PKG / module).read_text(encoding="utf-8")
        assert "keybd_event" not in src, (
            f"{module} synthesises a key event directly — route the SHIFT "
            f"bypass through core._press_shift_bypass/_release_shift instead"
        )

    core_src = (PKG / "core.py").read_text(encoding="utf-8")
    # Two call sites only: the press (gated) and the release (idempotent).
    assert core_src.count("keybd_event(_VK_SHIFT") == 2
    press = core_src.split("def _press_shift_bypass")[1]
    body = press.split("\ndef ")[0]
    assert "shift_bypass_enabled()" in body, (
        "core._press_shift_bypass no longer consults the gate"
    )
