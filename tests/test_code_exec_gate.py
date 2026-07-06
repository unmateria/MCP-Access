"""
Pure-Python tests for the opt-in code-execution gate (v0.7.51).

The three tools that can run arbitrary VBA/Shell (`access_run_vba`,
`access_eval_vba`, `access_run_macro`) are closed by default and only enabled
by the `MCP_ACCESS_ALLOW_CODE_EXEC` environment variable. The gate is enforced
in two layers: `list_tools()` hides them (hygiene) and `call_tool_sync` rejects
a direct call before touching COM (the real barrier).

No COM / no Access needed: the dispatch rejection returns before any `_Session`.

Run with:
    python -m pytest tests/test_code_exec_gate.py
"""

import asyncio
import json
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_access.security import (  # noqa: E402
    CODE_EXEC_TOOLS,
    code_exec_enabled,
    code_exec_denied_message,
)
from mcp_access.server import list_tools  # noqa: E402
from mcp_access.dispatcher import call_tool_sync  # noqa: E402


GATED = sorted(CODE_EXEC_TOOLS)


@pytest.mark.parametrize("value", ["1", "true", "TRUE", "Yes", "on", " on ", "On"])
def test_enabled_truthy(monkeypatch, value):
    monkeypatch.setenv("MCP_ACCESS_ALLOW_CODE_EXEC", value)
    assert code_exec_enabled() is True


@pytest.mark.parametrize("value", ["", "0", "false", "no", "off", "  ", "2", "disable"])
def test_disabled_falsy(monkeypatch, value):
    monkeypatch.setenv("MCP_ACCESS_ALLOW_CODE_EXEC", value)
    assert code_exec_enabled() is False


def test_disabled_when_unset(monkeypatch):
    monkeypatch.delenv("MCP_ACCESS_ALLOW_CODE_EXEC", raising=False)
    assert code_exec_enabled() is False


def test_list_tools_hides_gated_when_closed(monkeypatch):
    monkeypatch.delenv("MCP_ACCESS_ALLOW_CODE_EXEC", raising=False)
    names = {t.name for t in asyncio.run(list_tools())}
    assert names.isdisjoint(CODE_EXEC_TOOLS)


def test_list_tools_shows_gated_when_open(monkeypatch):
    monkeypatch.setenv("MCP_ACCESS_ALLOW_CODE_EXEC", "1")
    names = {t.name for t in asyncio.run(list_tools())}
    assert CODE_EXEC_TOOLS.issubset(names)


@pytest.mark.parametrize("name", GATED)
def test_dispatch_rejects_gated_when_closed(monkeypatch, name):
    monkeypatch.delenv("MCP_ACCESS_ALLOW_CODE_EXEC", raising=False)
    # Would need Access/COM if it got past the gate; it must not.
    text = call_tool_sync(name, {"db_path": "x"})
    payload = json.loads(text)
    assert "error" in payload
    assert "MCP_ACCESS_ALLOW_CODE_EXEC" in payload["hint"]
    assert payload["gated_tool"] == name


def test_denied_message_shape():
    msg = code_exec_denied_message("access_run_vba")
    assert msg["gated_tool"] == "access_run_vba"
    assert "MCP_ACCESS_ALLOW_CODE_EXEC" in msg["hint"]
    # Not a self-executable instruction: it never tells the model to call a tool.
    assert "call " not in msg["hint"].lower()


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
