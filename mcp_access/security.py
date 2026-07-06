"""
Opt-in capability gate for code-execution tools.

Single source of truth for the ``MCP_ACCESS_ALLOW_CODE_EXEC`` gate. The three
tools that can run arbitrary VBA (and therefore ``Shell "cmd /c ..."`` -> RCE)
are closed by default and only enabled when a human operator sets the
environment variable on the server process. This is the only control that
survives prompt injection: injected text can ask for ``confirm=true`` but it
cannot set an env var on the already-running process. See SECURITY.md.
"""

import os

# Tools that can execute arbitrary code. ``access_run_macro`` is included
# because a macro can carry a RunCode action that executes VBA/Shell.
CODE_EXEC_TOOLS = frozenset({
    "access_run_vba",
    "access_eval_vba",
    "access_run_macro",
})

_TRUTHY = frozenset({"1", "true", "yes", "on"})


def code_exec_enabled() -> bool:
    """True when the operator has opted into code execution.

    Read on every call (not at import time) so import order is irrelevant and
    tests can flip it with ``monkeypatch.setenv``/``delenv``.
    """
    return os.environ.get("MCP_ACCESS_ALLOW_CODE_EXEC", "").strip().lower() in _TRUTHY


def code_exec_denied_message(name: str) -> dict:
    """Standard error dict returned when a gated tool is called while closed.

    The hint describes the *user-initiated* enabling path (edit config + restart)
    and is deliberately NOT phrased as a self-executable instruction: enabling
    requires an out-of-band human action, so an in-session injection cannot use
    this message to escalate.
    """
    return {
        "error": f"'{name}' is disabled: code execution is turned off by default.",
        "hint": (
            "Code execution (arbitrary VBA/Shell) is disabled by default. If the "
            "user asks to enable it, add `MCP_ACCESS_ALLOW_CODE_EXEC=1` to this "
            "server's `env` in the MCP client config and restart the server "
            "(takes effect on restart only). See SECURITY.md."
        ),
        "gated_tool": name,
    }
