"""
Environment switches: the code-execution gate and the SHIFT-bypass opt-out.

Both are read from the environment on every call so import order is irrelevant
and tests can flip them, but they are opposites by design — see
``shift_bypass_enabled`` for why one fails closed and the other fails open.

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
_FALSY = frozenset({"0", "false", "no", "off"})


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


def shift_bypass_enabled() -> bool:
    """True unless the operator has switched the synthetic SHIFT bypass off.

    Holding SHIFT across ``OpenCurrentDatabase`` / ``MSACCESS /decompile`` is how
    this server skips a target database's AutoExec macro and startup form. It
    works, but ``keybd_event(VK_SHIFT, ...)`` is a **global** OS-level key-down:
    it is not scoped to Access, so every keystroke the human types anywhere on
    the machine during the hold arrives shifted. The open path holds it ~0.3 s on
    every database switch; the decompile path holds it ~3 s. On a box where
    someone is working while the server runs, that is a repeated nuisance with
    nothing on screen to explain it (reported by @CustomDataNZ).

    **Default ON** — unlike ``MCP_ACCESS_ALLOW_CODE_EXEC``, which is a security
    gate and fails closed. This one is ergonomics, and defaulting it off would
    silently change behaviour for every existing user whose databases rely on the
    bypass: their AutoExec would start running again with no error to explain it.
    So it stays on and the people who don't need it turn it off. Hence the name —
    an ``ALLOW_`` prefix would wrongly imply default-off, and a ``DISABLE_`` flag
    would force every reader through a double negative.

    Set ``MCP_ACCESS_SHIFT_BYPASS`` to ``0`` / ``false`` / ``no`` / ``off`` to
    disable. It fails OPEN: anything else (including a typo or an empty value)
    keeps the bypass, because quietly dropping it would let AutoExec run on
    someone's database. With it disabled:

    - ``AutomationSecurity = msoAutomationSecurityForceDisable`` still runs, which
      blocks VBA auto-run code but NOT an AutoExec *macro object* (tested —
      Access ignores it for those), so an unguarded AutoExec macro WILL execute;
    - the dialog watchdog still runs, so a modal raised by that startup code is
      still detected and dismissed.

    Turn it off when the target databases guard their own startup, which is the
    clean fix and belongs there rather than in a global input hack::

        If Not Application.UserControl Then
          Exit Function
        End If

    ``Application.UserControl`` is False when Access was started via COM and True
    when a human launched it, so the database opts itself out under automation
    and needs no bypass at all.
    """
    raw = os.environ.get("MCP_ACCESS_SHIFT_BYPASS")
    if raw is None:
        return True
    return raw.strip().lower() not in _FALSY
