"""
Environment switches: the code-execution gate, the SHIFT-bypass opt-out and the
exclusive-open switch.

All three are read from the environment on every call so import order is
irrelevant and tests can flip them, but they are not all alike — see
``shift_bypass_enabled`` for why one fails open while the other two fail closed.

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


def exclusive_open_enabled() -> bool:
    """True when the operator wants the session's database opened exclusively.

    ``OpenCurrentDatabase(filepath, Exclusive, bstrPassword)`` takes the mode as
    its second argument and defaults to shared. Shared is the right default for
    reading and for most edits, but it cannot take a **design lock**: attaching
    Data Macros through ``SaveAsText``/``LoadFromText``, or anything that routes
    through ``DoCmd.OpenTable acViewDesign``, needs one. When another Access
    session has that table open the lock is refused, the change is dropped for
    that table alone, and the run can still finish and report success. An
    exclusive open turns that into one visible failure at open time instead
    (requested by @GPGeorge, issue #36).

    There is no explicit open tool — the open is implicit in
    ``_Session.connect()`` — so a per-call argument would have to be threaded
    through every tool that takes ``db_path`` and kept consistent across all of
    them. A session either wants exclusive or it doesn't, so it is one
    server-level switch: ``MCP_ACCESS_EXCLUSIVE`` set to ``1``/``true``/``yes``/
    ``on``.

    **Default OFF, and it fails CLOSED** like ``MCP_ACCESS_ALLOW_CODE_EXEC``
    (and unlike ``MCP_ACCESS_SHIFT_BYPASS``, which is ergonomics and fails
    open). Not because exclusivity is dangerous to the file — it isn't — but
    because it is *exclusive*: the COM session holds the database open between
    tool calls, so while the server is connected nobody else can get in. On a
    shared front-end that locks out every other user for as long as the session
    lives. Falling closed on a typo costs a design lock; falling open on one
    would throw a workgroup off their database with nothing on screen to say
    why. ``access_close`` releases it without stopping the server.

    Turning it on also stops the server attaching to an already-running Access
    instance (``_Session._launch``): that instance holds the file shared, and
    reusing it would report an exclusive session that is nothing of the sort —
    the same silent-success problem one level up.
    """
    return os.environ.get("MCP_ACCESS_EXCLUSIVE", "").strip().lower() in _TRUTHY
