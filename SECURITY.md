# Security policy

## Transport and trust boundary

`mcp-access` is a **local stdio MCP server**. It does not open a socket, bind a
port, or listen on the network — there is no remote surface. The entry point
(`server.main()`) speaks the MCP protocol over stdin/stdout only.

Because of that, **there is no login by design** — this is explicit, not a gap.
Whoever controls the server's stdin already has local code execution with the
privileges of the process. Adding a password would protect nothing: there is no
network peer to authenticate, and the trust boundary is the operator/agent that
launches the process. The server runs with the launching user's privileges.

## Primary risk: prompt injection / confused deputy

The realistic threat is **not** an unauthenticated remote attacker. It is
**prompt injection**: the agent driving this server ingests instructions that
were smuggled into its context and then calls a code-execution tool on the
attacker's behalf (a *confused deputy*). Injected text can arrive via:

- a crafted `db_path` argument (patched — see `_sanitize_db_path`,
  GHSA-9jp6-hph9-jm5f), or
- **data the agent reads out of the database itself**: VBA source, object names,
  or table contents. A database is untrusted input.

The code-execution sinks are `access_run_vba`, `access_eval_vba` and
`access_run_macro` (`mcp_access/vba_exec.py`). They run arbitrary VBA, which can
call `Shell "cmd /c ..."` → arbitrary OS command execution.

### Why `confirm_*` flags are not an injection defense

The existing `confirm_destructive` / `confirm=true` flags protect against
**model mistakes**, not injection: injected text can just ask for `confirm=true`
too. The only control that survives injection is one the model cannot set from
inside the session — an **environment variable** the human operator puts on the
process. That is the gate below.

## Mitigations present

| Mitigation | Protects against |
|------------|------------------|
| `_sanitize_db_path` (GHSA-9jp6-hph9-jm5f) | prompt injection via `db_path` |
| `MCP_ACCESS_ALLOW_CODE_EXEC` gate (default **closed**) | injection → arbitrary code execution |
| `confirm_*` flags on destructive SQL/deletes | model mistakes (not injection) |
| `AutomationSecurity = 3` + SHIFT AutoExec bypass | malicious startup macros on DB open |

## Code-execution gate: `MCP_ACCESS_ALLOW_CODE_EXEC`

The three code-execution tools are **disabled by default**. A fresh install
straight from PyPI cannot be turned into RCE by a single injection: the sinks
are not advertised and are rejected at dispatch before any COM call.

`access_run_macro` is gated too, because a macro can carry a `RunCode` action
that runs VBA/Shell.

The gate is enforced in two layers:

1. **Not advertised** — `list_tools()` omits the three tools when the gate is
   closed, so the model never sees them.
2. **Rejected at dispatch** (the real barrier) — `call_tool_sync` refuses a
   gated tool *before* touching COM, even if a client calls the name directly
   without ever seeing it advertised.

### How to enable

Add the variable to this server's `env` block in your MCP client config (e.g.
the repo's `.mcp.json`) and **restart the server**:

```json
{
  "mcpServers": {
    "access": {
      "command": "...",
      "args": ["..."],
      "env": { "MCP_ACCESS_ALLOW_CODE_EXEC": "1" }
    }
  }
}
```

Accepted truthy values (case-insensitive): `1`, `true`, `yes`, `on`.

**Enabling grants arbitrary OS command execution** through the VBA `Shell`
function. Only point the server at databases you trust, and treat all database
content as untrusted input regardless.

### Enabling on request (edit config + restart)

When **the user explicitly asks** to enable VBA execution, the assistant may
edit the client config to add `"MCP_ACCESS_ALLOW_CODE_EXEC": "1"`, after warning
what it grants. It takes effect **only after a restart**, because the gate is
read at startup. This is deliberate: the gate stays out of band, so an injection
cannot escalate a running session.

There is **no** MCP tool that turns the gate on at runtime, and there never will
be — an injection could call it. Enabling always requires the out-of-band action
plus a restart.

## Non-goals of this iteration

- A mode system (`MCP_ACCESS_MODE` readonly/safe/full) gating whole tool families.
- Gating `access_ui_type` / `access_ui_click` (they can type/click outside the
  Access window) — possible future work.
- Heuristic `Shell()` detection in incoming VBA — weak and easily obfuscated, no
  real guarantee.

## Reporting a vulnerability

Please report vulnerabilities privately via the repository's **GitHub Security
Advisories** ("Report a vulnerability"), not as a public issue.
