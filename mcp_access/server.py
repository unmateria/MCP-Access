"""
MCP Server setup: list_tools, list_prompts, get_prompt, call_tool, main.
"""

import asyncio
import json
import re
import time
import traceback

import mcp.types as types
from mcp.server import Server
from mcp.server.stdio import stdio_server

from .core import _Session, _com_executor, log
from .tools import TOOLS, coerce_arguments
from .security import code_exec_enabled, CODE_EXEC_TOOLS


server = Server("access-mcp")


@server.list_tools()
async def list_tools() -> list[types.Tool]:
    # Hide code-execution tools unless the operator opted in. This is hygiene
    # only (the model never sees them); the real barrier is the dispatch-time
    # check in dispatcher.call_tool_sync, which rejects a direct call too.
    if code_exec_enabled():
        return TOOLS
    return [t for t in TOOLS if t.name not in CODE_EXEC_TOOLS]


@server.list_prompts()
async def list_prompts() -> list[types.Prompt]:
    return [
        types.Prompt(
            name="access-workflow",
            description=(
                "Usage instructions for the MCP access server for working with "
                "Microsoft Access databases (.accdb/.mdb) from Claude Code."
            ),
            arguments=[
                types.PromptArgument(
                    name="db_path",
                    description="Full path to the .accdb or .mdb file",
                    required=False,
                )
            ],
        )
    ]


# A legitimate Windows Access path never contains newlines or control
# characters, so any that appear are an attempt to inject extra prompt
# instructions (GHSA-9jp6-hph9-jm5f). Collapse to a single sanitized line and
# cap the length before reflecting it into the prompt template.
_CONTROL_CHARS = re.compile(r"[\x00-\x1f\x7f]")
_MAX_PATH_LEN = 260  # Windows MAX_PATH


def _sanitize_db_path(raw: object) -> str:
    """Reduce an untrusted db_path to a single harmless line of path text."""
    if not isinstance(raw, str) or not raw.strip():
        return "<path_to_file.accdb>"
    # Drop everything from the first control character onward: this removes the
    # newline the injection relies on and anything smuggled after it.
    cleaned = _CONTROL_CHARS.split(raw, 1)[0].strip()
    if not cleaned:
        return "<path_to_file.accdb>"
    if len(cleaned) > _MAX_PATH_LEN:
        cleaned = cleaned[:_MAX_PATH_LEN] + "..."
    return cleaned


@server.get_prompt()
async def get_prompt(name: str, arguments: dict | None) -> types.GetPromptResult:
    db_path = _sanitize_db_path((arguments or {}).get("db_path"))
    return types.GetPromptResult(
        description="Required workflow for working with Access databases",
        messages=[
            types.PromptMessage(
                role="user",
                content=types.TextContent(
                    type="text",
                    text=f"""
I'm working with a Microsoft Access database (file path): `{db_path}`

REQUIRED RULES for the agent:
1. Any operation on .accdb or .mdb files MUST be done through the MCP access server.
   No other tool or shell command can read or modify Access.

2. Required workflow for editing VBA or object definitions:
   a) access_list_objects  -> discover which objects exist (forms, modules, reports...)
   b) access_get_code      -> read the current code of the object
   c) modify the text
   d) access_set_code      -> save the result to the database

3. For small edits (more efficient):
   a) access_vbe_module_info  -> procedure index with line numbers
   b) access_vbe_get_proc     -> code of the specific procedure
   c) access_vbe_replace_lines -> replace only the modified lines

4. Never guess form, module, or control names.
   Always call access_list_objects or access_list_controls first.

5. Never write VBA code without first reading the original with access_get_code
   or access_vbe_get_proc. The internal Access format is strict.
""",
                ),
            )
        ],
    )


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:
    """Async wrapper -- delegates COM work to a thread so the event loop stays free."""
    from .dispatcher import call_tool_sync

    coerce_arguments(name, arguments)
    # Safe logging — guard against non-string `code` (clients sometimes send int/None)
    safe_args = {}
    for k, v in arguments.items():
        if k == "code" and isinstance(v, str):
            safe_args[k] = f"<VBA code: {len(v)} chars>"
        else:
            safe_args[k] = v
    log.info(">>> %s  %s", name, safe_args)

    loop = asyncio.get_running_loop()
    # Mark a tool call as in flight so the global dialog watchdog knows a
    # persistent modal on an ATTACHED Access instance was provoked by our
    # (now blocked) COM call and may be dismissed.  Cleared in finally so an
    # idle session never auto-dismisses the interactive user's dialogs.
    started = time.monotonic()
    _Session._tool_started = started
    try:
        text = await loop.run_in_executor(_com_executor, call_tool_sync, name, arguments)
    finally:
        _Session._tool_started = None
    # A watchdog dismissed a modal while this call was in flight: surface it.
    # A silently dismissed dialog (e.g. a save-changes prompt cancelled by the
    # Cancel-first policy) can make an otherwise clean result wrong (#31).
    d = _Session._last_dismissed
    if d and d[0] >= started:
        text += (
            f"\n\n[mcp-access] Note: a modal dialog ({d[1]!r}) was "
            "auto-dismissed by the dialog watchdog during this call. "
            "If the result looks wrong, the dialog is the likely cause."
        )
    return [types.TextContent(type="text", text=text)]


async def main() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


def cli() -> None:
    asyncio.run(main())
