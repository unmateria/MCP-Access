"""
MCP Server setup: list_tools, list_prompts, get_prompt, call_tool, main.
"""

import asyncio
import json
import traceback

import mcp.types as types
from mcp.server import Server
from mcp.server.stdio import stdio_server

from .core import _Session, _com_executor, log
from .tools import TOOLS, coerce_arguments


server = Server("access-mcp")


@server.list_tools()
async def list_tools() -> list[types.Tool]:
    return TOOLS


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


@server.get_prompt()
async def get_prompt(name: str, arguments: dict | None) -> types.GetPromptResult:
    db_path = (arguments or {}).get("db_path", "<path_to_file.accdb>")
    return types.GetPromptResult(
        description="Required workflow for working with Access databases",
        messages=[
            types.PromptMessage(
                role="user",
                content=types.TextContent(
                    type="text",
                    text=f"""
I'm working with a Microsoft Access database: {db_path}

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
    # Safe logging
    safe_args = {}
    for k, v in arguments.items():
        if k == "code":
            safe_args[k] = f"<VBA code: {len(v)} chars>"
        else:
            safe_args[k] = v
    log.info(">>> %s  %s", name, safe_args)

    loop = asyncio.get_running_loop()
    text = await loop.run_in_executor(_com_executor, call_tool_sync, name, arguments)
    return [types.TextContent(type="text", text=text)]


async def main() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )
