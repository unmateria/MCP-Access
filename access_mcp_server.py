#!/usr/bin/env python3
"""
access_mcp_server.py
====================
MCP Server for reading and editing Microsoft Access databases (.accdb/.mdb)
via COM automation (pywin32). Requires Windows + Microsoft Access installed.

Install dependencies:
    pip install mcp pywin32

Register in Claude Code (one of two methods):
    # Option A -- global
    claude mcp add access -- python /path/to/access_mcp_server.py

    # Option B -- this project only (creates .mcp.json in current directory)
    claude mcp add --scope project access -- python /path/to/access_mcp_server.py

Typical workflow for editing VBA:
    1. access_list_objects  -> see which modules/forms exist
    2. access_get_code      -> export the object to text
    3. (Claude edits the text)
    4. access_set_code      -> reimport the modified text
    5. access_close         -> release Access (optional)
"""

import asyncio
from mcp_access.server import main

if __name__ == "__main__":
    asyncio.run(main())
