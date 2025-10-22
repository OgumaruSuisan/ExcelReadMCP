"""MCP server exposing read-only Excel tools."""

from __future__ import annotations

import asyncio
import json
import logging
from typing import Any, Dict, List

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from .core import ExcelReadTools

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


class ExcelReadMCPServer:
    """Server wiring ExcelReadTools into the Model Context Protocol."""

    def __init__(self) -> None:
        self.server = Server("excel-read-tools")
        self.tools = ExcelReadTools()
        self._register_tools()

    def _register_tools(self) -> None:
        """Set up tool listing and execution hooks."""

        @self.server.list_tools()
        async def list_tools() -> List[Tool]:
            return [
                Tool(
                    name="excel_read_info",
                    description="Return workbook metadata such as sheet names and size.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute or relative path to an Excel file.",
                            }
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="excel_read_range",
                    description="Read a sheet and return its contents as JSON records.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Excel file to read.",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Sheet name (defaults to the first sheet).",
                            },
                            "range_spec": {
                                "type": "string",
                                "description": "Optional Excel range such as A1:C10.",
                            },
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="excel_read_all_sheets",
                    description="Read every sheet with optional row limits.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Excel file to read.",
                            },
                            "max_rows_per_sheet": {
                                "type": "number",
                                "description": "Maximum number of rows to return per sheet.",
                                "default": 1000,
                            },
                            "include_empty_sheets": {
                                "type": "boolean",
                                "description": "Include sheets with no data.",
                                "default": False,
                            },
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="excel_quick_overview",
                    description="Provide workbook metadata and a small sample per sheet.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Excel file to read.",
                            },
                            "sample_rows": {
                                "type": "number",
                                "description": "Number of rows to sample per sheet.",
                                "default": 5,
                            },
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="excel_search",
                    description="Search for text across the workbook or a specific sheet.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Excel file to inspect.",
                            },
                            "search_term": {
                                "type": "string",
                                "description": "Text to search for (case-insensitive).",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Optional sheet to restrict the search.",
                            },
                        },
                        "required": ["file_path", "search_term"],
                    },
                ),
            ]

        @self.server.call_tool()
        async def call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent]:
            try:
                if name == "excel_read_info":
                    result = self.tools.excel_read_info(arguments["file_path"])
                elif name == "excel_read_range":
                    result = self.tools.excel_read_range(
                        file_path=arguments["file_path"],
                        sheet_name=arguments.get("sheet_name"),
                        range_spec=arguments.get("range_spec"),
                    )
                elif name == "excel_read_all_sheets":
                    result = self.tools.excel_read_all_sheets(
                        file_path=arguments["file_path"],
                        max_rows_per_sheet=int(arguments.get("max_rows_per_sheet", 1000)),
                        include_empty_sheets=bool(arguments.get("include_empty_sheets", False)),
                    )
                elif name == "excel_quick_overview":
                    result = self.tools.excel_quick_overview(
                        file_path=arguments["file_path"],
                        sample_rows=int(arguments.get("sample_rows", 5)),
                    )
                elif name == "excel_search":
                    result = self.tools.excel_search(
                        file_path=arguments["file_path"],
                        search_term=arguments["search_term"],
                        sheet_name=arguments.get("sheet_name"),
                    )
                else:
                    result = {"success": False, "error": f"Unknown tool: {name}"}

                return [
                    TextContent(
                        type="text",
                        text=json.dumps(result, ensure_ascii=False, indent=2),
                    )
                ]
            except Exception as exc:
                logger.error("Tool execution failed (%s): %s", name, exc)
                error_result = {"success": False, "error": str(exc)}
                return [
                    TextContent(
                        type="text",
                        text=json.dumps(error_result, ensure_ascii=False, indent=2),
                    )
                ]

    async def run(self) -> None:
        """Launch the MCP stdio server."""
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options(),
            )


async def main() -> None:
    server = ExcelReadMCPServer()
    await server.run()


if __name__ == "__main__":
    asyncio.run(main())
