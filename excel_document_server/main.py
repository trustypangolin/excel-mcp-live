"""
Main entry point for the Excel Live MCP Server.

Live-attach tools that operate on workbooks already open in Excel via COM
automation. Requires Windows with Excel installed and running — this server
never launches or quits Excel itself.
"""

import sys

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

from excel_document_server.tools import (
    workbook_tools,
    sheet_tools,
    range_tools,
    format_tools,
    comment_tools,
    structure_tools,
    conditional_format_tools,
    named_range_tools,
)

mcp = FastMCP("Excel Live MCP Server")


def register_tools():
    """Register all tools with the MCP server using FastMCP decorators."""

    # --- Workbook tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="List Open Workbooks", readOnlyHint=True),
        description=workbook_tools.list_open_workbooks.__doc__,
    )
    def list_open_workbooks():
        return workbook_tools.list_open_workbooks()

    @mcp.tool(
        annotations=ToolAnnotations(title="Get Workbook Info", readOnlyHint=True),
        description=workbook_tools.get_workbook_info.__doc__,
    )
    def get_workbook_info(workbook: str = None):
        return workbook_tools.get_workbook_info(workbook)

    @mcp.tool(
        annotations=ToolAnnotations(title="Save Workbook"),
        description=workbook_tools.save_workbook.__doc__,
    )
    def save_workbook(workbook: str = None):
        return workbook_tools.save_workbook(workbook)

    # --- Sheet tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="List Worksheets", readOnlyHint=True),
        description=sheet_tools.list_worksheets.__doc__,
    )
    def list_worksheets(workbook: str = None):
        return sheet_tools.list_worksheets(workbook)

    @mcp.tool(
        annotations=ToolAnnotations(title="Add Worksheet"),
        description=sheet_tools.add_worksheet.__doc__,
    )
    def add_worksheet(workbook: str = None, name: str = None, index: int = None):
        return sheet_tools.add_worksheet(workbook, name, index)

    @mcp.tool(
        annotations=ToolAnnotations(title="Delete Worksheet", destructiveHint=True),
        description=sheet_tools.delete_worksheet.__doc__,
    )
    def delete_worksheet(workbook: str = None, name: str = None):
        return sheet_tools.delete_worksheet(workbook, name)

    @mcp.tool(
        annotations=ToolAnnotations(title="Rename Worksheet"),
        description=sheet_tools.rename_worksheet.__doc__,
    )
    def rename_worksheet(workbook: str = None, name: str = None, new_name: str = None):
        return sheet_tools.rename_worksheet(workbook, name, new_name)

    @mcp.tool(
        annotations=ToolAnnotations(title="Activate Worksheet"),
        description=sheet_tools.activate_worksheet.__doc__,
    )
    def activate_worksheet(workbook: str = None, name: str = None):
        return sheet_tools.activate_worksheet(workbook, name)

    # --- Range / cell tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Get Range Values", readOnlyHint=True),
        description=range_tools.get_range_values.__doc__,
    )
    def get_range_values(workbook: str = None, sheet: str = None, range_address: str = "A1"):
        return range_tools.get_range_values(workbook, sheet, range_address)

    @mcp.tool(
        annotations=ToolAnnotations(title="Set Range Values"),
        description=range_tools.set_range_values.__doc__,
    )
    def set_range_values(
        workbook: str = None,
        sheet: str = None,
        range_address: str = "A1",
        values: list | str | float | int | bool | None = None,
    ):
        return range_tools.set_range_values(workbook, sheet, range_address, values)

    @mcp.tool(
        annotations=ToolAnnotations(title="Get Range Formulas", readOnlyHint=True),
        description=range_tools.get_range_formulas.__doc__,
    )
    def get_range_formulas(workbook: str = None, sheet: str = None, range_address: str = "A1"):
        return range_tools.get_range_formulas(workbook, sheet, range_address)

    @mcp.tool(
        annotations=ToolAnnotations(title="Set Cell Formula"),
        description=range_tools.set_cell_formula.__doc__,
    )
    def set_cell_formula(workbook: str = None, sheet: str = None, cell: str = "A1", formula: str = ""):
        return range_tools.set_cell_formula(workbook, sheet, cell, formula)

    @mcp.tool(
        annotations=ToolAnnotations(title="Find and Replace", destructiveHint=True),
        description=range_tools.find_replace.__doc__,
    )
    def find_replace(
        workbook: str = None,
        sheet: str = None,
        find_text: str = "",
        replace_text: str = "",
        match_case: bool = False,
        whole_cell: bool = False,
    ):
        return range_tools.find_replace(workbook, sheet, find_text, replace_text, match_case, whole_cell)

    @mcp.tool(
        annotations=ToolAnnotations(title="Clear Range", destructiveHint=True),
        description=range_tools.clear_range.__doc__,
    )
    def clear_range(workbook: str = None, sheet: str = None, range_address: str = "A1", clear_formatting: bool = False):
        return range_tools.clear_range(workbook, sheet, range_address, clear_formatting)

    # --- Formatting tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Format Range"),
        description=format_tools.format_range.__doc__,
    )
    def format_range(
        workbook: str = None,
        sheet: str = None,
        range_address: str = "A1",
        bold: bool = None,
        italic: bool = None,
        underline: bool = None,
        font_name: str = None,
        font_size: float = None,
        font_color: str = None,
        fill_color: str = None,
        number_format: str = None,
        horizontal_align: str = None,
        wrap_text: bool = None,
    ):
        return format_tools.format_range(
            workbook, sheet, range_address, bold, italic, underline,
            font_name, font_size, font_color, fill_color, number_format,
            horizontal_align, wrap_text,
        )

    @mcp.tool(
        annotations=ToolAnnotations(title="Autofit Columns"),
        description=format_tools.autofit_columns.__doc__,
    )
    def autofit_columns(workbook: str = None, sheet: str = None, range_address: str = None):
        return format_tools.autofit_columns(workbook, sheet, range_address)

    # --- Comment tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Add Comment"),
        description=comment_tools.add_comment.__doc__,
    )
    def add_comment(workbook: str = None, sheet: str = None, cell: str = "A1", text: str = "", author: str = None):
        return comment_tools.add_comment(workbook, sheet, cell, text, author)

    @mcp.tool(
        annotations=ToolAnnotations(title="Get Comments", readOnlyHint=True),
        description=comment_tools.get_comments.__doc__,
    )
    def get_comments(workbook: str = None, sheet: str = None):
        return comment_tools.get_comments(workbook, sheet)

    @mcp.tool(
        annotations=ToolAnnotations(title="Delete Comment", destructiveHint=True),
        description=comment_tools.delete_comment.__doc__,
    )
    def delete_comment(workbook: str = None, sheet: str = None, cell: str = "A1"):
        return comment_tools.delete_comment(workbook, sheet, cell)

    # --- Row/column structure tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Insert Rows"),
        description=structure_tools.insert_rows.__doc__,
    )
    def insert_rows(workbook: str = None, sheet: str = None, row_index: int = None, count: int = 1):
        return structure_tools.insert_rows(workbook, sheet, row_index, count)

    @mcp.tool(
        annotations=ToolAnnotations(title="Delete Rows", destructiveHint=True),
        description=structure_tools.delete_rows.__doc__,
    )
    def delete_rows(workbook: str = None, sheet: str = None, row_index: int = None, count: int = 1):
        return structure_tools.delete_rows(workbook, sheet, row_index, count)

    @mcp.tool(
        annotations=ToolAnnotations(title="Insert Columns"),
        description=structure_tools.insert_columns.__doc__,
    )
    def insert_columns(workbook: str = None, sheet: str = None, column: str = None, count: int = 1):
        return structure_tools.insert_columns(workbook, sheet, column, count)

    @mcp.tool(
        annotations=ToolAnnotations(title="Delete Columns", destructiveHint=True),
        description=structure_tools.delete_columns.__doc__,
    )
    def delete_columns(workbook: str = None, sheet: str = None, column: str = None, count: int = 1):
        return structure_tools.delete_columns(workbook, sheet, column, count)

    @mcp.tool(
        annotations=ToolAnnotations(title="Sort Range", destructiveHint=True),
        description=structure_tools.sort_range.__doc__,
    )
    def sort_range(
        workbook: str = None,
        sheet: str = None,
        range_address: str = "A1",
        key_column: int = 1,
        ascending: bool = True,
        has_header: bool = True,
    ):
        return structure_tools.sort_range(workbook, sheet, range_address, key_column, ascending, has_header)

    # --- Conditional formatting tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="Add Conditional Formatting Rule"),
        description=conditional_format_tools.add_rule.__doc__,
    )
    def add_conditional_format_rule(
        workbook: str = None,
        sheet: str = None,
        range_address: str = "A1",
        condition: str = "greater_than",
        value1: str | float | int | None = None,
        value2: str | float | int | None = None,
        fill_color: str = None,
        font_color: str = None,
        bold: bool = None,
    ):
        return conditional_format_tools.add_rule(
            workbook, sheet, range_address, condition, value1, value2, fill_color, font_color, bold,
        )

    @mcp.tool(
        annotations=ToolAnnotations(title="Add Color Scale"),
        description=conditional_format_tools.add_color_scale.__doc__,
    )
    def add_color_scale(
        workbook: str = None,
        sheet: str = None,
        range_address: str = "A1",
        min_color: str = "F8696B",
        max_color: str = "63BE7B",
        mid_color: str = None,
    ):
        return conditional_format_tools.add_color_scale(workbook, sheet, range_address, min_color, max_color, mid_color)

    @mcp.tool(
        annotations=ToolAnnotations(title="List Conditional Formatting Rules", readOnlyHint=True),
        description=conditional_format_tools.list_rules.__doc__,
    )
    def list_conditional_format_rules(workbook: str = None, sheet: str = None, range_address: str = "A1"):
        return conditional_format_tools.list_rules(workbook, sheet, range_address)

    @mcp.tool(
        annotations=ToolAnnotations(title="Clear Conditional Formatting Rules", destructiveHint=True),
        description=conditional_format_tools.clear_rules.__doc__,
    )
    def clear_conditional_format_rules(workbook: str = None, sheet: str = None, range_address: str = "A1"):
        return conditional_format_tools.clear_rules(workbook, sheet, range_address)

    # --- Named range tools ---

    @mcp.tool(
        annotations=ToolAnnotations(title="List Named Ranges", readOnlyHint=True),
        description=named_range_tools.list_named_ranges.__doc__,
    )
    def list_named_ranges(workbook: str = None):
        return named_range_tools.list_named_ranges(workbook)

    @mcp.tool(
        annotations=ToolAnnotations(title="Read Named Range", readOnlyHint=True),
        description=named_range_tools.read_named_range.__doc__,
    )
    def read_named_range(workbook: str = None, name: str = None):
        return named_range_tools.read_named_range(workbook, name)

    @mcp.tool(
        annotations=ToolAnnotations(title="Write Named Range"),
        description=named_range_tools.write_named_range.__doc__,
    )
    def write_named_range(workbook: str = None, name: str = None, value: str | float | int | bool | None = None):
        return named_range_tools.write_named_range(workbook, name, value)

    @mcp.tool(
        annotations=ToolAnnotations(title="Create Named Range"),
        description=named_range_tools.create_named_range.__doc__,
    )
    def create_named_range(workbook: str = None, name: str = None, range_address: str = "A1", sheet: str = None):
        return named_range_tools.create_named_range(workbook, name, range_address, sheet)

    @mcp.tool(
        annotations=ToolAnnotations(title="Update Named Range"),
        description=named_range_tools.update_named_range.__doc__,
    )
    def update_named_range(workbook: str = None, name: str = None, range_address: str = "A1", sheet: str = None):
        return named_range_tools.update_named_range(workbook, name, range_address, sheet)

    @mcp.tool(
        annotations=ToolAnnotations(title="Delete Named Range", destructiveHint=True),
        description=named_range_tools.delete_named_range.__doc__,
    )
    def delete_named_range(workbook: str = None, name: str = None):
        return named_range_tools.delete_named_range(workbook, name)


def run_server():
    """Run the Excel Live MCP Server over stdio."""
    register_tools()
    print("Excel Live MCP Server starting (stdio transport)...", file=sys.stderr)
    mcp.run(transport="stdio")


def main():
    """Main entry point for the server."""
    run_server()


if __name__ == "__main__":
    main()
