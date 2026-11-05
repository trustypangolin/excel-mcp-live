"""Startup-time tool filtering via the EXCEL_MCP_TOOLSETS environment variable.

Mirrors word-mcp-live's WORD_MCP_TOOLSETS: a comma-separated union of groups,
where unset, empty or ``all`` registers every tool. ``list_open_workbooks`` is
always registered, since every other tool needs a workbook name to target.
"""

from fastmcp.tools import Tool

from .tool_metadata import drop_null_defaults, prepare_tool


ENV_VAR = "EXCEL_MCP_TOOLSETS"

# Always registered, whatever the selection.
CORE_TOOLS = frozenset({"list_open_workbooks"})

TOOL_GROUPS = {
    "workbooks": (
        "list_open_workbooks", "get_workbook_info", "save_workbook",
    ),
    "sheets": (
        "list_worksheets", "add_worksheet", "delete_worksheet", "rename_worksheet",
        "activate_worksheet",
    ),
    "ranges": (
        "get_range_values", "set_range_values", "get_range_formulas", "set_cell_formula",
        "find_replace", "clear_range", "insert_rows", "delete_rows", "insert_columns",
        "delete_columns", "sort_range", "merge_cells", "unmerge_cells", "get_merge_info",
    ),
    "formatting": (
        "format_range", "autofit_columns", "add_conditional_format_rule", "add_color_scale",
        "list_conditional_format_rules", "clear_conditional_format_rules",
    ),
    "comments": (
        "add_comment", "get_comments", "delete_comment",
    ),
    "named_ranges": (
        "list_named_ranges", "read_named_range", "write_named_range", "create_named_range",
        "update_named_range", "delete_named_range",
    ),
    "validation": (
        "add_dropdown_validation", "add_number_validation", "get_validation",
        "remove_validation",
    ),
    "hyperlinks": (
        "add_hyperlink", "remove_hyperlink", "list_hyperlinks",
    ),
    "tables": (
        "create_table", "list_tables", "get_table_info", "append_table_rows", "get_table_data",
        "apply_table_style", "delete_table",
    ),
    "charts": (
        "create_chart", "list_charts", "set_chart_type", "set_chart_title", "add_series",
        "delete_chart", "format_chart_series", "set_chart_legend",
    ),
}

TOOLSETS = frozenset(TOOL_GROUPS)

_GROUP_OF = {name: group for group, names in TOOL_GROUPS.items() for name in names}


def tool_group(name):
    """The toolset a tool belongs to. Every registered tool must have one."""
    try:
        return _GROUP_OF[name]
    except KeyError:
        raise KeyError(f"Tool {name!r} is not assigned to a toolset in toolsets.TOOL_GROUPS") from None


def parse_toolsets(value):
    """Parse a comma-separated union of toolsets; None means all tools."""
    if value is None or not value.strip():
        return None
    parts = [part.strip().lower() for part in value.split(",")]
    if any(not part for part in parts):
        raise ValueError(f"{ENV_VAR} contains an empty toolset name")
    if parts == ["all"]:
        return None
    unknown = set(parts) - TOOLSETS
    if unknown:
        raise ValueError(
            f"Unknown {ENV_VAR} value(s): " + ", ".join(sorted(unknown))
            + ". Choose from all, " + ", ".join(sorted(TOOLSETS))
        )
    return frozenset(parts)


class ToolRegistrar:
    """Stand-in for ``FastMCP`` in register_tools(): filters by toolset and compacts metadata."""

    def __init__(self, server, selected):
        self.server = server
        self.selected = selected

    def tool(self, **kwargs):
        def register(function):
            name = function.__name__
            group = tool_group(name)
            if self.selected is None or name in CORE_TOOLS or group in self.selected:
                prepared, options = prepare_tool(function, **kwargs)
                tool = Tool.from_function(prepared, **options)
                drop_null_defaults(tool.parameters)
                self.server.add_tool(tool)
            return function
        return register
