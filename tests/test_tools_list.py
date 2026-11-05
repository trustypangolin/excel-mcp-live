"""tools/list shape, payload budget and toolset filtering, via the MCP protocol.

None of these need Excel: registration and listing never touch COM.
"""

import asyncio
import json
import re

import pytest
from fastmcp import Client, FastMCP

from excel_document_server.main import register_tools
from excel_document_server.tool_metadata import SERVER_INSTRUCTIONS, parse_docstring
from excel_document_server.toolsets import (
    CORE_TOOLS, ENV_VAR, TOOL_GROUPS, TOOLSETS, parse_toolsets, tool_group,
)


def make_server(selection=None):
    server = FastMCP("Excel test", instructions=SERVER_INSTRUCTIONS)
    register_tools(server, selection)
    return server


def tools_for(selection=None):
    async def list_tools():
        async with Client(make_server(selection)) as client:
            return await client.list_tools()

    return asyncio.run(list_tools())


@pytest.fixture(scope="module")
def all_tools():
    return {tool.name: tool for tool in tools_for()}


def payload_size(tools):
    payload = [{"name": t.name, "description": t.description, "inputSchema": t.inputSchema} for t in tools]
    return len(json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode())


def props(tool):
    return tool.inputSchema["properties"]


# --- Descriptions and schemas ---

def test_descriptions_do_not_repeat_parameter_lists(all_tools):
    for tool in all_tools.values():
        assert not re.search(r"\b(?:Args|Arguments|Returns):", tool.description), tool.name
        assert "\n" not in tool.description, tool.name


def test_parameter_help_is_in_the_schema(all_tools):
    fmt = props(all_tools["format_range"])
    assert "Hex RGB" in fmt["font_color"]["description"]
    # A comma-grouped "bold, italic, underline:" entry documents each parameter.
    assert {fmt[p]["description"] for p in ("bold", "italic", "underline")} == {"Font style toggles."}
    assert "1-based" in props(all_tools["insert_rows"])["row_index"]["description"]
    # Caveats from Returns: survive in the description.
    assert "only remaining visible sheet" in all_tools["delete_worksheet"].description


def test_shared_boilerplate_is_stated_once(all_tools):
    workbook_docs = [props(t)["workbook"] for t in all_tools.values() if "workbook" in props(t)]
    assert len(workbook_docs) > 50
    assert all("description" not in p for p in workbook_docs)
    assert "description" not in props(all_tools["get_range_values"])["sheet"]
    # A non-default meaning is kept rather than dropped as boilerplate.
    assert "every sheet" in props(all_tools["find_replace"])["sheet"]["description"]
    assert "active workbook" in SERVER_INSTRUCTIONS and "active sheet" in SERVER_INSTRUCTIONS


def test_null_defaults_are_omitted_but_real_defaults_kept(all_tools):
    for tool in all_tools.values():
        for name, prop in props(tool).items():
            assert prop.get("default", "unset") is not None, (tool.name, name)
    assert props(all_tools["create_chart"])["chart_type"]["default"] == "column_clustered"


def test_payload_budget(all_tools):
    # Was ~38 KB before docs moved into the schema. Most of what remains is
    # schema structure for 59 tools; use EXCEL_MCP_TOOLSETS to go smaller.
    assert len(all_tools) == 59
    assert payload_size(all_tools.values()) < 27000


def test_hints(all_tools):
    for name in ("get_merge_info", "list_conditional_format_rules"):
        assert all_tools[name].annotations.readOnlyHint is True
        assert not all_tools[name].annotations.destructiveHint
    destructive = {n for n, t in all_tools.items() if t.annotations.destructiveHint}
    assert len(destructive) == 13
    assert not any(all_tools[n].annotations.readOnlyHint for n in destructive)


def test_server_instructions_are_sent():
    async def init():
        async with Client(make_server()) as client:
            return client.initialize_result.instructions

    assert asyncio.run(init()) == SERVER_INSTRUCTIONS


def test_parse_docstring_handles_grouped_and_wrapped_args():
    prose, args, returns = parse_docstring('''Do a thing.

    More detail.

    Args:
        left, top: Position in
                   points.
        name: Name (required).

    Returns:
        JSON confirmation. Excel may refuse.
    ''')
    assert prose == "Do a thing. More detail."
    assert args == {"left": "Position in points.", "top": "Position in points.", "name": "Name (required)."}
    assert returns == "JSON confirmation. Excel may refuse."


# --- Toolsets ---

def test_every_tool_belongs_to_exactly_one_group(all_tools):
    grouped = [name for names in TOOL_GROUPS.values() for name in names]
    assert len(grouped) == len(set(grouped))
    assert set(grouped) == set(all_tools)
    with pytest.raises(KeyError, match="TOOL_GROUPS"):
        tool_group("not_a_tool")


@pytest.mark.parametrize("value,expected", [
    (None, None), ("", None), (" all ", None),
    (" CHARTS, tables,charts ", frozenset({"charts", "tables"})),
])
def test_parse_toolsets(value, expected):
    assert parse_toolsets(value) == expected


@pytest.mark.parametrize("value", ["charts,", ",charts", "charts,,tables", "all,charts", "unknown"])
def test_invalid_toolsets_fail(value):
    with pytest.raises(ValueError, match=ENV_VAR):
        parse_toolsets(value)


@pytest.mark.parametrize("group", sorted(TOOLSETS))
def test_each_toolset_registers_its_tools_plus_core(group):
    names = {t.name for t in tools_for(frozenset({group}))}
    assert names == set(TOOL_GROUPS[group]) | CORE_TOOLS


def test_toolsets_are_a_union():
    names = {t.name for t in tools_for(frozenset({"charts", "tables"}))}
    assert names == set(TOOL_GROUPS["charts"]) | set(TOOL_GROUPS["tables"]) | CORE_TOOLS


def test_selection_is_read_from_environment(monkeypatch):
    monkeypatch.setenv(ENV_VAR, "comments")
    assert {t.name for t in tools_for()} == set(TOOL_GROUPS["comments"]) | CORE_TOOLS
    monkeypatch.setenv(ENV_VAR, "bogus")
    with pytest.raises(ValueError, match=ENV_VAR):
        tools_for()


def test_excluded_tool_cannot_be_called():
    async def call():
        async with Client(make_server(frozenset({"charts"}))) as client:
            with pytest.raises(Exception, match="(?i)(unknown|not found)"):
                await client.call_tool("get_range_values", {"range_address": "A1"})

    asyncio.run(call())
