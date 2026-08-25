# excel-mcp-live

**Edit a Microsoft Excel workbook while it's open — the AI attaches to the window you already
have, it doesn't start its own.**

`Live editing` &middot; `Windows (COM)` &middot; `19 tools`

---

Most Excel MCP servers (including this workspace's own
[mcp-server-excel](../mcp-server-excel)) work by starting a fresh, invisible Excel process, opening
your file in it, and closing it again when done. That's great for scripted automation, but it
means the AI is never looking at the same window you are.

`excel-mcp-live` does the opposite: it attaches to whatever Excel window(s) you already have open
via COM, using the same "find the running application" approach as
[word-mcp-live](../word-mcp-live)'s live editing tools. It never launches Excel and never closes
it — it only touches workbooks you opened yourself.

## Requirements

- **Windows** (COM automation is Windows-only — there is no macOS/Linux mode for this project)
- **Microsoft Excel** installed and **already running with a workbook open**
- **Python 3.11+**

## Installation

```bash
git clone <this-repo-url>
cd excel-mcp-live
pip install -e .
```

Or run straight from source with [uv](https://docs.astral.sh/uv/) — no install step needed:

```bash
uv run excel_mcp_server.py
```

## Client setup (`.mcp.json`)

Point your MCP client at the project directory. For Claude Code, add this to `.mcp.json` in your
project (or `~/.claude.json` for a user-wide config):

```json
{
  "mcpServers": {
    "excel-live": {
      "command": "uv",
      "args": [
        "--directory",
        "G:/Ai/MCP/excel-mcp-live",
        "run",
        "excel_mcp_server.py"
      ],
      "env": {
        "MCP_AUTHOR": "Your Name"
      }
    }
  }
}
```

If you installed it with `pip install -e .` instead, you can use the console script directly:

```json
{
  "mcpServers": {
    "excel-live": {
      "command": "excel-mcp-live",
      "env": {
        "MCP_AUTHOR": "Your Name"
      }
    }
  }
}
```

`MCP_AUTHOR` sets the author name attached to new cell comments (default: `"Author"`).

## Supported functions

Every tool operates on a workbook that's already open in Excel. `workbook` and `sheet` arguments
are optional almost everywhere — omit them to target the active workbook / active sheet.

| Tool | Description |
|---|---|
| **Workbook** | |
| `list_open_workbooks` | List every workbook open across all running Excel windows |
| `get_workbook_info` | Sheet names, active sheet, and used range per sheet |
| `save_workbook` | Save in place (Ctrl+S) |
| **Worksheets** | |
| `list_worksheets` | List sheet names in tab order |
| `add_worksheet` | Add a new sheet, optionally named and positioned |
| `delete_worksheet` | Delete a sheet |
| `rename_worksheet` | Rename a sheet |
| `activate_worksheet` | Switch the active tab |
| **Cells & ranges** | |
| `get_range_values` | Read cell/range values |
| `set_range_values` | Write a value, a row, or a 2D block of values |
| `get_range_formulas` | Read formulas (A1-style) |
| `set_cell_formula` | Set a formula or literal on one cell |
| `find_replace` | Find and replace text across a sheet or the whole workbook |
| `clear_range` | Clear contents, optionally including formatting |
| `insert_rows` / `delete_rows` | Insert or delete rows, shifting others up/down |
| `insert_columns` / `delete_columns` | Insert or delete columns, shifting others left/right |
| `sort_range` | Sort a range by one column, optionally keeping a header row in place |
| **Formatting** | |
| `format_range` | Bold/italic/underline, font, fill color, number format, alignment, wrap |
| `autofit_columns` | Autofit column widths |
| **Conditional formatting** | |
| `add_conditional_format_rule` | Highlight cells matching a comparison (greater than, between, etc.) |
| `add_color_scale` | Apply a 2- or 3-color scale across a range |
| `list_conditional_format_rules` | List the rules applied to a range |
| `clear_conditional_format_rules` | Remove all rules from a range |
| **Named ranges** | |
| `list_named_ranges` | List user-defined named ranges |
| `read_named_range` | Read a named range's current value(s) |
| `write_named_range` | Write a value into a named range's top-left cell |
| `create_named_range` | Create a new named range |
| `update_named_range` | Repoint an existing named range to a new address |
| `delete_named_range` | Delete a named range |
| **Data validation** | |
| `add_dropdown_validation` | Restrict a range to a dropdown list of values |
| `add_number_validation` | Restrict a range to numbers matching a comparison |
| `get_validation` | Read a cell's validation rule |
| `remove_validation` | Remove validation from a range |
| **Merging** | |
| `merge_cells` / `unmerge_cells` | Merge a range into one cell, or undo it |
| `get_merge_info` | Check whether a cell is part of a merged range |
| **Hyperlinks** | |
| `add_hyperlink` | Add a hyperlink (external URL or internal cell reference) |
| `remove_hyperlink` | Remove a cell's hyperlink |
| `list_hyperlinks` | List all hyperlinks on a sheet |
| **Excel Tables** | |
| `create_table` | Create an Excel Table (ListObject) from a range |
| `list_tables` | List tables on a sheet |
| `get_table_info` | Read a table's columns, range, row count, and style |
| `append_table_rows` | Add one or more rows to the end of a table |
| `get_table_data` | Read a table's data rows as a 2D array |
| `apply_table_style` | Apply a built-in table style |
| `delete_table` | Remove a table's definition, keeping the underlying data |
| **Comments** | |
| `add_comment` | Add a cell comment (note) |
| `get_comments` | List comments on a sheet |
| `delete_comment` | Remove a cell's comment |

Not yet covered: charts, PivotTables, Power Query, DAX, VBA — for those, use
[mcp-server-excel](../mcp-server-excel), which starts its own Excel instance and has much broader
operation coverage.

## Example prompts

```
"I have this budget open in Excel — read A1:F20 and tell me if any totals look wrong."
"Add a new sheet called 'Summary' before the current one."
"Bold row 1 in the sheet I have open and give it a light gray fill."
"Find every 'Q3' in this workbook and replace it with 'Q4'."
"Set B2 to =SUM(B3:B20) and tell me what it calculates to."
"Add a comment on cell D5 asking whether this figure includes tax."
```

## Known limitations

- **Windows only.** Excel COM automation doesn't exist on macOS/Linux.
- **The workbook must already be open in Excel.** This server doesn't create or open files — see
  mcp-server-excel for that.
- **No single-undo grouping.** Word's COM API can wrap several edits into one Ctrl+Z; Excel's
  cannot. A tool call that performs multiple writes (e.g. writing several cells) may need several
  Ctrl+Z's to fully undo.
- **Multiple Excel windows:** if you have more than one workbook open (possibly across separate
  Excel processes), pass `workbook` explicitly to target the right one — otherwise tools default to
  the active workbook of whichever Excel instance is found first.

## Development

See [CLAUDE.md](CLAUDE.md) for architecture notes and design constraints — in particular, why this
server must never launch or quit Excel itself.
