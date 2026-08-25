"""Excel Table (ListObject) tools for live Excel COM automation.

A lean subset: create, list, inspect, append rows, read data, style, and
remove. Not yet covered: filters, column add/remove, structured
references, DAX-backed tables — use mcp-server-excel for those.
"""

import json
import sys

_XL_SRC_RANGE = 1
_XL_YES, _XL_NO = 1, 2


def _find_table(ws, name: str):
    """Find a ListObject by name on a worksheet. Raises ValueError if not found."""
    try:
        return ws.ListObjects(name)
    except Exception:
        names = [ws.ListObjects(i).Name for i in range(1, ws.ListObjects.Count + 1)]
        raise ValueError(f"Table '{name}' not found. Available tables: {names}")


def create_table(
    workbook: str = None,
    sheet: str = None,
    range_address: str = "A1",
    table_name: str = None,
    has_headers: bool = True,
) -> str:
    """Create an Excel Table from a range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address the table will cover, e.g. "A1:C10".
        table_name: Name for the new table (None = Excel's default, e.g. "Table1").
        has_headers: Whether the range's first row is a header row.

    Returns:
        JSON with the new table's name and range.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        tbl = ws.ListObjects.Add(SourceType=_XL_SRC_RANGE, Source=rng, XlListObjectHasHeaders=(_XL_YES if has_headers else _XL_NO))
        if table_name:
            tbl.Name = table_name

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "table_name": tbl.Name,
            "range": com_range_address(tbl.Range),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def list_tables(workbook: str = None, sheet: str = None) -> str:
    """List Excel Tables on a worksheet.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).

    Returns:
        JSON with each table's name and range.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)

        tables = []
        for i in range(1, ws.ListObjects.Count + 1):
            tbl = ws.ListObjects(i)
            tables.append({"name": tbl.Name, "range": com_range_address(tbl.Range)})

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "tables": tables}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_table_info(workbook: str = None, sheet: str = None, table_name: str = None) -> str:
    """Get a table's structure: columns, range, and style.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        table_name: Table to inspect (required).

    Returns:
        JSON with column names, range, row count, and style.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not table_name:
        return json.dumps({"error": "table_name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        tbl = _find_table(ws, table_name)

        columns = [tbl.ListColumns(i).Name for i in range(1, tbl.ListColumns.Count + 1)]

        try:
            row_count = tbl.ListRows.Count
        except Exception:
            row_count = 0

        # TableStyle reads back as a TableStyle COM object, not a plain
        # string, even though it's settable with a plain style-name string
        # (confirmed in live testing — json.dumps() raised "Object of type
        # CDispatch is not JSON serializable" when this wasn't unwrapped).
        style = tbl.TableStyle
        style_name = style.Name if hasattr(style, "Name") else str(style)

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "table_name": tbl.Name,
            "range": com_range_address(tbl.Range),
            "columns": columns,
            "row_count": row_count,
            "style": style_name,
        }, ensure_ascii=False)
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def append_table_rows(workbook: str = None, sheet: str = None, table_name: str = None, rows: list = None) -> str:
    """Append one or more rows to the end of a table.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        table_name: Table to append to (required).
        rows: A list of rows, each a list of values matching the table's
              column count (required), e.g. [["Widget", 12, 4.5]].

    Returns:
        JSON with the table's resulting row count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not table_name:
        return json.dumps({"error": "table_name is required"})
    if not rows:
        return json.dumps({"error": "rows is required (a non-empty list of row lists)"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        tbl = _find_table(ws, table_name)

        for row in rows:
            new_row = tbl.ListRows.Add()
            new_row.Range.Value = tuple(row)

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "table_name": tbl.Name, "row_count": tbl.ListRows.Count})
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_table_data(workbook: str = None, sheet: str = None, table_name: str = None) -> str:
    """Read a table's data rows (excluding the header) as a 2D array.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        table_name: Table to read (required).

    Returns:
        JSON with the column names and data rows.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not table_name:
        return json.dumps({"error": "table_name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet
        from excel_document_server.utils.conversions import to_jsonable

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        tbl = _find_table(ws, table_name)

        columns = [tbl.ListColumns(i).Name for i in range(1, tbl.ListColumns.Count + 1)]

        try:
            raw = tbl.DataBodyRange.Value
        except Exception:
            raw = None

        if raw is None:
            values = []
        elif isinstance(raw, tuple):
            values = to_jsonable(raw)
        else:
            values = [[to_jsonable(raw)]]

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "table_name": tbl.Name,
            "columns": columns,
            "values": values,
        }, ensure_ascii=False)
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def apply_table_style(workbook: str = None, sheet: str = None, table_name: str = None, style_name: str = "TableStyleMedium2") -> str:
    """Apply a built-in table style.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        table_name: Table to style (required).
        style_name: Built-in style name, e.g. "TableStyleMedium2",
                    "TableStyleLight9", "TableStyleDark3".

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not table_name:
        return json.dumps({"error": "table_name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        tbl = _find_table(ws, table_name)
        tbl.TableStyle = style_name

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "table_name": tbl.Name, "style": style_name})
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_table(workbook: str = None, sheet: str = None, table_name: str = None) -> str:
    """Remove a table's definition, keeping the underlying cell data.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        table_name: Table to remove (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not table_name:
        return json.dumps({"error": "table_name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        tbl = _find_table(ws, table_name)
        tbl.Unlist()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "deleted": table_name})
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})
