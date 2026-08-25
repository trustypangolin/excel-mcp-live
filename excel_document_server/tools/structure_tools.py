"""Row/column insert-delete and range-sorting tools for live Excel COM automation."""

import json
import sys

_XL_ASCENDING, _XL_DESCENDING = 1, 2
_XL_YES, _XL_NO = 1, 2


def insert_rows(workbook: str = None, sheet: str = None, row_index: int = None, count: int = 1) -> str:
    """Insert one or more blank rows, shifting existing rows down.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        row_index: 1-based row number the new rows are inserted before (required).
        count: How many rows to insert.

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if row_index is None:
        return json.dumps({"error": "row_index is required"})
    if count < 1:
        return json.dumps({"error": "count must be at least 1"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        ws.Range(f"A{row_index}:A{row_index + count - 1}").EntireRow.Insert()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "inserted_at": row_index, "count": count})
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_rows(workbook: str = None, sheet: str = None, row_index: int = None, count: int = 1) -> str:
    """Delete one or more rows, shifting rows below them up.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        row_index: 1-based row number to start deleting from (required).
        count: How many rows to delete.

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if row_index is None:
        return json.dumps({"error": "row_index is required"})
    if count < 1:
        return json.dumps({"error": "count must be at least 1"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        ws.Range(f"A{row_index}:A{row_index + count - 1}").EntireRow.Delete()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "deleted_from": row_index, "count": count})
    except Exception as e:
        return json.dumps({"error": str(e)})


def insert_columns(workbook: str = None, sheet: str = None, column: str = None, count: int = 1) -> str:
    """Insert one or more blank columns, shifting existing columns right.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        column: Column letter(s) the new columns are inserted before, e.g. "B" (required).
        count: How many columns to insert.

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not column:
        return json.dumps({"error": "column is required"})
    if count < 1:
        return json.dumps({"error": "count must be at least 1"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet
        from excel_document_server.utils.conversions import column_letter, column_number

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        start_col = column_number(column)
        end_col_letter = column_letter(start_col + count - 1)
        ws.Range(f"{column.upper()}1:{end_col_letter}1").EntireColumn.Insert()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "inserted_at": column.upper(), "count": count})
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_columns(workbook: str = None, sheet: str = None, column: str = None, count: int = 1) -> str:
    """Delete one or more columns, shifting columns to their right left.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        column: Column letter(s) to start deleting from, e.g. "B" (required).
        count: How many columns to delete.

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not column:
        return json.dumps({"error": "column is required"})
    if count < 1:
        return json.dumps({"error": "count must be at least 1"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet
        from excel_document_server.utils.conversions import column_letter, column_number

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        start_col = column_number(column)
        end_col_letter = column_letter(start_col + count - 1)
        ws.Range(f"{column.upper()}1:{end_col_letter}1").EntireColumn.Delete()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "deleted_from": column.upper(), "count": count})
    except Exception as e:
        return json.dumps({"error": str(e)})


def sort_range(
    workbook: str = None,
    sheet: str = None,
    range_address: str = "A1",
    key_column: int = 1,
    ascending: bool = True,
    has_header: bool = True,
) -> str:
    """Sort a range by one column.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: The range to sort, e.g. "A1:C10".
        key_column: 1-based column offset within range_address to sort by
                    (1 = the range's leftmost column).
        ascending: Sort direction.
        has_header: Whether the range's first row is a header (excluded
                    from sorting, kept in place).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address
        from excel_document_server.utils.conversions import column_letter

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        if key_column < 1 or key_column > rng.Columns.Count:
            return json.dumps({"error": f"key_column {key_column} is outside the range's {rng.Columns.Count} column(s)"})

        # Exclude the header row ourselves rather than trusting Sort's own
        # Header parameter — passing it as a keyword argument (Header=...)
        # was silently ignored in live testing (the header row got sorted
        # into the data instead of staying put), a COM keyword-argument
        # binding issue in the same family as the Address/Resize quirks
        # documented in excel_com.py. Excluding the row ourselves and never
        # relying on Header sidesteps that ambiguity entirely.
        start_col_letter = column_letter(rng.Column)
        end_col_letter = column_letter(rng.Column + rng.Columns.Count - 1)
        if has_header:
            if rng.Rows.Count < 2:
                return json.dumps({"error": "Range has no data rows below the header"})
            data_start_row = rng.Row + 1
        else:
            data_start_row = rng.Row
        data_end_row = rng.Row + rng.Rows.Count - 1
        data_rng = ws.Range(f"{start_col_letter}{data_start_row}:{end_col_letter}{data_end_row}")

        key_col_letter = column_letter(rng.Column + key_column - 1)
        key_range = ws.Range(f"{key_col_letter}{data_start_row}")

        # All positional — Sort's own keyword arguments proved unreliable
        # via COM automation (see note above). Signature order per the VBA
        # object model: Key1, Order1, Key2, Type, Order2, Key3, Order3, Header.
        order_value = _XL_ASCENDING if ascending else _XL_DESCENDING
        data_rng.Sort(key_range, order_value, None, None, None, None, None, _XL_NO)

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(data_rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})
