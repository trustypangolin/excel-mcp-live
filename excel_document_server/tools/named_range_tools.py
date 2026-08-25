"""Named range tools for live Excel COM automation.

Named ranges are workbook-scoped labels for cells/ranges — useful for
formulas ("=SUM(SalesData)") and as parameters that Power Query or other
formulas react to when their value changes.
"""

import json
import sys


def list_named_ranges(workbook: str = None) -> str:
    """List user-defined named ranges in a workbook.

    Hidden/internal names (e.g. Excel's own Print_Area, filter ranges) are
    omitted.

    Args:
        workbook: Workbook name or path (None = active workbook).

    Returns:
        JSON with each name and what it refers to.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook

        _app, wb = find_workbook(workbook)

        names = []
        for i in range(1, wb.Names.Count + 1):
            nm = wb.Names(i)
            try:
                if not nm.Visible:
                    continue
            except Exception:
                pass
            names.append({"name": nm.Name, "refers_to": nm.RefersTo})

        return json.dumps({"success": True, "workbook": wb.Name, "names": names}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def read_named_range(workbook: str = None, name: str = None) -> str:
    """Read the current value(s) of a named range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Named range to read (required).

    Returns:
        JSON with the value(s) as a 2D array.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not name:
        return json.dumps({"error": "name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook
        from excel_document_server.utils.conversions import to_jsonable

        _app, wb = find_workbook(workbook)

        try:
            rng = wb.Names(name).RefersToRange
        except Exception:
            return json.dumps({"error": f"Named range '{name}' not found or does not refer to a range"})

        raw = rng.Value
        if raw is None:
            values = [[None]]
        elif isinstance(raw, tuple):
            values = to_jsonable(raw)
        else:
            values = [[to_jsonable(raw)]]

        return json.dumps({"success": True, "workbook": wb.Name, "name": name, "values": values}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def write_named_range(workbook: str = None, name: str = None, value: str | float | int | bool | None = None) -> str:
    """Write a single value into a named range's top-left cell.

    For multi-cell named ranges, only the top-left cell is set — use
    set_range_values with the range's address for block writes.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Named range to write to (required).
        value: New value (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not name:
        return json.dumps({"error": "name is required"})
    if value is None:
        return json.dumps({"error": "value is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook

        _app, wb = find_workbook(workbook)

        try:
            rng = wb.Names(name).RefersToRange
        except Exception:
            return json.dumps({"error": f"Named range '{name}' not found or does not refer to a range"})

        rng.Cells(1, 1).Value = value

        return json.dumps({"success": True, "workbook": wb.Name, "name": name})
    except Exception as e:
        return json.dumps({"error": str(e)})


def create_named_range(workbook: str = None, name: str = None, range_address: str = "A1", sheet: str = None) -> str:
    """Create a new named range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Name to create (required). Must start with a letter or
              underscore and contain no spaces.
        range_address: A1-style address the name refers to, e.g. "A1:C10".
        sheet: Worksheet the range is on (None = active sheet).

    Returns:
        JSON confirmation with the resolved absolute reference.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not name:
        return json.dumps({"error": "name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet
        from excel_document_server.utils.conversions import column_letter

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        start_row, start_col = rng.Row, rng.Column
        end_row = start_row + rng.Rows.Count - 1
        end_col = start_col + rng.Columns.Count - 1
        refers_to = f"='{ws.Name}'!${column_letter(start_col)}${start_row}:${column_letter(end_col)}${end_row}"

        wb.Names.Add(name, refers_to)

        return json.dumps({"success": True, "workbook": wb.Name, "name": name, "refers_to": refers_to})
    except Exception as e:
        return json.dumps({"error": str(e)})


def update_named_range(workbook: str = None, name: str = None, range_address: str = "A1", sheet: str = None) -> str:
    """Change what an existing named range refers to.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Existing named range to update (required).
        range_address: New A1-style address, e.g. "A1:C10".
        sheet: Worksheet the new range is on (None = active sheet).

    Returns:
        JSON confirmation with the resolved absolute reference.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not name:
        return json.dumps({"error": "name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet
        from excel_document_server.utils.conversions import column_letter

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        try:
            existing = wb.Names(name)
        except Exception:
            return json.dumps({"error": f"Named range '{name}' not found"})

        start_row, start_col = rng.Row, rng.Column
        end_row = start_row + rng.Rows.Count - 1
        end_col = start_col + rng.Columns.Count - 1
        refers_to = f"='{ws.Name}'!${column_letter(start_col)}${start_row}:${column_letter(end_col)}${end_row}"

        existing.RefersTo = refers_to

        return json.dumps({"success": True, "workbook": wb.Name, "name": name, "refers_to": refers_to})
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_named_range(workbook: str = None, name: str = None) -> str:
    """Delete a named range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Named range to delete (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not name:
        return json.dumps({"error": "name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook

        _app, wb = find_workbook(workbook)

        try:
            wb.Names(name).Delete()
        except Exception:
            return json.dumps({"error": f"Named range '{name}' not found"})

        return json.dumps({"success": True, "workbook": wb.Name, "deleted": name})
    except Exception as e:
        return json.dumps({"error": str(e)})
