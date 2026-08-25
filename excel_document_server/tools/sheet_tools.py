"""Worksheet management tools for live Excel COM automation."""

import json
import sys


def list_worksheets(workbook: str = None) -> str:
    """List worksheet names, in tab order, for an open workbook.

    Args:
        workbook: Workbook name or path (None = active workbook).

    Returns:
        JSON with sheet names and which one is active.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook

        _app, wb = find_workbook(workbook)
        sheets = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "active_sheet": wb.ActiveSheet.Name,
            "sheets": sheets,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def add_worksheet(workbook: str = None, name: str = None, index: int = None) -> str:
    """Add a new worksheet to an open workbook.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Name for the new sheet (None = Excel's default, e.g. "Sheet4").
        index: 1-based tab position to insert before (None = before the
               active sheet).

    Returns:
        JSON with the new sheet's name and position.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook

        _app, wb = find_workbook(workbook)

        if index is not None:
            before = wb.Worksheets(index)
            new_sheet = wb.Worksheets.Add(Before=before)
        else:
            new_sheet = wb.Worksheets.Add()

        if name:
            new_sheet.Name = name

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": new_sheet.Name,
            "index": new_sheet.Index,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_worksheet(workbook: str = None, name: str = None) -> str:
    """Delete a worksheet from an open workbook.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Sheet to delete (required).

    Returns:
        JSON confirmation. Excel will refuse to delete a workbook's only
        remaining visible sheet.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not name:
        return json.dumps({"error": "name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, name)

        prev_alerts = app.DisplayAlerts
        app.DisplayAlerts = False
        try:
            ws.Delete()
        finally:
            app.DisplayAlerts = prev_alerts

        return json.dumps({"success": True, "workbook": wb.Name, "deleted": name})
    except Exception as e:
        return json.dumps({"error": str(e)})


def rename_worksheet(workbook: str = None, name: str = None, new_name: str = None) -> str:
    """Rename a worksheet in an open workbook.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Current sheet name (None = active sheet).
        new_name: New sheet name (required).

    Returns:
        JSON with the old and new names.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not new_name:
        return json.dumps({"error": "new_name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, name)
        old_name = ws.Name
        ws.Name = new_name

        return json.dumps({"success": True, "workbook": wb.Name, "old_name": old_name, "new_name": ws.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})


def activate_worksheet(workbook: str = None, name: str = None) -> str:
    """Make a worksheet the active tab in an open workbook.

    Args:
        workbook: Workbook name or path (None = active workbook).
        name: Sheet to activate (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not name:
        return json.dumps({"error": "name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, name)
        ws.Activate()

        return json.dumps({"success": True, "workbook": wb.Name, "active_sheet": wb.ActiveSheet.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})
