"""Workbook-level tools for live Excel COM automation.

These tools operate on workbooks that are currently open in Excel.
"""

import json
import sys


def list_open_workbooks() -> str:
    """List all workbooks currently open, across every running Excel instance.

    Returns:
        JSON with each workbook's name, full path, save state, and which
        Excel window (by Hwnd) it belongs to.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import list_all_open_workbooks

        pairs = list_all_open_workbooks()
        workbooks = []
        for app, wb in pairs:
            try:
                active_name = app.ActiveWorkbook.Name if app.ActiveWorkbook else None
            except Exception:
                active_name = None
            workbooks.append({
                "name": wb.Name,
                "full_path": wb.FullName,
                "saved": bool(wb.Saved),
                "excel_window": app.Hwnd,
                "is_active_in_its_window": wb.Name == active_name,
            })

        return json.dumps({"success": True, "workbooks": workbooks}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_workbook_info(workbook: str = None) -> str:
    """Get worksheet names, the active sheet, and each sheet's used range.

    Args:
        workbook: Workbook name or path (None = the active workbook of the
                  first Excel instance found).

    Returns:
        JSON with sheet names, the active sheet, and each sheet's used
        range and visibility.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, com_range_address

        _app, wb = find_workbook(workbook)

        sheets = []
        for i in range(1, wb.Worksheets.Count + 1):
            sheet = wb.Worksheets(i)
            used = sheet.UsedRange
            sheets.append({
                "name": sheet.Name,
                "index": i,
                "used_range": com_range_address(used),
                "visible": sheet.Visible == -1,
            })

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "full_path": wb.FullName,
            "active_sheet": wb.ActiveSheet.Name,
            "sheets": sheets,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def save_workbook(workbook: str = None) -> str:
    """Save an open workbook in place (equivalent to pressing Ctrl+S).

    Args:
        workbook: Workbook name or path (None = the active workbook of the
                  first Excel instance found).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook

        _app, wb = find_workbook(workbook)
        wb.Save()

        return json.dumps({"success": True, "workbook": wb.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})
