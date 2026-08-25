"""COM connection manager for Microsoft Excel on Windows.

Provides functions to attach to a running Excel instance and locate open
workbooks/worksheets. Only works on Windows with pywin32 installed.

Unlike mcp-server-excel (the sibling project), this module never launches or
quits an Excel.Application — it only attaches to processes the user already
started. Every lookup here must remain read-only with respect to process
lifetime: no Quit(), no Close() on a workbook the caller didn't explicitly
ask to close.
"""

import os
import sys
import unicodedata


def _find_active_excel():
    """Return the Excel.Application from GetActiveObject, or None."""
    import win32com.client

    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return None


def list_all_open_workbooks():
    """Enumerate every open workbook across every running Excel instance.

    Windows commonly hosts multiple separate Excel.Application processes at
    once (unlike Word, which normally has one). This walks the full COM
    Running Object Table (ROT) collecting every distinct Application found
    there, rather than trusting a single GetActiveObject() call to find the
    one the caller actually wants.

    Returns:
        A list of (app, workbook) COM object pairs.

    Raises:
        RuntimeError: If no Excel process could be found at all, or if not
                      running on Windows.
    """
    if sys.platform != "win32":
        raise RuntimeError("Excel COM automation is only available on Windows")

    import pythoncom
    import win32com.client

    # COM is apartment-threaded: each OS thread must initialize its own
    # apartment before touching COM objects. FastMCP may dispatch this call
    # on a fresh worker thread per request, so this can't be a one-time
    # process-level init — it must run on every call. CoInitialize() is a
    # no-op (returns S_FALSE) if the thread is already initialized, so this
    # is safe to call unconditionally.
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    apps = []
    seen_hwnds = set()

    def _try_add(app):
        if app is None:
            return
        try:
            hwnd = app.Hwnd
        except Exception:
            return
        if hwnd in seen_hwnds:
            return
        seen_hwnds.add(hwnd)
        apps.append(app)

    _try_add(_find_active_excel())

    try:
        rot = pythoncom.GetRunningObjectTable(0)
        enum = rot.EnumRunning()
        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                ctx = pythoncom.CreateBindCtx(0)
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                com_obj = win32com.client.Dispatch(dispatch)
                if hasattr(com_obj, "Workbooks") and hasattr(com_obj, "ActiveWorkbook"):
                    # Direct Application entry
                    _try_add(com_obj)
                elif hasattr(com_obj, "Application") and hasattr(com_obj, "FullName"):
                    # A Workbook moniker — reach its owning Application
                    _try_add(com_obj.Application)
            except Exception:
                continue
    except Exception:
        pass

    if not apps:
        raise RuntimeError(
            "Microsoft Excel is not running. Please open Excel and a workbook first."
        )

    pairs = []
    for app in apps:
        try:
            for i in range(1, app.Workbooks.Count + 1):
                pairs.append((app, app.Workbooks(i)))
        except Exception:
            continue

    return pairs


def find_workbook(name: str = None):
    """Find an open workbook by name/path across all running Excel instances.

    Args:
        name: Workbook name (basename) or full path. If None, returns the
              active workbook of the first Excel instance found.

    Returns:
        (app, workbook) — the owning Excel.Application COM object and the
        Workbook COM object.

    Raises:
        ValueError: If no open workbook matches, or none are open at all.
        RuntimeError: If no Excel process is running, or not on Windows.
    """
    pairs = list_all_open_workbooks()

    if not pairs:
        raise ValueError("No workbooks are open in Excel. Please open a workbook first.")

    if not name:
        app = pairs[0][0]
        try:
            active = app.ActiveWorkbook
        except Exception:
            active = None
        if active is not None:
            return app, active
        return pairs[0]

    target_basename = unicodedata.normalize("NFC", os.path.basename(name)).lower()
    target_fullpath = (
        unicodedata.normalize("NFC", os.path.normpath(name)).lower()
        if os.path.isabs(name)
        else None
    )

    for app, wb in pairs:
        if unicodedata.normalize("NFC", wb.Name).lower() == target_basename:
            return app, wb
        if target_fullpath and unicodedata.normalize("NFC", os.path.normpath(wb.FullName)).lower() == target_fullpath:
            return app, wb

    open_books = [wb.Name for _, wb in pairs]
    raise ValueError(f"Workbook '{name}' is not open in Excel. Open workbooks: {open_books}")


def find_worksheet(workbook, name: str = None):
    """Find a worksheet by name in an open workbook, or return the active sheet.

    Args:
        workbook: Workbook COM object (from find_workbook()).
        name: Worksheet name. If None, returns the workbook's active sheet.

    Returns:
        Worksheet COM object.

    Raises:
        ValueError: If no worksheet matches the given name.
    """
    if not name:
        return workbook.ActiveSheet

    for i in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets(i)
        if sheet.Name == name:
            return sheet

    sheet_names = [workbook.Worksheets(i).Name for i in range(1, workbook.Worksheets.Count + 1)]
    raise ValueError(
        f"Worksheet '{name}' not found in '{workbook.Name}'. Available sheets: {sheet_names}"
    )


def com_range_address(rng) -> str:
    """Return a relative A1-style address for a Range (e.g. "A1:C10", no '$').

    Named com_range_address (not range_address) because every tool function
    in tools/range_tools.py etc. already has a `range_address` parameter (an
    A1 string like "A1:C10") — importing a same-named function would be
    shadowed by that local parameter inside those function bodies.

    Range.Address is a COM property that optionally takes RowAbsolute/
    ColumnAbsolute arguments. Whether it's callable from Python depends on
    whether pywin32 is using early-bound (gencache) or late-bound dispatch
    for Excel on the current machine: early-bound resolves it eagerly to a
    plain string using the default (absolute) arguments, so calling it like
    ``rng.Address(False, False)`` raises "'str' object is not callable" in
    that mode. This works either way — always use it instead of calling
    ``.Address(...)`` directly.
    """
    addr = rng.Address
    if callable(addr):
        return addr(False, False)
    return addr.replace("$", "")
