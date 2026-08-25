"""Data validation, cell merging, and hyperlink tools for live Excel COM automation."""

import json
import sys

_XL_VALIDATE_WHOLE_NUMBER, _XL_VALIDATE_DECIMAL, _XL_VALIDATE_LIST = 1, 2, 3
_XL_VALID_ALERT_STOP = 1
_XL_BETWEEN, _XL_NOT_BETWEEN = 1, 2

_OPERATORS = {
    "greater_than": 5,
    "less_than": 6,
    "greater_equal": 7,
    "less_equal": 8,
    "equal_to": 3,
    "not_equal": 4,
    "between": 1,
    "not_between": 2,
}


def add_dropdown_validation(workbook: str = None, sheet: str = None, range_address: str = "A1", values: list = None) -> str:
    """Restrict a range to a dropdown list of allowed values.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "B2:B20".
        values: List of allowed string values (required), e.g. ["Yes", "No"].

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not values:
        return json.dumps({"error": "values is required (a non-empty list)"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        try:
            rng.Validation.Delete()
        except Exception:
            pass

        formula1 = ",".join(str(v) for v in values)
        rng.Validation.Add(_XL_VALIDATE_LIST, _XL_VALID_ALERT_STOP, _XL_BETWEEN, formula1)

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng), "values": values})
    except Exception as e:
        return json.dumps({"error": str(e)})


def add_number_validation(
    workbook: str = None,
    sheet: str = None,
    range_address: str = "A1",
    condition: str = "greater_than",
    value1: float = None,
    value2: float = None,
    decimal: bool = False,
) -> str:
    """Restrict a range to numbers matching a comparison.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "B2:B20".
        condition: One of "greater_than", "less_than", "greater_equal",
                   "less_equal", "equal_to", "not_equal", "between",
                   "not_between".
        value1: Comparison value (required).
        value2: Second comparison value, required only for "between"/
                "not_between".
        decimal: Allow decimals (False = whole numbers only).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    condition_key = (condition or "").lower()
    if condition_key not in _OPERATORS:
        return json.dumps({"error": f"Invalid condition: {condition}. Use one of {list(_OPERATORS)}"})
    if value1 is None:
        return json.dumps({"error": "value1 is required"})
    if condition_key in ("between", "not_between") and value2 is None:
        return json.dumps({"error": f"value2 is required for condition '{condition_key}'"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        try:
            rng.Validation.Delete()
        except Exception:
            pass

        validate_type = _XL_VALIDATE_DECIMAL if decimal else _XL_VALIDATE_WHOLE_NUMBER
        operator = _OPERATORS[condition_key]

        if value2 is not None:
            rng.Validation.Add(validate_type, _XL_VALID_ALERT_STOP, operator, str(value1), str(value2))
        else:
            rng.Validation.Add(validate_type, _XL_VALID_ALERT_STOP, operator, str(value1))

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_validation(workbook: str = None, sheet: str = None, cell: str = "A1") -> str:
    """Read the data validation rule applied to a cell, if any.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        cell: A1-style cell reference, e.g. "B2".

    Returns:
        JSON with the validation type/operator/formulas, or has_validation:
        false if none is set.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(cell)

        try:
            validation = rng.Validation
            result = {
                "has_validation": True,
                "type": validation.Type,
                "operator": validation.Operator,
                "formula1": validation.Formula1,
            }
            try:
                result["formula2"] = validation.Formula2
            except Exception:
                result["formula2"] = None
        except Exception:
            result = {"has_validation": False}

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "cell": cell, **result}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def remove_validation(workbook: str = None, sheet: str = None, range_address: str = "A1") -> str:
    """Remove data validation from a range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "B2:B20".

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)
        rng.Validation.Delete()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def merge_cells(workbook: str = None, sheet: str = None, range_address: str = "A1") -> str:
    """Merge a range into a single cell.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1:C1".

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)
        rng.Merge()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def unmerge_cells(workbook: str = None, sheet: str = None, range_address: str = "A1") -> str:
    """Undo a cell merge.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address covering the merged area, e.g. "A1:C1".

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)
        rng.UnMerge()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_merge_info(workbook: str = None, sheet: str = None, cell: str = "A1") -> str:
    """Check whether a cell is part of a merged range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        cell: A1-style cell reference, e.g. "B2".

    Returns:
        JSON with is_merged and, if true, the full merged area's address.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(cell)
        is_merged = bool(rng.MergeCells)

        result = {"success": True, "workbook": wb.Name, "sheet": ws.Name, "cell": cell, "is_merged": is_merged}
        if is_merged:
            result["merge_area"] = com_range_address(rng.MergeArea)

        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e)})


def add_hyperlink(
    workbook: str = None,
    sheet: str = None,
    cell: str = "A1",
    address: str = "",
    text_to_display: str = None,
    screen_tip: str = None,
) -> str:
    """Add a hyperlink to a cell.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        cell: A1-style single-cell reference, e.g. "B2".
        address: Target URL (e.g. "https://example.com"), or an internal
                 reference prefixed with '#' (e.g. "#Sheet2!A1").
        text_to_display: Text shown in the cell (None = the address itself).
        screen_tip: Tooltip text shown on hover.

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not address:
        return json.dumps({"error": "address is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(cell)

        if address.startswith("#"):
            sub_address = address[1:]
            display = text_to_display or sub_address
            hl = ws.Hyperlinks.Add(Anchor=rng, Address="", SubAddress=sub_address, TextToDisplay=display)
        else:
            display = text_to_display or address
            hl = ws.Hyperlinks.Add(Anchor=rng, Address=address, TextToDisplay=display)

        if screen_tip:
            hl.ScreenTip = screen_tip

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "cell": cell, "address": address})
    except Exception as e:
        return json.dumps({"error": str(e)})


def remove_hyperlink(workbook: str = None, sheet: str = None, cell: str = "A1") -> str:
    """Remove the hyperlink from a cell, if any.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        cell: A1-style single-cell reference, e.g. "B2".

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(cell)
        rng.Hyperlinks.Delete()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "cell": cell})
    except Exception as e:
        return json.dumps({"error": str(e)})


def list_hyperlinks(workbook: str = None, sheet: str = None) -> str:
    """List all hyperlinks on a worksheet.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).

    Returns:
        JSON with each hyperlink's cell, address, and display text.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)

        links = []
        for i in range(1, ws.Hyperlinks.Count + 1):
            hl = ws.Hyperlinks(i)
            links.append({
                "cell": com_range_address(hl.Range),
                "address": hl.Address,
                "sub_address": hl.SubAddress,
                "text_to_display": hl.TextToDisplay,
            })

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "hyperlinks": links}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})
