"""Cell and range read/write tools for live Excel COM automation."""

import json
import sys


def get_range_values(workbook: str = None, sheet: str = None, range_address: str = "A1") -> str:
    """Read values from a cell or range in an open workbook.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1" or "A1:C10".

    Returns:
        JSON with the values as a 2D array (even for a single cell).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address
        from excel_document_server.utils.conversions import to_jsonable

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        raw = rng.Value
        if raw is None:
            values = [[None]]
        elif isinstance(raw, tuple):
            values = to_jsonable(raw)
        else:
            values = [[to_jsonable(raw)]]

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "range": com_range_address(rng),
            "values": values,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_range_values(
    workbook: str = None,
    sheet: str = None,
    range_address: str = "A1",
    values: list | str | float | int | bool | None = None,
) -> str:
    """Write values into a cell or range in an open workbook.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: The top-left cell to write into, e.g. "A1" or "B2".
        values: A single value, a flat list (written as one row), or a list
                of lists (2D block) anchored at range_address's top-left cell.

    Returns:
        JSON with the range actually written to.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if values is None:
        return json.dumps({"error": "values is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address
        from excel_document_server.utils.conversions import column_letter

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        anchor = ws.Range(range_address)
        start_row, start_col = anchor.Row, anchor.Column

        if isinstance(values, list) and values and isinstance(values[0], list):
            rows, cols = len(values), max(len(r) for r in values)
            end_row, end_col = start_row + rows - 1, start_col + cols - 1
            target = ws.Range(f"{column_letter(start_col)}{start_row}:{column_letter(end_col)}{end_row}")
            target.Value = [tuple(row) for row in values]
        elif isinstance(values, list):
            end_col = start_col + len(values) - 1
            target = ws.Range(f"{column_letter(start_col)}{start_row}:{column_letter(end_col)}{start_row}")
            target.Value = [tuple(values)]
        else:
            target = anchor
            target.Value = values

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "range": com_range_address(target),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_range_formulas(workbook: str = None, sheet: str = None, range_address: str = "A1") -> str:
    """Read formulas from a cell or range.

    Formulas are returned A1-style (e.g. "=SUM(A1:A10)"). Cells without a
    formula return their literal value instead.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1" or "A1:C10".

    Returns:
        JSON with the formulas/values as a 2D array.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address
        from excel_document_server.utils.conversions import to_jsonable

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        raw = rng.Formula
        formulas = to_jsonable(raw) if isinstance(raw, tuple) else [[to_jsonable(raw)]]

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "range": com_range_address(rng),
            "formulas": formulas,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_cell_formula(workbook: str = None, sheet: str = None, cell: str = "A1", formula: str = "") -> str:
    """Set a formula (or literal value) on a single cell.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        cell: A1-style cell reference, e.g. "B2".
        formula: Formula string starting with '=' (e.g. "=SUM(A1:A10)"), or
                 a literal value.

    Returns:
        JSON with the cell's resulting calculated value.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address
        from excel_document_server.utils.conversions import to_jsonable

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(cell)
        rng.Formula = formula

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "cell": com_range_address(rng),
            "value": to_jsonable(rng.Value),
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def find_replace(
    workbook: str = None,
    sheet: str = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = False,
    whole_cell: bool = False,
) -> str:
    """Find and replace text in an open workbook.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = every sheet in the workbook).
        find_text: Text to search for (required).
        replace_text: Replacement text.
        match_case: Case-sensitive match.
        whole_cell: Match the entire cell contents rather than a substring.

    Returns:
        JSON with which sheets had at least one replacement.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not find_text:
        return json.dumps({"error": "find_text is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        sheets = (
            [find_worksheet(wb, sheet)]
            if sheet
            else [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
        )

        XL_PART, XL_WHOLE = 2, 1
        results = []
        for ws in sheets:
            changed = bool(ws.UsedRange.Replace(
                What=find_text,
                Replacement=replace_text,
                LookAt=XL_WHOLE if whole_cell else XL_PART,
                MatchCase=match_case,
            ))
            results.append({"sheet": ws.Name, "changed": changed})

        return json.dumps({"success": True, "workbook": wb.Name, "sheets": results})
    except Exception as e:
        return json.dumps({"error": str(e)})


def clear_range(workbook: str = None, sheet: str = None, range_address: str = "A1", clear_formatting: bool = False) -> str:
    """Clear cell contents from a range, optionally including formatting.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1:C10".
        clear_formatting: Also reset fonts, fills, and number formats
                           (Range.Clear instead of Range.ClearContents).

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

        if clear_formatting:
            rng.Clear()
        else:
            rng.ClearContents()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})
