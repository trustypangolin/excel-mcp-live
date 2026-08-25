"""Conditional formatting tools for live Excel COM automation.

Covers cell-value comparison rules and color scales — the most commonly
requested conditional formatting. Not yet covered: data bars, icon sets,
and formula/expression-based rules; use mcp-server-excel for those.
"""

import json
import sys

_XL_CELL_VALUE = 1

_OPERATORS = {
    "greater_than": 5,   # xlGreater
    "less_than": 6,      # xlLess
    "greater_equal": 7,  # xlGreaterEqual
    "less_equal": 8,     # xlLessEqual
    "equal_to": 3,       # xlEqual
    "not_equal": 4,      # xlNotEqual
    "between": 1,        # xlBetween
    "not_between": 2,    # xlNotBetween
}


def add_rule(
    workbook: str = None,
    sheet: str = None,
    range_address: str = "A1",
    condition: str = "greater_than",
    value1=None,
    value2=None,
    fill_color: str = None,
    font_color: str = None,
    bold: bool = None,
) -> str:
    """Add a cell-value conditional formatting rule to a range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1:A20".
        condition: One of "greater_than", "less_than", "greater_equal",
                   "less_equal", "equal_to", "not_equal", "between",
                   "not_between".
        value1: Comparison value (required). Numbers or text.
        value2: Second comparison value, required only for "between"/
                "not_between".
        fill_color: Hex RGB cell fill to apply when the rule matches.
        font_color: Hex RGB font color to apply when the rule matches.
        bold: Bold the font when the rule matches.

    Returns:
        JSON with the new rule's index (rules apply in index order).
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
        from excel_document_server.utils.conversions import hex_to_bgr

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        formula1 = str(value1)
        operator = _OPERATORS[condition_key]

        if value2 is not None:
            fc = rng.FormatConditions.Add(_XL_CELL_VALUE, operator, formula1, str(value2))
        else:
            fc = rng.FormatConditions.Add(_XL_CELL_VALUE, operator, formula1)

        if fill_color is not None:
            fc.Interior.Color = hex_to_bgr(fill_color)
        if font_color is not None:
            fc.Font.Color = hex_to_bgr(font_color)
        if bold is not None:
            fc.Font.Bold = bold

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "range": com_range_address(rng),
            "rule_index": rng.FormatConditions.Count,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def add_color_scale(
    workbook: str = None,
    sheet: str = None,
    range_address: str = "A1",
    min_color: str = "F8696B",
    max_color: str = "63BE7B",
    mid_color: str = None,
) -> str:
    """Add a 2-color or 3-color scale to a range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1:A20".
        min_color: Hex RGB for the lowest value (default: a red).
        max_color: Hex RGB for the highest value (default: a green).
        mid_color: Hex RGB for the midpoint (None = 2-color scale; set
                   this for a 3-color scale, e.g. yellow "FFEB84").

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address
        from excel_document_server.utils.conversions import hex_to_bgr

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        scale_type = 3 if mid_color is not None else 2
        cs = rng.FormatConditions.AddColorScale(scale_type)
        cs.ColorScaleCriteria(1).FormatColor.Color = hex_to_bgr(min_color)
        if mid_color is not None:
            cs.ColorScaleCriteria(2).FormatColor.Color = hex_to_bgr(mid_color)
        cs.ColorScaleCriteria(scale_type).FormatColor.Color = hex_to_bgr(max_color)

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "range": com_range_address(rng),
            "scale_type": scale_type,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def list_rules(workbook: str = None, sheet: str = None, range_address: str = "A1") -> str:
    """List conditional formatting rules applied to a range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1:A20".

    Returns:
        JSON with each rule's index, COM Type/Operator codes, and formulas
        (cell-value rules only report formulas; color scales report just
        the type).
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        rules = []
        for i in range(1, rng.FormatConditions.Count + 1):
            fc = rng.FormatConditions(i)
            rule = {"index": i, "type": fc.Type}
            if fc.Type == _XL_CELL_VALUE:
                rule["operator"] = fc.Operator
                rule["formula1"] = fc.Formula1
                try:
                    rule["formula2"] = fc.Formula2
                except Exception:
                    rule["formula2"] = None
            rules.append(rule)

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "range": com_range_address(rng),
            "rules": rules,
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def clear_rules(workbook: str = None, sheet: str = None, range_address: str = "A1") -> str:
    """Remove all conditional formatting rules from a range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1:A20".

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
        rng.FormatConditions.Delete()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})
