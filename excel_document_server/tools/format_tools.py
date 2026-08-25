"""Cell formatting tools for live Excel COM automation."""

import json
import sys

_HALIGN = {
    "left": -4131,
    "center": -4108,
    "right": -4152,
    "general": 1,
}


def format_range(
    workbook: str = None,
    sheet: str = None,
    range_address: str = "A1",
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    fill_color: str = None,
    number_format: str = None,
    horizontal_align: str = None,
    wrap_text: bool = None,
) -> str:
    """Apply font, fill, number-format, and alignment to a cell or range.

    Only arguments that are explicitly set are changed; everything else is
    left as-is.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1:C1".
        bold, italic, underline: Font style toggles.
        font_name: e.g. "Calibri".
        font_size: Point size.
        font_color: Hex RGB, e.g. "FF0000" (with or without leading '#').
        fill_color: Hex RGB cell background color.
        number_format: Excel number format code, e.g. "$#,##0.00" or "0.0%".
        horizontal_align: One of "left", "center", "right", "general".
        wrap_text: Whether to wrap text within the cell.

    Returns:
        JSON confirmation of the range formatted.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address
        from excel_document_server.utils.conversions import hex_to_bgr

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        if bold is not None:
            rng.Font.Bold = bold
        if italic is not None:
            rng.Font.Italic = italic
        if underline is not None:
            rng.Font.Underline = 2 if underline else -4142  # xlUnderlineStyleSingle / xlUnderlineStyleNone
        if font_name is not None:
            rng.Font.Name = font_name
        if font_size is not None:
            rng.Font.Size = font_size
        if font_color is not None:
            rng.Font.Color = hex_to_bgr(font_color)
        if fill_color is not None:
            rng.Interior.Color = hex_to_bgr(fill_color)
        if number_format is not None:
            rng.NumberFormat = number_format
        if horizontal_align is not None:
            key = horizontal_align.lower()
            if key not in _HALIGN:
                return json.dumps({"error": f"Invalid horizontal_align: {horizontal_align}. Use one of {list(_HALIGN)}"})
            rng.HorizontalAlignment = _HALIGN[key]
        if wrap_text is not None:
            rng.WrapText = wrap_text

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def autofit_columns(workbook: str = None, sheet: str = None, range_address: str = None) -> str:
    """Autofit column widths for a range, or the sheet's used range if none given.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address (None = the sheet's used range).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address) if range_address else ws.UsedRange
        rng.Columns.AutoFit()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "range": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})
