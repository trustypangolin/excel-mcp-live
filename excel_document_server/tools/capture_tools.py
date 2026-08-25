"""Range/sheet screenshot tools for live Excel COM automation.

Excel has no direct "export range to image file" API. The standard
workaround (used throughout VBA automation) is: copy the range as a
picture, paste it onto a temporary chart object sized to match, export
that chart as an image, then delete the temporary chart.
"""

import json
import os
import sys

_XL_SCREEN = 1
_XL_BITMAP = 2


def _export_range_as_image(ws, rng, output_path: str, image_format: str):
    """Shared implementation for capturing a Range to an image file.

    Excel COM quirks hit while building this (see CLAUDE.md for the full
    history): Appearance=xlPrinter raises outright; calling ws.Activate()
    right before CopyPicture(xlScreen) also raises, even with the message
    queue pumped afterward. Plain xlScreen with no Activate() is the only
    combination that doesn't raise — but Chart.Paste() doesn't reliably
    accept CopyPicture's clipboard format directly either: it silently
    produces an empty chart (no error, but a blank exported image).
    """
    rng.CopyPicture(Appearance=_XL_SCREEN, Format=_XL_BITMAP)

    # Paste onto the worksheet first to materialize a genuine Picture
    # shape, then re-copy that shape — its clipboard format is a plain
    # "Picture" that Chart.Paste() does accept. The temporary worksheet
    # picture is deleted afterward.
    ws.Paste()
    temp_shape = ws.Shapes(ws.Shapes.Count)
    temp_shape.Copy()

    chart_obj = ws.ChartObjects().Add(rng.Left, rng.Top + rng.Height + 40, rng.Width, rng.Height)
    try:
        chart_obj.Chart.Paste()
        chart_obj.Chart.Export(output_path, image_format)
    finally:
        chart_obj.Delete()
        temp_shape.Delete()
        chart_obj.Delete()


def capture_range(workbook: str = None, sheet: str = None, range_address: str = "A1", output_path: str = None) -> str:
    """Capture a range as an image file, exactly as Excel displays it
    (formatting, conditional formatting, charts included).

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: A1-style address, e.g. "A1:F20".
        output_path: Destination file path (required). Extension selects
                     the format: .png, .jpg, .gif, or .bmp.

    Returns:
        JSON with the exported file's absolute path. A vision-capable
        assistant can read the image back to check the result.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not output_path:
        return json.dumps({"error": "output_path is required"})

    ext = os.path.splitext(output_path)[1].lstrip(".").upper()
    if ext not in ("PNG", "JPG", "GIF", "BMP"):
        return json.dumps({"error": f"Unsupported image extension '.{ext}'. Use .png, .jpg, .gif, or .bmp"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        full_output = os.path.abspath(output_path)
        directory = os.path.dirname(full_output)
        if directory and not os.path.isdir(directory):
            return json.dumps({"error": f"Output directory does not exist: {directory}"})

        _export_range_as_image(ws, rng, full_output, ext)

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "range": com_range_address(rng),
            "output_path": full_output,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


def capture_sheet(workbook: str = None, sheet: str = None, output_path: str = None) -> str:
    """Capture a worksheet's entire used area as an image file.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        output_path: Destination file path (required). Extension selects
                     the format: .png, .jpg, .gif, or .bmp.

    Returns:
        JSON with the exported file's absolute path.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not output_path:
        return json.dumps({"error": "output_path is required"})

    ext = os.path.splitext(output_path)[1].lstrip(".").upper()
    if ext not in ("PNG", "JPG", "GIF", "BMP"):
        return json.dumps({"error": f"Unsupported image extension '.{ext}'. Use .png, .jpg, .gif, or .bmp"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.UsedRange

        full_output = os.path.abspath(output_path)
        directory = os.path.dirname(full_output)
        if directory and not os.path.isdir(directory):
            return json.dumps({"error": f"Output directory does not exist: {directory}"})

        _export_range_as_image(ws, rng, full_output, ext)

        return json.dumps({
            "success": True,
            "workbook": wb.Name,
            "sheet": ws.Name,
            "range": com_range_address(rng),
            "output_path": full_output,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
