"""Cell comment/note tools for live Excel COM automation."""

import json
import sys

from excel_document_server.defaults import DEFAULT_AUTHOR


def add_comment(workbook: str = None, sheet: str = None, cell: str = "A1", text: str = "", author: str = None) -> str:
    """Add a comment (note) to a cell in an open workbook.

    Replaces any existing comment on the cell.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        cell: A1-style single-cell reference, e.g. "B2".
        text: Comment text (required).
        author: Comment author (default: MCP_AUTHOR env var, or "Author").

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not text:
        return json.dumps({"error": "text is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(cell)

        if rng.Comment is not None:
            rng.Comment.Delete()

        author_name = author or DEFAULT_AUTHOR
        comment = rng.AddComment(f"{author_name}:\n{text}")
        comment.Visible = False

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "cell": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def get_comments(workbook: str = None, sheet: str = None) -> str:
    """List all comments on a worksheet.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).

    Returns:
        JSON with each comment's cell address and text.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)

        comments = []
        for i in range(1, ws.Comments.Count + 1):
            com = ws.Comments(i)
            # Comment.Text is the same kind of parameterized COM property as
            # Range.Address — callable under late-bound dispatch, a plain
            # string under early-bound (gencache) dispatch.
            comment_text = com.Text() if callable(com.Text) else com.Text
            comments.append({
                "cell": com_range_address(com.Parent),
                "text": comment_text,
            })

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "comments": comments}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_comment(workbook: str = None, sheet: str = None, cell: str = "A1") -> str:
    """Delete the comment on a specific cell, if any.

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
        from excel_document_server.core.excel_com import find_workbook, find_worksheet, com_range_address

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(cell)

        if rng.Comment is None:
            return json.dumps({
                "success": True,
                "workbook": wb.Name,
                "sheet": ws.Name,
                "cell": com_range_address(rng),
                "message": "No comment to delete",
            })

        rng.Comment.Delete()
        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "cell": com_range_address(rng)})
    except Exception as e:
        return json.dumps({"error": str(e)})
