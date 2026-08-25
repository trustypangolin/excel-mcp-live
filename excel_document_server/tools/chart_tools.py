"""Chart tools for live Excel COM automation.

A lean subset: create, list, retype, retitle, add a series, and delete.
Not yet covered: axes, legends, data labels, trendlines, PivotCharts —
use mcp-server-excel for those.
"""

import json
import sys

_CHART_TYPES = {
    "column_clustered": 51,  # xlColumnClustered
    "bar_clustered": 57,     # xlBarClustered
    "line": 4,               # xlLine
    "pie": 5,                # xlPie
    "area": 1,               # xlArea
    "scatter": -4169,        # xlXYScatter
    "doughnut": -4120,       # xlDoughnut
}


def _find_chart(ws, name: str):
    """Find a ChartObject by name on a worksheet. Raises ValueError if not found."""
    try:
        return ws.ChartObjects(name)
    except Exception:
        names = [ws.ChartObjects(i).Name for i in range(1, ws.ChartObjects().Count + 1)]
        raise ValueError(f"Chart '{name}' not found. Available charts: {names}")


def create_chart(
    workbook: str = None,
    sheet: str = None,
    range_address: str = "A1",
    chart_type: str = "column_clustered",
    title: str = None,
    chart_name: str = None,
    left: float = 400,
    top: float = 50,
    width: float = 400,
    height: float = 250,
) -> str:
    """Create a chart from a range.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        range_address: Source data range, e.g. "A1:C10".
        chart_type: One of "column_clustered", "bar_clustered", "line",
                    "pie", "area", "scatter", "doughnut".
        title: Chart title (None = no title).
        chart_name: Name for the new chart (None = Excel's default, e.g. "Chart 1").
        left, top, width, height: Position and size in points.

    Returns:
        JSON with the new chart's name.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    type_key = (chart_type or "").lower()
    if type_key not in _CHART_TYPES:
        return json.dumps({"error": f"Invalid chart_type: {chart_type}. Use one of {list(_CHART_TYPES)}"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        rng = ws.Range(range_address)

        chart_obj = ws.ChartObjects().Add(left, top, width, height)
        chart_obj.Chart.SetSourceData(Source=rng)
        chart_obj.Chart.ChartType = _CHART_TYPES[type_key]

        if chart_name:
            chart_obj.Name = chart_name
        if title:
            chart_obj.Chart.HasTitle = True
            chart_obj.Chart.ChartTitle.Text = title

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "chart_name": chart_obj.Name})
    except Exception as e:
        return json.dumps({"error": str(e)})


def list_charts(workbook: str = None, sheet: str = None) -> str:
    """List charts on a worksheet.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).

    Returns:
        JSON with each chart's name and position/size.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)

        charts = []
        for i in range(1, ws.ChartObjects().Count + 1):
            co = ws.ChartObjects(i)
            charts.append({"name": co.Name, "left": co.Left, "top": co.Top, "width": co.Width, "height": co.Height})

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "charts": charts}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_chart_type(workbook: str = None, sheet: str = None, chart_name: str = None, chart_type: str = "column_clustered") -> str:
    """Change a chart's type.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        chart_name: Chart to change (required).
        chart_type: One of "column_clustered", "bar_clustered", "line",
                    "pie", "area", "scatter", "doughnut".

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not chart_name:
        return json.dumps({"error": "chart_name is required"})

    type_key = (chart_type or "").lower()
    if type_key not in _CHART_TYPES:
        return json.dumps({"error": f"Invalid chart_type: {chart_type}. Use one of {list(_CHART_TYPES)}"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        chart_obj = _find_chart(ws, chart_name)
        chart_obj.Chart.ChartType = _CHART_TYPES[type_key]

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "chart_name": chart_obj.Name, "chart_type": type_key})
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def set_chart_title(workbook: str = None, sheet: str = None, chart_name: str = None, title: str = None) -> str:
    """Set or clear a chart's title.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        chart_name: Chart to update (required).
        title: New title text. Empty string or None removes the title.

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not chart_name:
        return json.dumps({"error": "chart_name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        chart_obj = _find_chart(ws, chart_name)

        if title:
            chart_obj.Chart.HasTitle = True
            chart_obj.Chart.ChartTitle.Text = title
        else:
            chart_obj.Chart.HasTitle = False

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "chart_name": chart_obj.Name, "title": title})
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def add_series(workbook: str = None, sheet: str = None, chart_name: str = None, series_range: str = None, series_name: str = None) -> str:
    """Add a data series to a chart.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        chart_name: Chart to add the series to (required).
        series_range: A1-style address of the series' values, e.g. "D2:D10" (required).
        series_name: Legend label for the series (None = Excel's default).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not chart_name:
        return json.dumps({"error": "chart_name is required"})
    if not series_range:
        return json.dumps({"error": "series_range is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        chart_obj = _find_chart(ws, chart_name)
        rng = ws.Range(series_range)

        new_series = chart_obj.Chart.SeriesCollection().NewSeries()
        new_series.Values = rng
        if series_name:
            new_series.Name = series_name

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "chart_name": chart_obj.Name})
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})


def delete_chart(workbook: str = None, sheet: str = None, chart_name: str = None) -> str:
    """Delete a chart.

    Args:
        workbook: Workbook name or path (None = active workbook).
        sheet: Worksheet name (None = active sheet).
        chart_name: Chart to delete (required).

    Returns:
        JSON confirmation.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})
    if not chart_name:
        return json.dumps({"error": "chart_name is required"})

    try:
        from excel_document_server.core.excel_com import find_workbook, find_worksheet

        _app, wb = find_workbook(workbook)
        ws = find_worksheet(wb, sheet)
        chart_obj = _find_chart(ws, chart_name)
        chart_obj.Delete()

        return json.dumps({"success": True, "workbook": wb.Name, "sheet": ws.Name, "deleted": chart_name})
    except ValueError as e:
        return json.dumps({"error": str(e)})
    except Exception as e:
        return json.dumps({"error": str(e)})
