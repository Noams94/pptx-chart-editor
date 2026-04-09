"""Extract chart data from PPTX files into pandas DataFrames."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from io import BytesIO

import pandas as pd
from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn


# Chart types that use XyChartData (scatter plots)
XY_CHART_TYPES = {
    XL_CHART_TYPE.XY_SCATTER,
    XL_CHART_TYPE.XY_SCATTER_LINES,
    XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
}


def _is_percentage_format(fmt: str) -> bool:
    """Check if a format code represents percentages (e.g., '0%', '0.0%', '#,##0%')."""
    if not fmt:
        return False
    # Remove escaped characters and quoted strings
    cleaned = re.sub(r'"[^"]*"', '', fmt)
    cleaned = re.sub(r'\\\.', '', cleaned)
    return '%' in cleaned


def _extract_series_formats(chart: Chart) -> dict[str, str]:
    """Extract number format codes per series from chart XML.

    Returns dict mapping series name -> formatCode (e.g., '0%', '#,##0', 'General')
    """
    formats = {}
    chart_xml = chart.part._element

    for ser in chart_xml.iter(qn('c:ser')):
        # Get series name
        name = None
        tx = ser.find(qn('c:tx'))
        if tx is not None:
            str_ref = tx.find(qn('c:strRef'))
            if str_ref is not None:
                str_cache = str_ref.find(qn('c:strCache'))
                if str_cache is not None:
                    pt = str_cache.find(qn('c:pt'))
                    if pt is not None:
                        v = pt.find(qn('c:v'))
                        if v is not None:
                            name = v.text

        # Get formatCode from val > numRef > numCache > formatCode
        fmt_code = None
        val = ser.find(qn('c:val'))
        if val is not None:
            num_ref = val.find(qn('c:numRef'))
            if num_ref is not None:
                num_cache = num_ref.find(qn('c:numCache'))
                if num_cache is not None:
                    fc = num_cache.find(qn('c:formatCode'))
                    if fc is not None and fc.text:
                        fmt_code = fc.text

        if name and fmt_code:
            formats[name] = fmt_code

    return formats


@dataclass
class ChartInfo:
    slide_index: int
    shape_name: str
    chart_type: int
    chart_type_name: str
    dataframe: pd.DataFrame          # Display values (67 for 67%)
    raw_dataframe: pd.DataFrame       # Raw values (0.67 for 67%)
    is_xy: bool = False
    series_names: list = field(default_factory=list)
    series_formats: dict = field(default_factory=dict)  # series_name -> formatCode


def _extract_chart_data(chart: Chart) -> tuple[pd.DataFrame, pd.DataFrame, bool, list[str], dict]:
    """Extract data from a chart into DataFrames (display + raw)."""
    chart_type = chart.chart_type
    is_xy = chart_type in XY_CHART_TYPES

    plot = chart.plots[0]
    series_list = list(plot.series)
    series_names = [s.name if s.name else f"סדרה {i+1}" for i, s in enumerate(series_list)]

    # Extract number formats
    series_formats = _extract_series_formats(chart)

    if is_xy:
        data = {}
        for i, series in enumerate(series_list):
            x_vals = list(series.values)
            y_vals = list(series.values)
            data[f"X_{series_names[i]}"] = x_vals
            data[f"Y_{series_names[i]}"] = y_vals
        raw_df = pd.DataFrame(data)
        display_df = raw_df.copy()
    else:
        try:
            categories = [str(c) for c in plot.categories]
        except Exception:
            categories = [str(i + 1) for i in range(len(list(series_list[0].values)))]

        raw_data = {"קטגוריה": categories}
        display_data = {"קטגוריה": categories}

        for i, series in enumerate(series_list):
            values = list(series.values)
            while len(values) < len(categories):
                values.append(None)
            values = values[:len(categories)]

            name = series_names[i]
            raw_data[name] = values

            # Convert to display values based on format
            fmt = series_formats.get(name, "General")
            if _is_percentage_format(fmt):
                display_values = [
                    round(v * 100, 2) if v is not None else None
                    for v in values
                ]
                display_data[name] = display_values
            else:
                display_data[name] = values

        raw_df = pd.DataFrame(raw_data)
        display_df = pd.DataFrame(display_data)

    return display_df, raw_df, is_xy, series_names, series_formats


def _chart_type_display_name(chart_type: int) -> str:
    """Get a Hebrew display name for the chart type."""
    names = {
        XL_CHART_TYPE.COLUMN_CLUSTERED: "עמודות מקובצות",
        XL_CHART_TYPE.COLUMN_STACKED: "עמודות מוערמות",
        XL_CHART_TYPE.COLUMN_STACKED_100: "עמודות מוערמות 100%",
        XL_CHART_TYPE.BAR_CLUSTERED: "מוטות מקובצות",
        XL_CHART_TYPE.BAR_STACKED: "מוטות מוערמות",
        XL_CHART_TYPE.BAR_STACKED_100: "מוטות מוערמות 100%",
        XL_CHART_TYPE.LINE: "קו",
        XL_CHART_TYPE.LINE_MARKERS: "קו עם סמנים",
        XL_CHART_TYPE.LINE_STACKED: "קו מוערם",
        XL_CHART_TYPE.PIE: "עוגה",
        XL_CHART_TYPE.PIE_EXPLODED: "עוגה מפוצלת",
        XL_CHART_TYPE.DOUGHNUT: "סופגנייה",
        XL_CHART_TYPE.AREA: "שטח",
        XL_CHART_TYPE.AREA_STACKED: "שטח מוערם",
        XL_CHART_TYPE.XY_SCATTER: "פיזור",
    }
    return names.get(chart_type, "גרף")


def extract_all_charts(pptx_bytes: bytes) -> list[ChartInfo]:
    """Extract all charts from a PPTX file."""
    prs = Presentation(BytesIO(pptx_bytes))
    charts = []

    for slide_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.has_chart:
                continue

            chart = shape.chart
            try:
                display_df, raw_df, is_xy, series_names, series_formats = _extract_chart_data(chart)
                info = ChartInfo(
                    slide_index=slide_idx,
                    shape_name=shape.name,
                    chart_type=chart.chart_type,
                    chart_type_name=_chart_type_display_name(chart.chart_type),
                    dataframe=display_df,
                    raw_dataframe=raw_df,
                    is_xy=is_xy,
                    series_names=series_names,
                    series_formats=series_formats,
                )
                charts.append(info)
            except Exception as e:
                print(f"Warning: Could not extract chart '{shape.name}' on slide {slide_idx + 1}: {e}")

    return charts
