"""Extract chart data from PPTX files into pandas DataFrames."""

from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO

import pandas as pd
from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.enum.chart import XL_CHART_TYPE


# Chart types that use XyChartData (scatter plots)
XY_CHART_TYPES = {
    XL_CHART_TYPE.XY_SCATTER,
    XL_CHART_TYPE.XY_SCATTER_LINES,
    XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
}


@dataclass
class ChartInfo:
    slide_index: int
    shape_name: str
    chart_type: int
    chart_type_name: str
    dataframe: pd.DataFrame
    is_xy: bool = False
    series_names: list = field(default_factory=list)


def _extract_chart_data(chart: Chart) -> tuple[pd.DataFrame, bool, list[str]]:
    """Extract data from a chart into a DataFrame."""
    chart_type = chart.chart_type
    is_xy = chart_type in XY_CHART_TYPES

    plot = chart.plots[0]
    series_list = list(plot.series)
    series_names = [s.name if s.name else f"סדרה {i+1}" for i, s in enumerate(series_list)]

    if is_xy:
        # XY charts: each series has its own x values
        data = {}
        for i, series in enumerate(series_list):
            x_vals = list(series.values)  # In XY, x values
            y_vals = list(series.values)  # Simplified - may need adjustment
            data[f"X_{series_names[i]}"] = x_vals
            data[f"Y_{series_names[i]}"] = y_vals
        df = pd.DataFrame(data)
    else:
        # Category charts: shared categories, series have values
        try:
            categories = [str(c) for c in plot.categories]
        except Exception:
            categories = [str(i + 1) for i in range(len(list(series_list[0].values)))]

        data = {"קטגוריה": categories}
        for i, series in enumerate(series_list):
            values = list(series.values)
            # Pad if needed
            while len(values) < len(categories):
                values.append(None)
            data[series_names[i]] = values[:len(categories)]

        df = pd.DataFrame(data)

    return df, is_xy, series_names


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
    """Extract all charts from a PPTX file.

    Args:
        pptx_bytes: Raw bytes of the .pptx file

    Returns:
        List of ChartInfo objects with extracted data
    """
    prs = Presentation(BytesIO(pptx_bytes))
    charts = []

    for slide_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.has_chart:
                continue

            chart = shape.chart
            try:
                df, is_xy, series_names = _extract_chart_data(chart)
                info = ChartInfo(
                    slide_index=slide_idx,
                    shape_name=shape.name,
                    chart_type=chart.chart_type,
                    chart_type_name=_chart_type_display_name(chart.chart_type),
                    dataframe=df,
                    is_xy=is_xy,
                    series_names=series_names,
                )
                charts.append(info)
            except Exception as e:
                print(f"Warning: Could not extract chart '{shape.name}' on slide {slide_idx + 1}: {e}")

    return charts
