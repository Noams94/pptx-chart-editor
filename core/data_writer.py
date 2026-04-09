"""Write edited DataFrames back into PPTX charts."""

from __future__ import annotations

from io import BytesIO

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.oxml.ns import qn
from lxml import etree

from core.data_extractor import _is_percentage_format


def _display_to_raw(df: pd.DataFrame, series_formats: dict) -> pd.DataFrame:
    """Convert display values back to raw values.

    E.g., 67 (displayed as 67%) -> 0.67 (raw value stored in chart)
    """
    raw_df = df.copy()
    for col in df.columns[1:]:  # Skip first column (categories)
        fmt = series_formats.get(col, "General")
        if _is_percentage_format(fmt):
            raw_df[col] = df[col].apply(
                lambda v: v / 100.0 if pd.notna(v) else None
            )
    return raw_df


def update_chart_data(
    pptx_bytes: bytes,
    slide_index: int,
    shape_name: str,
    display_df: pd.DataFrame,
    is_xy: bool = False,
    series_formats: dict = None,
) -> bytes:
    """Update a single chart's data in the PPTX and return updated bytes.

    Args:
        pptx_bytes: Original PPTX file bytes
        slide_index: Zero-based slide index
        shape_name: Name of the chart shape
        display_df: DataFrame with display values (67 for 67%)
        is_xy: Whether this is an XY/scatter chart
        series_formats: Dict of series_name -> formatCode for converting display->raw

    Returns:
        Updated PPTX file bytes
    """
    # Convert display values to raw values
    if series_formats:
        df = _display_to_raw(display_df, series_formats)
    else:
        df = display_df

    prs = Presentation(BytesIO(pptx_bytes))
    slide = prs.slides[slide_index]

    # Find the chart shape
    chart_shape = None
    for shape in slide.shapes:
        if shape.has_chart and shape.name == shape_name:
            chart_shape = shape
            break

    if chart_shape is None:
        raise ValueError(f"Chart '{shape_name}' not found on slide {slide_index + 1}")

    chart = chart_shape.chart

    if is_xy:
        chart_data = XyChartData()
        col_names = df.columns.tolist()
        for i in range(0, len(col_names), 2):
            series_name = col_names[i].replace("X_", "")
            series = chart_data.add_series(series_name)
            x_vals = df.iloc[:, i].dropna().tolist()
            y_vals = df.iloc[:, i + 1].dropna().tolist()
            for x, y in zip(x_vals, y_vals):
                series.add_data_point(x, y)
    else:
        chart_data = CategoryChartData()
        categories = df.iloc[:, 0].dropna().astype(str).tolist()
        chart_data.categories = categories

        for col in df.columns[1:]:
            values = df[col].tolist()
            values = [None if pd.isna(v) else float(v) for v in values]
            values = values[:len(categories)]
            while len(values) < len(categories):
                values.append(None)
            chart_data.add_series(col, values)

    chart.replace_data(chart_data)

    # Restore number format codes that replace_data() resets to "General"
    if series_formats:
        _restore_format_codes(chart, series_formats)

    output = BytesIO()
    prs.save(output)
    return output.getvalue()


def _restore_format_codes(chart, series_formats: dict):
    """Restore formatCode in chart XML after replace_data() resets them."""
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

        if not name or name not in series_formats:
            continue

        fmt_code = series_formats[name]

        # Set formatCode in val > numRef > numCache > formatCode
        val = ser.find(qn('c:val'))
        if val is not None:
            num_ref = val.find(qn('c:numRef'))
            if num_ref is not None:
                num_cache = num_ref.find(qn('c:numCache'))
                if num_cache is not None:
                    fc = num_cache.find(qn('c:formatCode'))
                    if fc is None:
                        fc = etree.SubElement(num_cache, qn('c:formatCode'))
                    fc.text = fmt_code
