"""Write edited DataFrames back into PPTX charts."""

from io import BytesIO

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData


def update_chart_data(
    pptx_bytes: bytes,
    slide_index: int,
    shape_name: str,
    df: pd.DataFrame,
    is_xy: bool = False,
) -> bytes:
    """Update a single chart's data in the PPTX and return updated bytes.

    Args:
        pptx_bytes: Original PPTX file bytes
        slide_index: Zero-based slide index
        shape_name: Name of the chart shape
        df: DataFrame with updated data (first column = categories, rest = series)
        is_xy: Whether this is an XY/scatter chart

    Returns:
        Updated PPTX file bytes
    """
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
        # XY charts: pairs of X/Y columns
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
        # First column is categories
        categories = df.iloc[:, 0].dropna().astype(str).tolist()
        chart_data.categories = categories

        # Remaining columns are series
        for col in df.columns[1:]:
            values = df[col].tolist()
            # Replace NaN with None for python-pptx
            values = [None if pd.isna(v) else float(v) for v in values]
            # Trim to match categories length
            values = values[:len(categories)]
            # Pad if needed
            while len(values) < len(categories):
                values.append(None)
            chart_data.add_series(col, values)

    chart.replace_data(chart_data)

    # Save to bytes
    output = BytesIO()
    prs.save(output)
    return output.getvalue()


def update_multiple_charts(
    pptx_bytes: bytes,
    edits: list[dict],
) -> bytes:
    """Apply multiple chart edits to a PPTX file.

    Args:
        pptx_bytes: Original PPTX file bytes
        edits: List of dicts with keys: slide_index, shape_name, df, is_xy

    Returns:
        Updated PPTX file bytes
    """
    current_bytes = pptx_bytes
    for edit in edits:
        current_bytes = update_chart_data(
            current_bytes,
            edit["slide_index"],
            edit["shape_name"],
            edit["df"],
            edit.get("is_xy", False),
        )
    return current_bytes
