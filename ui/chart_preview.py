"""Render chart data as interactive Plotly figures."""

from __future__ import annotations

import plotly.graph_objects as go
from pptx.enum.chart import XL_CHART_TYPE

from core.data_extractor import ChartInfo, is_percentage_format


# Map PPTX chart types to Plotly rendering style
_BAR_TYPES = {
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    XL_CHART_TYPE.COLUMN_STACKED,
    XL_CHART_TYPE.COLUMN_STACKED_100,
}
_HBAR_TYPES = {
    XL_CHART_TYPE.BAR_CLUSTERED,
    XL_CHART_TYPE.BAR_STACKED,
    XL_CHART_TYPE.BAR_STACKED_100,
}
_STACKED_TYPES = {
    XL_CHART_TYPE.COLUMN_STACKED,
    XL_CHART_TYPE.COLUMN_STACKED_100,
    XL_CHART_TYPE.BAR_STACKED,
    XL_CHART_TYPE.BAR_STACKED_100,
    XL_CHART_TYPE.LINE_STACKED,
    XL_CHART_TYPE.AREA_STACKED,
}
_LINE_TYPES = {
    XL_CHART_TYPE.LINE,
    XL_CHART_TYPE.LINE_MARKERS,
    XL_CHART_TYPE.LINE_STACKED,
}
_AREA_TYPES = {
    XL_CHART_TYPE.AREA,
    XL_CHART_TYPE.AREA_STACKED,
}
_PIE_TYPES = {
    XL_CHART_TYPE.PIE,
    XL_CHART_TYPE.PIE_EXPLODED,
    XL_CHART_TYPE.DOUGHNUT,
}
_SCATTER_TYPES = {
    XL_CHART_TYPE.XY_SCATTER,
    XL_CHART_TYPE.XY_SCATTER_LINES,
    XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
}


def render_chart_plotly(
    df,
    chart_type: int,
    series_visibility: dict[str, bool],
    series_formats: dict[str, str],
) -> go.Figure:
    """Build a Plotly figure from chart data and metadata."""
    fig = go.Figure()
    cat_col = df.columns[0]
    categories = df[cat_col].tolist()
    series_cols = [c for c in df.columns[1:] if series_visibility.get(c, True)]

    has_pct = any(is_percentage_format(series_formats.get(c, "")) for c in series_cols)

    if chart_type in _PIE_TYPES:
        # For pie/doughnut, use first visible series
        if series_cols:
            col = series_cols[0]
            hole = 0.4 if chart_type == XL_CHART_TYPE.DOUGHNUT else 0
            fig.add_trace(go.Pie(
                labels=categories,
                values=df[col].tolist(),
                hole=hole,
                name=col,
            ))
        fig.update_layout(margin=dict(t=30, b=30, l=30, r=30))
        return fig

    if chart_type in _SCATTER_TYPES:
        # XY scatter — columns come in pairs: X_name, Y_name
        for i in range(0, len(series_cols), 2):
            x_col = series_cols[i] if i < len(series_cols) else None
            y_col = series_cols[i + 1] if i + 1 < len(series_cols) else None
            if x_col and y_col:
                fig.add_trace(go.Scatter(
                    x=df[x_col].tolist(),
                    y=df[y_col].tolist(),
                    mode="markers",
                    name=x_col.replace("X_", ""),
                ))
        fig.update_layout(margin=dict(t=30, b=30, l=50, r=30))
        return fig

    # Bar, line, area charts
    barmode = "stack" if chart_type in _STACKED_TYPES else "group"
    is_horizontal = chart_type in _HBAR_TYPES

    for col in series_cols:
        values = df[col].tolist()

        if chart_type in _BAR_TYPES:
            fig.add_trace(go.Bar(x=categories, y=values, name=col))
        elif chart_type in _HBAR_TYPES:
            fig.add_trace(go.Bar(x=values, y=categories, name=col, orientation="h"))
        elif chart_type in _LINE_TYPES:
            mode = "lines+markers" if chart_type == XL_CHART_TYPE.LINE_MARKERS else "lines"
            fig.add_trace(go.Scatter(x=categories, y=values, mode=mode, name=col))
        elif chart_type in _AREA_TYPES:
            fig.add_trace(go.Scatter(
                x=categories, y=values, mode="lines", fill="tonexty", name=col,
            ))
        else:
            # Fallback: bar chart
            fig.add_trace(go.Bar(x=categories, y=values, name=col))

    fig.update_layout(
        barmode=barmode,
        margin=dict(t=40, b=40, l=50, r=30),
        height=400,
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        hovermode="x unified",
    )

    if has_pct:
        axis = "xaxis" if is_horizontal else "yaxis"
        fig.update_layout(**{axis: dict(ticksuffix="%")})

    return fig
