"""Convert ECharts option dicts to native, editable PowerPoint charts."""

from io import BytesIO

from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

ECHART_TYPE_MAP = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
}

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

CHART_LEFT = Inches(0.5)
CHART_TOP = Inches(1.2)
CHART_WIDTH = Inches(12.333)
CHART_HEIGHT = Inches(5.8)


def _extract_title(option: dict) -> str | None:
    title = option.get("title")
    if isinstance(title, dict):
        return title.get("text")
    return None


def _extract_categories(option: dict, series_list: list[dict]) -> list[str]:
    """Extract category labels from xAxis.data or from series data point names."""
    x_axis = option.get("xAxis")
    if isinstance(x_axis, dict) and "data" in x_axis:
        return [str(c) for c in x_axis["data"]]
    if isinstance(x_axis, list) and x_axis and "data" in x_axis[0]:
        return [str(c) for c in x_axis[0]["data"]]

    for s in series_list:
        data = s.get("data", [])
        if data and isinstance(data[0], dict) and "name" in data[0]:
            return [str(d["name"]) for d in data]

    max_len = max((len(s.get("data", [])) for s in series_list), default=0)
    return [str(i + 1) for i in range(max_len)]


def _extract_series_values(series: dict) -> list[float]:
    """Extract numeric values from a series data array."""
    values = []
    for item in series.get("data", []):
        if isinstance(item, dict):
            values.append(float(item.get("value", 0)))
        elif isinstance(item, (int, float)):
            values.append(float(item))
        else:
            values.append(0.0)
    return values


def _extract_scatter_points(series: dict) -> list[tuple[float, float]]:
    """Extract (x, y) pairs from scatter series data."""
    points = []
    for item in series.get("data", []):
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            points.append((float(item[0]), float(item[1])))
    return points


def _add_title_textbox(slide, text: str):
    txBox = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2), Inches(12.333), Inches(0.8)
    )
    p = txBox.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(24)
    p.font.bold = True


def _build_category_chart(
    prs: Presentation,
    option: dict,
    series_list: list[dict],
    chart_type_str: str,
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = _extract_title(option)
    if title:
        _add_title_textbox(slide, title)

    categories = _extract_categories(option, series_list)

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for s in series_list:
        chart_data.add_series(s.get("name", "Series"), _extract_series_values(s))

    pptx_type = ECHART_TYPE_MAP.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)
    graphic_frame = slide.shapes.add_chart(
        pptx_type, CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT, chart_data
    )
    chart = graphic_frame.chart

    if len(series_list) > 1:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False


def _build_pie_chart(prs: Presentation, option: dict, series_list: list[dict]):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = _extract_title(option)
    if title:
        _add_title_textbox(slide, title)

    series = series_list[0]
    categories = _extract_categories(option, [series])
    values = _extract_series_values(series)

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(series.get("name", "Series"), values)

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT, chart_data
    )
    chart = graphic_frame.chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False


def _build_scatter_chart(prs: Presentation, option: dict, series_list: list[dict]):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = _extract_title(option)
    if title:
        _add_title_textbox(slide, title)

    chart_data = XyChartData()
    for s in series_list:
        xy_series = chart_data.add_series(s.get("name", "Series"))
        for x, y in _extract_scatter_points(s):
            xy_series.add_data_point(x, y)

    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER,
        CHART_LEFT,
        CHART_TOP,
        CHART_WIDTH,
        CHART_HEIGHT,
        chart_data,
    )
    chart = graphic_frame.chart

    if len(series_list) > 1:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False


def echarts_to_pptx(option: dict) -> bytes:
    """Convert an ECharts option dict to a PowerPoint file returned as bytes.

    The resulting PPTX contains native, editable charts backed by embedded
    Excel worksheets.  Supported ECharts series types: bar, line, pie, scatter.
    """
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    series_list = option.get("series", [])
    if not series_list:
        raise ValueError("ECharts option must contain at least one series")

    chart_type = series_list[0].get("type", "bar")

    if chart_type == "pie":
        _build_pie_chart(prs, option, series_list)
    elif chart_type == "scatter":
        _build_scatter_chart(prs, option, series_list)
    else:
        _build_category_chart(prs, option, series_list, chart_type)

    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()
