"""Round-trip tests: ECharts option -> PPTX -> read back and verify."""

from io import BytesIO
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE

from converter import echarts_to_pptx

TEST_OUTPUT = Path(__file__).resolve().parent.parent / "test_output"


@pytest.fixture(autouse=True)
def _output_dir():
    TEST_OUTPUT.mkdir(exist_ok=True)


def _load_chart(pptx_bytes: bytes):
    """Load PPTX bytes and return the first chart object found."""
    prs = Presentation(BytesIO(pptx_bytes))
    assert len(prs.slides) == 1
    for shape in prs.slides[0].shapes:
        if shape.has_chart:
            return shape.chart
    pytest.fail("No chart shape found on the slide")


# ── Bar chart ────────────────────────────────────────────────────────

BAR_OPTION = {
    "title": {"text": "Bar Test"},
    "xAxis": {"type": "category", "data": ["Mon", "Tue", "Wed"]},
    "yAxis": {"type": "value"},
    "series": [{"name": "Sales", "type": "bar", "data": [120, 200, 150]}],
}


def test_bar_chart_type():
    data = echarts_to_pptx(BAR_OPTION)
    (TEST_OUTPUT / "bar_chart.pptx").write_bytes(data)
    chart = _load_chart(data)
    assert chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED


def test_bar_chart_categories():
    chart = _load_chart(echarts_to_pptx(BAR_OPTION))
    categories = [str(c) for c in chart.plots[0].categories]
    assert categories == ["Mon", "Tue", "Wed"]


def test_bar_chart_values():
    chart = _load_chart(echarts_to_pptx(BAR_OPTION))
    assert list(chart.series[0].values) == [120.0, 200.0, 150.0]


# ── Multi-series bar ────────────────────────────────────────────────

MULTI_BAR_OPTION = {
    "title": {"text": "Multi-series Bar"},
    "xAxis": {"type": "category", "data": ["A", "B", "C"]},
    "yAxis": {"type": "value"},
    "series": [
        {"name": "Series 1", "type": "bar", "data": [10, 20, 30]},
        {"name": "Series 2", "type": "bar", "data": [15, 25, 35]},
    ],
}


def test_multi_series_bar():
    data = echarts_to_pptx(MULTI_BAR_OPTION)
    (TEST_OUTPUT / "multi_bar_chart.pptx").write_bytes(data)
    chart = _load_chart(data)
    assert len(chart.series) == 2
    assert list(chart.series[0].values) == [10.0, 20.0, 30.0]
    assert list(chart.series[1].values) == [15.0, 25.0, 35.0]


# ── Bar with named data points ──────────────────────────────────────

NAMED_BAR_OPTION = {
    "title": {"text": "Named Bar"},
    "xAxis": {"type": "category"},
    "yAxis": {"type": "value"},
    "series": [
        {
            "name": "sales",
            "type": "bar",
            "data": [
                {"name": "Shirts", "value": 5},
                {"name": "Pants", "value": 10},
                {"name": "Socks", "value": 20},
            ],
        }
    ],
}


def test_named_data_categories():
    data = echarts_to_pptx(NAMED_BAR_OPTION)
    (TEST_OUTPUT / "named_bar_chart.pptx").write_bytes(data)
    chart = _load_chart(data)
    categories = [str(c) for c in chart.plots[0].categories]
    assert categories == ["Shirts", "Pants", "Socks"]


def test_named_data_values():
    chart = _load_chart(echarts_to_pptx(NAMED_BAR_OPTION))
    assert list(chart.series[0].values) == [5.0, 10.0, 20.0]


# ── Line chart ──────────────────────────────────────────────────────

LINE_OPTION = {
    "title": {"text": "Line Test"},
    "xAxis": {"type": "category", "data": ["Jan", "Feb", "Mar"]},
    "yAxis": {"type": "value"},
    "series": [
        {"name": "Revenue", "type": "line", "data": [300, 450, 500]},
    ],
}


def test_line_chart_type():
    data = echarts_to_pptx(LINE_OPTION)
    (TEST_OUTPUT / "line_chart.pptx").write_bytes(data)
    chart = _load_chart(data)
    assert chart.chart_type == XL_CHART_TYPE.LINE_MARKERS


def test_line_chart_categories():
    chart = _load_chart(echarts_to_pptx(LINE_OPTION))
    assert [str(c) for c in chart.plots[0].categories] == ["Jan", "Feb", "Mar"]


def test_line_chart_values():
    chart = _load_chart(echarts_to_pptx(LINE_OPTION))
    assert list(chart.series[0].values) == [300.0, 450.0, 500.0]


# ── Pie chart ───────────────────────────────────────────────────────

PIE_OPTION = {
    "title": {"text": "Pie Test"},
    "series": [
        {
            "name": "Share",
            "type": "pie",
            "data": [
                {"name": "A", "value": 100},
                {"name": "B", "value": 200},
                {"name": "C", "value": 300},
            ],
        }
    ],
}


def test_pie_chart_type():
    data = echarts_to_pptx(PIE_OPTION)
    (TEST_OUTPUT / "pie_chart.pptx").write_bytes(data)
    chart = _load_chart(data)
    assert chart.chart_type == XL_CHART_TYPE.PIE


def test_pie_chart_categories():
    chart = _load_chart(echarts_to_pptx(PIE_OPTION))
    assert [str(c) for c in chart.plots[0].categories] == ["A", "B", "C"]


def test_pie_chart_values():
    chart = _load_chart(echarts_to_pptx(PIE_OPTION))
    assert list(chart.series[0].values) == [100.0, 200.0, 300.0]


# ── Scatter chart ───────────────────────────────────────────────────

SCATTER_OPTION = {
    "title": {"text": "Scatter Test"},
    "xAxis": {"type": "value"},
    "yAxis": {"type": "value"},
    "series": [
        {
            "name": "Points",
            "type": "scatter",
            "data": [[1, 10], [2, 20], [3, 30]],
        }
    ],
}


def test_scatter_chart_type():
    data = echarts_to_pptx(SCATTER_OPTION)
    (TEST_OUTPUT / "scatter_chart.pptx").write_bytes(data)
    chart = _load_chart(data)
    assert chart.chart_type == XL_CHART_TYPE.XY_SCATTER


def test_scatter_series_count():
    chart = _load_chart(echarts_to_pptx(SCATTER_OPTION))
    assert len(chart.series) == 1


# ── Edge cases ──────────────────────────────────────────────────────


def test_empty_series_raises():
    with pytest.raises(ValueError, match="at least one series"):
        echarts_to_pptx({"series": []})


def test_no_series_key_raises():
    with pytest.raises(ValueError, match="at least one series"):
        echarts_to_pptx({})


def test_no_title_does_not_crash():
    option = {
        "xAxis": {"type": "category", "data": ["X"]},
        "yAxis": {"type": "value"},
        "series": [{"type": "bar", "data": [42]}],
    }
    data = echarts_to_pptx(option)
    chart = _load_chart(data)
    assert chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED


def test_xaxis_as_list():
    """ECharts allows xAxis to be an array of axis objects."""
    option = {
        "xAxis": [{"type": "category", "data": ["Q1", "Q2"]}],
        "yAxis": {"type": "value"},
        "series": [{"type": "bar", "data": [10, 20]}],
    }
    chart = _load_chart(echarts_to_pptx(option))
    assert [str(c) for c in chart.plots[0].categories] == ["Q1", "Q2"]
