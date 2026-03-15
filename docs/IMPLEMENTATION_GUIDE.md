# ECharts to PowerPoint -- Implementation Guide

## Core Idea

ECharts and PowerPoint both represent charts as structured data (categories + series). Convert between their formats to produce **native, editable** PPTX charts -- not images.

## Dependencies

```
pip install python-pptx solara
```

## Type Mapping

| ECharts `series.type` | python-pptx chart type | Data class |
|---|---|---|
| `bar` | `XL_CHART_TYPE.COLUMN_CLUSTERED` | `CategoryChartData` |
| `line` | `XL_CHART_TYPE.LINE_MARKERS` | `CategoryChartData` |
| `pie` | `XL_CHART_TYPE.PIE` | `CategoryChartData` |
| `scatter` | `XL_CHART_TYPE.XY_SCATTER` | `XyChartData` |

## ECharts Data Formats

ECharts uses three data shapes. Extract categories and values accordingly:

**1. Plain numbers** -- categories come from `xAxis.data`:
```python
{"xAxis": {"data": ["Mon", "Tue"]}, "series": [{"data": [120, 200]}]}
#  categories = option["xAxis"]["data"]
#  values     = series["data"]
```

**2. Named objects** -- categories come from each item's `name`:
```python
{"series": [{"data": [{"name": "Shirts", "value": 5}, {"name": "Pants", "value": 10}]}]}
#  categories = [d["name"] for d in series["data"]]
#  values     = [d["value"] for d in series["data"]]
```

**3. XY pairs** (scatter) -- no categories, use `XyChartData`:
```python
{"series": [{"type": "scatter", "data": [[1, 10], [2, 20]]}]}
#  points = [(d[0], d[1]) for d in series["data"]]
```

## Minimal Conversion Example

```python
from io import BytesIO
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

def echarts_to_pptx(option):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

    series = option["series"][0]
    chart_data = CategoryChartData()
    chart_data.categories = option["xAxis"]["data"]
    chart_data.add_series(series.get("name", "Series"), series["data"])

    TYPE_MAP = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
                "line": XL_CHART_TYPE.LINE_MARKERS,
                "pie": XL_CHART_TYPE.PIE}

    slide.shapes.add_chart(
        TYPE_MAP[series["type"]],
        Inches(0.5), Inches(1), Inches(12), Inches(5.5),
        chart_data,
    )
    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()
```

## Solara Wiring

Display the chart with `FigureEcharts` and offer download via `FileDownload`:

```python
import solara
from converter import echarts_to_pptx

@solara.component
def Page():
    option = {
        "xAxis": {"type": "category", "data": ["Q1", "Q2", "Q3"]},
        "yAxis": {"type": "value"},
        "series": [{"name": "Revenue", "type": "bar", "data": [100, 200, 150]}],
    }

    solara.FigureEcharts(option=option, responsive=True,
                         attributes={"style": "height:500px"})
    solara.FileDownload(
        lambda: echarts_to_pptx(option),
        filename="chart.pptx",
        mime_type="application/vnd.openxmlformats-officedocument"
               ".presentationml.presentation",
        label="Download PowerPoint",
    )
```

Run with `solara run app.py`.

## Key Details

- Use `slide_layouts[6]` (Blank) to avoid placeholder clutter.
- For scatter charts, use `XyChartData` with `add_data_point(x, y)` instead of `CategoryChartData`.
- Pie charts only use the first series. Categories come from `data[].name`.
- Enable legends for multi-series charts: `chart.has_legend = True`.
- Title goes in a textbox via `slide.shapes.add_textbox(...)`, not the layout title placeholder.
- All exported charts are backed by embedded Excel worksheets -- fully editable in PowerPoint.
