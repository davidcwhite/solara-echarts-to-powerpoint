---
name: echarts-to-powerpoint
description: >
  Convert Apache ECharts option dicts to native, editable PowerPoint charts
  using python-pptx. Use when exporting ECharts visualizations to .pptx files,
  building Solara apps with PowerPoint download, or creating editable chart
  slides from structured data. Covers bar, line, pie, and scatter chart types.
compatibility: opencode
---

# ECharts to PowerPoint

Convert ECharts `option` dicts into native PowerPoint charts that are fully
editable (data backed by embedded Excel worksheets).

## Dependencies

Install `python-pptx` (chart generation) and optionally `solara` (web app):

```
pip install python-pptx solara
```

## Type Mapping

Map ECharts `series[].type` to python-pptx constants:

| ECharts type | `XL_CHART_TYPE` | Data class |
|---|---|---|
| `bar` | `COLUMN_CLUSTERED` | `CategoryChartData` |
| `line` | `LINE_MARKERS` | `CategoryChartData` |
| `pie` | `PIE` | `CategoryChartData` |
| `scatter` | `XY_SCATTER` | `XyChartData` |

```python
from pptx.enum.chart import XL_CHART_TYPE
TYPE_MAP = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
}
```

## Data Extraction

ECharts uses three data shapes. Handle each:

**Plain numbers** -- categories from `xAxis.data`:
```python
categories = option["xAxis"]["data"]        # ["Mon", "Tue"]
values = series["data"]                     # [120, 200]
```

**Named objects** -- categories from item names (used by pie and some bar charts):
```python
categories = [d["name"] for d in series["data"]]   # ["Shirts", "Pants"]
values = [d["value"] for d in series["data"]]       # [5, 10]
```

**XY pairs** (scatter only) -- no categories:
```python
points = [(d[0], d[1]) for d in series["data"]]    # [(1, 10), (2, 20)]
```

Also handle `xAxis` as a list (`option["xAxis"][0]["data"]`).

## Conversion Steps

1. Create a `Presentation()` and set widescreen dimensions.
2. Add a blank slide (`slide_layouts[6]`).
3. Read `series[0]["type"]` to determine chart type.
4. Build the appropriate `ChartData` object:
   - Category charts: `CategoryChartData` with `.categories` and `.add_series()`.
   - Scatter: `XyChartData` with `.add_series()` then `.add_data_point(x, y)`.
5. Call `slide.shapes.add_chart(chart_type, x, y, w, h, chart_data)`.
6. Optionally add title via `slide.shapes.add_textbox()`.
7. Enable legend for multi-series: `chart.has_legend = True`.
8. Save to `BytesIO` and return `.getvalue()`.

### Category chart example (bar/line)

```python
from io import BytesIO
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

option = {
    "xAxis": {"data": ["Q1", "Q2", "Q3"]},
    "series": [{"name": "Revenue", "type": "bar", "data": [100, 200, 150]}],
}

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

chart_data = CategoryChartData()
chart_data.categories = option["xAxis"]["data"]
chart_data.add_series("Revenue", option["series"][0]["data"])

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(0.5), Inches(1), Inches(12), Inches(5.5),
    chart_data,
)

buf = BytesIO()
prs.save(buf)
pptx_bytes = buf.getvalue()
```

### Pie chart example

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

series_data = [
    {"name": "Product A", "value": 335},
    {"name": "Product B", "value": 234},
]

chart_data = CategoryChartData()
chart_data.categories = [d["name"] for d in series_data]
chart_data.add_series("Share", [d["value"] for d in series_data])

slide.shapes.add_chart(
    XL_CHART_TYPE.PIE,
    Inches(0.5), Inches(1), Inches(12), Inches(5.5),
    chart_data,
)
```

### Scatter chart example

```python
from pptx.chart.data import XyChartData
from pptx.enum.chart import XL_CHART_TYPE

points = [[167, 65], [170, 70], [175, 75]]

chart_data = XyChartData()
xy = chart_data.add_series("Measurements")
for x, y in points:
    xy.add_data_point(x, y)

slide.shapes.add_chart(
    XL_CHART_TYPE.XY_SCATTER,
    Inches(0.5), Inches(1), Inches(12), Inches(5.5),
    chart_data,
)
```

## Solara Integration

Display the ECharts chart and offer PPTX download:

```python
import solara
from converter import echarts_to_pptx

@solara.component
def Page():
    option = {"xAxis": {"data": ["Q1", "Q2"]}, "yAxis": {},
              "series": [{"name": "Sales", "type": "bar", "data": [10, 20]}]}

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

## Gotchas

- Use `slide_layouts[6]` (Blank) to avoid placeholder boxes.
- Pie charts only use the first series.
- `xAxis` can be a dict or a list of dicts -- handle both.
- Title goes in a textbox, not a layout placeholder.
- All charts are backed by embedded Excel -- fully editable in PowerPoint.
