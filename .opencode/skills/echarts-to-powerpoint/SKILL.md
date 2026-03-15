---
name: echarts-to-powerpoint
description: >
  Convert Apache ECharts option dicts to native, editable PowerPoint charts
  using python-pptx. Use when exporting ECharts visualizations to .pptx files,
  building Solara apps with PowerPoint download, or creating editable chart
  slides from structured data. Supports bar, stacked bar, horizontal bar,
  line, area, pie, donut, scatter, effectScatter, and radar chart types.
  Unsupported types (sunburst, treemap, etc.) are detected via is_exportable().
compatibility: opencode
---

# ECharts to PowerPoint

Convert ECharts `option` dicts into native PowerPoint charts that are fully
editable (data backed by embedded Excel worksheets).

## Dependencies

```
pip install python-pptx solara
```

## Type Mapping

Map ECharts series configuration to python-pptx chart types. Sub-variants
are auto-detected from series properties:

```python
from pptx.enum.chart import XL_CHART_TYPE

# Base types
"bar"           -> COLUMN_CLUSTERED    # CategoryChartData
"line"          -> LINE_MARKERS        # CategoryChartData
"pie"           -> PIE                 # CategoryChartData
"scatter"       -> XY_SCATTER          # XyChartData
"effectScatter" -> XY_SCATTER          # XyChartData (same as scatter)
"radar"         -> RADAR_MARKERS       # CategoryChartData

# Auto-detected sub-variants
"bar" + stack          -> COLUMN_STACKED
"bar" + yAxis=category -> BAR_CLUSTERED       # horizontal
"bar" + stack + horiz  -> BAR_STACKED
"line" + areaStyle     -> AREA
"line" + area + stack  -> AREA_STACKED
"pie" + radius[0] > 0  -> DOUGHNUT           # donut
```

### Unsupported (no native PPTX equivalent)

sunburst, treemap, tree, graph, sankey, funnel, gauge, heatmap, parallel,
pictorialBar, themeRiver, map, boxplot, custom.

Check before converting:

```python
from converter import is_exportable
if is_exportable(option):
    pptx_bytes = echarts_to_pptx(option)
```

## Data Extraction

ECharts uses four data shapes:

**Plain numbers** -- categories from `xAxis.data`:
```python
categories = option["xAxis"]["data"]
values = series["data"]
```

**Named objects** -- categories from item names (pie, some bar):
```python
categories = [d["name"] for d in series["data"]]
values = [d["value"] for d in series["data"]]
```

**XY pairs** (scatter) -- no categories:
```python
points = [(d[0], d[1]) for d in series["data"]]
```

**Radar** -- categories from indicators, values from nested data:
```python
categories = [ind["name"] for ind in option["radar"]["indicator"]]
# Each entry in series[0]["data"] is {"value": [...], "name": "..."}
```

Handle `xAxis` as either dict or list (`option["xAxis"][0]["data"]`).

## Conversion Steps

1. Create `Presentation()`, set widescreen (`Inches(13.333) x Inches(7.5)`).
2. Add blank slide (`slide_layouts[6]`).
3. Read `series[0]["type"]` and detect sub-variants:
   - `stack` key -> stacked
   - `areaStyle` key -> area
   - `yAxis.type == "category"` -> horizontal bar
   - `pie` with `radius[0] > 0` -> donut
4. Build appropriate `ChartData`:
   - Category charts: `CategoryChartData` with `.categories` and `.add_series()`.
   - Scatter: `XyChartData` with `.add_series()` then `.add_data_point(x, y)`.
   - Radar: `CategoryChartData` with categories from `radar.indicator`.
5. Call `slide.shapes.add_chart(chart_type, x, y, w, h, chart_data)`.
6. Add title via `slide.shapes.add_textbox()` from `option.title.text`.
7. Enable legend for multi-series: `chart.has_legend = True`.
8. Save to `BytesIO` and return `.getvalue()`.

### Category chart example (bar/line/area/stacked)

```python
from io import BytesIO
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

chart_data = CategoryChartData()
chart_data.categories = ["Q1", "Q2", "Q3"]
chart_data.add_series("Revenue", [100, 200, 150])

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(0.5), Inches(1), Inches(12), Inches(5.5), chart_data,
)
buf = BytesIO()
prs.save(buf)
```

### Pie / donut example

```python
chart_data = CategoryChartData()
chart_data.categories = ["Chrome", "Firefox", "Safari"]
chart_data.add_series("Browsers", [65, 12, 15])
# Use PIE for standard pie, DOUGHNUT for donut
slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, ...)
```

### Scatter example

```python
from pptx.chart.data import XyChartData

chart_data = XyChartData()
xy = chart_data.add_series("Measurements")
for x, y in [(167, 65), (170, 70), (175, 75)]:
    xy.add_data_point(x, y)
slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, ...)
```

### Radar example

```python
chart_data = CategoryChartData()
chart_data.categories = ["Eng", "Design", "Sales"]
chart_data.add_series("Team A", [90, 65, 80])
chart_data.add_series("Team B", [70, 85, 60])
slide.shapes.add_chart(XL_CHART_TYPE.RADAR_MARKERS, ...)
```

## Solara Integration

Display the ECharts chart and conditionally offer PPTX download:

```python
import solara
from converter import echarts_to_pptx, is_exportable

@solara.component
def Page():
    option = { ... }
    solara.FigureEcharts(option=option, responsive=True)
    if is_exportable(option):
        solara.FileDownload(
            lambda: echarts_to_pptx(option),
            filename="chart.pptx",
            mime_type="application/vnd.openxmlformats-officedocument"
                     ".presentationml.presentation",
        )
    else:
        solara.Button("Export not available", disabled=True)
```

## Gotchas

- Use `slide_layouts[6]` (Blank) to avoid placeholder boxes.
- Pie/donut only uses the first series.
- For horizontal bar, categories come from yAxis, not xAxis.
- `xAxis` can be a dict or a list -- handle both.
- `radar` option can be a dict or a list -- handle both.
- Title goes in a textbox, not a layout placeholder.
- All charts are backed by embedded Excel -- fully editable in PowerPoint.
