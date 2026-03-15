# ECharts to PowerPoint -- Implementation Guide

## Core Idea

ECharts and PowerPoint both represent charts as structured data (categories + series). Convert between their formats to produce **native, editable** PPTX charts -- not images.

## Dependencies

```
pip install python-pptx solara
```

## Type Mapping

| ECharts pattern | python-pptx chart type | Data class |
|---|---|---|
| `bar` | `COLUMN_CLUSTERED` | `CategoryChartData` |
| `bar` + `stack` | `COLUMN_STACKED` | `CategoryChartData` |
| `bar` + `yAxis.type="category"` | `BAR_CLUSTERED` (horizontal) | `CategoryChartData` |
| `bar` + `stack` + horizontal | `BAR_STACKED` | `CategoryChartData` |
| `line` | `LINE_MARKERS` | `CategoryChartData` |
| `line` + `areaStyle` | `AREA` | `CategoryChartData` |
| `line` + `areaStyle` + `stack` | `AREA_STACKED` | `CategoryChartData` |
| `pie` | `PIE` | `CategoryChartData` |
| `pie` + `radius[0] > 0` (donut) | `DOUGHNUT` | `CategoryChartData` |
| `scatter` / `effectScatter` | `XY_SCATTER` | `XyChartData` |
| `radar` | `RADAR_MARKERS` | `CategoryChartData` |

### Unsupported (no native PPTX equivalent)

sunburst, treemap, tree, graph, sankey, funnel, gauge, heatmap, parallel, pictorialBar, themeRiver, map, boxplot, custom.

Use `is_exportable(option)` to check before attempting conversion.

## ECharts Data Formats

**1. Plain numbers** -- categories from `xAxis.data`:
```python
{"xAxis": {"data": ["Mon", "Tue"]}, "series": [{"data": [120, 200]}]}
```

**2. Named objects** -- categories from item names (pie, some bar):
```python
{"series": [{"data": [{"name": "Shirts", "value": 5}, {"name": "Pants", "value": 10}]}]}
```

**3. XY pairs** (scatter) -- no categories:
```python
{"series": [{"type": "scatter", "data": [[1, 10], [2, 20]]}]}
```

**4. Radar** -- categories from `radar.indicator[].name`, values from `series[].data[].value`:
```python
{"radar": {"indicator": [{"name": "Eng", "max": 100}, ...]},
 "series": [{"type": "radar", "data": [{"value": [90, 65], "name": "Team A"}]}]}
```

## Sub-variant Detection

The converter auto-detects these variants from the ECharts option:

- **Horizontal bar**: `yAxis.type == "category"`
- **Stacked**: any series has a `stack` key
- **Area**: any series has an `areaStyle` key
- **Donut**: pie with `radius[0]` non-zero (e.g. `["40%", "70%"]`)

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
    slide = prs.slides.add_slide(prs.slide_layouts[6])

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

Display the chart with `FigureEcharts`, conditionally show download via `is_exportable`:

```python
import solara
from converter import echarts_to_pptx, is_exportable

@solara.component
def Page():
    option = { ... }  # ECharts option dict

    solara.FigureEcharts(option=option, responsive=True,
                         attributes={"style": "height:500px"})

    if is_exportable(option):
        solara.FileDownload(
            lambda: echarts_to_pptx(option),
            filename="chart.pptx",
            mime_type="application/vnd.openxmlformats-officedocument"
                   ".presentationml.presentation",
            label="Download PowerPoint",
        )
    else:
        solara.Button("Export not available", disabled=True)
```

Run with `solara run app.py`.

## Key Details

- Use `slide_layouts[6]` (Blank) to avoid placeholder clutter.
- For scatter/effectScatter, use `XyChartData` with `add_data_point(x, y)`.
- Pie charts only use the first series. Categories come from `data[].name`.
- For horizontal bar, swap xAxis/yAxis when extracting categories.
- For radar, categories come from `option.radar.indicator`, not xAxis.
- Enable legends for multi-series charts: `chart.has_legend = True`.
- Title goes in a textbox via `slide.shapes.add_textbox(...)`.
- All exported charts are backed by embedded Excel worksheets -- fully editable in PowerPoint.
