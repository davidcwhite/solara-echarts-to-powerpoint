# Solara ECharts to PowerPoint

A lightweight Solara web app that displays ECharts charts and exports them to PowerPoint as native, editable charts.

## Features

- Render ECharts charts in the browser via Solara's `FigureEcharts`
- Export to PowerPoint with fully editable native charts (not images)
- Chart data backed by embedded Excel worksheets -- double-click any chart in PowerPoint to edit the data
- Supported chart types: **bar**, **line**, **pie**, **scatter**

## Installation

```bash
pip install -r requirements.txt
```

## Running the App

```bash
solara run app.py
```

Then open http://localhost:8765 in your browser. Use the toggle buttons to switch chart types, and click "Download PowerPoint" to export.

## Running Tests

```bash
pytest tests/ -v
```

Generated test PPTX files are saved to `test_output/` for manual inspection in PowerPoint.

## How It Works

The converter reads the ECharts `option` dict (the same JSON structure used by [Apache ECharts](https://echarts.apache.org/)) and maps it to native PowerPoint chart objects via `python-pptx`:

| ECharts type | PowerPoint chart type | python-pptx data class |
|---|---|---|
| `bar` | Clustered Column | `CategoryChartData` |
| `line` | Line with Markers | `CategoryChartData` |
| `pie` | Pie | `CategoryChartData` |
| `scatter` | XY Scatter | `XyChartData` |

Because the charts are native PowerPoint objects, all data, labels, colours, and formatting are fully editable after export.
