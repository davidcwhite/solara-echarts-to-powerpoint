"""Solara app: display ECharts charts and export to PowerPoint."""

import solara

from converter import echarts_to_pptx

CHARTS = {
    "Bar": {
        "title": {"text": "Quarterly Sales by Region"},
        "tooltip": {},
        "legend": {"data": ["Q1", "Q2", "Q3"]},
        "xAxis": {
            "type": "category",
            "data": ["East", "West", "North", "South"],
        },
        "yAxis": {"type": "value"},
        "series": [
            {"name": "Q1", "type": "bar", "data": [120, 200, 150, 80]},
            {"name": "Q2", "type": "bar", "data": [180, 230, 170, 65]},
            {"name": "Q3", "type": "bar", "data": [150, 210, 190, 90]},
        ],
    },
    "Line": {
        "title": {"text": "Monthly Temperature Trends"},
        "tooltip": {"trigger": "axis"},
        "legend": {"data": ["New York", "London"]},
        "xAxis": {
            "type": "category",
            "data": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        },
        "yAxis": {"type": "value"},
        "series": [
            {
                "name": "New York",
                "type": "line",
                "data": [2, 4, 10, 16, 22, 27],
            },
            {
                "name": "London",
                "type": "line",
                "data": [5, 6, 9, 12, 16, 19],
            },
        ],
    },
    "Pie": {
        "title": {"text": "Market Share"},
        "tooltip": {"trigger": "item"},
        "series": [
            {
                "name": "Market Share",
                "type": "pie",
                "radius": ["0%", "60%"],
                "data": [
                    {"name": "Product A", "value": 335},
                    {"name": "Product B", "value": 234},
                    {"name": "Product C", "value": 154},
                    {"name": "Product D", "value": 135},
                    {"name": "Product E", "value": 105},
                ],
            }
        ],
    },
    "Scatter": {
        "title": {"text": "Height vs Weight"},
        "tooltip": {"trigger": "item"},
        "xAxis": {"type": "value", "name": "Height (cm)"},
        "yAxis": {"type": "value", "name": "Weight (kg)"},
        "series": [
            {
                "name": "Male",
                "type": "scatter",
                "data": [
                    [167, 65],
                    [170, 70],
                    [175, 75],
                    [180, 82],
                    [185, 90],
                    [172, 68],
                    [178, 78],
                ],
            },
            {
                "name": "Female",
                "type": "scatter",
                "data": [
                    [155, 50],
                    [160, 55],
                    [165, 60],
                    [158, 53],
                    [162, 57],
                    [168, 62],
                    [170, 65],
                ],
            },
        ],
    },
}

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


@solara.component
def Page():
    chart_type, set_chart_type = solara.use_state("Bar")

    option = CHARTS[chart_type]

    with solara.AppLayout(title="ECharts to PowerPoint"):
        with solara.Card(margin=2):
            with solara.Row(
                justify="space-between",
                style={"align-items": "center", "flex-wrap": "wrap", "gap": "12px"},
            ):
                with solara.ToggleButtonsSingle(
                    chart_type, on_value=set_chart_type
                ):
                    solara.Button("Bar")
                    solara.Button("Line")
                    solara.Button("Pie")
                    solara.Button("Scatter")

                solara.FileDownload(
                    lambda: echarts_to_pptx(option),
                    filename=f"{chart_type.lower()}_chart.pptx",
                    mime_type=PPTX_MIME,
                    label="Download PowerPoint",
                    icon_name="mdi-microsoft-powerpoint",
                )

            solara.FigureEcharts(
                option=option,
                attributes={"style": "height: 500px; width: 100%;"},
                responsive=True,
            )
