"""Solara app: display ECharts charts and export to PowerPoint."""

import solara

from converter import echarts_to_pptx, is_exportable
from theme import THEMES

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
    "Stacked Bar": {
        "title": {"text": "Stacked Revenue"},
        "tooltip": {"trigger": "axis"},
        "legend": {"data": ["Online", "In-Store"]},
        "xAxis": {
            "type": "category",
            "data": ["Jan", "Feb", "Mar", "Apr"],
        },
        "yAxis": {"type": "value"},
        "series": [
            {
                "name": "Online",
                "type": "bar",
                "stack": "total",
                "data": [320, 302, 341, 374],
            },
            {
                "name": "In-Store",
                "type": "bar",
                "stack": "total",
                "data": [120, 132, 101, 134],
            },
        ],
    },
    "Horizontal Bar": {
        "title": {"text": "Feature Satisfaction"},
        "tooltip": {},
        "xAxis": {"type": "value"},
        "yAxis": {
            "type": "category",
            "data": ["Speed", "Reliability", "UX", "Support"],
        },
        "series": [
            {"name": "Score", "type": "bar", "data": [85, 92, 78, 88]},
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
    "Area": {
        "title": {"text": "Website Traffic"},
        "tooltip": {"trigger": "axis"},
        "legend": {"data": ["Desktop", "Mobile"]},
        "xAxis": {
            "type": "category",
            "data": ["Mon", "Tue", "Wed", "Thu", "Fri"],
        },
        "yAxis": {"type": "value"},
        "series": [
            {
                "name": "Desktop",
                "type": "line",
                "areaStyle": {},
                "data": [820, 932, 901, 934, 1290],
            },
            {
                "name": "Mobile",
                "type": "line",
                "areaStyle": {},
                "data": [620, 732, 801, 834, 990],
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
    "Donut": {
        "title": {"text": "Browser Usage"},
        "tooltip": {"trigger": "item"},
        "series": [
            {
                "name": "Browsers",
                "type": "pie",
                "radius": ["40%", "70%"],
                "data": [
                    {"name": "Chrome", "value": 65},
                    {"name": "Firefox", "value": 12},
                    {"name": "Safari", "value": 15},
                    {"name": "Edge", "value": 8},
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
    "Radar": {
        "title": {"text": "Team Skill Assessment"},
        "tooltip": {},
        "radar": {
            "indicator": [
                {"name": "Engineering", "max": 100},
                {"name": "Design", "max": 100},
                {"name": "Marketing", "max": 100},
                {"name": "Sales", "max": 100},
                {"name": "Support", "max": 100},
            ]
        },
        "series": [
            {
                "type": "radar",
                "data": [
                    {
                        "value": [90, 65, 70, 80, 75],
                        "name": "Team A",
                    },
                    {
                        "value": [70, 85, 90, 60, 80],
                        "name": "Team B",
                    },
                ],
            }
        ],
    },
    "Sunburst (unsupported)": {
        "title": {"text": "Sunburst -- no PPTX equivalent"},
        "series": [
            {
                "type": "sunburst",
                "data": [
                    {
                        "name": "A",
                        "children": [
                            {"name": "A1", "value": 10},
                            {"name": "A2", "value": 20},
                        ],
                    },
                    {
                        "name": "B",
                        "children": [
                            {"name": "B1", "value": 30},
                            {"name": "B2", "value": 15},
                        ],
                    },
                ],
                "radius": ["15%", "80%"],
            }
        ],
    },
}

PPTX_MIME = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


@solara.component
def Page():
    chart_type, set_chart_type = solara.use_state("Bar")
    theme_name, set_theme_name = solara.use_state("Default")

    option = CHARTS[chart_type]
    theme = THEMES[theme_name]
    exportable = is_exportable(option)

    with solara.AppLayout(title="ECharts to PowerPoint"):
        with solara.Card(margin=2):
            with solara.Row(
                justify="space-between",
                style={
                    "align-items": "center",
                    "flex-wrap": "wrap",
                    "gap": "12px",
                },
            ):
                with solara.ToggleButtonsSingle(
                    chart_type, on_value=set_chart_type
                ):
                    for label in CHARTS:
                        solara.Button(label)

                solara.Select(
                    label="PPTX Theme",
                    value=theme_name,
                    values=list(THEMES.keys()),
                    on_value=set_theme_name,
                    style={"min-width": "160px"},
                )

                if exportable:
                    solara.FileDownload(
                        lambda: echarts_to_pptx(option, theme=theme),
                        filename=f"{chart_type.lower().replace(' ', '_')}_chart.pptx",
                        mime_type=PPTX_MIME,
                        label="Download PowerPoint",
                        icon_name="mdi-microsoft-powerpoint",
                    )
                else:
                    solara.Button(
                        "Export not available for this chart type",
                        disabled=True,
                        icon_name="mdi-microsoft-powerpoint",
                    )

            solara.FigureEcharts(
                option=option,
                attributes={"style": "height: 500px; width: 100%;"},
                responsive=True,
            )
