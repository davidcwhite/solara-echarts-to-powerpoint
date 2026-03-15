"""PPTX theme definitions: color palettes, fonts, and layout preferences."""

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class PptxTheme:
    name: str
    palette: List[str]
    title_font: str
    title_size: int
    label_font: str
    label_size: int
    legend_position: str  # "bottom" or "right"


DEFAULT = PptxTheme(
    name="Default",
    palette=["5470c6", "91cc75", "fac858", "ee6666", "73c0de", "3ba272",
             "fc8452", "9a60b4", "ea7ccc"],
    title_font="Calibri",
    title_size=24,
    label_font="Calibri",
    label_size=11,
    legend_position="bottom",
)

CORPORATE = PptxTheme(
    name="Corporate",
    palette=["2b579a", "4f81bd", "7eaed2", "a5c8e1", "bfd8ec", "d6e6f4",
             "3b6e8f", "5a8fb0", "8bb5cc"],
    title_font="Arial",
    title_size=22,
    label_font="Arial",
    label_size=10,
    legend_position="right",
)

VIBRANT = PptxTheme(
    name="Vibrant",
    palette=["e6194b", "3cb44b", "ffe119", "4363d8", "f58231", "911eb4",
             "42d4f4", "f032e6", "bfef45"],
    title_font="Helvetica",
    title_size=26,
    label_font="Helvetica",
    label_size=12,
    legend_position="bottom",
)

MONOCHROME = PptxTheme(
    name="Monochrome",
    palette=["1a1a2e", "3d3d5c", "5e5e85", "8080aa", "a3a3c2", "c6c6db",
             "d9d9e6", "ececf2", "4a4a6a"],
    title_font="Calibri",
    title_size=24,
    label_font="Calibri",
    label_size=11,
    legend_position="bottom",
)

THEMES: Dict[str, PptxTheme] = {
    "Default": DEFAULT,
    "Corporate": CORPORATE,
    "Vibrant": VIBRANT,
    "Monochrome": MONOCHROME,
}
