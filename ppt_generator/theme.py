"""
Theme constants for PPT generation: colors, dimensions, and styling.
"""

from pptx.util import Inches
from pptx.dml.color import RGBColor


# ─── Color Palette ─────────────────────────────────────────────────────────

COLORS = {
    "dark": RGBColor(0x0F, 0x17, 0x2A),
    "slate800": RGBColor(0x1E, 0x29, 0x3B),
    "slate600": RGBColor(0x47, 0x55, 0x69),
    "slate400": RGBColor(0x94, 0xA3, 0xB8),
    "text": RGBColor(0x1E, 0x29, 0x3B),
    "text_secondary": RGBColor(0x47, 0x55, 0x69),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "primary": RGBColor(0x25, 0x63, 0xEB),
    "primary_light": RGBColor(0x3B, 0x82, 0xF6),
    "accent": RGBColor(0x05, 0x96, 0x69),
    "accent_light": RGBColor(0xD1, 0xFA, 0xE5),
    "warn": RGBColor(0xD9, 0x77, 0x06),
    "bg_light": RGBColor(0xF8, 0xFA, 0xFC),
    "bg_card": RGBColor(0xF1, 0xF5, 0xF9),
}


# ─── Slide Dimensions ──────────────────────────────────────────────────────

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.8)
CONTENT_W = Inches(11.7)
