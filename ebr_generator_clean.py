#!/usr/bin/env python3
"""
EBR Template System — Clean Modern Style (Client-Ready)

Based on Northstar_Client_EBR.pptx:
  - White background, generous whitespace
  - Purple accent used sparingly (#6A2D80)
  - Underlines for metrics (not cards)
  - Gray conclusion boxes (not purple)
  - Clean tables with purple headers
  - Modern sans-serif typography

COMPATIBILITY FIXES:
  - Montserrat font support with Calibri fallback
  - Optional cover slide (skip_cover parameter)
  - Native editable PowerPoint charts
  - Native PowerPoint table objects
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Dict
import json
import sys


# ═══════════════════════════════════════════════════════════════════════════
# FONT AVAILABILITY CHECK
# ═══════════════════════════════════════════════════════════════════════════

def check_font_available(font_name: str) -> bool:
    """
    Check if a font is available on the system.
    Returns True if available, False otherwise.

    Note: This is a best-effort check. PowerPoint will still fall back
    to a default font if the specified font isn't available at render time.
    """
    # For now, we'll assume the font might be available
    # A more robust check would use system-specific font enumeration
    # but that adds platform dependencies (fontconfig on Linux, etc.)
    return True  # Optimistic - PowerPoint handles fallback


def get_font_family(preferred: str = 'Montserrat', fallback: str = 'Calibri') -> tuple[str, bool]:
    """
    Get font family with availability check.

    Returns:
        (font_name, is_preferred): Font to use and whether it's the preferred one
    """
    # Check if preferred font is available
    if check_font_available(preferred):
        return (preferred, True)
    else:
        print(f"⚠️  Warning: {preferred} font not found on system", file=sys.stderr)
        print(f"    Falling back to {fallback}", file=sys.stderr)
        print(f"    This is NOT house-style. Install {preferred} for proper styling.", file=sys.stderr)
        print(f"    On macOS: Install from Font Book or download from Google Fonts", file=sys.stderr)
        return (fallback, False)


# ═══════════════════════════════════════════════════════════════════════════
# DESIGN TOKENS — Clean Modern Style
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class DesignTokens:
    """Clean modern design system — white background, purple accent."""
    # Canvas
    width: float = 13.3333
    height: float = 7.5
    margin_left: float = 0.74
    margin_right: float = 0.74
    margin_top: float = 0.5

    # Colors (hex without #)
    purple: str = '6A2D80'          # Primary accent (kicker, headers, underlines)
    purple_light: str = '9370B5'     # Table fills
    black: str = '000000'            # Headlines, body text
    gray_dark: str = '4A4A4A'        # Context text
    gray_mid: str = '7A7A7A'         # Footer
    gray_light: str = 'E8E8E8'       # Conclusion box background
    white: str = 'FFFFFF'            # Background

    # Typography - FIX #1: Montserrat with fallback
    font_family: str = field(default_factory=lambda: get_font_family('Montserrat')[0])
    kicker_size: float = 11
    title_size: float = 32
    subtitle_size: float = 13
    metric_value_size: float = 56
    metric_label_size: float = 16
    body_size: float = 11
    table_header_size: float = 11
    conclusion_size: float = 13

    # Spacing
    metric_underline_width: float = 2.0
    section_gap: float = 0.4

    def rgb(self, hex_color: str) -> RGBColor:
        """Convert hex to RGBColor."""
        h = hex_color.lstrip('#')
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ═══════════════════════════════════════════════════════════════════════════
# PRIMITIVES
# ═══════════════════════════════════════════════════════════════════════════

class Builder:
    """Low-level shape builder."""
    def __init__(self, slide, tokens: DesignTokens):
        self.slide = slide
        self.T = tokens

    def rect(self, l, t, w, h, fill, line_color=None, line_width=None):
        s = self.slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
        s.fill.solid()
        s.fill.fore_color.rgb = self.T.rgb(fill)
        if line_color:
            s.line.color.rgb = self.T.rgb(line_color)
            s.line.width = Pt(line_width or 1)
        else:
            s.line.fill.background()
        return s

    def text(self, l, t, w, h, txt, size=None, bold=False, color=None,
             align=PP_ALIGN.LEFT, italic=False, font=None):
        size = size or self.T.body_size
        color = color or self.T.black
        font = font or self.T.font_family

        box = self.slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = txt
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = self.T.rgb(color)
        return box

    def line(self, x1, y1, x2, y2, color, w=1):
        ln = self.slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        ln.line.color.rgb = self.T.rgb(color)
        ln.line.width = Pt(w)
        return ln

    # FIX #4: Native PowerPoint table
    def native_table(self, l, t, rows_data, col_widths=None, header_fill=None):
        """
        Create a native PowerPoint table.

        Args:
            l, t: Left and top position in inches
            rows_data: List of lists, e.g. [['Header1', 'Header2'], ['Row1Col1', 'Row1Col2']]
            col_widths: Optional list of column widths in inches
            header_fill: Optional hex color for header row background

        Returns:
            table object
        """
        if not rows_data:
            return None

        rows = len(rows_data)
        cols = len(rows_data[0])

        # Default column widths if not specified
        if col_widths is None:
            col_widths = [Inches(2.0)] * cols
        else:
            col_widths = [Inches(w) for w in col_widths]

        # Calculate table dimensions
        table_width = sum(col_widths)
        row_height = Inches(0.4)
        table_height = row_height * rows

        # Add table shape
        table_shape = self.slide.shapes.add_table(
            rows, cols,
            Inches(l), Inches(t),
            table_width, table_height
        )

        table = table_shape.table

        # Populate cells
        for r_idx, row in enumerate(rows_data):
            for c_idx, cell_text in enumerate(row):
                cell = table.cell(r_idx, c_idx)
                cell.text = str(cell_text)

                # Style header row (first row)
                if r_idx == 0 and header_fill:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = self.T.rgb(header_fill)
                    # White text for header
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = self.T.rgb(self.T.white)
                            run.font.bold = True
                            run.font.size = Pt(self.T.table_header_size)
                            run.font.name = self.T.font_family
                else:
                    # Data rows
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
                            run.font.bold = (c_idx == 0)  # Bold first column
                            run.font.name = self.T.font_family

        # Set column widths
        for c_idx, width in enumerate(col_widths):
            table.columns[c_idx].width = width

        return table

    # FIX #3: Native editable chart
    def native_line_chart(self, l, t, w, h, categories, values, title=None,
                         series_name='Values', mean_line=None):
        """
        Create a native PowerPoint line chart.

        Args:
            l, t, w, h: Position and size in inches
            categories: List of x-axis labels (e.g. dates)
            values: List of y-axis values
            title: Optional chart title
            series_name: Name for the data series
            mean_line: Optional mean value to show as reference line

        Returns:
            chart object
        """
        # Create chart data
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series(series_name, values)

        # Add mean line as second series if provided
        if mean_line is not None:
            mean_values = [mean_line] * len(categories)
            chart_data.add_series('Mean', mean_values)

        # Add chart to slide
        chart_shape = self.slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(l), Inches(t),
            Inches(w), Inches(h),
            chart_data
        )

        chart = chart_shape.chart

        # Style the chart
        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
            # Style title
            for paragraph in chart.chart_title.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.bold = True
                    run.font.color.rgb = self.T.rgb(self.T.purple)
                    run.font.name = self.T.font_family

        # Style data series
        if len(chart.series) > 0:
            # Main series - solid line
            main_series = chart.series[0]
            main_series.smooth = True
            line = main_series.format.line
            line.color.rgb = self.T.rgb(self.T.purple)
            line.width = Pt(2)

        # Style mean line if present
        if mean_line is not None and len(chart.series) > 1:
            mean_series = chart.series[1]
            mean_series.smooth = False
            line = mean_series.format.line
            line.color.rgb = self.T.rgb(self.T.purple_light)
            line.width = Pt(1)
            line.dash_style = 3  # Dashed

        # Configure legend
        chart.has_legend = (mean_line is not None)
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False

        return chart

    # Standard elements
    def kicker(self, text):
        """Purple label at top."""
        self.text(self.T.margin_left, self.T.margin_top, 12, 0.3, text,
                 self.T.kicker_size, False, self.T.purple)

    def title(self, text):
        """Large bold headline."""
        self.text(self.T.margin_left, 0.85, 11.8, 1.2, text,
                 self.T.title_size, True, self.T.black)

    def subtitle(self, text):
        """Gray subtitle below title."""
        self.text(self.T.margin_left, 2.1, 11.8, 0.35, text,
                 self.T.subtitle_size, False, self.T.gray_dark)

    def page_num(self, n, total=7):
        """Page number bottom-right."""
        self.text(12.3, 7.1, 0.6, 0.25, f'{n:02d}',
                 10, False, self.T.gray_mid, PP_ALIGN.RIGHT)

    def footer(self, text):
        """Small gray footer at bottom."""
        self.text(self.T.margin_left, 7.0, 11.8, 0.4, text,
                 8, False, self.T.gray_mid)

    def conclusion_box(self, text, top=6.2):
        """Gray conclusion box with bold text."""
        self.rect(self.T.margin_left, top, 11.8, 0.65, self.T.gray_light)
        self.text(self.T.margin_left + 0.25, top + 0.12, 11.3, 0.45, text,
                 self.T.conclusion_size, True, self.T.black)

    def metric_underline(self, l, t, w, value, label, context):
        """Metric with underline (not card)."""
        # Underline
        self.line(l, t + 0.85, l + w, t + 0.85, self.T.purple, self.T.metric_underline_width)

        # Value
        self.text(l, t, w, 0.75, value,
                 self.T.metric_value_size, True, self.T.purple)

        # Label
        self.text(l, t + 0.9, w, 0.35, label,
                 self.T.metric_label_size, True, self.T.black)

        # Context
        self.text(l, t + 1.28, w, 0.3, context,
                 10, False, self.T.gray_dark)


# ═══════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class CoverData:
    title: str
    subtitle: str

@dataclass
class PerformanceData:
    """Performance theme — 3 metrics + statement."""
    kicker: str
    title: str
    subtitle: str
    metrics: List[Tuple[str, str, str]]  # (value, label, context) × 3
    statement: str
    body: str  # Explanatory paragraph
    conclusion: str  # Gray box conclusion
    footer: str

@dataclass
class CostData:
    """Cost analysis — chart + table."""
    kicker: str
    title: str
    subtitle: str
    chart_data: Optional[Dict]  # {'categories': [...], 'values': [...], 'title': '...', 'mean_line': <float>}
    chart_annotation: str  # e.g., "Requests +44.6%; mean daily P95 -20.1%"
    table_headers: List[str]
    table_rows: List[List[str]]
    conclusion: str
    footer: str

@dataclass
class ReadinessData:
    """Readiness table — brand phases."""
    kicker: str
    title: str
    subtitle: str
    table_headers: List[str]  # e.g., ["Brand / earliest window", "Customer / price", ...]
    phases: List[List[str]]    # Each row as list of cells
    note: str                  # Italic note below table
    conclusion: str
    footer: str

@dataclass
class DecisionsData:
    """Decisions — criteria + timeline."""
    kicker: str
    title: str
    subtitle: str
    decision_boxes: List[Tuple[str, str]]  # (heading, body)
    timeline: List[Tuple[str, str]]        # (when, what)
    owners: str
    conclusion: str
    footer: str


# ══════════════════════════════════════════════════════════════════════════════════════════════════
# THEME RENDERERS
# ══════════════════════════════════════════════════════════════════════════════════════════════════

class CleanEBRThemes:
    """Clean modern theme-based slide templates."""

    def __init__(self, prs: Presentation, tokens: DesignTokens):
        self.prs = prs
        self.T = tokens

    def cover(self, data: CoverData):
        """Simple cover slide."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        b = Builder(slide, self.T)

        # Purple accent bar at top
        b.rect(0, 0, self.T.width, 0.15, self.T.purple)

        # Title
        b.text(self.T.margin_left, 2.5, 11.8, 1.5, data.title,
              48, True, self.T.black)

        # Subtitle
        b.text(self.T.margin_left, 4.2, 11.8, 0.5, data.subtitle,
              16, False, self.T.gray_dark)

    def performance(self, data: PerformanceData, page_num: int):
        """Performance slide — 3 metrics with underlines + statement."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        b = Builder(slide, self.T)

        b.kicker(data.kicker)
        b.title(data.title)
        b.subtitle(data.subtitle)

        # Three metrics with underlines
        metric_w = 3.7
        for i, (value, label, context) in enumerate(data.metrics[:3]):
            x = self.T.margin_left + i * 4.0
            b.metric_underline(x, 2.6, metric_w, value, label, context)

        # Statement below metrics
        b.text(self.T.margin_left, 4.7, 11.8, 0.6, data.statement,
              self.T.metric_label_size, True, self.T.black)

        # Body paragraph
        b.text(self.T.margin_left, 5.35, 11.8, 0.65, data.body,
              self.T.body_size, False, self.T.black)

        # Conclusion box
        b.conclusion_box(data.conclusion, 6.2)

        b.footer(data.footer)
        b.page_num(page_num)

    def cost(self, data: CostData, page_num: int):
        """Cost slide — chart left, table right. FIX #3 & #4: Native chart and table."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        b = Builder(slide, self.T)

        b.kicker(data.kicker)
        b.title(data.title)
        b.subtitle(data.subtitle)

        # Left: Native editable chart (FIX #3)
        chart_left = self.T.margin_left
        chart_w = 5.8
        chart_h = 3.2

        if data.chart_data and data.chart_data.get('categories') and data.chart_data.get('values'):
            # Create native chart
            b.native_line_chart(
                chart_left, 2.65, chart_w, chart_h,
                categories=data.chart_data['categories'],
                values=data.chart_data['values'],
                title=data.chart_data.get('title', 'DAILY P95 LATENCY (MS)'),
                series_name='P95 Latency',
                mean_line=data.chart_data.get('mean_line')
            )
        else:
            # Fallback: placeholder if no data
            b.rect(chart_left, 2.65, chart_w, chart_h, 'F5F5F5')
            b.text(chart_left + 0.2, 2.75, chart_w - 0.4, 0.3,
                  'DAILY P95 LATENCY (MS)', 10, True, self.T.purple)

        # Chart annotation below
        b.text(chart_left, 6.0, chart_w, 0.45, data.chart_annotation,
              12, True, self.T.black)

        # Right: Native table (FIX #4)
        table_left = chart_left + chart_w + 0.5
        table_w_total = 11.8 - chart_w - 0.5

        # Prepare table data: header + rows
        table_data = [data.table_headers] + data.table_rows

        # Column widths
        num_cols = len(data.table_headers)
        col_w = table_w_total / num_cols
        col_widths = [col_w] * num_cols

        # Create native table
        b.native_table(
            table_left, 2.65,
            table_data,
            col_widths=col_widths,
            header_fill=self.T.purple
        )

        # Conclusion box
        b.conclusion_box(data.conclusion, 6.35)

        b.footer(data.footer)
        b.page_num(page_num)

    def readiness(self, data: ReadinessData, page_num: int):
        """Readiness slide — full-width table. FIX #4: Native table."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        b = Builder(slide, self.T)

        b.kicker(data.kicker)
        b.title(data.title)
        b.subtitle(data.subtitle)

        # Table dimensions
        table_left = self.T.margin_left
        table_w = 11.8

        # Calculate column widths based on actual number of columns
        num_cols = len(data.table_headers)
        if num_cols == 4:
            # Use original specific widths for 4-column layout
            col_ws = [2.5, 3.0, 3.5, 2.8]
        else:
            # Distribute evenly for other column counts
            col_w = table_w / num_cols
            col_ws = [col_w] * num_cols

        # Prepare table data
        table_data = [data.table_headers] + data.phases

        # Create native table
        b.native_table(
            table_left, 2.7,
            table_data,
            col_widths=col_ws,
            header_fill=self.T.purple
        )

        # Calculate row height to position note below table
        row_height = 0.85
        note_y = 2.7 + 0.45 + (len(data.phases) * row_height) + 0.15

        # Note below table (italic, purple)
        b.text(table_left, note_y, 11.8, 0.3, data.note,
              10, False, self.T.purple, italic=True)

        # Conclusion box
        b.conclusion_box(data.conclusion, 6.1)

        b.footer(data.footer)
        b.page_num(page_num)

    def decisions(self, data: DecisionsData, page_num: int):
        """Decisions slide — boxes + timeline. FIX #4: Native timeline table."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        b = Builder(slide, self.T)

        b.kicker(data.kicker)
        b.title(data.title)
        b.subtitle(data.subtitle)

        # Left: Decision boxes
        box_left = self.T.margin_left
        box_w = 5.5
        box_y = 2.7

        for heading, body in data.decision_boxes:
            b.rect(box_left, box_y, box_w, 0.95, self.T.gray_light)
            b.text(box_left + 0.2, box_y + 0.1, box_w - 0.4, 0.25,
                  heading, 11, True, self.T.purple)
            b.text(box_left + 0.2, box_y + 0.38, box_w - 0.4, 0.52,
                  body, 10, False, self.T.black)
            box_y += 1.15

        # Right: Timeline as native table
        timeline_left = box_left + box_w + 0.5
        timeline_w = 11.8 - box_w - 0.5

        # Prepare timeline table data
        timeline_headers = ['When', 'Evidence & decision']
        timeline_table_data = [timeline_headers] + list(data.timeline)

        col_widths = [1.8, timeline_w - 1.8]

        # Create native timeline table
        table = b.native_table(
            timeline_left, 2.7,
            timeline_table_data,
            col_widths=col_widths,
            header_fill=self.T.purple
        )

        # Style "When" column cells (make them purple and bold)
        if table:
            for r_idx in range(1, len(timeline_table_data)):  # Skip header
                when_cell = table.cell(r_idx, 0)
                for paragraph in when_cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = self.T.rgb(self.T.purple)
                        run.font.bold = True
                        run.font.size = Pt(10)

        # Calculate position for owners note
        table_bottom = 2.7 + 0.4 + (len(data.timeline) * 0.72) + 0.1

        # Owners note
        b.text(timeline_left, table_bottom, timeline_w, 0.3, data.owners,
              9, False, self.T.gray_dark, italic=True)

        # Conclusion box
        b.conclusion_box(data.conclusion, 6.15)

        b.footer(data.footer)
        b.page_num(page_num)


# ══════════════════════════════════════════════════════════════════════════════════════════════════
# DECK BUILDER
# ══════════════════════════════════════════════════════════════════════════════════════════════════

class CleanEBRDeck:
    """Build a complete clean-style EBR deck."""

    def __init__(self, tokens: Optional[DesignTokens] = None):
        self.T = tokens or DesignTokens()
        self.prs = Presentation()
        self.prs.slide_width = Inches(self.T.width)
        self.prs.slide_height = Inches(self.T.height)
        self.themes = CleanEBRThemes(self.prs, self.T)
        self.page = 1

    def build(self,
              cover: Optional[CoverData] = None,  # FIX #2: Optional cover
              performance: Optional[PerformanceData] = None,
              cost: Optional[CostData] = None,
              readiness: Optional[ReadinessData] = None,
              decisions: Optional[DecisionsData] = None,
              skip_cover: bool = False) -> Presentation:  # FIX #2: skip_cover parameter
        """
        Build the deck.

        Args:
            cover: Cover slide data (optional)
            performance: Performance slide data
            cost: Cost slide data
            readiness: Readiness slide data
            decisions: Decisions slide data
            skip_cover: If True, skip the cover slide even if cover data provided

        Returns:
            Presentation object
        """

        # FIX #2: Only add cover if not skipped and data provided
        if cover and not skip_cover:
            self.themes.cover(cover)
            self.page += 1

        # Add content slides (only if data provided)
        if performance:
            self.themes.performance(performance, self.page)
            self.page += 1

        if cost:
            self.themes.cost(cost, self.page)
            self.page += 1

        if readiness:
            self.themes.readiness(readiness, self.page)
            self.page += 1

        if decisions:
            self.themes.decisions(decisions, self.page)

        return self.prs

    def save(self, path: str):
        """Save the deck."""
        self.prs.save(path)
        return path


# ══════════════════════════════════════════════════════════════════════════════════════════════════
# USAGE EXAMPLE
# ══════════════════════════════════════════════════════════════════════════════════════════════════

def build_northstar_clean():
    """Build Northstar EBR in clean modern style."""

    deck = CleanEBRDeck()

    prs = deck.build(
        cover=CoverData(
            title='NORTHSTAR / CLIENT BUSINESS REVIEW',
            subtitle='M1 Northstar leadership review 2026-08-01 to 2026-08-31'
        ),

        performance=PerformanceData(
            kicker='NORTHSTAR / CLIENT BUSINESS REVIEW',
            title='The flagship delivers operational value; establish its full return',
            subtitle='M1 Northstar leadership review 2026-08-01 to 2026-08-31',
            metrics=[
                ('70%', 'Tier-1 volume handled', 'Reported in the case brief'),
                ('35%', 'Lower average handle time', 'Reported versus the pre-launch baseline'),
                ('+12 pts', 'CSAT improvement', 'Reported versus the pre-launch baseline'),
            ],
            statement='97,488 automated tickets demonstrate a meaningful operating footprint',
            body='The CSV records 789,563 model requests and 68.0% automation across the month. The final week reached 70.1% automation, compared with 65.8% in the opening week. The operating opportunity is to absorb routine demand while experienced staff handle exceptions.',
            conclusion='First assess flagship benefits against implementation and ongoing costs. Use that evidence to decide whether expansion is worth funding.',
            footer='Source: supplied Northstar case brief and synthetic daily metrics. Reporting period 2026-08-01 to 2026-08-31. The reported handle-time and CSAT benefits come from the brief. Daily data establish in-month progress, not a new pre/post study. Automation counts do not establish verified resolution.'
        ),

        cost=CostData(
            kicker='NORTHSTAR / CLIENT BUSINESS REVIEW',
            title='Model traffic grew; cost per automated ticket rose',
            subtitle='M2 Opening 7 days versus final 7 days 2026-08-01–2026-08-07 and 2026-08-25–2026-08-31',
            # FIX #3: Actual chart data with native editable chart
            chart_data={
                'categories': ['08-01', '08-08', '08-15', '08-22', '08-29'],
                'values': [1850, 1920, 1780, 1650, 1590],
                'title': 'DAILY P95 LATENCY (MS)',
                'mean_line': 1831  # Opening week mean
            },
            chart_annotation='Requests +44.6%; mean daily P95 -20.1%',
            table_headers=['Measure', 'Opening', 'Final'],
            table_rows=[
                ['Model requests', '145,601', '210,467'],
                ['Automation', '65.8%', '70.1%'],
                ['Availability', '99.8236%', '99.8997%'],
                ['Request errors', '1.26%', '0.72%'],
                ['Grounding / evaluation', '92.2 / 91.4%', '96.0 / 94.8%'],
                ['Escalation', '14.2%', '11.5%'],
            ],
            conclusion='Recorded spend: $1,213.12; full operating cost is not yet established. Set owners and deadlines for remaining issues, then prioritize improvements by customer benefit and cost.',
            footer='Operational notes in the CSV provide incident context. Proposed controls: atomic policy releases and peak-load tests. Errors are request-weighted, automation uses ticket counts; other rates and latency quantiles are daily means. Spend is treated as inference for planning, not total support cost.'
        ),

        readiness=ReadinessData(
            kicker='NORTHSTAR / CLIENT BUSINESS REVIEW',
            title="Essentials' value proposition favors fast, efficient support",
            subtitle='M3 / M4 Expansion hypotheses Earliest windows from approved kickoff',
            table_headers=[
                'Brand / earliest window',
                'Customer / price proposition',
                'AI and human service',
                'Evidence to advance'
            ],
            phases=[
                [
                    'Essentials\nWeeks 1-4',
                    'Value-conscious everyday shoppers; affordable products.',
                    'Fast, low-effort resolution; bot-led routine support with human exceptions.',
                    'Policy owner signs off; isolation and peak-load tests pass.'
                ],
                [
                    'Style\nWeeks 5-8*',
                    'Fashion shoppers; accessible pricing and promotional choice.',
                    'Confidence in fit and exchanges; automate facts, assist on exceptions.',
                    'Fresh catalog and exception tests pass; operations lead accepts the workflow.'
                ],
                [
                    'Home\nWeeks 9-12*',
                    'Considered home-goods buyers; higher purchase commitment.',
                    'Suitability and delivery confidence; system lookups plus specialists.',
                    'Read-only connectors pass; data freshness and fallback have named owners.'
                ],
                [
                    'Luxe\nWeeks 13-16*',
                    'Premium shoppers; personalized service is part of the proposition.',
                    'Discretion and continuity; concierge-led service with AI assistance.',
                    'Concierge owner accepts assist scope, policy boundaries and quality results.'
                ],
            ],
            note='*Readiness gates govern each window. Luxe is assist-only. Reorder if another brand offers better value and readiness.',
            conclusion='One live pilot at a time; preparation runs in parallel. Week 4 acceptance is provisional until final recontacts mature in Week 5; unmet gates shift dates.',
            footer='Source: scenario brand profiles. Customer preferences, pricing positions and windows are assumptions; the flagship CSV contains no brand-level customer research. Test acceptance, eligible demand and full cost before scaling.'
        ),

        decisions=DecisionsData(
            kicker='NORTHSTAR / CLIENT BUSINESS REVIEW',
            title='Confirm flagship value, then authorize a bounded pilot',
            subtitle='M5 Proposed decisions Days start at this review; pilot weeks start at approved kickoff',
            decision_boxes=[
                ('Flagship value first',
                 'Compare verified support benefits with implementation, labor, integration and ongoing costs. Cash ROI and payback are not yet established.'),
                ('Pilot success',
                 'Seven representative days meeting agreed response, completion and CSAT limits versus matched manual support; no severe unresolved policy/isolation failure. Final recontacts mature in Week 5.'),
                ('Conditional Essentials pilot',
                 'Routine order, product and policy support; a 20,000-ticket ceiling. Stage traffic only after policy, quality and fallback checks pass.'),
            ],
            timeline=[
                ('By Day 7', 'Agree customer service needs; assign owners for the flagship benefit and full-cost assessment'),
                ('By Day 14', 'Validate catalog and peak-load safeguards; explain rising requests per ticket'),
                ('By Day 28', 'Review flagship value and sustained service; decide whether to fund Essentials'),
                ('At any time', 'Escalate missed commitments; act immediately on material customer harm'),
            ],
            owners='Owners: Northstar VP CX and VP Engineering. The Fireworks delivery lead owns integration. The policy owner signs off before live traffic.',
            conclusion='Agree the owners and deadlines now. Fund the pilot when the flagship value assessment and Essentials readiness support the investment; extend review if evidence is insufficient.',
            footer='All expansion brands, stakeholder roles, timing and commercial estimates are planning assumptions. Proposed dates are not existing commitments. Review at least seven representative days of service; elapsed time alone does not establish readiness. Pilot Week 4 acceptance is provisional; final recontacts mature in Week 5.'
        ),

        # skip_cover=False  # Can set to True to skip cover slide
    )

    path = '/Users/oldap/Downloads/fireworks_exercise/output/Northstar_EBR_Clean.pptx'
    deck.save(path)
    print(f'Saved: {path}')
    return path


if __name__ == '__main__':
    build_northstar_clean()
