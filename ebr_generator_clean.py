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
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict
import json


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

    # Typography
    font_family: str = 'Calibri'     # Modern, clean (fallback to Arial)
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
    chart_data: Optional[List[Tuple[float, float]]]  # (x, y) points for line chart
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


# ═══════════════════════════════════════════════════════════════════════════
# THEME RENDERERS
# ═══════════════════════════════════════════════════════════════════════════

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
        """Cost slide — chart left, table right."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        b = Builder(slide, self.T)

        b.kicker(data.kicker)
        b.title(data.title)
        b.subtitle(data.subtitle)

        # Left: Chart placeholder + annotation
        chart_left = self.T.margin_left
        chart_w = 5.8

        b.rect(chart_left, 2.65, chart_w, 3.2, 'F5F5F5')  # Chart background
        b.text(chart_left + 0.2, 2.75, chart_w - 0.4, 0.3,
              'DAILY P95 LATENCY (MS)', 10, True, self.T.purple)

        # Chart annotation below
        b.text(chart_left, 6.0, chart_w, 0.45, data.chart_annotation,
              12, True, self.T.black)

        # Right: Table
        table_left = chart_left + chart_w + 0.5
        table_w = 11.8 - chart_w - 0.5

        # Header row
        b.rect(table_left, 2.65, table_w, 0.4, self.T.purple)
        col_w = table_w / len(data.table_headers)
        for i, hdr in enumerate(data.table_headers):
            b.text(table_left + i * col_w + 0.15, 2.7, col_w - 0.2, 0.3,
                  hdr, self.T.table_header_size, True, self.T.white)

        # Data rows
        row_y = 3.05
        for row in data.table_rows:
            for i, cell in enumerate(row):
                bold = (i == 0)
                b.text(table_left + i * col_w + 0.15, row_y, col_w - 0.2, 0.35,
                      cell, 10, bold, self.T.black)
            row_y += 0.42

        # Conclusion box
        b.conclusion_box(data.conclusion, 6.35)

        b.footer(data.footer)
        b.page_num(page_num)

    def readiness(self, data: ReadinessData, page_num: int):
        """Readiness slide — full-width table."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        b = Builder(slide, self.T)

        b.kicker(data.kicker)
        b.title(data.title)
        b.subtitle(data.subtitle)

        # Table dimensions
        table_left = self.T.margin_left
        table_w = 11.8
        col_ws = [2.5, 3.0, 3.5, 2.8]  # Column widths

        # Header row
        b.rect(table_left, 2.7, table_w, 0.45, self.T.purple)
        x = table_left
        for hdr, cw in zip(data.table_headers, col_ws):
            b.text(x + 0.15, 2.77, cw - 0.2, 0.35,
                  hdr, self.T.table_header_size, True, self.T.white)
            x += cw

        # Data rows
        row_y = 3.15
        for phase_row in data.phases:
            x = table_left
            for i, (cell, cw) in enumerate(zip(phase_row, col_ws)):
                bold = (i == 0)
                b.text(x + 0.15, row_y + 0.08, cw - 0.25, 0.7,
                      cell, 10, bold, self.T.black)
                x += cw
            row_y += 0.85

        # Note below table (italic, purple)
        b.text(table_left, row_y + 0.15, table_w, 0.3, data.note,
              10, False, self.T.purple, italic=True)

        # Conclusion box
        b.conclusion_box(data.conclusion, 6.1)

        b.footer(data.footer)
        b.page_num(page_num)

    def decisions(self, data: DecisionsData, page_num: int):
        """Decisions slide — boxes + timeline."""
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

        # Right: Timeline table
        timeline_left = box_left + box_w + 0.5
        timeline_w = 11.8 - box_w - 0.5

        # Header
        b.rect(timeline_left, 2.7, timeline_w, 0.4, self.T.purple)
        b.text(timeline_left + 0.15, 2.77, 1.8, 0.3,
              'When', 10, True, self.T.white)
        b.text(timeline_left + 2.0, 2.77, timeline_w - 2.15, 0.3,
              'Evidence & decision', 10, True, self.T.white)

        # Timeline rows
        row_y = 3.1
        for when, what in data.timeline:
            b.text(timeline_left + 0.15, row_y, 1.7, 0.6,
                  when, 10, True, self.T.purple)
            b.text(timeline_left + 2.0, row_y, timeline_w - 2.15, 0.6,
                  what, 9, False, self.T.black)
            row_y += 0.72

        # Owners note
        b.text(timeline_left, row_y + 0.1, timeline_w, 0.3, data.owners,
              9, False, self.T.gray_dark, italic=True)

        # Conclusion box
        b.conclusion_box(data.conclusion, 6.15)

        b.footer(data.footer)
        b.page_num(page_num)


# ═══════════════════════════════════════════════════════════════════════════
# DECK BUILDER
# ═══════════════════════════════════════════════════════════════════════════

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
              cover: CoverData,
              performance: PerformanceData,
              cost: Optional[CostData] = None,
              readiness: Optional[ReadinessData] = None,
              decisions: Optional[DecisionsData] = None) -> Presentation:
        """Build the deck."""

        self.themes.cover(cover)
        self.page += 1

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


# ═══════════════════════════════════════════════════════════════════════════
# USAGE EXAMPLE
# ═══════════════════════════════════════════════════════════════════════════

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
            chart_data=None,  # Would need actual data points
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
    )

    path = '/Users/oldap/Downloads/fireworks_exercise/output/Northstar_EBR_Clean.pptx'
    deck.save(path)
    print(f'Saved: {path}')
    return path


if __name__ == '__main__':
    build_northstar_clean()
