#!/usr/bin/env python3
"""
Northstar EBR Adapter

Converts Northstar's ebr-slides.json format to PowerPoint via the template system.
Accepts the real Northstar array format: five content slides with kind fields.

Usage:
    python northstar_adapter.py <northstar-json> <output.pptx>
"""

import json
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

from ebr_generator_clean import (
    PerformanceData, CostData, ReadinessData, DecisionsData,
    CleanEBRDeck
)


class NorthstarAdapter:
    """Converts Northstar JSON array to template slide data."""

    def __init__(self):
        self.warnings: List[str] = []
        self.errors: List[str] = []

    def convert(self, slides: List[Dict[str, Any]]) -> tuple:
        """
        Convert Northstar slides array to slide data.

        Args:
            slides: Array of slide objects with 'kind' fields

        Returns:
            Tuple of (performance, cost, readiness, decisions) data objects
        """
        self.warnings = []
        self.errors = []

        # Index slides by kind
        by_kind = {slide['kind']: slide for slide in slides}

        # Convert each slide
        performance = None
        cost = None
        readiness = None
        decisions = None

        # summary → Performance slide (3 metrics + statement)
        if 'summary' in by_kind:
            performance = self._convert_summary(by_kind['summary'])
        else:
            self.errors.append("Missing 'summary' slide")

        # performance → Cost slide (chart + table)
        if 'performance' in by_kind:
            cost = self._convert_performance(by_kind['performance'])
        else:
            self.errors.append("Missing 'performance' slide")

        # economics → Readiness slide (table-only)
        if 'economics' in by_kind:
            readiness = self._convert_economics(by_kind['economics'])
        else:
            self.errors.append("Missing 'economics' slide")

        # pipeline → (no template slot)
        if 'pipeline' in by_kind:
            self.warnings.append("Pipeline slide not rendered (no template slot available)")

        # pilot → Decisions slide (timeline + decision boxes)
        if 'pilot' in by_kind:
            decisions = self._convert_pilot(by_kind['pilot'])
        else:
            self.errors.append("Missing 'pilot' slide")

        return performance, cost, readiness, decisions

    def _convert_summary(self, slide: Dict) -> PerformanceData:
        """Convert summary slide to PerformanceData (3 metrics + statement)."""
        # Convert metrics to tuples
        metrics = []
        for m in slide.get('metrics', []):
            metrics.append((
                m.get('value', ''),
                m.get('label', ''),
                m.get('note', '')
            ))

        # Extract notes including talking points
        notes = []
        if slide.get('note'):
            notes.append(slide['note'])
        if slide.get('talkingPoint'):
            notes.append(f"Talking point: {slide['talkingPoint']}")

        return PerformanceData(
            kicker='',
            title=slide.get('title', ''),
            subtitle=slide.get('subtitle', ''),
            metrics=metrics,
            statement=slide.get('lead', ''),
            body=slide.get('body', ''),
            conclusion=slide.get('decision', ''),
            footer='\n\n'.join(notes)
        )

    def _convert_performance(self, slide: Dict) -> CostData:
        """Convert performance slide to CostData (chart + table)."""
        # Extract chart data
        chart_data = None
        chart_annotation = ''

        if 'chart' in slide and slide['chart']:
            chart = slide['chart']
            # Convert to template format
            categories = [v['date'] for v in chart.get('values', [])]
            values = [v['value'] for v in chart.get('values', [])]

            chart_data = {
                'categories': categories,
                'values': values,
                'title': chart.get('label', ''),
                'mean_line': chart.get('target')  # Reference line
            }

            # Build annotation from chart metadata
            if chart.get('unit'):
                chart_annotation = f"Unit: {chart['unit']}"
            if chart.get('targetLabel'):
                chart_annotation += f" | Reference: {chart['targetLabel']}"

        # Extract table
        table_data = slide.get('table', {})
        headers = table_data.get('headers', [])
        rows = table_data.get('rows', [])

        # Extract notes
        notes = []
        if slide.get('note'):
            notes.append(slide['note'])
        if slide.get('talkingPoint'):
            notes.append(f"Talking point: {slide['talkingPoint']}")

        return CostData(
            kicker='',
            title=slide.get('title', ''),
            subtitle=slide.get('subtitle', ''),
            chart_data=chart_data,
            chart_annotation=chart_annotation,
            table_headers=headers,
            table_rows=rows,
            conclusion=slide.get('decision', ''),
            footer='\n\n'.join(notes)
        )

    def _convert_economics(self, slide: Dict) -> ReadinessData:
        """Convert economics slide to ReadinessData (table-only)."""
        table_data = slide.get('table', {})
        headers = table_data.get('headers', [])
        rows = table_data.get('rows', [])

        # Extract notes
        notes = []
        if slide.get('note'):
            notes.append(slide['note'])
        if slide.get('talkingPoint'):
            notes.append(f"Talking point: {slide['talkingPoint']}")

        # Note field (italic text below table)
        note_text = slide.get('lead', '') or slide.get('body', '')

        return ReadinessData(
            kicker='',
            title=slide.get('title', ''),
            subtitle=slide.get('subtitle', ''),
            table_headers=headers,
            phases=rows,
            note=note_text,
            conclusion=slide.get('decision', ''),
            footer='\n\n'.join(notes)
        )

    def _convert_pilot(self, slide: Dict) -> DecisionsData:
        """Convert pilot slide to DecisionsData (criteria + timeline)."""
        # Extract decision boxes from sections
        decision_boxes = []
        if 'sections' in slide:
            for section in slide['sections']:
                decision_boxes.append((
                    section.get('heading', ''),
                    section.get('text', '')
                ))

        # Extract timeline from commitments/actions
        timeline = []

        if 'commitments' in slide:
            for commit in slide['commitments']:
                when = commit.get('proposedDueDate', '')
                what = f"{commit.get('issue', '')}"
                if commit.get('owner'):
                    what += f"\nOwner: {commit['owner']}"
                if commit.get('evidence'):
                    what += f"\nEvidence: {commit['evidence']}"
                timeline.append((when, what))

        elif 'actions' in slide:
            for action in slide['actions']:
                when = action.get('target', '')
                what = f"{action.get('action', '')}"
                if action.get('owner'):
                    what += f"\nOwner: {action['owner']}"
                timeline.append((when, what))

        elif 'table' in slide:
            # Fallback: extract from table
            table = slide['table']
            for row in table.get('rows', []):
                if len(row) >= 2:
                    timeline.append((row[0], row[1]))

        # Extract owners
        owners = ''
        if 'commitments' in slide:
            owner_list = list(set(c.get('owner', '') for c in slide['commitments'] if c.get('owner')))
            owners = ', '.join(owner_list)

        # Extract notes including fullGates
        notes = []
        if slide.get('note'):
            notes.append(slide['note'])
        if slide.get('fullGates') and slide['fullGates']:
            notes.append("Full gates:\n" + '\n'.join(f"• {gate}" for gate in slide['fullGates']))
        if slide.get('talkingPoint'):
            notes.append(f"Talking point: {slide['talkingPoint']}")

        return DecisionsData(
            kicker='',
            title=slide.get('title', ''),
            subtitle=slide.get('subtitle', ''),
            decision_boxes=decision_boxes,
            timeline=timeline,
            owners=owners,
            conclusion=slide.get('decision', ''),
            footer='\n\n'.join(notes)
        )


def load_json(path: Path) -> List[Dict]:
    """Load Northstar JSON from file."""
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if not isinstance(data, list):
        raise ValueError(
            f"Expected array of slides, got {type(data).__name__}. "
            "Northstar output should be: [{kind: 'summary', ...}, ...]"
        )

    return data


def main():
    if len(sys.argv) != 3:
        print("Usage: python northstar_adapter.py <northstar-json> <output.pptx>")
        print("")
        print("Example:")
        print("  python northstar_adapter.py ebr-slides.json output.pptx")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    if not input_path.exists():
        print(f"❌ Error: Input file not found: {input_path}")
        sys.exit(1)

    try:
        print(f"Loading {input_path}...")
        slides = load_json(input_path)

        print("Converting to template format...")
        adapter = NorthstarAdapter()
        performance, cost, readiness, decisions = adapter.convert(slides)

        if adapter.errors:
            print("❌ Conversion errors:")
            for error in adapter.errors:
                print(f"   - {error}")
            sys.exit(1)

        if adapter.warnings:
            print("⚠️  Warnings:")
            for warning in adapter.warnings:
                print(f"   - {warning}")

        print("Generating PowerPoint deck...")
        deck = CleanEBRDeck()
        prs = deck.build(
            performance=performance,
            cost=cost,
            readiness=readiness,
            decisions=decisions,
            skip_cover=True
        )
        deck.save(str(output_path))

        size = output_path.stat().st_size
        print("")
        print(f"✅ Success! Deck created: {output_path}")
        print(f"   File size: {size:,} bytes")
        print("")
        print("📋 Verification Checklist:")
        print("   1. Open in PowerPoint")
        print("   2. Check 4 content slides (no cover, pipeline skipped)")
        print("   3. Verify chart is editable (performance → cost slide)")
        print("   4. Verify tables are editable")
        print("   5. Check speaker notes for talking points and full gates")

    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
