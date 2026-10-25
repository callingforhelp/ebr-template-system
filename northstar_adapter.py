#!/usr/bin/env python3
"""
Northstar EBR Adapter

Converts Northstar's ebr-slides.json format to the template system's input format.
Maps the Northstar dashboard output directly to slide data classes for rendering.

Usage:
    python northstar_adapter.py <northstar-json> <output.pptx>
"""

import json
import sys
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple
from dataclasses import dataclass

from ebr_generator_clean import (
    CoverData, PerformanceData, CostData,
    ReadinessData, DecisionsData, CleanEBRDeck
)


@dataclass
class ConvertedData:
    """Converted slide data ready for rendering."""
    cover: CoverData
    performance: PerformanceData
    cost: CostData
    readiness: ReadinessData
    decisions: DecisionsData


@dataclass
class AdapterResult:
    """Result of adapter conversion with validation info."""
    data: Optional[ConvertedData]
    warnings: List[str]
    errors: List[str]

    @property
    def is_valid(self) -> bool:
        return self.data is not None and len(self.errors) == 0


class NorthstarAdapter:
    """Converts Northstar JSON to template slide data."""

    def __init__(self):
        self.warnings: List[str] = []
        self.errors: List[str] = []

    def convert(self, northstar_json: Dict[str, Any]) -> AdapterResult:
        """
        Convert Northstar JSON to slide data classes.

        Args:
            northstar_json: Parsed Northstar ebr-slides.json

        Returns:
            AdapterResult with converted data or errors
        """
        self.warnings = []
        self.errors = []

        try:
            # Extract and convert each slide type
            cover_data = self._convert_cover(northstar_json.get('cover'))
            performance_data = self._convert_performance(northstar_json.get('performance'))
            cost_data = self._convert_cost(northstar_json.get('cost'))
            readiness_data = self._convert_readiness(northstar_json.get('readiness'))
            decisions_data = self._convert_decisions(northstar_json.get('decisions'))

            # Build ConvertedData
            converted = ConvertedData(
                cover=cover_data,
                performance=performance_data,
                cost=cost_data,
                readiness=readiness_data,
                decisions=decisions_data
            )

            return AdapterResult(
                data=converted,
                warnings=self.warnings,
                errors=self.errors
            )

        except Exception as e:
            self.errors.append(f"Conversion failed: {str(e)}")
            return AdapterResult(
                data=None,
                warnings=self.warnings,
                errors=self.errors
            )

    def _convert_cover(self, cover: Optional[Dict]) -> CoverData:
        """Convert cover slide data."""
        if not cover:
            self.errors.append("Missing 'cover' section")
            raise ValueError("Cover data required")

        return CoverData(
            title=cover.get('title', ''),
            subtitle=cover.get('subtitle', '')
        )

    def _convert_performance(self, perf: Optional[Dict]) -> PerformanceData:
        """Convert performance slide data."""
        if not perf:
            self.errors.append("Missing 'performance' section")
            raise ValueError("Performance data required")

        # Convert metrics: [[value, label, context], ...] → List[Tuple[str, str, str]]
        raw_metrics = perf.get('metrics', [])
        if len(raw_metrics) != 3:
            self.warnings.append(f"Expected 3 metrics, got {len(raw_metrics)}")

        metrics = [tuple(m) if len(m) == 3 else (str(m[0]), str(m[1]), str(m[2]) if len(m) > 2 else '')
                   for m in raw_metrics]

        return PerformanceData(
            kicker=perf.get('kicker', ''),
            title=perf.get('title', ''),
            subtitle=perf.get('subtitle', ''),
            metrics=metrics,
            statement=perf.get('statement', ''),
            body=perf.get('body', ''),
            conclusion=perf.get('conclusion', ''),
            footer=perf.get('footer', '')
        )

    def _convert_cost(self, cost: Optional[Dict]) -> CostData:
        """Convert cost slide data."""
        if not cost:
            self.errors.append("Missing 'cost' section")
            raise ValueError("Cost data required")

        # Chart data is optional (can be None)
        chart_data = cost.get('chart_data')
        if chart_data and not isinstance(chart_data, dict):
            self.warnings.append("chart_data should be dict or null")
            chart_data = None

        return CostData(
            kicker=cost.get('kicker', ''),
            title=cost.get('title', ''),
            subtitle=cost.get('subtitle', ''),
            chart_data=chart_data,
            chart_annotation=cost.get('chart_annotation', ''),
            table_headers=cost.get('table_headers', []),
            table_rows=cost.get('table_rows', []),
            conclusion=cost.get('conclusion', ''),
            footer=cost.get('footer', '')
        )

    def _convert_readiness(self, ready: Optional[Dict]) -> ReadinessData:
        """Convert readiness slide data."""
        if not ready:
            self.errors.append("Missing 'readiness' section")
            raise ValueError("Readiness data required")

        return ReadinessData(
            kicker=ready.get('kicker', ''),
            title=ready.get('title', ''),
            subtitle=ready.get('subtitle', ''),
            table_headers=ready.get('table_headers', []),
            phases=ready.get('phases', []),
            note=ready.get('note', ''),
            conclusion=ready.get('conclusion', ''),
            footer=ready.get('footer', '')
        )

    def _convert_decisions(self, decisions: Optional[Dict]) -> DecisionsData:
        """Convert decisions slide data."""
        if not decisions:
            self.errors.append("Missing 'decisions' section")
            raise ValueError("Decisions data required")

        # Convert decision_boxes and timeline to tuples
        decision_boxes = [tuple(box) if isinstance(box, list) else box
                          for box in decisions.get('decision_boxes', [])]
        timeline = [tuple(item) if isinstance(item, list) else item
                    for item in decisions.get('timeline', [])]

        return DecisionsData(
            kicker=decisions.get('kicker', ''),
            title=decisions.get('title', ''),
            subtitle=decisions.get('subtitle', ''),
            decision_boxes=decision_boxes,
            timeline=timeline,
            owners=decisions.get('owners', ''),
            conclusion=decisions.get('conclusion', ''),
            footer=decisions.get('footer', '')
        )


def load_northstar_json(path: Path) -> Dict[str, Any]:
    """Load and parse Northstar JSON file."""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON: {e}")
    except Exception as e:
        raise IOError(f"Cannot read {path}: {e}")


def main():
    """CLI entry point."""
    if len(sys.argv) != 3:
        print("Usage: python northstar_adapter.py <northstar-json> <output.pptx>")
        print("\nExample:")
        print("  python northstar_adapter.py artifacts/example-northstar-data.json output.pptx")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    # Validate input
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        sys.exit(1)

    if not input_path.suffix == '.json':
        print(f"Warning: Input file should be .json, got {input_path.suffix}")

    # Load Northstar JSON
    try:
        print(f"Loading {input_path}...")
        northstar_data = load_northstar_json(input_path)
    except Exception as e:
        print(f"Error loading JSON: {e}")
        sys.exit(1)

    # Convert
    print("Converting to template format...")
    adapter = NorthstarAdapter()
    result = adapter.convert(northstar_data)

    # Report warnings
    if result.warnings:
        print("\nWarnings:")
        for warn in result.warnings:
            print(f"  ⚠️  {warn}")

    # Check for errors
    if not result.is_valid:
        print("\nErrors:")
        for err in result.errors:
            print(f"  ❌ {err}")
        sys.exit(1)

    # Generate PowerPoint
    print(f"Generating PowerPoint deck...")
    try:
        # Create deck and build slides
        deck = CleanEBRDeck()
        deck.build(
            cover=result.data.cover,
            performance=result.data.performance,
            cost=result.data.cost,
            readiness=result.data.readiness,
            decisions=result.data.decisions
        )
        deck.save(str(output_path))

        print(f"\n✅ Success! Deck created: {output_path}")
        print(f"   File size: {output_path.stat().st_size:,} bytes")

        # Verification checklist
        print("\n📋 Verification Checklist:")
        print("   1. Open in PowerPoint")
        print("   2. Check font is Montserrat (or Calibri fallback)")
        print("   3. Verify chart is editable (click → Edit Data)")
        print("   4. Verify tables are editable (click cells)")
        print("   5. Check all 5 slides render correctly")

    except Exception as e:
        print(f"\n❌ Generation failed: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
