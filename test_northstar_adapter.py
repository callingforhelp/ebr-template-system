#!/usr/bin/env python3
"""Test suite for Northstar adapter."""

import json
import unittest
from pathlib import Path
from northstar_adapter import NorthstarAdapter, load_northstar_json


class TestNorthstarAdapter(unittest.TestCase):
    """Test Northstar JSON → Template slide data conversion."""

    def setUp(self):
        """Load example Northstar data."""
        self.example_path = Path(__file__).parent / 'artifacts' / 'example-northstar-data.json'
        self.northstar_data = load_northstar_json(self.example_path)
        self.adapter = NorthstarAdapter()

    def test_example_data_exists(self):
        """Verify example data file exists and loads."""
        self.assertTrue(self.example_path.exists(), "Example data file missing")
        self.assertIsInstance(self.northstar_data, dict)
        self.assertIn('cover', self.northstar_data)
        self.assertIn('performance', self.northstar_data)
        self.assertIn('cost', self.northstar_data)
        self.assertIn('readiness', self.northstar_data)
        self.assertIn('decisions', self.northstar_data)

    def test_full_conversion(self):
        """Test complete conversion with example data."""
        result = self.adapter.convert(self.northstar_data)

        self.assertTrue(result.is_valid, f"Conversion failed: {result.errors}")
        self.assertIsNotNone(result.data)
        self.assertEqual(len(result.errors), 0)

    def test_cover_conversion(self):
        """Test cover slide conversion."""
        result = self.adapter.convert(self.northstar_data)
        cover = result.data.cover

        self.assertEqual(cover.title, "NORTHSTAR / CLIENT BUSINESS REVIEW")
        self.assertIn("M1", cover.subtitle)

    def test_performance_conversion(self):
        """Test performance slide conversion."""
        result = self.adapter.convert(self.northstar_data)
        perf = result.data.performance

        self.assertEqual(len(perf.metrics), 3)
        self.assertEqual(perf.metrics[0][0], "70%")
        self.assertEqual(perf.metrics[0][1], "Tier-1 volume handled")
        self.assertIn("97,488", perf.statement)

    def test_cost_conversion(self):
        """Test cost slide conversion."""
        result = self.adapter.convert(self.northstar_data)
        cost = result.data.cost

        self.assertIsNone(cost.chart_data, "Example has null chart_data")
        self.assertEqual(len(cost.table_headers), 3)
        self.assertGreater(len(cost.table_rows), 0)
        self.assertEqual(cost.table_rows[0][0], "Model requests")

    def test_readiness_conversion(self):
        """Test readiness slide conversion."""
        result = self.adapter.convert(self.northstar_data)
        ready = result.data.readiness

        self.assertEqual(len(ready.table_headers), 4)
        self.assertEqual(len(ready.phases), 4)
        self.assertIn("Essentials", ready.phases[0][0])

    def test_decisions_conversion(self):
        """Test decisions slide conversion."""
        result = self.adapter.convert(self.northstar_data)
        decisions = result.data.decisions

        self.assertEqual(len(decisions.decision_boxes), 3)
        self.assertEqual(len(decisions.timeline), 4)
        self.assertIn("VP CX", decisions.owners)

    def test_missing_cover(self):
        """Test error handling for missing cover."""
        bad_data = {k: v for k, v in self.northstar_data.items() if k != 'cover'}
        result = self.adapter.convert(bad_data)

        self.assertFalse(result.is_valid)
        self.assertIn("cover", str(result.errors).lower())

    def test_missing_performance(self):
        """Test error handling for missing performance."""
        bad_data = {k: v for k, v in self.northstar_data.items() if k != 'performance'}
        result = self.adapter.convert(bad_data)

        self.assertFalse(result.is_valid)
        self.assertIn("performance", str(result.errors).lower())

    def test_chart_data_optional(self):
        """Test that chart_data can be null."""
        # Example already has null chart_data
        result = self.adapter.convert(self.northstar_data)
        self.assertTrue(result.is_valid)
        self.assertIsNone(result.data.cost.chart_data)

    def test_chart_data_with_values(self):
        """Test chart_data with actual values."""
        data_with_chart = self.northstar_data.copy()
        data_with_chart['cost'] = data_with_chart['cost'].copy()
        data_with_chart['cost']['chart_data'] = {
            'categories': ['Day 1', 'Day 2', 'Day 3'],
            'values': [1500, 1600, 1550],
            'title': 'P95 Latency',
            'mean_line': 1550
        }

        result = self.adapter.convert(data_with_chart)
        self.assertTrue(result.is_valid)
        self.assertIsNotNone(result.data.cost.chart_data)
        self.assertEqual(len(result.data.cost.chart_data['categories']), 3)

    def test_metrics_count_warning(self):
        """Test warning for incorrect metrics count."""
        data_bad_metrics = self.northstar_data.copy()
        data_bad_metrics['performance'] = data_bad_metrics['performance'].copy()
        data_bad_metrics['performance']['metrics'] = [["50%", "Single metric", "context"]]

        result = self.adapter.convert(data_bad_metrics)
        self.assertTrue(result.is_valid, "Should still be valid")
        self.assertGreater(len(result.warnings), 0, "Should have warning")
        self.assertIn("3 metrics", str(result.warnings))


class TestEndToEnd(unittest.TestCase):
    """End-to-end integration tests."""

    def test_generate_deck(self):
        """Test full pipeline: JSON → Adapter → PowerPoint."""
        from ebr_generator_clean import CleanEBRDeck

        example_path = Path(__file__).parent / 'artifacts' / 'example-northstar-data.json'
        output_path = Path(__file__).parent / 'test_northstar_output.pptx'

        # Clean up any previous test output
        if output_path.exists():
            output_path.unlink()

        try:
            # Load and convert
            northstar_data = load_northstar_json(example_path)
            adapter = NorthstarAdapter()
            result = adapter.convert(northstar_data)

            self.assertTrue(result.is_valid, f"Conversion failed: {result.errors}")

            # Generate deck
            deck = CleanEBRDeck()
            deck.build(
                cover=result.data.cover,
                performance=result.data.performance,
                cost=result.data.cost,
                readiness=result.data.readiness,
                decisions=result.data.decisions
            )
            deck.save(str(output_path))

            # Verify file created
            self.assertTrue(output_path.exists(), "Output file not created")
            self.assertGreater(output_path.stat().st_size, 1000, "Output file too small")

        finally:
            # Clean up
            if output_path.exists():
                output_path.unlink()


if __name__ == '__main__':
    unittest.main()
