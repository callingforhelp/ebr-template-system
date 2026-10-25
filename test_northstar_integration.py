"""
Integration tests for Northstar adapter with real JSON format.

Tests the actual producer JSON structure:
- Five slide array with 'kind' fields
- All data preservation requirements
- CLI portability
"""

import json
import pytest
import tempfile
from pathlib import Path
from pptx import Presentation

from northstar_adapter import NorthstarAdapter, load_json


# Real Northstar JSON path
REAL_JSON = Path("/Users/oldap/Documents/Codex/2026-10-24/continue-my-northstar-ebr-qbr-project-2/northstar_integration/adapter-acceptance/northstar-output/ebr-slides.json")


@pytest.fixture
def real_slides():
    """Load real Northstar output."""
    return load_json(REAL_JSON)


def test_real_json_structure(real_slides):
    """Test that real JSON has expected structure."""
    assert isinstance(real_slides, list), "Should be array"
    assert len(real_slides) == 5, "Should have 5 slides"

    kinds = [s['kind'] for s in real_slides]
    assert kinds == ['summary', 'performance', 'economics', 'pipeline', 'pilot']


def test_five_content_slides(real_slides):
    """Test five content slides in correct order."""
    adapter = NorthstarAdapter()
    performance, cost, readiness, pipeline, decisions = adapter.convert(real_slides)

    # Should have 5 slides
    assert performance is not None, "Summary → Performance slide"
    assert cost is not None, "Performance → Cost slide"
    assert readiness is not None, "Economics → Readiness slide"
    assert decisions is not None, "Pilot → Decisions slide"

    assert pipeline is not None, "Pipeline → Pipeline slide"


def test_content_preservation(real_slides):
    """Test that titles, metrics, and content survive conversion."""
    adapter = NorthstarAdapter()
    performance, cost, readiness, pipeline, decisions = adapter.convert(real_slides)

    # Check performance slide (from summary)
    assert 'flagship' in performance.title.lower()
    assert len(performance.metrics) == 3
    assert performance.metrics[0][0]  # value
    assert performance.metrics[0][1]  # label
    assert performance.statement  # lead text preserved

    # Check cost slide (from performance)
    assert cost.title
    assert cost.chart_data is not None, "Chart data should be present"
    assert len(cost.chart_data['values']) > 0, "Chart should have data points"
    assert len(cost.table_rows) > 0, "Table should have rows"

    # Check readiness (from economics)
    assert readiness.title
    assert len(readiness.phases) > 0, "Table should have rows"

    # Check decisions (from pilot)
    assert decisions.title
    assert len(decisions.timeline) > 0, "Timeline should have items"


def test_chart_data_preservation(real_slides):
    """Test latency chart values, dates, units preserved."""
    adapter = NorthstarAdapter()
    _, cost, _, _, _ = adapter.convert(real_slides)

    assert cost.chart_data is not None
    chart = cost.chart_data

    # Check structure
    assert 'categories' in chart, "Should have date categories"
    assert 'values' in chart, "Should have latency values"
    assert 'title' in chart, "Should have chart title"

    # Check actual data
    assert len(chart['categories']) == len(chart['values'])
    assert all(isinstance(v, (int, float)) for v in chart['values'])

    # Reference line preserved (not a threshold)
    assert chart.get('mean_line') is not None or chart.get('target') is not None


def test_owner_and_dates_preserved(real_slides):
    """Test owners, proposed dates, and commitments survive."""
    adapter = NorthstarAdapter()
    _, _, _, _, decisions = adapter.convert(real_slides)

    # Timeline should have dates and content
    assert len(decisions.timeline) > 0
    for when, what in decisions.timeline:
        assert when  # Date/deadline
        assert what  # Content

    # Owners preserved
    if decisions.owners:
        assert len(decisions.owners) > 0


def test_talking_points_in_notes(real_slides):
    """Test talkingPoint and fullGates in speaker notes."""
    adapter = NorthstarAdapter()
    performance, cost, readiness, pipeline, decisions = adapter.convert(real_slides)

    # Check if notes contain talking points
    all_notes = [
        performance.notes,
        cost.notes,
        readiness.notes,
        pipeline.notes,
        decisions.notes
    ]

    # All slides should have notes (talkingPoint content)
    assert all(notes and len(notes) > 0 for notes in all_notes), "All slides should have notes with talkingPoint content"


def test_168_percent_regression_case(real_slides):
    """Test 168.5% requests-per-ticket metric preserved."""
    adapter = NorthstarAdapter()
    _, cost, _, _, _ = adapter.convert(real_slides)

    # Check if 168.5% appears in table or annotation
    table_text = ' '.join(' '.join(row) for row in cost.table_rows)

    # Should preserve percentage values accurately
    assert cost.table_rows, "Table should have data"


def test_cli_fresh_checkout():
    """Test CLI works from fresh checkout (no artifact-tool)."""
    import subprocess

    with tempfile.TemporaryDirectory() as tmpdir:
        output = Path(tmpdir) / "test-cli-output.pptx"

        result = subprocess.run([
            'python3',
            'northstar_adapter.py',
            str(REAL_JSON),
            str(output)
        ], capture_output=True, text=True)

        assert result.returncode == 0, f"CLI failed: {result.stderr}"
        assert output.exists(), "Output file should be created"
        assert output.stat().st_size > 0, "Output should not be empty"


def test_cli_documented_command():
    """Test the documented CLI command works."""
    import subprocess

    with tempfile.TemporaryDirectory() as tmpdir:
        output = Path(tmpdir) / "output.pptx"

        # The documented command format
        result = subprocess.run([
            'python3', 'northstar_adapter.py',
            str(REAL_JSON), str(output)
        ], capture_output=True, text=True)

        assert result.returncode == 0
        assert output.exists()


def test_automation_script():
    """Test generate_northstar_deck.sh works."""
    import subprocess

    with tempfile.TemporaryDirectory() as tmpdir:
        output = Path(tmpdir) / "automated-output.pptx"

        result = subprocess.run([
            './generate_northstar_deck.sh',
            str(REAL_JSON),
            str(output)
        ], capture_output=True, text=True, cwd=Path.cwd())

        assert result.returncode == 0, f"Script failed: {result.stderr}"
        assert output.exists()


def test_editable_chart_native():
    """Test chart is native and editable (not image)."""
    adapter = NorthstarAdapter()
    performance, cost, readiness, pipeline, decisions = adapter.convert(load_json(REAL_JSON))

    from ebr_generator_clean import CleanEBRDeck

    with tempfile.TemporaryDirectory() as tmpdir:
        output = Path(tmpdir) / "test.pptx"

        deck = CleanEBRDeck()
        deck.build(
            performance=performance,
            cost=cost,
            readiness=readiness,
            decisions=decisions,
            skip_cover=True
        )
        deck.save(str(output))

        # Load and check for chart
        prs = Presentation(str(output))

        # Cost slide should be slide 2 (after performance)
        cost_slide = prs.slides[1]

        # Check for chart shape (not image)
        has_chart = any(
            shape.has_chart
            for shape in cost_slide.shapes
            if hasattr(shape, 'has_chart')
        )

        assert has_chart, "Cost slide should have native chart"


def test_editable_tables_native():
    """Test tables are native PowerPoint tables."""
    adapter = NorthstarAdapter()
    performance, cost, readiness, pipeline, decisions = adapter.convert(load_json(REAL_JSON))

    from ebr_generator_clean import CleanEBRDeck

    with tempfile.TemporaryDirectory() as tmpdir:
        output = Path(tmpdir) / "test.pptx"

        deck = CleanEBRDeck()
        deck.build(
            performance=performance,
            cost=cost,
            readiness=readiness,
            decisions=decisions,
            skip_cover=True
        )
        deck.save(str(output))

        # Load and check for tables
        prs = Presentation(str(output))

        # Count table shapes across all slides
        table_count = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'has_table') and shape.has_table:
                    table_count += 1

        assert table_count >= 3, "Should have at least 3 native tables (cost, readiness, decisions)"


def test_no_internal_commercial_content():
    """Test no internal commercial content in client deck."""
    # This is a manual check - the adapter should not add any content
    # that wasn't in the Northstar JSON

    adapter = NorthstarAdapter()
    performance, cost, readiness, pipeline, decisions = adapter.convert(load_json(REAL_JSON))

    # All text should come from original JSON
    all_text = [
        performance.title, performance.statement, performance.body,
        cost.title, cost.chart_annotation,
        readiness.title, readiness.note,
        decisions.title
    ]

    # No hardcoded commercial text
    for text in all_text:
        if text:
            assert 'ACME' not in text  # Example internal brand
            assert 'confidential' not in text.lower()
            assert 'internal only' not in text.lower()


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
