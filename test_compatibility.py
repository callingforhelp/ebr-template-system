#!/usr/bin/env python3
"""
Compatibility Test Suite for Template System

Tests all 4 fixes:
1. Montserrat font support with Calibri fallback
2. Optional cover slide (skip_cover parameter)
3. Native editable PowerPoint charts
4. Native PowerPoint table objects
"""

import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent))

from ebr_generator_clean import (
    CleanEBRDeck, DesignTokens,
    CoverData, PerformanceData, CostData, ReadinessData, DecisionsData,
    get_font_family
)
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE


def test_font_configuration():
    """Test Fix #1: Montserrat font with fallback."""
    print("\n" + "="*70)
    print("TEST 1: Font Configuration")
    print("="*70)

    # Test font retrieval
    font, is_preferred = get_font_family('Montserrat', 'Calibri')
    print(f"✓ Font selected: {font}")
    print(f"✓ Is preferred (Montserrat): {is_preferred}")

    # Test tokens use correct font
    tokens = DesignTokens()
    assert tokens.font_family in ['Montserrat', 'Calibri'], \
        f"Font must be Montserrat or Calibri, got {tokens.font_family}"
    print(f"✓ DesignTokens font_family: {tokens.font_family}")

    # Test fallback warning (if Calibri is used)
    if tokens.font_family == 'Calibri':
        print("⚠️  Note: Falling back to Calibri (Montserrat not available)")
        print("   This is expected behavior if Montserrat isn't installed")

    print("✅ PASS: Font configuration working correctly\n")
    return True


def test_skip_cover_parameter():
    """Test Fix #2: Optional cover slide."""
    print("="*70)
    print("TEST 2: Optional Cover Slide")
    print("="*70)

    # Test with cover (default)
    deck_with_cover = CleanEBRDeck()
    deck_with_cover.build(
        cover=CoverData(
            title='Test Deck With Cover',
            subtitle='This should have a cover slide'
        ),
        performance=PerformanceData(
            kicker='TEST',
            title='Performance Slide',
            subtitle='Test subtitle',
            metrics=[('100%', 'Test Metric', 'Context')],
            statement='Test statement',
            body='Test body',
            conclusion='Test conclusion',
            footer='Test footer'
        ),
        skip_cover=False
    )

    slides_with_cover = len(deck_with_cover.prs.slides)
    print(f"✓ Deck with cover: {slides_with_cover} slides")
    assert slides_with_cover == 2, f"Expected 2 slides (cover + performance), got {slides_with_cover}"

    # Test without cover (skip_cover=True)
    deck_without_cover = CleanEBRDeck()
    deck_without_cover.build(
        cover=CoverData(
            title='Test Deck Without Cover',
            subtitle='This cover should be skipped'
        ),
        performance=PerformanceData(
            kicker='TEST',
            title='Performance Slide',
            subtitle='Test subtitle',
            metrics=[('100%', 'Test Metric', 'Context')],
            statement='Test statement',
            body='Test body',
            conclusion='Test conclusion',
            footer='Test footer'
        ),
        skip_cover=True  # Skip cover
    )

    slides_without_cover = len(deck_without_cover.prs.slides)
    print(f"✓ Deck with skip_cover=True: {slides_without_cover} slides")
    assert slides_without_cover == 1, f"Expected 1 slide (performance only), got {slides_without_cover}"

    # Test exact 5 content slides (Northstar requirement)
    deck_five_slides = CleanEBRDeck()
    deck_five_slides.build(
        cover=CoverData(title='Cover', subtitle='Should be skipped'),
        performance=PerformanceData(
            kicker='TEST', title='Slide 1', subtitle='',
            metrics=[('1', 'M1', 'C1')], statement='S', body='B', conclusion='C', footer='F'
        ),
        cost=CostData(
            kicker='TEST', title='Slide 2', subtitle='',
            chart_data={'categories': ['A'], 'values': [100]},
            chart_annotation='Ann',
            table_headers=['H1'], table_rows=[['R1']],
            conclusion='C', footer='F'
        ),
        readiness=ReadinessData(
            kicker='TEST', title='Slide 3', subtitle='',
            table_headers=['H1'], phases=[['P1']], note='N',
            conclusion='C', footer='F'
        ),
        decisions=DecisionsData(
            kicker='TEST', title='Slide 4', subtitle='',
            decision_boxes=[('D1', 'B1')], timeline=[('T1', 'W1')],
            owners='O', conclusion='C', footer='F'
        ),
        skip_cover=True
    )

    five_slides_count = len(deck_five_slides.prs.slides)
    print(f"✓ Northstar-style deck (skip_cover=True): {five_slides_count} slides")
    assert five_slides_count == 4, f"Expected 4 content slides, got {five_slides_count}"

    print("✅ PASS: Cover slide skip working correctly\n")
    return True


def test_native_charts():
    """Test Fix #3: Native editable PowerPoint charts."""
    print("="*70)
    print("TEST 3: Native Editable Charts")
    print("="*70)

    # Create deck with chart data
    deck = CleanEBRDeck()
    deck.build(
        cost=CostData(
            kicker='TEST',
            title='Chart Test',
            subtitle='Testing native PowerPoint charts',
            chart_data={
                'categories': ['Day 1', 'Day 2', 'Day 3', 'Day 4', 'Day 5'],
                'values': [1850, 1920, 1780, 1650, 1590],
                'title': 'Test Latency Chart',
                'mean_line': 1831
            },
            chart_annotation='Test annotation',
            table_headers=['Header'],
            table_rows=[['Data']],
            conclusion='Test conclusion',
            footer='Test footer'
        ),
        skip_cover=True
    )

    # Find chart shapes
    slide = deck.prs.slides[0]
    chart_shapes = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.CHART]

    print(f"✓ Found {len(chart_shapes)} chart shape(s)")
    assert len(chart_shapes) == 1, f"Expected 1 chart, found {len(chart_shapes)}"

    # Verify chart properties
    chart_shape = chart_shapes[0]
    chart = chart_shape.chart

    assert chart is not None, "Chart object should not be None"
    print(f"✓ Chart type: {chart.chart_type}")
    assert chart.chart_type == XL_CHART_TYPE.LINE, \
        f"Expected LINE chart, got {chart.chart_type}"

    # Check series
    assert len(chart.series) > 0, "Chart should have at least one series"
    print(f"✓ Chart has {len(chart.series)} series")

    # Check if chart has data points
    main_series = chart.series[0]
    assert len(main_series.values) == 5, \
        f"Expected 5 data points, got {len(main_series.values)}"
    print(f"✓ Series has {len(main_series.values)} data points")

    # Check title
    if chart.has_title:
        title_text = chart.chart_title.text_frame.text
        print(f"✓ Chart title: '{title_text}'")

    # Test chart with mean line
    assert len(chart.series) == 2, \
        f"Expected 2 series (data + mean line), got {len(chart.series)}"
    print("✓ Mean reference line series present")

    print("✅ PASS: Native charts working correctly\n")
    return True


def test_native_tables():
    """Test Fix #4: Native PowerPoint table objects."""
    print("="*70)
    print("TEST 4: Native PowerPoint Tables")
    print("="*70)

    # Test cost slide table
    print("\n--- Testing Cost Slide Table ---")
    deck_cost = CleanEBRDeck()
    deck_cost.build(
        cost=CostData(
            kicker='TEST',
            title='Table Test',
            subtitle='Testing native PowerPoint tables',
            chart_data=None,  # No chart
            chart_annotation='N/A',
            table_headers=['Measure', 'Value 1', 'Value 2'],
            table_rows=[
                ['Metric A', '100', '200'],
                ['Metric B', '150', '250'],
                ['Metric C', '175', '225'],
            ],
            conclusion='Test conclusion',
            footer='Test footer'
        ),
        skip_cover=True
    )

    slide_cost = deck_cost.prs.slides[0]
    table_shapes_cost = [s for s in slide_cost.shapes if s.has_table]

    print(f"✓ Found {len(table_shapes_cost)} table(s) on cost slide")
    assert len(table_shapes_cost) == 1, \
        f"Expected 1 table on cost slide, found {len(table_shapes_cost)}"

    # Verify table structure
    table_cost = table_shapes_cost[0].table
    print(f"✓ Table dimensions: {len(table_cost.rows)} rows × {len(table_cost.columns)} columns")
    assert len(table_cost.rows) == 4, f"Expected 4 rows (1 header + 3 data), got {len(table_cost.rows)}"
    assert len(table_cost.columns) == 3, f"Expected 3 columns, got {len(table_cost.columns)}"

    # Check header cell
    header_cell = table_cost.cell(0, 0)
    print(f"✓ Header cell text: '{header_cell.text}'")
    assert header_cell.text == 'Measure', f"Expected 'Measure', got '{header_cell.text}'"

    # Check data cell
    data_cell = table_cost.cell(1, 0)
    print(f"✓ First data cell text: '{data_cell.text}'")
    assert data_cell.text == 'Metric A', f"Expected 'Metric A', got '{data_cell.text}'"

    # Test readiness slide table
    print("\n--- Testing Readiness Slide Table ---")
    deck_readiness = CleanEBRDeck()
    deck_readiness.build(
        readiness=ReadinessData(
            kicker='TEST',
            title='Readiness Table Test',
            subtitle='Testing readiness table',
            table_headers=['Phase', 'Details', 'Outcome', 'Gate'],
            phases=[
                ['Phase 1', 'Description 1', 'Outcome 1', 'Gate 1'],
                ['Phase 2', 'Description 2', 'Outcome 2', 'Gate 2'],
            ],
            note='Test note',
            conclusion='Test conclusion',
            footer='Test footer'
        ),
        skip_cover=True
    )

    slide_readiness = deck_readiness.prs.slides[0]
    table_shapes_readiness = [s for s in slide_readiness.shapes if s.has_table]

    print(f"✓ Found {len(table_shapes_readiness)} table(s) on readiness slide")
    assert len(table_shapes_readiness) == 1, \
        f"Expected 1 table on readiness slide, found {len(table_shapes_readiness)}"

    table_readiness = table_shapes_readiness[0].table
    print(f"✓ Table dimensions: {len(table_readiness.rows)} rows × {len(table_readiness.columns)} columns")
    assert len(table_readiness.rows) == 3, f"Expected 3 rows, got {len(table_readiness.rows)}"
    assert len(table_readiness.columns) == 4, f"Expected 4 columns, got {len(table_readiness.columns)}"

    # Test decisions slide table (timeline)
    print("\n--- Testing Decisions Slide Timeline Table ---")
    deck_decisions = CleanEBRDeck()
    deck_decisions.build(
        decisions=DecisionsData(
            kicker='TEST',
            title='Decisions Table Test',
            subtitle='Testing timeline table',
            decision_boxes=[('Box 1', 'Description 1')],
            timeline=[
                ('Day 1', 'Event 1'),
                ('Day 7', 'Event 2'),
                ('Day 14', 'Event 3'),
            ],
            owners='Test owners',
            conclusion='Test conclusion',
            footer='Test footer'
        ),
        skip_cover=True
    )

    slide_decisions = deck_decisions.prs.slides[0]
    table_shapes_decisions = [s for s in slide_decisions.shapes if s.has_table]

    print(f"✓ Found {len(table_shapes_decisions)} table(s) on decisions slide")
    assert len(table_shapes_decisions) == 1, \
        f"Expected 1 table on decisions slide, found {len(table_shapes_decisions)}"

    table_decisions = table_shapes_decisions[0].table
    print(f"✓ Table dimensions: {len(table_decisions.rows)} rows × {len(table_decisions.columns)} columns")
    assert len(table_decisions.rows) == 4, f"Expected 4 rows, got {len(table_decisions.rows)}"
    assert len(table_decisions.columns) == 2, f"Expected 2 columns, got {len(table_decisions.columns)}"

    print("✅ PASS: Native tables working correctly\n")
    return True


def test_integration_full_deck(tmp_path=None):
    """Integration test: Generate full deck with all fixes."""
    print("="*70)
    print("INTEGRATION TEST: Full Deck Generation")
    print("="*70)

    deck = CleanEBRDeck()

    prs = deck.build(
        cover=CoverData(
            title='INTEGRATION TEST DECK',
            subtitle='Testing all 4 compatibility fixes'
        ),
        performance=PerformanceData(
            kicker='TEST / INTEGRATION',
            title='All fixes integrated',
            subtitle='This deck tests font, cover skip, charts, and tables',
            metrics=[
                ('✓', 'Montserrat font', 'With Calibri fallback'),
                ('✓', 'Optional cover', 'Can be skipped'),
                ('✓', 'Native charts', 'Editable in PowerPoint'),
            ],
            statement='Native tables also implemented',
            body='This test verifies that all 4 compatibility fixes work together: (1) Montserrat font with fallback, (2) Optional cover slide, (3) Native editable charts, (4) Native PowerPoint tables.',
            conclusion='All compatibility fixes integrated successfully.',
            footer='Integration test generated on 2026-10-24'
        ),
        cost=CostData(
            kicker='TEST / INTEGRATION',
            title='Native chart and table demonstration',
            subtitle='Both should be fully editable in PowerPoint',
            chart_data={
                'categories': ['W1', 'W2', 'W3', 'W4'],
                'values': [100, 95, 90, 85],
                'title': 'Test Metric Over Time',
                'mean_line': 92.5
            },
            chart_annotation='Downward trend with mean reference line',
            table_headers=['Feature', 'Status'],
            table_rows=[
                ['Font', '✓ Working'],
                ['Cover skip', '✓ Working'],
                ['Charts', '✓ Working'],
                ['Tables', '✓ Working'],
            ],
            conclusion='All features validated.',
            footer='Test execution successful'
        ),
        skip_cover=False  # Include cover for this test
    )

    total_slides = len(prs.slides)
    print(f"✓ Generated deck with {total_slides} slides")
    assert total_slides == 3, f"Expected 3 slides (cover + 2 content), got {total_slides}"

    # Verify chart exists on cost slide (slide index 2: cover=0, performance=1, cost=2)
    slide_cost = prs.slides[2]
    charts = [s for s in slide_cost.shapes if hasattr(s, 'has_chart') and s.has_chart]
    print(f"✓ Slide 3 (cost) has {len(charts)} chart(s)")
    assert len(charts) == 1, "Cost slide should have 1 chart"

    # Verify table exists
    tables = [s for s in slide_cost.shapes if s.has_table]
    print(f"✓ Slide 3 (cost) has {len(tables)} table(s)")
    assert len(tables) == 1, "Cost slide should have 1 table"

    # Save test deck - use tmp_path if provided (pytest), otherwise temp location
    if tmp_path:
        output_path = tmp_path / 'test_compatibility_deck.pptx'
    else:
        import tempfile
        temp_dir = tempfile.gettempdir()
        output_path = f"{temp_dir}/test_compatibility_deck.pptx"

    deck.save(str(output_path))
    print(f"✓ Saved test deck: {output_path}")
    print("  → Open this file in PowerPoint to verify:")
    print("     - Font is Montserrat (or Calibri if not installed)")
    print("     - Chart is editable (click on it)")
    print("     - Table is editable (click in cells)")

    print("✅ PASS: Integration test successful\n")
    return True


def run_all_tests():
    """Run all compatibility tests."""
    print("\n" + "🎯 "*35)
    print("TEMPLATE SYSTEM COMPATIBILITY TEST SUITE")
    print("Testing 4 fixes for Northstar integration")
    print("🎯 "*35 + "\n")

    tests = [
        ("Font Configuration (Fix #1)", test_font_configuration),
        ("Optional Cover Slide (Fix #2)", test_skip_cover_parameter),
        ("Native Charts (Fix #3)", test_native_charts),
        ("Native Tables (Fix #4)", test_native_tables),
        ("Full Deck Integration", test_integration_full_deck),
    ]

    results = []
    for test_name, test_func in tests:
        try:
            passed = test_func()
            results.append((test_name, "PASS" if passed else "FAIL", None))
        except Exception as e:
            results.append((test_name, "FAIL", str(e)))
            print(f"❌ FAIL: {test_name}")
            print(f"   Error: {e}\n")

    # Summary
    print("="*70)
    print("TEST SUMMARY")
    print("="*70)

    passed_count = sum(1 for _, status, _ in results if status == "PASS")
    total_count = len(results)

    for test_name, status, error in results:
        symbol = "✅" if status == "PASS" else "❌"
        print(f"{symbol} {test_name}: {status}")
        if error:
            print(f"   └─ {error}")

    print(f"\n{passed_count}/{total_count} tests passed")

    if passed_count == total_count:
        print("\n🎉 ALL TESTS PASSED! Template system is compatible with Northstar.")
        return 0
    else:
        print("\n⚠️  SOME TESTS FAILED. Review errors above.")
        return 1


if __name__ == '__main__':
    sys.exit(run_all_tests())
