#!/usr/bin/env python3
"""
EBR CLI - Command-line interface for Executive Business Review deck generator.

Usage:
    ebr generate data.json output.pptx
    ebr validate data.json
    ebr init my-data.json
    ebr themes
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Optional
import jsonschema

from ebr_generator_clean import (
    CleanEBRDeck,
    CoverData,
    PerformanceData,
    CostData,
    ReadinessData,
    DecisionsData
)


def load_schema() -> dict:
    """Load the JSON schema for validation."""
    schema_path = Path(__file__).parent / "artifacts" / "theme-data-schema.json"
    if not schema_path.exists():
        print(f"⚠️  Schema not found at {schema_path}", file=sys.stderr)
        sys.exit(1)

    with open(schema_path) as f:
        return json.load(f)


def validate_json(json_path: Path, verbose: bool = False) -> bool:
    """Validate JSON data against schema."""
    schema = load_schema()

    try:
        with open(json_path) as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"❌ Invalid JSON: {e}", file=sys.stderr)
        return False
    except FileNotFoundError:
        print(f"❌ File not found: {json_path}", file=sys.stderr)
        return False

    try:
        jsonschema.validate(instance=data, schema=schema)
        print(f"✅ {json_path.name} is valid")
        if verbose:
            print(f"   Found themes: {', '.join(data.keys())}")
        return True
    except jsonschema.exceptions.ValidationError as e:
        print(f"❌ Validation failed:", file=sys.stderr)
        print(f"   {e.message}", file=sys.stderr)
        print(f"   Path: {' -> '.join(str(p) for p in e.path)}", file=sys.stderr)
        return False


def load_data_from_json(json_path: Path) -> dict:
    """Load and parse JSON data file."""
    with open(json_path) as f:
        data = json.load(f)

    # Parse into dataclass instances
    parsed = {}

    if "cover" in data:
        parsed["cover"] = CoverData(**data["cover"])

    if "performance" in data:
        parsed["performance"] = PerformanceData(**data["performance"])

    if "cost" in data:
        parsed["cost"] = CostData(**data["cost"])

    if "readiness" in data:
        parsed["readiness"] = ReadinessData(**data["readiness"])

    if "decisions" in data:
        parsed["decisions"] = DecisionsData(**data["decisions"])

    return parsed


def generate_deck(json_path: Path, output_path: Path, themes: Optional[list] = None):
    """Generate PowerPoint deck from JSON data."""
    # Validate first
    print(f"📋 Validating {json_path.name}...")
    if not validate_json(json_path):
        sys.exit(1)

    # Load data
    print(f"📂 Loading data...")
    try:
        data = load_data_from_json(json_path)
    except Exception as e:
        print(f"❌ Failed to parse data: {e}", file=sys.stderr)
        sys.exit(1)

    # Filter themes if specified
    if themes:
        available = set(data.keys())
        requested = set(themes)
        missing = requested - available
        if missing:
            print(f"⚠️  Themes not in data: {', '.join(missing)}", file=sys.stderr)
        data = {k: v for k, v in data.items() if k in requested}

    # Generate
    print(f"🎨 Generating deck with themes: {', '.join(data.keys())}...")
    deck = CleanEBRDeck()
    prs = deck.build(**data)

    # Save
    print(f"💾 Saving to {output_path}...")
    deck.save(str(output_path))

    # Report
    slide_count = len(prs.slides)
    file_size = output_path.stat().st_size / 1024  # KB
    print(f"✅ Generated {slide_count} slides ({file_size:.1f} KB)")
    print(f"   Output: {output_path.absolute()}")


def init_template(output_path: Path):
    """Create a blank JSON template for user to fill in."""
    template = {
        "cover": {
            "kicker": "YOUR COMPANY / BUSINESS REVIEW",
            "title": "Executive summary title",
            "subtitle": "Q1 2027 / Your team / Date range"
        },
        "performance": {
            "kicker": "YOUR COMPANY / BUSINESS REVIEW",
            "title": "Key performance highlights",
            "subtitle": "Q1 2027 / Jan 1 - Mar 31",
            "metrics": [
                ("70%", "First metric", "Versus baseline or target"),
                ("35%", "Second metric", "Comparison context"),
                ("+12", "Third metric", "Additional context")
            ],
            "statement": "The key message in one sentence",
            "body": "A paragraph explaining the operational context and achievements.",
            "conclusion": "The key recommendation or next action.",
            "footer": "Source: data sources. Reporting period. Important caveats."
        },
        "cost": {
            "kicker": "YOUR COMPANY / BUSINESS REVIEW",
            "title": "Cost and traffic analysis",
            "subtitle": "Opening 7 days vs final 7 days",
            "opening_metrics": [
                ("145K", "Opening metric", "Description"),
                ("65.8%", "Another metric", "Context")
            ],
            "final_metrics": [
                ("210K", "Final metric", "Description"),
                ("70.1%", "Another metric", "Context")
            ],
            "chart_note": "Explanation of the chart data",
            "table_data": [
                ("Measure", "Opening", "Final"),
                ("Requests", "145,601", "210,467"),
                ("Automation", "65.8%", "70.1%")
            ],
            "conclusion": "Key findings and recommendations",
            "footer": "Source and caveats"
        }
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w') as f:
        json.dump(template, f, indent=2)

    print(f"✅ Created template at {output_path}")
    print(f"   Edit this file, then run: ebr generate {output_path.name} output.pptx")
    print(f"\n💡 Note: Metrics use tuple format: (value, label, context)")


def list_themes():
    """List available theme types."""
    themes = [
        ("cover", "Simple title page with purple accent bar"),
        ("performance", "3 metrics with underlines + statement"),
        ("cost", "Chart left, table right (for cost analysis)"),
        ("readiness", "Phase table with gates (for expansion planning)"),
        ("decisions", "Decision boxes + timeline (for decision frameworks)")
    ]

    print("📊 Available Themes:\n")
    for name, desc in themes:
        print(f"  {name:20s} {desc}")

    print("\n💡 Include any themes in your JSON file.")
    print("   Run 'python ebr_cli.py init template.json' to see structure.")


def main():
    parser = argparse.ArgumentParser(
        description="Generate Executive Business Review decks from JSON",
        epilog="Examples:\n"
               "  python ebr_cli.py generate data.json output.pptx\n"
               "  python ebr_cli.py validate data.json\n"
               "  python ebr_cli.py init my-template.json\n"
               "  python ebr_cli.py themes",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # Generate command
    gen_parser = subparsers.add_parser("generate", help="Generate PowerPoint deck")
    gen_parser.add_argument("input", type=Path, help="Input JSON file")
    gen_parser.add_argument("output", type=Path, help="Output PPTX file")
    gen_parser.add_argument(
        "--themes",
        nargs="+",
        help="Only generate specific themes (e.g., --themes cover performance)"
    )

    # Validate command
    val_parser = subparsers.add_parser("validate", help="Validate JSON against schema")
    val_parser.add_argument("input", type=Path, help="JSON file to validate")
    val_parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")

    # Init command
    init_parser = subparsers.add_parser("init", help="Create blank JSON template")
    init_parser.add_argument("output", type=Path, help="Output JSON template file")

    # Themes command
    subparsers.add_parser("themes", help="List available theme types")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    # Dispatch
    try:
        if args.command == "generate":
            generate_deck(args.input, args.output, args.themes)
        elif args.command == "validate":
            success = validate_json(args.input, args.verbose)
            sys.exit(0 if success else 1)
        elif args.command == "init":
            init_template(args.output)
        elif args.command == "themes":
            list_themes()
    except KeyboardInterrupt:
        print("\n⚠️  Cancelled", file=sys.stderr)
        sys.exit(130)
    except Exception as e:
        print(f"❌ Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
