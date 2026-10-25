#!/usr/bin/env bash
# Northstar EBR Pipeline - Generate PowerPoint from Dashboard Data
#
# Usage:
#   ./generate_northstar_deck.sh [northstar-json] [output-pptx]
#
# Defaults:
#   Input:  ebr-slides.json (current directory)
#   Output: Northstar_EBR_<date>.pptx

set -euo pipefail

# Configuration
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ADAPTER="$SCRIPT_DIR/northstar_adapter.py"
DEFAULT_INPUT="ebr-slides.json"
DEFAULT_OUTPUT="Northstar_EBR_$(date +%Y%m%d).pptx"

# Parse arguments
INPUT="${1:-$DEFAULT_INPUT}"
OUTPUT="${2:-$DEFAULT_OUTPUT}"

# Validate input exists
if [ ! -f "$INPUT" ]; then
    echo "❌ Error: Input file not found: $INPUT"
    echo ""
    echo "Usage: $0 [northstar-json] [output-pptx]"
    echo ""
    echo "Example:"
    echo "  $0 ebr-slides.json output.pptx"
    exit 1
fi

# Validate adapter exists
if [ ! -f "$ADAPTER" ]; then
    echo "❌ Error: Adapter not found: $ADAPTER"
    echo ""
    echo "Make sure you're running this from the template system directory."
    exit 1
fi

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: python3 not found"
    exit 1
fi

# Generate
echo "🎨 Northstar EBR Pipeline"
echo "=========================="
echo ""
echo "Input:  $INPUT"
echo "Output: $OUTPUT"
echo ""

python3 "$ADAPTER" "$INPUT" "$OUTPUT"

# Verify output
if [ -f "$OUTPUT" ]; then
    SIZE=$(stat -f%z "$OUTPUT" 2>/dev/null || stat -c%s "$OUTPUT" 2>/dev/null || echo "unknown")
    echo ""
    echo "✅ Pipeline complete!"
    echo "   File: $OUTPUT"
    echo "   Size: $SIZE bytes"
    echo ""
    echo "Next steps:"
    echo "  1. Open in PowerPoint"
    echo "  2. Verify content accuracy"
    echo "  3. Check charts are editable (click → Edit Data)"
    echo "  4. Check tables are editable (click cells)"
else
    echo ""
    echo "❌ Pipeline failed - output file not created"
    exit 1
fi
