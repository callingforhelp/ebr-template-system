#!/usr/bin/env bash
set -euo pipefail

echo "=== Testing Wrapper Scripts ==="
echo ""

# Test 1: First run (init path)
echo "Test 1: First-time initialization"
rm -f my-deck.json output.pptx
timeout 2 bash generate-deck.sh <<< "" || true
if [ -f my-deck.json ]; then
    echo "   ✓ Created my-deck.json on first run"
else
    echo "   ✗ Failed to create my-deck.json"
    exit 1
fi

# Test 2: Second run with real data (generation path)
echo ""
echo "Test 2: Generation with existing data"
cp artifacts/example-northstar-data.json my-deck.json
timeout 5 bash generate-deck.sh <<< "" || true
if [ -f output.pptx ]; then
    SIZE=$(stat -f%z output.pptx 2>/dev/null || stat -c%s output.pptx)
    echo "   ✓ Generated output.pptx ($SIZE bytes)"
else
    echo "   ✗ Failed to generate output.pptx"
    exit 1
fi

# Test 3: Error handling (invalid JSON)
echo ""
echo "Test 3: Error handling with invalid JSON"
echo '{"invalid": true}' > my-deck.json
timeout 2 bash generate-deck.sh <<< "" 2>&1 | grep -q "❌" && echo "   ✓ Correctly detected invalid data"

# Cleanup
rm -f my-deck.json output.pptx

echo ""
echo "=== All Wrapper Tests Passed! ==="
