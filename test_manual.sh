#!/usr/bin/env bash
set -euo pipefail

echo "=== Manual Test: End-to-End Flow ==="
echo ""

# Clean start
rm -f my-test-deck.json my-test-output.pptx

# Step 1: Init
echo "1. Creating template..."
python3 ebr_cli.py init my-test-deck.json
[ -f my-test-deck.json ] && echo "   ✓ Template created"

# Step 2: Validate
echo "2. Validating example data..."
python3 ebr_cli.py validate artifacts/example-northstar-data.json
echo "   ✓ Validation passed"

# Step 3: Generate
echo "3. Generating PowerPoint..."
python3 ebr_cli.py generate artifacts/example-northstar-data.json my-test-output.pptx
[ -f my-test-output.pptx ] && echo "   ✓ PowerPoint created"

# Step 4: Check file size
SIZE=$(stat -f%z my-test-output.pptx 2>/dev/null || stat -c%s my-test-output.pptx)
echo "4. Output file: my-test-output.pptx ($SIZE bytes)"
[ $SIZE -gt 10000 ] && echo "   ✓ File size reasonable"

# Step 5: Test themes list
echo "5. Available themes:"
python3 ebr_cli.py themes | head -5

echo ""
echo "=== All Tests Passed! ==="
echo ""
echo "Output file: my-test-output.pptx"
echo "You can open it to verify the content."

# Cleanup test files
rm -f my-test-deck.json
