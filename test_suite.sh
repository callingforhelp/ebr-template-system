#!/usr/bin/env bash
# Comprehensive test suite for EBR CLI and wrapper scripts

set -euo pipefail

cd "$(dirname "$0")"

PASSED=0
FAILED=0
TEST_DIR="test_output_$$"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Test runner
run_test() {
    local test_name="$1"
    local test_cmd="$2"

    echo -n "Testing: $test_name... "

    if eval "$test_cmd" > /dev/null 2>&1; then
        echo -e "${GREEN}✓ PASS${NC}"
        ((PASSED++))
        return 0
    else
        echo -e "${RED}✗ FAIL${NC}"
        ((FAILED++))
        return 1
    fi
}

# Setup
mkdir -p "$TEST_DIR"
cd "$TEST_DIR"

echo "========================================"
echo "  EBR CLI Test Suite"
echo "========================================"
echo ""

# Test 1: CLI help
run_test "CLI help command" "python3 ../ebr_cli.py --help"

# Test 2: Themes list
run_test "List themes" "python3 ../ebr_cli.py themes"

# Test 3: Init creates template
run_test "Init creates template" "python3 ../ebr_cli.py init test-init.json && [ -f test-init.json ]"

# Test 4: Validate example data
run_test "Validate example data" "python3 ../ebr_cli.py validate ../artifacts/example-northstar-data.json"

# Test 5: Generate from example
run_test "Generate from example" "python3 ../ebr_cli.py generate ../artifacts/example-northstar-data.json test-output.pptx && [ -f test-output.pptx ]"

# Test 6: Generated file is PowerPoint
run_test "Output is valid PPTX" "file test-output.pptx | grep -q 'Microsoft PowerPoint'"

# Test 7: Generated file has reasonable size
run_test "Output file size check" "[ \$(stat -f%z test-output.pptx 2>/dev/null || stat -c%s test-output.pptx) -gt 10000 ]"

# Test 8: Selective theme generation
run_test "Selective themes (cover only)" "python3 ../ebr_cli.py generate ../artifacts/example-northstar-data.json selective-test.pptx --themes cover && [ -f selective-test.pptx ]"

# Test 9: Invalid JSON fails validation
echo '{"invalid": "json"' > invalid.json
if python3 ../ebr_cli.py validate invalid.json 2>/dev/null; then
    echo -e "Testing: Invalid JSON detection... ${RED}✗ FAIL${NC}"
    ((FAILED++))
else
    echo -e "Testing: Invalid JSON detection... ${GREEN}✓ PASS${NC}"
    ((PASSED++))
fi

# Test 10: Missing required field fails validation
echo '{"title": "Test"}' > missing-fields.json
if python3 ../ebr_cli.py validate missing-fields.json 2>/dev/null; then
    echo -e "Testing: Missing fields detection... ${RED}✗ FAIL${NC}"
    ((FAILED++))
else
    echo -e "Testing: Missing fields detection... ${GREEN}✓ PASS${NC}"
    ((PASSED++))
fi

# Test 11: Dependencies are installed
run_test "python-pptx installed" "python3 -c 'import pptx'"
run_test "jsonschema installed" "python3 -c 'import jsonschema'"

# Test 12: Core generator module loads
run_test "Core generator imports" "python3 -c 'from ebr_generator_clean import create_ebr_deck'"

# Test 13: Schema file exists
run_test "Schema file present" "[ -f ../artifacts/theme-data-schema.json ]"

# Test 14: Shell script is executable
run_test "generate-deck.sh executable" "[ -x ../generate-deck.sh ]"

# Test 15: Requirements file exists
run_test "requirements.txt present" "[ -f ../requirements.txt ]"

# Test 16: README exists
run_test "README.md present" "[ -f ../README.md ]"

# Test 17: INSTALL guide exists
run_test "INSTALL.md present" "[ -f ../INSTALL.md ]"

# Cleanup
cd ..
rm -rf "$TEST_DIR"

echo ""
echo "========================================"
echo "  Test Results"
echo "========================================"
echo -e "  ${GREEN}Passed:${NC} $PASSED"
echo -e "  ${RED}Failed:${NC} $FAILED"
echo "========================================"
echo ""

if [ $FAILED -eq 0 ]; then
    echo -e "${GREEN}All tests passed!${NC}"
    exit 0
else
    echo -e "${RED}Some tests failed.${NC}"
    exit 1
fi
