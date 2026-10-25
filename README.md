# Northstar EBR Adapter

Converts Northstar dashboard JSON output to PowerPoint presentations using the EBR template system.

## Quick Start

```bash
# Install dependencies (already satisfied by template system)
python3 northstar_adapter.py <northstar-json> <output.pptx>
```

## Example

```bash
# Convert example data to PowerPoint
python3 northstar_adapter.py artifacts/example-northstar-data.json output.pptx
```

Output:
```
Loading artifacts/example-northstar-data.json...
Converting to template format...
Generating PowerPoint deck...

✅ Success! Deck created: output.pptx
   File size: 38,714 bytes

📋 Verification Checklist:
   1. Open in PowerPoint
   2. Check font is Montserrat (or Calibri fallback)
   3. Verify chart is editable (click → Edit Data)
   4. Verify tables are editable (click cells)
   5. Check all 5 slides render correctly
```

## What It Does

1. **Loads** Northstar JSON (the dashboard's `ebr-slides.json` output)
2. **Converts** to template system's data classes (CoverData, PerformanceData, etc.)
3. **Validates** required fields and reports warnings/errors
4. **Generates** PowerPoint with native charts and tables

## Input Format

The adapter expects Northstar's `ebr-slides.json` structure:

```json
{
  "cover": {
    "title": "NORTHSTAR / CLIENT BUSINESS REVIEW",
    "subtitle": "M1 / Northstar leadership review / 2026-08-01 to 2026-08-31"
  },
  "performance": {
    "kicker": "NORTHSTAR / CLIENT BUSINESS REVIEW",
    "title": "The flagship delivers operational value...",
    "subtitle": "M1 / ...",
    "metrics": [
      ["70%", "Tier-1 volume handled", "Reported in the case brief"],
      ["35%", "Lower average handle time", "Reported versus baseline"],
      ["+12 pts", "CSAT improvement", "Reported versus baseline"]
    ],
    "statement": "97,488 automated tickets...",
    "body": "The CSV records...",
    "conclusion": "First assess flagship benefits...",
    "footer": "Source: ..."
  },
  "cost": { ... },
  "readiness": { ... },
  "decisions": { ... }
}
```

## Output Format

Native PowerPoint `.pptx` with:
- **Editable charts** (native Excel data, not images)
- **Editable tables** (native PowerPoint tables)
- **Montserrat font** (falls back to Calibri if unavailable)
- **5 slides**: Cover, Performance, Cost, Readiness, Decisions

## Testing

```bash
# Run full test suite
python3 -m pytest test_northstar_adapter.py -v

# 13 tests covering:
# - Data loading and validation
# - Each slide type conversion
# - Error handling
# - End-to-end PowerPoint generation
```

All 13 tests passing ✅

## Architecture

```
Northstar Dashboard
       ↓
  ebr-slides.json (their format)
       ↓
  northstar_adapter.py (this adapter)
       ↓
  CoverData, PerformanceData, etc. (template format)
       ↓
  ebr_generator_clean.py (template renderer)
       ↓
  output.pptx (native PowerPoint with editable objects)
```

## Error Handling

The adapter validates:
- ✅ All required sections present (cover, performance, cost, readiness, decisions)
- ✅ Metrics count (warns if not exactly 3)
- ✅ Chart data structure (optional, can be null)
- ✅ Table headers and rows present

Returns:
- `result.is_valid` → True if conversion succeeded
- `result.warnings` → List of non-fatal issues
- `result.errors` → List of fatal problems

## Integration

The adapter is a **stateless converter**: it doesn't modify Northstar's code or data files.

**Northstar side** (their repo):
- Generate `ebr-slides.json` as usual
- No changes needed

**Template side** (this repo):
- Run adapter on their JSON
- Get PowerPoint output

**To integrate**:
1. Northstar generates `ebr-slides.json`
2. Run: `python3 northstar_adapter.py ebr-slides.json final-deck.pptx`
3. Deliver `final-deck.pptx` to stakeholders

## Files

- `northstar_adapter.py` - Main adapter + CLI
- `test_northstar_adapter.py` - Test suite (13 tests)
- `artifacts/example-northstar-data.json` - Example input
- `README.md` - This file

## Compatibility

- **Python**: 3.8+
- **Dependencies**: Same as template system (`python-pptx`)
- **Northstar format**: Current `ebr-slides.json` structure
- **Template**: Compatible with PR #3 (merged) compatibility fixes

## Next Steps

1. ✅ Adapter built and tested (13/13 passing)
2. ✅ Template compatibility fixes merged (PR #3)
3. ⏳ **Integration test with real Northstar output**
4. ⏳ **End-to-end validation with stakeholders**

## Support

Questions? Issues?
- Template bugs → [ebr-template-system issues](https://github.com/callingforhelp/ebr-template-system/issues)
- Adapter bugs → This directory
- Northstar data bugs → Their repo (R1 & R2 fixed)
