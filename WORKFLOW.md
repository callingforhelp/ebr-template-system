# My Workflow: From Request to Delivery

## The Complete Journey

```
┌─────────────────────────────────────────────────────────────────┐
│ 1. USER REQUEST: "take a look at the only deck pdf and make    │
│                   it better"                                    │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 2. INVESTIGATE: What exists?                                    │
├─────────────────────────────────────────────────────────────────┤
│ • Read the PDF (5 slides, Northstar EBR)                       │
│ • Content: metrics, cost analysis, expansion strategy           │
│ • Current design: functional but basic                          │
│ • My assessment: solid data, weak visual hierarchy              │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 3. LOAD DESIGN KNOWLEDGE: /dataviz skill                        │
├─────────────────────────────────────────────────────────────────┤
│ • Read: choosing-a-form.md, color-formula.md, palette.md       │
│ • Learn: validated color system, accessibility checks           │
│ • Internalize: dark surfaces ladder, accent rules, typography   │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 4. CREATE IMPROVED VERSION (HTML first)                         │
├─────────────────────────────────────────────────────────────────┤
│ Why HTML first?                                                 │
│ • Faster iteration (CSS > PPTX API)                             │
│ • Better design tools (grid, fluid type, real preview)          │
│ • Can generate PDF from it (headless Chrome)                    │
│                                                                 │
│ Design decisions:                                               │
│ • Dark theme: #05070C → #1E2636 (5-level surface ladder)       │
│ • Cyan accent: #38BDF8 (data/tech context)                      │
│ • Fluid typography: clamp() for responsive sizing               │
│ • Card-based metrics: better hierarchy than flat layout         │
│ • Proper grid: CSS Grid for section rhythm                      │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 5. USER: "we need a ppt/pdf version"                            │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 6. GENERATE PDF: Install playwright, render HTML → PDF          │
├─────────────────────────────────────────────────────────────────┤
│ • pip install playwright                                        │
│ • playwright install chromium                                   │
│ • node script to render HTML → PDF (landscape, 16:9)            │
│ • Output: Northstar_EBR_Improved.pdf (36KB, 5 slides)           │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 7. USER: "/code-review" (triggered background skill)            │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 8. CODE REVIEW RUNS (parallel, in background)                   │
├─────────────────────────────────────────────────────────────────┤
│ Found 15 bugs in wrapper scripts:                               │
│ • #1: python vs python3 (Windows compatibility)                 │
│ • #6: Unquoted paths (spaces in filenames fail)                 │
│ • #9: `open` is macOS-only (Linux fails)                        │
│ • #11: Missing `set -e` (silent failures)                       │
│ • #16: Double validation (wasteful)                             │
│ • ... (10 more issues)                                          │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 9. FIX ALL 15 BUGS                                              │
├─────────────────────────────────────────────────────────────────┤
│ generate-deck.bat:                                              │
│ • Added `py -3` fallback (robust Python detection)              │
│ • Quoted all paths: cd "%~dp0" not cd %~dp0                     │
│ • Added existence checks before `start output.pptx`             │
│                                                                 │
│ generate-deck.sh:                                               │
│ • Added `set -e` (fail fast on errors)                          │
│ • Fixed `open` → check for Linux, use `xdg-open`                │
│ • Added shebang portability: `#!/usr/bin/env bash`              │
│ • Removed double validation (generate validates internally)     │
│                                                                 │
│ Deleted:                                                        │
│ • test-template.json (duplicates `ebr_cli.py init`)             │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 10. MERGE CONFLICTS (uh oh)                                     │
├─────────────────────────────────────────────────────────────────┤
│ • Worktree had fixes, but origin/main had old buggy versions    │
│ • Used git merge strategy: keep our fixed versions              │
│ • Resolved conflicts manually                                   │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 11. CREATE TEST SUITE                                           │
├─────────────────────────────────────────────────────────────────┤
│ test-everything.sh:                                             │
│ • ✅ CLI validation works                                       │
│ • ✅ Generation creates valid PPTX                              │
│ • ✅ Wrapper scripts work                                       │
│ • ✅ All 15 code review findings addressed                      │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 12. USER: "remove all of the coauthored"                        │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 13. CLEAN GIT HISTORY                                           │
├─────────────────────────────────────────────────────────────────┤
│ Problem: Commits had "Co-Authored-By: Claude..." trailers       │
│                                                                 │
│ Solution:                                                       │
│ • git filter-branch on feature branch                           │
│ • git commit --amend on main branch                             │
│ • Stripped all Co-Authored-By lines                             │
│ • Force-pushed clean history                                    │
│                                                                 │
│ Result: Clean commit messages, no attribution trailers          │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 14. USER: "what to tell the other agent?"                       │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 15. READ EXISTING POWERPOINT GENERATOR                          │
├─────────────────────────────────────────────────────────────────┤
│ • Found: ebr_generator_clean.py (current PPTX generator)        │
│ • Uses: white background, purple accent, Calibri font           │
│ • Architecture: DesignTokens → Builder → Themes → Deck          │
│ • Clean code: dataclasses, type hints, documented               │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 16. CREATE MERGE GUIDE                                          │
├─────────────────────────────────────────────────────────────────┤
│ MERGE_GUIDE.md:                                                 │
│ • Token mapping table (HTML dark → PPTX light)                  │
│ • Three merge options (both themes / dark only / refactor)      │
│ • Recommended: Option 1 (add DarkDesignTokens class)            │
│ • Step-by-step instructions                                     │
│ • Code examples for dark theme implementation                   │
│ • CLI flag: --theme light|dark                                  │
│ • Validation checklist                                          │
│ • What transfers well, what doesn't                             │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 17. USER: "can you walk me through your workflow"               │
└─────────────────────────────────────────────────────────────────┘
                              ↓
                         (You are here!)
```

## Key Workflow Principles I Follow

### 1. **Investigate Before Building**
- Read existing code/files first
- Understand current state and constraints
- Identify what's missing vs what needs improvement

### 2. **Load Relevant Knowledge**
- Use skills (/dataviz) when available
- Read references before designing
- Apply validated patterns, don't improvise

### 3. **Start with Best Tool for the Job**
- HTML for visual design (faster iteration)
- Python for data processing
- Git worktrees for isolation

### 4. **Respond to User Feedback**
- User said "need PDF" → generated PDF
- User said "/code-review" → reviewed and fixed bugs
- User said "remove coauthored" → cleaned git history
- Adapt to new requirements without restarting

### 5. **Verify Everything**
- Code review found 15 bugs → fixed all 15
- Created test suite to prove fixes work
- Checked git history was actually clean

### 6. **Document Handoffs**
- MERGE_GUIDE.md has token mappings, code examples
- Explained three options, recommended one
- Gave validation checklist

### 7. **Use Git Properly**
- Worktrees for isolation
- Feature branches, not direct-to-main
- Clean history (no duplicate trailers)
- Test before merge

## Tools & Techniques Used

### Development
- **Python**: PPTX generation, PDF rendering
- **HTML/CSS**: Visual design prototype
- **Node.js**: PDF rendering (Playwright)
- **Git worktrees**: Isolated development

### Quality
- **Code review skill**: Found 15 bugs automatically
- **Test suite**: Validated all fixes work
- **Manual verification**: Checked each script runs

### Documentation
- **MERGE_GUIDE.md**: Technical handoff
- **TEST_REPORT.md**: Validation evidence
- **This workflow doc**: Process transparency

## What Made This Efficient

1. **Ponytail mode active**: Lazy = efficient
   - Reused existing PPTX generator (didn't rebuild)
   - HTML first (faster design iteration)
   - Fixed bugs in place (didn't refactor unnecessarily)

2. **Parallel work**: Code review ran while I continued
   - Didn't block on review results
   - Fixed issues when review completed

3. **Right tool per task**:
   - CSS for visual design (not PPTX API)
   - Playwright for PDF (not manual export)
   - Git filter-branch for history rewrite (not manual edits)

4. **Incremental delivery**:
   - HTML version → delivered
   - PDF version → delivered
   - Bug fixes → committed
   - Clean history → force-pushed
   - Merge guide → documented

## Common Patterns

### Pattern: Read → Understand → Act
```
User request → Read existing → Identify gap → Build solution → Deliver
```

### Pattern: Use Skills for Specialized Knowledge
```
Need design system → /dataviz → Read references → Apply patterns
```

### Pattern: Fix Root Cause, Not Symptom
```
Code review found bugs → Fix all 15 (not just critical ones)
Test suite proves fixes work (not just "looks OK")
```

### Pattern: Document for Handoff
```
Built HTML dark theme → Other agent owns PPTX generator → 
MERGE_GUIDE.md explains how to integrate
```

## Questions I Ask Myself

1. **What already exists?** (Read before building)
2. **What's the fastest path?** (HTML before PPTX)
3. **Is this verified?** (Test suite, not assumptions)
4. **Can someone else continue this?** (Merge guide)
5. **Is the history clean?** (Git hygiene matters)

---

**That's my workflow!** The key is: investigate, adapt, verify, document.
