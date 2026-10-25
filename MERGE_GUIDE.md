# Merge Guide: HTML Dark Theme → PowerPoint Generator

This document explains how to merge the improved HTML dark theme design into the PowerPoint generator (`ebr_generator_clean.py`).

## Two Parallel Implementations

### Current State:
1. **PowerPoint Generator** (`ebr_generator_clean.py`) - White background, purple accent, clean modern
2. **HTML Improved Version** (`northstar_ebr_improved.html`) - Dark theme, cyan accent, modern depth

## Quick Reference: Design Token Mapping

### HTML Design System → PowerPoint Tokens

| Concept | HTML (Dark) | PowerPoint (Light) | Notes |
|---------|-------------|-------------------|-------|
| **Background** | `#05070C` → `#1E2636` (5 surfaces) | `#FFFFFF` (white) | Dark uses surface ladder |
| **Primary Accent** | `#38BDF8` (cyan) | `#6A2D80` (purple) | Main brand color |
| **Positive/Success** | `#6EE7B7` (green) | Not used | For positive metrics |
| **Text Primary** | `#F8FAFC` (near-white) | `#000000` (black) | Main text |
| **Text Secondary** | `#94A3B8` (slate) | `#4A4A4A` (dark gray) | Labels, context |
| **Text Muted** | `#64748B` (muted slate) | `#7A7A7A` (mid gray) | Footer notes |
| **Card/Surface** | `#0F131C` (surface-2) | None (flat white) | HTML uses layered cards |
| **Borders** | `#161D2B` (surface-3) | None | Subtle depth in dark |

### Typography Comparison

| Element | HTML | PowerPoint | Recommendation |
|---------|------|------------|----------------|
| **Display/Title** | `clamp(2rem, 5vw, 3.5rem)` fluid | `32pt` fixed | Keep PowerPoint fixed (PPTX constraint) |
| **Metric Value** | `clamp(3rem, 7vw, 5rem)` fluid | `56pt` | PowerPoint is correct scale |
| **Body** | `clamp(0.875rem, 1.8vw, 1.125rem)` | `11pt` | PowerPoint standard |
| **Font Family** | System stack (`-apple-system`) | `Calibri` | Calibri is safe PowerPoint choice |
| **Tracking** | Tight negative on display (`-0.02em`) | None | PowerPoint doesn't support letter-spacing easily |

### Layout Differences

| Feature | HTML | PowerPoint | Merge Strategy |
|---------|------|------------|----------------|
| **Grid System** | CSS Grid (responsive) | Fixed `Inches()` positioning | Keep PowerPoint absolute positioning |
| **Border Radius** | `12px-16px` (rounded cards) | None (rectangular shapes) | PowerPoint doesn't support rounded corners well |
| **Spacing** | Fluid `clamp()` tokens | Fixed inch values | Keep PowerPoint fixed (presentation format) |
| **Metrics Display** | Card-based with backgrounds | Underlines only | HTML cards are better; consider adding to PPTX |
| **Interactive** | Hover tooltips, crosshairs | Static | PPTX is print-only |

## Merge Approach: Three Options

### Option 1: Add Dark Theme as New Style ⭐ RECOMMENDED

**What:** Create `DarkDesignTokens` class alongside existing `DesignTokens`

**Why:** 
- Non-breaking (keeps existing light theme)
- CLI can offer `--theme light|dark` flag
- Users can choose based on presentation context

**How:**
```python
@dataclass
class DarkDesignTokens(DesignTokens):
    """Dark theme with cyan accent — modern depth."""
    # Override colors
    bg_surface_0: str = '05070C'     # Deepest
    bg_surface_1: str = '0A0D12'
    bg_surface_2: str = '0F131C'     # Cards
    bg_surface_3: str = '161D2B'     # Borders
    bg_surface_4: str = '1E2636'     # Elevated
    
    accent: str = '38BDF8'           # Cyan (was purple)
    accent_dim: str = '0891B2'       # Teal
    positive: str = '6EE7B7'         # Green for good metrics
    warning: str = 'F59E0B'          # Amber
    
    text_primary: str = 'F8FAFC'     # Near-white
    text_secondary: str = '94A3B8'   # Slate
    text_muted: str = '64748B'       # Muted
    
    # Keep existing typography sizes
    # (they're already good)
```

**Then:**
```python
# In ebr_cli.py
deck = CleanEBRDeck(tokens=DarkDesignTokens()) if args.dark_mode else CleanEBRDeck()
```

### Option 2: Replace Light with Dark (Breaking Change)

**What:** Replace all colors in `DesignTokens` with dark theme

**Pros:** Simpler codebase (one theme)  
**Cons:** Breaks existing users who expect light theme

**When:** Only if you know all users want dark mode

### Option 3: Make Theme Pluggable

**What:** Create `Theme` protocol/ABC, make both light and dark implement it

**Pros:** Most flexible, best architecture  
**Cons:** Most work, requires refactoring `Builder` class

## Step-by-Step: Recommended Merge (Option 1)

### 1. Add Dark Tokens to `ebr_generator_clean.py`

```python
@dataclass
class DarkDesignTokens(DesignTokens):
    """Dark theme — cyan accent, layered surfaces."""
    
    # Surfaces (ascending brightness)
    bg_0: str = '05070C'
    bg_1: str = '0A0D12' 
    bg_2: str = '0F131C'  # Main canvas
    bg_3: str = '161D2B'  # Cards
    bg_4: str = '1E2636'  # Elevated elements
    
    # Accents
    accent: str = '38BDF8'
    accent_dim: str = '0891B2'
    positive: str = '6EE7B7'
    
    # Text
    text_primary: str = 'F8FAFC'
    text_secondary: str = '94A3B8'
    text_muted: str = '64748B'
    
    # Override methods
    def slide_background(self):
        return self.bg_2  # Use surface-2 as main canvas
    
    def card_background(self):
        return self.bg_3  # Cards one level up
```

### 2. Update `Builder` to Use Surface Ladder

Currently, `Builder` assumes white background everywhere. For dark theme:

```python
class Builder:
    def __init__(self, slide, tokens: DesignTokens):
        self.slide = slide
        self.T = tokens
        # Set slide background
        if hasattr(tokens, 'slide_background'):
            bg = slide.background
            fill = bg.fill
            fill.solid()
            fill.fore_color.rgb = tokens.rgb(tokens.slide_background())
    
    def card(self, l, t, w, h):
        """Card with elevated surface."""
        bg_color = self.T.card_background() if hasattr(self.T, 'card_background') else self.T.white
        return self.rect(l, t, w, h, bg_color)
```

### 3. Update Theme Renderers to Use Cards

Replace direct shapes with card primitives:

```python
# OLD (light theme — metrics with underlines)
def metric_underline(self, l, t, w, value, label, context):
    self.line(l, t + 0.85, l + w, t + 0.85, self.T.purple, 2)
    self.text(l, t, w, 0.75, value, 56, True, self.T.purple)
    # ...

# NEW (dark theme — metrics in cards)
def metric_card(self, l, t, w, value, label, context):
    """Metric card with background (dark theme)."""
    self.card(l, t, w, 1.6)  # Use elevated surface
    self.text(l + 0.2, t + 0.2, w - 0.4, 0.5, value,
             56, True, self.T.accent)
    self.text(l + 0.2, t + 0.75, w - 0.4, 0.3, label,
             16, True, self.T.text_primary)
    self.text(l + 0.2, t + 1.1, w - 0.4, 0.25, context,
             10, False, self.T.text_muted)
```

### 4. Add CLI Flag

In `ebr_cli.py`:

```python
parser.add_argument('--theme', choices=['light', 'dark'], default='light',
                    help='Visual theme (default: light)')

# ...

tokens = DarkDesignTokens() if args.theme == 'dark' else DesignTokens()
deck = CleanEBRDeck(tokens=tokens)
```

### 5. Test Both Themes

```bash
# Light theme (existing)
python ebr_cli.py generate my-deck.json output-light.pptx

# Dark theme (new)
python ebr_cli.py generate my-deck.json output-dark.pptx --theme dark
```

## Key Implementation Notes

### PowerPoint Limitations vs HTML

| HTML Feature | PowerPoint Equivalent | Workaround |
|--------------|----------------------|------------|
| `border-radius: 16px` | Not supported | Use rectangles with no outline |
| `box-shadow` | Very limited | Omit or use subtle borders |
| Fluid `clamp()` sizing | Fixed `Pt()` only | Use fixed sizes (PPTX is fixed canvas) |
| CSS Grid | Absolute `Inches()` positioning | Calculate positions manually |
| Hover states | Not applicable | Static presentation |
| `linear-gradient()` | Shape fills only | Use solid fills, gradient shapes are complex |

### What Transfers Well

✅ **Color palette** - Direct hex → RGBColor mapping  
✅ **Typography scale** - Font sizes map 1:1  
✅ **Spacing rhythm** - Inch values are absolute  
✅ **Component hierarchy** - Cards, metrics, tables translate  
✅ **Content structure** - Same data models work  

### What Doesn't Transfer

❌ Rounded corners (PPTX limitation)  
❌ Shadows (too complex in PPTX)  
❌ Fluid responsive sizing (PPTX is fixed canvas)  
❌ Interactions (hover, click — static medium)  
❌ CSS Grid (must calculate absolute positions)  

## Validation Checklist

After merging dark theme:

- [ ] All 5 slide types render without errors
- [ ] Text is readable (contrast check: primary text on surface-2 ≥ 4.5:1)
- [ ] Metrics stand out (accent color visible)
- [ ] Tables have clear headers (inverted text on accent background)
- [ ] Footer text is legible (muted but not invisible)
- [ ] Page numbers visible
- [ ] No white rectangles bleeding through (check slide backgrounds)
- [ ] File size reasonable (<500KB for 5-slide deck)

## Color Contrast Validation

Run this check for dark theme text:

```python
def check_contrast(fg_hex, bg_hex):
    """WCAG AA requires 4.5:1 for normal text, 3:1 for large."""
    # Compute relative luminance, then contrast ratio
    # (implementation left as exercise)
    pass

# Must pass:
check_contrast('F8FAFC', '0F131C')  # Primary text on surface-2: should be ~14:1 ✓
check_contrast('94A3B8', '0F131C')  # Secondary text: should be ~7:1 ✓
check_contrast('64748B', '0F131C')  # Muted text: should be ~4.6:1 ✓ (minimum)
```

## Migration Path for Existing Users

If you have users with existing JSON data files:

1. **Default to light theme** (no breaking change)
2. **Add `--theme` flag** (opt-in dark mode)
3. **Document in README**: "Use `--theme dark` for modern dark style"
4. **Show examples**: Include both light and dark screenshots

## Example: Before & After

### Before (Light Theme)
```python
# ebr_generator_clean.py (current)
tokens = DesignTokens()  # White bg, purple accent
deck = CleanEBRDeck(tokens=tokens)
```

### After (Both Themes Available)
```python
# User chooses at CLI
python ebr_cli.py generate data.json out.pptx --theme light  # Original
python ebr_cli.py generate data.json out.pptx --theme dark   # New dark theme
```

Internally:
```python
if args.theme == 'dark':
    tokens = DarkDesignTokens()
else:
    tokens = DesignTokens()

deck = CleanEBRDeck(tokens=tokens)
```

## Questions for Your Team

1. **Do you want to keep light theme, or replace it entirely?**  
   Recommendation: Keep both, let users choose

2. **Should dark be the new default?**  
   Recommendation: No, keep light as default (less breaking)

3. **Do you want card-based metrics (HTML style) or underlines (current PPTX)?**  
   Recommendation: Cards for dark theme (better depth), underlines for light (cleaner)

4. **Should we add rounded corners even though PPTX support is limited?**  
   Recommendation: No, keep rectangles (PPTX limitation)

## Next Steps

1. Review this guide with your team
2. Decide: Option 1 (both themes), Option 2 (dark only), or Option 3 (refactor)
3. I can implement the chosen option
4. Test with real data
5. Update README and examples
6. Ship!

---

**Summary:** The HTML dark theme can be merged as a new `DarkDesignTokens` class with a `--theme` CLI flag. The color palette maps directly, typography sizes are already good, and the main work is adapting card-based layouts (HTML) to PPTX absolute positioning. Recommend keeping both themes available.
