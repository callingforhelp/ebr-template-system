# Clean EBR Template System

**Modern, client-ready Executive Business Review deck generator**

Based on the clean design from `Northstar_Client_EBR.pptx`:
- White background with generous whitespace
- Purple accent used sparingly (#6A2D80)
- Underlined metrics (not cards)
- Gray conclusion boxes (no labels)
- Clean tables with purple headers
- Modern typography

## Quick Start

```python
from ebr_generator_clean import CleanEBRDeck, CoverData, PerformanceData

deck = CleanEBRDeck()

prs = deck.build(
    cover=CoverData(
        title='YOUR COMPANY / BUSINESS REVIEW',
        subtitle='Q4 2026 Performance Review'
    ),
    
    performance=PerformanceData(
        kicker='YOUR COMPANY / BUSINESS REVIEW',
        title='Key achievements this quarter',
        subtitle='Q4 2026 / Oct 1 - Dec 31',
        metrics=[
            ('42%', 'Revenue growth', 'Year-over-year'),
            ('98%', 'Customer satisfaction', 'Up from 94% last quarter'),
            ('15k', 'New customers', 'Exceeds target by 23%'),
        ],
        statement='Record quarter demonstrates product-market fit',
        body='Our Q4 results reflect strong market demand...',
        conclusion='Recommend doubling down on this momentum...',
        footer='Source: Company financials, Q4 2026.'
    )
)

deck.save('my_ebr.pptx')
```

## Themes (Slide Types)

### Essential (Always Include)

1. **Cover** — Simple title page with purple accent bar
2. **Performance** — 3 metrics with underlines + statement

### Optional (Include As Needed)

3. **Cost** — Chart left, table right (for cost analysis)
4. **Readiness** — Phase table with gates (for expansion planning)
5. **Decisions** — Decision boxes + timeline (for decision frameworks)

## Design Tokens

```json
{
  "colors": {
    "purple": "#6A2D80",
    "black": "#000000",
    "gray_light": "#E8E8E8",
    "white": "#FFFFFF"
  },
  "components": {
    "metrics": "Underlines (not cards)",
    "conclusions": "Gray boxes, unlabeled",
    "tables": "Purple headers, white text"
  }
}
```

See `artifacts/design-tokens.json` for full specification.

## Schema

All theme data structures are documented in `artifacts/theme-data-schema.json`.

Example data: `artifacts/example-northstar-data.json`

## File Structure

```
ebr-template-system-clean/
├── README.md                           # This file
├── ebr_generator_clean.py              # Generator script
└── artifacts/
    ├── design-tokens.json              # Design system specification
    ├── theme-data-schema.json          # JSON schema for all themes
    └── example-northstar-data.json     # Full Northstar example
```

## When To Use Clean vs House System

| Use Case | System |
|----------|--------|
| Client-facing EBRs | Clean (this) |
| Internal reviews for technical teams | House (darker, more data-dense) |
| Executive presentations | Clean (this) |
| Data dashboards | House or custom |

## Differences from House System

| Element | House System | Clean System |
|---------|-------------|--------------|
| Background | Dark surfaces | White |
| Accent usage | Purple throughout | Purple sparingly |
| Metrics | Purple mist cards | Underlines |
| Conclusions | Purple with "SO WHAT" label | Gray, unlabeled |
| Whitespace | Dense | Generous |
| Typography | Bold throughout | Lighter weight |

## Customization

### Change Colors

Edit `DesignTokens` class in `ebr_generator_clean.py`:

```python
@dataclass
class DesignTokens:
    purple: str = 'YOUR_HEX'  # Primary accent
    gray_light: str = 'YOUR_HEX'  # Conclusion boxes
    # ...
```

### Add Custom Theme

1. Define data class:
```python
@dataclass
class YourThemeData:
    kicker: str
    title: str
    # ...
```

2. Add renderer to `CleanEBRThemes`:
```python
def your_theme(self, data: YourThemeData, page_num: int):
    slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
    b = Builder(slide, self.T)
    # ... build slide
```

3. Wire into `CleanEBRDeck.build()`:
```python
if your_theme:
    self.themes.your_theme(your_theme, self.page)
    self.page += 1
```

## Examples

See `example-northstar-data.json` for a complete deck with all 5 themes.

## License

Internal use only. Design tokens based on client-ready reference deck.
