# research-deck · a Claude skill for research-backed PowerPoint decks

The research, analysis, chart-design and sources parts of the green-carousel skill, made brand-neutral and adapted for **.pptx** output. Brand identity, illustrations and the carousel and social-caption steps are not included.

## What it does

| Step | Output |
|---|---|
| 1. Research, newest data first | Every source saved to `research/sources/` (PDFs, page snapshots, CSV) with a `Sources.md` index |
| 2. Verify | `verify_numbers.py` confirms each number appears in the saved source text |
| 3. Fact Sheet | `Fact Sheet.html`: the user ticks facts and leads, and picks autosave to disk |
| 4. Storyline | 2–3 storylines: a takeaway title, a chart type and fact IDs per slide |
| 5. Charts | `charts.py`: sankey, waterfall, bubbles, scatter, slope, dumbbell, treemap and line, as SVG or 2x PNG, with colour-blind-safe palettes |
| 6. Build | python-pptx helpers: a takeaway title, the chart, one standard source footer and speaker notes with full links |
| 7. Check | `check_deck.py`: missing or moved footers, old data years, shapes off the slide, overlaps |

## Install
Copy the `research-deck` folder into your skills folder (`~/.claude/skills/research-deck/`, or `C:\Users\<you>\.claude\skills\research-deck\`). Then ask for a research deck, e.g. "Research the global LNG market and make a 10-slide deck".

## Requirements
```bash
pip install -r requirements.txt
python -m playwright install chromium   # for source snapshots and chart PNGs
```
Node.js is needed for `validate_palette.js` (palette checks only).

## Folder contents
```
research-deck/
├── SKILL.md                      workflow + quality bar (read by Claude)
├── references/
│   ├── research.md               search order, newest editions, saving + verifying sources
│   ├── analysis.md               data integrity, Fact Sheet, storyline options
│   ├── chart-gallery.md          chart by the data's job, chart rules, palettes, charts.py usage
│   └── sources.md                footer standard, speaker notes, Sources slide
├── scripts/
│   ├── save_sources.py           saves every source + writes Sources.md
│   ├── verify_numbers.py         finds each claimed number in the saved files
│   ├── build_fact_sheet.py       Fact Sheet page from facts.json
│   ├── sheet_server.py           local server so Fact Sheet picks autosave
│   ├── charts.py                 SVG/PNG chart generator (no extra packages for SVG)
│   ├── chart_png.py              SVG → 2x transparent PNG (Playwright)
│   ├── validate_palette.js       colour-blind + contrast checks for a palette
│   ├── pptx_helpers.py           chart slide, source footer and notes helpers (python-pptx)
│   └── check_deck.py             pre-flight check of the finished .pptx
└── assets/
    ├── chart-gallery.jpg         every chart type rendered
    ├── chart-gallery-specs.json  working specs to copy
    └── fact-sheet-template.html
```

## Using your own template or colours
- Open your template with `Presentation("template.pptx")` and use its layouts. `pptx_helpers.LAYOUT` holds the title, chart and footer positions.
- Pass your colours to charts with `"palette": [...]` (last = grey/other) after checking them with `validate_palette.js`.
