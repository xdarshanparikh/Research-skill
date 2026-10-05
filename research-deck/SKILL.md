---
name: research-deck
description: Research a topic with the latest authentic data, analyse it, design the right charts, and cite every source, then build a PowerPoint (.pptx) deck from it. Covers the whole evidence chain - newest-edition research, saving every source as a file, checking each number against the saved text, a Fact Sheet so the user picks the facts, chart choice by the data's job (sankey, waterfall, bubbles, slope, dumbbell, treemap, line), colour-blind-safe palettes, one standard source footer per slide, and a pre-flight deck check. Use this whenever someone asks for a data-backed or research-based presentation, deck, slides, PPTX, briefing, market overview, industry analysis or report deck, or asks to "research X and make slides", "turn this data into charts for a deck", or "add sources to my slides". Builds slides in the MGL (Mahanagar Gas) template look by default; another template can be swapped in.
---

# Research Deck

Make decks whose every number can be traced to a saved source, with charts that show what a bar chart can't. This skill covers **research, analysis, chart design and sources**. **Every deck uses the MGL template look** (`template/mgl-template.pptx` via `scripts/mgl_kit.py`): a green takeaway title, a grey subtitle, stat cards with icon circles, native charts in the template colours, tinted scorecards, a "Key takeaway" bar and a sources line. If the user gives a different template, use it with `scripts/pptx_helpers.py` instead. For general PPTX mechanics (editing an existing deck, layouts, speaker notes), also use the `pptx` skill if it's available.

Read these before you start:
- `references/research.md`: how to search, which sources come first, saving and verifying.
- `references/analysis.md`: data integrity, derived numbers, the Fact Sheet and building the storyline.
- `references/chart-gallery.md`: picking a chart for the data's job, chart rules, palettes, `charts.py`.
- `references/sources.md`: footers, speaker notes, `Sources.md` and the closing sources slide.
- `references/template-design.md`: the MGL slide anatomy, components, colours and composition patterns. Look at `assets/template-preview.jpg`.

## Quality bar

**Data**
1. **Use the latest data.** For every number, look for a newer edition first (the newest annual report, budget, statistical release or survey round). For anything still going on, show where it stands now, not only a target year. When a newer number changes the story, show the change ("22% → 11% since 2024").
2. **Check every number against the saved file**, not a search summary. Summaries invent and blend figures. Run `scripts/verify_numbers.py`, and drop or replace anything it can't find.
3. **Prefer primary sources.** Use regulators, ministries, statistics offices, central banks, multilaterals and company filings for numbers. Use news for events, and analysts for interpretation.
4. **Cross-check totals yourself.** Add up the parts from the primary source before you print a headline total. Compute percentages from primary values and make them add to 100.
5. **Fact-check "it stopped / it failed / it's the first" claims** for later developments.
6. **Label derived numbers as derived**, in the speaker notes, and in the footer when this changes the meaning.

**Charts**
7. **Choose the form before the colour.** Decide what the data has to show, then pick the chart from `references/chart-gallery.md`. Use richer charts (sankey, waterfall, slope, dumbbell, treemap, bubbles) when they show something a bar can't, never to look clever. Never use a dual axis.
8. **Keep one fixed palette order, validated** with `scripts/validate_palette.js`: colour-blind separation must PASS. Colour follows the entity across slides. If the story is one item, colour that item and grey the rest. Past five categories, fold the rest into "Other".
9. **Label marks directly.** Skip legend boxes when every mark is labelled. On lines, label only the start, peak and end.

**Slides**
10. **The title is the takeaway** ("Imports doubled since 2019"), not the topic ("Imports"). The subtitle says what is measured, the unit and the period.
11. **Every content slide has a source footer**, in the same place on every slide, one or two lines: `Source: Publisher, Document, Year`. Full links go in the speaker notes and `research/Sources.md`.
12. **Spell out an abbreviation the first time it appears** (in the footer: `bcm = billion cubic metres`), then use the short form on later slides.
13. **Use plain words.** Write "Demand was 3.8x the amount offered", not "oversubscribed". No invented metaphors.
14. **The user chooses** the facts, the storyline and the closing message. Offer options; don't decide for them unless they ask you to run it end to end.

## Workflow

### 1. Scope
Confirm the topic, the audience, the deck length, and the currency and units. The look is the MGL template unless the user supplies another one. Set up a project folder:
```
<project>/
├── research/            facts.json, Fact Sheet.html, claims.json
│   └── sources/         every saved PDF / snapshot / CSV + Sources.md
├── charts/              one spec .json, .svg and .png per chart
└── <Deck name>.pptx
```

### 2. Research (newest data first)
Follow `references/research.md`. Run several searches in parallel. Collect **chart-ready data**, not just headline numbers: flows for a sankey, a base plus increments for a waterfall, two dates per item for a slope or dumbbell. Save every source:
```bash
python "<skill>/scripts/save_sources.py" research/sources.json research/sources
```
Then fix every FAILED, BAD or CHECK row in `Sources.md`. Verify the numbers you plan to use:
```bash
python "<skill>/scripts/verify_numbers.py" research/sources research/claims.json
```
For a currency conversion, use a live rate (`curl -s https://open.er-api.com/v6/latest/USD`) and state the rate and date in the footer.

### 3. Fact Sheet (the user picks the facts)
Write `research/facts.json` (the schema is at the top of `scripts/build_fact_sheet.py`) and run:
```bash
python "<skill>/scripts/build_fact_sheet.py" research/facts.json
python "<skill>/scripts/sheet_server.py" 8765 "<project>"     # then open http://localhost:8765/research/Fact%20Sheet.html
```
- **Themes** are what the audience needs to understand, in the order they'd need it. Tick the 2–4 strongest facts per theme. Aim for 20–35 facts in all, so there are real choices.
- Each fact carries the number, its plain meaning, the source with a link, the year, "newer edition checked?" and the chart data it holds.
- **Leads** are angles you found but didn't research yet.
- Ticks and comments autosave to `Fact Sheet picks.txt` next to the sheet. Read that file when the user says they're done. Don't overwrite a sheet they may be filling in: pass a new name (`"Fact Sheet (round 2).html"`).
- If no browser or local server is available, show the facts as a numbered table in chat and ask for picks.

### 4. Storyline
From the picked facts, propose **2–3 genuinely different storylines**. For each, give one line per slide: the takeaway title, the chart type, and the fact IDs behind it. Also give 3 options for the closing message. Let the user pick. See `references/analysis.md`.

### 5. Charts
For each chart, write a spec in `charts/` and render it:
```bash
python "<skill>/scripts/charts.py" charts/imports.json charts/imports.png    # or .svg
```
- Size the canvas at about **140 px per inch** of the space it fills on the slide. A full-width 16:9 chart area of 12.1 × 4.9 in is about `"w": 1700, "h": 690`. Font sizes in the spec are in those px.
- Use `"theme": "mgl"` so image charts match the template's colours and fonts (`"light"` and `"dark"` are neutral). For a different palette, pass `"palette"` and validate it first:
  ```bash
  node "<skill>/scripts/validate_palette.js" "#hex,#hex,#hex" --mode light --surface "#ffffff"
  ```
- Make column, bar, stacked, line and doughnut charts **native** with `mgl_kit` (`column_chart`, `line_chart`, `doughnut_chart`), so the user can edit them in PowerPoint. Use `charts.py` images only for chart types PowerPoint lacks (sankey, waterfall with an upside bar, slope, dumbbell, treemap, bubbles).
- Look at each PNG once. Fix clipped or overlapping labels by shortening labels, changing `fs`, or resizing the canvas.

### 6. Build the deck (MGL template)
Write a small build script with `scripts/mgl_kit.py` (copy patterns from `examples/build_example_deck.py`):
```python
import sys; sys.path.insert(0, "<skill>/scripts")
from mgl_kit import Deck, C
d = Deck()                                         # template/mgl-template.pptx
d.title_slide("Title", "Subtitle", "25th Sep 2026")
s = d.content_slide("Takeaway title, one line", "What is measured, unit, period")
d.heading(s, 0.9, 1.42, 7.5, "Chart heading, unit"); d.column_chart(s, 0.81, 1.67, 7.57, 2.2, cats, {"Series": vals})
d.stat_card(s, 8.71, 1.42, 3.72, 1.0, "+53%", "what the number means", icon="chart-up", tone="green")
d.takeaway(s, "One or two sentences: the so-what.")
d.sources(s, "Sources: Publisher, Document, Year; ...")
d.notes(s, "Full citations with URLs, and derivations.")
d.closing_slide(); d.save("<Deck name>.pptx")
```
- One message per slide. Every content slide gets a takeaway title, a subtitle, at least one visual (chart, cards or table), a **Key takeaway** bar and a **sources** line. Speaker notes hold the full links and derivations.
- Vary the composition from slide to slide using the patterns in `references/template-design.md` (outlook, market snapshot, segments, flow, scenarios, plan response). Don't repeat one layout across the deck.
- Use colour with meaning: red for risk or the pain year, green for positive or recovery, amber for watch or forecast, grey for history or context. Keep data labels on.
- Pick icons from `template/icons/` (the list is in the design reference).
- End with a **Sources** slide if the deck is shared outside, then `closing_slide()`.

### 7. Check, then hand over
```bash
python "<skill>/scripts/check_deck.py" "<Deck name>.pptx" --skip 1,<closing slides>
```
This flags missing or moved footers, long footers, old data years, shapes off the slide and overlaps. Fix everything until it says "all clear". If an old year really is the latest that exists, write "(latest available)" in the footer or `data-checked: YYYY` in the notes, and tell the user. Then render and look at every slide (needs LibreOffice Impress): `soffice --headless --convert-to pdf deck.pptx` and `pdftoppm -jpeg -r 110 deck.pdf slide`. Check for text overflowing its cards, labels colliding, and titles longer than one line. Aptos may render a little wider or narrower in LibreOffice than in PowerPoint, so leave slack in tight boxes.

Hand over the `.pptx`, `research/Sources.md` and the `charts/` folder. Then list in chat anything still open: derived numbers, series that have stopped, and facts you dropped because you couldn't verify them.

## Working economically
- Run independent searches together, and use extended or deep search only for hard-to-find data.
- Look at chart PNGs once, and look again only at the ones that changed. Let `check_deck.py` do the layout checks.
- Edit decks with small targeted python-pptx scripts, never by rebuilding everything for one fix.
