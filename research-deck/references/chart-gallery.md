# Chart gallery: pick the chart by the data's job

Use a richer chart only when it shows something a bar can't (a flow, a build-up, a before and after). A typical analytical deck can carry **two or more** of these when the data supports them. Rendered examples of every chart: `assets/chart-gallery.jpg` (specs in `assets/chart-gallery-specs.json`).

## 1. The data's job → the chart

| The slide has to show... | Use | Example | Generator |
|---|---|---|---|
| Where something **comes from and where it goes** (money, energy, materials, users) | **Sankey** | suppliers → country → sectors | `sankey` |
| How a total **builds up** from today to a target, step by step (a forecast, a budget, a gap) | **Waterfall** | 65 → +15 → +4 → +19 → 103 | `waterfall` |
| A few **sizes side by side** where the gap is the story (one giant, many small) | **Packed bubbles** | new capacity: 112 vs the rest | `bubble` |
| Two measures per item + size (countries, companies) | **Bubble scatter** (label at most 4; fold the rest into "Other") | imports vs import share | `scatter` |
| **Ranks or values changing** between two dates for 3–6 items | **Slope chart** (highlight 1–2, grey the rest) | supplier shares 2025 → 2026 | `slope` |
| **Before → after** per item (prices, costs, emissions) | **Dumbbell** | price before vs now | `dumbbell` |
| **Part of a whole** with 5–8 parts | **Treemap** (better than a pie past 4 slices) | use by sector | `treemap` |
| A **level over time** against benchmarks | **Line + reference lines** (one axis only) | spot price vs average and floor | `line` |
| One headline number | **Stat tile**: a big number + one line, not a one-bar chart | "+40%" | native text |
| Amounts over 3–8 periods | Column chart | | native PPTX chart |
| Ranking of many items | Horizontal bar, sorted | | native PPTX chart |
| Shares, at most 4 parts | Donut or a 100% stacked bar | | native PPTX chart |
| Events in order | Timeline (dots + dates + one line each, at most 6) | | shapes |
| Ranges or uncertainty | Range or dot chart | | shapes |
| A share against a limit | Meter: one track + a fill | "47% of capacity used" | shapes |

**Not a chart:** when more than 7 classes all matter, use a small table. When two measures have different units, use two charts, **never a dual axis**.

## 2. Rules

- **Form before colour.** Decide the job, then the chart, then the colours.
- **Fixed colour order, never cycled.** Past the palette, fold the tail into "Other" in grey. Colour follows the entity: if Qatar is blue on slide 3, it stays blue on slide 4.
- **Emphasis beats rainbow.** If the story is one item, colour it and grey the rest.
- **Direct labels, used selectively.** Label each mark once (name + % or value). Skip the legend box when everything is labelled. On a line, label only the start, peak and end.
- **Text in ink colours** (near-black, grey or white), not the series colour. The one exception is the % accent in sankey labels.
- **Thin, calm marks:** rounded bar ends, 2–4 px gaps between fills, solid hairline axes (dashes only for a benchmark or a projection), no gridlines, no 3D, no shadows.
- **Percentages add to 100.** Totals match the source. If sector data and supply totals differ, scale one or say so.
- **Derived numbers are labelled as derived** in the notes, and in the footer when this changes the meaning.
- **Axis and units:** put the unit in the subtitle or the axis label once, not on every value. Start bar axes at zero.

## 3. Palettes

Defaults in `scripts/charts.py`. Both pass `validate_palette.js` (CVD separation, normal-vision floor, lightness band):

| Theme (surface) | Fixed order (slot 1 → last = grey / other) |
|---|---|
| `light` (#ffffff) | blue #2a78d6 · orange #eb6834 · green #1baf7a · amber #eda100 · pink #e87ba4 · grey #b8b8b8 |
| `dark` (#1f2328) | blue #3b82d6 · orange #d95926 · green #199e70 · amber #c28400 · pink #cf5f8c · grey #6b6f76 |

To use the template's own colours, pass `"palette": [...]` in each spec (the last entry = grey/other) and validate them on the slide background first:
```bash
node "<skill>/scripts/validate_palette.js" "#hex,#hex,#hex,#hex" --mode light --surface "#ffffff"
```
CVD separation and the normal-vision floor must PASS. A contrast WARN means every mark needs a visible label, which these charts always have. Saturated brand backgrounds fail the lightness band by design. What matters there is CVD separation.

## 4. Using `scripts/charts.py`

Write a spec (one object, or a list for several charts), then render:
```bash
python "<skill>/scripts/charts.py" charts/flows.json charts/flows.png   # PNG at 2x, transparent
python "<skill>/scripts/charts.py" charts/flows.json charts/flows.svg   # editable SVG
python "<skill>/scripts/chart_png.py" charts/flows.svg                   # SVG → PNG later
```
Every spec has `type`, `theme` (light/dark), `w`, `h` (canvas px, about 140 px per inch of the slide area). Optional: `palette`, `font` (default Arial; use the template's body font), `cw` (that font's average character width ÷ size, about 0.5–0.6), `bg` (paint a background; the default is transparent). Each function's docstring lists its type fields. Copy a working spec from `assets/chart-gallery-specs.json`.

Tips:
- **Sankey:** 3 columns (sources → hub → uses) read best. Use at most 8 nodes per side, and merge nodes under ~2.5% into "Others". Give the hub a label (it's rotated inside the bar). Use `"sub"` for a second label line. Colour links per link (`"color"`) or with `"link_color": "source"|"target"`.
- **Waterfall:** make the first and last steps `"kind": "total"`. Use `\n` in labels for 2–3 line captions. Colour an optional "upside" bar differently (`"color"`).
- **Bubble:** use at most 7 bubbles. The biggest should be at least 3x the smallest, or a bar chart is clearer.
- **Line:** give x values as days or months from the start, `xticks` for labels, `marks` for the 2–3 points worth naming, and `refs` for benchmarks.
- **Font sizes** (`fs`, `lfs`) are canvas px. At 140 px per inch, 28 px ≈ 14 pt on the slide. Keep chart text at 12 pt or more.
- Look at the PNG once. Text widths are estimated, so shorten labels or change `fs` if anything is clipped.
