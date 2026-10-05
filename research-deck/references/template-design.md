# Template design: the MGL look

Every deck uses `template/mgl-template.pptx`: the master, layouts, logo, green band, page-number dot and "MAHANAGAR GAS LIMITED" footer, with no content slides. Build slides with `scripts/mgl_kit.py`, which encodes the rules below. `assets/template-preview.jpg` shows the result, and `examples/build_example_deck.py` builds every component.

## Slide anatomy (13.33 × 7.5 in)
| Zone | Position (in) | Style |
|---|---|---|
| Title | layout placeholder, top-left | Aptos 24 pt, brand green `#51AD30`, one line. The title is the takeaway |
| Subtitle | x 0.9, y 0.95, w 11.53 | 12 pt `#404040`: what is measured, unit, period, or which sources |
| Body | x 0.9 → 12.43, y 1.33 → 5.9 | charts, cards and tables on a 2- or 3-column grid with 0.15–0.3 in gaps |
| Block heading | just above its chart | 11 pt bold black: "What, unit (source)" |
| Key takeaway bar | x 0.9, y 6.0, w 11.53, h 0.6 | fill `#EAF5E6`, "Key takeaway" bold blue `#0067B2` + 1–2 sentences, 11 pt black |
| Sources | x 0.9, y 6.68 | 7.5 pt `#6E6E6E`, one or two lines, shape named `Source` |
| Page number | layout (green dot, bottom right) | added automatically by the kit |

Every content slide ends with a takeaway bar and a sources line. The title slide (`title_slide`) and the closing "Thank You…" slide (`closing_slide`) carry neither.

## Components (`mgl_kit.Deck`)
| Component | Use it for | Look |
|---|---|---|
| `stat_card` | one headline number with context | grey `#F5F5F5` card 3.7 × 1.0, icon in a coloured circle, number 19 pt bold (red `#B9382A` for risk, dark green `#2F7A17` for positive), 9 pt grey label |
| `stat_row` | 3–4 KPIs across the top of a slide | green-tint cards 0.84 tall, black numbers |
| `info_card` | a region, segment, scenario or option | grey panel: bold title, optional blue highlight line, grey body, optional icon and status `tag` |
| `callout` | a caveat or warning | tinted box with a coloured lead-in: `watch` amber `#FFF3D6`, `risk` red `#F6E1DE`, `good` green, `note` grey with a blue lead |
| `tag` | a status pill ("High risk", "Tailwind", "Positive") | outlined, 7.5 pt bold, tinted |
| `table` | scorecards, comparisons, scenario matrices | green-tint header, grey rows, white 2.5 pt gaps; tint cells with `(text, "good"/"bad"/"watch")` |
| `column_chart` | amounts, balances, highlighted years | native chart, data labels on, no gridlines, gap 45%, grey axis; `point_colors` for one-series highlights; `stacked=True`; `horizontal=True`; `line={...}` overlays a line on bars |
| `line_chart` | paths over time vs benchmarks | 2.25 pt lines, round markers, end labels only on series that don't collide |
| `doughnut_chart` | at most 4–5 shares | white value labels, legend on the right |
| `picture` | sankey, waterfall, slope, dumbbell, treemap, bubbles from `charts.py` | render with `"theme": "mgl"` so the colours and fonts match |
| `bullets` | 2–4 implications beside a chart | 10 pt, bold lead-in words |
| `icon` | a white pictogram in a coloured circle | 28 icons in `template/icons/` |

Icons: waves, warehouse, factory, chart-up, flag, handshake, buildings, rupee, people, train, bus, metro, ship, flame, fuel-pump, leaf, pie, bolt, tag, percent, scales, network, shield, bulb, sprout, exchange, layers, home.

## Colour
| Role | Hex |
|---|---|
| Brand green: titles, icon circles, positive bars | `#51AD30` |
| Dark green: positive numbers | `#2F7A17` |
| Blue: primary series, "Key takeaway", highlight text | `#0067B2` |
| Amber: second series, "watch" | `#FCB000` (text `#9A6A00`) |
| Red: risk, negative, peak values | `#B9382A` |
| Text / secondary / source / axis | `#000000` / `#404040` / `#6E6E6E` / `#888888` |
| Tints: card / green / amber / red | `#F5F5F5` / `#EAF5E6` / `#FFF3D6` / `#F6E1DE` |

Chart series order: blue → amber → green → red → light blue `#7FB2D9` → grey. This passes the colour-blind check with direct labels (amber and green sit in the warning band). So **always keep data labels on**, and use the colours with meaning: red for the pain year or risk, green for a recovery or a positive, grey for history or context, amber for a forecast or partial year.

## Composition patterns from the template
- **Outlook:** a chart top-left and a smaller chart bottom-left, with 4 stacked stat cards in a right column.
- **Market snapshot:** a stat row on top, two charts side by side, then 3 info cards and a `watch` callout.
- **Segments:** a stacked chart on the left, a tinted scorecard table and bullets on the right, and a `risk` callout.
- **Flow:** a `charts.py` sankey on the left, and 4 icon cards with tags on the right.
- **Policy or scenarios:** a full-width table (or 3 info cards) with traffic-light cells and a legend of tints.
- **Plan response:** chevron-style phase cards across the top, a chart on the left and 6 icon cards on the right.

## Type and density
- Font: Aptos, as in the template. For readers on older Office without Aptos, set `mgl_kit.FONT = "Calibri"`.
- The template is dense by design: 8.5–9 pt card text and 7.5 pt sources. Don't go smaller. If text doesn't fit, cut words or split the slide.
- No accent lines under titles, and no extra colour bars: the green band and dotted rule come from the layout.
