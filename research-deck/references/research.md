# Research: newest data, saved and verified

## Search
- Read what the user already has first: their notes, files and any seed link. Don't browse a whole workspace.
- Start with standard web search and run several queries in parallel. Use extended or deep search only for hard-to-find data, very recent events or prices.
- Build a running fact list: **number · meaning in plain words · source · year · newer edition checked?**
- For every number, ask "is there a newer edition?" and search for it by name: the current budget or annual report, the latest statistical release, the newest survey round, the latest regulator update. Users reject old data ("priorities change"), and a number from 5 years ago undermines the whole deck.
- For anything still going on (a market, a scheme, a project, a price), find where it stands **now** as well as the target year.
- Fact-check "it stopped / it failed / it's the first / it's the largest" claims for later developments.

## Source order
| Use | Prefer |
|---|---|
| Numbers | Primary sources: ministries, regulators, statistics offices, central banks, multilaterals (World Bank, IMF, UN agencies, IEA, OECD), company filings and annual reports, standard setters |
| Events and quotes | Reputable news with a named reporter and date |
| Interpretation | Independent analysts and think tanks, named in the footer |

Avoid unsourced aggregators, SEO "statistics" pages and AI-written summaries. If a figure only exists there, trace it back to its origin or drop it.

## Collect chart-ready data
Headline numbers make stat tiles. Charts need structure, so collect it while you research:
- **Flows** (from → to, same unit and period) for a sankey.
- **Base + increments + end** for a waterfall. Check that the steps add up to the end value.
- **Two dates per item** for a slope or dumbbell.
- **Parts of a whole** (5–8 parts) for a treemap. Check they add up to the published total.
- **A time series** plus 1–3 benchmarks for a line with reference lines.
Write down any derived number and how you got it (e.g. "+19 = 103 − 65 − 15 − 4").

## Save every source
Write `research/sources.json`, numbered in the order the slides use them:
```json
[{"name": "03 Annual Statistics 2026", "url": "https://...", "note": "Slides 2, 5: total demand 68.8 bcm"}]
```
```bash
python "<skill>/scripts/save_sources.py" research/sources.json research/sources
```
- PDFs are downloaded as published. Web pages are saved as PDF snapshots (Playwright). `Sources.md` indexes the file, what it supports, the link and its status.
- Fix every **FAILED** (dead link: find a replacement), **BAD** (paywall or error page: save a free article on the same fact and keep the original as "link only") and **CHECK** (blank or scanned page: look at it) row.
- Some sites block automated saving or sit behind a bot check. Never try to get around one. Keep the link as "link only" and save a reputable article that quotes the same figure.
- When several agents research in parallel, give each its own json and number range, then merge them and run the script once at the end.

## Verify every number
Write `research/claims.json` with the numbers you plan to show:
```json
[{"id": "F1", "number": "68.8", "file": "03 Annual Statistics 2026.pdf", "context": "bcm"}]
```
```bash
python "<skill>/scripts/verify_numbers.py" research/sources research/claims.json
```
Every MISSING number is dropped, replaced, or traced to the file that really contains it. Look by eye at hits on short numbers.

## Currency and units
- Use one currency and one unit system per deck. Convert with a live rate (`curl -s https://open.er-api.com/v6/latest/USD`) and state it once ("USD at 0.92 EUR, Oct 2026").
- Keep the source's unit unless converting helps the audience, and say so when you convert.
