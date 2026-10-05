# Sources on slides

## The footer
- **Every content slide gets one**, including the title slide if it shows a number. Use one text box named `Source`, in the same position on every slide, one or two lines, in a small grey type. `pptx_helpers.add_source` does this, and `check_deck.py` checks it.
- Format: `Source: Publisher, Document, Year`, with several sources separated by `;`. Write publisher names the way they appear, without expanding them or adding tags ("World Bank", not "World Bank (WB)").
- Drop sub-titles, "via" chains and page or table numbers from the footer. They go in the speaker notes.
- Abbreviations: add `ABBR = full name` to the footer of the **first** slide that uses the term, and only there.
- Derived numbers: add a short note when it changes the meaning ("Shares derived from 2025 data").
- Currency conversions: state the rate once ("USD at 0.92 EUR, Oct 2026").
- An old year that really is the latest available: write "(latest available)", or put `data-checked: YYYY` in the notes.

## Speaker notes
For each slide, add the full citation for every number on it: `Title, Publisher (year): URL`. Add any derivation ("+19 = 103 − 65 − 15 − 4") and caveats (a different period, a scaled total).

## Sources slide and Sources.md
- The last content slide, or an appendix, lists every source: `Publisher, Document title, Year`, grouped by slide or theme.
- `research/sources/Sources.md` (written by `save_sources.py`) holds the file, what it supports, the full link and its status. Hand it over with the deck.
- Link-only sources (paywalled or bot-protected) are marked so in `Sources.md`, together with the saved article that quotes the same figure.
