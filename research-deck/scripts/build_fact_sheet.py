"""Build the Fact Sheet page from facts.json.

Usage:
  python build_fact_sheet.py "<project>/research/facts.json" ["Out name.html"]

Writes "Fact Sheet.html" next to facts.json. Open it via sheet_server.py so ticks autosave. facts.json:
{
  "title": "Climate-Proof Cities",
  "date": "3 Oct 2026",
  "fx": "USD at 0.92 EUR (open.er-api.com, 3 Oct 2026)",
  "story": "My one or two sentence read of the story (a draft the user can change).",
  "themes": [{"id": "risk", "name": "The risk", "why": "What the reader needs from this theme"}],
  "facts": [{"id": "F1", "theme": "risk", "number": "1,44,000", "meaning": "Heat deaths a year in Indian cities now",
             "source": "World Bank, Towards Resilient and Prosperous Cities in India", "url": "https://...",
             "year": "2025", "fresh": "✓ newest edition", "data": "series 2010/2030/2050", "pick": true}],
  "leads": [{"id": "L1", "text": "Surat's flood early-warning system", "why": "A city that fixed it; a concrete case for the deck"}]
}
"fresh" starting with ✗ / old / no is shown in amber. "pick": true = default coverage.
"""
import json
import pathlib
import sys

SKILL = pathlib.Path(__file__).resolve().parent.parent
TEMPLATE = SKILL / "assets" / "fact-sheet-template.html"


def main():
    src = pathlib.Path(sys.argv[1])
    data = json.loads(src.read_text(encoding="utf-8"))
    ids = [f["id"] for f in data["facts"]]
    assert len(ids) == len(set(ids)), "duplicate fact ids"
    themes = {t["id"] for t in data["themes"]}
    for f in data["facts"]:
        assert f["theme"] in themes, f"{f['id']}: unknown theme {f['theme']}"
    for t in data["themes"]:
        if not any(f["theme"] == t["id"] and f.get("pick") for f in data["facts"]):
            print(f"note: theme '{t['name']}' has no default pick")
    html = TEMPLATE.read_text(encoding="utf-8")
    start, end = html.index("/*DATA*/") + len("/*DATA*/"), html.index("/*END*/")
    blob = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    out = src.parent / (sys.argv[2] if len(sys.argv) > 2 else "Fact Sheet.html")
    out.write_text(html[:start] + blob + html[end:], encoding="utf-8")
    print(f"wrote {out} ({len(ids)} facts, {sum(bool(f.get('pick')) for f in data['facts'])} default picks)")


if __name__ == "__main__":
    main()
