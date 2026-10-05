"""Check that every number you plan to put on a slide really appears in a saved source.

usage: python verify_numbers.py <sources_dir> <claims.json>
claims.json = [{"id": "F1", "number": "72,697", "file": "03 Union Budget Statement 15A.pdf", "context": "crore"}, ...]
  "file" optional (searches every PDF/TXT/CSV/HTML in sources_dir when absent);
  "context" optional: a word that should sit within ~120 characters of the number.

Search summaries invent and blend figures; a number goes on a slide only after this says FOUND.
Number formats are normalised (72,697 = 72697 = 72 697; 3.8 also matches 3·8), so check the hit by eye
when a short number (e.g. "12") could match something unrelated.
"""
import json
import pathlib
import re
import sys


def text_of(path):
    if path.suffix.lower() == ".pdf":
        import pypdf
        try:
            return " ".join((pg.extract_text() or "") for pg in pypdf.PdfReader(str(path)).pages)
        except Exception as e:
            return f"<<unreadable: {e}>>"
    return path.read_text(encoding="utf-8", errors="ignore")


def norm(s):
    s = s.replace("·", ".").replace(" ", "").replace(" ", " ")
    return re.sub(r"(?<=\d)[ ,](?=\d)", "", s)


def main():
    root = pathlib.Path(sys.argv[1]); claims = json.loads(pathlib.Path(sys.argv[2]).read_text(encoding="utf-8"))
    files = [p for p in root.rglob("*") if p.suffix.lower() in (".pdf", ".txt", ".csv", ".html", ".md") and p.name != "Sources.md"]
    cache = {}
    bad = 0
    for c in claims:
        targets = [root / c["file"]] if c.get("file") else files
        num = norm(c["number"]); pat = re.compile(r"(?<![\d.])" + re.escape(num) + r"(?![\d])")
        hit = None
        raw_pat = re.compile(r"(?<![\d.,])" + re.escape(c["number"]) + r"(?![\d])")
        for f in targets:
            if f not in cache:
                raw = " ".join(text_of(f).split()) if f.exists() else ""
                cache[f] = (norm(raw), raw)
            # normalised text catches 72,697 = 72697; raw text catches numbers inside tables, where
            # normalising would glue neighbouring columns together ("1 56 100" -> "156100")
            for t, p_ in ((cache[f][0], pat), (cache[f][1], raw_pat)):
              for m in p_.finditer(t):
                win = t[max(0, m.start() - 120): m.end() + 120]
                if not c.get("context") or c["context"].lower() in win.lower():
                    hit = (f.name, t[max(0, m.start() - 60): m.end() + 60].strip()); break
              if hit: break
            if hit: break
        if hit:
            print(f"FOUND    {c['id']:6s} {c['number']:>12s}  in {hit[0]}: '...{hit[1]}...'")
        else:
            bad += 1
            print(f"MISSING  {c['id']:6s} {c['number']:>12s}  (not in {'that file' if c.get('file') else 'any saved source'}; drop it or find the real source)")
    print(f"\n{len(claims) - bad} of {len(claims)} numbers found in saved sources.")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
