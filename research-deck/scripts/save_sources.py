"""Save every source behind a deck into <project>/research/sources/ so the evidence lives with the slides.

usage: python save_sources.py <sources.json> <out_dir>
sources.json = [{"name": "short-file-name", "url": "https://...", "note": "what it backs up"}, ...]
- URLs ending in .pdf (or serving application/pdf) are downloaded as-is.
- Web pages are printed to PDF with Playwright (a snapshot, in case the page changes or dies).
- Writes Sources.md (index: file, URL, what it supports, date saved) in out_dir.
"""
import json, os, sys, datetime, urllib.request, ssl

src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
items = json.load(open(src, encoding="utf8"))
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}
ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE

def try_pdf(url, path):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=60, context=ctx) as r:
        data = r.read()
    if data[:4] != b"%PDF":
        return False
    open(path, "wb").write(data); return True

rows, pages = [], []
for it in items:
    path = os.path.join(out, it["name"] + ".pdf")
    status = ""
    if os.path.exists(path) and os.path.getsize(path) > 2000:
        status = "already saved"
    else:
        try:
            status = "PDF downloaded" if try_pdf(it["url"], path) else ""
        except Exception as e:
            status = ""
        if not status:
            pages.append((it, path)); continue
    rows.append((it, path, status))

if pages:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctxp = b.new_context(user_agent=UA["User-Agent"], viewport={"width": 1280, "height": 1600})
        for it, path in pages:
            pg = ctxp.new_page()
            try:
                pg.goto(it["url"], wait_until="domcontentloaded", timeout=60000)
                pg.wait_for_timeout(2500)
                pg.pdf(path=path, format="A4", print_background=True)
                rows.append((it, path, "page saved as PDF"))
            except Exception as e:
                rows.append((it, None, "FAILED: " + str(e)[:80]))
            pg.close()
        b.close()

# flag snapshots that saved a paywall / error / blank page instead of the article
BAD = ("access denied", "page not found", "404", "subscribe to read", "you don't have permission", "not found")
def check(path):
    try:
        import pypdf
        r = pypdf.PdfReader(path)
        t = " ".join("".join((pg.extract_text() or "") for pg in r.pages[:2]).split()).lower()
        if len(t) < 200 and len(r.pages) <= 1:
            return "CHECK: almost no text (blank page or scanned PDF)"
        if any(b in t[:400] for b in BAD):
            return "BAD: paywall or error page, find a free article on the same fact"
    except Exception as e:
        return "CHECK: unreadable PDF"
    return ""
rows = [(it, path, (status + ("; " + check(path) if path and check(path) else ""))) for it, path, status in rows]

today = datetime.date.today().isoformat()
md = ["# Sources", "", f"Saved {today}. Each file backs up a number or claim in the deck.", "",
      "| File | What it supports | Link | Status |", "|---|---|---|---|"]
for it, path, status in rows:
    fn = os.path.basename(path) if path else "(not saved)"
    md.append(f"| {fn} | {it.get('note','')} | {it['url']} | {status} |")
open(os.path.join(out, "Sources.md"), "w", encoding="utf8").write("\n".join(md) + "\n")
for it, path, status in rows:
    print(f"{status:20s} {it['name']}")
