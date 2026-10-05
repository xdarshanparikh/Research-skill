"""Serve the project folder for Fact Sheets, and save picks to disk.

Usage (from the project folder, or via .claude/launch.json "fact-sheets"):
  python sheet_server.py [port] [folder]

GET  serves files as usual.
POST /save  {"sheet": "/research/Fact%20Sheet.html", "state": {...}, "summary": "..."}
     writes "<sheet name> picks.json" and "<sheet name> picks.txt" next to the sheet,
     so picks survive reloads and Claude can read them straight from disk.
"""
import json
import pathlib
import sys
import urllib.parse
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
ROOT = pathlib.Path(sys.argv[2] if len(sys.argv) > 2 else ".").resolve()


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **k):
        super().__init__(*a, directory=str(ROOT), **k)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_POST(self):
        if self.path != "/save":
            return self.send_error(404)
        try:
            body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))).decode("utf-8"))
            rel = urllib.parse.unquote(body["sheet"]).lstrip("/")
            sheet = (ROOT / rel).resolve()
            assert ROOT in sheet.parents and sheet.suffix == ".html", "bad path"
            base = sheet.with_name(sheet.stem + " picks")
            base.with_suffix(".json").write_text(json.dumps(body.get("state", {}), ensure_ascii=False, indent=1), encoding="utf-8")
            base.with_suffix(".txt").write_text(body.get("summary", ""), encoding="utf-8")
            self.send_response(200); self.send_header("Content-Type", "application/json"); self.end_headers()
            self.wfile.write(b'{"ok":true}')
        except Exception as e:  # report, never crash the server
            self.send_response(400); self.end_headers(); self.wfile.write(str(e).encode())


if __name__ == "__main__":
    print(f"serving {ROOT} on http://127.0.0.1:{PORT}")
    ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
