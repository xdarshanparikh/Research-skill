"""Chart generator for presentation slides: writes one standalone <svg> per spec.
Turn it into a PNG for PowerPoint with chart_png.py (or pass a .png output path here).

usage: python charts.py <spec.json> [out.svg | out.png]
spec = {"type": "sankey"|"waterfall"|"bubble"|"scatter"|"slope"|"dumbbell"|"treemap"|"line",
        "theme": "light"|"dark"|"mgl",      # background / template the chart sits on (picks ink + palette + font)
        "w": 1600, "h": 800,                 # canvas size in px (~140 px per inch of the PPTX placeholder)
        "palette": ["#hex", ...],            # optional: your own validated categorical order (last = grey/other)
        "font": "Arial", "cw": 0.55,         # optional: font family and its average char width / font size
        "bg": "#ffffff",                     # optional: paint a background (default transparent)
        ... type-specific fields, see each function below ... }

Rules baked in (dataviz best practice):
- Fixed palette order per background, validated for colour-blind separation (validate_palette.js).
  Colour follows the entity; text stays in ink colours, never the series colour (except the % accent).
- Thin marks, 2 px gaps between fills, rounded ends, solid hairline axes, no dual axes.
- Selective direct labels (no legend boxes needed when every mark is labelled).
- Label decluttering: labels in a column are pushed apart so they never overlap.
Text width is estimated at cw x font size per character (good enough for layout; check the PNG).
"""
import json, math, sys

INK = {"light": {"t": "#1a1a1a", "t2": "#555555", "axis": "#1a1a1a", "accent": "#2a78d6", "edge": "#ffffff", "on_dark": "#ffffff", "on_light": "#1a1a1a"},
       "dark":  {"t": "#ffffff", "t2": "#c9ccd1", "axis": "#ffffff", "accent": "#5b9be3", "edge": "#1f2328", "on_dark": "#ffffff", "on_light": "#1a1a1a"},
       # MGL template (template/mgl-template.pptx): black labels, 404040 secondary text, 888888 axis, blue accent
       "mgl":   {"t": "#000000", "t2": "#404040", "axis": "#888888", "accent": "#0067B2", "edge": "#ffffff", "on_dark": "#ffffff", "on_light": "#000000"}}
# fixed categorical order per background; last slot = "other / context" grey.
# Both pass validate_palette.js (CVD separation + normal-vision floor) on #ffffff / #1f2328.
PALETTE = {"light": ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#b8b8b8"],
           "dark":  ["#3b82d6", "#d95926", "#199e70", "#c28400", "#cf5f8c", "#6b6f76"],
           # template series order: blue, amber, green, red, light blue, grey (CVD >= 7 with direct labels)
           "mgl":   ["#0067B2", "#FCB000", "#51AD30", "#B9382A", "#7FB2D9", "#888888"]}
FONTS = {"mgl": "Aptos, Calibri, Carlito, Arial, sans-serif"}  # per theme; the first installed font wins when rendering
ALIAS = {"white": "light", "char": "dark", "charcoal": "dark"}
CW = 0.55  # average char width / font size (Arial); set per spec with "cw"
FONT = "Arial, Helvetica, sans-serif"


def theme(sp):
    """Resolve theme -> (name, ink, palette); honours a custom "palette" in the spec."""
    th = ALIAS.get(sp.get("theme", "light"), sp.get("theme", "light"))
    if th not in INK: raise SystemExit(f"unknown theme {sp.get('theme')!r}: use light or dark")
    return th, INK[th], sp.get("palette") or PALETTE[th]


def lum(col):
    """WCAG relative luminance of a hex colour."""
    c = col.lstrip("#")
    if len(c) == 3: c = "".join(ch * 2 for ch in c)
    r, g, b = (int(c[i:i + 2], 16) / 255 for i in (0, 2, 4))
    lin = lambda v: v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


def is_light(col):
    """True when dark text reads better on this fill than white text (higher WCAG contrast)."""
    L = lum(col)
    return (L + 0.05) / (lum("#1a1a1a") + 0.05) > 1.05 / (L + 0.05)


def cat(pal, i):
    """Fixed order, never cycled: items past the palette fold into the last slot (grey / context)."""
    return pal[i] if i < len(pal) - 1 else pal[-1]


def tw(s, fs): return len(str(s)) * fs * CW
def esc(s): return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
def fmt(v, d=None):
    if d is None: d = 0 if abs(v) >= 100 or float(v).is_integer() else 1
    s = f"{v:,.{d}f}"
    return s


def pct_round(vals):
    """Largest-remainder rounding so the percentages add up to exactly 100."""
    tot = sum(vals) or 1
    raw = [v * 100 / tot for v in vals]; fl = [int(r) for r in raw]
    for i in sorted(range(len(raw)), key=lambda i: raw[i] - fl[i], reverse=True)[:100 - sum(fl)]: fl[i] += 1
    return fl


def declutter(items, min_gap, lo, hi):
    """items: list of [y_center, ...]. Push apart so centres are >= min_gap apart, within [lo, hi]."""
    order = sorted(range(len(items)), key=lambda i: items[i][0])
    ys = [items[i][0] for i in order]
    for k in range(1, len(ys)): ys[k] = max(ys[k], ys[k - 1] + min_gap)
    over = ys[-1] - hi if ys else 0
    if over > 0:
        ys[-1] -= over
        for k in range(len(ys) - 2, -1, -1): ys[k] = min(ys[k], ys[k + 1] - min_gap)
    for k, i in enumerate(order): items[i][0] = max(lo, ys[k])
    return items


def wrap(spec, body, vb_top=0):
    w, h = spec["w"], spec["h"]
    bg = f'  <rect x="0" y="{vb_top}" width="{w}" height="{h - vb_top}" fill="{spec["bg"]}"/>\n' if spec.get("bg") else ""
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h - vb_top}" viewBox="0 {vb_top} {w} {h - vb_top}">\n'
            f'{bg}  <g font-family="{esc(spec.get("font", FONT))}">\n{body}\n  </g>\n</svg>')


# ---------------------------------------------------------------- SANKEY
def sankey(sp):
    """nodes: [{"id","label","col","color"?,"sub"?}]  links: [{"s","t","v","color"?}]
    options: node_w 22, gap 14, fs 32, pct true (adds % of column total), unit "" (adds value),
             link_color "source"|"target", opacity .55, mid_label "rotate"|"none"."""
    th, ink, pal = theme(sp)
    W, H = sp["w"], sp["h"]; nw = sp.get("node_w", 22); gap = sp.get("gap", 14); fs = sp.get("fs", 32)
    nodes = {n["id"]: dict(n, i=i) for i, n in enumerate(sp["nodes"])}
    for i, n in enumerate(nodes.values()): n.setdefault("color", cat(pal, i))
    for n in nodes.values(): n["in"] = sum(l["v"] for l in sp["links"] if l["t"] == n["id"]); n["out"] = sum(l["v"] for l in sp["links"] if l["s"] == n["id"])
    for n in nodes.values(): n["v"] = max(n["in"], n["out"])
    ncol = max(n["col"] for n in nodes.values()) + 1
    cols = [[n for n in nodes.values() if n["col"] == c] for c in range(ncol)]
    # percentages per column
    lab = {}
    for c in cols:
        ps = pct_round([n["v"] for n in c]) if sp.get("pct", True) else [None] * len(c)
        for n, p in zip(c, ps):
            extra = (f" {p}%" if p is not None else "") + (f" {fmt(n['v'])}{sp['unit']}" if sp.get("unit") else "")
            lab[n["id"]] = (n["label"], extra.strip())
    lw = lambda n: tw(lab[n["id"]][0] + " " + lab[n["id"]][1], fs) + 16
    left = max(lw(n) for n in cols[0]); right = max(lw(n) for n in cols[-1])
    xs = [left + c * (W - left - right - nw) / (ncol - 1) for c in range(ncol)]
    S = min((H - gap * (len(c) - 1)) / sum(n["v"] for n in c) for c in cols)
    for c, x in zip(cols, xs):
        tot = sum(n["v"] for n in c) * S + gap * (len(c) - 1); y = (H - tot) / 2
        for n in c: n["x"], n["y"], n["h"] = x, y, n["v"] * S; y += n["h"] + gap
    out, flows = [], []
    so = {k: 0.0 for k in nodes}; ti = {k: 0.0 for k in nodes}
    links = sorted(sp["links"], key=lambda l: (nodes[l["s"]]["y"], nodes[l["t"]]["y"]))
    for l in links:
        a, b = nodes[l["s"]], nodes[l["t"]]; h = l["v"] * S
        y0 = a["y"] + so[a["id"]]; so[a["id"]] += h
        # incoming order at the target follows source position
        y1 = b["y"] + ti[b["id"]]; ti[b["id"]] += h
        x1, x2 = a["x"] + nw, b["x"]; m = (x1 + x2) / 2
        col = l.get("color") or (b if sp.get("link_color") == "target" else a)["color"]
        flows.append(f'    <path d="M{x1:.1f} {y0:.1f}C{m:.1f} {y0:.1f} {m:.1f} {y1:.1f} {x2:.1f} {y1:.1f}V{y1 + h:.1f}C{m:.1f} {y1 + h:.1f} {m:.1f} {y0 + h:.1f} {x1:.1f} {y0 + h:.1f}Z" fill="{col}" opacity="{sp.get("opacity", .55)}"/>')
    for n in nodes.values():
        out.append(f'    <rect x="{n["x"]:.1f}" y="{n["y"]:.1f}" width="{nw}" height="{max(n["h"], 2):.1f}" rx="4" fill="{n["color"]}"/>')
    for ci, c in enumerate(cols):
        if 0 < ci < ncol - 1:
            for n in c:
                if sp.get("mid_label", "rotate") == "rotate" and n["h"] > tw(n["label"], fs * .9) + 20:
                    cx, cy = n["x"] + nw / 2, n["y"] + n["h"] / 2
                    out.append(f'    <text x="{cx:.1f}" y="{cy:.1f}" dy="{fs * .32:.1f}" text-anchor="middle" font-size="{fs * .9:.0f}" fill="#fff" transform="rotate(-90 {cx:.1f} {cy:.1f})">{esc(n["label"])}</text>')
            continue
        items = [[n["y"] + n["h"] / 2, n] for n in c]
        declutter(items, fs * 1.08 + (fs * .9 if any(n.get("sub") for n in c) else 0) * 0, fs * .5, H - fs * .4)
        for yc, n in items:
            name, extra = lab[n["id"]]; sub = n.get("sub")
            yb = yc + fs * .34 - (fs * .45 if sub else 0)
            if ci == 0: x, anc = n["x"] - 14, "end"
            else: x, anc = n["x"] + nw + 14, "start"
            out.append(f'    <text x="{x:.1f}" y="{yb:.1f}" text-anchor="{anc}" font-size="{fs}" fill="{ink["t"]}">{esc(name)} <tspan fill="{ink["accent"]}">{esc(extra)}</tspan></text>')
            if sub: out.append(f'    <text x="{x:.1f}" y="{yb + fs * .9:.1f}" text-anchor="{anc}" font-size="{fs * .8:.0f}" fill="{ink["t2"]}">{esc(sub)}</text>')
    return wrap(sp, "\n".join(flows + out))


# ---------------------------------------------------------------- WATERFALL
def waterfall(sp):
    """steps: [{"label": "City gas\\npower", "v": 19, "kind": "delta"|"total"}]
    options: fs 40 (values), lfs 26 (labels), total_color, up_color, down_color, zero_color, vmax, prefix/suffix."""
    th, ink, pal = theme(sp)
    W, H = sp["w"], sp["h"]; fs, lfs = sp.get("fs", 40), sp.get("lfs", 26)
    tot_c = sp.get("total_color", pal[0]); up = sp.get("up_color", pal[2]); dn = sp.get("down_color", pal[1])
    steps = sp["steps"]; n = len(steps); nl = max(len(s["label"].split("\n")) for s in steps)
    base = H - nl * lfs * 1.1 - 14; top = fs + 10
    run, bars = 0, []
    for s in steps:
        if s.get("kind") == "total": a, b = 0, s["v"]; run = s["v"]
        else: a, b = run, run + s["v"]; run = b
        bars.append((min(a, b), max(a, b), s))
    vmax = sp.get("vmax") or max(b for _, b, _ in bars)
    k = (base - top) / vmax; gapx = sp.get("gapx", .22)
    bw = W / (n + (n - 1) * gapx); out = [f'    <path d="M0 {base}H{W}" stroke="{ink["axis"]}" stroke-width="3"/>']
    pv = None
    for i, (a, b, s) in enumerate(bars):
        x = i * bw * (1 + gapx); y1, y2 = base - b * k, base - a * k
        h = max(y2 - y1, 4); kind = s.get("kind", "delta")
        col = s.get("color") or (tot_c if kind == "total" else (dn if s["v"] < 0 else (sp.get("zero_color", pal[3]) if abs(s["v"]) * k < 4 else up)))
        out.append(f'    <rect x="{x:.1f}" y="{y1 if h > 4 else (y1 - 2):.1f}" width="{bw:.1f}" height="{h:.1f}" rx="{8 if h > 16 else 2}" fill="{col}"/>')
        if pv is not None:
            out.append(f'    <path d="M{x - bw * gapx:.1f} {pv:.1f}H{x:.1f}" stroke="{ink["axis"]}" stroke-width="2"/>')
        pv = base - (b if (kind == "total" or s["v"] >= 0) else a) * k
        txt = s.get("text") or ((sp.get("prefix", "") + fmt(s["v"]) + sp.get("suffix", "")) if kind == "total" else (("+" if s["v"] > 0 else ("" if s["v"] < 0 else "~")) + fmt(s["v"]) + sp.get("suffix", "")))
        out.append(f'    <text x="{x + bw / 2:.1f}" y="{y1 - 12:.1f}" text-anchor="middle" font-size="{fs}" fill="{ink["t"]}">{esc(txt)}</text>')
        for j, line in enumerate(s["label"].split("\n")):
            lx = min(max(x + bw / 2, tw(line, lfs) / 2 + 2), W - tw(line, lfs) / 2 - 2)
            out.append(f'    <text x="{lx:.1f}" y="{base + 34 + j * lfs * 1.08:.1f}" text-anchor="middle" font-size="{lfs}" fill="{ink["t2"]}">{esc(line)}</text>')
    return wrap(sp, "\n".join(out))


# ---------------------------------------------------------------- BUBBLE (packed)
def bubble(sp):
    """items: [{"label","v","color"?,"text"?}] area proportional to v; packed into a cluster.
    options: fs 34, min_label_r 60. Labels inside when they fit, else below the bubble."""
    th, ink, pal = theme(sp)
    W, H = sp["w"], sp["h"]; fs = sp.get("fs", 34)
    it = sorted(sp["items"], key=lambda d: -d["v"])
    r = [math.sqrt(d["v"]) for d in it]; P = []
    ang = 0
    for i, ri in enumerate(r):  # golden-angle spiral start, then relax
        d = 0 if i == 0 else (r[0] + ri) * (0.9 + 0.15 * i ** .5); ang += 2.39996
        P.append([d * math.cos(ang), d * math.sin(ang)])
    for _ in range(600):
        for i in range(len(P)):
            for j in range(i + 1, len(P)):
                dx, dy = P[j][0] - P[i][0], P[j][1] - P[i][1]; d = math.hypot(dx, dy) or 1e-6
                need = r[i] + r[j] + 0.06 * (r[i] + r[j])
                if d < need:
                    m = (need - d) / 2; ux, uy = dx / d, dy / d
                    P[i][0] -= ux * m; P[i][1] -= uy * m; P[j][0] += ux * m; P[j][1] += uy * m
            P[i][0] *= 0.985; P[i][1] *= 0.985 * (W / H if W < H else 1)  # gentle pull to centre
    minx = min(p[0] - ri for p, ri in zip(P, r)); maxx = max(p[0] + ri for p, ri in zip(P, r))
    miny = min(p[1] - ri for p, ri in zip(P, r)); maxy = max(p[1] + ri for p, ri in zip(P, r))
    s = min(W / (maxx - minx), (H - fs * 1.3) / (maxy - miny))
    ox = (W - (maxx - minx) * s) / 2 - minx * s; oy = (H - fs * 1.3 - (maxy - miny) * s) / 2 - miny * s
    out = []
    for i, (d, p, ri) in enumerate(zip(it, P, r)):
        cx, cy, R = p[0] * s + ox, p[1] * s + oy, ri * s; col = d.get("color", cat(pal, i))
        out.append(f'    <circle cx="{cx:.1f}" cy="{cy:.1f}" r="{R:.1f}" fill="{col}" stroke="{ink["edge"]}" stroke-width="3"/>')
        val = d.get("text", fmt(d["v"]))
        tc = ink["on_light"] if is_light(col) else ink["on_dark"]
        fv = min(fs * 1.5, R * 2 * .8 / len(val) / CW, R * .7); fl = min(fs * .85, R * 2 * .7 / len(d["label"]) / CW)
        if fl >= 20 and fv >= 24:
            out.append(f'    <text x="{cx:.1f}" y="{cy + fv * .2:.1f}" text-anchor="middle" font-size="{fv:.0f}" fill="{tc}">{esc(val)}</text>')
            out.append(f'    <text x="{cx:.1f}" y="{cy + fv * .2 + fl * 1.1:.1f}" text-anchor="middle" font-size="{fl:.0f}" fill="{tc}">{esc(d["label"])}</text>')
        else:
            out.append(f'    <text x="{cx:.1f}" y="{cy + R + fs * .95:.1f}" text-anchor="middle" font-size="{fs * .8:.0f}" fill="{ink["t"]}">{esc(d["label"])} {esc(val)}</text>')
    return wrap(sp, "\n".join(out))


# ---------------------------------------------------------------- SCATTER / BUBBLE ON AXES
def scatter(sp):
    """items: [{"label","x","y","v"?,"color"?}] ; xlabel, ylabel, xmax, ymax, rmax 70. Size = sqrt(v)."""
    th, ink, pal = theme(sp)
    W, H = sp["w"], sp["h"]; fs = sp.get("fs", 28); pad_l, pad_b = 20, 60
    xmax = sp.get("xmax") or max(d["x"] for d in sp["items"]) * 1.1; ymax = sp.get("ymax") or max(d["y"] for d in sp["items"]) * 1.15
    vmax = max(d.get("v", 1) for d in sp["items"]); rmax = sp.get("rmax", 70)
    X = lambda v: pad_l + v / xmax * (W - pad_l - 20 - rmax); Y = lambda v: H - pad_b - v / ymax * (H - pad_b - 20 - rmax)
    out = [f'    <path d="M{pad_l} 10V{H - pad_b}H{W}" fill="none" stroke="{ink["axis"]}" stroke-width="3"/>',
           f'    <text x="{W}" y="{H - pad_b + 44}" text-anchor="end" font-size="{fs}" fill="{ink["t2"]}">{esc(sp.get("xlabel", ""))}</text>',
           f'    <text x="{pad_l + 14}" y="{30}" font-size="{fs}" fill="{ink["t2"]}">{esc(sp.get("ylabel", ""))}</text>']
    circ = [(X(d["x"]), Y(d["y"]), max(8, rmax * math.sqrt(d.get("v", 1) / vmax))) for d in sp["items"]]
    def hits(x0, x1, yc):  # does a label box [x0,x1] at row yc touch any bubble?
        return any(x0 < ccx + cr and x1 > ccx - cr and abs(yc - ccy) < cr + fs * .5 for ccx, ccy, cr in circ)
    for i, d in enumerate(sorted(sp["items"], key=lambda d: -d.get("v", 1))):
        R = max(8, rmax * math.sqrt(d.get("v", 1) / vmax)); col = d.get("color", cat(pal, i))
        cx, cy = X(d["x"]), Y(d["y"])
        out.append(f'    <circle cx="{cx:.1f}" cy="{cy:.1f}" r="{R:.1f}" fill="{col}" opacity=".85" stroke="{ink["edge"]}" stroke-width="2"/>')
        L = tw(d["label"], fs); right = cx + R + 8 + L <= W and not hits(cx + R + 8, cx + R + 8 + L, cy)
        if not right and hits(cx - R - 8 - L, cx - R - 8, cy):  # both sides blocked: label under the bubble
            out.append(f'    <circle cx="{cx:.1f}" cy="{cy:.1f}" r="{R:.1f}" fill="{col}" opacity=".85" stroke="{ink["edge"]}" stroke-width="2"/>'); out.pop()
            out.append(f'    <text x="{cx:.1f}" y="{cy + R + fs:.1f}" text-anchor="middle" font-size="{fs}" fill="{ink["t"]}">{esc(d["label"])}</text>'); continue
        out.append(f'    <text x="{(cx + R + 8) if right else (cx - R - 8):.1f}" y="{cy + fs * .34:.1f}" text-anchor="{"start" if right else "end"}" font-size="{fs}" fill="{ink["t"]}">{esc(d["label"])}</text>')
    return wrap(sp, "\n".join(out))


# ---------------------------------------------------------------- SLOPE
def slope(sp):
    """left, right: column titles; items: [{"label","a","b","hi"?: true}]; unit ''. Highlighted items in accent, rest grey."""
    th, ink, pal = theme(sp)
    W, H = sp["w"], sp["h"]; fs = sp.get("fs", 32)
    vals = [v for d in sp["items"] for v in (d["a"], d["b"])]; lo, hi = sp.get("vmin", min(vals)), sp.get("vmax", max(vals))
    top, bot = fs * 2.2, H - fs
    Y = lambda v: bot - (v - lo) / ((hi - lo) or 1) * (bot - top)
    xl = max(tw(d["label"] + " " + fmt(d["a"]) + sp.get("unit", ""), fs) for d in sp["items"]) + 10
    xr = W - max(tw(fmt(d["b"]) + sp.get("unit", "") + " " + d["label"], fs) for d in sp["items"]) - 10
    out = [f'    <text x="{xl}" y="{fs}" text-anchor="middle" font-size="{fs}" fill="{ink["t2"]}">{esc(sp.get("left", ""))}</text>',
           f'    <text x="{xr}" y="{fs}" text-anchor="middle" font-size="{fs}" fill="{ink["t2"]}">{esc(sp.get("right", ""))}</text>',
           f'    <path d="M{xl} {top - 20}V{bot + 10}M{xr} {top - 20}V{bot + 10}" stroke="{ink["axis"]}" stroke-width="2" opacity=".4"/>']
    L = declutter([[Y(d["a"]), d] for d in sp["items"]], fs * 1.1, fs, H)
    R = declutter([[Y(d["b"]), d] for d in sp["items"]], fs * 1.1, fs, H)
    grey = pal[-1]
    for d in sp["items"]:
        col = d.get("color") or (pal[0] if d.get("hi") else grey)
        out.insert(3, f'    <path d="M{xl} {Y(d["a"]):.1f}L{xr} {Y(d["b"]):.1f}" stroke="{col}" stroke-width="{7 if d.get("hi") else 4}" stroke-linecap="round"/>')
        out.append(f'    <circle cx="{xl}" cy="{Y(d["a"]):.1f}" r="9" fill="{col}"/><circle cx="{xr}" cy="{Y(d["b"]):.1f}" r="9" fill="{col}"/>')
    for (y, d) in L: out.append(f'    <text x="{xl - 18}" y="{y + fs * .34:.1f}" text-anchor="end" font-size="{fs}" fill="{ink["t"]}">{esc(d["label"])} {fmt(d["a"])}{sp.get("unit", "")}</text>')
    for (y, d) in R: out.append(f'    <text x="{xr + 18}" y="{y + fs * .34:.1f}" font-size="{fs}" fill="{ink["t"]}">{fmt(d["b"])}{sp.get("unit", "")} {esc(d["label"])}</text>')
    return wrap(sp, "\n".join(out))


# ---------------------------------------------------------------- DUMBBELL
def dumbbell(sp):
    """items: [{"label","a","b"}]; a_name, b_name (legend); vmax; unit. a = before (grey), b = after (accent)."""
    th, ink, pal = theme(sp)
    W, H = sp["w"], sp["h"]; fs = sp.get("fs", 32); n = len(sp["items"])
    lw = max(tw(d["label"], fs) for d in sp["items"]) + 24
    vmax = sp.get("vmax") or max(max(d["a"], d["b"]) for d in sp["items"]) * 1.08
    X = lambda v: lw + v / vmax * (W - lw - 90); row = (H - fs * 1.6) / n
    ca, cb = pal[-1], pal[0]
    out = [f'    <circle cx="{lw + 12}" cy="{fs * .5}" r="11" fill="{ca}"/><text x="{lw + 30}" y="{fs * .8}" font-size="{fs * .8:.0f}" fill="{ink["t2"]}">{esc(sp.get("a_name", ""))}</text>',
           f'    <circle cx="{lw + 42 + tw(sp.get("a_name", ""), fs * .8) + 30}" cy="{fs * .5}" r="11" fill="{cb}"/><text x="{lw + 60 + tw(sp.get("a_name", ""), fs * .8) + 30}" y="{fs * .8}" font-size="{fs * .8:.0f}" fill="{ink["t2"]}">{esc(sp.get("b_name", ""))}</text>']
    for i, d in enumerate(sp["items"]):
        y = fs * 1.6 + row * i + row / 2
        out.append(f'    <text x="0" y="{y + fs * .34:.1f}" font-size="{fs}" fill="{ink["t"]}">{esc(d["label"])}</text>')
        out.append(f'    <path d="M{X(d["a"]):.1f} {y:.1f}H{X(d["b"]):.1f}" stroke="{ink["axis"]}" stroke-width="4" opacity=".35"/>')
        out.append(f'    <circle cx="{X(d["a"]):.1f}" cy="{y:.1f}" r="13" fill="{ca}"/><circle cx="{X(d["b"]):.1f}" cy="{y:.1f}" r="15" fill="{cb}"/>')
        xb = X(d["b"]) + (24 if d["b"] >= d["a"] else -24)
        out.append(f'    <text x="{xb:.1f}" y="{y + fs * .34:.1f}" text-anchor="{"start" if d["b"] >= d["a"] else "end"}" font-size="{fs}" fill="{ink["t"]}">{fmt(d["b"])}{sp.get("unit", "")}</text>')
    return wrap(sp, "\n".join(out))


# ---------------------------------------------------------------- TREEMAP (squarified)
def treemap(sp):
    """items: [{"label","v","color"?,"text"?}] ; fs 34. Tiles separated by 4 px gaps; labels only where they fit."""
    th, ink, pal = theme(sp)
    W, H = sp["w"], sp["h"]; fs = sp.get("fs", 34)
    it = sorted(sp["items"], key=lambda d: -d["v"]); tot = sum(d["v"] for d in it)
    areas = [d["v"] / tot * W * H for d in it]; rects = []

    def worst(row, w):
        s = sum(row); return max(max(w * w * r / (s * s), s * s / (w * w * r)) for r in row)
    x, y, w, h = 0.0, 0.0, float(W), float(H); i = 0
    while i < len(areas):
        side = min(w, h); row = [areas[i]]; j = i + 1
        while j < len(areas) and worst(row + [areas[j]], side) <= worst(row, side): row.append(areas[j]); j += 1
        s = sum(row)
        if w >= h:
            cw = s / h; yy = y
            for a in row: rects.append((x, yy, cw, a / cw)); yy += a / cw
            x += cw; w -= cw
        else:
            ch = s / w; xx = x
            for a in row: rects.append((xx, y, a / ch, ch)); xx += a / ch
            y += ch; h -= ch
        i = j
    ps = pct_round([d["v"] for d in it]); out = []
    for k, (d, (rx, ry, rw, rh), p) in enumerate(zip(it, rects, ps)):
        col = d.get("color", cat(pal, k))
        tc = ink["on_light"] if is_light(col) else ink["on_dark"]
        out.append(f'    <rect x="{rx + 2:.1f}" y="{ry + 2:.1f}" width="{rw - 4:.1f}" height="{rh - 4:.1f}" rx="10" fill="{col}"/>')
        txt = d.get("text", f"{p}%"); f2 = min(fs, (rw - 30) / max(len(d["label"]), len(txt)) / CW)
        if f2 >= 20 and rh > f2 * 2.4:
            out.append(f'    <text x="{rx + 16:.1f}" y="{ry + 14 + f2 * 1.0:.1f}" font-size="{f2 * 1.3:.0f}" fill="{tc}">{esc(txt)}</text>')
            out.append(f'    <text x="{rx + 16:.1f}" y="{ry + 14 + f2 * 2.05:.1f}" font-size="{f2 * .85:.0f}" fill="{tc}">{esc(d["label"])}</text>')
    return wrap(sp, "\n".join(out))


# ---------------------------------------------------------------- LINE (+ reference lines)
def line(sp):
    """series: [{"name","color"?,"points":[[x,y],...],"label_end": true}] ; refs: [{"y","label","color"?,"dash"?}]
    xmin,xmax,ymin,ymax ; xticks: [[x,"label"]] ; marks: [{"x","y","text","anchor"?,"dy"?}] ; fs 30."""
    th, ink, pal = theme(sp)
    W, H = sp["w"], sp["h"]; fs = sp.get("fs", 30); base = H - 60
    pts = [p for s in sp["series"] for p in s["points"]]
    xmin, xmax = sp.get("xmin", min(p[0] for p in pts)), sp.get("xmax", max(p[0] for p in pts))
    ymin, ymax = sp.get("ymin", 0), sp.get("ymax", max(p[1] for p in pts) * 1.1)
    X = lambda v: 20 + (v - xmin) / ((xmax - xmin) or 1) * (W - 40); Y = lambda v: base - (v - ymin) / ((ymax - ymin) or 1) * (base - fs * 1.4)
    out = [f'    <path d="M0 {base}H{W}" stroke="{ink["axis"]}" stroke-width="3"/>']
    for xv, t in sp.get("xticks", []):
        out.append(f'    <path d="M{X(xv):.1f} {base}v12" stroke="{ink["axis"]}" stroke-width="3"/><text x="{X(xv):.1f}" y="{base + 46}" text-anchor="middle" font-size="{fs}" fill="{ink["t2"]}">{esc(t)}</text>')
    for i, r in enumerate(sp.get("refs", [])):
        c = r.get("color", pal[(i + 1) % len(pal)])
        dash = ' stroke-dasharray="12 9"' if r.get("dash") else ""
        out.append(f'    <path d="M0 {Y(r["y"]):.1f}H{W}" stroke="{c}" stroke-width="4"{dash}/>')
        out.append(f'    <text x="{W}" y="{Y(r["y"]) - 12:.1f}" text-anchor="end" font-size="{fs}" fill="{c}">{esc(r["label"])}</text>')
    for i, s in enumerate(sp["series"]):
        c = s.get("color", pal[i]); P = " ".join(f"{X(x):.1f},{Y(y):.1f}" for x, y in s["points"])
        out.append(f'    <polyline points="{P}" fill="none" stroke="{c}" stroke-width="{s.get("width", 8)}" stroke-linejoin="round" stroke-linecap="round"/>')
        x, y = s["points"][-1]; out.append(f'    <circle cx="{X(x):.1f}" cy="{Y(y):.1f}" r="11" fill="{c}"/>')
    for m in sp.get("marks", []):
        out.append(f'    <circle cx="{X(m["x"]):.1f}" cy="{Y(m["y"]):.1f}" r="11" fill="{m.get("color", pal[0])}"/>')
        out.append(f'    <text x="{X(m["x"]) + m.get("dx", 18):.1f}" y="{Y(m["y"]) + m.get("dy", 10):.1f}" text-anchor="{m.get("anchor", "start")}" font-size="{fs * 1.1:.0f}" fill="{ink["t"]}">{esc(m["text"])}</text>')
    return wrap(sp, "\n".join(out))


TYPES = dict(sankey=sankey, waterfall=waterfall, bubble=bubble, scatter=scatter, slope=slope, dumbbell=dumbbell, treemap=treemap, line=line)

def render(sp):
    global CW
    CW = sp.get("cw", 0.55)
    th = ALIAS.get(sp.get("theme", "light"), sp.get("theme", "light"))
    if "font" not in sp and th in FONTS: sp = dict(sp, font=FONTS[th])
    return TYPES[sp["type"]](sp)


if __name__ == "__main__":
    spec = json.load(open(sys.argv[1], encoding="utf8"))
    specs = spec if isinstance(spec, list) else [spec]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    if out and out.lower().endswith(".png"):
        import pathlib
        sys.path.insert(0, str(pathlib.Path(__file__).parent))
        from chart_png import svgs_to_pngs
        stem = out[:-4]
        names = [out] if len(specs) == 1 else [f"{stem}-{i + 1}.png" for i in range(len(specs))]
        svgs_to_pngs([(render(s), s["w"], s["h"], n) for s, n in zip(specs, names)])
        for n in names: print("wrote", n)
    elif out:
        open(out, "w", encoding="utf8").write("\n".join(render(s) for s in specs)); print("wrote", out)
    else:
        sys.stdout.reconfigure(encoding="utf8"); print("\n".join(render(s) for s in specs))
