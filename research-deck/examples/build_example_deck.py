"""Example: every component of the MGL template kit on one short deck.

usage: python examples/build_example_deck.py [out.pptx]
All numbers are illustrative placeholders, not research findings. Copy a slide's code, then swap in real,
verified data and real sources (see SKILL.md).
"""
import json
import pathlib
import subprocess
import sys

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts"))
from mgl_kit import C, Deck, GRID  # noqa: E402

OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else HERE / "example-deck.pptx")
CH = OUT.parent / "example-charts"; CH.mkdir(exist_ok=True)
SRC = "Source: Illustrative data for the template example; replace with Publisher, Document, Year"

d = Deck()
d.title_slide("Example Research Deck", "Market context in the MGL template", "Month Year")

# 1. Two charts + four stat cards (the template's "outlook" slide)
s = d.content_slide("Supply turns from shortage to surplus within three years",
                    "What the outlooks say: base case vs a 12-month delay case, million tonnes")
d.heading(s, 0.9, 1.42, 7.48, "Market balance, surplus / (shortage), million tonnes")
d.column_chart(s, 0.81, 1.67, 7.57, 2.15, ["FY27", "FY28", "FY29", "FY30", "FY31", "FY32"],
               {"Base case": [0, -3, 28, 59, 70, 104], "Delay case": [0, -17, -29, 4, 37, 68]}, legend=True)
d.heading(s, 0.9, 3.9, 4.8, "Spot price path, $/MMBtu")
d.column_chart(s, 0.81, 4.15, 4.86, 1.73, ["FY27", "FY28", "FY29", "FY30", "FY31", "FY32"],
               {"Spot": [23.8, 15, 10, 10, 10, 10]}, fmt='"$"0.0',
               point_colors=[C["red"], C["amber"]] + [C["green"]] * 4)
d.callout(s, 5.81, 4.22, 2.57, 1.59, "Cross-check", "Forward curve and a second forecaster agree within 10% for FY27-FY29.", "note")
for i, (v, l, ic, tone) in enumerate([("~20%", "of supply offline today; shipping still constrained", "waves", "red"),
                                      ("61%", "storage level vs a 78% seasonal norm", "warehouse", "red"),
                                      ("+53%", "growth in global supply to FY32", "chart-up", "green"),
                                      ("$8-10", "long-run price range across forecasters", "rupee", "green")]):
    d.stat_card(s, 8.71, 1.42 + i * 1.115, 3.72, 1.0, v, l, ic, tone)
d.takeaway(s, "FY27 is the peak-pain year. Prices roughly halve in FY28 and settle from FY29; plan on the delay case and treat an earlier surplus as upside.")
d.sources(s, SRC)
d.notes(s, "Example notes: put the full citation for every number here: Title, Publisher (year): URL. Add derivations.")

# 2. Stat row + charts + info cards (the template's "market" slide)
s = d.content_slide("Gas vehicles: 4 in 10 new registrations, led by one region",
                    "New 3-wheeler and 4-wheeler registrations, FY20 to FY27 year-to-date")
d.stat_row(s, [("1.01 lakh", "gas vehicles added in FY26, 2.9x FY21", "bus", "green"),
               ("40.4%", "share of new vehicles in FY26", "pie", "green"),
               ("+23% a year", "growth in additions, FY21-FY26", "chart-up", "green"),
               ("6.4% → 10.7%", "EV share of new vehicles", "bolt", "green")])
d.heading(s, 0.9, 2.29, 5.42, "Gas vehicles added each year")
d.column_chart(s, 0.81, 2.54, 5.51, 2.2, ["FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26", "FY27 YTD"],
               {"Added": [57597, 35165, 50819, 62838, 77821, 94349, 100801, 37961]},
               point_colors=[C["blue"]] * 7 + [C["amber"]])
d.heading(s, 6.74, 2.29, 5.7, "Share of new vehicles by region, FY22 vs FY26 (%)")
d.column_chart(s, 6.65, 2.54, 5.84, 2.2, ["North", "South", "East", "West", "Central"],
               {"FY22": [22, 29, 34, 36, 50], "FY26": [29, 34, 41, 45, 62]}, colors=["C8C8C8", C["blue"]])
for i, (t, hi, b) in enumerate([("Region A", "40.2k adds, 34.9% share", "Lowest car share"),
                                ("Region B", "54.7k adds, 43.8% share", "Largest pool"),
                                ("Region C", "5.9k adds, 61.6% share", "Highest penetration")]):
    d.info_card(s, 0.9 + i * 2.94, 4.85, 2.8, 1.07, t, b, highlight=hi)
d.callout(s, 9.71, 4.85, 2.72, 1.07, "Since April", "the gas share slipped to 36-38% while EV rose to ~10%.", "watch")
d.takeaway(s, "Gas vehicles out-grew every fuel since FY21 and every region raised its share; watch the EV shift in cars.", y=6.06, h=0.58)
d.sources(s, SRC)

# 3. Scorecard table + bullets (the template's "segments" slide)
s = d.content_slide("Segments: cars and cabs now drive growth; autos slumped",
                    "Additions and penetration by segment, with a traffic-light read")
d.heading(s, 0.9, 1.42, 5.4, "Additions by segment (thousands)")
d.column_chart(s, 0.81, 1.67, 5.5, 3.2, ["FY22", "FY23", "FY24", "FY25", "FY26"],
               {"Autos": [21, 22, 24, 26, 27], "Cars": [12, 16, 22, 30, 38], "Cabs": [10, 14, 20, 25, 29]},
               stacked=True, fmt="0")
d.heading(s, 6.28, 1.42, 6.15, "Segment scorecard")
d.table(s, 6.28, 1.67, 6.15, [
    ["Segment", "FY26 adds", "Share FY22 → FY26", "EV share", "Read"],
    ["Cars", "38,056", ("24.7% → 26.5%", "good"), ("9.7%", "watch"), ("Growth; EVs rising", "watch")],
    ["Cabs", "28,849", ("64.3% → 83.9%", "good"), "1.2%", ("Dominant", "good")],
    ["Autos", "27,283", ("88.2% → 89.6%", "good"), "1.7%", ("Volumes falling", "bad")],
    ["Small goods", "6,373", ("47.8% → 31.0%", "bad"), "3.0%", ("Diesel still 64%", "watch")],
], col_w=[1.2, 0.95, 1.45, 0.85, 1.7])
d.bullets(s, 6.28, 3.75, 6.15, 1.2, [("Cars and cabs:", "two-thirds of additions, replacing autos as the engine."),
                                      ("Small goods:", "the biggest untapped pool, still mostly diesel.")])
d.callout(s, 0.9, 5.2, 11.53, 0.5, "Watch:", "auto registrations fell sharply after the price rise; check permit data before drawing conclusions.", "risk")
d.takeaway(s, "Cars and cabs now supply two-thirds of additions. Cabs are almost all gas; cars are the battleground with EVs.")
d.sources(s, SRC)

# 4. A charts.py image chart (sankey) beside icon cards with tags
spec = json.loads((HERE.parent / "assets" / "chart-gallery-specs.json").read_text(encoding="utf-8"))[0]
spec.update(theme="mgl", w=1050, h=600, fs=22)
for n in spec["nodes"]:
    n.pop("color", None)
    if n["col"] == 1: n["color"] = "#888888"
    if n["col"] == 2: n["color"] = "#0067B2"
for l in spec["links"]:
    if "color" in l: l["color"] = "#0067B2"
(CH / "sankey.json").write_text(json.dumps(spec), encoding="utf-8")
subprocess.run([sys.executable, str(HERE.parent / "scripts" / "charts.py"), str(CH / "sankey.json"), str(CH / "sankey.png")], check=True)
s = d.content_slide("Where supply comes from and where it goes",
                    "Supply by source and use by sector, billion cubic metres, latest year")
d.heading(s, 0.9, 1.42, 7.3, "Supply → uses (bcm)")
d.picture(s, CH / "sankey.png", 0.9, 1.72, 7.3, 4.15)
for i, (t, tag, kind, ic, b) in enumerate([("Domestic gas", "Tailwind", "good", "home", "Half of supply; priority allocation to homes and transport."),
                                           ("Imports", "Watch", "watch", "ship", "One route carries a large share; contract cover is flat."),
                                           ("Industry", "Upside", "good", "factory", "Switching from liquid fuels pays below a price threshold."),
                                           ("Power", "Mixed", "watch", "bolt", "Runs at peaks only; renewables take the base load.")]):
    y = 1.42 + i * 1.12
    d.info_card(s, 8.5, y, 3.93, 1.0, t, b, icon=ic, tag=(tag, kind))
d.takeaway(s, "Domestic gas covers half the market; the import half is where price and route risk sit.")
d.sources(s, SRC)

# 5. Doughnut + line chart + bullets
s = d.content_slide("Prices peak this year, then settle below the cost of oil products",
                    "Spot price path vs benchmarks, and where the scheme money goes")
d.heading(s, 0.9, 1.42, 6.6, "Spot price vs benchmarks, $/MMBtu")
d.line_chart(s, 0.81, 1.67, 6.7, 3.9, ["FY27", "FY28", "FY29", "FY30", "FY31", "FY32"],
             {"Spot (forecast)": [23.8, 15, 10, 10, 9, 8], "Fuel oil parity": [8.3, 8.3, 8.3, 8.3, 8.3, 8.3],
              "Propane parity": [9.3, 9.3, 9.3, 9.3, 9.3, 9.3]}, fmt='"$"0.0', end_labels=["Spot (forecast)"])
d.heading(s, 7.9, 1.42, 4.5, "Outlay by component (crore)")
d.doughnut_chart(s, 7.8, 1.67, 4.63, 2.6, ["Offtake support", "Capital assistance", "Pipelines", "Credit guarantee"],
                 [10410, 7063, 5133, 1125])
d.bullets(s, 7.9, 4.4, 4.5, 1.4, [("Switching window:", "gas beats fuel oil below ~$8.3 from FY29."),
                                   ("Contracts:", "offer multi-year prices so industry can plan the switch.")])
d.takeaway(s, "From FY29 the spot price sits at or below oil-product parity, opening the industrial switching window.")
d.sources(s, SRC)

d.closing_slide()
d.save(OUT)
print("wrote", OUT)
