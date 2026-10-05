"""Rebuild "MGL CBG Business Case" in the MGL template look, keeping every slide's content and wording.

usage: python build_deck.py [out.pptx]
Source deck: original.pptx (same folder). Pictures reused from it are in assets/.
"""
import pathlib
import sys

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent.parent / "research-deck" / "scripts"))
from mgl_kit import C, Deck, GRID  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402

A = HERE / "assets"
OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else HERE / "MGL_CBG_Business_Case_MGL_design.pptx")
L, W = GRID["left"], GRID["width"]
d = Deck()


def footnote(s, text, y=6.72):
    """Original footnotes sit where the template puts its source line."""
    d.sources(s, text, y=y)


def line(s, x, y, w, h, color=C["axis"]):
    return d.rect(s, x, y, w, h, color, name="Connector")


# 1 ------------------------------------------------------------------ title
d.title_slide("Corporate Presentation", "Mahanagar Gas Limited", "05th Oct 2026")

# 2 ------------------------------------------------------------------ section
d.section_slide("MGL Overview")

# 3 ------------------------------------------------------------------ Our History (timeline)
s = d.content_slide("Our History")
steps = [("1976", "Mumbai High", "Natural Gas Flared off", "waves"),
         ("Late 1980’s", "GOI", "Planned to utilize the gas instead of flaring and nominated GAIL to implement CGD project in Mumbai", "buildings"),
         ("Early 1990’s", "GAIL Partnered British Gas", "To leverage on their technical Know-how and robust safety practices of BG", "handshake"),
         ("8th May 1995", "MGL", "Was incorporated as a JV between GAIL and BG", "flame"),
         ("MGL Today", "", "Mahanagar Gas Limited (MGL) sowed the seeds of success in city  gas  distribution  in India and  today stands as a leader in the CGD sector, known for its excellence in operations and safety.", "flag")]
ly = 3.62
line(s, L + 0.6, ly + 0.33, W - 1.2, 0.04, C["green"])
cw = W / 5
for i, (yr, head, body, ic) in enumerate(steps):
    cx = L + i * cw + cw / 2
    d.icon(s, cx - 0.36, ly, 0.72, ic, C["green"] if i < 4 else C["blue"])
    up = i % 2 == 0
    by = 1.45 if up else 4.85
    d.text(s, cx - cw / 2 + 0.1, ly - 0.42 if up else ly + 0.86, cw - 0.2, 0.3, yr, 13, C["dark_green"], bold=True, align="c")
    d.rect(s, cx - cw / 2 + 0.08, by, cw - 0.16, 1.62 if up else 1.55, C["card"], name="Card")
    yy = by + 0.12
    if head:
        d.text(s, cx - cw / 2 + 0.2, yy, cw - 0.4, 0.3, head, 11, C["ink"], bold=True, align="c"); yy += 0.36
    d.text(s, cx - cw / 2 + 0.2, yy, cw - 0.4, by + 1.55 - yy, body, 9, C["text"], align="c")

# 4 ------------------------------------------------------------------ Company Overview
s = d.content_slide("Company Overview")
facts = [("calendar", "Mahanagar Gas Limited (MGL) was incorporated on 8th May 1995 with GAIL and British Gas as its promoters."),
         ("exchange", "In August 2019, BGAPH completely divested its equity."),
         ("chart-up", "MGL got listed on 1st July 2016; Government of Maharashtra (GoM) increased its shareholding to 10%."),
         ("pie", "GAIL owns 32.5% , GoM owns 10% stake in the company, with the remaining stake distributed amongst FPIs/NRIs (~24.25%), DI/ Investors (~23.82%), Others (9.4%).")]
for i, (ic, t) in enumerate(facts):
    y = 1.3 + i * 1.12
    d.rect(s, L, y, 6.1, 1.0, C["card"], name="Card")
    d.icon(s, L + 0.15, y + 0.22, 0.56, {"calendar": "flag"}.get(ic, ic))
    d.text(s, L + 0.88, y + 0.1, 5.05, 0.8, t, 11, C["ink"], anchor="m")
d.heading(s, 7.3, 1.3, 5.1, "Shareholding")
d.doughnut_chart(s, 7.2, 1.55, 5.2, 2.55, ["GAIL", "GoM", "FPIs/NRIs", "DI/ Investors", "Others"],
                 [32.5, 10, 24.25, 23.82, 9.4], fmt='0.0#"%"')
d.rect(s, 7.3, 4.25, 5.13, 1.95, C["green_tint"], name="Card")
d.icon(s, 7.45, 4.4, 0.52, "bulb")
d.text(s, 8.1, 4.43, 4.2, 0.45, "Corporate Vision", 14, C["dark_green"], bold=True, anchor="m")
d.bullets(s, 7.5, 5.0, 4.85, 1.15, ["To be world class, consumer and environment-friendly company, an employer of choice",
                                    "Committed to provide safe, efficient, and reliable energy",
                                    "While creating value for all our stakeholders"], size=10)

# 5 ------------------------------------------------------------------ Business Verticals (org chart)
s = d.content_slide("MGL Business Verticals")
top = d.rect(s, 5.17, 1.35, 3.0, 0.62, C["green"], name="Org box")
d.text(s, 5.17, 1.35, 3.0, 0.62, "Mahanagar Gas Limited", 14, C["white"], bold=True, align="c", anchor="m")
line(s, 6.66, 1.97, 0.03, 0.38)
line(s, 2.35, 2.33, 8.65, 0.03)
cols = [("Merger & Acquisition", ["Unison Enviro(UEPL)"], "handshake"),
        ("Strategic Investment", ["3ev Industries Pvt. Ltd."], "chart-up"),
        ("Joint Ventures", ["Mahanagar LNG Pvt. Ltd. (MLPL)", "IBC India (IBCIPL)"], "layers")]
for i, (head, kids, ic) in enumerate(cols):
    cx = 2.35 + i * 4.32
    line(s, cx - 0.01, 2.33, 0.03, 0.35)
    d.rect(s, cx - 1.6, 2.68, 3.2, 0.7, C["blue"], name="Org box")
    d.icon(s, cx - 1.48, 2.79, 0.48, ic, C["green"])
    d.text(s, cx - 0.9, 2.68, 2.4, 0.7, head, 13, C["white"], bold=True, anchor="m")
    for j, k in enumerate(kids):
        y = 3.85 + j * 1.0
        line(s, cx - 0.01, 3.38 if j == 0 else 3.85 + (j - 1) * 1.0 + 0.72, 0.03, 0.47 if j == 0 else 0.28)
        d.rect(s, cx - 1.45, y, 2.9, 0.72, C["card"], line=C["green"], name="Org box")
        d.text(s, cx - 1.4, y, 2.8, 0.72, k, 12, C["ink"], bold=True, align="c", anchor="m")

# 6 ------------------------------------------------------------------ Company Overview (numbers)
s = d.content_slide("Company Overview")
d.stat_card(s, L, 1.3, 3.7, 1.15, "4.54 MMSCMD", "Average Sales¹", "chart-up", "blue")
d.stat_card(s, L + 3.92, 1.3, 3.7, 1.15, "5.34 MMSCMD", "Peak Sales · September 2026", "bolt", "green")
d.stat_card(s, L + 7.84, 1.3, 3.69, 1.15, "Gas Sourcing", "Mix of Domestic, Term and Spot gas", "ship", "ink", value_size=16)
d.heading(s, L, 2.75, W, "Financial Metrics FY 26* (in INR Cr.)", 13)
fin = [("8,240²", "Revenue", "rupee", C["amber"]), ("7,135²", "Fixed Asset", "buildings", C["axis"]),
       ("847", "PAT", "chart-up", C["green"]), ("10,882³", "Market Cap", "pie", C["blue"])]
for i, (v, lab, ic, col) in enumerate(fin):
    x = L + i * 2.94
    d.rect(s, x, 3.15, 2.79, 2.75, C["card"], name="Card")
    d.icon(s, x + 0.2, 3.35, 0.7, ic, col)
    d.text(s, x + 0.2, 4.3, 2.4, 0.75, v, 30, col if col != C["axis"] else C["ink"], bold=True)
    d.text(s, x + 0.2, 5.1, 2.4, 0.5, lab, 14, C["text"])
footnote(s, "(1) Sales Volume is April to August  FY 2026-27 annual average   (2) Fixed Assets is as on 31st March 2025 and Revenue from operations (Net) excluding Other Income   (3) Market Cap as on September 09, 2026", y=6.25)

# 7 ------------------------------------------------------------------ Area of Operation
s = d.content_slide("Area of Operation & Infrastructure")
d.picture(s, A / "map-maharashtra.png", L - 0.2, 1.3, 4.2, 2.85)
d.picture(s, A / "map-karnataka.png", L + 0.2, 4.05, 2.6, 2.2)
for i, (col, lab) in enumerate([(C["green"], "MGL’s existing Geographical Areas (GA)"), ("A65CC6", "GAs of erstwhile UEPL")]):
    d.rect(s, L + 2.85, 5.0 + i * 0.55, 0.2, 0.2, col, shape=MSO_SHAPE.OVAL, name="Legend")
    d.text(s, L + 3.15, 4.96 + i * 0.55, 1.85, 0.5, lab, 9, C["text"])
grid = [[("45,691", "Area (sq Km)", "layers"), ("47,96,180", "Urban Households", "home")],
        [("10", "CGS & LNG Station (No)", "fuel-pump"), ("522", "CNG Stations (No)", "fuel-pump"),
         ("759", "Steel Pipeline Length (Km)", "network"), ("7,757", "PE Pipeline Length (Km)", "network")],
        [("13,30,819", "CNG Vehicles (No)", "bus"), ("32,61,514", "PNG Household Connections", "home"),
         ("6,307", "Industrial Commercial Customers", "factory")]]
gx, gw = 6.0, 6.43
for r, row in enumerate(grid):
    cw = (gw - 0.12 * (len(row) - 1)) / len(row)
    for c, (v, lab, ic) in enumerate(row):
        x, y = gx + c * (cw + 0.12), 1.3 + r * 1.62
        d.rect(s, x, y, cw, 1.48, C["green_tint"] if r == 0 else C["card"], name="Card")
        d.icon(s, x + 0.14, y + 0.14, 0.46, ic, C["blue"] if r == 1 else C["green"])
        d.text(s, x + 0.14, y + 0.66, cw - 0.25, 0.4, v, 17 if len(row) < 4 else 15, C["ink"], bold=True)
        d.text(s, x + 0.14, y + 1.04, cw - 0.25, 0.4, lab, 8.5, C["text"])
footnote(s, "Note: Data till  31st March 2026")

# 8 ------------------------------------------------------------------ Network expansion to integrated operations
s = d.content_slide("From network expansion to integrated operations")
kp = [("10", "CGS & LNG Station (No)"), ("522", "CNG Stations (No)"), ("759", "Steel Pipeline Length (Km)"),
      ("7,757", "PE Pipeline Length (Km)"), ("13,30,819", "CNG Vehicles (No)"), ("32,61,514", "PNG Household Connections"),
      ("6,307", "Industrial Commercial Customers")]
kw = (W - 0.1 * 6) / 7
for i, (v, lab) in enumerate(kp):
    x = L + i * (kw + 0.1)
    d.rect(s, x, 1.0, kw, 0.72, C["green_tint"], name="Card")
    d.text(s, x + 0.1, 1.04, kw - 0.15, 0.3, v, 13, C["ink"], bold=True)
    d.text(s, x + 0.1, 1.35, kw - 0.15, 0.35, lab, 7.5, C["text"])
# left: ICCC
d.rect(s, L, 1.85, 3.55, 4.75, C["card"], name="Card")
d.icon(s, L + 0.12, 1.95, 0.42, "network", C["blue"])
d.text(s, L + 0.62, 1.95, 2.85, 0.42, "Integrated Command & Control Centre - Mahape", 10, C["ink"], bold=True, anchor="m")
d.text(s, L + 0.15, 2.45, 3.3, 0.35, "MGL moved from individual systems to an integrated operational view.", 8.5, C["text"], italic=True)
chips = ["SCADA", "GIS", "IVMS", "CNG Transport Software", "Forecourt Automation", "AMR"]
for i, ch in enumerate(chips):
    x, y = L + 0.15 + (i % 3) * 1.1, 2.9 + (i // 3) * 0.5
    d.rect(s, x, y, 1.02, 0.42, C["white"], line=C["blue"], name="Chip")
    d.text(s, x, y, 1.02, 0.42, ch, 7.5, C["blue"], bold=True, align="c", anchor="m")
d.rect(s, L + 0.15, 3.98, 3.25, 0.32, C["blue"], name="Chip")
d.text(s, L + 0.15, 3.98, 3.25, 0.32, "ICCC enables", 9, C["white"], bold=True, align="c", anchor="m")
en = ["Centralised monitoring", "Faster emergency response", "Resource mobilisation", "Outage management",
      "Nearest emergency response team identification", "Operational dashboards and analytics",
      "Future SAP integration for maintenance planning and expenditure management"]
d.bullets(s, L + 0.15, 4.4, 1.6, 2.1, en[:4], size=8)
d.bullets(s, L + 1.8, 4.4, 1.65, 2.1, en[4:], size=8)
# right: three rows
rows = [("Projects", "network", ["CNG station network scaled from 290 in FY 22 to 500+ in FY 26",
                                 "~3.2 Mn DPNG connections, accounting for ~20% of India’s ~16 Mn customer base",
                                 "Adoption of MLC1 piping (~25-30% cheaper) and upgradation in metering technologies"]),
        ("O&M", "bulb", ["GIS mapping for asset visibility, maintenance and emergency handling",
                         "In-house SCADA + IVMS2 tool to cut dry-outs at daughter booster stations",
                         "Powering COCO CNG stations through renewable energy"]),
        ("HSE", "shield", ["Integrated ISO 45001, 14001 & 9001 framework; zero LTIFR3, 96% employees trained in FY25",
                           "“No STC4 No Work”, Dial-before-Dig and near-miss reporting embedded",
                           "Digital safety audits and HIRA / QRA / HAZOP reviews across projects and operations"])]
rx = L + 3.7
for r, (lab, ic, items) in enumerate(rows):
    y = 1.85 + r * 1.6
    d.rect(s, rx, y, 1.3, 1.48, C["blue"], name="Row label")
    d.icon(s, rx + 0.4, y + 0.2, 0.5, ic, C["green"])
    d.text(s, rx, y + 0.82, 1.3, 0.4, lab, 13, C["white"], bold=True, align="c")
    iw = (W - 3.7 - 1.3 - 0.1 * 3) / 3
    for j, t in enumerate(items):
        x = rx + 1.4 + j * (iw + 0.1)
        d.rect(s, x, y, iw, 1.48, C["card"], name="Card")
        d.rect(s, x + 0.12, y + 0.14, 0.24, 0.24, C["green"], shape=MSO_SHAPE.OVAL, name="Tick")
        d.text(s, x + 0.12, y + 0.48, iw - 0.22, 0.95, t, 8.5, C["ink"])
footnote(s, "1. Multi Layer Composite | 2. In Vehicle Monitoring System | 3. Lost Time Injury Frequency Rate | 4. Safety and Technical Competency")

# 9 ------------------------------------------------------------------ Digital platform
s = d.content_slide("The digital platform is now ready for the next phase of MGL")
cw = (W - 0.3) / 3
# Utkarsh
x = L
d.info_card(s, x, 1.0, cw, 2.75, "Project Utkarsh - RISE with SAP", "", icon="layers",
            highlight="First CGD company in India to implement RISE with SAP on private cloud - SAP ECC 6.0 to SAP S/4HANA.")
d.bullets(s, x + 0.66, 2.05, cw - 0.8, 1.65, ["Modern enterprise backbone", "Reduced IT complexity", "Cloud-based scalability",
                                              "Improved integration", "Platform for future expansion and innovation"], size=8.5)
# PRISM
x = L + cw + 0.15
d.info_card(s, x, 1.0, cw, 2.75, "Project PRISM", "", icon="people",
            highlight="New digital customer/operations architecture with Salesforce and Appstrail. Phase 1:")
for i, ch in enumerate(["CRM", "Collect360", "Meter Reading App", "Five O&M apps", "Single customer view", "Breaking data silos"]):
    cx, cy = x + 0.2 + (i % 2) * ((cw - 0.4) / 2 + 0.05), 2.05 + (i // 2) * 0.42
    d.rect(s, cx, cy, (cw - 0.5) / 2, 0.34, C["white"], line=C["blue"], name="Chip")
    d.text(s, cx, cy, (cw - 0.5) / 2, 0.34, ch, 8, C["blue"], bold=True, align="c", anchor="m")
d.rect(s, x + 0.2, 3.32, cw - 0.4, 0.32, C["blue"], name="Chip")
d.text(s, x + 0.2, 3.32, cw - 0.4, 0.32, "First AI-enabled use case", 8.5, C["white"], bold=True, align="c", anchor="m")
# Digital HSE
x = L + 2 * (cw + 0.15)
d.info_card(s, x, 1.0, cw, 2.75, "Digital HSE", "", icon="shield")
for i, t in enumerate(["Digital safety audits", "Real-time compliance", "Incident tracking", "Root-cause analysis", "Corrective-action monitoring"]):
    y = 1.45 + i * 0.36
    d.rect(s, x + 0.66, y, 0.26, 0.26, C["blue"], shape=MSO_SHAPE.OVAL, name="Num")
    d.text(s, x + 0.66, y, 0.26, 0.26, str(i + 1), 8, C["white"], bold=True, align="c", anchor="m")
    d.text(s, x + 1.02, y, cw - 1.15, 0.28, t, 9, C["ink"], anchor="m")
d.text(s, x + 0.66, 3.3, cw - 0.8, 0.4, "Supports virtual audits, photographic evidence and centralised HSE monitoring.", 8, C["text"], italic=True)
# Digital O&M flow
y = 4.0
d.rect(s, L, y, 1.45, 0.95, C["card"], name="Card")
d.icon(s, L + 0.08, y + 0.24, 0.46, "network", C["blue"])
d.text(s, L + 0.6, y, 0.85, 0.95, "Digital O&M", 11, C["ink"], bold=True, anchor="m")
d.rect(s, L + 1.6, y, 2.55, 0.95, C["blue"], name="Chip")
d.text(s, L + 1.65, y, 2.45, 0.95, "SCADA + GIS + AMR + IVMS\n+ ICCC + SAP + PRISM", 10, C["white"], bold=True, align="c", anchor="m")
steps = ["Predictive maintenance", "Asset optimisation", "Demand forecasting", "Automated workflows", "AI-enabled operations"]
sw = (W - 4.3) / 5
for i, t in enumerate(steps):
    xx = L + 4.3 + i * sw
    shp = d.rect(s, xx, y, sw + 0.12, 0.95, C["green"] if i == 4 else C["green_tint"], shape=MSO_SHAPE.CHEVRON, name="Step")
    d.text(s, xx + 0.32, y, sw - 0.45, 0.95, t, 9, C["white"] if i == 4 else C["dark_green"], bold=True, align="c", anchor="m")
d.takeaway(s, "Over the last 5-6 years, MGL has built more than physical infrastructure - it has built an execution platform combining network scale, asset integrity, HSE, project execution, digital systems and operational data.",
           y=5.2, h=0.95, label="FROM CAPABILITY BUILDING TO CAPABILITY LEVERAGE")

# 10 ----------------------------------------------------------------- Financial performance
s = d.content_slide("Financial performance over past 5 years", "Values in INR Crores")
fy = ["FY21-22", "FY22-23", "FY 23-24", "FY 24-25", "FY 25-26"]
charts = [("Revenue", [3970.6343, 7032.84, 7037.25, 7774.0, 9178.97], C["blue"]),
          ("PAT", [596.9472, 790.05, 1289.07, 1044.89, 846.82], C["green"]),
          ("EBITDA", [923.9435, 1183.4, 1842.8, 1509.66, 1450.02], C["amber"]),
          ("Gross Margin", [1514.4085, 1863.64, 2626.63, 2582.52, 2607.97], C["light_blue"])]
for i, (t, v, col) in enumerate(charts):
    x, y = L + (i % 2) * 5.85, 1.42 + (i // 2) * 2.62
    d.heading(s, x, y, 5.6, t)
    d.column_chart(s, x - 0.1, y + 0.25, 5.75, 2.25, fy, {t: v}, colors=[col], fmt="#,##0.00", gap=60)

# 11 + 12 ------------------------------------------------------------ Infrastructure and volumes (same content twice)
def volumes(title):
    s = d.content_slide(title)
    yrs = ["FY 21", "FY 22", "FY 23", "FY 24", "FY 25", "FY 26"]
    d.heading(s, L, 1.3, 5.6, "CNG Customers (end users) Growth (in '000s)")
    d.column_chart(s, L - 0.1, 1.55, 5.75, 2.1, yrs, {"CNG Customer": [795, 855, 925, 1005, 1124, 1284]}, colors=[C["blue"]])
    d.heading(s, L + 5.95, 1.3, 5.6, "PNG Customers (end users) Growth (in '000s)")
    d.column_chart(s, L + 5.85, 1.55, 5.75, 2.1, yrs, {"PNG Customer": [1257, 1381, 1520, 1697, 1875, 2057]}, colors=[C["green"]])
    d.heading(s, L, 3.85, 8.5, "Sales Volume")
    d.column_chart(s, L - 0.1, 4.1, 8.7, 2.5, yrs,
                   {"CNG": [1.42, 2.16, 2.58, 2.71, 3.05, 3.267], "DPNG": [0.46, 0.47, 0.48, 0.52, 0.55, 0.59],
                    "Commercial": [0.11, 0.15, 0.14, 0.14, 0.14, 0.15], "Industrial": [0.23, 0.27, 0.31, 0.37, 0.49, 0.58]},
                   stacked=True, colors=[C["blue"], C["green"], C["amber"], C["light_blue"]], fmt="0.00",
                   label_series=[0, 1], line={"Total": [2.22, 3.05, 3.51, 3.74, 4.23, 4.587]}, line_color=C["red"],
                   line_labels=True, gap=70)
    d.rect(s, L + 8.85, 4.1, 2.68, 2.4, C["green_tint"], name="Card")
    d.icon(s, L + 9.0, 4.25, 0.6, "chart-up")
    d.text(s, L + 9.0, 4.95, 2.4, 0.35, "YoY Volume Growth", 12, C["text"])
    d.text(s, L + 9.0, 5.3, 2.4, 0.3, "FY 25", 11, C["text"], bold=True)
    d.text(s, L + 9.0, 5.6, 2.4, 0.75, "09%", 36, C["dark_green"], bold=True)
    return s


volumes("Infrastructure and volumes")
volumes("Growth- Infrastructure and Volume")

# 13 ----------------------------------------------------------------- CAPEX
s = d.content_slide("CAPEX Outlay Plan")
d.stat_card(s, L, 1.3, 3.7, 1.15, "Rs. 9,300 Crore*", "Total capex till 31st March 2026", "rupee", "blue")
d.stat_card(s, L, 2.6, 3.7, 1.15, "~Rs.6,850 Crore", "Plan to Invest ~Rs.6,850 Crore in next 5 years", "chart-up", "green")
d.callout(s, L, 3.9, 3.7, 2.3, "", "Reinvesting Profits to enhance CGD network for benefit of Consumers primarily PNG Customers.\n\nTargeting addition of about 1 million and 500K DPNG customers in next 5 years in MGL and UEPL GAs, respectively.", "good")
d.heading(s, L + 4.0, 1.3, 7.5, "Total Capex (Rs. Cr)")
d.column_chart(s, L + 3.9, 1.55, 7.7, 4.6,
               ["Actual\nFY 2023-24", "Target\n(FY 2024-25)", "Target\n(FY 2025-26)", "Target\n(FY 2026-27)", "Target\n(FY 2027-28)", "Target\n(FY 2028-29)"],
               {"Total Capex": [847, 1454, 1346, 1453, 1297, 1297]}, point_colors=[C["blue"]] + [C["green"]] * 5)
footnote(s, "*The figures include MGL and UEPL Capex")

# 14 ----------------------------------------------------------------- 5-year Business Plan
s = d.content_slide("5-year Business Plan")
plan = [("home", "Connect ~12.5 lakhs additional households", C["green"]),
        ("fuel-pump", "New LNG stations at multiple locations to cater to the long-haul segment", C["blue"]),
        ("fuel-pump", "250+ additional CNG stations for meeting rising demand of CNG", C["green"]),
        ("chart-up", "Sales Volume upto ~ 6.80 mmscmd   (CAGR 10%)", C["blue"]),
        ("network", "Expansion of pipeline network of more than~1,000 kms", C["green"]),
        ("rupee", "Invest about Rs.6,000 crore to meet the expansion plan", C["blue"])]
pw = (W - 0.3) / 3
for i, (ic, t, col) in enumerate(plan):
    x, y = L + (i % 3) * (pw + 0.15), 1.4 + (i // 3) * 2.45
    d.rect(s, x, y, pw, 2.2, C["card"], name="Card")
    d.icon(s, x + 0.25, y + 0.25, 0.75, ic, col)
    d.text(s, x + 0.25, y + 1.15, pw - 0.5, 0.95, t, 15, C["ink"], bold=True)

# 15 ----------------------------------------------------------------- section
d.section_slide("O&M Infrastructure")

# 16 ----------------------------------------------------------------- O&M infrastructure
s = d.content_slide("O&M infrastructure")
half = (W - 0.3) / 2
for k, (head, x) in enumerate([("Emergency Control Stations", L), ("Security Control Room", L + half + 0.3)]):
    d.rect(s, x, 1.3, half, 0.38, C["green"], name="Header")
    d.text(s, x + 0.15, 1.3, half - 0.3, 0.38, head, 12, C["white"], bold=True, anchor="m")
line(s, L + half + 0.14, 1.3, 0.02, 5.3, C["grid"])
em = [("buildings", "MGL has divided its authorized GAs into 16 districts", ""),
      ("bolt", "24 x 7 Control room", "To receive complaints of damages and stoppage of gas"),
      ("bus", "29 Emergency Van", "Fully Equipped and operational round the clock")]
sec = [("CCTV SURVEILLANCE", "-Monitoring through cameras installed at strategic locations\n-Establishing two-way communication through D3D cameras"),
       ("IN VEHICLE MONITORING SYSTEM", "-Vehicle & trip monitoring of CTV\n-Driver behavior monitoring\n-Live information streams"),
       ("VEHICLE TRACKING SYSTEM", "-Monitor pipeline patrolling activity\n-Real time tracking & alerts")]
cw3 = (half - 0.2) / 3
for i, (ic, t, b) in enumerate(em):
    x = L + i * (cw3 + 0.1)
    d.rect(s, x, 1.8, cw3, 1.95, C["card"], name="Card")
    d.icon(s, x + 0.12, 1.92, 0.46, ic)
    d.text(s, x + 0.12, 2.48, cw3 - 0.2, 0.6, t, 9.5, C["ink"], bold=True)
    if b: d.text(s, x + 0.12, 3.08, cw3 - 0.2, 0.62, b, 8.5, C["text"])
for i, (t, b) in enumerate(sec):
    x = L + half + 0.3 + i * (cw3 + 0.1)
    d.rect(s, x, 1.8, cw3, 1.95, C["card"], name="Card")
    d.icon(s, x + 0.12, 1.92, 0.46, ["shield", "bus", "network"][i], C["blue"])
    d.text(s, x + 0.12, 2.48, cw3 - 0.2, 0.5, t, 9, C["ink"], bold=True)
    d.text(s, x + 0.12, 2.98, cw3 - 0.2, 0.75, b, 8, C["text"])
for i, (img, x) in enumerate([("emergency-control-room.jpg", L), ("emergency-van.jpg", L + half / 2 + 0.05),
                              ("security-room-1.jpg", L + half + 0.3), ("security-room-2.jpg", L + half * 1.5 + 0.35)]):
    d.picture(s, A / img, x, 3.92, half / 2 - 0.05, 2.65)

# 17 ----------------------------------------------------------------- ICCC photo
s = d.content_slide("O&M infrastructure")
d.picture(s, A / "iccc.jpg", L + 1.0, 1.25, W - 2.0, 5.0)
d.rect(s, L + 1.0, 6.0, 4.2, 0.45, C["green"], name="Caption")
d.text(s, L + 1.15, 6.0, 4.0, 0.45, "Integrated Command and Control Centre", 12, C["white"], bold=True, anchor="m")

# 18 ----------------------------------------------------------------- Vision
s = d.content_slide("MGL’s Vision and Long-term Strategy")
d.rect(s, L, 1.3, 4.2, 5.1, C["green_tint"], name="Card")
d.text(s, L + 0.3, 1.45, 3.6, 0.5, "LAKSHYA", 22, C["blue"], bold=True)
d.text(s, L + 0.3, 1.95, 3.6, 0.4, "Vision 2027", 16, C["dark_green"], bold=True)
d.text(s, L + 0.3, 2.4, 3.6, 0.35, "A diversified energy utility:", 12, C["ink"])
for i, (v, t) in enumerate([("1st", "quartile in employee engagement"), ("2x", "volume growth"), ("3x", "margins"),
                            ("1/4th", "revenue from new businesses"), ("in 5", "years")]):
    y = 2.9 + i * 0.66
    d.text(s, L + 0.3, y, 1.1, 0.55, v, 22, C["green"], bold=True, anchor="m")
    d.text(s, L + 1.45, y, 2.6, 0.55, t, 12, C["ink"], anchor="m")
vis = [("bolt", "Be a diversified Clean Energy company", "Diversify into new energy areas (e-mobility, biofuels)"),
       ("bulb", "Be a Digital Leader", "Digitize  across  the  value  chain (network planning to operations & maintenance to customer service)"),
       ("people", "Among India's top energy companies to work for", "Create and sustain an Employee friendly workplace")]
vw = (W - 4.5 - 0.3) / 3
for i, (ic, t, b) in enumerate(vis):
    x = L + 4.5 + i * (vw + 0.15)
    d.rect(s, x, 1.3, vw, 5.1, C["card"], name="Card")
    d.icon(s, x + (vw - 1.1) / 2, 1.6, 1.1, ic, [C["green"], C["blue"], C["amber"]][i])
    d.text(s, x + 0.2, 3.0, vw - 0.4, 0.9, t, 15, C["ink"], bold=True, align="c")
    d.text(s, x + 0.2, 4.05, vw - 0.4, 1.6, b, 11, C["text"], italic=True, align="c")

# 19 + 20 ------------------------------------------------------------ Business Expansion & Diversification
def expansion(rows, last_col, first_no, closing=None):
    s = d.content_slide("Business Expansion & Diversification")
    head = ["New business opportunity", "Description", "Adjacency with MGL's core", last_col]
    rh = 0.98 if len(rows) == 4 else 1.05
    d.table(s, L + 0.7, 1.3, W - 0.7, [head] + rows, col_w=[2.25, 2.95, 2.9, W - 0.7 - 8.1], row_h=rh, size=8.5)
    for i in range(len(rows)):
        y = 1.3 + rh * (i + 1) + (rh - 0.55) / 2
        d.rect(s, L, y, 0.55, 0.55, C["green"], shape=MSO_SHAPE.OVAL, name="Num")
        d.text(s, L, y, 0.55, 0.55, str(first_no + i), 14, C["white"], bold=True, align="c", anchor="m")
    if closing:
        d.callout(s, L, 1.3 + rh * (len(rows) + 1) + 0.2, W, 0.6, "", closing, "good")
    return s


expansion([
    ["CGD acquisition", "Augment Growth in the core sector", "Leveraging core competencies of MGL in the CGD sector",
     ("MGL auqired 100% stake in UEPL. Merger completed.", "good")],
    ["Long Haul/ Circular movement", "Cluster-based approach\nReplace existing diesel fleet with LNG for closed loop logistics",
     "Extension of MGL‘s business through participation in long haul logistics",
     "Rolling out LNG stations.\nSavroli (started in Aug-23) Launched Aurangabad in Oct-24. target to add 3 more LNG stations within 1 year period."],
    ["Compressed Bio Gas", "Set up Asia‘s biggest CBG plant at Mumbai, having capacity to process 1,000 TPD MSW\nBMC to provide land and feedstock",
     "Similar composition and energy potential to CNG; green renewable alternative for MGL",
     "Concession agreement to be signed with BMC.\nJV agreement with the partner to be signed after CA"],
    ["Green Hydrogen", "Develop use cases for Hydrogen for the mobility and Industrial segments through pilot projects",
     "Hydrogen to be blended in the CGD network.\nHydrogen to be dispensed in FCEV‘s deployed as part of project.",
     "Land for pilot project has been acquired at Rasayani, Patalganga.\nA pilot project of 0.5 MW for blending application has been planned to be implemented."],
], "Status", 1)
expansion([
    ["EV", "Exploring opportunities in the rapidly  expanding  e-mobility space",
     "Aligns with MGL‘s objective of being a diversified clean energy company.", "Invested in 3W EV manufacturing company"],
    ["Lithium Ion Battery", "Setting up a Joint Venture to manufacture Li-Ion battery cells",
     "Complements MGL‘s interest in the mobility sector.\nWill immensely help India indigenize the core of EV value chain.",
     "Land has been acquired from KIADB in Bangalore.\nMEP & CSE Design in final stages."],
    ["Renewable", "Exploring opportunities in the renewable energy space – Captive use as well as solar/ wind assets/ platforms.",
     "Aligns with MGL‘s objective of being a diversified clean energy company",
     "Invested in group captive Solar Power generation to meet demand at MGL‘s COCO CNG stations."],
], "Initiative Taken", 5,
    closing="By 2040, MGL aims to be a diversified clean energy company and is actively working towards developing capabilities in new energy segments viz. Hydrogen, CBG and Renewable Energy.")

# 21 ----------------------------------------------------------------- section
d.section_slide("CBG")

# 22 ----------------------------------------------------------------- CBG project updates
s = d.content_slide("Strategic CBG Initiatives & Project Updates – Mumbai Region")
panels = [
    ("1. Municipal Solid Waste (MSW) Project – Mankhurd & Deonar", "buildings", 4.75, [
        ("MOU signed with BMC", "on 08th June 2023 for setting up CBG plant"),
        ("Government Support:", "GR issued (28th May 2025) for land allocation at concessional rates to MGL."),
        ("Project Scope:", "Total 1000 TPD capacity plant, to be executed in two phases (500 TPD each)."),
        ("Plot at Phase-1 Mankhurd:", "Dismantling of the existing Mumbai Compost Structure is 95% complete by BMC."),
        ("Key Milestone:", "The structure of the Mumbai compost plant has been removed by BMC from the Mankhurd Bulldozer Garage site."),
        ("Key Milestone:", "Changes proposed in draft CA by BMC, under review. Concession Agreement (CA) with BMC expected to be signed in March 2026.")]),
    ("2. Cow Dung-Based Project – Aarey Colony", "sprout", 3.3, [
        ("Proposal Status:", "Submitted to Secretary (Animal Husbandry), GoM, for allotment of 6 acres of land."),
        ("Stakeholder Engagement:", "Multiple rounds of discussions held with the Commissioner (Dairy Dev.) and CEO (Aarey) to expedite the process."),
        "The CEO of Aarey Colony has informed MGL to have discussions with NDDB, which will be one of the partners for executing the CBG Project at Aarey Colony.",
        "Subsequently, MGL officials have met with NDDB officials for further action on the proposal."]),
    ("3. Municipal Solid Waste (MSW) Project – Davangere, Karnataka", "leaf", 3.18, [
        "Government of Karnataka(GoK) has allotted 10 acres of land to MGL for the establishment of CBG Plant at Avaragolla, Davangere, on 10th December 2025 for a period of 25 years, subject to a lease at the rate of ₹72,000 per acre per annum.",
        "MGL has requested City Corporation of Davangere (CCD), Karnataka, for reconsideration of the lease rent. Matter is being followed up."])]
x = L
for t, ic, w, items in panels:
    d.rect(s, x, 1.3, w, 5.3, C["card"], name="Card")
    d.icon(s, x + 0.15, 1.42, 0.5, ic)
    d.text(s, x + 0.78, 1.38, w - 0.9, 0.6, t, 11, C["dark_green"], bold=True, anchor="m")
    d.bullets(s, x + 0.18, 2.15, w - 0.33, 4.4, items, size=11)
    x += w + 0.15

d.closing_slide()
d.save(OUT)
print("wrote", OUT)
