"""Derived numbers for the Japan-India energy / carbon research. Every input cites a saved source.
Run: python calculations.py   (prints each result with its formula)"""

# ---- inputs (source file number in brackets) ----
EF_NG_KG_PER_TJ = 56_100        # IPCC 2006 default CO2 EF, natural gas [23]
NCV_NG_TJ_PER_GG = 48.0         # IPCC 2006 NCV natural gas [23]
EF_NG_KG_PER_SCF = 0.05444      # US EPA EF Hub 2025, natural gas / CNG [27]
SCF_PER_M3 = 35.3147            # unit conversion (1 m3 = 35.3147 ft3)
GWP_CH4 = 28                    # IPCC AR5 GWP100, non-fossil methane [26]
DOC_FOOD, DOCF, F_CH4 = 0.15, 0.5, 0.5   # IPCC 2006 waste defaults [24][25]
MCF_MANAGED, MCF_UNMANAGED_DEEP = 1.0, 0.8  # IPCC 2006 [25]
KANJUR_CH4_KG_H = 3_300         # Carbon Mapper, 3 Apr 2026 overpass [18]
MUMBAI_WASTE_TPD = 6_500        # [21]
DEONAR_SHARE = 0.10             # ~10% to Deonar, rest Kanjurmarg [31]
DEONAR_CBG_WASTE_TPD = 1_000    # MGL Deonar project wet waste [20]
DEONAR_CBG_T_PER_DAY = 2 * 18   # two plants x 18 t CBG/day [20]
MGL_350_TPD = 350               # MGL-BMC Mankhurd Phase I [19]
CAP_ASSIST_RS_PER_TPD_CBG = 2.0e7  # GOBARdhan: Rs 20 million per TPD of CBG capacity [09]
MGL_CNG_MSCM_Q1, MGL_DPNG_MSCM_Q1, DAYS_Q1 = 318.09, 56.67, 91   # MGL Q1 FY27 [22]
BLEND = {"FY27": 0.03, "FY28": 0.04, "FY29+": 0.05}  # CBG obligation [09]
INDORE_TPD, INDORE_TCO2E = 550, 70_000                # EverEnviro Indore benchmark [32]

out = []
def r(k, v, f): out.append((k, v, f)); return v

co2_per_kg_cbg = r("CO2 avoided per kg CBG replacing natural gas (kg)", EF_NG_KG_PER_TJ * NCV_NG_TJ_PER_GG / 1e6,
                   "56,100 kg/TJ x 48.0 TJ/Gg / 1e6")
co2_per_scm = r("CO2 per SCM natural gas burned (kg)", EF_NG_KG_PER_SCF * SCF_PER_M3, "0.05444 kg/scf x 35.3147 scf/m3")

# landfill: IPCC default (lifetime) potential per tonne of wet food waste
ch4_t_per_t = DOC_FOOD * DOCF * F_CH4 * 16 / 12
r("IPCC lifetime CH4 per t wet food waste, managed site (t CH4)", ch4_t_per_t * MCF_MANAGED, "MCF 1.0 x DOC 0.15 x DOCf 0.5 x F 0.5 x 16/12")
hi = r("  = tCO2e per t wet waste, managed site", ch4_t_per_t * MCF_MANAGED * GWP_CH4, "x GWP 28")
lo = r("  = tCO2e per t wet waste, unmanaged deep site", ch4_t_per_t * MCF_UNMANAGED_DEEP * GWP_CH4, "MCF 0.8 x ... x GWP 28")

# landfill: observed at Kanjurmarg (snapshot)
kanjur_t_yr = MUMBAI_WASTE_TPD * (1 - DEONAR_SHARE) * 365
kanjur_co2e = r("Kanjurmarg CH4 as tCO2e/yr if the Apr-2026 rate held all year", KANJUR_CH4_KG_H * 8760 / 1000 * GWP_CH4,
                "3,300 kg/h x 8,760 h / 1000 x 28")
obs = r("  = tCO2e per tonne of waste received now (observed, indicative)", kanjur_co2e / kanjur_t_yr,
        "/ (6,500 t/d x 90% x 365)")
r("Indore benchmark: tCO2e avoided per tonne organic waste", INDORE_TCO2E / (INDORE_TPD * 365), "70,000 / (550 x 365)")

# CBG route
cbg_kg_per_t = r("CBG yield, kg per tonne wet waste (MGL Deonar design)", DEONAR_CBG_T_PER_DAY * 1000 / DEONAR_CBG_WASTE_TPD, "36 t / 1,000 t")
fuel = r("  fossil CO2 displaced per tonne waste via CBG (t)", cbg_kg_per_t * co2_per_kg_cbg / 1000, "36 kg x 2.69")
r("Gross saving per tonne diverted, low (observed + fuel)", obs + fuel, "observed methane + fuel displacement")
r("Gross saving per tonne diverted, high (IPCC managed + fuel)", hi + fuel, "IPCC methane + fuel displacement")
for name, tpd in (("MGL Mankhurd 350 TPD", MGL_350_TPD), ("MGL Deonar 1,000 TPD", DEONAR_CBG_WASTE_TPD)):
    t = tpd * 365
    r(f"{name}: tCO2e/yr, low", t * (obs + fuel), "t/yr x low per-tonne")
    r(f"{name}: tCO2e/yr, high", t * (hi + fuel), "t/yr x high per-tonne")
    r(f"{name}: CBG t/day", tpd * cbg_kg_per_t / 1000, "TPD x 36 kg/t")
    r(f"{name}: GOBARdhan capital assistance (Rs crore)", tpd * cbg_kg_per_t / 1000 * CAP_ASSIST_RS_PER_TPD_CBG / 1e7, "CBG TPD x Rs 2 cr")

# PNG blending
dpng = r("MGL domestic PNG, MMSCMD (Q1 FY27)", MGL_DPNG_MSCM_Q1 / DAYS_Q1, "56.67 MSCM / 91 days")
cng = r("MGL CNG, MMSCMD (Q1 FY27)", MGL_CNG_MSCM_Q1 / DAYS_Q1, "318.09 MSCM / 91 days")
dpng_co2 = r("CO2 from MGL domestic PNG, t/yr", dpng * 1e6 * 365 * co2_per_scm / 1000, "MMSCMD x 365 x 1.92 kg")
r("  saving per 1% CBG blended in domestic PNG, tCO2/yr", dpng_co2 * 0.01, "1% of above")
for fy, b in BLEND.items():
    r(f"  domestic PNG saving at {b:.0%} ({fy}), tCO2/yr", dpng_co2 * b, f"{b:.0%} of PNG CO2")
    r(f"  CNG + domestic PNG saving at {b:.0%} ({fy}), tCO2/yr", (dpng + cng) * 1e6 * 365 * co2_per_scm / 1000 * b, "same on CNG + PNG")
    r(f"  CBG volume needed for CNG + PNG at {b:.0%}, SCM/day", (dpng + cng) * 1e6 * b, "volume x blend")

if __name__ == "__main__":
    for k, v, f in out:
        print(f"{k:72s} {v:>14,.3f}   [{f}]")
