# -*- coding: utf-8 -*-
"""manifest.py — the single source of truth that turns 16 folder converters into one
iterable pipeline. n8n (and assemble_payload.py) walk this list.

Per folder:
  folder            display name
  converter         the convert_*.py to run
  std_json          the standardized JSON it emits (assembler reads this)
  input_glob        SharePoint filename pattern; {MON}=Jun {MMM}=jun {YY}=26 (roll per month)
  sections          dashboard section key(s) in MONTHS[month] this folder feeds
                    ('CPI' is a SEPARATE stream — see stream field)
  stream            'month_payload' (goes into MONTHS[month]) | 'cpi' (separate CPI_DATA[month])
  validate_blocking True  -> if this folder's validate() fails, HALT the whole publish
                    False -> quarantine this folder, continue, mark its sections stale

NOTE on section mapping: the four folders standardized THIS session (Growth & Margin,
Executive Summary, CPI, Sojern) have confirmed section mappings. The other twelve are the
best inference from the dashboard's MONTHS[] structure and should be confirmed against each
converter's actual output the first time it runs (the assembler will flag any it can't map).
"""

MANIFEST = [
    # folder                         converter                    std_json                          input_glob                                      sections                                             stream           block
    {"folder": "Full P&L",                 "converter": "convert_full_pnl.py",      "std_json": "full_pnl_STANDARDIZED.json",     "input_glob": "P&L/Full P&L {MON}'{YY}.xlsx",              "sections": ["PL_CONSOL", "PL_BUS", "PLF"],                                    "stream": "month_payload", "validate_blocking": True},
    {"folder": "Consolidated P&L Summary", "converter": "convert_consol_pnl.py",    "std_json": "consol_pnl_STANDARDIZED.json",   "input_glob": "P&L/Consolidated P&L {MON}'{YY}.xlsx",      "sections": ["ESPL"],                                                  "stream": "month_payload", "validate_blocking": True},
    {"folder": "Monetization",             "converter": "convert_monetization.py",  "std_json": "monetization_STANDARDIZED.json", "input_glob": "Monetization/Monetization {MON}'{YY}.xlsx", "sections": ["MON", "MON_CONSOL", "MON_CHARTS"],                        "stream": "month_payload", "validate_blocking": True},
    {"folder": "Infra Cost",               "converter": "convert_infra.py",         "std_json": "infra_STANDARDIZED.json",        "input_glob": "Infra/Infra Cost {MON}'{YY}.xlsx",         "sections": ["INFRA_BU", "INFRA_DETAIL", "IC"],                              "stream": "month_payload", "validate_blocking": True},
    {"folder": "Travel Expense",           "converter": "convert_travel.py",        "std_json": "travel_STANDARDIZED.json",       "input_glob": "Travel/Travel Expense {MON}'{YY}.xlsx",    "sections": ["TRAV"],                                                  "stream": "month_payload", "validate_blocking": False},
    {"folder": "SG&A Expense",             "converter": "convert_sga.py",           "std_json": "sga_STANDARDIZED.json",          "input_glob": "SGA/SG&A {MON}'{YY}.xlsx",                 "sections": ["SGA"],                                                   "stream": "month_payload", "validate_blocking": False},
    {"folder": "GRR-NRR Ratios",           "converter": "convert_grr.py",           "std_json": "grr_STANDARDIZED.json",          "input_glob": "GRR/GRR NRR {MON}'{YY}.xlsx",              "sections": ["GRR_BRIDGE", "GRR_RATIOS", "GRR_LAST"],                  "stream": "month_payload", "validate_blocking": True},
    {"folder": "Revenue by BU",            "converter": "convert_revbu.py",         "std_json": "revbu_STANDARDIZED.json",        "input_glob": "Revenue/Revenue by BU {MON}'{YY}.xlsx",    "sections": ["REV_ROWS"],                                              "stream": "month_payload", "validate_blocking": True},
    {"folder": "HR Headcount",             "converter": "convert_hr.py",            "std_json": "hr_STANDARDIZED.json",           "input_glob": "HR/HR Headcount {MON}'{YY}.xlsx",          "sections": ["HR", "HR_PREV"],                                         "stream": "month_payload", "validate_blocking": False},
    {"folder": "Regional Revenue",         "converter": "convert_regional.py",      "std_json": "regional_STANDARDIZED.json",     "input_glob": "Regional/Regional Revenue {MON}'{YY}.xlsx","sections": ["REG_CONSOL", "REG_NORAM", "REG_EUROPE", "REG_APMEA", "REG_LATAM"], "stream": "month_payload", "validate_blocking": True},
    {"folder": "Top Accounts",             "converter": "convert_top_accounts.py",  "std_json": "top_accounts_STANDARDIZED.json", "input_glob": "TopAccounts/Top Accounts {MON}'{YY}.xlsx", "sections": ["PRODUCTS"],                                              "stream": "month_payload", "validate_blocking": False},
    {"folder": "Key KPIs",                 "converter": "convert_keykpis.py",       "std_json": "keykpis_STANDARDIZED.json",      "input_glob": "KPIs/Key KPIs {MON}'{YY}.xlsx",           "sections": ["CEO_DASH"],                                              "stream": "month_payload", "validate_blocking": True},
    {"folder": "Growth & Margin Snapshot", "converter": "convert_growth_margin.py", "std_json": "growth_margin_STANDARDIZED.json","input_glob": "Growth/Growth & Margin Snapshot {MON}'{YY}.xlsx", "sections": ["GM_SNAP"],                                        "stream": "month_payload", "validate_blocking": True},
    {"folder": "Executive Summary",        "converter": "convert_exec_summary.py",  "std_json": "exec_summary_STANDARDIZED.json", "input_glob": "ExecSummary/Executive Summary {MON}'{YY}.xlsx", "sections": ["ES"],                                              "stream": "month_payload", "validate_blocking": False},
    {"folder": "CPI Tracker",              "converter": "convert_cpi.py",           "std_json": "cpi_STANDARDIZED.json",          "input_glob": "CPI/CPI Tracker {MON}'{YY}.xlsx",          "sections": ["CPI"],                                                   "stream": "month_payload",           "validate_blocking": True},
    {"folder": "Sojern",                   "converter": "convert_sojern.py",        "std_json": "sojern_STANDARDIZED.json",       "input_glob": "Sojern/Sojern {MON}'{YY}.xlsx",           "sections": ["SOJERN"],                                                "stream": "month_payload", "validate_blocking": False},
]

# Sections that live in MONTHS[month] but are NOT produced by any folder converter
# (static / marketing was removed from the ticker) -> always carry forward from prior month.
CARRY_ALWAYS = ["MKT_EX", "MKT_SOJ", "PRODUCTS"]

# The full set of section keys a complete MONTHS[month] payload must contain (28).
REQUIRED_SECTIONS = [
    "CEO_DASH", "ES", "ESPL", "GM_SNAP", "GRR_BRIDGE", "GRR_LAST", "GRR_RATIOS",
    "HR", "HR_PREV", "INFRA_BU", "INFRA_DETAIL", "MKT_EX", "MKT_SOJ", "MON",
    "MON_CHARTS", "MON_CONSOL", "PL_BUS", "PL_CONSOL", "PRODUCTS", "REG_APMEA",
    "REG_CONSOL", "REG_EUROPE", "REG_LATAM", "REG_NORAM", "REV_ROWS", "SGA",
    "SOJERN", "TRAV",
]

if __name__ == "__main__":
    import json, sys
    if "--json" in sys.argv:            # emit JSON for n8n
        print(json.dumps({"manifest": MANIFEST, "carry_always": CARRY_ALWAYS,
                          "required_sections": REQUIRED_SECTIONS}, indent=2))
    else:
        covered = {s for f in MANIFEST for s in f["sections"] if f["stream"] == "month_payload"}
        covered |= set(CARRY_ALWAYS)
        missing = [s for s in REQUIRED_SECTIONS if s not in covered]
        print(f"{len(MANIFEST)} folders; month_payload sections covered: {len(covered & set(REQUIRED_SECTIONS))}/{len(REQUIRED_SECTIONS)}")
        print("sections with no folder (carry-forward):", missing or "none")
