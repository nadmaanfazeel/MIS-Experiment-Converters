# -*- coding: utf-8 -*-
"""assemble_payload.py — MAP -> ASSEMBLE.

Turns the standardized JSON produced by the folder converters into a complete
MONTHS[month] payload (28 sections) that seed_supabase.mjs upserts to Supabase.

How it works:
  1. Reads manifest.py to know which folder feeds which section.
  2. For each section that HAS a mapper and whose converter output is present,
     maps the standardized JSON -> the dashboard section shape (fresh data).
  3. Every other required section is CARRIED FORWARD from the previous month's
     payload, and reported loudly, so nothing stale is published silently.
  4. Emits {month: payload} in the exact shape seed_supabase.mjs consumes, plus a
     report of mapped / carried-forward / missing sections for the publish gate.

Reference mappers implemented now: GM_SNAP, ES  (the two folders standardized this
session whose target shapes are confirmed). Add the others by writing one function
and registering it in MAPPERS — each is ~10 lines. Until then they carry forward.

Usage:
  # assemble July from July's converter outputs, carrying forward from June:
  python3 assemble_payload.py --month jul --std-dir ./std_jul \
      --prev months_all.json --prev-month jun --out month_jul.json

  # self-test: rebuild June from this session's outputs and diff vs the real payload:
  python3 assemble_payload.py --month jun --std-dir . \
      --prev months_all.json --prev-month jun --out /tmp/month_jun.json --selftest
"""
import json, os, sys, argparse
from manifest import MANIFEST, REQUIRED_SECTIONS, CARRY_ALWAYS
import section_mappers

def load(path):
    with open(path) as f: return json.load(f)

# ------------------------------------------------------------------ mappers
def map_gm_snap(std, prev=None):
    """growth_margin_STANDARDIZED.json -> GM_SNAP  [{t,c,v,s,x?}]"""
    out = []
    for c in sorted(std["dashboard_preview"], key=lambda x: x.get("order", 99)):
        card = {"t": c["t"], "c": c["color"], "v": c["v"], "s": c["s"]}
        if c.get("x"):
            card["x"] = c["x"]
        out.append(card)
    return out

def map_es(std, prev=None):
    """exec_summary_STANDARDIZED.json -> ES {consolidated:{stats,narr}, bus:{...}}"""
    ep = std["es_preview"]
    bus, n = {}, 0
    for key, b in ep["bus"].items():
        n += 1
        prods = [{"name": p["name"], "metric": p.get("metric", ""), "points": p["points"]}
                 for p in sorted(b["products"], key=lambda x: x.get("order", 99))]
        bus[key] = {"num": n, "name": b.get("name", key), "desc": b.get("teaser", ""),
                    "headline": b.get("headline", ""), "teaser": b.get("teaser", ""),
                    "products": prods}
    return {"consolidated": {"stats": ep["consolidated"]["stats"],
                             "narr": ep["consolidated"]["narr"]},
            "bus": bus}

def map_ic(std, prev=None):
    """infra_STANDARDIZED.json -> IC node tree (dashboard IC_JUN shape)."""
    nodes = std["standardized"]["nodes"]
    return {nid: {k: nd[k] for k in ("name", "parent", "col", "strips", "tables")}
            for nid, nd in nodes.items()}

def _cpi_money(v):
    return "-" if (v is None or v == 0) else "$" + format(int(round(v)), ",")

def _cpi_pct(v, tok=None):
    if tok:
        return "–" if tok == "na" else tok            # 'New' / 'Pending' / '–'
    return "–" if v is None else f"{v * 100:.1f}%"

def map_cpi(std, prev=None):
    """cpi_STANDARDIZED.json -> CPI_DATA[month] cards [{title,cols,rows,total}]."""
    meta = std["standardized"].get("meta", {})
    cur_tok, app_tok = meta.get("cur_month_token", ""), meta.get("applied_month_token", "")
    cards = []
    for g in std["standardized"]["groups"]:
        cols = g["cols"]
        rows = []
        for a in g["accounts"]:
            row = []
            seen_month = 0
            for col in cols:
                cl = col.lower()
                if "account name" in cl: row.append(a.get("account") or "")
                elif cl == "product":    row.append(a.get("product") or "")
                elif cl == "am":         row.append(a.get("am") or "")
                elif cl == "group":      row.append(a.get("group") or "")
                elif cl == "due":        row.append(a.get("due") or "")
                elif "revenue" in cl:    row.append("–" if a.get("revenue") is None else "$" + format(int(round(a["revenue"])), ","))
                elif "contract cpi" in cl: row.append(_cpi_pct(a.get("contract_cpi")))
                elif "impact" in cl:     row.append(_cpi_money(a.get("cpi_impact")))
                elif cl == "status":     row.append(a.get("status") or "")
                elif "comment" in cl:    row.append(a.get("comments") or "")
                elif "cpi" in cl and "%" in cl:                  # a month CPI column
                    if cur_tok and cur_tok.lower() in cl:
                        row.append(_cpi_pct(a.get("cur_month_cpi"), a.get("cur_month_cpi_token")))
                    elif app_tok and app_tok.lower() in cl:
                        row.append(_cpi_pct(a.get("applied_cpi"), a.get("applied_cpi_token")))
                    else:                                        # fallback by position
                        row.append(_cpi_pct(a.get("applied_cpi") if seen_month else a.get("cur_month_cpi")))
                    seen_month += 1
                else: row.append("")
            rows.append(row)
        tr = g.get("total_row") or {}
        total = []
        for col in cols:
            cl = col.lower()
            if "account name" in cl: total.append("TOTAL")
            elif "revenue" in cl:    total.append("$" + format(int(round(tr.get("revenue") or 0)), ","))
            elif "impact" in cl:     total.append("$" + format(int(round(tr.get("cpi_impact") or 0)), ","))
            else: total.append("")
        cards.append({"title": g["group"], "cols": cols, "rows": rows, "total": total})
    return cards

# section_key -> (std_json filename, mapper fn)
MAPPERS = {
    "GM_SNAP": ("growth_margin_STANDARDIZED.json", map_gm_snap),
    "ES":      ("exec_summary_STANDARDIZED.json",  map_es),
    "IC":      ("infra_STANDARDIZED.json",         map_ic),   # infra drill-down tree
    "CPI":     ("cpi_STANDARDIZED.json",           map_cpi),  # CPI tracker cards
}
# extra payload sections that aren't required every month (IC/CPI absent months are fine)
EXTRA_SECTIONS = ["IC", "CPI", "PLF"]   # PLF = product P&L drill-down (incl p-unorez)

MAPPERS.update(section_mappers.MAPPERS)   # SGA, GRR_*, TRAV, REG_*, REV_ROWS, HR
DERIVED = ["GRR_LAST", "HR_PREV"]         # prior-month derivations, built in main()

# ------------------------------------------------------------------ assemble
def assemble(month, std_dir, prev_payload):
    fresh, carried, missing = {}, [], []
    payload = {}
    for sec in REQUIRED_SECTIONS:
        mp = MAPPERS.get(sec)
        std_path = os.path.join(std_dir, mp[0]) if mp else None
        if mp and std_path and os.path.exists(std_path):
            try:
                payload[sec] = mp[1](load(std_path), prev_payload.get(sec))
                fresh[sec] = mp[0]
                continue
            except Exception as e:
                print(f"  ! mapper for {sec} failed: {e}  -> carrying forward")
        if prev_payload and sec in prev_payload:
            payload[sec] = prev_payload[sec]
            carried.append(sec)
        else:
            missing.append(sec)
    # optional extras (IC / CPI): map if the converter output is present; else carry if prior had it
    for sec in EXTRA_SECTIONS:
        mp = MAPPERS.get(sec); std_path = os.path.join(std_dir, mp[0]) if mp else None
        if mp and std_path and os.path.exists(std_path):
            try:
                payload[sec] = mp[1](load(std_path), prev_payload.get(sec)); fresh[sec] = mp[0]; continue
            except Exception as e:
                print(f"  ! mapper for {sec} failed: {e}")
        if prev_payload and sec in prev_payload:
            payload[sec] = prev_payload[sec]; carried.append(sec)
    return payload, {"fresh": fresh, "carried": carried, "missing": missing}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True)
    ap.add_argument("--std-dir", default=".")
    ap.add_argument("--prev", help="months_all.json (carry-forward source)")
    ap.add_argument("--prev-month", default=None)
    ap.add_argument("--out", required=True)
    ap.add_argument("--selftest", action="store_true")
    a = ap.parse_args()

    prev_all = load(a.prev) if a.prev else {}
    prev_payload = prev_all.get(a.prev_month or a.month, {})

    payload, rep = assemble(a.month, a.std_dir, prev_payload)
    # prior-month derivations
    if "GRR_RATIOS" in prev_payload:
        payload["GRR_LAST"] = [{"l": r["l"] + " Last Month", "type": r["type"], "c": r["c"]}
                               for r in prev_payload["GRR_RATIOS"]]
        rep["fresh"]["GRR_LAST"] = "(derived: prev GRR_RATIOS)"
    if "HR" in prev_payload:
        payload["HR_PREV"] = prev_payload["HR"]; rep["fresh"]["HR_PREV"] = "(derived: prev HR)"
    for d in ("GRR_LAST","HR_PREV"):
        if d in rep["carried"]: rep["carried"].remove(d)


    print(f"\n== assemble {a.month} ==")
    print(f"  fresh-mapped ({len(rep['fresh'])}): {', '.join(sorted(rep['fresh'])) or '—'}")
    print(f"  carried-forward ({len(rep['carried'])}): {', '.join(rep['carried']) or '—'}")
    print(f"  MISSING ({len(rep['missing'])}): {', '.join(rep['missing']) or '—'}")
    clean = not rep["missing"] and not rep["carried"]
    print(f"  publish-clean (all fresh, nothing carried/missing): {clean}")

    if a.selftest and prev_payload:
        print("\n  self-test vs real payload:")
        for sec in sorted(rep["fresh"]):
            same = json.dumps(payload[sec], sort_keys=True) == json.dumps(prev_payload.get(sec), sort_keys=True)
            print(f"    {sec}: {'EXACT match' if same else 'differs (expected where prose is hand-polished)'}")

    with open(a.out, "w") as f:
        json.dump({a.month: payload}, f, ensure_ascii=False)
    print(f"\n  wrote {a.out}  ({len(payload)} sections) — feed this to seed_supabase.mjs")
    # exit non-zero if anything missing, so n8n can gate on it
    sys.exit(1 if rep["missing"] else 0)

if __name__ == "__main__":
    main()
