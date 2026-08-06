# -*- coding: utf-8 -*-
"""Mappers: folder converter output -> dashboard payload sections.
mapper(std, prev) -> section value. `prev` = same section from previous month's payload,
used to carry hand-authored text/structure while numbers refresh from `std`.
GRR_LAST and HR_PREV are prior-month derivations handled in the assembler."""
import json, math

def rh(x):                        # round half up, symmetric
    if x is None: return None
    return int(math.floor(x + 0.5))   # JS Math.round semantics (half up toward +inf)
def _c0(v):  return "–" if v is None else format(rh(v), ",")            # "7,970" no $
def _usdL(v):                                                            # "($40)" / "$3,771"
    if v is None: return "–"
    n = rh(v); return ("($" + format(-n, ",") + ")") if n < 0 else ("$" + format(n, ","))
def _usdR(v):                                                            # "$(13)" / "$219,882"
    if v is None: return "–"
    n = rh(v); return ("$(" + format(-n, ",") + ")") if n < 0 else ("$" + format(n, ","))
def _pct(v): return "n/m" if v is None else f"{rh(v*100)}%"
def _i100(v): return None if v is None else rh(v * 100)

# ---------- GRR (transpose) ----------
_BRIDGE = [("YTD FY 25-26","ytd_pfy","tot"),("Churn","churn","neg"),("Downsell","downsell","neg"),
           ("Upsell","upsell","pos"),("New Revenue","new_revenue","num"),
           ("Exceptional Items","exceptional","exc"),("YTD FY 26-27","ytd_cur","tot2")]
def map_grr_bridge(std, prev=None):
    e = std["entities"]; return [{"l": l, "k": k, "c": [_usdL(x.get(f)) for x in e]} for l,f,k in _BRIDGE]
def map_grr_ratios(std, prev=None):
    e = std["entities"]
    return [{"l":"GRR","type":"grr","c":[_i100(x.get("grr")) for x in e]},
            {"l":"NRR","type":"nrr","c":[_i100(x.get("nrr")) for x in e]}]

# ---------- Revenue by BU (template-carry: keep dashboard rows, refresh #s by label) ----------
def map_rev_rows(std, prev=None):
    out=[]
    for r in std["rows"]:
        k={"bu":"bu","sub_bu":"sub","total":"total","adjustment":"total"}.get(r["role"],"leaf")
        row={"l":r["node"],"k":k,"v":[_c0(r.get("ytd_cur")),_c0(r.get("ytd_pfy")),_pct(r.get("yoy_growth"))]}
        if r.get("level"): row["i"]=r["level"]
        out.append(row)
    return out

# ---------- Regional (one positional array per region; empty -> '–') ----------
_REGMAP={"REG_CONSOL":"Consol","REG_NORAM":"NORAM","REG_EUROPE":"EU","REG_APMEA":"APMEA","REG_LATAM":"LATAM"}
def _regcur(v): return "–" if (v is None or v == 0) else _usdR(v)
def make_region_mapper(key):
    reg=_REGMAP[key]
    def _m(std, prev=None):
        out=[]
        for r in std["rows"]:
            d=r["regions"].get(reg,{})
            pfy,cur,g=d.get("ytd_pfy"),d.get("ytd_cur"),d.get("growth")
            empty=(pfy in (None,0) and cur in (None,0))
            out.append([_regcur(pfy),_regcur(cur),("" if empty else _pct(g))])
        return out
    return _m

# ---------- SG&A (consol 5-col in $'000, legal 3-col + nature; notes carried) ----------
def _sga_rows(rows, total_label):
    C=["YTD FY 25-26 Act","YTD FY 26-27 Act","YTD FY 26-27 Bud","YTD Bud Var","YoY Var"]
    out=[]; sc=lambda v:(None if v is None else v/1000.0)
    for r in rows:
        li=r.get("line_item") or r.get("item") or ""
        k="total" if total_label.lower() in li.lower() else ("grp" if r.get("is_parent") else "leaf")
        out.append({"l":li,"k":k,"c":[_usdR(sc(r.get(C[0]))),_usdR(sc(r.get(C[1]))),_usdR(sc(r.get(C[2]))),
                                       _pct(r.get(C[3])),_pct(r.get(C[4]))]})
    tot=next((r for r in rows if total_label.lower() in (r.get("line_item") or r.get("item") or "").lower()),None)
    return out,tot
def map_sga(std, prev=None):
    prev=prev or {}; sc=lambda v:(None if v is None else v/1000.0)
    crows,ctot=_sga_rows(std.get("sga",{}).get("rows",[]),"SG&A Costs")
    consol=dict(prev.get("consol") or {}); consol["rows"]=crows
    if ctot:
        consol["ytd"]=rh((ctot.get("YTD FY 26-27 Act") or 0)/1000.0); consol["py"]=rh((ctot.get("YTD FY 25-26 Act") or 0)/1000.0)
        consol["yoyVar"]=_pct(ctot.get("YoY Var")); consol["budVar"]=_pct(ctot.get("YTD Bud Var"))
    lrows=[]
    for r in std.get("legal_prof",{}).get("rows",[]):
        item=r.get("item") or ""
        k="total" if "total professional" in item.lower() else ("grp" if r.get("is_parent") else "leaf")
        lrows.append({"l":item,"k":k,"nat":r.get("nature",""),
                      "c":[_usdR(sc(r.get("YTD FY 25-26 Act"))),_usdR(sc(r.get("YTD FY 26-27 Act"))),_pct(r.get("YoY Var"))]})
    legal=dict(prev.get("legal") or {}); legal["rows"]=lrows
    ltot=next((r for r in std.get("legal_prof",{}).get("rows",[]) if "total professional" in (r.get("item") or "").lower()),None)
    if ltot:
        legal["ytd"]=rh((ltot.get("YTD FY 26-27 Act") or 0)/1000.0); legal["py"]=rh((ltot.get("YTD FY 25-26 Act") or 0)/1000.0); legal["yoyVar"]=_pct(ltot.get("YoY Var"))
    return {"consol":consol,"legal":legal}

# ---------- Travel (full USD; ($x) negatives; labels + notes carried; cols roll) ----------
def map_trav(std, prev=None):
    prev=prev or {}
    lt=std.get("leader_table",{}); months=lt.get("months",[])
    cols=["Leader","Approved","FY27 YTD Bud"]+[m.replace("\u0027","\u2019")+" Expense" for m in months]+["FY27 YTD Expense","Variance"]
    prev_labels=[r["l"] for r in (prev.get("consol") or {}).get("rows",[])]
    rows=[]
    for i,r in enumerate(lt.get("rows",[])):
        c=[_usdL(r.get("approved")),_usdL(r.get("ytd_bud"))]+[_usdL(r["monthly"].get(m)) for m in months]+[_usdL(r.get("ytd_expense")),_usdL(r.get("variance"))]
        lbl=prev_labels[i] if i<len(prev_labels) else r["leader"]     # carry cleaned labels positionally
        rows.append({"l":lbl,"c":c})
    tot=next((r for r in lt.get("rows",[]) if str(r.get("leader","")).lower()=="total"),None)
    consol=dict(prev.get("consol") or {}); consol["cols"]=cols; consol["rows"]=rows
    if tot: consol["ytd"]=rh(tot.get("ytd_expense") or 0); consol["bud"]=rh(tot.get("ytd_bud") or 0)
    sj=std.get("sojern_dept",{}); ents=sj.get("entities",[])
    sojern=dict(prev.get("sojern") or {}); sojern["cols"]=["Dept"]+ents
    sojern["rows"]=[{"l":r["dept"],"c":[_c0(r.get(e)) for e in ents]} for r in sj.get("rows",[])]
    return {"consol":consol,"sojern":sojern}

# ---------- HR (7 tabs; explicit per-view col-0 header) ----------
_HRMAP=[("Total Headcount By Region","region","Headcount by Region","Region / Category"),
        ("SoHo","soho","SoHo","Department"),("Sojern","sojern","Sojern","Department"),
        ("RG","rg","RG","Department"),
        ("Combine_SoHo_Sojern_RG","combine","Combined \u00b7 SoHo \u00b7 Sojern \u00b7 RG","Entity / Category"),
        ("Total Headcount Deepak Kapoor","deepak","Deepak Kapoor Org","Department"),
        ("Total Headcount Finance","finance","Finance","Department")]
def _hrc(v): return "" if v is None else format(int(round(v)),",")
def _hr_view(tb,name,col0):
    rows=[]
    if tb["type"]=="flat":
        for r in tb["rows"]: rows.append([r["label"],_hrc(r["count"]),"total" if r.get("is_total") else ""])
    else:
        for r in tb["rows"]:
            tag="total" if r.get("is_total") else ("grp" if r["level"]==0 else "")
            rows.append([r["node"],_hrc(r["count"]),tag])
    return {"name":name,"cols":[col0,tb.get("month","")],"rows":rows}
def map_hr(std, prev=None):
    T=std["tables"]; out={}; prev=prev or {}
    for tabname,key,name,col0 in _HRMAP:
        tb=T.get(tabname) or T.get(tabname.strip())
        if tb: out[key]=_hr_view(tb,name,col0)
        elif key in prev: out[key]=prev[key]
    return out


# ---------- Top Accounts -> PRODUCTS (rows/totals refreshed; card copy carried) ----------
_TA2KEY={"TravelBI":"travelbi","HospiBI":"hospibi","Rez+UNO":"rezuno","Enterprise Connectivity":"ec",
         "SoHo":"soho","Sojern":"sojern","Sojern - Destination":"sojern-dest","Sojern - Corporate":"sojern-corp"}
def _acct(n): return (n or "").rstrip(". ").strip()
def _tarow(a):
    return {"n":a.get("rank"),"acct":_acct(a.get("account")),
            "vals":[rh(a.get("fy_pfy_act")),rh(a.get("ytd_pfy")),rh(a.get("ytd_cur")),
                    rh(a.get("ytd_bud")),rh((a.get("yoy_pct") or 0)*100),rh((a.get("bud_var_pct") or 0)*100)],
            "cmt":a.get("comments","")}
def map_products(std, prev=None):
    prev=prev or {}; out=dict(prev)   # start from carried presentation, refresh per tab
    for tab in std["tabs"]:
        key=_TA2KEY.get(tab.get("tab"))
        if not key: continue
        card=dict(prev.get(key) or {})
        accts=tab.get("accounts")
        if accts is not None:
            card["rows"]=[_tarow(a) for a in accts]
            summ=tab.get("summary") or []
            if summ:
                card["totals"]=[{"label":x.get("label"),"kind":x.get("kind","sub"),
                                 "vals":[rh(v) if isinstance(v,(int,float)) else v for v in x.get("vals",[])]} for x in summ]
        out[key]=card
    return out


# ---------- Sojern (16 sub-tables: 12 "groups" mapped, 4 curated carried) ----------
_SJ={"grossreg":"Gross Revenue by Region","netreg":"Net Revenue by Region","gmreg":"Revenue and GM by Region",
     "yoy":"Revenue YoY Growth Rates","retvert":"Annual Retention by Vertical","retreg":"Annual Retention by Region",
     "waterfall":"Property Account Waterfall","proprev":"Property Revenue-Gross Profit","prpa":"Property Revenue Per Account",
     "momret":"Monthly Retention Property","destrev":"Destinations Rev-Gross Margin","corprev":"Corporate Revenue-Gross Profit"}
_SJ_CARRY=["alloc","ar","recon","ltvcac"]   # curated/messy -> carry hand-built version
def _sjfmt(v,col,unit,label=""):
    if v is None: return "–"
    if isinstance(v,str): return v
    if "%" in col or "%" in (label or "") or unit in ("percent","ratio/percent"):
        n=v*100
        return f"({abs(n):.1f}%)" if n<0 else f"{n:.1f}%"
    if unit=="count":
        n=rh(v); return f"({format(-n,',')})" if n<0 else format(n,",")
    n=rh(v)
    return f"(${format(-n,',')})" if n<0 else f"${format(n,',')}"
def map_sojern(std, prev=None):
    prev=prev or {}
    S=std.get("standardized",std); tabs={t["tab"]:t for t in S["tabs"]}
    out=dict(prev)
    for key,tabname in _SJ.items():
        tab=tabs.get(tabname); pv=prev.get(key,{})
        if not tab or not tab.get("blocks"): continue
        blocks=tab["blocks"]; tabunit=tab.get("unit")
        has_group=any(r.get("group") for blk in blocks for r in blk["rows"])
        groups=[]
        if has_group:
            order=[]; buckets={}; cols_for={}
            for blk in blocks:
                u=blk.get("unit") or tabunit
                sec=(blk.get("section") or "").lower()
                if u=="count" and ("revenue" in sec or "per account" in sec): u="usd"
                for r in blk["rows"]:
                    g=r.get("group") or (blk.get("section") or "")
                    if g not in buckets: buckets[g]=[]; order.append(g)
                    buckets[g].append([r["label"]]+[_sjfmt(r["cells"].get(c),c,u,r["label"]) for c in blk["columns"]])
            groups=[{"g":g,"rows":buckets[g]} for g in order]
        else:
            for blk in blocks:
                u=blk.get("unit") or tabunit
                sec=(blk.get("section") or "").lower()
                if u=="count" and ("revenue" in sec or "per account" in sec): u="usd"
                rows=[[r["label"]]+[_sjfmt(r["cells"].get(c),c,u,r["label"]) for c in blk["columns"]] for r in blk["rows"]]
                groups.append({"g":blk.get("section") or "","rows":rows})
        cols=pv.get("cols") or (["($000s)"]+blocks[0]["columns"])
        out[key]={"note":pv.get("note",""),"cols":cols,"groups":groups}
    return out


# ---------- Key KPIs -> CEO_DASH (per-unit formatting; section + 40%-rule rows) ----------
def _kfmt(v,unit):
    if v is None: return "n/a"
    if unit=="percent": return f"{rh(v*100)}%"
    if unit in ("currency","count"): return format(rh(v),",")
    if unit=="multiple": return f"{v:.2f}x"
    if unit in ("months","ratio"): return f"{v:.2f}"
    return str(v)
def map_ceo_dash(std, prev=None):
    bus=std["bus"]; out=[]; seen=set()
    for m in std["metrics"]:
        sec=m.get("section")
        if sec and sec not in seen: out.append({"section":sec}); seen.add(sec)
        v=[_kfmt(m["values"].get(b),m["unit"]) for b in bus]
        label=m["metric"].replace(" - "," \u2014 ")     # normalise hyphen -> em-dash (Monetization)
        row={"l":label,"v":v}
        if sec=="40% Rule Check": row["l"]="\u2013 "+m["metric"]; row["rule"]=True
        out.append(row)
    return out


# ---------- Consolidated P&L -> ESPL (4 BUs x Revenue/Cost/EBITDA/EBITDA%) ----------
_ESPL_ROLES=["prior_act","cur_act","cur_bud","mom_bud_var","ytd_act","ytd_bud","yoy_bud_var","pfy_ytd_act","yoy_var","py_month_act","yoy_month_var"]
_ESPL_PCTPOS={3,6,8,10}
_ESPL_K={"Revenue":"rev","Net revenue":"rev","Gross Revenue":"rev","Cost":"cost","EBITDA":"eb","EBITDA %":"pct"}
def map_espl(std, prev=None):
    order=["DaaS","Distribution","Martech","Consolidated"]; bybu={}
    for r in std: bybu.setdefault(r["bu"],[]).append(r)
    out=[]
    for bu in order:
        rows=[]
        for rec in bybu.get(bu,[]):
            k=_ESPL_K.get(rec["metric"],"")
            c=[]
            for i,role in enumerate(_ESPL_ROLES):
                v=rec.get(role)
                if k=="pct":
                    c.append("" if i in _ESPL_PCTPOS else ("n/m" if v is None else f"{rh(v*100)}%"))
                else:
                    if i in _ESPL_PCTPOS: c.append("" if v is None else f"{rh(v*100)}%")
                    else: c.append("–" if v is None else format(rh(v),","))
            rows.append({"l":rec["metric"],"k":k,"c":c})
        card={"bu":bu,"rows":rows}
        if bu=="Consolidated": card={"bu":bu,"strong":True,"rows":rows}
        out.append(card)
    return out
def map_carry(std, prev=None): return prev   # PL_BUS / PL_CONSOL: hand-authored summary cards


# ---------- Full P&L -> PLF product tables (P1 refreshed from converter; P2 carried) ----------
_PROD2PKEY={"PG - OTA":"p-ota","PG - Air + PG - Cruise":"p-air","PG - Car + Rev.AI":"p-car",
            "HospiBI":"p-hospibi","Enterprise Connectivity":"p-ec","RezGain":"p-rez","UNO":"p-uno",
            "UNO + RezGain":"p-unorez","SoHo":"p-soho"}
_PLF_ORDER=["prior_act","cur_act","cur_bud","ytd_act","ytd_bud","pfy_ytd_act","py_month_act"]
_PLF_ALIAS={"Revenue":"GAAP Revenue","Gross Profit":"GM","Gross Profit %":"GM %","IT & Telecom":"IT",
            "Account Management":"AM","Bad & Doubtful Debt":"BadDebts Others","GAAP EBITDA %":"Net EBITDA %",
            "Rev Share":"Rev Share Rev Share","Client Services":"Client Services Client Services"}
def _plf_tgt(line,kind):
    l=line.strip()
    if l in _PLF_ALIAS: return _PLF_ALIAS[l]
    if kind=="csub": return "COGS"
    return l
def _plf_round(v,kind):
    if v is None: return None
    return round(v*100,2) if kind=="pct" else round(v)
def map_plf(std, prev=None):
    tmpl=prev or {}
    if "P1" not in tmpl: return prev          # need the template structure
    prods={}
    for rec in std.get("daas_dist_soho",[]):
        pk=_PROD2PKEY.get(rec["product"]); 
        if not pk: continue
        prods.setdefault(pk,{})[rec["line_item"].strip()]=[rec.get(k) for k in _PLF_ORDER]
    out={"P1":{}, "P2":tmpl.get("P2",{})}     # P2 (Sojern segments) carried for now
    for pk, rows in tmpl["P1"].items():
        conv=prods.get(pk)
        if not conv: out["P1"][pk]=rows; continue
        newrows=[]; csub=None; texp=None; rev=None; gm=None; eb=None
        def _pctrow(numer):
            if not (numer and rev): return None
            return [ (round(numer[i]/rev[i]*100,2) if (numer[i] is not None and rev[i]) else None) for i in range(7) ]
        for grp,line,kind,vals in rows:
            l=line.strip()
            if kind=="pct":                    # recompute from reliable numerator/Revenue
                if "Gross Profit" in l: nv=_pctrow(gm)
                else: nv=_pctrow(eb)           # GAAP / Gross EBITDA %
                if nv is None: nv=vals
            else:
                v=conv.get(_plf_tgt(line,kind))
                if v is not None:
                    nv=[_plf_round(x,kind) for x in v]
                elif kind=="osub" and csub and texp:
                    nv=[(texp[i]-csub[i]) if (texp[i] is not None and csub[i] is not None) else None for i in range(7)]
                else:
                    nv=vals
            newrows.append([grp,line,kind,nv])
            if kind=="csub": csub=nv
            if kind=="texp": texp=nv
            if kind=="rev" and l=="Revenue": rev=nv
            if kind=="gp": gm=nv
            if kind=="eb": eb=nv
        out["P1"][pk]=newrows
    return out


# ---------- Consolidated P&L -> PL_CONSOL (headline stats computed; narrative carried) ----------
def _plc_M(v,dp): return "\u2013" if v is None else f"${v/1000:.{dp}f}M"
def _plc_var(a,b): return 0 if not (a and b) else round((a-b)/b*100)
def map_pl_consol(std, prev=None):
    prev=prev or {}
    idx={(r["bu"],r["metric"]):r for r in std}
    gr=idx.get(("Consolidated","Gross Revenue"),{}); eb=idx.get(("Consolidated","EBITDA"),{})
    sj=idx.get(("Martech","EBITDA"),{})
    stats=[
      {"k":"Group revenue","v":_plc_M(gr.get("cur_act"),2),"s":"gross revenue"},
      {"k":"Group EBITDA","v":_plc_M(eb.get("cur_act"),1),
       "s":f"vs {_plc_M(eb.get('cur_bud'),2)} budget \u00b7 +{_plc_var(eb.get('cur_act'),eb.get('cur_bud'))}%"},
      {"k":"Sojern EBITDA \u00b7 YTD","v":_plc_M(sj.get("ytd_act"),1),
       "s":f"vs {_plc_M(sj.get('ytd_bud'),1)} budget \u00b7 +{_plc_var(sj.get('ytd_act'),sj.get('ytd_bud'))}%"},
    ]
    # keep the exact k-labels (they carry the month suffix) from prev when present
    if prev.get("stats"):
        for i in range(min(len(stats),len(prev["stats"]))): stats[i]["k"]=prev["stats"][i]["k"]
    return {"stats":stats, "narr":prev.get("narr",[])}   # narrative is hand-authored -> carry


# ---------- Monetization -> MON / MON_CONSOL / MON_CHARTS ----------
_MONKEY={"OTA":"mon-ota","Car":"mon-car","Rev.AI":"mon-revai","Air + Cruise":"mon-air","HospiBI":"mon-hospibi",
         "RezGain":"mon-rez","UNO":"mon-uno","Enterprise Connectivity":"mon-ec","SoHo":"mon-soho"}
def _mon_vals(c): return [rh(c["opp_value"]),rh(c["total_invoicing"]),rh(c["orderbook"]),rh(c["monetization_pct"]*100)]
def map_mon(std, prev=None):
    prev=prev or {}; byp={_MONKEY.get(p["product"]):p for p in std["products"] if _MONKEY.get(p["product"])}
    out=dict(prev)
    for k,tmpl in prev.items():
        p=byp.get(k)
        if not p: continue
        cohs=p["cohorts"]; card=dict(tmpl); rows=[]
        for i,row in enumerate(tmpl.get("rows",[])):
            r=dict(row)
            if i<len(cohs): r["vals"]=_mon_vals(cohs[i])
            rows.append(r)
        card["rows"]=rows; out[k]=card
    return out
def map_mon_consol(std, prev=None):
    prev=prev or {}; cohs=std["consolidated"]["cohorts"]
    rollup=[]
    for i,row in enumerate(prev.get("rollup",[])):
        r=dict(row)
        if i<len(cohs): r["vals"]=_mon_vals(cohs[i])
        rollup.append(r)
    total=next((c for c in cohs if str(c["period"]).lower()=="total"),None)
    cards=[dict(x) for x in prev.get("cards",[])]
    if total and len(cards)>=4:
        cards[0]["v"]=rh(total["opp_value"]); cards[1]["v"]=rh(total["total_invoicing"])
        cards[2]["v"]=rh(total["orderbook"]); cards[3]["v"]=rh(total["monetization_pct"]*100)
    ob=[{"n":x["bu"].replace("DaaS - ",""),"v":rh(x["orderbook"])} for x in std["consolidated"].get("orderbook_by_bu",[])]
    ob.sort(key=lambda z:-z["v"])
    return {"cards":cards,"rollup":rollup,"ob":ob,"obTotal":(rh(total["orderbook"]) if total else prev.get("obTotal")),"note":prev.get("note","")}
def map_mon_charts(std, prev=None):
    prev=prev or {}; out={}
    tr=std["consolidated"]["orderbook_trend"]
    out["consol"]={"line":{"cats":[x["month"] for x in tr],"vals":[rh(x["orderbook"]) for x in tr]}}
    byp={_MONKEY.get(p["product"]):p for p in std["products"] if _MONKEY.get(p["product"])}
    for k,tmpl in prev.items():
        if k=="consol": continue
        p=byp.get(k)
        if not p: out[k]=tmpl; continue
        mo={x["month"]:x for x in p["monthly"]}
        ch=json.loads(json.dumps(tmpl))
        if "bars" in ch:
            cats=ch["bars"]["cats"]
            ch["bars"]["ns"]=[rh(mo[c]["opp_value"]) if c in mo else ch["bars"]["ns"][i] for i,c in enumerate(cats)]
            ch["bars"]["mon"]=[rh(mo[c]["monetised"]) if c in mo else ch["bars"]["mon"][i] for i,c in enumerate(cats)]
        if "line" in ch:
            cats=ch["line"]["cats"]
            ch["line"]["vals"]=[rh(mo[c]["orderbook"]) if c in mo else ch["line"]["vals"][i] for i,c in enumerate(cats)]
        out[k]=ch
    return out


# ---------- Consolidated P&L -> PL_BUS (BU cards: rev/ebitda = YTD; names/products carried) ----------
_PLBUS_BU={"daas":"DaaS","dist":"Distribution","martech":"Martech"}
def map_pl_bus(std, prev=None):
    prev=prev or []
    idx={(r["bu"],r["metric"]):r for r in std}
    out=[]
    for card in prev:
        c=dict(card); bu=_PLBUS_BU.get(card.get("key"))
        if bu:
            rev=idx.get((bu,"Revenue")) or idx.get((bu,"Gross Revenue"))
            eb=idx.get((bu,"EBITDA"))
            if rev and rev.get("ytd_act") is not None: c["rev"]=rh(rev["ytd_act"])
            if eb and eb.get("ytd_act") is not None: c["ebitda"]=rh(eb["ytd_act"])
        out.append(c)
    return out

MAPPERS={
 "GRR_BRIDGE":("grr_STANDARDIZED.json",map_grr_bridge),
 "GRR_RATIOS":("grr_STANDARDIZED.json",map_grr_ratios),
 "REV_ROWS":("revbu_STANDARDIZED.json",map_rev_rows),
 "REG_CONSOL":("regional_STANDARDIZED.json",make_region_mapper("REG_CONSOL")),
 "REG_NORAM":("regional_STANDARDIZED.json",make_region_mapper("REG_NORAM")),
 "REG_EUROPE":("regional_STANDARDIZED.json",make_region_mapper("REG_EUROPE")),
 "REG_APMEA":("regional_STANDARDIZED.json",make_region_mapper("REG_APMEA")),
 "REG_LATAM":("regional_STANDARDIZED.json",make_region_mapper("REG_LATAM")),
 "SGA":("sga_STANDARDIZED.json",map_sga),
 "TRAV":("travel_STANDARDIZED.json",map_trav),
 "HR":("hr_STANDARDIZED.json",map_hr),
 "PRODUCTS":("top_accounts_STANDARDIZED.json",map_products),
 "SOJERN":("sojern_STANDARDIZED.json",map_sojern),
 "PLF":("full_pnl_STANDARDIZED.json",map_plf),
 "PL_CONSOL":("consol_pnl_STANDARDIZED.json",map_pl_consol),
 "MON":("monetization_STANDARDIZED.json",map_mon),
 "MON_CONSOL":("monetization_STANDARDIZED.json",map_mon_consol),
 "MON_CHARTS":("monetization_STANDARDIZED.json",map_mon_charts),
 "PL_BUS":("consol_pnl_STANDARDIZED.json",map_pl_bus),
 "CEO_DASH":("keykpis_STANDARDIZED.json",map_ceo_dash),
 "ESPL":("consol_pnl_STANDARDIZED.json",map_espl),
}

if __name__=="__main__":
    STD="/mnt/user-data/outputs"; ALL=json.load(open("months_all.json")); jun=ALL["jun"]; prev=ALL.get("may",{})
    def load(f): return json.load(open(STD+"/"+f))
    def diff(a,b,p=""):
        o=[]
        if type(a)!=type(b): return [f"{p}: type"]
        if isinstance(a,list):
            if len(a)!=len(b): o.append(f"{p}: len {len(a)}!={len(b)}")
            for i in range(min(len(a),len(b))): o+=diff(a[i],b[i],f"{p}[{i}]")
        elif isinstance(a,dict):
            for k in set(a)|set(b):
                if k not in a: o.append(f"{p}.{k}: miss-mine")
                elif k not in b: o.append(f"{p}.{k}: extra-mine")
                else: o+=diff(a[k],b[k],f"{p}.{k}")
        elif a!=b: o.append(f"{p}: {a!r}!={b!r}")
        return o
    for sec,(f,fn) in MAPPERS.items():
        mine=fn(load(f),prev.get(sec)); d=diff(mine,jun.get(sec))
        # ignore carried hand-written notes/static text
        d=[x for x in d if not any(t in x for t in ['.note:','.sub:','.name:','.desc:','.card','.summary','.cmt'])]
        print(f"{sec:12}: {'EXACT (data)' if not d else str(len(d))+' diffs'}")
        for x in d[:4]: print("     ",x)
