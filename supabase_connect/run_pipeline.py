# -*- coding: utf-8 -*-
"""run_pipeline.py — ONE command that does the whole monthly update.

    python3 run_pipeline.py --month jun --prev-month may --files-dir <folder-with-the-16-xlsx>

Steps (no manual anything):
  1. find each folder's Excel in --files-dir (matched by filename prefix, month-agnostic)
  2. run all 16 converters  -> *_STANDARDIZED.json  (validate() gate per folder)
  3. fetch the previous published month's payload from Supabase (for GRR_LAST / HR_PREV / carry)
  4. assemble the dashboard payload for the month
  5. upsert month_payload + mark the month published  (single row; other months untouched)

Env required for the Supabase step (skip with --no-upload to just build the payload):
  SUPABASE_URL   e.g. https://cttvaqfnkmwoqedzsyyd.supabase.co
  SUPABASE_SERVICE_KEY   service_role key (server-side only)

This is the exact body the automation runs — n8n / GitHub Action just invokes this after the
files land in SharePoint.  Reuses the existing converters + assemble_payload.py unchanged."""
import os, sys, json, glob, tempfile, subprocess, argparse, importlib, datetime, urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
CONV = os.path.dirname(HERE)                      # converters live one level up (/mnt/user-data/outputs)
for p in (HERE, CONV):
    if p not in sys.path: sys.path.insert(0, p)

# folder file-prefix -> converter module, entry fn, standardized-json name, how to extract std from the return
PIPE = [
    ("Full_P_L",          "convert_full_pnl",      "convert_all", "full_pnl_STANDARDIZED.json",     "direct"),
    ("Consolidated_P_L",  "convert_consol_pnl",    "convert",     "consol_pnl_STANDARDIZED.json",   "tuple0"),
    ("Monetization",      "convert_monetization",  "convert",     "monetization_STANDARDIZED.json", "direct"),
    ("Infra_Cost",        "convert_infra",         "convert",     "infra_STANDARDIZED.json",        "direct"),
    ("Travel_Expense",    "convert_travel",        "convert",     "travel_STANDARDIZED.json",       "direct"),
    ("SG_A",              "convert_sga",           "convert",     "sga_STANDARDIZED.json",          "direct"),
    ("GRR_NRR",           "convert_grr",           "convert",     "grr_STANDARDIZED.json",          "direct"),
    ("Revenue_by_BU",     "convert_revbu",         "convert",     "revbu_STANDARDIZED.json",        "direct"),
    ("HR_",               "convert_hr",            "convert",     "hr_STANDARDIZED.json",           "direct"),
    ("Regional_Revenue",  "convert_regional",      "convert",     "regional_STANDARDIZED.json",     "direct"),
    ("Top_Accounts",      "convert_top_accounts",  "convert",     "top_accounts_STANDARDIZED.json", "direct"),
    ("Key_KPIs",          "convert_keykpis",       "convert",     "keykpis_STANDARDIZED.json",      "direct"),
    ("Growth",            "convert_growth_margin", "convert",     "growth_margin_STANDARDIZED.json","direct"),
    ("Executive_Summary", "convert_exec_summary",  "convert",     "exec_summary_STANDARDIZED.json", "direct"),
    ("CPI_Tracker",       "convert_cpi",           "convert",     "cpi_STANDARDIZED.json",          "direct"),
    ("Sojern",            "convert_sojern",        "convert",     "sojern_STANDARDIZED.json",       "direct"),
]

def _find(files_dir, prefix):
    hits = [f for f in glob.glob(os.path.join(files_dir, "*.xls*"))
            if os.path.basename(f).lower().startswith(prefix.lower())]
    return sorted(hits)[0] if hits else None

def run_converters(files_dir, std_dir):
    """Run each converter as its own process (writes STD_OUT/<name>_STANDARDIZED.json exactly
    as its __main__ does) so the on-disk shape matches what the mappers expect."""
    ok, missing, failed = [], [], []
    env = dict(os.environ, STD_OUT=std_dir)
    for prefix, mod_name, entry, out_name, how in PIPE:
        path = _find(files_dir, prefix)
        if not path:
            missing.append(prefix); continue
        r = subprocess.run([sys.executable, os.path.join(CONV, mod_name + ".py"), path],
                           capture_output=True, text=True, env=env)
        if r.returncode != 0 or not os.path.exists(os.path.join(std_dir, out_name)):
            failed.append(f"{mod_name}: {(r.stderr or r.stdout)[-160:]}")
        else:
            # surface a FAIL/WARN from the converter's own validate() line if present
            if "VALIDATE: FAIL" in (r.stdout or ""):
                failed.append(f"{mod_name}: validate FAIL")
            else:
                ok.append(out_name)
    return ok, missing, failed

def sb_get_prev(url, key, prev_month):
    req = urllib.request.Request(
        f"{url}/rest/v1/month_payload?month_id=eq.{prev_month}&select=payload",
        headers={"apikey": key, "Authorization": f"Bearer {key}"})
    with urllib.request.urlopen(req, timeout=30) as r:
        rows = json.load(r)
    return rows[0]["payload"] if rows else {}

def sb_upsert(url, key, month, payload):
    body = json.dumps([{"month_id": month, "payload": payload}]).encode()
    req = urllib.request.Request(
        f"{url}/rest/v1/month_payload",
        data=body, method="POST",
        headers={"apikey": key, "Authorization": f"Bearer {key}",
                 "Content-Type": "application/json",
                 "Prefer": "resolution=merge-duplicates,return=minimal"})
    urllib.request.urlopen(req, timeout=60).read()
    # publish + touch updated_at (drives the Sync Now poll)
    body2 = json.dumps({"published": True,
                        "updated_at": datetime.datetime.utcnow().isoformat() + "Z"}).encode()
    req2 = urllib.request.Request(
        f"{url}/rest/v1/months?month_id=eq.{month}",
        data=body2, method="PATCH",
        headers={"apikey": key, "Authorization": f"Bearer {key}",
                 "Content-Type": "application/json", "Prefer": "return=minimal"})
    urllib.request.urlopen(req2, timeout=30).read()

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True)
    ap.add_argument("--prev-month", default=None)
    ap.add_argument("--files-dir", required=True)
    ap.add_argument("--std-dir", default=None)
    ap.add_argument("--prev-file", default=None, help="local {month:payload} json instead of Supabase fetch")
    ap.add_argument("--no-upload", action="store_true")
    a = ap.parse_args()

    std_dir = a.std_dir or tempfile.mkdtemp(prefix="std_")
    url = os.environ.get("SUPABASE_URL"); key = os.environ.get("SUPABASE_SERVICE_KEY")

    print(f"== run_pipeline {a.month} ==\n  files-dir: {a.files_dir}\n  std-dir:   {std_dir}")
    ok, missing, failed = run_converters(a.files_dir, std_dir)
    print(f"  converters ok: {len(ok)}/16 | missing: {missing or '-'} | failed: {failed or '-'}")
    if failed: sys.exit("HALT: converter failure")

    # previous month payload (for derivations / carry)
    prev_payload = {}
    if a.prev_month:
        if a.prev_file:
            prev_payload = json.load(open(a.prev_file)).get(a.prev_month, {})
        elif url and key and not a.no_upload:
            prev_payload = sb_get_prev(url, key, a.prev_month)
    prev_tmp = os.path.join(std_dir, "_prev.json")
    json.dump({a.prev_month: prev_payload} if a.prev_month else {}, open(prev_tmp, "w"))

    out_file = os.path.join(std_dir, "payload.json")
    cmd = [sys.executable, os.path.join(HERE, "assemble_payload.py"),
           "--month", a.month, "--std-dir", std_dir, "--out", out_file]
    if a.prev_month: cmd += ["--prev", prev_tmp, "--prev-month", a.prev_month]
    r = subprocess.run(cmd, capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode != 0:
        print(r.stderr); sys.exit("HALT: assemble failed / missing sections")

    payload = json.load(open(out_file))[a.month]
    print(f"  payload built: {len(payload)} sections")

    if a.no_upload or not (url and key):
        print("  (skipping Supabase upload — no creds / --no-upload)")
        print(f"  payload written to: {out_file}")
        return
    sb_upsert(url, key, a.month, payload)
    print(f"  UPSERTED + PUBLISHED {a.month} to Supabase — dashboard will refresh on next poll. boom.")

if __name__ == "__main__":
    main()
