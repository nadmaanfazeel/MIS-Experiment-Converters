# -*- coding: utf-8 -*-
"""fetch_sharepoint.py - download ONE month's 16 Excel files from SharePoint."""
import os, sys, argparse, requests, urllib.parse

GRAPH = "https://graph.microsoft.com/v1.0"

SECTIONS = {
    "CPI Tracker": "CPI_Tracker", "Executive Summary": "Executive_Summary",
    "Growth & Margin Snapshot": "Growth", "GRR NRR Ratios": "GRR_NRR", "HR": "HR_",
    "Infra Cost": "Infra_Cost", "Key KPIs": "Key_KPIs", "Monetization": "Monetization",
    "Regional Revenue": "Regional_Revenue", "Revenue by BU": "Revenue_by_BU",
    "SG&A Expense": "SG_A", "Sojern": "Sojern", "Top Accounts": "Top_Accounts",
    "Travel Expense": "Travel_Expense",
}

def token():
    t = os.environ["TENANT_ID"]
    r = requests.post("https://login.microsoftonline.com/%s/oauth2/v2.0/token" % t,
        data={"client_id": os.environ["CLIENT_ID"], "client_secret": os.environ["CLIENT_SECRET"],
              "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}, timeout=30)
    r.raise_for_status(); return r.json()["access_token"]

def resolve_drive(h):
    d = os.environ.get("SHAREPOINT_DRIVE_ID", "").strip()
    if d:
        return d
    host = os.environ["SHAREPOINT_HOST"]; site = os.environ["SHAREPOINT_SITE"]
    site_id = requests.get("%s/sites/%s:%s" % (GRAPH, host, site), headers=h, timeout=30).json()["id"]
    lib = os.environ.get("SHAREPOINT_LIBRARY", "").strip()
    if not lib:
        return requests.get("%s/sites/%s/drive" % (GRAPH, site_id), headers=h, timeout=30).json()["id"]
    for dv in requests.get("%s/sites/%s/drives" % (GRAPH, site_id), headers=h, timeout=30).json().get("value", []):
        if dv.get("name", "").lower() == lib.lower():
            return dv["id"]
    sys.exit("HALT: library '%s' not found" % lib)

def children(h, drive, path):
    p = urllib.parse.quote(path)
    url = "%s/drives/%s/root:/%s:/children" % (GRAPH, drive, p)
    out = []
    while url:
        r = requests.get(url, headers=h, timeout=30); r.raise_for_status(); j = r.json()
        out += j.get("value", []); url = j.get("@odata.nextLink")
    return out

def dl(item, dest):
    u = item.get("@microsoft.graph.downloadUrl")
    if not u: return False
    open(dest, "wb").write(requests.get(u, timeout=120).content); return True

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="./month_files")
    ap.add_argument("--month-folder", required=True, help="e.g. Jun'26")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    h = {"Authorization": "Bearer " + token()}
    drive = resolve_drive(h)
    base = os.environ.get("SHAREPOINT_BASE", "MIS/2026-2027/MIS_Automation_FY27").strip("/")
    month = a.month_folder
    tag = month.replace("'", "").replace("\u2019", "")

    got, missing = 0, []
    for folder, prefix in SECTIONS.items():
        try:
            kids = children(h, drive, "%s/%s/%s" % (base, folder, month))
        except Exception:
            missing.append(folder); continue
        xls = [k for k in kids if k.get("name", "").lower().endswith((".xlsx", ".xlsm"))]
        if not xls:
            missing.append(folder); continue
        ext = os.path.splitext(xls[0]["name"])[1]
        if dl(xls[0], os.path.join(a.out, prefix + "_" + tag + ext)):
            print("  %s -> %s%s" % (folder, prefix, ext)); got += 1

    try:
        for k in children(h, drive, "%s/P&L/%s" % (base, month)):
            n = k.get("name", "")
            if not n.lower().endswith((".xlsx", ".xlsm")): continue
            pref = "Consolidated_P_L" if "consol" in n.lower() else ("Full_P_L" if "full" in n.lower() else None)
            if pref and dl(k, os.path.join(a.out, pref + "_" + tag + os.path.splitext(n)[1])):
                print("  P&L -> %s" % pref); got += 1
    except Exception:
        missing.append("P&L")

    print("fetch_sharepoint: %d files -> %s" % (got, a.out))
    if missing: print("  WARNING missing:", ", ".join(missing))
    if got < 16: sys.exit("HALT: expected 16 files, got %d (month folder '%s')." % (got, month))

if __name__ == "__main__":
    main()
