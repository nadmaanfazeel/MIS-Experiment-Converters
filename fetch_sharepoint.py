# -*- coding: utf-8 -*-
"""fetch_sharepoint.py — download this month's Excel files from a SharePoint folder.

Uses Microsoft Graph app-only auth (client credentials).  Downloads every .xlsx in the
configured folder into --out.  run_pipeline.py then matches each file to its converter by
filename prefix, so the folder just needs the 16 files named consistently, e.g.
    GRR_NRR_Ratios_Jul_26.xlsx, HR_Jul_26.xlsx, Full_P_L_Jul_26.xlsx, ...

Env (set as GitHub secrets):
  TENANT_ID           Azure AD tenant id
  CLIENT_ID           app registration (client) id
  CLIENT_SECRET       app registration client secret
  SHAREPOINT_HOST     e.g. rategain.sharepoint.com
  SHAREPOINT_SITE     site path, e.g. /sites/FPandA           (the part after the host)
  SHAREPOINT_FOLDER   folder path inside the site's default drive, e.g. MIS/2026-07
"""
import os, sys, argparse, requests

GRAPH = "https://graph.microsoft.com/v1.0"

def token():
    t = os.environ["TENANT_ID"]
    r = requests.post(
        f"https://login.microsoftonline.com/{t}/oauth2/v2.0/token",
        data={"client_id": os.environ["CLIENT_ID"],
              "client_secret": os.environ["CLIENT_SECRET"],
              "scope": "https://graph.microsoft.com/.default",
              "grant_type": "client_credentials"}, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="./month_files")
    ap.add_argument("--folder", default=os.environ.get("SHAREPOINT_FOLDER", ""))
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    h = {"Authorization": "Bearer " + token()}

    host = os.environ["SHAREPOINT_HOST"]; site = os.environ["SHAREPOINT_SITE"]
    site_meta = requests.get(f"{GRAPH}/sites/{host}:{site}", headers=h, timeout=30)
    site_meta.raise_for_status()
    site_id = site_meta.json()["id"]

    # list children of the target folder in the site's default document library
    folder = a.folder.strip("/")
    url = (f"{GRAPH}/sites/{site_id}/drive/root:/{folder}:/children"
           if folder else f"{GRAPH}/sites/{site_id}/drive/root/children")
    got = 0
    while url:
        resp = requests.get(url, headers=h, timeout=30); resp.raise_for_status()
        data = resp.json()
        for it in data.get("value", []):
            name = it.get("name", "")
            if not name.lower().endswith((".xlsx", ".xlsm")):
                continue
            dl = it.get("@microsoft.graph.downloadUrl")
            if not dl:
                continue
            content = requests.get(dl, timeout=120).content
            with open(os.path.join(a.out, name), "wb") as f:
                f.write(content)
            print(f"  downloaded {name} ({len(content)//1024} KB)")
            got += 1
        url = data.get("@odata.nextLink")
    print(f"fetch_sharepoint: {got} files -> {a.out}")
    if got == 0:
        sys.exit("HALT: no .xlsx files found in the SharePoint folder")

if __name__ == "__main__":
    main()
