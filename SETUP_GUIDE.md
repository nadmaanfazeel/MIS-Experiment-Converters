# MIS Dashboard — Automation Setup (GitHub Actions)

**What this does:** every month you drop the 16 Excel files into the SharePoint folder,
click **Sync Now** on the dashboard, and a GitHub Action runs the whole pipeline and updates
Supabase. The dashboard refreshes on its own. You never run a script by hand.

You set this up **once**. After that, monthly = drop files → click Sync Now.

---

## How it flows

```
Sync Now (dashboard)  →  n8n (relay)  →  GitHub Action
                                             │  1. downloads the 16 files from SharePoint
                                             │  2. runs the 16 converters + assembler
                                             │  3. writes the month to Supabase + publishes
                                             ▼
                                        Dashboard refreshes automatically
```

If anything fails, you see it in **two easy places**: the n8n run (red step) and the GitHub
Actions tab (a red ✗ with the exact error). No command line.

---

## ONE-TIME SETUP

### Phase 1 — Put the code on GitHub (10 min)
1. Go to https://github.com → sign in (or create a free account).
2. Click **+** (top right) → **New repository**. Name it `mis-dashboard`. Set **Private**. Click **Create**.
3. On the new repo page, click **uploading an existing file**.
4. Unzip the package I gave you and **drag ALL of it in** (keep the folders — `convert_*.py`,
   `supabase_connect/`, `fetch_sharepoint.py`, `requirements.txt`, and the hidden
   `.github/workflows/sync.yml`). Click **Commit changes**.

### Phase 2 — Add the secret keys (5 min)
In the repo: **Settings** → **Secrets and variables** → **Actions** → **New repository secret**.
Add each of these (name on the left, value on the right):

| Secret name | What to paste |
|---|---|
| `SUPABASE_URL` | `https://cttvaqfnkmwoqedzsyyd.supabase.co` |
| `SUPABASE_SERVICE_KEY` | your Supabase **service_role** key (Supabase → Settings → API) |
| `SHAREPOINT_HOST` | e.g. `rategain.sharepoint.com` |
| `SHAREPOINT_SITE` | e.g. `/sites/FPandA` |
| `SHAREPOINT_FOLDER` | folder holding the 16 files, e.g. `MIS/current` |
| `TENANT_ID` | from the Azure app (Phase 4) |
| `CLIENT_ID` | from the Azure app (Phase 4) |
| `CLIENT_SECRET` | from the Azure app (Phase 4) |

### Phase 3 — Test it once, by hand (2 min)
1. Repo → **Actions** tab → click **MIS Dashboard Sync** → **Run workflow**.
2. Type the month (e.g. `jul`) and previous month (e.g. `jun`) → **Run workflow**.
3. Watch it go green. Green ✓ = the dashboard is updated. (If red, click it — the failing
   step shows the exact reason, usually "this file looked different this month".)

### Phase 4 — SharePoint access (your IT does this once, ~10 min)
The Action needs permission to read the SharePoint folder. Ask IT to:
1. Azure Portal → **App registrations** → **New registration** (name: `mis-dashboard-reader`).
2. Copy the **Directory (tenant) ID** → that's `TENANT_ID`. Copy **Application (client) ID** → `CLIENT_ID`.
3. **Certificates & secrets** → **New client secret** → copy the value → `CLIENT_SECRET`.
4. **API permissions** → **Add** → Microsoft Graph → **Application permissions** → add
   **Sites.Read.All** → then **Grant admin consent**.
5. Put those three values into the GitHub secrets from Phase 2.

*(If IT can't do the Azure app, tell me — there's an alternative where n8n pulls the files
using the Microsoft 365 login you already have, so no Azure setup is needed.)*

### Phase 5 — Wire the "Sync Now" button (I build this with you)
I'll create a tiny n8n workflow: **Sync Now → n8n → tells GitHub to run**. You'll paste one
webhook URL into the dashboard and redeploy. That's the last step and I'll walk each click.

---

## MONTHLY (after setup) — this is all you do
1. Put the 16 Excel files in the SharePoint folder (same names each month, month changes: `..._Jul_26.xlsx`).
2. Open the dashboard → click **Sync Now**.
3. Done. It refreshes in ~1–2 minutes.

## If something looks off
- Dashboard didn't change → open **Actions** tab, look at the latest run.
  - Green ✓ = it worked (try a hard refresh).
  - Red ✗ = click it; the red step names the problem. 99% of the time it's one Excel file that
    came in a different shape — re-check that file against last month's and click Sync Now again.
- You never edit code. The converters don't change month to month.
