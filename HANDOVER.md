# Saras — App Spend Dashboard · Handover & Maintenance

A single-page dashboard that reads the app/laptop spend Excel sheet from SharePoint,
shows spend analytics, and stores/auto-mirrors per-app invoice PDFs. Hosted on Vercel,
deployed automatically from GitHub.

- **Live URL:** https://it-spend-dashboard-app.vercel.app
- **Repo:** https://github.com/anudeepkolla16/IT-spend-dashboard-app (branch `main`)
- **Hosting:** Vercel project `it-spend-dashboard-app` (deploys on every push to `main`)
- **Auth:** Microsoft SSO (Entra ID), restricted to an email allowlist

---

## How it runs day-to-day

| Data | How it updates |
|---|---|
| **Spend amounts** | The page re-fetches the Excel sheet from SharePoint every 5 minutes while open, and on "Refresh now". Edit the sheet → dashboard follows. |
| **Invoices** | A daily Vercel Cron job (2 AM UTC) mirrors *new* invoice PDFs from the source SharePoint folders into `Invoices/{App}/…`. Only files not already present are copied. |
| **Access** | Microsoft sign-in; only emails in `ALLOWED_EMAILS` are let in. |

**Manual tasks going forward:** update the Excel sheet when amounts change; drop new
invoice PDFs into the source `Procurement bills/{vendor}/…` folders. Everything else is automatic.

---

## Environment variables (set in Vercel → Settings → Environment Variables)

Secrets live only in Vercel, never in the repo. Names and purpose:

| Variable | Purpose |
|---|---|
| `AZURE_TENANT_ID` | Entra ID tenant for Graph auth |
| `AZURE_CLIENT_ID` | App registration (client) ID |
| `AZURE_CLIENT_SECRET` | App registration client secret — **rotate before it expires** (see below) |
| `TARGET_USER_UPN` | Whose OneDrive holds the sheet + invoices (`anudeep.kolla@sarasanalytics.com`) |
| `TARGET_FILE_PATH` | **Relative path** to the spend sheet, e.g. `Anudeep Excel sheets/Saras Apps & Subscriptions Purchase from Jan 26 .xlsx`. Must be a path, **never a share URL** (that causes a 400 "Resource not found for the segment 'root:'"). |
| `SESSION_SECRET` | Signs the login session cookie |
| `ALLOWED_EMAILS` | Comma-separated allowlist of who can sign in. Add/remove people here — no redeploy needed. |
| `PUBLIC_APP_URL` | `https://it-spend-dashboard-app.vercel.app` (used to build the OAuth redirect) |
| `CRON_SECRET` | Authorizes the daily invoice-sync cron. Vercel auto-sends it as a Bearer token on scheduled runs. |

> Env-var changes take effect only on the **next deployment**. To apply: push any commit
> (an empty commit works: `git commit --allow-empty -m "redeploy"`), or redeploy in Vercel.

---

## Azure app registration

- Find it in **portal.azure.com → Entra ID → App registrations** by searching the client ID.
- **API permissions (application):** `Files.ReadWrite.All` (Graph) with admin consent granted.
- **Authentication → Web redirect URI:** `https://it-spend-dashboard-app.vercel.app/api/auth/callback`
- **Client secret expiry:** secrets expire. When it does, Graph calls start failing —
  create a new secret in the app registration and update `AZURE_CLIENT_SECRET` in Vercel, then redeploy.

---

## SharePoint / OneDrive layout (under `TARGET_USER_UPN`'s OneDrive)

- **Spend sheet:** the file at `TARGET_FILE_PATH`. Amounts are all in USD; month columns are headers like `Jan-26`.
- **Dashboard invoice store:** `Invoices/{App Name}/…` — one folder per app (exact dashboard app name). Written by the app.
- **Invoice source:** `Desktop/Anudeep files/Procurement bills/{vendor}/…` — where you drop new invoices. Folder names don't match app names, which is why there's a saved mapping.
- **Auto-sync config:** `Invoices/_sync-config.json` — holds the source folder link and the
  folder→app mapping the cron uses. Created/updated whenever you run **Import Invoices** and confirm.

---

## Common tasks

**Add someone to the dashboard:** edit `ALLOWED_EMAILS` in Vercel (comma-separated). Effective immediately.

**Point at a different / renamed sheet:** update `TARGET_FILE_PATH` to the new relative path, then redeploy.

**Re-map invoice folders (e.g. new vendor folder):** click **Import Invoices**, paste the source
`Procurement bills` folder link, review the folder→app dropdowns, **Confirm & Import**. This also
re-saves `_sync-config.json`, so the daily cron picks up the new mapping.

**Manually trigger the invoice sync (test):**
```
curl -s -H "Authorization: Bearer <CRON_SECRET>" https://it-spend-dashboard-app.vercel.app/api/invoices/sync-cron
```
Returns `copiedNew` / `alreadyPresent` counts, or a note if no config is saved yet.

**"Numbers look stale/wrong":** first check for the red *"Sync failed"* banner at the top. If present,
it's a sync problem (usually `TARGET_FILE_PATH` set to a URL instead of a path, or an expired client secret),
not a calculation bug. Fix the env var / secret and redeploy.

---

## Code map

```
index.html                     Single-page dashboard (UI + all client logic)
middleware.js                  Auth gate: requires a valid SSO session; lets /api/auth/* and the cron through
vercel.json                    Function timeouts, the daily cron schedule, no-cache headers for the HTML
lib/
  session.js                   Signs/verifies the session cookie
  graph.js                     Shared Microsoft Graph helpers (token, drive, list, upload, share resolve)
api/
  spend-data.js                Reads + parses the Excel sheet → JSON the dashboard renders (60s cache)
  auth/{login,callback,logout,me}.js   Microsoft OAuth sign-in flow
  invoices/
    list.js                    Lists an app's invoices (recurses into month subfolders)
    upload.js                  Manual single-PDF upload from a drill-down modal
    import.js                  Bulk import: preview (suggest matches) + batched commit (skips existing)
    save-sync-config.js        Persists the folder→app mapping to _sync-config.json
    sync-cron.js               Daily job: mirrors new source invoices using the saved mapping
```

## Notes / gotchas

- **Amounts are USD.** The sheet's native currency; no conversion is applied.
- **Run-rate = trailing 3-month average per app, annualised** (smooths volatile cloud/API costs); future
  budgeted months already in the sheet are excluded from run-rate and "current month" figures.
- **Invoice uploads handle any size** — files over 4 MB go via a Graph resumable upload session.
- **Browsers won't cache the HTML** (no-cache headers), so deployed changes show up on a normal refresh.
- **Cron is once-daily** on the current Vercel plan. If more frequent mirroring is ever needed, that's a plan/scheduler change.
