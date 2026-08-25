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
| **Spend amounts** | Click **💰 Update Amounts**, hand it the statement workbook Finance sends, review, apply. It writes the month column for you. Editing the sheet by hand still works — the page re-fetches every 5 minutes and on "Refresh now". |
| **Invoices (mailbox)** | A daily Cron job (3 AM UTC) reads `invoices@sarasanalytics.com`, files each invoice PDF into `Invoices/{App}/{YYYY-MM}/` and ticks that app's month in the **Invoices tracker** sheet. **📧 Fetch Invoices** runs it on demand. |
| **Invoices (folders)** | A daily Cron job (2 AM UTC) mirrors *new* invoice PDFs from the source SharePoint folders into `Invoices/{App}/…`. Only files not already present are copied. |
| **Access** | Microsoft sign-in; only emails in `ALLOWED_EMAILS` are let in. |

**Manual tasks going forward:** ask Finance for the monthly statement workbook and run
**Update Amounts** with it, and occasionally map a vendor the mailbox sync couldn't place.
Invoices forwarded to `invoices@sarasanalytics.com` file themselves.

---

## Updating amounts from a statement

Amounts used to be copied into the sheet by hand each month. **💰 Update Amounts** in the
dashboard header does that job: you give it the statement workbook Finance sends, it works
out what each app's month cell should say, applies the unambiguous ones, and asks about the rest.

**How to run it**

1. Click **💰 Update Amounts** and pick the statement file (`.xlsx`/`.csv`, up to 3 MB).
2. Check the **Month** it detected — taken from the data, not the filename, since statement
   files are named inconsistently (`apps Apr.xlsx`, `Saras appas & subscription June 2026.xlsx`).
3. Work through the review sections, then **Apply to sheet**.

**What lands in each section**

| Section | What it means |
|---|---|
| **Ready to apply** | Vendor matched exactly one app row and the month cell is empty. Ticked by default. |
| **Need a look** | Cell already holds a different figure, the total is wildly off this app's trailing 3-month average, refunds outweigh charges, or the sheet has no column for that month. Unticked — tick what you want written. |
| **Unrecognised vendors** | The vendor label matched no app row. Pick the row it belongs to; the mapping is saved and resolves itself next month. |
| **Already correct** | The sheet already holds the figure the statement gives. Nothing is written. |

Every row has a transaction drill-down showing the individual charges behind the figure, so a
number can be checked before it is written.

**How a charge is assigned to a month.** Default is *the statement's month*. The other option,
*the month it was purchased*, reads the date inside the bank descriptor (`ANTHROPIC 05/31 PURCHASE`),
which can fall in the previous month. Both rules exist because the sheet's own history follows
both: Bubble Starter's May-dated purchases were filed under June, but Chargebee's `05/31` charge
went to May. Switch the rule if a month looks off, and the figures regroup.

**Vendor labels.** Matching is on the statement's `Comments` column, normalised — so
`bubble Starter` and `Bubble Starter` are one app, and known typos like `Likedin` are seeded in.
The statement sheet with no `Comments` column is matched on the bank descriptor instead
(`ANTHROPIC* CLAUDE SUB` is the seat subscription, a different row from the API console).

**Refunds** net against charges rather than adding to them, and any app whose month includes a
refund is sent to review rather than applied silently.

**Safety.** Writes go through the Graph Excel API one cell at a time, so the `=SUM()` totals,
number formats and the other two sheets are untouched — nothing is rewritten except the exact
cells you approve. App rows and month columns are re-read from the live file at write time, so a
row inserted between preview and apply can't misdirect a write. Every write is recorded in
`Invoices/_amount-log.json` with who approved it and what the cell held before.

**Files it keeps**

- `Invoices/_amount-map.json` — vendor label → app mappings you've confirmed
- `Invoices/_amount-log.json` — audit trail of the last 200 writes (`GET /api/amounts/log`)

**Tests:** `npm install && npm test`. The suite checks the parser reproduces figures already in
the live sheet (Bubble Starter `751.96`, DBT Cloud `531.25`, Google cloud `38686.27`,
Claude seats `118`) and that writes land on the right cell and never on the Total row.

---

## Invoices from the shared mailbox

Invoices sent or forwarded to **`invoices@sarasanalytics.com`** are filed automatically. The job
runs on the daily cron (3 AM UTC) and from **📧 Fetch Invoices** in the header.

**What it does per message**

1. Looks only at mail with a real PDF attachment, and skips account-admin noise
   ("Verify your email", "Password reset") even when it mentions billing.
2. Works out the app from the subject, the attachment filename and the sender's domain —
   taking the *original* sender out of a forwarded message, so `FW: [Bubble] Invoice`
   forwarded by a colleague is filed under Bubble Starter, not against the colleague.
3. Saves the PDF to `Invoices/{App}/{YYYY-MM}/`, skipping anything already there.
4. Ticks that app's month in the **Invoices tracker** sheet (the TRUE/FALSE grid).

**Anything it can't place** goes to `Invoices/_Unmatched/{YYYY-MM}/` and is listed in the run
summary with its subject and sender, so nothing is silently dropped. Add the vendor once via
**Update Amounts → Unrecognised vendors** and the mapping applies to mail too — both use the
same `Invoices/_amount-map.json`.

**It does not write spend amounts, on purpose.** The figures in the sheet reconcile against the
bank statement, and an invoice total often differs from what was actually charged — tax, currency
conversion, prepaid or partial billing. Writing invoice totals into the sheet would quietly change
what those numbers mean. The mailbox sync tracks *which invoices arrived*; **Update Amounts**
remains the source of the figures.

**Files it keeps**

- `Invoices/_mail-sync.json` — last run time and recently seen message IDs (so nothing is filed twice)
- `Invoices/_invoice-index.json` — what was filed, from whom, for which app and month

**Manual run:**
```
curl -s -X POST -H "Authorization: Bearer <CRON_SECRET>" https://it-spend-dashboard-app.vercel.app/api/invoices/mail-sync
```

> **Setup required before this works:** the app registration needs the **`Mail.Read`**
> *application* permission with admin consent (see Azure section below). Until that is granted
> every run returns a 403 saying exactly that.

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
| `CRON_SECRET` | Authorizes the daily invoice-sync crons. Vercel auto-sends it as a Bearer token on scheduled runs. |
| `INVOICE_MAILBOX` | Shared mailbox the invoice sync reads. Defaults to `invoices@sarasanalytics.com` (note the plural) if unset. |
| `SPEND_SHEET_NAME` | Worksheet holding the amounts. Defaults to `Spendings`; only set it if that tab is renamed. |

> Env-var changes take effect only on the **next deployment**. To apply: push any commit
> (an empty commit works: `git commit --allow-empty -m "redeploy"`), or redeploy in Vercel.

---

## Azure app registration

- Find it in **portal.azure.com → Entra ID → App registrations** by searching the client ID.
- **API permissions (application):** `Files.ReadWrite.All` **and `Mail.Read`** (Graph), both with
  admin consent granted. `Mail.Read` is what lets the invoice mailbox sync run — without it that
  job returns 403 and nothing else is affected.
- **Restrict `Mail.Read` to the one mailbox.** Granted plainly it can read *every* mailbox in the
  tenant, so scope it with an application access policy (Exchange Online PowerShell):
  ```
  New-ApplicationAccessPolicy -AppId <AZURE_CLIENT_ID> `
    -PolicyScopeGroupId invoices@sarasanalytics.com `
    -AccessRight RestrictAccess -Description "Spend dashboard: invoices mailbox only"
  ```
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
- **Mailbox invoices:** `Invoices/{App}/{YYYY-MM}/…` — written by the mailbox sync.
- **Unmatched invoices:** `Invoices/_Unmatched/{YYYY-MM}/…` — invoices whose vendor didn't match an app row.

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
  excel.js                     Graph Excel API helpers — in-place cell writes, sheet/grid location
  statement.js                 Parses Finance's statement workbook into transactions
  vendor-map.js                Vendor label → app row matching (seeded aliases + descriptor rules)
  spend-sheet.js               Opens the spend workbook, alias map and audit log
  mail.js                      Graph mail helpers — invoice detection, forwarded-sender parsing
api/
  spend-data.js                Reads + parses the Excel sheet → JSON the dashboard renders (60s cache)
  amounts/
    preview.js                 Reads a statement, proposes per-app month figures, flags exceptions
    apply.js                   Writes approved figures into the sheet + appends the audit log
    log.js                     Recent amount writes
  auth/{login,callback,logout,me}.js   Microsoft OAuth sign-in flow
  invoices/
    list.js                    Lists an app's invoices (recurses into month subfolders)
    upload.js                  Manual single-PDF upload from a drill-down modal
    import.js                  Bulk import: preview (suggest matches) + batched commit (skips existing)
    save-sync-config.js        Persists the folder→app mapping to _sync-config.json
    sync-cron.js               Daily job: mirrors new source invoices using the saved mapping
    mail-sync.js               Daily job: files invoice PDFs from the shared mailbox, ticks the tracker
```

## Notes / gotchas

- **Amounts are USD.** The sheet's native currency; no conversion is applied.
- **Run-rate = trailing 3-month average per app, annualised** (smooths volatile cloud/API costs); future
  budgeted months already in the sheet are excluded from run-rate and "current month" figures.
- **Invoice uploads handle any size** — files over 4 MB go via a Graph resumable upload session.
- **Browsers won't cache the HTML** (no-cache headers), so deployed changes show up on a normal refresh.
- **Cron is once-daily** on the current Vercel plan. If more frequent mirroring is ever needed, that's a plan/scheduler change.
