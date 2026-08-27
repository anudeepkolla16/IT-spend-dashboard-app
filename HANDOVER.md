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
| **Spend amounts** | Two sources. **💰 Update Amounts** with Finance's statement writes a whole month (reviewed). The invoice sync fills *empty* month cells from invoice totals, USD only, never overwriting. Editing by hand still works. |
| **Invoices** | One daily Cron job (2 AM UTC) does both halves: mirrors *new* PDFs from the source SharePoint folders, then files invoices out of `invoices@sarasanalytics.com` into `Invoices/{App}/{YYYY-MM}/` and ticks that app's month in the **Invoices tracker** sheet. **📧 Fetch Invoices** runs the mailbox half on demand. |
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

### Checking the Total row

**🧮 Check Totals** verifies that the Total row's `=SUM()` ranges span every app row, and offers to
repair the ones that don't. The ranges were written once and go stale every time an app is added —
the month columns had drifted apart, so some months were leaving rows out of the total.

It always shows the current formula against the proposed one and asks before changing anything,
since this edits formulas rather than values. The range is derived from the live sheet each run
rather than hardcoded, so it stays correct as rows are inserted and is safe to re-run whenever the
sheet grows. Repairs are recorded in `Invoices/_amount-log.json` with `source: "fix-totals"`.

**Files it keeps**

- `Invoices/_amount-map.json` — vendor label → app mappings you've confirmed
- `Invoices/_amount-log.json` — audit trail of the last 200 writes (`GET /api/amounts?action=log`)

**Tests:** `npm install && npm test`. The suite checks the parser reproduces figures already in
the live sheet (Bubble Starter `751.96`, DBT Cloud `531.25`, Google cloud `38686.27`,
Claude seats `118`) and that writes land on the right cell and never on the Total row.

---

## Invoices from the shared mailbox

Invoices sent or forwarded to **`invoices@sarasanalytics.com`** are filed automatically, as the
second half of the daily invoice cron (2 AM UTC), and from **📧 Fetch Invoices** in the header.

**What it does per message**

1. Looks only at mail with a real PDF attachment, and skips account-admin noise
   ("Verify your email", "Password reset") even when it mentions billing.
2. Works out the app from the subject, the attachment filename and the sender's domain —
   taking the *original* sender out of a forwarded message, so `FW: [Bubble] Invoice`
   forwarded by a colleague is filed under Bubble Starter, not against the colleague.
3. Saves the PDF into the procurement archive at
   `Desktop/Anudeep files/Procurment bills/{vendor}/{month}/`, skipping anything already there.
   The vendor folder is the one that app is already mapped to in `_sync-config.json`
   (`Bubble Starter` → `Bubble`, `Cursor pro` → `Cursor`), and the month subfolder **reuses
   whatever is already there** — `Bubble/Aug`, `Cursor/July` — rather than adding a second
   folder beside it. Only when no month folder exists is one created, named `Aug-26`.
4. Ticks that app's month in the **Invoices tracker** sheet (the TRUE/FALSE grid).

**Which month an invoice belongs to.** The invoice's own **billing period** decides it, not the
day the mail arrived: an invoice is filed under the month holding **most of the period it pays
for**. Luzmo's invoice of Aug 26 reads *"period from 2026-08-26 until 2026-09-26"* — six days of
that fall in August and twenty-five in September, and the charge falls due on Sep 09 — so it is
September's, and the folder, the tracker tick and the amount all follow it there. A period inside
one calendar month (`01-JUN-2026 to 30-JUN-2026`) resolves to that month exactly as before, and a
quarterly or annual period, which has no majority month, stays in the month it starts.

Only a **labelled** period moves anything — "period", "billing period", "service term". Two things
on an ordinary invoice look just like a date range and are not one: the invoice date beside the due
date, and a line item's own dates (`Starter Web Plan 8/12/26 - 9/12/26`, of which Bubble sends nine
a month on separate cycles). An invoice that names no period, or whose PDF can't be read, is filed
by the mail's date as it always was. Anything moved is listed in the run summary, and a period more
than two months from the mail is treated as a misread and ignored.

**Invoices already filed under the old rule stay where they are.** A re-read (**📧 Fetch Invoices** →
*Re-read the last 60 days*) files them again under the right month, but the copy the earlier run put
in the mail's month is not deleted — nothing here deletes from the archive. Both folders would then
be totalled and the charge would count in two months, so the run summary names any leftover copy and
its full path; delete it in SharePoint, and clear the old month's cell if it was written from it.

Invoices are filed where they have always been filed by hand, so there is **one archive, not
two**. The folder mirror then copies them on to `Invoices/{App}/` for the dashboard — and
because the mailbox pass runs *first* in the cron, an invoice that arrives by mail reaches the
dashboard in the same run instead of waiting a day.

**Anything it can't place** goes to `Procurment bills/_Unmatched/{month}/` and is listed in the run
summary with its subject and sender, so nothing is silently dropped. If an app has no procurement
folder mapped at all, a folder named after the app is created and reported under `newFolders` —
map it via **Import Invoices** so the mirror starts picking it up. Add the vendor once via
**Update Amounts → Unrecognised vendors** and the mapping applies to mail too — both use the
same `Invoices/_amount-map.json`.

### Re-checking billing periods on invoices already filed

**🗓 Recheck Periods** in the header applies the billing-period rule to the archive as it stands.
Everything filed before that rule existed sits where its email happened to land — Luzmo's invoice
of Aug 26, which bills `2026-08-26 → 2026-09-26`, is in the August folder with its 557.28 in the
August cell.

It runs in two steps, and **nothing moves until you have read the list and confirmed it**:

1. **The scan** opens every PDF in `Procurment bills/{vendor}/{month}/`, reads the billing period
   and the total, and reports which invoices belong in another month and what that would do to the
   sheet. It writes nothing except a cache of what it read, so re-running it is cheap — only PDFs
   it has never opened cost a download. A run reads 25 of them; if the archive is bigger, the
   summary says how many are left and you run it again.
2. **The apply** moves the files you confirmed — in the procurement archive *and* in the
   `Invoices/{App}/` mirror the dashboard reads, or the checklist would keep showing them under the
   old month — then totals each affected month **from the folder as it now stands** and writes those
   cells. The figure written is never the previewed one: a file that failed to move cannot leave
   behind a total that assumes it did.

Nothing is deleted. A Graph move keeps the file's identity, its version history and anyone's link
to it, so putting one back is just a move the other way, and every cell written goes to the audit
log in `_amount-log.json` with what it held before.

**This one lowers a cell**, which the invoice sync never does. That is the point: an invoice that
has moved out of a month is not a missing invoice, it is one that was never that month's. A month
holding any invoice that could not be read, or read as a USD total, is reported and left alone
rather than written short.

Left alone as a matter of course: invoices that state no period (most of them — their arrival month
is the best thing known about them), a period more than two months from where the file sits (a
contract term misread as a cycle), an annual or quarterly period (no majority month — it stays where
it starts), and anything filed deeper than `{vendor}/{month}`.

### Invoice totals and the Spendings sheet

The sync totals **every invoice in the app's month folder** — not just the ones that arrived by
email. Invoices reach the archive by several routes (hand-filed, mirrored from the source folders,
filed by this sync), and an email-only total undercounts badly: Bubble's August mail carried 2 of
its 9 charges, so an email-only total read `64.00` against an actual `524.27`.

It then keeps the app's month cell in Spendings up to date:

| Cell | What happens |
|---|---|
| Empty | Written with the folder total |
| Invoice total is **higher** than the cell | **Updated** — a new invoice has arrived |
| Invoice total is **lower** than the cell | **Left alone**, and reported |
| Equal | Nothing |

Invoices for a month arrive across it — Bubble's ninth August invoice lands on the 28th — so the
cell has to keep up or it sits short for ever while looking final.

**Why a lower total never overwrites:** a new invoice can only add. A shortfall means invoices are
missing from the folder, not that less was spent. Bubble's folder read `492.27` against a correct
`524.27` with the ninth invoice still pending; overwriting there would have replaced a right figure
with a wrong one.

**An upward update will replace a hand correction or a statement figure.** That is deliberate — the
sheet is meant to track invoices — but it is never silent: the run summary flags any figure the
sync did not itself write, and `Invoices/_amount-log.json` records the previous value of every cell
with `source: "invoice"`. If you want a figure to stand regardless, it has to be higher than the
invoice total, or the app needs leaving out of the mailbox flow.

**Only unambiguously USD invoices are used.** This matters more than it sounds. The sheet is
entirely in USD, but Indian vendors bill in INR — Tata Tele's June invoice reads
`Net Payable (INR) 150591.60`, where the sheet correctly holds `1574.40`, the converted figure.
Writing the face value would have been a ~95x error. So an invoice whose currency is not plainly
USD is reported with its amount and currency and left for a human; it is never totalled into the
sheet.

**One vendor, several rows.** Anthropic bills the API console and the Claude seat subscriptions
from the same address, with near-identical subjects and filenames — only the invoice body tells
them apart, so the PDF is read *before* the file is placed:

- line items reading `Claude <model> Usage …` → **Anthropic(Api Console)**
- line items naming a plan or seats (`Enterprise plan`, `Team plan`, `3 accounts`) → **Claude Ai**
- neither → falls back to the amount, above 10k being the API console

Verified against both July 2026 invoices: the API console one totals `13,479.42` and the seats one
`0.00 paid`, which is exactly what the sheet holds for `Jul-26` in each row. Note the threshold is
the *last* resort, not the first — it does not hold historically (Claude Ai was `14,160` in Feb-26,
the API console `371.88` in Jan-26).

**An invoice settled from a prepaid balance counts as nothing charged.** The July Claude receipt
prints `Total $4,037.39` but `$0.00 paid` against an `Applied balance`; the sheet holds `0.00`.
Where a document shows a balance was applied, what was actually paid wins over what was billed.

Totals are matched most-specific first, because invoices state several. Adobe prints
`NET AMOUNT (USD) 34.97` (pre-tax) *before* `GRAND TOTAL (USD) 37.16`; the payable total wins, and
37.16 is what the sheet holds.

**Be aware of what this changes.** Spend figures sourced from invoices will no longer tie exactly
to the bank statement — invoice totals and card charges differ on tax, FX and proration. Apps whose
cells are filled from invoices are recorded in `Invoices/_amount-log.json` with `source: "invoice"`,
so the two can always be told apart.

**Re-reading invoices already seen:** the sync skips mail it has processed before. To pick up totals
from invoices filed before amounts were being read, add `?rescan=1`:
```
curl -s -X POST -H "Authorization: Bearer <CRON_SECRET>" \
  "https://it-spend-dashboard-app.vercel.app/api/invoices/sync-cron?mode=mail&rescan=1"
```

**Files it keeps**

- `Invoices/_mail-sync.json` — last run time and recently seen message IDs (so nothing is filed twice)
- `Invoices/_sync-config.json` — the folder→app mapping, read in **both** directions: the mirror
  uses folder→app, the mailbox sync uses app→folder
- `Invoices/_invoice-index.json` — what was filed, from whom, for which app and month

**Manual run** (`mode=mail` for the mailbox only, `mode=folders` for the folder mirror, omit for both):
```
curl -s -X POST -H "Authorization: Bearer <CRON_SECRET>" \
  "https://it-spend-dashboard-app.vercel.app/api/invoices/sync-cron?mode=mail"
```
The response nests each half under `folders` and `mail`. A mailbox failure (most likely a missing
`Mail.Read` grant) is reported under `mail.error` and does not stop the folder mirror.

> **Setup required before this works:** the app registration needs the **`Mail.Read`**
> *application* permission with admin consent (see Azure section below). Until that is granted
> every run returns a 403 saying exactly that.

---

## The Invoice Checklist tab

The **Invoice Checklist** card on the dashboard answers one question: for every month the sheet says
we were charged, do we actually have the invoice? It has two views, switched by the tabs on the card:

- **Checklist** — a grid of applications down the side and months across the top. Each cell is one of:

  | Mark | Meaning |
  |---|---|
  | `✓` (green) | The sheet records a charge and there is at least one invoice on file. A number instead of the tick means that many invoices. |
  | `!` (amber) | The sheet records a charge for a month already billed, and **no invoice was found** — the gap worth chasing. |
  | `●` (blue) | Invoices are on file but the sheet records no charge for that month. Either the amount has not been entered yet, or the invoice belongs elsewhere. |
  | `·` | Nothing charged, nothing filed. |
  | `–` | A **future** month the sheet has budgeted. Not a gap — the invoice has not been issued yet. |

  Clicking a row opens that app's drill-down, which lists the invoice files themselves.

- **All invoices** — every PDF in the archive, flat, with the app, the month, a link to the file, and
  the total that was read out of it. A non-USD total is shown greyed with its currency (e.g.
  `INR 150,591.60`) because it is not comparable to the sheet, which is USD throughout.

The filter picks all apps, only apps with a gap, or only apps with invoices on file; the search box
matches app, department or file name.

**Where the numbers come from.** The grid is a join, not a single source:

- the **archive** — `Invoices/{App}/{Month}/`, the mirrored copy the folder sync maintains — says which
  PDFs exist;
- the **spend sheet** (already loaded by the dashboard) says which app-months were charged;
- `Invoices/_invoice-index.json` supplies the per-file totals, joined **by file name**, because the
  amounts were parsed against the procurement folder's paths and the archive is a mirror at a different
  path. A file name that carries two conflicting totals is shown with no total rather than a guessed one.

**Coverage only counts months already billed.** Including the sheet's budgeted future months would
report a shortfall that no amount of filing could ever close.

**Month subfolders are read leniently.** They were named by hand over the years and are not consistent —
`Aug-26` under one vendor, `July` under another, `2026-08` elsewhere; all are understood. A folder named
only for a month takes its year from when the file landed, rolling back a year when the folder names a
month well ahead of that date (a December invoice filed in January belongs to the year before). A PDF
sitting directly in the app folder with no month subfolder is counted and flagged as *undated* rather
than silently dropped.

**Scanning.** The tab loads with the dashboard and the result is cached for five minutes; **↻ Rescan
archive** forces a fresh crawl, and a run of **📧 Fetch Invoices** that files anything triggers one
automatically. A crawl that runs past its deadline is reported as partial and is **not** cached — a
truncated scan would otherwise make closed gaps look open for the next five minutes.

**Endpoint:** `GET /api/invoices/list?mode=checklist` (add `&refresh=1` to bypass the cache). It shares
a route with the per-app listing on purpose — see the function-count note below.

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
| `INVOICE_SOURCE_PATH` | Where invoices are archived. Defaults to `Desktop/Anudeep files/Procurment bills` — note the folder really is spelled *Procurment*. Set this if the archive ever moves. |
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
- **Invoice source:** `Desktop/Anudeep files/Procurment bills/{vendor}/…` — where you drop new invoices. Folder names don't match app names, which is why there's a saved mapping.
- **Auto-sync config:** `Invoices/_sync-config.json` — holds the source folder link and the
  folder→app mapping the cron uses. Created/updated whenever you run **Import Invoices** and confirm.
- **Mailbox invoices:** written into the procurement archive above, under
  `Procurment bills/{vendor}/{month}/` — the same folders used for hand-filed invoices.
- **Unmatched invoices:** `Procurment bills/_Unmatched/{month}/…` — invoices whose vendor didn't match an app row.

---

## Common tasks

**Add someone to the dashboard:** edit `ALLOWED_EMAILS` in Vercel (comma-separated). Effective immediately.

**Point at a different / renamed sheet:** update `TARGET_FILE_PATH` to the new relative path, then redeploy.

**Re-map invoice folders (e.g. new vendor folder):** click **Import Invoices**, paste the source
`Procurment bills` folder link, review the folder→app dropdowns, **Confirm & Import**. This also
re-saves `_sync-config.json`, so the daily cron picks up the new mapping.

**Manually trigger the invoice sync (test):**
```
curl -s -H "Authorization: Bearer <CRON_SECRET>" https://it-spend-dashboard-app.vercel.app/api/invoices/sync-cron
```
Returns `folders` (`copiedNew` / `alreadyPresent`, or a note if no config is saved yet) and
`mail` (what the mailbox pass filed).

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
  mail-sync.js                 Files invoice PDFs from the shared mailbox, ticks the tracker, fills empty amounts
  invoice-amount.js            Reads the payable total + currency out of an invoice PDF
  invoice-period.js            Reads the billing period, and picks the month it mostly covers
  invoices/period-backfill.js  Re-files invoices archived before that rule, in the folders and the sheet
  invoices/inventory.js        Crawls the invoice archive for the checklist tab (month folders, per-file totals)
  amounts/{preview,apply,log}.js  The amount-import handlers, behind api/amounts.js
api/
  spend-data.js                Reads + parses the Excel sheet → JSON the dashboard renders (60s cache)
  amounts.js                   One route for the amount import; dispatches on `action`
  auth/{login,callback,logout,me}.js   Microsoft OAuth sign-in flow
  invoices/
    list.js                    Lists an app's invoices; `?mode=checklist` returns the whole archive
    upload.js                  Manual single-PDF upload from a drill-down modal
    import.js                  Bulk import: preview (suggest matches) + batched commit (skips existing)
    save-sync-config.js        Persists the folder→app mapping to _sync-config.json
    sync-cron.js               Daily job: source-folder mirror + mailbox invoice filing
```

## Notes / gotchas

- **Amounts are USD.** The sheet's native currency; no conversion is applied.
- **Run-rate = trailing 3-month average per app, annualised** (smooths volatile cloud/API costs); future
  budgeted months already in the sheet are excluded from run-rate and "current month" figures.
- **Invoice uploads handle any size** — files over 4 MB go via a Graph resumable upload session.
- **Every Graph call goes through `graphFetch`** (`lib/graph.js`), which retries `429` and the
  `502/503/504` family plus transient network faults, honouring Graph's `Retry-After` (capped at 8s)
  and never sleeping past the run's deadline. A run makes a few hundred Graph calls and Graph
  throttles routinely; before this, the first 429 on a file became a line in the run summary beside
  the invoice it lost. Use it for any new Graph call rather than bare `fetch`.
- **Listings are paged** — `@odata.nextLink` is followed everywhere children are listed
  (`graphListAll`, `listFilesRecursive`, `folderChildren`, the archive's app folders, the mirror's
  source folders). A listing cut off at 200 is worse than one that fails: it reads as "these are all
  the invoices there are", so folder totals come out short and the checklist shows gaps that aren't.
- **A failed listing is never treated as an empty folder.** It used to be: a throttled listing read
  as "this vendor has no month folders", so the run created `Aug-26` beside the existing `Aug` and
  split that month's invoices across two folders, each totalling short. Listings now throw and the
  run reports the file it could not place. Only a real `404` counts as "not there yet".
- **Browsers won't cache the HTML** (no-cache headers), so deployed changes show up on a normal refresh.
- **Cron is once-daily** on the current Vercel plan. If more frequent mirroring is ever needed, that's a plan/scheduler change.
- **`pdf-parse` is pinned to 1.x on purpose, and required by its inner path**
  (`require('pdf-parse/lib/pdf-parse.js')`). Version 2 wraps modern pdf.js, which needs browser
  globals (`DOMMatrix`, `Path2D`) and `@napi-rs/canvas`; on Vercel's Node runtime it throws
  `ReferenceError: DOMMatrix is not defined` **at require time**, which kills the entire function
  before any handler runs — that took the folder mirror down along with invoice reading. The inner
  path also dodges 1.1.1's debug branch, which reads a bundled sample file that isn't deployed.
  The require is lazy and guarded so a future incompatibility degrades to "no amount read" instead
  of a 500. Do not upgrade it without testing a deployed run.
- **The Hobby plan allows 12 Serverless Functions per deployment, and the project is at exactly 12.**
  That is the 11 files under `api/` **plus `middleware.js`**, which also compiles to a function —
  easy to forget when counting. Vercel reports the total as `lambdaRuntimeStats` on a deployment
  (`{"nodejs":12}`). There is **no headroom**: adding one more file under `api/` fails the build with
  `exceeded_serverless_functions_per_deployment`, which is what happened when the amount and mailbox
  endpoints were first added as four separate routes.

  That is why the three amount handlers share `api/amounts.js` (dispatching on `action`) and the
  mailbox sync runs inside `api/invoices/sync-cron.js` (`?mode=mail`), and the invoice checklist
  inside `api/invoices/list.js` (`?mode=checklist`). **Add any new endpoint the same way** — logic in
  `lib/`, dispatched from an existing route — or move the project to a Pro team, which raises the limit.
