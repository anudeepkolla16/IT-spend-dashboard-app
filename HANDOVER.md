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

**The process it follows is the owner's, set out in September 2026, and every rule below is one
of its points:**

1. *Check the sender domain or subject for the application.* Which row an invoice belongs to
   is decided by **`_vendor-rules.json`** in the archive (editable from the dashboard under
   **❓ Needs your answer → ⚙ Filing rules**). A vendor rule names the sender `domains` it
   mails from, `subject` words for vendors that bill through Stripe (where the domain is
   `stripe.com` and says nothing), and either one `app` or several `apps`. Nothing fuzzy is
   used to file: the old alias/descriptor resolver only supplies a *suggestion* beside a
   question.
2. *Check the billing period.* An invoice belongs to **the month its billing period starts in**:
   a bill for `01-08-2026 to 31-08-2026` that arrives on 1 September is August's, and Luzmo's
   `26 Aug → 26 Sep` cycle is August's too. With no labelled period, the invoice's own **issue
   date** decides (`Date of issue`, `Invoice date`, a bare `Date:` — never the due date). With
   neither, the sync **asks** — the month the mail arrived is never assumed.
3. *Check the line items.* Where one vendor bills several rows, each row's rule lists `text`
   phrases and the first row whose phrase appears in the invoice wins. Anthropic's three rows
   are told apart by the **invoice number's account prefix** — `Q8MUNTUC-…` is the API console,
   `2FSKIDHO-…` the Claude Team plan, `XQRYKLO3-…` the Max accounts (checked against the
   invoices in the archive: the API console's read "One-time credit purchase", billed to
   `ai@`; the Team plan's "Auto recharge extra usage, Team plan", billed to `krishna@`).
   Google's four rows are told apart by the product named on the invoice (`Google Voice`,
   `Google Ads`, `Google Cloud`, `Google Workspace`, in that order — Voice's invoice also says
   "Google Workspace Telecom"). The old `google → GOOGLE ADS` catch-all, which filed Google
   Voice under Ads, is gone.
4. *Never guess — ask.* Anything the rules do not settle (an unknown sender, a several-row vendor
   with no matching phrase, a rule naming a row the sheet lacks, no readable month, an unreadable
   PDF) is **held**: the PDF is parked in `{archive}/_Pending/` and a question is recorded in
   `_pending.json`. Nothing is filed into a vendor or month folder on a guess, and nothing goes to
   `_Unmatched` any more. See "Answering the sync's questions" below.
5. *Update the amounts.* The month's invoices are added up and the cell is **set to the total** —
   see "Invoice totals and the Spendings sheet".
6. *Slack DM on every change.* See "Slack".
7. *Application mapping.* Every answer is written back into the rules, so each vendor is asked
   about once.
8. *Read every invoice type.* See "Reading the PDF".

Per message it: looks only at mail with a real PDF attachment, skipping account-admin noise;
reads the PDF *first* (the rules look at its wording and the month comes from its period); takes the
*original* sender out of a forwarded message; saves the PDF into `{archive}/{vendor}/{month}/`,
reusing whatever the vendor already calls that month (`Bubble/Aug`, `Cursor/July`) and creating
`Aug-26` only when none exists; ticks the app's month in the **Invoices tracker** sheet; and totals
the month.

**Invoices already filed under the old rule stay where they are.** A re-read (**📧 Fetch Invoices** →
*Re-read the last 60 days*) files them again under the right month, but the copy the earlier run put
in the mail's month is not deleted — nothing here deletes from the archive. Both folders would then
be totalled and the charge would count in two months, so the run summary names any leftover copy and
its full path; delete it in SharePoint, and clear the old month's cell if it was written from it.

Invoices are filed where they have always been filed by hand, so there is **one archive, not
two**. The folder mirror then copies them on to `Invoices/{App}/` for the dashboard — and
because the mailbox pass runs *first* in the cron, an invoice that arrives by mail reaches the
dashboard in the same run instead of waiting a day.

**Anything it can't place** goes to `{archive}/_Unmatched/{month}/` and is listed in the run
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

1. **The scan** opens every PDF in `{archive}/{vendor}/{month}/`, reads the billing period and the
   total, and reports which invoices belong in another month and what that would do to the sheet.
   It writes nothing except a cache of what it read, so re-running it is cheap — only PDFs it has
   never opened cost a download. Reads run six at a time and stop on the run's own clock rather than
   at a fixed count, and the dashboard keeps going until nothing is left unread, so an archive of a
   few hundred is one click rather than a dozen. A round that makes no progress stops the loop.
2. **The apply** moves the files you confirmed, then totals each affected month **from the folder as
   it now stands** and writes those cells. The figure written is never the previewed one: a file
   that failed to move cannot leave behind a total that assumes it did.

**Invoices filed loose are reported, never re-filed.** Most of the archive keeps its invoices flat in
the app folder with the month in the name (`jan 26.pdf`, `Aug 2026.pdf`), which is how the checklist
dates them. There is no month folder to move such a file out of, and renaming somebody's files to
impose one is a different job — so when a loose invoice's billing period disagrees with the month
its name reads as, the scan says so and leaves it alone. Whether to rename it or file it under a
month folder is yours to decide.

Both halves work off the archive root `resolveArchiveRoot` finds, never a hardcoded path, so a
rename moves the backfill with it. A vendor folder is matched to its sheet row through the saved
`_sync-config.json` mapping first and the vendor aliases second (`Cumul` → `Cumul(Luzmo)`); a folder
that resolves to no row still has its files moved, but no cell is proposed for it.

Nothing is deleted. A Graph move keeps the file's identity, its version history and anyone's link
to it, so putting one back is just a move the other way, and every cell written goes to the audit
log in `_amount-log.json` with what it held before.

**This one lowers a cell**, which the invoice sync never does. That is the point: an invoice that
has moved out of a month is not a missing invoice, it is one that was never that month's.

**But only a cell the invoices demonstrably account for.** A cell is rewritten when it is empty, when
it matches what its folder totalled before the moves, or when it matches what this app last wrote
there. Anything else is reported with both figures and left alone. The first live run showed why:
Cumul(Luzmo)'s August cell held `14,081.00` while its August folder held one `557.28` invoice, and
the earlier rule offered to replace 14,081.00 with `0.00` once that invoice moved out. A statement or
hand-entered figure is not the invoices' to overwrite, and the backfill cannot know what it is made
of. The same test runs again at write time against the live sheet — approval says which months to
consider, never that a cell may be replaced. A month holding any invoice that could not be read as a
USD total is likewise reported and left alone rather than written short.

Left alone as a matter of course: invoices that state no period (most of them — their arrival month
is the best thing known about them), a period more than two months from where the file sits (a
contract term misread as a cycle), an annual or quarterly period (no majority month — it stays where
it starts), and anything filed deeper than `{vendor}/{month}`.

### Invoice totals and the Spendings sheet

The sync totals **every invoice in the app's month folder** — not just the ones that arrived by
email. Invoices reach the archive by several routes (hand-filed, mirrored from the source folders,
filed by this sync), and an email-only total undercounts badly: Bubble's August mail carried 2 of
its 9 charges, so an email-only total read `64.00` against an actual `524.27`.

**The cell is the month's invoice total.** The owner's rule: "certain applications have 3 to 4
separate invoices in a month; calculate the total for all invoices in that month and show that in
the sheet". So the sum of the month's invoices replaces whatever the cell holds — higher *or*
lower:

| Cell | What happens |
|---|---|
| Empty | Written with the month's total |
| Different from the total, either way | **Set to the total**; a lowering is flagged in the report (it usually means invoices are missing from the folder) |
| Equal | Nothing |
| Some invoice in the month could not be used (unreadable, or not in USD) | **Held** — a partial sum is not the total, and a total that is not the total is never written; the report names the file to fix |

Nothing is ever *added* to a cell, which is what makes a re-run harmless: the same folder writes the
same number twice, never twice the number.

**A replaced figure is never silent.** The report says whose number went — one the sync wrote
itself, or a hand correction / statement figure — and `Invoices/_amount-log.json` records the
previous value of every cell with `source: "invoice"`.

**Only unambiguously USD invoices are used.** This matters more than it sounds. The sheet is
entirely in USD, but Indian vendors bill in INR — Tata Tele's June invoice reads
`Net Payable (INR) 150591.60`, where the sheet correctly holds `1574.40`, the converted figure.
Writing the face value would have been a ~95x error. So an invoice whose currency is not plainly
USD is reported with its amount and currency and left for a human; it is never totalled into the
sheet.

**One vendor, several rows** is settled by the rules file (point 3 above), never by the wording
heuristics or the amount threshold this used to apply — those filed Claude API invoices under
Claude Ai.

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

All directly inside the invoice archive (see the layout section for where that is):

- `_mail-sync.json` — last run time and recently seen message IDs (so nothing is filed twice)
- `_sync-config.json` — the folder→app mapping, read in **both** directions: the mirror
  uses folder→app, the mailbox sync uses app→folder
- `_invoice-index.json` — what was filed, from whom, for which app and month

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

**"not in sheet" and "no spend recorded" are different things.** A folder whose name no row matched is
the first; a folder that resolved to a real row which simply has no amounts entered yet is the
second. Posthog was in the sheet with four invoices on file and nothing charged, and the checklist
called it "not in sheet" because it joined against the pivot — which only carries rows with money in
them. `/api/spend-data` returns every app row it reads (`apps`) alongside the records, and the join
uses both.

- **All invoices** — every PDF in the archive, flat, with the app, the month, a link to the file, and
  the total that was read out of it. A non-USD total is shown greyed with its currency (e.g.
  `INR 150,591.60`) because it is not comparable to the sheet, which is USD throughout.

The filter picks all apps, only apps with a gap, or only apps with invoices on file; the search box
matches app, department or file name.

**Where the numbers come from.** The grid is a join, not a single source:

- the **archive** — `{archive}/{Vendor}/{Month}/`, located by `resolveArchiveRoot` — says which PDFs exist;
- the **spend sheet** (already loaded by the dashboard) says which app-months were charged;
- `{archive}/_invoice-index.json` supplies the per-file totals, joined **by file name** rather than by
  path, so a total survives the file being moved or the archive being renamed. A file name that carries
  two conflicting totals is shown with no total rather than a guessed one.

**The columns are the sheet's months, not the archive's.** The tab answers "for a month we were
charged, do we have the invoice?", so a file dated outside the sheet's period has no column to belong
to. Such files are still counted and still listed — the *dated outside the sheet's months* pill opens
them — they just do not widen the grid.

**Every pill on the card is a filter.** Clicking one narrows the view to exactly the things it
counted, so each number can be opened up rather than taken on trust; clicking it again clears it. The
file-level pills (*on file*, *no month to read*, *dated from the file name*, *dated outside the
sheet's months*) switch to the **All invoices** tab, since that is the list they are talking about.
A pill and the dropdown are two ways of narrowing the same list, so picking either clears the other
rather than silently intersecting.

**Vendor folders are matched to sheet rows before joining.** The archive's folders carry *vendor*
names and the sheet's rows carry *app* names, and the two often disagree — `Bubble` is
`Bubble Starter`, `Luzmo` is `Cumul(Luzmo)`, `Claude Api` is `Anthropic(Api Console)`. Joining on the
raw name reports invoices the archive is holding as missing, which is the very thing this tab exists
to catch. Each folder is resolved in this order:

1. the folder→app mapping saved in `{archive}/_sync-config.json` (written by **Import Invoices**);
2. an exact match after normalising case and punctuation;
3. the **same curated vendor aliases the statement importer uses** (`SEED_ALIASES` in
   `lib/vendor-map.js`), served to the page by the endpoint so the two cannot drift. This is what knows
   `Claude Api` is the API console row. A trailing filing word is stripped first, so `click up invoices`
   reaches the `clickup` alias;
4. one-sided containment, **only when exactly one row matches** — `Cursor` is unambiguously
   `Cursor pro`, but `Claude Ai` must not sweep up `Claude Ai Max 6 Accounts`, which is why the exact
   test comes first and an ambiguous folder is left unresolved rather than guessed.

A folder that resolves to nothing gets its own row marked *not in sheet*, and the count of such
folders is shown as a pill — a wrong tick is worse than a visible gap. When a row's invoices live
under a differently-named folder, that folder name is shown on the row as a chip, so the join is
visible rather than magic.

> **If folders still show as "not matched to a sheet row"**, they are vendors the alias list doesn't
> know. Either add them to `SEED_ALIASES` in `lib/vendor-map.js`, or run **📥 Import Invoices** once
> against the archive folder and confirm the dropdowns — that saves the mapping into
> `_sync-config.json` permanently, and takes precedence over everything else.
>
> Running **Import Invoices** against the archive itself is safe and is now the normal case: the
> import detects that the source folder *is* the archive and copies nothing, saving only the mapping.
> It has to, because copying a source subfolder into `{archive}/{app}/` when the two are the same
> folder would duplicate it under a second name — 85 Anthropic PDFs in both `Claude Api` and
> `Anthropic(Api Console)`, with no easy way back. The review dialog says so before you confirm.

**Coverage only counts months already billed.** Including the sheet's budgeted future months would
report a shortfall that no amount of filing could ever close.

**Where an invoice's month comes from.** Two sources, in this order:

1. **The month subfolder**, when there is one — somebody put the file there deliberately, and that beats
   anything inferred. They were named by hand over the years and are not consistent (`Aug-26` under one
   vendor, `July` under another, `2026-08` elsewhere); all are understood.
2. **The file name**, when there is no month subfolder — which is most of the archive. Adobe, AWS and
   Chargebee keep their invoices flat and put the month in the name: `jan 26.pdf`, `Apr-26.pdf`,
   `June 26.pdf`, `Aug 2026.pdf`. Reading only the subfolder left **220 of 472 invoices undated**, which
   the checklist showed as months charged with no invoice while the PDF sat right there.

A month name never matches inside a longer word, so `Marchant`, `Augustine` and `Decision` are not
months. The upload timestamp this app prefixes onto hand-uploaded files is stripped first — it is when
the file arrived, never which month it bills.

**Two digits after a month name are a year in half the archive and a day in the other half.** Adobe's
`jan 26.pdf` is January 2026; Slack's `March 31 invoice.pdf` is the 31st. Nothing in the name says
which, so the year reading is checked against when the file landed: within 24 months behind and 2
ahead it is taken as a year, otherwise the digits were a day and only the month survives. Reading
every one as a year is what put **Nov 2004, Feb 2012 and Mar 2031** across the grid as real columns.
An explicit **four**-digit year is never second-guessed — a human who wrote `2024` meant 2024.

**Ambiguous numeric dates are refused, not guessed.** `27-01-2026` can only be day-first and `01-13-2026`
can only be month-first, but `03-12-2025` is neither on its own. The order is settled **per folder** from
whichever of its files happen to be unambiguous (Anthropic's `Claude Purchase 27-01-2026.pdf` settles the
rest of that folder), and a folder that offers no evidence gets no guess. Filing an invoice under the
wrong month is worse than leaving it undated: a wrong tick hides a real gap.

**A bare month with no year** — in a folder name or a file name — takes its year from when the file
landed, rolling back a year when the month is well ahead of that date, so a December invoice filed in
January is not dated a year forward. A PDF with no readable month at all is counted and reported rather
than silently dropped.

**Scanning.** The tab loads with the dashboard and the result is cached for five minutes; **↻ Rescan
archive** forces a fresh crawl, and a run of **📧 Fetch Invoices** that files anything triggers one
automatically. A crawl that runs past its deadline is reported as partial and is **not** cached — a
truncated scan would otherwise make closed gaps look open for the next five minutes.

**Endpoint:** `GET /api/invoices/list?mode=checklist` (add `&refresh=1` to bypass the cache). It shares
a route with the per-app listing on purpose — see the function-count note below.

---

### The cell is the total — no top-up

The earlier rule (`max(cell, folderTotal)`, "a cell holds the invoices on file plus whatever the
sheet already knows about") is gone. The cell is set to the month's invoice total and nothing else;
see "Invoice totals and the Spendings sheet". The **🗓 Recheck Periods** backfill follows the same
rule: a cell holding a figure the invoices do not account for is *proposed* (for you to tick) with
the figure it replaces and the invoices behind it, rather than blocked.

**Several folders can map to one app, and the sync files into only one of them.** `Luzmo` and
`Cumul(Luzmo)` both mean `Cumul(Luzmo)`; `Bubble` and `Bubble Starter` both mean `Bubble Starter`;
`Laptop procurment`, `Laptop Repair` and `Laptops sold` all mean `Laptops Procurement`. Whichever
folder the mapping lists first is where invoices are filed and the only one totalled, so months whose
invoices live in the other folder are missed entirely.

**A folder that exists now beats one the mapping names but which is gone.** `Bubble` had been renamed
to `Bubble Starter`, and filing into the stale name would have recreated it and split the vendor in
two. The sync lists the archive once per run and prefers a mapped folder that is really there; it
only ever trades up, so a folder still present is never swapped for a later one, and if none of the
mapped folders exists the first is still used so filing has a home.

> Where one app genuinely has invoices in two folders, nothing merges them — that is a filing
> decision. Move them together, or drop the extra folder from `_sync-config.json`.

**A month's invoices are not all inside that month's folder.** Vendors get filed both ways: Apollo's
`Aug-26/` folder holds only the 27 August invoice, while the 4 August one sits loose in the vendor
folder as `Invoice-A0589F17-0016-Aug 2026.pdf`. Totalling the month folder alone gave `85.00` for a
month that really cost `138.12` — and being higher than the `53.12` already in the cell, it
overwrote and **lost the 4 August charge**.

A month's total is now the month folder **plus** any invoice sitting loose in the vendor folder whose
file name dates it to that month — dated by exactly the rule the checklist uses, so the two always
agree on which month an invoice belongs to. Loose invoices counted this way are named in the run
summary.

**An invoice and its own payment receipt are one charge.** Stripe-billed vendors send both, and a
folder holding both would be totalled twice. Apollo's August folder is exactly that:
`Invoice-A0589F17-0017.pdf` and `Receipt-2601-5895.pdf`, both `$85.00`, the receipt being the payment
for that invoice — summing every PDF gives `170.00`. They are paired by the **invoice** number, which
the receipt carries too; the receipt's own *receipt* number is deliberately not matched, since that
would give the two documents different keys. Only a reference actually read out of the PDF counts:
without one, two files that merely share an amount are two charges, because a vendor billing the same
figure twice in a month is ordinary. A duplicate is reported and does **not** count as an unread PDF
— it is accounted for, not missing, so it must not make the folder look partially read.

**A partial folder total is never written, not even into an empty cell.** When some PDF in an
app-month cannot be used, the figure the sync has is not the month's total. The month is held and
the report names the file; once it reads (or its amount is typed in when answering a question), the
next run writes the cell.

**Non-USD totals are never written**, whatever the comparison says.

### Answering the sync's questions

**❓ Needs your answer** on the dashboard lists every held invoice: the file, the sender, the
subject, what was read from it (total, period, date), the question, and — for a several-row vendor —
the options. Pick the app, give the month if asked, type the amount only if the PDF's total could
not be read, and press **File it**: the PDF moves from `_Pending/` into `{vendor}/{month}/`, the
month is totalled and ticked exactly as a mailed invoice is, and the answer is remembered as a
rule (an unknown sender's domain, or a Stripe sender's name, or the invoice-number prefix for a
several-row vendor). **Ignore** moves it to `_Ignored/` and leaves the sheet alone.

The same questions go to the Slack DM, and a reply there is read at the start of the next run:

```
P12 = Google Voice
P12 = 2                    (the second option the question listed)
P12 = Google Voice, Aug-26
P12 = ignore
```

A reply that does not read cleanly is reported back, never guessed at.

`?mode=inspect&path=<archive path>` on `/api/invoices/sync-cron` reads one archived PDF the way
the sync does and returns the text, total, period, date and invoice number it saw — for an invoice
the sync got wrong, that is the first thing to look at.

### Slack

One DM per run that changed anything or has a question outstanding — what was filed, which cells
were set (and lowered, and whose figure was replaced), what was ticked, what could not be used,
and the open questions with the reply format. A run that filed nothing, wrote nothing and asks
nothing sends nothing. Needs `SLACK_BOT_TOKEN` (the workspace's own Slack app, scopes
`chat:write`, `im:write`, `im:history`) and `SLACK_DM_USER` (the member ID of the person to DM);
without them the run carries on and says the DM was not sent.

### Reading the PDF

Two readers, tried in order: **pdfjs-dist 3.x** (legacy build — runs on plain Node, rebuilds a
broken cross-reference table, which is what "bad XRef entry" was on twenty-one archived invoices
that read fine elsewhere; its text is put back into lines by position, so a label and its figure
in two table cells come out as `Total   USD 816.00`), then **pdf-parse 1.x**. The archive survey of
September 2026 (675 PDFs) found the totals that were missed were pattern gaps, not bad files, and
each is now covered: Google's column layout (figures then labels — `Total in USD` is paired by
position), Sentry's `Total $82.31 USD`, Webflow's `TotalUSD 816.00`, PostHog's `$0.00` invoices
(a real total of nothing, not an unread file), the currency detector reading "any currency" as
`ANY`, and a per-vendor `currency` in the rules for invoices stating two (PhantomBuster).

---

## 🧹 Tidy Archive

Filing drifts. A whole vendor folder gets dragged inside another and its month
folders end up a level too deep — `Luzmo` inside `Cumul(Luzmo)` left three July invoices at
`Cumul(Luzmo)/Luzmo/July/`. The sync totals `{vendor}/{month}/` and **nothing below it**, so those
invoices counted towards no month at all.

**Tidy Archive** fixes exactly that one thing: a month folder buried deeper than `{vendor}/{month}`
is lifted up to where the sync looks. Nothing moves until you have read the list and confirmed, and
the apply executes only the moves the plan proposed.

It is deliberately narrow:

- **It never deletes.** Empty folders left behind are reported and left in place.
- **It never overwrites.** A name already at the destination fails that one move and no other.
- **It never lifts a copy of an invoice already filed under that vendor.** `20260826_20260258.pdf`
  is in `Cumul(Luzmo)/Aug-26/` *and* in the stranded `Sep-26/`; lifting the second would put the
  identical charge in two months. Those are reported for a human — which of two copies is the real
  one is a filing decision, not a mechanical one.
- **It ignores subfolders that are not months.** `Quotations/` under `Laptop procurment` holds dated
  folders that are not months, and `Claude Ai/Max accounts/` is a different app's invoices. Neither
  is this job's business.
- **Every destination is re-resolved server-side**, and path traversal is rejected on the segments
  rather than by a prefix test — `…/Invoices/../elsewhere` starts with the archive path and still
  escapes it, so a prefix check alone would wave it through. The plan is echoed back by the browser,
  so an edited payload must not be able to move an invoice off somewhere else on the drive.

After tidying, run **📧 Fetch Invoices** so the months that changed are re-totalled.

**Endpoint:** `POST /api/invoices/import` with `mode: 'tidy-plan'` then `mode: 'tidy-apply'` plus the
plan's `moves`. It shares that route for the same reason everything else does — the function limit.

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
| `INVOICE_ARCHIVE_PATH` | Where invoices are archived, relative to the OneDrive root. Normally **unset** — the path is probed (see "The archive path is resolved, not hardcoded" below). Set it only if the folder is renamed to something the probe doesn't know. |
| `INVOICE_SOURCE_PATH` | Older name for the same thing; still honoured, second in the probe order. Prefer `INVOICE_ARCHIVE_PATH`. |
| `SPEND_SHEET_NAME` | Worksheet holding the amounts. Defaults to `Spendings`; only set it if that tab is renamed. |
| `SLACK_BOT_TOKEN` | `xoxb-…` token of the workspace's Slack app (scopes `chat:write`, `im:write`, `im:history`). Unset = no DMs, runs still work. Set on 2 Sep 2026. |
| `SLACK_DM_USER` | Slack member ID (`U…`) of the person the run reports to and reads answers from. Set on 2 Sep 2026 (`U09BGGLK338`). |

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
- **Invoice archive:** one folder, currently `Desktop/Anudeep files/Invoices/`. Everything reads and
  writes under it — hand-filed invoices, mailbox invoices, dashboard uploads, and the sync's own
  bookkeeping. Inside it: `{vendor or app}/{month}/invoice.pdf`.
- **Unmatched invoices:** `{archive}/_Unmatched/{month}/…` — invoices whose vendor didn't match an app row.
- **Bookkeeping files**, all directly inside the archive: `_sync-config.json` (source link +
  folder→app mapping), `_mail-sync.json` (last run, seen message IDs), `_invoice-index.json`
  (what was filed and each invoice's parsed total), `_amount-map.json` (confirmed vendor aliases),
  `_amount-log.json` (the write audit trail).

### The archive path is **resolved, not hardcoded**

Every path used to be written as `Invoices/…` relative to the OneDrive root, while the archive
actually lived at `Desktop/Anudeep files/Procurment bills` — later renamed to
`Desktop/Anudeep files/Invoices`. Nothing errored. The reads resolved to a folder that did not
exist and quietly returned nothing, so the dashboard reported **all 293 charged months as missing
an invoice**, the per-app invoice lists all read "No invoices uploaded yet", and the next mailbox
sync would have recreated `Procurment bills` and split the archive in two.

So `resolveArchiveRoot` (in `lib/graph.js`) probes for it and takes the first path that actually
exists, in this order:

1. `INVOICE_ARCHIVE_PATH` (if set)
2. `INVOICE_SOURCE_PATH` (the name this used to be configured under)
3. `Desktop/Anudeep files/Invoices`
4. `Desktop/Anudeep files/Procurment bills`
5. `Invoices`

The result is cached per drive for 10 minutes. A **renamed or moved** archive is picked up on its
own within that window. A transient Graph error on one candidate is **rethrown rather than skipped**
— falling through to a later candidate is how invoices end up filed in the wrong folder. If none of
the candidates exists, readers say so explicitly (the checklist tab prints the paths it tried)
rather than rendering an empty archive, because on screen "no invoices" and "no archive" look
identical and only one of them is true.

**If the folder is renamed to something not on that list**, set `INVOICE_ARCHIVE_PATH` in Vercel to
the new path, relative to the OneDrive root, and redeploy.

---

## Common tasks

**Add someone to the dashboard:** edit `ALLOWED_EMAILS` in Vercel (comma-separated). Effective immediately.

**Point at a different / renamed sheet:** update `TARGET_FILE_PATH` to the new relative path, then redeploy.

**Re-map invoice folders (e.g. new vendor folder):** click **Import Invoices**, paste the source
invoice archive folder link, review the folder→app dropdowns, **Confirm & Import**. This also
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
  mail-sync.js                 Files invoice PDFs from the shared mailbox, holds what it is unsure of, totals months, ticks the tracker, sets cells, reports by Slack
  slack.js                     Slack DM: post the run report, read the owner's replies
  invoice-amount.js            Reads the payable total + currency out of an invoice PDF (pdfjs, then pdf-parse)
  invoice-period.js            Reads the billing period (the month it starts in) and the invoice date
  invoices/rules.js            The owner's filing rules (_vendor-rules.json): sender → vendor → row; learns answers
  invoices/pending.js          The hold queue (_pending.json, _Pending/): hold, answer, parse a Slack reply
  invoices/period-backfill.js  Re-files invoices archived before that rule, in the folders and the sheet
  invoices/inventory.js        Crawls the invoice archive for the checklist tab (month folders, per-file totals)
                               (the archive's location comes from graph.js's resolveArchiveRoot)
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
    sync-cron.js               The invoice job (run from the dashboard; the cron is off): mailbox filing, Slack answers,
                               pending/pending-resolve, rules-save, inspect, periods, folder mirror
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
