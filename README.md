# Business Lead Duplicate Checker & Lead Management System

A complete Google Sheets + Apps Script system: no server, no database,
no paid API. Everything runs inside `SpreadsheetApp` / `HtmlService`.

## Install

1. Create (or open) a Google Sheet.
2. Extensions → Apps Script.
3. Delete the default `Code.gs` stub.
4. For every `.gs` file in this folder, create a matching script file
   (Files → + → Script) and paste in the contents. Keep the numeric
   prefixes (00_, 01_, 02_ …) as the file names — Apps Script runs
   files in a stable order and this keeps dependencies readable, though
   execution order doesn't actually matter since everything is
   function-scoped.
5. For every `.html` file, create a matching HTML file (Files → + →
   HTML) named exactly as listed (e.g. `AddLead`, without the `.html`
   extension in the Apps Script file name field) and paste in the
   contents.
6. Replace the generated `appsscript.json` (Project Settings → "Show
   'appsscript.json' manifest file in editor") with the one provided
   here.
7. Save, then reload the Google Sheet. A **Lead Manager** menu appears.
8. Run **Lead Manager → Setup / Repair System** once (it will ask for
   authorization the first time — approve the Sheets/UI scopes).
9. Optional: run `loadSampleData` from the Apps Script editor (select
   the function in the toolbar, click Run) to see the duplicate
   checker in action with three seeded leads, two of which are
   near-duplicates of each other.

## Follow-Up Tracking + PM Summary upgrade (latest update)

Follow-up automation, an improved duplicate-checker integration, and a
daily PM summary email were added **into this same app** — not as a
separate project. This required real integration work (not a
copy-paste of the earlier `LeadTracker-Upgrade.zip`, which was built
against a different, hypothetical sheet schema) since every new piece
had to work against this app's actual `Lead ID`-based design.

### Schema change (additive only)
`Leads` gained 10 new columns, appended after the existing 30
(`Record Status`, `Branch / Location`) — nothing existing moved,
renamed, or was overwritten:

`Contact Status`, `Date Contacted`, `Last Contact Date`, `Follow-Up Due`,
`Follow-Up Count`, `Last Follow-Up Date`, `Response Status`,
`Next Action`, `Assigned To`, `Follow-Up Status`

Existing lead rows simply have these columns blank until touched;
`repairSystem()` will backfill the headers and validation/formatting
on your live sheet without touching any existing row's data.

### New files
- **`09_FollowUp.gs`** — `flagFollowUpLeads()`, `getFollowUpLeads()`,
  `markContacted()`, `markFollowUpComplete()`, `recordResponse()`,
  `generateFollowUpMetrics()`. Everything is keyed by **Lead ID** (via
  the existing `getLeadById()`), matching how the rest of this app
  works — not by row number.
- **`10_DailySummary.gs`** — `generateDailyMetrics()`,
  `generateTeamMetrics()`, `sendDailyPMSummary()`. Duplicate counts in
  the summary reuse the existing **Duplicate Review** sheet; team
  activity logging reuses the existing **Audit Log** — no separate
  "automation log" sheet was created since Audit Log already covers
  that role here.
- **`11_Notifications.gs`** — `sendNotification(type, data)` (modular,
  email today, easy to extend later) and `sendFollowUpNotifications()`.
- **`FollowUps.html`** — new sidebar/Web-App page, same design system
  as the other four pages, listing leads with `Follow-Up Status =
  FOLLOW UP DUE`, sorted by most overdue.

### Modified files
- **`00_Config.gs`** — extended `LEAD_COL`/`LEADS_HEADERS` with the 10
  new columns; added `CONTACT_STATUS_VALUES` / `FOLLOWUP_STATUS_VALUES`
  / `RESPONSE_STATUS_VALUES`; extended `SETTINGS_DEFAULTS` and
  `getSettings()` with follow-up threshold days, daily-summary
  enabled/time, follow-up-notifications enabled, PM emails, and a
  default notification email.
- **`03_LeadService.gs`** — `buildLeadRow_()` now also writes sensible
  defaults for new leads (`Contact Status: Not Contacted`,
  `Follow-Up Status: NOT DUE`, plus `Assigned To` if provided). No
  existing field's behavior changed.
- **`06_Dashboard.gs`** — `refreshDashboard()` now also writes a
  **Follow-Up Overview** section (Due / Overdue / Avg days since
  contact) and returns those numbers alongside its existing stats.
  Everything it already computed is untouched.
- **`07_SetupMenu.gs`** — menu gained **View Follow-Ups**, **Run Daily
  Automation**, **Send PM Summary Now**; `setupSystem()`/
  `repairSystem()` now also apply follow-up validation/conditional
  formatting and install/repair the two daily triggers
  (`installTriggers()`, `removeTrackerTriggers()`,
  `repairTrackerTriggers()`, `runDailyTrackerAutomation()` were added).
- **`08_WebApp.gs`** — `FollowUps` added to the page router; **and**,
  while in this file, the Web App's `doGet()` was switched from
  `createHtmlOutputFromFile` to `createTemplateFromFile` so
  `Index.html`'s nav links can be absolute
  (`ScriptApp.getService().getUrl()`). This was the fix for the "Add
  Lead / Search / Bulk Scan pages load blank" issue from a few turns
  back — that fix had only ever been described in chat, never actually
  saved into the project files until now, so it's included here for
  real.
- **`Index.html`** — nav links now absolute (see above), added a
  Follow-Ups nav card and a "Follow-Ups Due" stat.
- **`AddLead.html`** — added an optional **Assigned To** field.
- **`SearchLeads.html`** — added **Assigned To** as a search filter,
  and a **Mark Contacted** button in the lead detail panel (for leads
  still at `Contact Status: Not Contacted` — this is what starts the
  3-day follow-up clock).
- **`04_SearchService.gs`** — `Assigned To` added to the searchable
  field map.

### Settings you should fill in
Open the **Settings** sheet (or run Lead Manager → Setup / Repair
System, which adds the rows with defaults automatically) and set:
- **PM Emails (comma-separated)** — required for the daily summary to
  send at all.
- **Default notification email** — used for follow-up alerts on leads
  with nobody in `Assigned To`.
- **Daily summary time (24h, HH:MM)** — defaults to `18:00`.
- **Follow-up threshold days** — defaults to `3`.

### How the daily automations run
`runDailyTrackerAutomation()` is installed on a trigger at ~7am
(spreadsheet timezone): flags follow-ups, refreshes the dashboard, and
sends follow-up emails. `sendDailyPMSummary()` runs on its own trigger
at the time set in Settings (default 6pm). Both triggers are
(re)installed by `setupSystem()`/`repairSystem()`, which always remove
any tracker-owned triggers first — so re-running setup never creates
duplicates.

### Testing this upgrade
- [ ] Run **Setup / Repair System** once on your live sheet — confirm the 10 new columns appear on Leads, and two new triggers exist (Apps Script editor → Triggers)
- [ ] Set PM Emails in Settings, then **Send PM Summary Now** from the menu — confirm the email arrives with real (likely all-zero on a fresh test) numbers
- [ ] Search for a lead, open its detail panel, click **Mark Contacted** — confirm `Contact Status`/`Last Contact Date` update on the sheet
- [ ] Manually set that lead's `Last Contact Date` 3+ days in the past on the sheet, then run **Run Daily Automation** from the menu — confirm `Follow-Up Status` becomes `FOLLOW UP DUE` and it appears in **View Follow-Ups**
- [ ] From Follow-Ups, click **Mark Complete** and, on a different lead, **Record Reply** — confirm counts/status update and the card disappears
- [ ] Confirm existing duplicate-check, merge, keep-both, create-anyway, bulk scan, and review-duplicates flows all still work exactly as before

## File map

| File | Purpose |
|---|---|
| `00_Config.gs` | Sheet/column constants, Settings reader, current-user resolution, Lead ID generator, formula-injection sanitizer, audit logging |
| `01_Normalize.gs` | All field normalizers (name, Instagram, LinkedIn, website, email, phone, city, sector) + Levenshtein/token fuzzy similarity |
| `02_DuplicateEngine.gs` | Candidate search (cheap exact-match narrowing) + weighted scoring + `checkForDuplicates()` |
| `03_LeadService.gs` | `saveLead`, `getLeadById`, `mergeLeads`, `mergeExistingLeads`, `keepBoth`, `createAnyway`, `archiveLead`, Duplicate Review bookkeeping |
| `04_SearchService.gs` | `searchLeads()` for the Search Leads sidebar |
| `05_BulkScan.gs` | Chunked, client-driven bulk duplicate scan with real progress reporting |
| `06_Dashboard.gs` | `refreshDashboard()` — computes and writes all dashboard stats (now also *returns* them — see Web App fix below) |
| `07_SetupMenu.gs` | `onOpen`, `setupSystem`, `repairSystem`, sidebar openers, `loadSampleData` |
| `08_WebApp.gs` | `doGet(e)` — the Web App entry point (template-based, for absolute nav URLs) — plus `getDashboardSheetUrl()` |
| `09_FollowUp.gs` **(new)** | Follow-up flagging, the follow-up list, mark-contacted/complete, record-response |
| `10_DailySummary.gs` **(new)** | Daily + team metrics, the PM summary email |
| `11_Notifications.gs` **(new)** | `sendNotification()` dispatcher, follow-up email builder |
| `AddLead.html` | Add New Lead sidebar — form, "Check Duplicate", results UI, Merge / Keep Both / Create Anyway |
| `SearchLeads.html` | Search Leads sidebar — now also shows follow-up status and a Mark Contacted action |
| `ReviewDuplicates.html` | Lists open Duplicate Review cases (or, when opened via "Check Selected Lead", runs one ad-hoc check) |
| `BulkScan.html` | "Scan All Leads" progress UI |
| `FollowUps.html` **(new)** | Leads due for follow-up, sorted by most overdue, with Mark Complete / Record Reply |
| `Index.html` | Web App landing page — live stats + navigation to all five screens above |

## UI/UX Redesign (latest update)

`AddLead.html`, `SearchLeads.html`, `BulkScan.html`, and `ReviewDuplicates.html`
were redesigned into a consistent, modern SaaS-style interface. **No
business logic changed** — every `google.script.run` call, server-side
function name, and element ID that existing code depends on was
preserved exactly. Verified mechanically before delivery:

- Every original `google.script.run.<fn>()` call in each file still
  targets the same real server function (`checkForDuplicates`,
  `saveLead`, `mergeLeads`, `keepBoth`, `createAnyway`, `getLeadById`,
  `getCurrentUserEmail`, `searchLeads`, `getManualCheckLead`,
  `getOpenReviews`, `resolveReview`, `scanLeadsBatch`,
  `refreshDashboard`).
- Every original element ID (`businessName`, `sector`, `city`, …,
  `q`, `field`, `results`, `detail`, `bar`, `startBtn`, `status`,
  `summary`, `globalMsg`, `heading`, `list`, etc.) is still present
  with the same meaning.
- All four files pass a JavaScript syntax check on their `<script>`
  contents.
- No duplicate function names and exactly one `doGet()` remain
  project-wide.

**Two small, additive backend helpers were added** (nothing existing
was changed in behavior):
- `getBulkScanInfo()` in `05_BulkScan.gs` — read-only, reuses
  `getBulkScanTotal()` and adds a "last scan time" reading (persisted
  via `PropertiesService` when `scanLeadsBatch()` finishes) so the
  redesigned Bulk Scan card can show real numbers instead of
  hard-coded placeholders.
- `refreshDashboard()` already returns its stats object (added during
  the earlier Web App fix) — reused as-is, not duplicated.

**Design system:** since Apps Script HTMLService serves each HTML file
independently (and sidebars use `createHtmlOutputFromFile`, which does
**not** evaluate template includes), a shared external stylesheet isn't
reliable across both the sidebar and Web App contexts. Per the brief's
own fallback guidance, the same `:root` CSS variables, typography,
button/card/badge/alert/spinner styles are duplicated identically at
the top of all four files and should be kept in sync manually if you
change one.

**Emoji policy applied:** section headings use restrained, decorative
emoji is avoided entirely in favor of inline SVG for every functional
icon (search, add, warning, check, close, filter, scan/layers, etc.) —
none of the four redesigned files use an emoji as a button/status/icon
symbol. `Index.html` (the Web App landing page, added during the
earlier Web App fix) still uses emoji as its nav-card icons — it
wasn't in scope for this redesign request, but if you want full
consistency, swap those four emoji for the same SVG icon set now used
in the app.

### Testing checklist
- [ ] **Add a lead** with all required fields → duplicate check runs → "No strong match" → Save Lead → success card shows Business, Lead ID, Added By
- [ ] **Add a lead that matches an existing one** → banner + match card(s) appear with score, tags, Review Existing / Merge / Keep Both actions, plus Continue Anyway / Cancel
- [ ] **Leave a required field blank** → the field is highlighted with an inline error and focus jumps to the first one (no more relying on `alert()`)
- [ ] **Search** by each filter option, including "All fields" → results render as cards; click a card → detail panel opens with full field list
- [ ] **Search with no matches** → empty state with "No leads found" message
- [ ] **Bulk Scan**: card shows current lead count and last scan time → Run Duplicate Scan → progress bar advances with live counts → completion summary cards (Total / Possible / Likely / Confirmed)
- [ ] **Review Duplicates** (via menu, with no lead selected) → lists open review pairs → Merge / Keep Both / Dismiss each resolve and remove the card
- [ ] **Check Selected Lead** (row selected, then menu action) → opens Review Duplicates in single-check mode, heading changes to "Check: <business>"
- [ ] **Error handling**: temporarily break a server function name to confirm `withFailureHandler` still surfaces a readable error alert in each file
- [ ] **Mobile layout** (narrow width, e.g. phone-sized sidebar or a phone browser on the Web App) → sections stack, buttons stay full-width and reachable, no horizontal scroll
- [ ] **Desktop layout** (Web App at full browser width) → content stays within its max-width column, doesn't stretch awkwardly
- [ ] **Authorization** — first open after redeploy may prompt for Sheets/UI scopes again if new functions were added; accept and retry
- [ ] Confirm sidebars still open correctly from **Lead Manager** menu (Add New Lead, Search Leads, Review Duplicates, Scan All Leads) — unaffected by this redesign since `07_SetupMenu.gs` was not touched

### Files changed
`AddLead.html`, `SearchLeads.html`, `BulkScan.html`, `ReviewDuplicates.html` (full visual redesign) · `05_BulkScan.gs` (added `getBulkScanInfo()`, one `PropertiesService` write on scan completion) — everything else in the project is untouched.



**Why the error happened:** `HtmlService.createHtmlOutputFromFile(...)` in
`07_SetupMenu.gs` opens sidebars from inside the Sheet (`showSidebar`,
called from the menu) — that's a completely different code path from a
Web App deployment. A Web App deployment *always* calls `doGet(e)` first,
and this project never had one, hence *"Script function not found:
doGet"*. A project-wide search (`grep -rn "doGet" *.gs *.html`) confirmed
there was no `doGet` hiding anywhere else either.

**What changed:**
- **Added** `08_WebApp.gs` — a single `doGet(e)` that reads `?page=` from
  the URL and serves the matching existing HTML file (`AddLead`,
  `SearchLeads`, `BulkScan`, `ReviewDuplicates`), defaulting to the new
  `Index.html` when no page is specified or the value is unrecognized.
- **Added** `Index.html` — a responsive landing/dashboard page with live
  stats and a card linking to each existing screen (`?page=AddLead`
  etc.), plus a link straight to the real Dashboard sheet.
- **Modified** `06_Dashboard.gs`: `refreshDashboard()` now also
  `return`s the stats object it already computes, instead of returning
  nothing. Nothing else in that function changed, and its existing
  caller (`openDashboard()` in `07_SetupMenu.gs`) already ignored the
  return value, so this is purely additive — the sheet-writing behavior
  is identical to before.
- Nothing was deleted, renamed, or overwritten. No sheet or lead data is
  touched by any of this.

**How `doGet()` works:**
```js
function doGet(e) {
  var requestedPage = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'Index';
  var fileName = WEBAPP_PAGES[requestedPage] || 'Index';
  return HtmlService.createHtmlOutputFromFile(fileName)
    .setTitle('LeadManagement')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```
`WEBAPP_PAGES` is a small allow-list, so an unexpected or malformed
`?page=` value can never throw — it just falls back to `Index`.

**Sidebar/menu system is untouched:** `07_SetupMenu.gs`'s
`showAddLeadSidebar()`, `showSearchSidebar()`, `showReviewSidebar()`,
`showBulkScanSidebar()`, and `checkSelectedLead()` still call
`HtmlService.createHtmlOutputFromFile(...).setWidth(...)` +
`SpreadsheetApp.getUi().showSidebar(...)` exactly as before — those are
unrelated to `doGet` and keep working inside the Sheet whether or not
the project is ever deployed as a Web App.

### Deploy / redeploy steps
1. Save all files in the Apps Script editor.
2. **Deploy → Manage deployments → (pencil/Edit icon on your Web App
   deployment) → Version: New version → Deploy.** Re-using the same
   deployment ID keeps the same URL, but a **new version** is required
   for it to actually contain today's code — the old deployment does
   **not** auto-update.
3. Copy the Web App URL shown after deploying (or reuse the existing
   one if it didn't change).

### Testing the Web App URL
- Open the deployment URL directly → you should see the **Lead
  Management** landing page with four live stat numbers and five
  navigation cards, not the old error.
- Click **Add New Lead** → URL becomes `...exec?page=AddLead` → the
  existing Add Lead form loads and "Check Duplicate" works exactly as
  it does from the sidebar (same `checkForDuplicates` / `saveLead` /
  `mergeLeads` / `keepBoth` / `createAnyway` functions).
- Click **Search Leads**, **Bulk Scan**, **Review Duplicates** the same
  way — each loads its existing, unmodified HTML file.
- Click **Full Dashboard** → opens the real Dashboard sheet in a new
  tab (via `getDashboardSheetUrl()`).
- Back in the Sheet itself, confirm **Lead Manager → Add New Lead**
  (etc.) still opens the sidebar as before — the Web App and the
  sidebar system share the same HTML files but are otherwise
  independent.

### Pre-flight checks already run against this package
- ✅ Exactly one `function doGet(` in the whole project (`08_WebApp.gs`)
- ✅ `doGet()` references `Index`, which exists as `Index.html`
- ✅ Every `google.script.run.____()` call in every HTML file
  (`getCurrentUserEmail`, `checkForDuplicates`, `mergeLeads`,
  `keepBoth`, `createAnyway`, `saveLead`, `getLeadById`, `searchLeads`,
  `getManualCheckLead`, `getOpenReviews`, `resolveReview`,
  `scanLeadsBatch`, `refreshDashboard`, `getDashboardSheetUrl`) matches
  a real, existing server-side function
- ✅ No duplicate function names anywhere in the project
- ✅ All `.gs` files pass a JavaScript syntax check
- ✅ No sheet, row, or column is deleted or renamed by any change here

## How duplicate scoring works

1. Every identifier is normalized (`01_Normalize.gs`).
2. Instead of comparing a new lead against every row, `findCandidates_`
   narrows the whole Leads sheet down to rows that share at least one
   exact normalized identifier, the same city, or the same first name
   token — that candidate set is what gets the more expensive fuzzy
   name comparison.
3. Each candidate is scored using the weights on the **Settings**
   sheet (Instagram 40, LinkedIn 40, Website 35, Email 35, Phone 40,
   Name similarity up to 35 tiered by similarity %, City 10, Sector 5).
4. Any exact strong-identifier match (Instagram/LinkedIn/Website/Email/
   Phone) forces the result to at least "Likely Duplicate" even if the
   business name looks different — this is the "same handle, different
   name" catch described in the spec.
5. Thresholds (also on Settings): ≥90 Confirmed, ≥70 Likely, ≥50
   Possible, below that "No Strong Match".

Change any weight or threshold on the **Settings** sheet — no code
edits needed.

## Notes on a couple of judgment calls

- **Branch handling**: an optional **Branch / Location** column was
  added (per the spec's suggestion) to `Leads`, the Add Lead form, and
  the merge field map, so branches of the same company can be kept as
  separate rows without being treated as pure duplicates.
- **Merge UI**: the Add New Lead sidebar's "Merge" button uses a
  sensible default (existing values win, blanks filled from the new
  lead, Observation/Source preserved from both) rather than a full
  per-field radio-button grid, to keep the sidebar usable on a phone
  screen. `mergeLeads(leadData, existingLeadId, fieldChoices)` in
  `03_LeadService.gs` already accepts an explicit `fieldChoices` map
  per field (`'new' | 'existing' | 'both'`) if you want to build a
  fuller field-by-field comparison screen on top of it.
- **Bulk scan progress**: true live progress isn't possible from a
  single synchronous Apps Script call, so `scanLeadsBatch()` processes
  leads in batches of 25 and the client calls it repeatedly, updating
  a real progress bar between calls (see `BulkScan.html`).
- **Dashboard charts**: the Dashboard sheet is built as clean labeled
  stat rows plus a Recent Duplicate Alerts table. Add a native Sheets
  chart (Insert → Chart) over the stat range if you want a visual —
  chart objects aren't created programmatically here to keep
  `refreshDashboard()` fast and safe to call often.
