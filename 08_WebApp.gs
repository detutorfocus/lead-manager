/**
 * ============================================================
 * 08_WebApp.gs   (NEW FILE — nothing existing was modified here)
 *
 * Web App entry point. Apps Script Web Apps only ever call one
 * function, doGet(e), to decide what HTML to serve. This project
 * previously had no doGet() at all (confirmed: no other file
 * defines one), which is exactly why deploying as a Web App
 * failed with "Script function not found: doGet".
 *
 * This file adds ONE doGet(e), plus a tiny helper the landing
 * page uses to link back to the real Dashboard sheet. It reuses
 * the existing HTML files and server functions; no business
 * logic is duplicated here.
 * ============================================================
 */

/**
 * Pages the Web App is allowed to serve, keyed by the ?page=
 * query parameter. Values are the exact existing HTML file names
 * (case-sensitive, matching HtmlService.createHtmlOutputFromFile
 * calls already used elsewhere in this project, e.g. in
 * 07_SetupMenu.gs's showAddLeadSidebar()). Add an entry here if
 * you ever add another page-able HTML file — nowhere else needs
 * to change.
 */
var WEBAPP_PAGES = {
  'Index': 'Index',
  'AddLead': 'AddLead',
  'SearchLeads': 'SearchLeads',
  'BulkScan': 'BulkScan',
  'ReviewDuplicates': 'ReviewDuplicates',
  'FollowUps': 'FollowUps'
};

/**
 * Web App entry point. Chooses which existing HTML file to render
 * based on ?page=, defaulting to the new Index.html landing page.
 * Unknown ?page= values safely fall back to Index rather than
 * throwing, so a bad or stale link never breaks the app.
 *
 * Uses createTemplateFromFile (not createHtmlOutputFromFile) so
 * Index.html's nav-card links can be built as absolute URLs via
 * ScriptApp.getService().getUrl(). Without this, relative links
 * inside the Web App's sandboxed iframe resolve against
 * script.googleusercontent.com instead of your /exec URL and load
 * blank. Templating a file that has no <?= ?> tags (AddLead,
 * SearchLeads, BulkScan, ReviewDuplicates, FollowUps) is a no-op,
 * so this is safe for all pages, not just Index.
 */
function doGet(e) {
  var requestedPage = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'Index';
  var fileName = WEBAPP_PAGES[requestedPage] || 'Index';

  var template = HtmlService.createTemplateFromFile(fileName);
  template.webAppUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle('LeadManagement')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Small helper for Index.html: a direct link to the real Dashboard
 * sheet (the existing 06_Dashboard.gs / refreshDashboard() already
 * owns all dashboard logic — this just hands back a URL to it,
 * nothing is recomputed or duplicated here).
 */
function getDashboardSheetUrl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreateSheet_(SHEET_DASHBOARD);
  return ss.getUrl() + '#gid=' + sh.getSheetId();
}
