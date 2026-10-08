/**
 * ============================================================
 * 07_SetupMenu.gs
 * onOpen menu, idempotent setupSystem(), repairSystem(),
 * and optional loadSampleData().
 * ============================================================
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Lead Manager')
    .addItem('Add New Lead', 'showAddLeadSidebar')
    .addItem('Check Selected Lead', 'checkSelectedLead')
    .addItem('Search Leads', 'showSearchSidebar')
    .addItem('Review Duplicates', 'showReviewSidebar')
    .addSeparator()
    .addItem('View Follow-Ups', 'showFollowUpSidebar')
    .addItem('Run Daily Automation', 'runDailyTrackerAutomation')
    .addItem('Send PM Summary Now', 'menuSendPMSummary_')
    .addSeparator()
    .addItem('Open Dashboard', 'openDashboard')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('Scan All Leads for Duplicates', 'showBulkScanSidebar')
    .addItem('Setup / Repair System', 'setupSystem')
    .addToUi();
}

function showAddLeadSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('AddLead').setTitle('Add New Lead').setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showSearchSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SearchLeads').setTitle('Search Leads').setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showReviewSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ReviewDuplicates').setTitle('Review Duplicates').setWidth(440);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showBulkScanSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('BulkScan').setTitle('Scan All Leads').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showFollowUpSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('FollowUps').setTitle('Follow-Ups').setWidth(440);
  SpreadsheetApp.getUi().showSidebar(html);
}

function menuSendPMSummary_() {
  var result = sendDailyPMSummary();
  var ui = SpreadsheetApp.getUi();
  if (result.skipped) ui.alert('Summary not sent: ' + result.reason);
  else if (result.success) ui.alert('PM summary sent.');
  else ui.alert('Failed to send: ' + result.error);
}

function openDashboard() {
  refreshDashboard();
  var sh = getOrCreateSheet_(SHEET_DASHBOARD);
  SpreadsheetApp.setActiveSheet(sh);
}

/** Lead Manager -> Check Selected Lead: run the duplicate checker against the active row. */
function checkSelectedLead() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_LEADS) {
    ui.alert('Please select a row on the Leads sheet first.');
    return;
  }
  var row = sh.getActiveCell().getRow();
  if (row < 2) {
    ui.alert('Please select a lead row (not the header row).');
    return;
  }
  var values = sh.getRange(row, 1, 1, LEADS_HEADERS.length).getValues()[0];
  var leadId = values[LEAD_COL.LEAD_ID - 1];
  if (!leadId) {
    ui.alert('No lead found in the selected row.');
    return;
  }
  var leadData = {
    businessName: values[LEAD_COL.BUSINESS_NAME - 1],
    sector: values[LEAD_COL.SECTOR - 1],
    city: values[LEAD_COL.CITY - 1],
    instagram: values[LEAD_COL.INSTAGRAM - 1],
    linkedin: values[LEAD_COL.LINKEDIN - 1],
    email: values[LEAD_COL.EMAIL - 1],
    whatsapp: values[LEAD_COL.WHATSAPP - 1],
    website: values[LEAD_COL.WEBSITE - 1],
    excludeLeadId: leadId
  };
  PropertiesService.getDocumentProperties().setProperty('MANUAL_CHECK_LEAD', JSON.stringify(leadData));
  var html = HtmlService.createHtmlOutputFromFile('ReviewDuplicates').setTitle('Check: ' + leadData.businessName).setWidth(440);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Retrieve the lead queued up by "Check Selected Lead", for the sidebar to run on load. */
function getManualCheckLead() {
  var raw = PropertiesService.getDocumentProperties().getProperty('MANUAL_CHECK_LEAD');
  return raw ? JSON.parse(raw) : null;
}

/**
 * Idempotent full setup. Safe to run repeatedly - never destroys
 * existing lead data.
 */
function setupSystem() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }

  // 1. Leads sheet
  var leadsSh = getOrCreateSheet_(SHEET_LEADS);
  ensureHeaders_(leadsSh, LEADS_HEADERS);
  applyLeadsValidation_(leadsSh);
  applyFollowUpValidation_(leadsSh);
  hideNormalizedColumns_(leadsSh);
  freezeAndFormatHeader_(leadsSh, LEADS_HEADERS.length);

  // 2. Duplicate Review sheet
  var reviewSh = getOrCreateSheet_(SHEET_REVIEW);
  ensureHeaders_(reviewSh, REVIEW_HEADERS);
  freezeAndFormatHeader_(reviewSh, REVIEW_HEADERS.length);

  // 3. Audit Log sheet
  var auditSh = getOrCreateSheet_(SHEET_AUDIT);
  ensureHeaders_(auditSh, AUDIT_HEADERS);
  freezeAndFormatHeader_(auditSh, AUDIT_HEADERS.length);

  // 4. Settings sheet
  var settingsSh = getOrCreateSheet_(SHEET_SETTINGS);
  ensureSettingsDefaults_(settingsSh);

  // 5. Dashboard sheet
  getOrCreateSheet_(SHEET_DASHBOARD);
  refreshDashboard();

  // 6. Menu
  onOpen();

  // 7. Follow-up + PM summary triggers
  installTriggers();

  logAudit('System Setup', '', '', '', '', 'setupSystem() run.');

  if (ui) ui.alert('Lead Manager setup complete. All sheets, validation, menus, and daily automation triggers are ready.');
}

/** Check + restore required structure without touching existing lead rows. */
function repairSystem() {
  var report = [];

  var leadsSh = getOrCreateSheet_(SHEET_LEADS);
  var existingHeaders = leadsSh.getRange(1, 1, 1, LEADS_HEADERS.length).getValues()[0];
  var headerOk = LEADS_HEADERS.every(function (h, i) { return existingHeaders[i] === h; });
  if (!headerOk) { ensureHeaders_(leadsSh, LEADS_HEADERS); report.push('Restored/extended Leads headers.'); }
  applyLeadsValidation_(leadsSh); report.push('Reapplied Leads validation rules.');
  applyFollowUpValidation_(leadsSh); report.push('Reapplied follow-up validation and formatting.');
  hideNormalizedColumns_(leadsSh); report.push('Re-hid normalized columns.');
  freezeAndFormatHeader_(leadsSh, LEADS_HEADERS.length);

  ensureHeaders_(getOrCreateSheet_(SHEET_REVIEW), REVIEW_HEADERS);
  ensureHeaders_(getOrCreateSheet_(SHEET_AUDIT), AUDIT_HEADERS);
  ensureSettingsDefaults_(getOrCreateSheet_(SHEET_SETTINGS));
  getOrCreateSheet_(SHEET_DASHBOARD);
  refreshDashboard();
  onOpen();
  var triggerResult = repairTrackerTriggers();
  if (triggerResult.repaired) report.push('Reinstalled missing daily automation triggers.');

  logAudit('System Setup', '', '', '', '', 'repairSystem() run: ' + report.join(' '));
  try { SpreadsheetApp.getUi().alert('Repair complete.\n' + (report.join('\n') || 'No issues found.')); } catch (e) { /* headless */ }
}

function ensureHeaders_(sh, headers) {
  var current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var matches = headers.every(function (h, i) { return current[i] === h; });
  if (!matches) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function freezeAndFormatHeader_(sh, numCols) {
  sh.setFrozenRows(1);
  var range = sh.getRange(1, 1, 1, numCols);
  range.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.autoResizeColumns(1, numCols);
}

function hideNormalizedColumns_(sh) {
  [LEAD_COL.NORM_NAME, LEAD_COL.NORM_INSTAGRAM, LEAD_COL.NORM_LINKEDIN,
   LEAD_COL.NORM_WEBSITE, LEAD_COL.NORM_EMAIL, LEAD_COL.NORM_PHONE].forEach(function (col) {
    try { sh.hideColumns(col); } catch (e) { /* ignore */ }
  });
}

function applyLeadsValidation_(sh) {
  var maxRows = 2000;
  var websiteStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['None', 'Basic', 'Decent', 'Strong'], true).setAllowInvalid(false).build();
  sh.getRange(2, LEAD_COL.WEBSITE_STATUS, maxRows, 1).setDataValidation(websiteStatusRule);

  var icpRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Growth-Stage', 'Investor-Facing', 'Consistency Seeker', 'Early-Stage'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, LEAD_COL.ICP, maxRows, 1).setDataValidation(icpRule);

  var triggerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Funding', 'Product launch', 'New branch', 'Hiring', 'Expansion',
      'Rebranding', 'Increased social activity', 'New website', 'New product/service', 'Other'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, LEAD_COL.TRIGGER, maxRows, 1).setDataValidation(triggerRule);

  var recordStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Merged', 'Archived'], true).setAllowInvalid(false).build();
  sh.getRange(2, LEAD_COL.RECORD_STATUS, maxRows, 1).setDataValidation(recordStatusRule);
}

function applyFollowUpValidation_(sh) {
  var maxRows = 2000;
  var contactStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONTACT_STATUS_VALUES, true).setAllowInvalid(true).build();
  sh.getRange(2, LEAD_COL.CONTACT_STATUS, maxRows, 1).setDataValidation(contactStatusRule);

  var followUpStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FOLLOWUP_STATUS_VALUES, true).setAllowInvalid(true).build();
  sh.getRange(2, LEAD_COL.FOLLOWUP_STATUS, maxRows, 1).setDataValidation(followUpStatusRule);

  var responseStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(RESPONSE_STATUS_VALUES, true).setAllowInvalid(true).build();
  sh.getRange(2, LEAD_COL.RESPONSE_STATUS, maxRows, 1).setDataValidation(responseStatusRule);

  var range = sh.getRange(2, LEAD_COL.FOLLOWUP_STATUS, Math.max(sh.getMaxRows() - 1, 1), 1);
  var rules = sh.getConditionalFormatRules().filter(function (r) {
    return r.getRanges().every(function (rr) { return rr.getColumn() !== LEAD_COL.FOLLOWUP_STATUS; });
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FOLLOW UP DUE').setBackground('#fce8e6').setFontColor('#c5221f').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('REPLIED').setBackground('#e6f4ea').setFontColor('#137333').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FOLLOWED UP').setBackground('#e8f0fe').setFontColor('#1a73e8').setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('NOT DUE').setBackground('#ffffff').setFontColor('#5f6368').setRanges([range]).build());
  sh.setConditionalFormatRules(rules);
}

/** Install (once) the daily automation + PM summary triggers. Removes any it owns first, so re-running never duplicates them. */
function installTriggers() {
  removeTrackerTriggers();
  var settings = getSettings();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  ScriptApp.newTrigger('runDailyTrackerAutomation')
    .timeBased().everyDays(1).atHour(7).nearMinute(0).inTimezone(tz)
    .create();

  var timeParts = (settings.dailySummaryTime || '18:00').split(':');
  var hour = Number(timeParts[0]) || 18;
  ScriptApp.newTrigger('sendDailyPMSummary')
    .timeBased().everyDays(1).atHour(hour).nearMinute(0).inTimezone(tz)
    .create();

  return { success: true };
}

/** Remove all triggers this project owns (identified by handler function name). */
function removeTrackerTriggers() {
  var handlers = ['runDailyTrackerAutomation', 'sendDailyPMSummary'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (handlers.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  return { success: true };
}

/** Checks trigger state and re-creates anything missing, without duplicating what's already there. */
function repairTrackerTriggers() {
  var handlers = ['runDailyTrackerAutomation', 'sendDailyPMSummary'];
  var existing = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { existing[t.getHandlerFunction()] = (existing[t.getHandlerFunction()] || 0) + 1; });
  var needsRepair = handlers.some(function (h) { return existing[h] !== 1; });
  if (needsRepair) installTriggers();
  return { success: true, repaired: needsRepair };
}

/**
 * The main scheduled function: flags follow-ups, refreshes the
 * dashboard, and sends follow-up notifications. The PM summary runs
 * on its own separate trigger (different time of day).
 */
function runDailyTrackerAutomation() {
  var warnings = [];
  var flagResult = { flagged: 0, checked: 0 };
  var notifResult = { sent: 0 };

  try {
    flagResult = flagFollowUpLeads();
    if (!flagResult.success) warnings.push('Follow-up flagging: ' + flagResult.error);
  } catch (err) {
    warnings.push('Follow-up flagging failed: ' + String(err && err.message ? err.message : err));
  }

  try {
    refreshDashboard();
  } catch (err) {
    warnings.push('Dashboard refresh failed: ' + String(err && err.message ? err.message : err));
  }

  try {
    notifResult = sendFollowUpNotifications();
    if (!notifResult.success) warnings.push('Follow-up notifications: ' + notifResult.error);
  } catch (err) {
    warnings.push('Follow-up notifications failed: ' + String(err && err.message ? err.message : err));
  }

  logAudit('Daily Tracker Run', '', '', '', flagResult.flagged || 0,
    (warnings.length ? 'Completed with warnings: ' + warnings.join(' | ') : 'OK') +
    ' — ' + (flagResult.checked || 0) + ' leads checked, ' + (flagResult.flagged || 0) + ' flagged, ' + (notifResult.sent || 0) + ' notified.');

  return { success: true, warnings: warnings, flagged: flagResult.flagged, notificationsSent: notifResult.sent };
}

function ensureSettingsDefaults_(sh) {
  var lastRow = sh.getLastRow();
  var existingKeys = {};
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) { if (r[0]) existingKeys[r[0]] = true; });
  }
  sh.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
  var toAppend = SETTINGS_DEFAULTS.filter(function (d) { return !existingKeys[d[0]]; });
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 2).setValues(toAppend);
  }
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 2);
}

/** Optional sample data for trying out the system. */
function loadSampleData() {
  var samples = [
    { businessName: "Mama T's Kitchen", sector: 'Food and catering', city: 'Yaba, Lagos',
      contactName: 'Temitope', contactRole: 'Owner', instagram: '@mamatkitchen', linkedin: '',
      email: '', whatsapp: '08012345678', website: '', websiteStatus: 'None', icp: 'Early-Stage',
      trigger: 'Increased social activity', observation: 'Posts every day and receives many comments asking for prices, but there is no visible menu or price list, so every order appears to go through DMs.',
      source: 'Instagram outreach' },
    { businessName: 'Mama T Kitchen Lagos', sector: 'Food and catering', city: 'Yaba',
      contactName: 'Chioma', contactRole: 'Manager', instagram: 'mamatkitchen', linkedin: '',
      email: '', whatsapp: '2348012345678', website: '', websiteStatus: 'None', icp: 'Early-Stage',
      trigger: 'New branch', observation: 'Appears to be opening a second branch closer to Yaba market, based on recent story posts showing renovation.',
      source: 'Field visit' },
    { businessName: 'Zenith Fitness Studio', sector: 'Fitness', city: 'Lekki, Lagos',
      contactName: 'Ade Okafor', contactRole: 'Founder', instagram: '@zenithfit', linkedin: '',
      email: 'hello@zenithfitness.ng', whatsapp: '08099887766', website: 'https://zenithfitness.ng',
      websiteStatus: 'Decent', icp: 'Growth-Stage', trigger: 'Hiring',
      observation: 'Running a hiring campaign for two new trainers and recently added a group-class schedule, suggesting rising membership demand.',
      source: 'LinkedIn research' }
  ];
  samples.forEach(function (s) { saveLead(s); });
  refreshDashboard();
  try { SpreadsheetApp.getUi().alert('Sample data loaded: ' + samples.length + ' leads added.'); } catch (e) { /* headless */ }
}
