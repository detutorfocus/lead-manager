/**
 * ============================================================
 * 00_Config.gs
 * Core configuration: sheet names, column maps, settings access,
 * current-user resolution, ID generation, sanitization helpers.
 * ============================================================
 */

var SHEET_LEADS = 'Leads';
var SHEET_REVIEW = 'Duplicate Review';
var SHEET_AUDIT = 'Audit Log';
var SHEET_DASHBOARD = 'Dashboard';
var SHEET_SETTINGS = 'Settings';

// 1-indexed column positions for the Leads sheet (must match LEADS_HEADERS order)
var LEAD_COL = {
  LEAD_ID: 1,
  BUSINESS_NAME: 2,
  SECTOR: 3,
  CITY: 4,
  CONTACT_NAME: 5,
  CONTACT_ROLE: 6,
  INSTAGRAM: 7,
  LINKEDIN: 8,
  EMAIL: 9,
  WHATSAPP: 10,
  WEBSITE: 11,
  WEBSITE_STATUS: 12,
  ICP: 13,
  TRIGGER: 14,
  OBSERVATION: 15,
  SOURCE: 16,
  DATE_ADDED: 17,
  ADDED_BY: 18,
  NORM_NAME: 19,
  NORM_INSTAGRAM: 20,
  NORM_LINKEDIN: 21,
  NORM_WEBSITE: 22,
  NORM_EMAIL: 23,
  NORM_PHONE: 24,
  DUP_STATUS: 25,
  DUP_SCORE: 26,
  MATCHED_LEAD_ID: 27,
  LAST_DUP_CHECK: 28,
  RECORD_STATUS: 29,
  BRANCH: 30, // optional "Branch / Location" field appended at the end

  // --- Follow-up tracking (added for the Follow-Up + PM Summary upgrade) ---
  CONTACT_STATUS: 31,
  DATE_CONTACTED: 32,
  LAST_CONTACT_DATE: 33,
  FOLLOWUP_DUE: 34,
  FOLLOWUP_COUNT: 35,
  LAST_FOLLOWUP_DATE: 36,
  RESPONSE_STATUS: 37,
  NEXT_ACTION: 38,
  ASSIGNED_TO: 39,
  FOLLOWUP_STATUS: 40
};

var CONTACT_STATUS_VALUES = [
  'Not Contacted', 'Contacted', 'Replied', 'Interested', 'Not Interested',
  'Follow-Up Needed', 'Meeting Scheduled', 'Converted', 'Closed', 'Do Not Contact'
];
var FOLLOWUP_STATUS_VALUES = ['NOT DUE', 'FOLLOW UP DUE', 'FOLLOWED UP', 'REPLIED'];
var RESPONSE_STATUS_VALUES = ['No Reply', 'Replied'];
var FOLLOWUP_CLOSED_LIKE = ['Converted', 'Closed', 'Do Not Contact', 'Not Interested'];

var LEADS_HEADERS = [
  'Lead ID', 'Business Name', 'Sector', 'City', 'Contact Name', 'Contact Role',
  'Instagram Handle', 'LinkedIn URL', 'Business Email', 'WhatsApp Business Line',
  'Website', 'Website Status', 'ICP', 'Trigger', 'Observation', 'Source',
  'Date Added', 'Added By', 'Normalized Business Name', 'Normalized Instagram',
  'Normalized LinkedIn', 'Normalized Website', 'Normalized Email', 'Normalized Phone',
  'Duplicate Status', 'Duplicate Score', 'Matched Lead ID', 'Last Duplicate Check',
  'Record Status', 'Branch / Location',
  'Contact Status', 'Date Contacted', 'Last Contact Date', 'Follow-Up Due',
  'Follow-Up Count', 'Last Follow-Up Date', 'Response Status', 'Next Action',
  'Assigned To', 'Follow-Up Status'
];

var REVIEW_HEADERS = [
  'Review ID', 'New Lead ID', 'Existing Lead ID', 'New Business', 'Existing Business',
  'Duplicate Score', 'Status', 'Matching Fields', 'Detected At', 'Detected By',
  'Decision', 'Decision Date', 'Decision By'
];

var AUDIT_HEADERS = [
  'Timestamp', 'User', 'Action', 'Lead ID', 'Related Lead ID', 'Business Name',
  'Duplicate Score', 'Details'
];

var SETTINGS_DEFAULTS = [
  ['Organization Name', 'My Organization'],
  ['Default Country', 'Nigeria'],
  ['Phone Country Code', '234'],
  ['Duplicate Threshold', 90],
  ['Likely Duplicate Threshold', 70],
  ['Possible Duplicate Threshold', 50],
  ['Name Similarity Weight', 35],
  ['Instagram Weight', 40],
  ['LinkedIn Weight', 40],
  ['Website Weight', 35],
  ['Email Weight', 35],
  ['Phone Weight', 40],
  ['City Weight', 10],
  ['Sector Weight', 5],
  ['Follow-up threshold days', 3],
  ['Daily summary enabled', 'TRUE'],
  ['Follow-up notifications enabled', 'TRUE'],
  ['Daily summary time (24h, HH:MM)', '18:00'],
  ['PM Emails (comma-separated)', ''],
  ['Default notification email', '']
];

/** Get a sheet by name, creating it if it does not exist. */
function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** Read Settings sheet into a plain object { key: value }, with defaults as fallback. */
function getSettings() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  var map = {};
  SETTINGS_DEFAULTS.forEach(function (row) { map[row[0]] = row[1]; });
  if (sh) {
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach(function (row) {
        if (row[0]) map[row[0]] = row[1];
      });
    }
  }
  var pmEmailsRaw = String(map['PM Emails (comma-separated)'] || '');
  return {
    orgName: map['Organization Name'],
    defaultCountry: map['Default Country'],
    phoneCountryCode: String(map['Phone Country Code'] || '234'),
    thresholds: {
      confirmed: Number(map['Duplicate Threshold']) || 90,
      likely: Number(map['Likely Duplicate Threshold']) || 70,
      possible: Number(map['Possible Duplicate Threshold']) || 50
    },
    weights: {
      name: Number(map['Name Similarity Weight']) || 35,
      instagram: Number(map['Instagram Weight']) || 40,
      linkedin: Number(map['LinkedIn Weight']) || 40,
      website: Number(map['Website Weight']) || 35,
      email: Number(map['Email Weight']) || 35,
      phone: Number(map['Phone Weight']) || 40,
      city: Number(map['City Weight']) || 10,
      sector: Number(map['Sector Weight']) || 5
    },
    followUpThresholdDays: Number(map['Follow-up threshold days']) || 3,
    dailySummaryEnabled: String(map['Daily summary enabled']).toUpperCase() !== 'FALSE',
    followUpNotificationsEnabled: String(map['Follow-up notifications enabled']).toUpperCase() !== 'FALSE',
    dailySummaryTime: String(map['Daily summary time (24h, HH:MM)'] || '18:00'),
    pmEmails: pmEmailsRaw.split(',').map(function (e) { return e.trim(); }).filter(Boolean),
    defaultNotificationEmail: String(map['Default notification email'] || '')
  };
}

/** Resolve the current user's identity, with graceful fallback. */
function getCurrentUser_() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) return email;
  } catch (e) { /* ignore */ }
  try {
    var effective = Session.getEffectiveUser().getEmail();
    if (effective) return effective;
  } catch (e) { /* ignore */ }
  var props = PropertiesService.getUserProperties();
  var fallback = props.getProperty('DISPLAY_NAME');
  return fallback || 'Unknown User';
}

/**
 * Generate a unique, never-reused Lead ID in the form "LD-000001".
 * Uses LockService + PropertiesService counter, cross-checked against
 * the sheet's actual last ID to survive manual edits or counter drift.
 */
function generateLeadId_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var props = PropertiesService.getDocumentProperties();
    var counter = Number(props.getProperty('LEAD_ID_COUNTER') || 0);

    var sh = getOrCreateSheet_(SHEET_LEADS);
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var ids = sh.getRange(2, LEAD_COL.LEAD_ID, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        var m = String(ids[i][0] || '').match(/^LD-(\d+)$/);
        if (m) {
          var n = parseInt(m[1], 10);
          if (n > counter) counter = n;
        }
      }
    }
    counter += 1;
    props.setProperty('LEAD_ID_COUNTER', String(counter));
    return 'LD-' + ('000000' + counter).slice(-6);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Prevent spreadsheet formula injection. If a string begins with a
 * character that Sheets would interpret as a formula trigger, prefix
 * it with an apostrophe so it is stored as plain text.
 */
function sanitizeValue_(val) {
  if (typeof val !== 'string') return val;
  var trimmed = val;
  if (/^[=+\-@]/.test(trimmed)) {
    return "'" + trimmed;
  }
  return trimmed;
}

/** Sanitize every string field on a lead-data object. */
function sanitizeLeadData_(data) {
  var out = {};
  Object.keys(data).forEach(function (k) {
    out[k] = sanitizeValue_(data[k]);
  });
  return out;
}

/** Exposed to the client so the Add Lead form can pre-fill "Added By". */
function getCurrentUserEmail() {
  return getCurrentUser_();
}

/** Append a row to the Audit Log sheet. */
function logAudit(action, leadId, relatedLeadId, businessName, score, details) {
  var sh = getOrCreateSheet_(SHEET_AUDIT);
  sh.appendRow([
    new Date(),
    getCurrentUser_(),
    action,
    leadId || '',
    relatedLeadId || '',
    businessName || '',
    (score === undefined || score === null) ? '' : score,
    details || ''
  ]);
}
