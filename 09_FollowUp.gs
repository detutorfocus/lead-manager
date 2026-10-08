/**
 * ============================================================
 * 09_FollowUp.gs
 * Follow-up flagging + the actions available from the Follow-Ups
 * sidebar. Everything is keyed by Lead ID (via getLeadById /
 * LEAD_COL), consistent with the rest of this app - not by row
 * number, so it stays correct even if rows get reordered, merged,
 * or archived.
 * ============================================================
 */

/** Whole days between a past date and today (0 = today). Null if the date can't be parsed. */
function calculateDaysSinceContact(dateValue) {
  if (!dateValue) return null;
  var d = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
  if (isNaN(d.getTime())) return null;
  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var dMid = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((today.getTime() - dMid.getTime()) / 86400000);
}

/**
 * Scan every active lead once, decide which ones need a follow-up
 * flag, and write Follow-Up Status / Follow-Up Due back in batched
 * column writes (not per-row calls). Never touches any other field.
 */
function flagFollowUpLeads() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var settings = getSettings();
    var sh = getOrCreateSheet_(SHEET_LEADS);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, flagged: 0, checked: 0 };

    var values = sh.getRange(2, 1, lastRow - 1, LEADS_HEADERS.length).getValues();
    var today = new Date();
    var flaggedCount = 0, checked = 0;
    var dueDates = [], statuses = [];

    values.forEach(function (row) {
      var leadId = row[LEAD_COL.LEAD_ID - 1];
      var recordStatus = row[LEAD_COL.RECORD_STATUS - 1] || 'Active';
      if (!leadId || recordStatus !== 'Active') {
        dueDates.push([row[LEAD_COL.FOLLOWUP_DUE - 1] || '']);
        statuses.push([row[LEAD_COL.FOLLOWUP_STATUS - 1] || '']);
        return;
      }
      checked++;

      var contactStatus = row[LEAD_COL.CONTACT_STATUS - 1];
      var responseStatus = row[LEAD_COL.RESPONSE_STATUS - 1];
      var lastContact = row[LEAD_COL.LAST_CONTACT_DATE - 1];
      var daysSince = calculateDaysSinceContact(lastContact);

      var isClosedLike = FOLLOWUP_CLOSED_LIKE.indexOf(contactStatus) !== -1;
      var alreadyReplied = responseStatus === 'Replied' || contactStatus === 'Replied';
      var eligibleStatus = contactStatus === 'Contacted' || contactStatus === 'Follow-Up Needed';

      var newStatus = row[LEAD_COL.FOLLOWUP_STATUS - 1] || 'NOT DUE';
      var newDue = row[LEAD_COL.FOLLOWUP_DUE - 1] || '';

      if (alreadyReplied) {
        newStatus = 'REPLIED';
      } else if (isClosedLike) {
        // leave as-is; never flag converted/closed/DNC/not-interested leads
      } else if (eligibleStatus && daysSince !== null && daysSince >= settings.followUpThresholdDays) {
        if (newStatus !== 'FOLLOW UP DUE') flaggedCount++;
        newStatus = 'FOLLOW UP DUE';
        newDue = today;
      } else if (newStatus === 'FOLLOW UP DUE') {
        newStatus = 'NOT DUE'; // status changed since last flag (e.g. contact status updated)
      }

      dueDates.push([newDue]);
      statuses.push([newStatus]);
    });

    sh.getRange(2, LEAD_COL.FOLLOWUP_DUE, dueDates.length, 1).setValues(dueDates);
    sh.getRange(2, LEAD_COL.FOLLOWUP_STATUS, statuses.length, 1).setValues(statuses);

    return { success: true, flagged: flaggedCount, checked: checked };
  } catch (err) {
    return { success: false, error: 'Follow-up flagging failed.', details: String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Current follow-up list, sorted by most overdue, then oldest last
 * contact date, then assigned team member.
 */
function getFollowUpLeads() {
  var sh = getOrCreateSheet_(SHEET_LEADS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, LEADS_HEADERS.length).getValues();

  var list = [];
  values.forEach(function (row) {
    var leadId = row[LEAD_COL.LEAD_ID - 1];
    if (!leadId || row[LEAD_COL.FOLLOWUP_STATUS - 1] !== 'FOLLOW UP DUE') return;
    var lastContact = row[LEAD_COL.LAST_CONTACT_DATE - 1];
    list.push({
      leadId: leadId,
      businessName: row[LEAD_COL.BUSINESS_NAME - 1],
      contactName: row[LEAD_COL.CONTACT_NAME - 1] || '',
      assignedTo: row[LEAD_COL.ASSIGNED_TO - 1] || '',
      channel: row[LEAD_COL.INSTAGRAM - 1] ? 'Instagram' : (row[LEAD_COL.WHATSAPP - 1] ? 'WhatsApp' : (row[LEAD_COL.EMAIL - 1] ? 'Email' : '')),
      lastContactDate: lastContact ? Utilities.formatDate(new Date(lastContact), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      daysSinceContact: calculateDaysSinceContact(lastContact),
      followUpCount: row[LEAD_COL.FOLLOWUP_COUNT - 1] || 0,
      contactStatus: row[LEAD_COL.CONTACT_STATUS - 1] || '',
      nextAction: row[LEAD_COL.NEXT_ACTION - 1] || ''
    });
  });

  list.sort(function (a, b) {
    if ((b.daysSinceContact || 0) !== (a.daysSinceContact || 0)) return (b.daysSinceContact || 0) - (a.daysSinceContact || 0);
    if (a.lastContactDate !== b.lastContactDate) return a.lastContactDate < b.lastContactDate ? -1 : 1;
    return (a.assignedTo || '').localeCompare(b.assignedTo || '');
  });
  return list;
}

/** Mark a lead's first (or next) contact today - moves it from "Not Contacted" into the follow-up clock. */
function markContacted(leadId, assignedTo) {
  var lead = getLeadById(leadId);
  if (!lead) return { success: false, error: 'Lead not found.' };
  var sh = getOrCreateSheet_(SHEET_LEADS);
  var now = new Date();
  if (!lead['Date Contacted']) sh.getRange(lead.rowIndex, LEAD_COL.DATE_CONTACTED).setValue(now);
  sh.getRange(lead.rowIndex, LEAD_COL.LAST_CONTACT_DATE).setValue(now);
  sh.getRange(lead.rowIndex, LEAD_COL.CONTACT_STATUS).setValue('Contacted');
  sh.getRange(lead.rowIndex, LEAD_COL.FOLLOWUP_STATUS).setValue('NOT DUE');
  if (assignedTo) sh.getRange(lead.rowIndex, LEAD_COL.ASSIGNED_TO).setValue(sanitizeValue_(assignedTo));
  logAudit('Lead Updated', leadId, '', lead['Business Name'], '', 'Marked Contacted.');
  return { success: true };
}

/**
 * Mark a follow-up as completed. Never sets Response Status to
 * "Replied" itself - that's only ever set by recordResponse(),
 * matching the rule that a follow-up being "done" isn't the same
 * as the prospect actually responding.
 */
function markFollowUpComplete(leadId, nextAction) {
  try {
    var lead = getLeadById(leadId);
    if (!lead) return { success: false, error: 'Lead not found.' };
    var sh = getOrCreateSheet_(SHEET_LEADS);
    var currentCount = Number(lead['Follow-Up Count']) || 0;

    sh.getRange(lead.rowIndex, LEAD_COL.FOLLOWUP_STATUS).setValue('FOLLOWED UP');
    sh.getRange(lead.rowIndex, LEAD_COL.LAST_FOLLOWUP_DATE).setValue(new Date());
    sh.getRange(lead.rowIndex, LEAD_COL.FOLLOWUP_COUNT).setValue(currentCount + 1);
    if (nextAction) sh.getRange(lead.rowIndex, LEAD_COL.NEXT_ACTION).setValue(sanitizeValue_(nextAction));

    logAudit('Lead Updated', leadId, '', lead['Business Name'], '', 'Follow-up marked complete.');
    return { success: true };
  } catch (err) {
    return { success: false, error: 'Unable to mark follow-up complete.', details: String(err && err.message ? err.message : err) };
  }
}

/** Record whether the prospect replied. "Replied" immediately stops future follow-up alerts for this lead. */
function recordResponse(leadId, responseStatus) {
  try {
    if (RESPONSE_STATUS_VALUES.indexOf(responseStatus) === -1) {
      return { success: false, error: 'Invalid response status: ' + responseStatus };
    }
    var lead = getLeadById(leadId);
    if (!lead) return { success: false, error: 'Lead not found.' };
    var sh = getOrCreateSheet_(SHEET_LEADS);
    sh.getRange(lead.rowIndex, LEAD_COL.RESPONSE_STATUS).setValue(responseStatus);
    if (responseStatus === 'Replied') {
      sh.getRange(lead.rowIndex, LEAD_COL.FOLLOWUP_STATUS).setValue('REPLIED');
      sh.getRange(lead.rowIndex, LEAD_COL.CONTACT_STATUS).setValue('Replied');
    }
    logAudit('Lead Updated', leadId, '', lead['Business Name'], '', 'Response recorded: ' + responseStatus);
    return { success: true };
  } catch (err) {
    return { success: false, error: 'Unable to record response.', details: String(err && err.message ? err.message : err) };
  }
}

/** Small read-only helper for dashboards/landing pages: current follow-up counts. */
function generateFollowUpMetrics() {
  var settings = getSettings();
  var leads = getFollowUpLeads();
  var overdue = leads.filter(function (l) { return (l.daysSinceContact || 0) >= (settings.followUpThresholdDays + 2); }).length;
  var avg = leads.length
    ? Math.round((leads.reduce(function (s, l) { return s + (l.daysSinceContact || 0); }, 0) / leads.length) * 10) / 10
    : 0;
  return { due: leads.length, overdue: overdue, averageDaysSinceContact: avg };
}
