/**
 * ============================================================
 * 10_DailySummary.gs
 * Daily team metrics + the PM summary email. Reuses the existing
 * Leads sheet, Duplicate Review sheet, and Audit Log - no separate
 * "automation log" sheet is created since Audit Log already covers
 * that role in this app.
 *
 * NOTE on "today" for status-based counts (Replies, Interested,
 * Meetings, Converted): there's no per-status-change timestamp, so
 * these are approximated using Last Contact Date / Date Contacted
 * falling on today. If you want exact same-day attribution, add a
 * "Status Changed Date" column - this file is the only place that
 * would need to change.
 * ============================================================
 */

function isSameCalendarDay_(a, b) {
  if (!a || !b) return false;
  var da = (a instanceof Date) ? a : new Date(a);
  var db = (b instanceof Date) ? b : new Date(b);
  if (isNaN(da.getTime()) || isNaN(db.getTime())) return false;
  return da.getFullYear() === db.getFullYear() && da.getMonth() === db.getMonth() && da.getDate() === db.getDate();
}

function generateDailyMetrics() {
  var sh = getOrCreateSheet_(SHEET_LEADS);
  var lastRow = sh.getLastRow();
  var today = new Date();
  var metrics = {
    newLeads: 0, contacted: 0, followUpsCompleted: 0, followUpsDue: 0,
    repliesReceived: 0, interested: 0, meetingsScheduled: 0, converted: 0, duplicatesDetected: 0
  };
  if (lastRow < 2) return metrics;

  var values = sh.getRange(2, 1, lastRow - 1, LEADS_HEADERS.length).getValues();
  values.forEach(function (row) {
    if (!row[LEAD_COL.LEAD_ID - 1]) return;

    if (isSameCalendarDay_(row[LEAD_COL.DATE_ADDED - 1], today)) metrics.newLeads++;
    if (isSameCalendarDay_(row[LEAD_COL.DATE_CONTACTED - 1], today)) metrics.contacted++;
    if (isSameCalendarDay_(row[LEAD_COL.LAST_FOLLOWUP_DATE - 1], today)) metrics.followUpsCompleted++;
    if (row[LEAD_COL.FOLLOWUP_STATUS - 1] === 'FOLLOW UP DUE') metrics.followUpsDue++;

    var lastContact = row[LEAD_COL.LAST_CONTACT_DATE - 1];
    var statusToday = isSameCalendarDay_(lastContact, today) || isSameCalendarDay_(row[LEAD_COL.DATE_CONTACTED - 1], today);
    if (row[LEAD_COL.RESPONSE_STATUS - 1] === 'Replied' && isSameCalendarDay_(lastContact, today)) metrics.repliesReceived++;

    var status = row[LEAD_COL.CONTACT_STATUS - 1];
    if (statusToday && status === 'Interested') metrics.interested++;
    if (statusToday && status === 'Meeting Scheduled') metrics.meetingsScheduled++;
    if (statusToday && status === 'Converted') metrics.converted++;
  });

  var reviewSh = getOrCreateSheet_(SHEET_REVIEW);
  var reviewLastRow = reviewSh.getLastRow();
  if (reviewLastRow >= 2) {
    var reviewRows = reviewSh.getRange(2, 9, reviewLastRow - 1, 1).getValues(); // 'Detected At' column
    reviewRows.forEach(function (r) { if (isSameCalendarDay_(r[0], today)) metrics.duplicatesDetected++; });
  }

  return metrics;
}

/** Per-team-member breakdown for today. Grouped by Assigned To, falling back to Added By for the "new leads" count. */
function generateTeamMetrics() {
  var sh = getOrCreateSheet_(SHEET_LEADS);
  var lastRow = sh.getLastRow();
  var today = new Date();
  var byPerson = {};
  function bucket(name) {
    var key = name || '(unassigned)';
    if (!byPerson[key]) byPerson[key] = { member: key, newLeads: 0, contacted: 0, followUps: 0, replies: 0, interested: 0, meetings: 0, converted: 0 };
    return byPerson[key];
  }
  if (lastRow < 2) return [];

  var values = sh.getRange(2, 1, lastRow - 1, LEADS_HEADERS.length).getValues();
  values.forEach(function (row) {
    if (!row[LEAD_COL.LEAD_ID - 1]) return;
    var addedBy = row[LEAD_COL.ADDED_BY - 1];
    var assignedTo = row[LEAD_COL.ASSIGNED_TO - 1] || addedBy;

    if (isSameCalendarDay_(row[LEAD_COL.DATE_ADDED - 1], today) && addedBy) bucket(addedBy).newLeads++;
    if (isSameCalendarDay_(row[LEAD_COL.DATE_CONTACTED - 1], today) && assignedTo) bucket(assignedTo).contacted++;
    if (isSameCalendarDay_(row[LEAD_COL.LAST_FOLLOWUP_DATE - 1], today) && assignedTo) bucket(assignedTo).followUps++;

    var lastContact = row[LEAD_COL.LAST_CONTACT_DATE - 1];
    if (row[LEAD_COL.RESPONSE_STATUS - 1] === 'Replied' && isSameCalendarDay_(lastContact, today) && assignedTo) bucket(assignedTo).replies++;

    var statusToday = isSameCalendarDay_(lastContact, today) || isSameCalendarDay_(row[LEAD_COL.DATE_CONTACTED - 1], today);
    var status = row[LEAD_COL.CONTACT_STATUS - 1];
    if (statusToday && assignedTo) {
      if (status === 'Interested') bucket(assignedTo).interested++;
      if (status === 'Meeting Scheduled') bucket(assignedTo).meetings++;
      if (status === 'Converted') bucket(assignedTo).converted++;
    }
  });

  return Object.keys(byPerson).map(function (k) { return byPerson[k]; }).sort(function (a, b) { return b.newLeads - a.newLeads; });
}

/** Factual "Attention Needed" bullets - never critical of individuals. */
function buildAttentionNeeded_(daily, team, followUp) {
  var bullets = [];

  var dueCounts = {};
  getFollowUpLeads().forEach(function (l) {
    var who = l.assignedTo || '(unassigned)';
    dueCounts[who] = (dueCounts[who] || 0) + 1;
  });
  Object.keys(dueCounts).forEach(function (who) {
    if (dueCounts[who] >= 5) bullets.push(who + ' has ' + dueCounts[who] + ' follow-ups currently due.');
  });

  if (followUp.overdue > 0) {
    bullets.push(followUp.overdue + ' follow-up' + (followUp.overdue === 1 ? ' is' : 's are') + ' significantly overdue.');
  }
  if (daily.duplicatesDetected >= 3) {
    bullets.push(daily.duplicatesDetected + ' possible duplicate leads were detected today - worth a quick review.');
  }
  if (team.length) {
    var top = team.reduce(function (best, m) { return (m.newLeads > (best ? best.newLeads : -1)) ? m : best; }, null);
    if (top && top.newLeads > 0) bullets.push(top.member + ' added the highest number of new leads today (' + top.newLeads + ').');
  }
  return bullets;
}

function sendDailyPMSummary() {
  var settings = getSettings();
  if (!settings.dailySummaryEnabled) return { success: true, skipped: true, reason: 'Daily summary disabled in Settings.' };

  var cache = CacheService.getScriptCache();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var todayKey = 'PM_SUMMARY_SENT_' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (cache.get(todayKey)) return { success: true, skipped: true, reason: 'Already sent today.' };

  var daily = generateDailyMetrics();
  var team = generateTeamMetrics();
  var followUp = generateFollowUpMetrics();
  var responseRate = daily.contacted > 0 ? Math.round((daily.repliesReceived / daily.contacted) * 1000) / 10 : 0;

  var lines = [];
  lines.push('Daily Lead Tracker Summary');
  lines.push('Date: ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));
  lines.push('');
  lines.push('Team Numbers');
  lines.push('New Leads: ' + daily.newLeads);
  lines.push('Contacted: ' + daily.contacted);
  lines.push('Follow-Ups Completed: ' + daily.followUpsCompleted);
  lines.push('Follow-Ups Due: ' + daily.followUpsDue);
  lines.push('Replies: ' + daily.repliesReceived);
  lines.push('Interested: ' + daily.interested);
  lines.push('Meetings: ' + daily.meetingsScheduled);
  lines.push('Conversions: ' + daily.converted);
  lines.push('Duplicates Detected: ' + daily.duplicatesDetected);
  lines.push('');
  lines.push('Team Performance');
  lines.push('Team Member | New Leads | Contacted | Follow-Ups | Replies | Interested | Meetings | Converted');
  team.forEach(function (m) {
    lines.push(m.member + ' | ' + m.newLeads + ' | ' + m.contacted + ' | ' + m.followUps + ' | ' + m.replies + ' | ' + m.interested + ' | ' + m.meetings + ' | ' + m.converted);
  });
  lines.push('');
  lines.push('Team activity ' + (daily.newLeads >= 10 ? 'was strong' : 'was moderate') + ' today with ' + daily.newLeads +
    ' new leads added and ' + daily.contacted + ' leads contacted. ' + followUp.due + ' follow-up' + (followUp.due === 1 ? ' is' : 's are') +
    ' currently due. Response rate is ' + responseRate + '%.');
  lines.push('');
  lines.push('Attention Needed');
  var attention = buildAttentionNeeded_(daily, team, followUp);
  if (attention.length) attention.forEach(function (a) { lines.push('- ' + a); });
  else lines.push('- Nothing unusual to flag today.');

  var recipients = settings.pmEmails.length ? settings.pmEmails : (settings.defaultNotificationEmail ? [settings.defaultNotificationEmail] : []);
  if (!recipients.length) {
    logAudit('Daily PM Summary', '', '', '', '', 'Not sent: no PM emails configured in Settings.');
    return { success: false, error: 'No PM emails configured in Settings. Add them under Settings → "PM Emails (comma-separated)".' };
  }

  var result = sendNotification('pm-summary', {
    recipients: recipients,
    subject: 'Daily Lead Tracker Summary — ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    body: lines.join('\n')
  });

  if (result.success) {
    cache.put(todayKey, '1', 21600); // 6h guard against duplicate same-day sends, short enough not to block tomorrow
    logAudit('Daily PM Summary', '', '', '', '', 'Sent to ' + recipients.join(', ') + '. ' + daily.newLeads + ' new leads, ' + daily.followUpsDue + ' follow-ups due.');
  } else {
    logAudit('Daily PM Summary', '', '', '', '', 'Send failed: ' + (result.error || 'unknown error'));
  }
  return result;
}
