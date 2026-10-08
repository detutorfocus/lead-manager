/**
 * ============================================================
 * 06_Dashboard.gs
 * Computes summary statistics and writes them onto the
 * Dashboard sheet.
 * ============================================================
 */

function refreshDashboard() {
  var leadsSh = getOrCreateSheet_(SHEET_LEADS);
  var reviewSh = getOrCreateSheet_(SHEET_REVIEW);
  var dashSh = getOrCreateSheet_(SHEET_DASHBOARD);

  var lastRow = leadsSh.getLastRow();
  var leads = lastRow >= 2 ? leadsSh.getRange(2, 1, lastRow - 1, LEADS_HEADERS.length).getValues() : [];

  var totalLeads = 0, activeLeads = 0, cleanLeads = 0;
  var today = 0, thisWeek = 0, thisMonth = 0;
  var now = new Date();
  var startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var startOfWeek = new Date(startOfDay); startOfWeek.setDate(startOfDay.getDate() - startOfDay.getDay());
  var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  leads.forEach(function (r) {
    if (!r[LEAD_COL.LEAD_ID - 1]) return;
    totalLeads++;
    var recordStatus = r[LEAD_COL.RECORD_STATUS - 1] || 'Active';
    if (recordStatus === 'Active') activeLeads++;
    var dupStatus = r[LEAD_COL.DUP_STATUS - 1];
    if (!dupStatus || dupStatus === 'No Strong Match' || dupStatus === 'Not Checked') cleanLeads++;

    var added = r[LEAD_COL.DATE_ADDED - 1];
    if (added instanceof Date) {
      if (added >= startOfDay) today++;
      if (added >= startOfWeek) thisWeek++;
      if (added >= startOfMonth) thisMonth++;
    }
  });

  var reviewLastRow = reviewSh.getLastRow();
  var reviews = reviewLastRow >= 2 ? reviewSh.getRange(2, 1, reviewLastRow - 1, REVIEW_HEADERS.length).getValues() : [];
  var confirmed = 0, likely = 0, possible = 0, needsReview = 0;
  var recent = [];
  reviews.forEach(function (r) {
    var status = r[6];
    if (status === 'Confirmed Duplicate') confirmed++;
    else if (status === 'Likely Duplicate') likely++;
    else if (status === 'Possible Duplicate') possible++;
    if (status === 'Confirmed Duplicate' || status === 'Likely Duplicate' || status === 'Possible Duplicate') needsReview++;
    recent.push(r);
  });
  recent.sort(function (a, b) { return new Date(b[8]) - new Date(a[8]); });
  recent = recent.slice(0, 10);

  var dupRate = totalLeads > 0 ? Math.round((needsReview / totalLeads) * 1000) / 10 : 0;

  dashSh.clear();
  dashSh.getRange(1, 1).setValue('Lead Management Dashboard').setFontSize(16).setFontWeight('bold');
  dashSh.getRange(2, 1).setValue('Last refreshed: ' + Utilities.formatDate(now, Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd HH:mm'));

  var stats = [
    ['Total Leads', totalLeads],
    ['Active Leads', activeLeads],
    ['Confirmed Duplicates', confirmed],
    ['Likely Duplicates', likely],
    ['Possible Duplicates', possible],
    ['Leads Requiring Review', needsReview],
    ['Clean Leads', cleanLeads],
    ['Duplicate Rate', dupRate + '%'],
    ['Leads Added Today', today],
    ['Leads Added This Week', thisWeek],
    ['Leads Added This Month', thisMonth]
  ];
  dashSh.getRange(4, 1, stats.length, 2).setValues(stats);
  dashSh.getRange(4, 1, stats.length, 1).setFontWeight('bold');

  // Follow-Up Overview - reuses generateFollowUpMetrics() (09_FollowUp.gs), nothing recomputed here.
  var followUp = generateFollowUpMetrics();
  var fuRow = 4 + stats.length + 2;
  dashSh.getRange(fuRow, 1).setValue('Follow-Up Overview').setFontWeight('bold').setFontSize(12);
  var fuStats = [
    ['Follow-Ups Due', followUp.due],
    ['Overdue', followUp.overdue],
    ['Average Days Since Contact', followUp.averageDaysSinceContact]
  ];
  dashSh.getRange(fuRow + 1, 1, fuStats.length, 2).setValues(fuStats);
  dashSh.getRange(fuRow + 1, 1, fuStats.length, 1).setFontWeight('bold');

  var headerRow = fuRow + 1 + fuStats.length + 2;
  dashSh.getRange(headerRow, 1).setValue('Recent Duplicate Alerts').setFontWeight('bold').setFontSize(12);
  var alertHeaders = ['Business', 'Existing Match', 'Score', 'Status', 'Action'];
  dashSh.getRange(headerRow + 1, 1, 1, alertHeaders.length).setValues([alertHeaders]).setFontWeight('bold');

  if (recent.length) {
    var alertRows = recent.map(function (r) {
      return [r[3], r[4], r[5], r[6], r[10] || '(pending review)'];
    });
    dashSh.getRange(headerRow + 2, 1, alertRows.length, alertHeaders.length).setValues(alertRows);
  }

  dashSh.autoResizeColumns(1, 5);

  // Returned (in addition to writing the sheet) so the Web App landing
  // page can show live stats without recomputing this logic itself.
  // Existing callers (the menu's openDashboard()) already ignore the
  // return value, so this is additive and does not change any existing
  // behavior.
  return {
    success: true,
    totalLeads: totalLeads,
    activeLeads: activeLeads,
    confirmed: confirmed,
    likely: likely,
    possible: possible,
    needsReview: needsReview,
    cleanLeads: cleanLeads,
    dupRate: dupRate,
    today: today,
    thisWeek: thisWeek,
    thisMonth: thisMonth,
    followUpsDue: followUp.due,
    followUpsOverdue: followUp.overdue
  };
}
