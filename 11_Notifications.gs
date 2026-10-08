/**
 * ============================================================
 * 11_Notifications.gs
 * Email is the reliable default. Everything goes through
 * sendNotification(type, data) so another channel (an approved
 * WhatsApp Business API, Slack, etc.) can be added later by adding
 * one more case here - nothing else in the project needs to change.
 * ============================================================
 */

function sendNotification(type, data) {
  if (!data || !data.recipients || !data.recipients.length) {
    return { success: false, error: 'No recipients configured.' };
  }
  try {
    switch (type) {
      case 'follow-up':
      case 'pm-summary':
        MailApp.sendEmail({ to: data.recipients.join(','), subject: data.subject, body: data.body });
        return { success: true, sentTo: data.recipients.length };
      default:
        return { success: false, error: 'Unknown notification type: ' + type };
    }
  } catch (err) {
    return { success: false, error: 'Email send failed.', details: String(err && err.message ? err.message : err) };
  }
}

/** Sends nothing when there are zero follow-ups due. One email per assigned person, falling back to the default notification email. */
function sendFollowUpNotifications() {
  var settings = getSettings();
  if (!settings.followUpNotificationsEnabled) return { success: true, skipped: true, reason: 'Follow-up notifications disabled in Settings.' };

  var leads = getFollowUpLeads();
  if (!leads.length) return { success: true, sent: 0, reason: 'No follow-ups due.' };

  var byAssignee = {};
  leads.forEach(function (l) {
    var who = l.assignedTo || settings.defaultNotificationEmail;
    if (!byAssignee[who]) byAssignee[who] = [];
    byAssignee[who].push(l);
  });

  var totalSent = 0;
  Object.keys(byAssignee).forEach(function (assignee) {
    var recipient = assignee.indexOf('@') !== -1 ? assignee : settings.defaultNotificationEmail;
    if (!recipient) return;
    var leadsForPerson = byAssignee[assignee];

    var subject = 'Follow-Up Required — ' + leadsForPerson.length + ' Lead' + (leadsForPerson.length === 1 ? '' : 's') + ' Need' + (leadsForPerson.length === 1 ? 's' : '') + ' Attention';
    var lines = ['These leads have been waiting 3+ days without a response.', ''];
    leadsForPerson.forEach(function (l) {
      lines.push('• ' + l.businessName + ' (' + l.leadId + ') — Contact: ' + (l.contactName || 'n/a') +
        ' — Last contacted: ' + (l.lastContactDate || 'unknown') + ' (' + l.daysSinceContact + ' days ago)' +
        ' — Channel: ' + (l.channel || 'n/a') + ' — Assigned: ' + (l.assignedTo || 'unassigned') +
        (l.nextAction ? ' — Suggested next action: ' + l.nextAction : ''));
    });
    lines.push('', 'Open the Lead Manager sidebar (View Follow-Ups) to update these.');

    var result = sendNotification('follow-up', { recipients: [recipient], subject: subject, body: lines.join('\n') });
    if (result.success) totalSent++;
  });

  return { success: true, sent: totalSent };
}
