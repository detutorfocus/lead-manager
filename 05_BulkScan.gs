/**
 * ============================================================
 * 05_BulkScan.gs
 * Bulk duplicate scan across the whole Leads sheet, for cleaning
 * up existing data. Runs in client-driven chunks so a real
 * progress indicator can be shown ("Scanning leads: 245 / 1,000").
 * Never modifies original lead field values — only the
 * Duplicate Status / Score / Matched Lead ID / Last Duplicate
 * Check columns, plus new Duplicate Review rows.
 * ============================================================
 */

/** How many leads the client should ask us to process per call. */
var BULK_SCAN_BATCH_SIZE = 25;

/** Total lead count, for the client to compute a progress bar. */
function getBulkScanTotal() {
  var sh = getOrCreateSheet_(SHEET_LEADS);
  var lastRow = sh.getLastRow();
  return Math.max(0, lastRow - 1);
}

/**
 * Small read-only helper for the Bulk Scan screen's summary card:
 * current lead count (reuses getBulkScanTotal(), not recomputed) plus
 * when the last scan finished, if any. Purely additive - nothing
 * about the scan logic below changes.
 */
function getBulkScanInfo() {
  var props = PropertiesService.getDocumentProperties();
  return {
    totalLeads: getBulkScanTotal(),
    lastScanTime: props.getProperty('LAST_BULK_SCAN_TIME') || null
  };
}

/**
 * Process one batch of the bulk scan, starting at (0-indexed) row offset
 * `startIndex` within the lead list. Returns updated counts and the next
 * startIndex to call with (or null when done).
 */
function scanLeadsBatch(startIndex) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = getOrCreateSheet_(SHEET_LEADS);
    var lastRow = sh.getLastRow();
    var total = Math.max(0, lastRow - 1);
    if (total === 0) return { success: true, done: true, processed: 0, total: 0, summary: emptyScanSummary_() };

    var settings = getSettings();
    var fullIndex = loadLeadIndex_(); // includes normalized fields already

    var endIndex = Math.min(startIndex + BULK_SCAN_BATCH_SIZE, fullIndex.length);
    var possible = 0, likely = 0, confirmed = 0;

    for (var i = startIndex; i < endIndex; i++) {
      var lead = fullIndex[i];
      if (!lead) continue;

      var leadData = {
        businessName: lead.businessName, sector: lead.sector, city: lead.city,
        instagram: '', linkedin: '', email: '', whatsapp: '', website: '',
        excludeLeadId: lead.leadId
      };
      // Reuse already-normalized values directly rather than re-normalizing raw fields.
      var norm = {
        rawName: lead.businessName, name: lead.normName, city: normalizeCity(lead.city),
        sector: normalizeSector(lead.sector), instagram: lead.normInstagram,
        linkedin: lead.normLinkedIn, website: lead.normWebsite, email: lead.normEmail, phone: lead.normPhone
      };

      var others = fullIndex.filter(function (l) { return l.leadId !== lead.leadId; });
      var candidates = findCandidates_(norm, others);
      var scored = candidates.map(function (c) {
        var r = scoreCandidate_(norm, c, settings.weights);
        return { leadId: c.leadId, businessName: c.businessName, score: r.score, matchingFields: r.matchingFields, strongMatch: r.strongMatch };
      });
      scored.forEach(function (m) { if (m.strongMatch && m.score < settings.thresholds.likely) m.score = settings.thresholds.likely; });
      scored.sort(function (a, b) { return b.score - a.score; });
      var top = scored[0];

      var status = top ? classifyScore_(top.score, settings.thresholds) : 'NO_STRONG_MATCH';
      var displayStatus = status === 'CONFIRMED_DUPLICATE' ? 'Confirmed Duplicate'
        : status === 'LIKELY_DUPLICATE' ? 'Likely Duplicate'
        : status === 'POSSIBLE_DUPLICATE' ? 'Possible Duplicate' : 'No Strong Match';

      sh.getRange(lead.rowIndex, LEAD_COL.DUP_STATUS).setValue(displayStatus);
      sh.getRange(lead.rowIndex, LEAD_COL.DUP_SCORE).setValue(top ? top.score : 0);
      sh.getRange(lead.rowIndex, LEAD_COL.MATCHED_LEAD_ID).setValue(top ? top.leadId : '');
      sh.getRange(lead.rowIndex, LEAD_COL.LAST_DUP_CHECK).setValue(new Date());

      if (top && top.score >= settings.thresholds.possible) {
        createReviewRecord_(lead.leadId, lead.businessName, top);
        if (status === 'POSSIBLE_DUPLICATE') possible++;
        else if (status === 'LIKELY_DUPLICATE') likely++;
        else if (status === 'CONFIRMED_DUPLICATE') confirmed++;
      }
    }

    var done = endIndex >= fullIndex.length;
    if (done) {
      logAudit('System Setup', '', '', '', '', 'Bulk duplicate scan completed: ' + fullIndex.length + ' leads checked.');
      PropertiesService.getDocumentProperties().setProperty('LAST_BULK_SCAN_TIME', new Date().toISOString());
    }

    return {
      success: true,
      done: done,
      processed: endIndex,
      total: fullIndex.length,
      nextIndex: done ? null : endIndex,
      batchCounts: { possible: possible, likely: likely, confirmed: confirmed }
    };
  } catch (err) {
    return { success: false, error: 'Bulk scan failed.', details: String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

function emptyScanSummary_() {
  return { total: 0, possible: 0, likely: 0, confirmed: 0 };
}
