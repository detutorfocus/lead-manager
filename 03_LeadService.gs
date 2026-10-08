/**
 * ============================================================
 * 03_LeadService.gs
 * Saving leads, merging, "keep both", "create anyway",
 * fetching a lead by ID, and Duplicate Review bookkeeping.
 * ============================================================
 */

var REQUIRED_FIELDS = ['businessName', 'sector', 'city', 'observation', 'source'];

/** Server-side validation. Never trust the client. */
function validateLeadData_(data) {
  var missing = REQUIRED_FIELDS.filter(function (f) { return !data[f] || !String(data[f]).trim(); });
  if (missing.length) {
    return { valid: false, error: 'Missing required field(s): ' + missing.join(', ') };
  }
  var hasChannel = data.instagram || data.linkedin || data.email || data.whatsapp || data.website;
  var warning = hasChannel ? null : 'No contact channel provided (Instagram, LinkedIn, Email, WhatsApp, or Website). Consider adding one before saving.';
  var obs = String(data.observation || '').trim();
  var obsWarning = null;
  if (obs.length < 25 || /^(good|nice|great|okay|ok)\s*(business)?\.?$/i.test(obs)) {
    obsWarning = 'Please provide a specific observation that explains something real about this business.';
  }
  return { valid: true, channelWarning: warning, observationWarning: obsWarning };
}

/** Build the full row array for a lead, including normalized and follow-up columns. */
function buildLeadRow_(leadId, data, recordStatus) {
  var settings = getSettings();
  var cc = settings.phoneCountryCode;
  var clean = sanitizeLeadData_(data);
  var now = new Date();

  return [
    leadId,
    clean.businessName || '',
    clean.sector || '',
    clean.city || '',
    clean.contactName || '',
    clean.contactRole || '',
    clean.instagram || '',
    clean.linkedin || '',
    clean.email || '',
    clean.whatsapp || '',
    clean.website || '',
    clean.websiteStatus || 'None',
    clean.icp || '',
    clean.trigger || '',
    clean.observation || '',
    clean.source || '',
    now,
    clean.addedBy || getCurrentUser_(),
    normalizeBusinessName(clean.businessName),
    normalizeInstagram(clean.instagram),
    normalizeLinkedIn(clean.linkedin),
    normalizeWebsite(clean.website),
    normalizeEmail(clean.email),
    normalizePhone(clean.whatsapp, cc),
    clean.dupStatus || 'Not Checked',
    clean.dupScore || 0,
    clean.matchedLeadId || '',
    now,
    recordStatus || 'Active',
    clean.branch || '',
    // Follow-up tracking - a brand-new lead always starts here regardless
    // of how it was saved (clean save, Keep Both, or Create Anyway).
    'Not Contacted', // Contact Status
    '', // Date Contacted
    '', // Last Contact Date
    '', // Follow-Up Due
    0,  // Follow-Up Count
    '', // Last Follow-Up Date
    '', // Response Status
    '', // Next Action
    clean.assignedTo || '', // Assigned To
    'NOT DUE' // Follow-Up Status
  ];
}

/**
 * Save a brand-new lead. This is the single write path used by
 * "Create Anyway", "Keep Both", and a clean (no-match) save.
 */
function saveLead(leadData) {
  try {
    var validation = validateLeadData_(leadData);
    if (!validation.valid) {
      return { success: false, error: validation.error };
    }
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    var leadId;
    try {
      leadId = generateLeadId_();
      var sh = getOrCreateSheet_(SHEET_LEADS);
      var row = buildLeadRow_(leadId, leadData, 'Active');
      sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    } finally {
      lock.releaseLock();
    }

    logAudit('Lead Created', leadId, '', leadData.businessName, '', 'Source: ' + (leadData.source || ''));

    return {
      success: true,
      leadId: leadId,
      channelWarning: validation.channelWarning,
      observationWarning: validation.observationWarning
    };
  } catch (err) {
    return { success: false, error: 'Unable to save lead.', details: String(err && err.message ? err.message : err) };
  }
}

/** Fetch a full lead record by ID as a plain object, or null. */
function getLeadById(leadId) {
  var sh = getOrCreateSheet_(SHEET_LEADS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var values = sh.getRange(2, 1, lastRow - 1, LEADS_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][LEAD_COL.LEAD_ID - 1] === leadId) {
      var r = values[i];
      var obj = { rowIndex: i + 2 };
      LEADS_HEADERS.forEach(function (h, idx) { obj[h] = r[idx]; });
      return obj;
    }
  }
  return null;
}

/** Find (if any) an existing, still-open Duplicate Review row for this exact pair. */
function findExistingReview_(newLeadId, existingLeadId) {
  var sh = getOrCreateSheet_(SHEET_REVIEW);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var values = sh.getRange(2, 1, lastRow - 1, REVIEW_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row[1] === newLeadId && row[2] === existingLeadId) {
      return { rowIndex: i + 2, values: row };
    }
  }
  return null;
}

/** Create (or skip if one already exists) a Duplicate Review record for a match. */
function createReviewRecord_(newLeadId, newBusiness, match) {
  var existing = findExistingReview_(newLeadId, match.leadId);
  if (existing) return; // don't spam repeat rows for the same pair

  var sh = getOrCreateSheet_(SHEET_REVIEW);
  var reviewId = 'RV-' + Utilities.getUuid().slice(0, 8);
  var status = match.score >= 90 ? 'Confirmed Duplicate'
    : match.score >= 70 ? 'Likely Duplicate'
    : 'Possible Duplicate';

  sh.appendRow([
    reviewId, newLeadId, match.leadId, newBusiness, match.businessName,
    match.score, status, match.matchingFields.join(', '), new Date(),
    getCurrentUser_(), '', '', ''
  ]);

  logAudit('Duplicate Detected', newLeadId, match.leadId, newBusiness, match.score,
    'Matching: ' + match.matchingFields.join(', '));
}

/** Update the Decision / Decision Date / Decision By / Status columns on a review row. */
function updateReviewDecision_(newLeadId, existingLeadId, status, decision) {
  var found = findExistingReview_(newLeadId, existingLeadId);
  if (!found) return;
  var sh = getOrCreateSheet_(SHEET_REVIEW);
  sh.getRange(found.rowIndex, 7).setValue(status);   // Status
  sh.getRange(found.rowIndex, 11).setValue(decision); // Decision
  sh.getRange(found.rowIndex, 12).setValue(new Date()); // Decision Date
  sh.getRange(found.rowIndex, 13).setValue(getCurrentUser_()); // Decision By
}

/**
 * "Create Anyway": save the new lead despite a high-confidence match,
 * recording the override decision against each flagged match.
 */
function createAnyway(leadData, matches) {
  var result = saveLead(leadData);
  if (!result.success) return result;

  (matches || []).forEach(function (m) {
    createReviewRecord_(result.leadId, leadData.businessName, m);
    updateReviewDecision_(result.leadId, m.leadId, 'Reviewed', 'Create Anyway');
  });

  logAudit('Create Anyway', result.leadId, (matches || []).map(function (m) { return m.leadId; }).join(', '),
    leadData.businessName, matches && matches[0] ? matches[0].score : '', 'User created despite duplicate warning.');

  return result;
}

/**
 * "Keep Both": the user has judged these are legitimate separate
 * businesses (e.g. different branches). Saves the new lead and
 * dismisses the review for this pair.
 */
function keepBoth(leadData, existingLeadId) {
  var result = saveLead(leadData);
  if (!result.success) return result;

  if (existingLeadId) {
    createReviewRecord_(result.leadId, leadData.businessName, {
      leadId: existingLeadId,
      businessName: (getLeadById(existingLeadId) || {})['Business Name'] || '',
      score: 0,
      matchingFields: []
    });
    updateReviewDecision_(result.leadId, existingLeadId, 'Dismissed', 'Keep Both');
  }

  logAudit('Duplicate Dismissed', result.leadId, existingLeadId, leadData.businessName, '', 'Kept as separate record (Keep Both).');
  return result;
}

/**
 * Merge the incoming (not-yet-saved) lead into an existing lead record.
 * fieldChoices: { fieldKey: 'new' | 'existing' | 'both' }
 * The existing Lead ID survives; nothing is silently destroyed — both
 * original values are preserved in the Audit Log details.
 */
function mergeLeads(leadData, existingLeadId, fieldChoices) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var sh = getOrCreateSheet_(SHEET_LEADS);
      var existing = getLeadById(existingLeadId);
      if (!existing) {
        return { success: false, error: 'Existing lead not found: ' + existingLeadId };
      }

      var fieldMap = {
        businessName: LEAD_COL.BUSINESS_NAME, sector: LEAD_COL.SECTOR, city: LEAD_COL.CITY,
        contactName: LEAD_COL.CONTACT_NAME, contactRole: LEAD_COL.CONTACT_ROLE,
        instagram: LEAD_COL.INSTAGRAM, linkedin: LEAD_COL.LINKEDIN, email: LEAD_COL.EMAIL,
        whatsapp: LEAD_COL.WHATSAPP, website: LEAD_COL.WEBSITE, websiteStatus: LEAD_COL.WEBSITE_STATUS,
        icp: LEAD_COL.ICP, trigger: LEAD_COL.TRIGGER, branch: LEAD_COL.BRANCH
      };

      var changes = [];
      Object.keys(fieldMap).forEach(function (key) {
        var choice = (fieldChoices && fieldChoices[key]) || 'existing';
        var newVal = leadData[key];
        var existingVal = existing[LEADS_HEADERS[fieldMap[key] - 1]];
        var finalVal = existingVal;
        // Recommend the populated value when one side is blank.
        if (choice === 'new' || (!existingVal && newVal)) finalVal = newVal;
        else if (choice === 'existing') finalVal = existingVal;
        if (finalVal !== existingVal) {
          sh.getRange(existing.rowIndex, fieldMap[key]).setValue(sanitizeValue_(finalVal));
          changes.push(key + ': "' + existingVal + '" -> "' + finalVal + '"');
        }
      });

      // Observation and Source: preserve both rather than deleting either.
      var obsChoice = (fieldChoices && fieldChoices.observation) || 'both';
      var existingObs = existing['Observation'] || '';
      var newObs = leadData.observation || '';
      var finalObs = existingObs;
      if (obsChoice === 'new') finalObs = newObs;
      else if (obsChoice === 'both' && newObs && newObs !== existingObs) {
        finalObs = existingObs ? (existingObs + '\n---\n[Merged ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + ']: ' + newObs) : newObs;
      }
      if (finalObs !== existingObs) {
        sh.getRange(existing.rowIndex, LEAD_COL.OBSERVATION).setValue(sanitizeValue_(finalObs));
      }

      var existingSource = existing['Source'] || '';
      var newSource = leadData.source || '';
      if (newSource && newSource !== existingSource) {
        var mergedSource = existingSource ? (existingSource + '; ' + newSource) : newSource;
        sh.getRange(existing.rowIndex, LEAD_COL.SOURCE).setValue(sanitizeValue_(mergedSource));
      }

      sh.getRange(existing.rowIndex, LEAD_COL.LAST_DUP_CHECK).setValue(new Date());

      logAudit('Lead Merged', existingLeadId, '', existing['Business Name'], '',
        'Merged incoming lead "' + leadData.businessName + '" into ' + existingLeadId +
        (changes.length ? '. Changes: ' + changes.join(' | ') : '. No field changes.'));

      if (leadData.tempReviewNewLeadId) {
        updateReviewDecision_(leadData.tempReviewNewLeadId, existingLeadId, 'Merged', 'Merge');
      }

      return { success: true, leadId: existingLeadId, message: 'These leads have been merged successfully.' };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return { success: false, error: 'Unable to merge leads.', details: String(err && err.message ? err.message : err) };
  }
}

/**
 * Merge two ALREADY-SAVED leads (used from Review Duplicates / bulk scan
 * results, as opposed to the pre-save merge flow in Add New Lead).
 * The existing/older leadId (`survivorId`) survives; `loserId`'s row is
 * marked Record Status = "Merged" and its data folded into the survivor.
 */
function mergeExistingLeads(loserId, survivorId) {
  var loser = getLeadById(loserId);
  if (!loser) return { success: false, error: 'Lead not found: ' + loserId };

  var leadData = {
    businessName: loser['Business Name'], sector: loser['Sector'], city: loser['City'],
    contactName: loser['Contact Name'], contactRole: loser['Contact Role'],
    instagram: loser['Instagram Handle'], linkedin: loser['LinkedIn URL'], email: loser['Business Email'],
    whatsapp: loser['WhatsApp Business Line'], website: loser['Website'], websiteStatus: loser['Website Status'],
    icp: loser['ICP'], trigger: loser['Trigger'], observation: loser['Observation'], source: loser['Source'],
    branch: loser['Branch / Location'], tempReviewNewLeadId: loserId
  };

  var result = mergeLeads(leadData, survivorId, {});
  if (!result.success) return result;

  var sh = getOrCreateSheet_(SHEET_LEADS);
  sh.getRange(loser.rowIndex, LEAD_COL.RECORD_STATUS).setValue('Merged');
  sh.getRange(loser.rowIndex, LEAD_COL.MATCHED_LEAD_ID).setValue(survivorId);

  return result;
}

/** List open (unresolved) Duplicate Review rows, most recent first. */
function getOpenReviews() {
  var sh = getOrCreateSheet_(SHEET_REVIEW);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, REVIEW_HEADERS.length).getValues();
  var openStatuses = ['Confirmed Duplicate', 'Likely Duplicate', 'Possible Duplicate'];
  var open = values.filter(function (r) { return openStatuses.indexOf(r[6]) !== -1; });
  open.sort(function (a, b) { return new Date(b[8]) - new Date(a[8]); });
  return open.map(function (r) {
    return {
      reviewId: r[0], newLeadId: r[1], existingLeadId: r[2], newBusiness: r[3],
      existingBusiness: r[4], score: r[5], status: r[6], matchingFields: r[7]
    };
  });
}

/** Resolve an open review from the Review Duplicates sidebar. */
function resolveReview(newLeadId, existingLeadId, decision) {
  try {
    if (decision === 'Merge') {
      var res = mergeExistingLeads(newLeadId, existingLeadId);
      if (!res.success) return res;
      updateReviewDecision_(newLeadId, existingLeadId, 'Merged', 'Merge');
      return { success: true, message: 'Merged ' + newLeadId + ' into ' + existingLeadId + '.' };
    }
    if (decision === 'Keep Both' || decision === 'Dismiss') {
      updateReviewDecision_(newLeadId, existingLeadId, 'Dismissed', decision);
      logAudit('Duplicate Dismissed', newLeadId, existingLeadId, '', '', 'Resolved via Review Duplicates: ' + decision);
      return { success: true, message: 'Marked as not a duplicate.' };
    }
    updateReviewDecision_(newLeadId, existingLeadId, 'Reviewed', decision || 'Reviewed');
    logAudit('Duplicate Reviewed', newLeadId, existingLeadId, '', '', decision || 'Reviewed');
    return { success: true, message: 'Review updated.' };
  } catch (err) {
    return { success: false, error: 'Unable to resolve review.', details: String(err && err.message ? err.message : err) };
  }
}

/** Archive a lead (soft-delete): Record Status -> Archived. Never deletes data. */
function archiveLead(leadId) {
  var lead = getLeadById(leadId);
  if (!lead) return { success: false, error: 'Lead not found.' };
  var sh = getOrCreateSheet_(SHEET_LEADS);
  sh.getRange(lead.rowIndex, LEAD_COL.RECORD_STATUS).setValue('Archived');
  logAudit('Lead Archived', leadId, '', lead['Business Name'], '', '');
  return { success: true };
}
