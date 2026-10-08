/**
 * ============================================================
 * 02_DuplicateEngine.gs
 * Candidate search + scoring. Loads existing leads into memory
 * once per check, narrows to a small candidate set using exact
 * normalized identifiers / city / first-token, then runs fuzzy
 * scoring only against that candidate set.
 * ============================================================
 */

/** Load all active (non-Archived, non-Merged-away) leads into memory as row objects. */
function loadLeadIndex_() {
  var sh = getOrCreateSheet_(SHEET_LEADS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var numCols = LEADS_HEADERS.length;
  var values = sh.getRange(2, 1, lastRow - 1, numCols).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[LEAD_COL.LEAD_ID - 1]) continue;
    var recordStatus = r[LEAD_COL.RECORD_STATUS - 1] || 'Active';
    if (recordStatus === 'Merged') continue; // superseded record, skip as a match target
    out.push({
      rowIndex: i + 2, // actual sheet row
      leadId: r[LEAD_COL.LEAD_ID - 1],
      businessName: r[LEAD_COL.BUSINESS_NAME - 1],
      city: r[LEAD_COL.CITY - 1],
      sector: r[LEAD_COL.SECTOR - 1],
      normName: r[LEAD_COL.NORM_NAME - 1] || normalizeBusinessName(r[LEAD_COL.BUSINESS_NAME - 1]),
      normInstagram: r[LEAD_COL.NORM_INSTAGRAM - 1] || normalizeInstagram(r[LEAD_COL.INSTAGRAM - 1]),
      normLinkedIn: r[LEAD_COL.NORM_LINKEDIN - 1] || normalizeLinkedIn(r[LEAD_COL.LINKEDIN - 1]),
      normWebsite: r[LEAD_COL.NORM_WEBSITE - 1] || normalizeWebsite(r[LEAD_COL.WEBSITE - 1]),
      normEmail: r[LEAD_COL.NORM_EMAIL - 1] || normalizeEmail(r[LEAD_COL.EMAIL - 1]),
      normPhone: r[LEAD_COL.NORM_PHONE - 1] || normalizePhone(r[LEAD_COL.WHATSAPP - 1], '234'),
      recordStatus: recordStatus
    });
  }
  return out;
}

/** Narrow the full lead index down to a small candidate set for a new lead. */
function findCandidates_(norm, index) {
  var candidates = {}; // leadId -> lead object, dedup
  var firstToken = (norm.name.split(' ')[0] || '');

  index.forEach(function (lead) {
    var isCandidate = false;
    if (norm.instagram && lead.normInstagram && norm.instagram === lead.normInstagram) isCandidate = true;
    if (norm.linkedin && lead.normLinkedIn && norm.linkedin === lead.normLinkedIn) isCandidate = true;
    if (norm.website && lead.normWebsite && websiteDomain_(norm.website) === websiteDomain_(lead.normWebsite)) isCandidate = true;
    if (norm.email && lead.normEmail && norm.email === lead.normEmail) isCandidate = true;
    if (norm.phone && lead.normPhone && norm.phone === lead.normPhone) isCandidate = true;
    if (norm.name && lead.normName && norm.name === lead.normName) isCandidate = true;
    if (norm.city && lead.city && normalizeCity(lead.city) === norm.city) isCandidate = true;
    if (firstToken && lead.normName && lead.normName.split(' ')[0] === firstToken) isCandidate = true;

    if (isCandidate) candidates[lead.leadId] = lead;
  });

  return Object.keys(candidates).map(function (k) { return candidates[k]; });
}

/**
 * Score a single candidate lead against the incoming normalized lead data.
 * Returns { score, matchingFields, strongMatch }
 */
function scoreCandidate_(norm, candidate, weights) {
  var score = 0;
  var matchingFields = [];
  var strongMatch = false;

  if (norm.instagram && candidate.normInstagram && norm.instagram === candidate.normInstagram) {
    score += weights.instagram; matchingFields.push('Instagram'); strongMatch = true;
  }
  if (norm.linkedin && candidate.normLinkedIn && norm.linkedin === candidate.normLinkedIn) {
    score += weights.linkedin; matchingFields.push('LinkedIn'); strongMatch = true;
  }
  if (norm.website && candidate.normWebsite && websiteDomain_(norm.website) === websiteDomain_(candidate.normWebsite)) {
    score += weights.website; matchingFields.push('Website'); strongMatch = true;
  }
  if (norm.email && candidate.normEmail && norm.email === candidate.normEmail) {
    score += weights.email; matchingFields.push('Email'); strongMatch = true;
  }
  if (norm.phone && candidate.normPhone && norm.phone === candidate.normPhone) {
    score += weights.phone; matchingFields.push('Phone/WhatsApp'); strongMatch = true;
  }

  var nameSim = calculateNameSimilarity(norm.rawName, candidate.businessName);
  var nameWeight = weights.name;
  var tierFactor = 0;
  if (nameSim >= 95) tierFactor = 1;
  else if (nameSim >= 90) tierFactor = 30 / 35;
  else if (nameSim >= 80) tierFactor = 20 / 35;
  else if (nameSim >= 70) tierFactor = 10 / 35;
  if (tierFactor > 0) {
    score += nameWeight * tierFactor;
    matchingFields.push('Business Name (' + Math.round(nameSim) + '% similar)');
  }

  if (norm.city && candidate.city && normalizeCity(candidate.city) === norm.city) {
    score += weights.city; matchingFields.push('City');
  }
  if (norm.sector && candidate.sector && normalizeSector(candidate.sector) === norm.sector) {
    score += weights.sector; matchingFields.push('Sector');
  }

  score = Math.min(100, Math.round(score * 10) / 10);
  return { score: score, matchingFields: matchingFields, strongMatch: strongMatch, nameSimilarity: nameSim };
}

/** Classify a score against configured thresholds. */
function classifyScore_(score, thresholds) {
  if (score >= thresholds.confirmed) return 'CONFIRMED_DUPLICATE';
  if (score >= thresholds.likely) return 'LIKELY_DUPLICATE';
  if (score >= thresholds.possible) return 'POSSIBLE_DUPLICATE';
  return 'NO_STRONG_MATCH';
}

/**
 * Main entry point: given raw lead form data, find and score potential
 * duplicates. Returns a structured result (never throws to the caller;
 * errors come back as { success:false, error, details }).
 *
 * leadData fields expected: businessName, sector, city, instagram,
 * linkedin, email, whatsapp, website
 * excludeLeadId (optional): skip this lead when checking existing rows
 * (used for manual re-checks / bulk scans so a row never matches itself).
 */
function checkForDuplicates(leadData) {
  try {
    var settings = getSettings();
    var cc = settings.phoneCountryCode;

    var norm = {
      rawName: leadData.businessName || '',
      name: normalizeBusinessName(leadData.businessName),
      city: normalizeCity(leadData.city),
      sector: normalizeSector(leadData.sector),
      instagram: normalizeInstagram(leadData.instagram),
      linkedin: normalizeLinkedIn(leadData.linkedin),
      website: normalizeWebsite(leadData.website),
      email: normalizeEmail(leadData.email),
      phone: normalizePhone(leadData.whatsapp, cc)
    };

    var index = loadLeadIndex_();
    if (leadData.excludeLeadId) {
      index = index.filter(function (l) { return l.leadId !== leadData.excludeLeadId; });
    }

    var candidates = findCandidates_(norm, index);

    var scored = candidates.map(function (c) {
      var result = scoreCandidate_(norm, c, settings.weights);
      return {
        leadId: c.leadId,
        businessName: c.businessName,
        city: c.city,
        sector: c.sector,
        score: result.score,
        matchingFields: result.matchingFields,
        strongMatch: result.strongMatch,
        nameSimilarity: result.nameSimilarity
      };
    });

    // Strong-identifier override: an exact strong-field match is flagged
    // at least "Likely Duplicate" even if the aggregate score fell short.
    scored.forEach(function (m) {
      if (m.strongMatch && m.score < settings.thresholds.likely) {
        m.score = settings.thresholds.likely;
      }
    });

    scored.sort(function (a, b) { return b.score - a.score; });
    var topMatches = scored.slice(0, 5);

    var topScore = topMatches.length ? topMatches[0].score : 0;
    var status = classifyScore_(topScore, settings.thresholds);

    return {
      success: true,
      status: status,
      score: topScore,
      matches: topMatches,
      confidence: topMatches.length ? (topMatches[0].strongMatch ? 'high' : 'moderate') : 'none',
      normalized: norm
    };
  } catch (err) {
    return {
      success: false,
      error: 'Unable to check duplicates.',
      details: String(err && err.message ? err.message : err)
    };
  }
}
