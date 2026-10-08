/**
 * ============================================================
 * 04_SearchService.gs
 * Search Leads: search by any of several fields, in memory,
 * against a single batch read of the sheet.
 * ============================================================
 */

var SEARCH_FIELD_COL = {
  'Business Name': LEAD_COL.BUSINESS_NAME,
  'Contact Name': LEAD_COL.CONTACT_NAME,
  'Instagram': LEAD_COL.INSTAGRAM,
  'LinkedIn': LEAD_COL.LINKEDIN,
  'Email': LEAD_COL.EMAIL,
  'Phone': LEAD_COL.WHATSAPP,
  'Website': LEAD_COL.WEBSITE,
  'City': LEAD_COL.CITY,
  'Sector': LEAD_COL.SECTOR,
  'Lead ID': LEAD_COL.LEAD_ID,
  'Assigned To': LEAD_COL.ASSIGNED_TO
};

/**
 * Search leads. field = one of SEARCH_FIELD_COL keys, or 'All' to search
 * every searchable column. Returns up to 100 matches.
 */
function searchLeads(query, field) {
  try {
    var q = String(query || '').trim().toLowerCase();
    if (!q) return { success: true, results: [] };

    var sh = getOrCreateSheet_(SHEET_LEADS);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, results: [] };
    var values = sh.getRange(2, 1, lastRow - 1, LEADS_HEADERS.length).getValues();

    var cols = (field && field !== 'All' && SEARCH_FIELD_COL[field])
      ? [SEARCH_FIELD_COL[field]]
      : Object.keys(SEARCH_FIELD_COL).map(function (k) { return SEARCH_FIELD_COL[k]; });

    var results = [];
    for (var i = 0; i < values.length && results.length < 100; i++) {
      var row = values[i];
      if (!row[LEAD_COL.LEAD_ID - 1]) continue;
      var matched = cols.some(function (c) {
        var cell = row[c - 1];
        return cell && String(cell).toLowerCase().indexOf(q) !== -1;
      });
      if (matched) {
        results.push({
          leadId: row[LEAD_COL.LEAD_ID - 1],
          businessName: row[LEAD_COL.BUSINESS_NAME - 1],
          city: row[LEAD_COL.CITY - 1],
          sector: row[LEAD_COL.SECTOR - 1],
          contactName: row[LEAD_COL.CONTACT_NAME - 1],
          instagram: row[LEAD_COL.INSTAGRAM - 1],
          website: row[LEAD_COL.WEBSITE - 1],
          recordStatus: row[LEAD_COL.RECORD_STATUS - 1],
          duplicateStatus: row[LEAD_COL.DUP_STATUS - 1]
        });
      }
    }
    return { success: true, results: results };
  } catch (err) {
    return { success: false, error: 'Search failed.', details: String(err && err.message ? err.message : err) };
  }
}
