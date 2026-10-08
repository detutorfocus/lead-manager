/**
 * ============================================================
 * 01_Normalize.gs
 * Normalization and fuzzy-matching primitives used by the
 * duplicate-detection engine. No external libraries.
 * ============================================================
 */

var BUSINESS_SUFFIXES = [
  'ltd', 'limited', 'llc', 'inc', 'incorporated', 'enterprise',
  'enterprises', 'company', 'co'
];

/** Normalize a business name for comparison. */
function normalizeBusinessName(name) {
  if (!name) return '';
  var s = String(name).toLowerCase().trim();
  s = s.replace(/[’‘`]/g, "'");        // normalize apostrophes
  s = s.replace(/[^a-z0-9\s']/g, ' '); // strip punctuation (keep apostrophes for now)
  s = s.replace(/'/g, '');             // drop apostrophes entirely (Mama T's -> mama ts)
  s = s.replace(/\s+/g, ' ').trim();

  var tokens = s.split(' ').filter(function (t) { return t.length > 0; });
  // Strip a trailing legal/business suffix token, if present.
  while (tokens.length > 1 && BUSINESS_SUFFIXES.indexOf(tokens[tokens.length - 1]) !== -1) {
    tokens.pop();
  }
  return tokens.join(' ').trim();
}

/** Normalize an Instagram handle or URL down to the bare handle. */
function normalizeInstagram(value) {
  if (!value) return '';
  var s = String(value).trim().toLowerCase();
  s = s.replace(/^https?:\/\/(www\.)?instagram\.com\//, '');
  s = s.replace(/^@/, '');
  s = s.replace(/\/.*$/, '');   // drop anything after a slash (query params, extra path)
  s = s.replace(/\?.*$/, '');
  s = s.replace(/\s+/g, '');
  return s;
}

/** Normalize a LinkedIn company/profile URL. */
function normalizeLinkedIn(value) {
  if (!value) return '';
  var s = String(value).trim().toLowerCase();
  s = s.replace(/^https?:\/\//, '');
  s = s.replace(/^www\./, '');
  s = s.replace(/^linkedin\.com\//, '');
  s = s.replace(/\/+$/, '');
  s = s.replace(/\?.*$/, '');
  return s;
}

/** Normalize a website URL down to the bare domain (+ path, minus tracking noise). */
function normalizeWebsite(value) {
  if (!value) return '';
  var s = String(value).trim().toLowerCase();
  s = s.replace(/^https?:\/\//, '');
  s = s.replace(/^www\./, '');
  s = s.replace(/\?.*$/, '');   // drop query/tracking params
  s = s.replace(/#.*$/, '');
  s = s.replace(/\/+$/, '');    // drop trailing slash
  return s;
}

/** Extract just the domain from a normalized website, for looser matching. */
function websiteDomain_(normalizedWebsite) {
  if (!normalizedWebsite) return '';
  return normalizedWebsite.split('/')[0];
}

/** Normalize an email address. */
function normalizeEmail(value) {
  if (!value) return '';
  return String(value).trim().toLowerCase().replace(/\s+/g, '');
}

/** Normalize a phone number to a country-coded digit string, e.g. 234801234567. */
function normalizePhone(value, countryCode) {
  if (!value) return '';
  var cc = String(countryCode || '234');
  var digits = String(value).replace(/[\s\-\(\)\.]/g, '').replace(/[^\d+]/g, '');
  digits = digits.replace(/^\+/, '');
  if (digits.indexOf(cc) === 0) {
    return digits;
  }
  if (digits.indexOf('0') === 0) {
    return cc + digits.slice(1);
  }
  // Already looks like it has some country code, or is ambiguous - return as-is.
  return digits;
}

/** Normalize a city name. */
function normalizeCity(value) {
  if (!value) return '';
  return String(value).toLowerCase().trim().replace(/\s+/g, ' ');
}

/** Normalize a sector/category label. */
function normalizeSector(value) {
  if (!value) return '';
  return String(value).toLowerCase().trim().replace(/\s+/g, ' ');
}

/** Classic Levenshtein edit distance. */
function levenshteinDistance_(a, b) {
  a = a || ''; b = b || '';
  var m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  var prev = new Array(n + 1);
  var curr = new Array(n + 1);
  for (var j = 0; j <= n; j++) prev[j] = j;
  for (var i = 1; i <= m; i++) {
    curr[0] = i;
    for (var j2 = 1; j2 <= n; j2++) {
      var cost = a.charAt(i - 1) === b.charAt(j2 - 1) ? 0 : 1;
      curr[j2] = Math.min(
        prev[j2] + 1,      // deletion
        curr[j2 - 1] + 1,  // insertion
        prev[j2 - 1] + cost // substitution
      );
    }
    var tmp = prev; prev = curr; curr = tmp;
  }
  return prev[n];
}

/** Levenshtein-based similarity, 0-100. */
function levenshteinSimilarity_(a, b) {
  var maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 100;
  var dist = levenshteinDistance_(a, b);
  return Math.max(0, (1 - dist / maxLen)) * 100;
}

/** Token (Jaccard) similarity, 0-100. */
function tokenSimilarity_(a, b) {
  var ta = a.split(' ').filter(Boolean);
  var tb = b.split(' ').filter(Boolean);
  if (ta.length === 0 && tb.length === 0) return 100;
  if (ta.length === 0 || tb.length === 0) return 0;
  var setA = {}; ta.forEach(function (t) { setA[t] = true; });
  var setB = {}; tb.forEach(function (t) { setB[t] = true; });
  var intersection = 0;
  Object.keys(setA).forEach(function (t) { if (setB[t]) intersection++; });
  var union = Object.keys(setA).length + Object.keys(setB).length - intersection;
  return union === 0 ? 100 : (intersection / union) * 100;
}

/**
 * Combined fuzzy similarity between two ALREADY-NORMALIZED business names.
 * Blends character-level (Levenshtein) and token-level (Jaccard) similarity.
 * Returns 0-100.
 */
function calculateNameSimilarity(nameA, nameB) {
  var a = normalizeBusinessName(nameA);
  var b = normalizeBusinessName(nameB);
  if (!a && !b) return 100;
  if (!a || !b) return 0;
  if (a === b) return 100;

  var charSim = levenshteinSimilarity_(a, b);
  var tokSim = tokenSimilarity_(a, b);
  // Weighted blend: token overlap is a strong signal for reordered/partial names,
  // character similarity catches typos and minor variants.
  var combined = (charSim * 0.5) + (tokSim * 0.5);
  return Math.round(combined * 10) / 10;
}
