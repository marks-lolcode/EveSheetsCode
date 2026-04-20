/**
 * filterInventoryByTypeFirst()
 *
 * WHAT this version adds:
 * - Discovers column indexes by reading the FIRST ROW (header) of "AllInventory".
 * - Accepts common header name variants via a synonyms map.
 * - Detailed, beginner-friendly logging at every step for easy debugging.
 * - Headless-safe notifications via alertSafe (no UI methods called directly).
 *
 * HOW it works (high level):
 * 1) Read the header row from "AllInventory" and build a { field -> columnIndex } map.
 * 2) First pass: record Storage Facilities and buffer Launchpads.
 * 3) Second pass: include buffered Launchpads only if their (owner|location) has a Storage Facility.
 * 4) Sort by Fill % desc, then Owner asc, then Location asc. Write to "FilteredInventoryAll".
 *
 * RETURNS:
 * - The final 2D array written to the sheet (including header).
 */
function filterInventoryByTypeFirst() {
  // -----------------------------
  // Config and constants
  // -----------------------------
  const SOURCE_SHEET_NAME = 'AllInventory';
  const OUTPUT_SHEET_NAME = 'FilteredInventoryAll';

  // Known container capacities used to compute Fill %
  const CAPACITY = {
    STORAGE_FACILITY: 12000, // m³
    LAUNCHPAD: 10000         // m³
  };

  // We will look for these canonical fields in the header row.
  // Each field lists acceptable header variants (synonyms). All comparisons are normalized.
  const FIELD_SYNONYMS = {
    name:        ['type name'],
    count:       ['count'],
    owner:       ['owner'],
    location:    ['location'],
    totalVolume: ['total volume'],
    container:   ['container']
  };

  // For readable logs
  const L = (fmt, ...args) => Logger.log('[filterInventoryByTypeFirst] ' + fmt, ...args);
  const t0 = Date.now();
  L('Start.');

  // -----------------------------
  // Open source sheet and read data
  // -----------------------------
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    L("ERROR: Source sheet '%s' not found; aborting.", SOURCE_SHEET_NAME);
    alertSafe("Source sheet 'AllInventory' not found.");
    return [];
  }

  const range = sourceSheet.getDataRange();
  const values = range.getValues();
  if (!values || values.length < 2) {
    L('No data rows detected (header only or empty). Writing just the header to output.');
    return _writeOutputAndReturn(ss, OUTPUT_SHEET_NAME, [['Character','Location','Name','Count','Owner','Type','Total Volume','Fill %']]);
  }

  const headerRow = values[0];
  const dataRows  = values.slice(1);
  L('Read %s rows x %s cols from %s.', values.length, headerRow.length, SOURCE_SHEET_NAME);

  // -----------------------------
  // Build header index map from synonyms
  // -----------------------------
  const headerIndexMap = _buildHeaderIndexMap(headerRow, FIELD_SYNONYMS);
  L('Header index map: %s', JSON.stringify(headerIndexMap, null, 2));

  // Validate that required fields were found; if not, tell the user exactly what is missing
  const required = ['name', 'count', 'owner', 'location', 'totalVolume', 'container'];
  const missing  = required.filter(f => typeof headerIndexMap[f] !== 'number');
  if (missing.length) {
    const msg = 'Missing required header(s): ' + missing.join(', ') +
                '. Update your header row or add more synonyms in FIELD_SYNONYMS.';
    L('ERROR: %s', msg);
    alertSafe(msg, 'Header not found');
    return [];
  }

  // Extract resolved column indices to local constants for clarity
  const IDX = {
    NAME:        headerIndexMap.name,
    COUNT:       headerIndexMap.count,
    OWNER:       headerIndexMap.owner,
    LOCATION:    headerIndexMap.location,
    TOTAL_VOL:   headerIndexMap.totalVolume,
    CONTAINER:   headerIndexMap.container
  };

  // -----------------------------
  // Pass 1: scan rows
  // -----------------------------
  const storageMap = Object.create(null); // key: `${owner}|${location}` → true if a Storage Facility exists there
  const launchpadBuffer = [];             // buffered launchpads considered later
  const output = [];                      // rows we will write (without header)

  let scanned = 0, storageFound = 0, launchpadsBuffered = 0;

  for (const row of dataRows) {
    scanned++;

    // Defensive extraction with default values to avoid undefined issues
    const containerRaw = String(row[IDX.CONTAINER] || '').toLowerCase();
    const isStorage    = containerRaw.indexOf('storage facility') > -1;
    const isLaunchpad  = containerRaw.indexOf('launchpad') > -1;
    if (!isStorage && !isLaunchpad) continue; // only care about these two container types

    const owner       = String(row[IDX.OWNER]    || '').trim();
    const location    = String(row[IDX.LOCATION] || '').trim();
    const name        = row[IDX.NAME];
    const count       = row[IDX.COUNT];
    const totalVolume = Number(row[IDX.TOTAL_VOL]) || 0;
    const key         = owner + '|' + location;

    if (isStorage) {
      // Compute Storage Facility fill % using known capacity
      const fillPct = Math.min((totalVolume / CAPACITY.STORAGE_FACILITY) * 100, 100);
      output.push([ owner, location, name, count, owner, 'Storage Facility', totalVolume, fillPct ]);
      storageMap[key] = true;
      storageFound++;
    } else {
      // Buffer Launchpad and decide later based on storage presence at the same (owner|location)
      launchpadBuffer.push({ owner, location, key, name, count, totalVolume });
      launchpadsBuffered++;
    }
  }

  L('Pass 1: scanned=%s, storageFound=%s, launchpadsBuffered=%s', scanned, storageFound, launchpadsBuffered);

  // -----------------------------
  // Pass 2: include only launchpads paired with a storage facility
  // -----------------------------
  let launchpadsIncluded = 0;
  for (const lp of launchpadBuffer) {
    if (!storageMap[lp.key]) continue;
    const fillPct = Math.min((lp.totalVolume / CAPACITY.LAUNCHPAD) * 100, 100);
    output.push([ lp.owner, lp.location, lp.name, lp.count, lp.owner, 'Launchpad', lp.totalVolume, fillPct ]);
    launchpadsIncluded++;
  }
  L('Pass 2: launchpadsIncluded=%s (paired with Storage Facility).', launchpadsIncluded);

  // -----------------------------
  // Sort results for readability
  // -----------------------------
  const tSort = Date.now();
  output.sort((a, b) => {
    const fillDiff = b[7] - a[7];              // Fill % descending
    if (fillDiff !== 0) return fillDiff;
    const ownerCmp = String(a[0]).localeCompare(String(b[0])); // Owner ascending
    if (ownerCmp !== 0) return ownerCmp;
    return String(a[1]).localeCompare(String(b[1]));           // Location ascending
  });
  L('Sort completed in %s ms.', Date.now() - tSort);

  // -----------------------------
  // Format Fill % for display and prepend header
  // -----------------------------
  const HEADER = ['Character', 'Location', 'Name', 'Count', 'Owner', 'Type', 'Total Volume', 'Fill %'];
  const finalOutput = output.map(r => {
    const pctText = (Number(r[7]) || 0).toFixed(1) + '%';
    return [ r[0], r[1], r[2], r[3], r[4], r[5], r[6], pctText ];
  });
  finalOutput.unshift(HEADER);

  // -----------------------------
  // Write results
  // -----------------------------
  const tWrite = Date.now();
  const written = _writeOutputAndReturn(ss, OUTPUT_SHEET_NAME, finalOutput);
  L('Wrote %s rows x %s cols to "%s" in %s ms.',
    written.length, written[0] ? written[0].length : 0, OUTPUT_SHEET_NAME, Date.now() - tWrite);

  alertSafe('Filtered ' + output.length + ' rows to "' + OUTPUT_SHEET_NAME + '"');
  L('Done in %s ms total.', Date.now() - t0);
  return finalOutput;
}

/**
 * Builds a map from canonical field names to column indices by scanning the header row
 * and comparing against a list of synonyms.
 *
 * WHAT this does:
 *  - Normalizes both header cells and synonyms (lowercase, trim, remove punctuation, collapse spaces).
 *  - First match wins; logs duplicates so you can tidy headers if needed.
 *
 * @param {string[]} headerRow
 * @param {Object<string,string[]>} synonymsMap  e.g., { owner: ['owner','character'], ... }
 * @return {Object<string,number>} e.g., { owner: 3, location: 4, ... }
 */
function _buildHeaderIndexMap(headerRow, synonymsMap) {
  const L = (fmt, ...args) => Logger.log('[filterInventoryByTypeFirst] ' + fmt, ...args);

  // Normalize a header string: lowercase, remove punctuation, collapse whitespace
  const norm = s => String(s || '')
      .toLowerCase()
      .replace(/[_\-\/\\(){}\[\],.:;]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

  // Precompute normalized synonyms for fast matching
  const normSynonyms = {};
  for (const key of Object.keys(synonymsMap)) {
    normSynonyms[key] = (synonymsMap[key] || []).map(norm);
  }

  const result = {};
  const seenHeaders = new Map(); // normalized header -> index (for duplicate detection)

  headerRow.forEach((cell, idx) => {
    const h = norm(cell);
    if (!h) return;

    // Track duplicate headers (same normalized text appearing more than once)
    if (seenHeaders.has(h)) {
      L('WARNING: Duplicate header detected: "%s" at columns %s and %s (normalized match).',
        cell, seenHeaders.get(h) + 1, idx + 1);
    } else {
      seenHeaders.set(h, idx);
    }

    // Try to assign this header to one of the fields based on synonyms
    for (const key of Object.keys(normSynonyms)) {
      if (typeof result[key] === 'number') continue; // already matched this field
      if (normSynonyms[key].includes(h)) {
        result[key] = idx;
        L('Matched field "%s" to header "%s" (column %s).', key, cell, idx + 1);
        break;
      }
    }
  });

  return result;
}

/**
 * Writes a 2D array to a sheet, creating it if needed, or clearing it if it exists.
 * Returns the same array for convenience.
 *
 * WHY a helper:
 * - Centralizes write behavior and makes logging consistent.
 */
function _writeOutputAndReturn(ss, sheetName, values) {
  const L = (fmt, ...args) => Logger.log('[filterInventoryByTypeFirst] ' + fmt, ...args);

  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    L('Output sheet "%s" not found; creating it.', sheetName);
    sh = ss.insertSheet(sheetName);
  } else {
    L('Clearing existing contents of "%s".', sheetName);
    sh.clearContents();
  }

  if (!values || !values.length || !values[0] || !values[0].length) {
    L('WARNING: Nothing to write to "%s".', sheetName);
    return values || [];
  }

  sh.getRange(1, 1, values.length, values[0].length).setValues(values);
  return values;
}
