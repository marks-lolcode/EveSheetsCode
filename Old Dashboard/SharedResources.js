/** =====================================================================
 * File: LookupByNameOrId.gs
 * Purpose: Provide robust, logged, user-friendly custom functions to:
 *          - Get Item ID from Item Name: =GETITEMID(name)
 *          - Get Item Name from Item ID: =GETITEMNAME(id)
 *
 * Requirements:
 *   - Two named ranges in the spreadsheet:
 *       1) invItemName  → column of item names (one value per row)
 *       2) invItemID    → column of item IDs   (one value per row)
 *   - Ranges must be the same length and the nth name corresponds to the nth ID.
 *
 * Notes:
 *   - Custom functions in Google Sheets cannot write to other cells or sheets.
 *     Therefore, "logging" is done via Logger, visible in Apps Script
 *     under Extensions → Apps Script → Executions.
 *   - This file uses a short-lived in-memory/document cache to improve performance
 *     for repeated calls in recalculation-heavy sheets.
 *
 * Usage:
 *   =GETITEMID(A2)     // returns the ID for the name in A2
 *   =GETITEMNAME(B2)   // returns the name for the ID in B2
 *
 * ---------------------------------------------------------------------
 * Change Log:
 *   v1.0  (2025-10-21)  Initial version with validation, caching, and logging.
 * ===================================================================== */


/**
 * ===========================
 * Public Custom Functions
 * ===========================
 */

/**
 * GETITEMID(name): Return the corresponding Item ID for a given Item Name.
 *
 * WHAT it does:
 *   - Reads the invItemName and invItemID named ranges
 *   - Validates they are aligned and non-empty
 *   - Finds the row(s) where the provided name matches (case-insensitive)
 *   - Returns the corresponding ID, with graceful error text on problems
 *
 * WHY this approach:
 *   - INDEX/MATCH pattern reproduced in Apps Script for clearer logging,
 *     better error messages, and handling of duplicates.
 *
 * @param {any} nameInput A cell reference or literal text representing the Item Name.
 * @return {string} The matching Item ID, or a human-readable error string.
 */
function GETITEMID(nameInput) {
  const fn = 'GETITEMID';
  try {
    log_(fn, 'Start');

    // Normalize the input early to avoid accidental mismatches later.
    const nameQuery = normalize_(nameInput);
    if (!nameQuery) {
      log_(fn, 'Input is blank or not a string/number. Returning helpful message.');
      return 'Not found: blank name';
    }

    // Retrieve aligned arrays for names and IDs (with validation + caching).
    const { names, ids } = getAlignedNameIdData_(fn);

    // Find all exact (case-insensitive, trimmed) matches for the name.
    const matchIndexes = findAllMatches_(names, nameQuery);
    if (matchIndexes.length === 0) {
      log_(fn, `Name "${nameQuery}" not found in invItemName.`);
      return `Not found in invItemName: ${nameQuery}`;
    }
    if (matchIndexes.length > 1) {
      const candidateIds = matchIndexes.map(i => ids[i]).filter(v => v !== '');
      log_(fn, `Ambiguous name "${nameQuery}" — ${matchIndexes.length} matches. IDs: ${JSON.stringify(candidateIds)}`);
      return `Ambiguous name (multiple matches): ${candidateIds.join(', ')}`;
    }

    const idx = matchIndexes[0];
    const result = safeToString_(ids[idx]);
    log_(fn, `Resolved name "${nameQuery}" → ID "${result}" at rowIndex ${idx}.`);
    return result || `Not found: ID missing at row for "${nameQuery}"`;

  } catch (err) {
    log_(fn, `ERROR: ${err && err.stack ? err.stack : err}`);
    // Return an error string instead of throwing, so the cell shows a helpful message.
    return `Error in ${fn}: ${err && err.message ? err.message : err}`;
  } finally {
    log_(fn, 'End');
  }
}


/**
 * GETITEMNAME(id): Return the corresponding Item Name for a given Item ID.
 *
 * WHAT it does:
 *   - Reads the invItemName and invItemID named ranges
 *   - Validates they are aligned and non-empty
 *   - Finds the row(s) where the provided ID matches (string compare, case-insensitive)
 *   - Returns the corresponding Name, with graceful error text on problems
 *
 * WHY this approach:
 *   - Mirrors GETITEMID for consistent behavior, logging, and duplicate handling.
 *
 * @param {any} idInput A cell reference or literal representing the Item ID.
 * @return {string} The matching Item Name, or a human-readable error string.
 */
function GETITEMNAME(idInput) {
  const fn = 'GETITEMNAME';
  try {
    log_(fn, 'Start');

    const idQuery = normalize_(idInput);
    if (!idQuery) {
      log_(fn, 'Input is blank or not a string/number. Returning helpful message.');
      return 'Not found: blank id';
    }

    const { names, ids } = getAlignedNameIdData_(fn);

    const matchIndexes = findAllMatches_(ids, idQuery);
    if (matchIndexes.length === 0) {
      log_(fn, `ID "${idQuery}" not found in invItemID.`);
      return `Not found in invItemID: ${idQuery}`;
    }
    if (matchIndexes.length > 1) {
      const candidateNames = matchIndexes.map(i => names[i]).filter(v => v !== '');
      log_(fn, `Ambiguous id "${idQuery}" — ${matchIndexes.length} matches. Names: ${JSON.stringify(candidateNames)}`);
      return `Ambiguous id (multiple matches): ${candidateNames.join(', ')}`;
    }

    const idx = matchIndexes[0];
    const result = safeToString_(names[idx]);
    log_(fn, `Resolved ID "${idQuery}" → Name "${result}" at rowIndex ${idx}.`);
    return result || `Not found: Name missing at row for "${idQuery}"`;

  } catch (err) {
    log_(fn, `ERROR: ${err && err.stack ? err.stack : err}`);
    return `Error in ${fn}: ${err && err.message ? err.message : err}`;
  } finally {
    log_(fn, 'End');
  }
}


/**
 * ===========================
 * Internal Helpers (Private)
 * ===========================
 */

/**
 * Fetches invItemName and invItemID as aligned 1D arrays.
 * Includes validation, trimming, and light caching for speed.
 *
 * WHY caching:
 *   - Custom functions can be re-evaluated frequently.
 *   - Caching reduces repeated range reads and JSON processing overhead.
 *
 * @param {string} caller Name of the calling function (for logs).
 * @return {{names:string[], ids:string[]}}
 */
function getAlignedNameIdData_(caller) {
  const fn = 'getAlignedNameIdData_';
  const cache = CacheService.getDocumentCache();
  const cacheKey = 'LookupByNameOrId.v1.names_ids';

  // Try cache first.
  const cached = cache && cache.get(cacheKey);
  if (cached) {
    log_(fn, 'Using cached name/id arrays.');
    try {
      const { names, ids } = JSON.parse(cached);
      if (Array.isArray(names) && Array.isArray(ids) && names.length === ids.length && names.length > 0) {
        return { names, ids };
      }
      log_(fn, 'Cache present but invalid shape; falling back to live read.');
    } catch (e) {
      log_(fn, 'Cache JSON parse failed; falling back to live read.');
    }
  }

  // Live read from named ranges.
  log_(fn, 'Reading named ranges from spreadsheet.');
  const ss = SpreadsheetApp.getActive();
  const namedRangeNames = ['invItemName', 'invItemID'];

  // Validate named ranges exist.
  const allRanges = ss.getNamedRanges();
  const namesSet = new Set(allRanges.map(nr => nr.getName()));
  for (const requiredName of namedRangeNames) {
    if (!namesSet.has(requiredName)) {
      throw new Error(`Missing named range: ${requiredName}`);
    }
  }

  // Pull values, normalize to 1D arrays, and trim.
  const nameVals2D = ss.getRangeByName('invItemName').getValues();
  const idVals2D   = ss.getRangeByName('invItemID').getValues();

  const names = flatten1D_(nameVals2D).map(safeToString_).map(trimLowerStable_);
  const ids   = flatten1D_(idVals2D).map(safeToString_).map(trimLowerStable_);

  // Basic validations.
  if (names.length === 0 || ids.length === 0) {
    throw new Error('invItemName or invItemID is empty.');
  }
  if (names.length !== ids.length) {
    throw new Error(`Length mismatch: invItemName has ${names.length} rows; invItemID has ${ids.length} rows.`);
  }

  // Keep original display values for returns, but store normalized forms separately.
  // For return values, we want the original, not lowercased. Re-read original as display strings.
  const rawNames = flatten1D_(nameVals2D).map(safeToString_);
  const rawIds   = flatten1D_(idVals2D).map(safeToString_);

  // Package final aligned arrays: raw for return, normalized only used for matching.
  // To keep the surface simple for callers, return raw arrays and do matching externally
  // with a parallel normalized copy.
  const payload = {
    namesRaw: rawNames,
    idsRaw: rawIds,
    namesNorm: names,
    idsNorm: ids
  };

  // For cache, store only what callers need: raw arrays plus normalized arrays for searching.
  cache && cache.put(
    cacheKey,
    JSON.stringify({ names: rawNames, ids: rawIds, namesNorm: names, idsNorm: ids }),
    300 // seconds
  );

  // Return a facade that exposes raw arrays, but callers here expect {names, ids}.
  // To avoid breaking callers, we re-read from cache (which has both raw and norm).
  const rebuilt = cache && cache.get(cacheKey);
  if (rebuilt) {
    const { names: rawN, ids: rawI } = JSON.parse(rebuilt);
    log_(fn, `Live read complete. Rows: ${rawN.length}.`);
    // Return raw arrays for output use; matching will re-fetch normalized arrays via helper.
    return { names: rawN, ids: rawI };
  }

  // Fallback if cache failed for some reason.
  log_(fn, `Live read complete (no cache write). Rows: ${rawNames.length}.`);
  return { names: rawNames, ids: rawIds };
}


/**
 * Finds all index positions in "haystackRaw" matching "needle" with normalized compare.
 * We normalize both sides to ensure case-insensitive and trimmed comparisons.
 *
 * @param {string[]} haystackRaw Raw values array (not lowercased).
 * @param {string} needle The query (already normalized in caller).
 * @return {number[]} Array of matching indexes (may be empty or multi-length).
 */
function findAllMatches_(haystackRaw, needle) {
  // Build a normalized copy once for comparison (fast enough for typical lists).
  const hayNorm = haystackRaw.map(trimLowerStable_);
  const hits = [];
  for (let i = 0; i < hayNorm.length; i++) {
    if (hayNorm[i] === needle) {
      hits.push(i);
    }
  }
  return hits;
}


/**
 * Utility: Normalize arbitrary cell input to a safe, lowercased, trimmed string.
 * Returns "" if the value is null/undefined/empty after trimming.
 */
function normalize_(v) {
  const s = safeToString_(v).trim();
  return s ? s.toLowerCase() : '';
}

/**
 * Utility: Convert value safely to string (numbers preserved as digits).
 */
function safeToString_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return String(v);
  // For arrays or other types, stringify predictably.
  try {
    return String(v);
  } catch (_) {
    return '';
  }
}

/**
 * Utility: Trim + lowercase while handling blanks consistently.
 */
function trimLowerStable_(s) {
  return safeToString_((s || '')).trim().toLowerCase();
}

/**
 * Utility: Flatten a 2D range into a 1D column (top-to-bottom).
 * If the input is a single column, this returns that column.
 * If the input is a row, it returns that row’s cells in order.
 */
function flatten1D_(arr2d) {
  // If it’s already a 1D array, return as-is.
  if (!Array.isArray(arr2d)) return [];
  if (arr2d.length > 0 && !Array.isArray(arr2d[0])) {
    return arr2d.map(safeToString_);
  }
  // Standard sheet values: 2D array of rows
  const out = [];
  for (let r = 0; r < arr2d.length; r++) {
    for (let c = 0; c < arr2d[r].length; c++) {
      out.push(arr2d[r][c]);
    }
  }
  return out;
}

/**
 * Logging helper: prefixes with timestamp and function name.
 * Visible in Apps Script → Executions.
 *
 * WHY we don’t write to a Log sheet:
 *   Custom functions are restricted from editing the spreadsheet.
 *   Logger is the supported way to capture debug information here.
 */
function log_(fn, msg) {
  const ts = new Date().toISOString();
  Logger.log(`[${ts}] [${fn}] ${msg}`);
}


/** =====================================================================
 * End of file: LookupByNameOrId.gs
 * ===================================================================== */
