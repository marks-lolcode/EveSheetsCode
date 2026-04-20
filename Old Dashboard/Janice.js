/* ============================================================================
 * File: Janice.gs
 * Purpose: Adds detailed logging (and beginner-friendly comments) to existing
 *          functions without changing any logic or behavior.
 * Notes:
 * - Logging uses Logger.log(...) throughout to trace inputs, decisions,
 *   loop progress, cache hits, HTTP calls, and outputs.
 * - All original code remains intact; only comments and Logger.log lines
 *   have been added.
 * ========================================================================== */

// Janice Prices Update with Multiple Dynamic Flag Cells
function updateJanice() {
  Logger.log("[updateJanice] START");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JaniceCommodity");
  Logger.log("[updateJanice] Target sheet resolved: %s", sheet ? sheet.getName() : "(null)");

  var range = sheet.getRange("JaniceUpdate");
  Logger.log("[updateJanice] Using named range 'JaniceUpdate' -> A1Notation: %s", range.getA1Notation());
  
  // Toggle checkbox value
  // WHAT: Reads the current checkbox value and flips it true<->false.
  // WHY: Toggling this cell is used to force recalculation in the sheet.
  var currentValue = range.getValue();
  Logger.log("[updateJanice] Current checkbox value: %s", currentValue);

  var toggledValue = !currentValue;
  Logger.log("[updateJanice] Toggled checkbox value about to be written: %s", toggledValue);

  range.setValue(toggledValue);
  Logger.log("[updateJanice] Checkbox value written.");

  // Determine the last row and calculate flag cells
  // WHAT: We identify rows whose cells will be checked to detect when the sheet
  //       has finished recalculating.
  // WHY: Polling these “flag” cells until they stop changing is a simple sync
  //      strategy when formulas update in chunks.
  var lastRow = sheet.getLastRow();
  Logger.log("[updateJanice] lastRow detected: %s", lastRow);

  var flagCells = [];
  
  // Add every 499th row
  // WHAT: Every 499th row (A499, A998, ...) is added to the watch list.
  // WHY: Large sheets often update in batches; sampling every ~500 rows is a
  //      compromise between coverage and speed.
  for (var i = 499; i < lastRow; i += 499) {
    var cellA1 = "A" + i;
    Logger.log("[updateJanice] Adding periodic flag cell: %s", cellA1);
    flagCells.push(sheet.getRange(cellA1));
  }
  
  // Add the last row
  // WHAT/WHY: Always include the last visible row to catch tail-end updates.
  var lastCellA1 = "A" + lastRow;
  Logger.log("[updateJanice] Adding last-row flag cell: %s", lastCellA1);
  flagCells.push(sheet.getRange(lastCellA1));
  
  // Function to get all flag cell values
  // WHAT: Reads current values from all flag cells and returns as an array.
  // WHY: We compare successive reads to see when values stop changing.
  function getFlagValues() {
    var values = flagCells.map(function(cell) {
      return cell.getValue();
    });
    Logger.log("[updateJanice/getFlagValues] Snapshot: %s", JSON.stringify(values));
    return values;
  }
  
  // Wait for recalculation to complete
  // WHAT: Loop until two consecutive snapshots of the flag values are identical.
  // WHY: Identical snapshots strongly suggest formulas have stabilized.
  var previousValues = [];
  var newValues = getFlagValues();
  var maxChecks = 50; // Max number of checks to avoid infinite loops
  var checks = 0;
  Logger.log("[updateJanice] Begin stabilization loop; maxChecks=%s", maxChecks);
  
  // Check all flag cells until their values stop changing
  do {
    previousValues = newValues;
    Utilities.sleep(100); // Wait 100 milliseconds
    Logger.log("[updateJanice] Slept 100ms before next snapshot. Iteration=%s", checks + 1);

    newValues = getFlagValues();
    checks++;
    Logger.log("[updateJanice] arraysEqual(previous, current)=%s", arraysEqual(previousValues, newValues));
  } while (!arraysEqual(previousValues, newValues) && checks < maxChecks);
  
  if (checks >= maxChecks) {
    Logger.log("[updateJanice] Recalculation did not complete within the time limit. checks=%s", checks);
  } else {
    Logger.log("[updateJanice] Recalculation completed. checks=%s", checks);
  }

  Logger.log("[updateJanice] END");
}

// Helper function to compare two arrays
// WHAT: Performs strict, position-by-position equality check.
// WHY: Used by the stabilization loop to detect no further changes.
function arraysEqual(arr1, arr2) {
  Logger.log("[arraysEqual] Compare arrays. len1=%s len2=%s", (arr1 || []).length, (arr2 || []).length);
  if (arr1.length !== arr2.length) return false;
  for (var i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) {
      Logger.log("[arraysEqual] Mismatch at index %s: %s !== %s", i, arr1[i], arr2[i]);
      return false;
    }
  }
  return true;
}




//  Get your api key by contacting me on discord kukki#3914
//  
//  Purpose of having own api key is so that I can contact/block people with excessive traffic.
//  If you use api key not bound to your name you might find yourself blocked at some point without warning.
var JANICE_API_KEY = 'hSevDkQLXrJFaJsAsnWttr7GXUKHBFdE';

var JaniceUtils = (function() {
  Logger.log("[JaniceUtils] Module init.");

  var API_URL = 'https://janice.e-351.com/api/rest/v2';
  var MARKETS = {
    'jita': 2,
    'r1o-gn': 3,
    'perimeter': 4,
    'jitameter': 5,
    'npc': 6,
    'mj-5f9': 114,
    'amarr': 115
  };
  Logger.log("[JaniceUtils] Constants set. API_URL=%s MARKETS=%s", API_URL, JSON.stringify(MARKETS));
  
  function fetch_(url, options, cacheBuster) {
    Logger.log("[JaniceUtils.fetch] ENTER url=%s cacheBuster=%s options=%s",
      url, cacheBuster, JSON.stringify(options));

    var cache = CacheService.getScriptCache();
    var requestHash = Utilities.base64Encode(
      Utilities.computeDigest(
        Utilities.DigestAlgorithm.MD5,
        JSON.stringify({ url: url, options: options, cacheBuster: cacheBuster })
      )
    );
    Logger.log("[JaniceUtils.fetch] requestHash=%s", requestHash);
    
    var cachedResponseString = cache.get(requestHash);
    if (cachedResponseString != null) {
      Logger.log("[JaniceUtils.fetch] Cache HIT for requestHash=%s (TTL ~6h).", requestHash);
      return cachedResponseString;
    }
    Logger.log("[JaniceUtils.fetch] Cache MISS for requestHash=%s.", requestHash);
    
    // Randomized short sleep to avoid thundering herd
    var jitterMs = Math.random() * 2000;
    Logger.log("[JaniceUtils.fetch] Sleeping jitter ms: %s", Math.round(jitterMs));
    Utilities.sleep(jitterMs);
    
    Logger.log("[JaniceUtils.fetch] Performing UrlFetchApp.fetch...");
    var response = UrlFetchApp.fetch(url, options);  
    var responseString = response.getContentText();
    Logger.log("[JaniceUtils.fetch] HTTP status code: %s, content length: %s",
      response.getResponseCode(), responseString ? responseString.length : 0);
    
    cache.put(requestHash, responseString, 21600); // 6 hours
    Logger.log("[JaniceUtils.fetch] Response cached for requestHash=%s", requestHash);
    
    Logger.log("[JaniceUtils.fetch] EXIT");
    return responseString;
  }
  
  function fetch_json_(url, options, cacheBuster) {
    Logger.log("[JaniceUtils.fetchJson] ENTER url=%s cacheBuster=%s", url, cacheBuster);
    var responseString = fetch_(url, options, cacheBuster);
    Logger.log("[JaniceUtils.fetchJson] Parsing JSON (length=%s)", responseString ? responseString.length : 0);
    var parsed = JSON.parse(responseString);
    Logger.log("[JaniceUtils.fetchJson] EXIT (parsed type=%s)", typeof parsed);
    return parsed;
  }
  
  function transpose_(a) {
    Logger.log("[JaniceUtils.transpose] ENTER rows=%s cols=%s",
      a && a.length, (a && a[0]) ? a[0].length : 0);
    var t = Object.keys(new Array(a[0].length).fill()).map(function (c) { return a.map(function (r) { return r[c]; }); });
    Logger.log("[JaniceUtils.transpose] EXIT rows=%s cols=%s", t.length, (t[0] || []).length);
    return t;
  }

  function to_string_(value) {
    // WHAT: Converts any value into a string safely.
    // WHY: Upstream functions expect string comparison and logging readability.
    var result;
    if (typeof value === 'string') {
      result = value;
    } else if (value === undefined || value === null) {
      result = '';
    } else {
      result = value.toString();
    }
    Logger.log("[JaniceUtils.to_string] InputType=%s Output='%s'", typeof value, result);
    return result;
  }
  
  function find_(where, what, by) {
    Logger.log("[JaniceUtils.find] ENTER by=%s", JSON.stringify(by));
    if (!Array.isArray(where)) {
      Logger.log("[JaniceUtils.find] 'where' is not an array. Returning null.");
      return null;
    }

    what = to_string_(what).toUpperCase();
    Logger.log("[JaniceUtils.find] Normalized search key: '%s'", what);

    for (var i = 0; i < where.length; i++) {
      for (var ki = 0; ki < by.length; ki++) {
        var value = navigate_(where[i], by[ki]);
        var valueStr = to_string_(value).toUpperCase();
        if (valueStr === what) {
          Logger.log("[JaniceUtils.find] MATCH at index=%s byKey=%s", i, by[ki]);
          return where[i];
        }
      }
    }
    
    Logger.log("[JaniceUtils.find] No match found. Returning null.");
    return null;
  }
  
  function navigate_(data, key) {
    // WHAT: Navigates nested objects using dotted/bracket notation (e.g., "a.b[0].c").
    // WHY: Allows flexible access to API response paths without hardcoding chains.
    if (typeof key === 'string') {
      key = key.split(/[\.\[\]]/gi).filter(function (x) { return x.length > 0; });
    } else {
      Logger.log("[JaniceUtils.navigate] ERROR: key is not a string.");
      throw new Error('Function \'navigate_\' expects parameter \'key\' to be a string');
    }
    
    for (var i = 0; i < key.length; i++) {
      if (typeof data !== 'object') {
        Logger.log("[JaniceUtils.navigate] Early exit: data is non-object at segment '%s'", key[i]);
        return null;
      }
      data = data[key[i]];
    }
    
    Logger.log("[JaniceUtils.navigate] Resolved value for path: %s", JSON.stringify(data));
    return data;
  }
  
  function format_row_(data, spec) {
    Logger.log("[JaniceUtils.format_row] ENTER spec=%s", JSON.stringify(spec));
    var result = [];
    for (var i = 0; i < spec.length; i++) {
      result[i] = navigate_(data, spec[i]);
    }
    Logger.log("[JaniceUtils.format_row] EXIT row=%s", JSON.stringify(result));
    return result;
  }
  
  function Range(items) {
    Logger.log("[JaniceUtils.Range] ENTER - constructing Range from items.");
    this.mode = 'error';
    this.input = [];
    
    if (Array.isArray(items)) {
      this.mode = 'row';
      if (items.length === 1) {
        var row = items[0];
        if (Array.isArray(row)) {
          for (var i = 0; i < row.length; i++) {
            var value = row[i];
            if ((typeof value === 'string' && value.length > 0) || typeof value === 'number') {
              this.input[i] = value.toString().trim();
            }
          }
        }
      } else {
        this.mode = 'column';
        var column = items;
        for (var i = 0; i < column.length; i++) {
          var row = column[i];
          if (Array.isArray(row) && row.length === 1) {
            var value = row[0];
            if ((typeof value === 'string' && value.length > 0) || typeof value === 'number') {
              this.input[i] = value.toString().trim();
            }
          } else {
            this.mode = 'error';
            break;
          }
        }
      }
    } else if ((typeof items === 'string' && items.length > 0) || typeof items === 'number') {
      this.mode = 'single';
      this.input[0] = items.toString().trim();
    } 
    
    if (this.mode === 'error') {
      Logger.log("[JaniceUtils.Range] ERROR: Provided items are not a single dimensional range.");
      throw new Error('Parameter \'items\' must be a single dimensional range.');
    }
    
    this.output = [];
    for (var i = 0; i < this.input.length; i++) {
      this.output.push([""]);
    }

    Logger.log("[JaniceUtils.Range] EXIT mode=%s inputLen=%s", this.mode, this.input.length);
  }
  
  Range.prototype.getResult = function () {  
    Logger.log("[JaniceUtils.Range.getResult] ENTER mode=%s", this.mode);
    if (this.mode === 'single' || this.mode === 'column') {
      Logger.log("[JaniceUtils.Range.getResult] EXIT (column/single) rows=%s cols=%s",
        this.output.length, (this.output[0] || []).length);
      return this.output;
    } else if (this.mode === 'row') {
      var t = transpose_(this.output);
      Logger.log("[JaniceUtils.Range.getResult] EXIT (row->transpose) rows=%s cols=%s",
        t.length, (t[0] || []).length);
      return t;
    } else {
      Logger.log("[JaniceUtils.Range.getResult] ERROR: Invalid mode '%s'", this.mode);
      throw new Error('Invalid mode \'' + this.mode + '\'');
    }  
  }
  
  return {
    API_URL: API_URL,
    MARKETS: MARKETS,
    
    fetch: fetch_,
    fetchJson: fetch_json_,
    
    transpose: transpose_,
    find: find_,
    navigate: navigate_,
    formatRow: format_row_,
    
    Range: Range,
  };
})();

/**
 * Return pricing information for given single dimensional item range.
 *
 * WHAT: Calls Janice API to fetch immediate buy/sell prices (or any requested fields).
 * WHY: Automates obtaining market pricing for a list of item IDs or names directly from Sheets.
 *
 * @param {range} items Items to be appraised. Either type name or type id can be used.
 * @param {string?} spec Column specification. Example: "itemType.eid|itemType.name|immediatePrices.buyPrice|immediatePrices.sellPrice"
 * @param {string?} market Market, options: "jita", "perimeter", "r1o-gn".
 * @param {string?} cacheBuster String to break through cache.
 * @customfunction
 */
function JANICE_PRICER(items, spec, market, cacheBuster) {
  Logger.log("[JANICE_PRICER] START");

  var range = new JaniceUtils.Range(items);
  Logger.log("[JANICE_PRICER] Range constructed. inputLen=%s", range.input.length);
  
  if (range.input.length <= 0) {
    Logger.log("[JANICE_PRICER] No input items. Returning [null].");
    return [null];
  }
  
  if (typeof spec === 'string' && spec.length > 0) {
    spec = spec.split('|');
  } else {
    spec = ['immediatePrices.buyPrice', 'immediatePrices.sellPrice'];
  }
  Logger.log("[JANICE_PRICER] Spec fields: %s", JSON.stringify(spec));
  
  if (typeof market === 'string') {
    var originalMarket = market;
    market = JaniceUtils.MARKETS[market.toLowerCase()]
    if (!market) {
      Logger.log("[JANICE_PRICER] ERROR: invalid market '%s'", originalMarket);
      throw new Error('invalid market');
    }
  } else {
    market = JaniceUtils.MARKETS['jita'];
  }
  Logger.log("[JANICE_PRICER] Market numeric id: %s", market);
  
  if (!cacheBuster) {
    cacheBuster = '-';
  }
  Logger.log("[JANICE_PRICER] cacheBuster: %s", cacheBuster);
  
  var url = JaniceUtils.API_URL + '/pricer?market=' + encodeURIComponent(market) + '&_=' + encodeURIComponent(cacheBuster);
  Logger.log("[JANICE_PRICER] Request URL: %s", url);

  var data = JaniceUtils.fetchJson(url, { 
    method: 'post',
    contentType: "text/plain",
    headers: {
      'X-ApiKey': JANICE_API_KEY,
    },
    payload: range.input.join('\n'),
  }, cacheBuster);
  Logger.log("[JANICE_PRICER] Response array length: %s", Array.isArray(data) ? data.length : "(not array)");
  
  for (var i = 0; i < range.input.length; i++) {
    var input = range.input[i];
    if (typeof input !== 'string' && typeof input !== 'number') {
      Logger.log("[JANICE_PRICER] Skipping non-string/number input at index=%s: %s", i, typeof input);
      continue;
    }
    
    var typeInfo = JaniceUtils.find(data, input, ['itemType.eid', 'itemType.name']);
    if (!typeInfo) {
      Logger.log("[JANICE_PRICER] No match for input='%s' (by eid or name). Leaving default output.", input);
      continue;
    }
    
    range.output[i] = JaniceUtils.formatRow(typeInfo, spec);
    Logger.log("[JANICE_PRICER] Output row %s -> %s", i, JSON.stringify(range.output[i]));
  }

  var result = range.getResult();
  Logger.log("[JANICE_PRICER] END. rows=%s cols=%s", result.length, (result[0] || []).length);
  return result;
}

/**
 * Obtain base job cost for specified items.
 *
 * WHAT: Queries base industry job cost for items, for a given activity.
 * WHY: Useful for planning and profitability calculations in manufacturing chains.
 *
 * @param {range} items Items to be evaluated. Either type name or type id can be used.
 * @param {string?} activity Activity, options: "manufacturing", "researchingtimeefficiency", "researchingmaterialefficiency", "copying", "invention", "reactions"
 * @param {string?} spec Column specification. Example: "itemType.eid|itemType.name|baseJobCost"
 * @param {string?} cacheBuster String to break through cache.
 * @customfunction
 */
function JANICE_BASE_JOB_COST(items, activity, spec, cacheBuster) {
  Logger.log("[JANICE_BASE_JOB_COST] START activity=%s", activity);

  var range = new JaniceUtils.Range(items);
  Logger.log("[JANICE_BASE_JOB_COST] Range constructed. inputLen=%s", range.input.length);

  if (range.input.length <= 0) {
    Logger.log("[JANICE_BASE_JOB_COST] No input items. Returning [null].");
    return [null];
  }

  if (typeof spec === 'string' && spec.length > 0) {
    spec = spec.split('|');
  } else {
    spec = ['baseJobCost'];
  }
  Logger.log("[JANICE_BASE_JOB_COST] Spec fields: %s", JSON.stringify(spec));
  
  if (!cacheBuster) {
    cacheBuster = '-';
  }
  Logger.log("[JANICE_BASE_JOB_COST] cacheBuster: %s", cacheBuster);
  
  var url = JaniceUtils.API_URL + '/industry/base-job-cost?activity=' + encodeURIComponent(activity) + '&_=' + encodeURIComponent(cacheBuster);
  Logger.log("[JANICE_BASE_JOB_COST] Request URL: %s", url);
  
  var data = JaniceUtils.fetchJson(url, { 
    method: 'post',
    contentType: "text/plain",
    headers: {
      'X-ApiKey': JANICE_API_KEY,
    },
    payload: range.input.join('\n'),
  }, cacheBuster);
  Logger.log("[JANICE_BASE_JOB_COST] Response array length: %s", Array.isArray(data) ? data.length : "(not array)");

  for (var i = 0; i < range.input.length; i++) {
    var input = range.input[i];
    if (typeof input !== 'string' && typeof input !== 'number') {
      Logger.log("[JANICE_BASE_JOB_COST] Skipping non-string/number input at index=%s: %s", i, typeof input);
      continue;
    }
    
    var typeInfo = JaniceUtils.find(data, input, ['itemType.eid', 'itemType.name']);
    if (!typeInfo) {
      Logger.log("[JANICE_BASE_JOB_COST] No match found for '%s'. Filling with empty columns (len=%s).", input, spec.length);
      range.output[i] = new Array(spec.length);
      continue;
    }
    
    range.output[i] = JaniceUtils.formatRow(typeInfo, spec);
    Logger.log("[JANICE_BASE_JOB_COST] Output row %s -> %s", i, JSON.stringify(range.output[i]));
  }

  var result = range.getResult();
  Logger.log("[JANICE_BASE_JOB_COST] END. rows=%s cols=%s", result.length, (result[0] || []).length);
  return result;
}

/* ============================================================================
 * End of File: Janice.gs
 * ========================================================================== */
