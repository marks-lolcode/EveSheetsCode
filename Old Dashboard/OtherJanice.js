/**
 * Batch-fill Jita/Amarr prices into JaniceCommodity as VALUES ONLY, with:
 * - Extensive logging (DEBUG/INFO/WARN/ERROR) + timings for each step.
 * - Time slicing to avoid "Exceeded maximum execution time".
 * - Skipping already-filled rows by default (fast reruns).
 * - Robust retry for names with leading apostrophes or curly quotes
 *   (e.g., "'Wetu' Mobile Depot" / "‘Wetu’ Mobile Depot").
 *
 * ENTRY POINTS (wire these to your own menu or a button; no onOpen here):
 *   runPriceFillBatches()            // Fast mode: skips rows with existing non-error values
 *   runPriceFillBatchesForceAll()    // Full refresh: recompute all rows
 *   runPriceFillForSelection()       // Spot refresh: only the current selection
 *
 * Beginner-friendly comments included to explain what and why.
 */

/* ===========================
 * GLOBAL CONFIG + LOGGING CFG
 * =========================== */

var PRICE_CFG = {
  MAIN_SHEET: 'JaniceCommodity',
  TEMP_SHEET: '_JANICE_TEMP_CALC_',
  HEADER_ROW: 1,
  FIRST_DATA_ROW: 2,
  COL_ITEMS: 1,                    // Column A
  BATCH_SIZE: 100,                 // Consider 300–500 if your environment is stable

  // Markets: each spills 2 columns (buy|sell) on the temp sheet
  MARKETS: [
    { label: 'Jita',  market: 'Jita',  destStartCol: 2 }, // -> B:C
    { label: 'Amarr', market: 'Amarr', destStartCol: 4 }  // -> D:E
  ],

  // JANICE_PRICER fields argument
  FIELDS_ARG: 'immediatePrices.buyPrice|immediatePrices.sellPrice',

  // Waiting for custom function calc on the temp tab (per-batch)
  CALC_MAX_WAIT_MS: 20000,  // per-batch wait budget
  CALC_POLL_INTERVAL_MS: 300,

  // Stop before hard cap (~6 min) so you can re-run without hitting a fatal timeout
  MAX_EXECUTION_MS: 330000, // ~5m30s

  // Default behavior for runPriceFillBatches()
  SKIP_ROWS_WITH_EXISTING_VALUES: false,    // skip rows with non-error values in B:E
  TREAT_ERROR_TEXT_AS_EMPTY_ON_WRITE: true // convert "#N/A"/"#ERROR"/etc. to blanks on write
};

// Logging configuration. Raise or lower verbosity here.
// LEVEL in {'DEBUG','INFO','WARN','ERROR'}
var LOG_CFG = {
  LEVEL: 'DEBUG',   // choose: 'DEBUG' (most verbose), 'INFO', 'WARN', 'ERROR'
  TAG: 'JaniceBatch'
};

/* ====================
 * LOGGING INFRASTRUCTURE
 * ==================== */

/**
 * Map log levels to numeric severity for filtering.
 */
var LOG_LEVELS = { DEBUG: 10, INFO: 20, WARN: 30, ERROR: 40 };

/**
 * Returns true if the given level should be emitted under current LOG_CFG.LEVEL.
 */
function shouldLog_(level) {
  var cur = LOG_LEVELS[String(LOG_CFG.LEVEL || 'INFO').toUpperCase()] || LOG_LEVELS.INFO;
  var want = LOG_LEVELS[String(level || 'INFO').toUpperCase()] || LOG_LEVELS.INFO;
  return want >= cur;
}

/**
 * Uniform, timestamped logging to both console and Logger, with optional metadata.
 * We avoid fancy Unicode to prevent parsing issues.
 */
function log_(level, msg, meta) {
  if (!shouldLog_(level)) return;
  if (!meta) meta = {};
  var ts = new Date().toISOString();
  var line = '[' + LOG_CFG.TAG + '][' + level + '] ' + ts + ' - ' + msg + (Object.keys(meta).length ? ' ' + JSON.stringify(meta) : '');
  console.log(line);
  Logger.log(line);
}

/**
 * Simple timer utility to measure elapsed ms for operations.
 * Usage:
 *   var t = startTimer_('my-step');
 *   ... do work ...
 *   endTimer_(t, 'optional override message');
 */
function startTimer_(label) {
  return { label: label || 'timer', start: Date.now() };
}
function endTimer_(timer, msg, extraMeta) {
  if (!timer) return;
  var ms = Date.now() - timer.start;
  log_('DEBUG', (msg || ('Timer ' + (timer.label || 'timer'))) + ' elapsed_ms=' + ms, extraMeta || { elapsed_ms: ms, label: timer.label });
}

/* =======================
 * PUBLIC ENTRY POINTS
 * ======================= */

/**
 * Fast rerun: only rows that are blank or have obvious error text in B:E.
 * WHY: Huge speed-up after first run; re-run as often as needed.
 */
function runPriceFillBatches() {
  log_('INFO', 'runPriceFillBatches() start', { skipExisting: true });
  runPriceFillBatchesInternal_({
    skipExisting: true,
    rowsOverride: null
  });
  log_('INFO', 'runPriceFillBatches() end');
}

/**
 * Full refresh: recompute all rows regardless of existing values.
 * WHY: Use when you want a fresh sweep (e.g., prices shifted materially).
 */
function runPriceFillBatchesForceAll() {
  log_('INFO', 'runPriceFillBatchesForceAll() start', { skipExisting: false });
  runPriceFillBatchesInternal_({
    skipExisting: false,
    rowsOverride: null
  });
  log_('INFO', 'runPriceFillBatchesForceAll() end');
}

/**
 * Spot refresh: recompute only the currently selected rows (ignores skipExisting).
 * WHY: Fastest for a handful of rows you want to update right now.
 */
function runPriceFillForSelection() {
  var ss = SpreadsheetApp.getActive();
  var sel = ss.getActiveRange();
  if (!sel) {
    SpreadsheetApp.getUi().alert('Select one or more rows first.');
    log_('WARN', 'runPriceFillForSelection() has no active selection');
    return;
  }
  var start = sel.getRow();
  var end = start + sel.getNumRows() - 1;
  var rows = [];
  for (var r = start; r <= end; r++) rows.push(r);
  log_('INFO', 'runPriceFillForSelection() start', { selectionStart: start, selectionEnd: end, count: rows.length });

  runPriceFillBatchesInternal_({
    skipExisting: false,
    rowsOverride: rows
  });

  log_('INFO', 'runPriceFillForSelection() end');
}

/* =======================
 * CORE IMPLEMENTATION
 * ======================= */

/**
 * Shared worker. Options:
 *  - skipExisting: boolean (skip rows with non-error values)
 *  - rowsOverride: array|null (absolute row numbers; if null, compute from A2:A)
 */
function runPriceFillBatchesInternal_(opts) {
  var execTimer = startTimer_('total-execution');
  var cfg = PRICE_CFG;
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(cfg.MAIN_SHEET);
  if (!sh) {
    log_('ERROR', 'Main sheet not found', { sheet: cfg.MAIN_SHEET });
    SpreadsheetApp.getUi().alert('Sheet "' + cfg.MAIN_SHEET + '" not found.');
    endTimer_(execTimer, 'total-execution (error: main sheet missing)');
    return;
  }

  var startExec = Date.now();
  log_('INFO', 'Plan start', {
    mainSheet: cfg.MAIN_SHEET,
    firstDataRow: cfg.FIRST_DATA_ROW,
    batchSize: cfg.BATCH_SIZE,
    skipExisting: !!opts.skipExisting,
    markets: cfg.MARKETS.map(function(m){return m.label;})
  });

  // Determine contiguous data region in Column A starting at A2 (stop at first blank).
  var boundsTimer = startTimer_('compute-data-bounds');
  var lastRow = sh.getLastRow();
  if (lastRow < cfg.FIRST_DATA_ROW) {
    log_('WARN', 'No data rows found in Column A.');
    SpreadsheetApp.getUi().alert('No item names found.');
    endTimer_(boundsTimer, 'compute-data-bounds (no data)');
    endTimer_(execTimer, 'total-execution (no data)');
    return;
  }

  var dataRowsCount = lastRow - cfg.HEADER_ROW;
  var itemsRange = sh.getRange(cfg.FIRST_DATA_ROW, cfg.COL_ITEMS, dataRowsCount, 1);
  var items = itemsRange.getValues();
  var lastDataRow = cfg.HEADER_ROW;
  for (var i = 0; i < items.length; i++) {
    var v = String(items[i][0] || '').trim();
    if (v) lastDataRow = cfg.FIRST_DATA_ROW + i;
    else break;
  }
  if (lastDataRow < cfg.FIRST_DATA_ROW) {
    log_('WARN', 'Column A empty below header.');
    SpreadsheetApp.getUi().alert('No item names found.');
    endTimer_(boundsTimer, 'compute-data-bounds (empty after header)');
    endTimer_(execTimer, 'total-execution (empty after header)');
    return;
  }
  endTimer_(boundsTimer, 'compute-data-bounds', { lastRow: lastRow, lastDataRow: lastDataRow });

  // Decide which rows to process.
  var planTimer = startTimer_('plan-rows');
  var rowsToProcess;
  if (opts.rowsOverride && opts.rowsOverride.length) {
    rowsToProcess = sanitizeRowsOverride_(opts.rowsOverride, cfg.FIRST_DATA_ROW, lastDataRow);
    log_('INFO', 'Using rowsOverride', { count: rowsToProcess.length, first: rowsToProcess[0], last: rowsToProcess[rowsToProcess.length - 1] });
  } else {
    rowsToProcess = computeRowsToProcess_(sh, cfg.FIRST_DATA_ROW, lastDataRow, opts.skipExisting);
    log_('INFO', 'Rows to process computed', { count: rowsToProcess.length, skippedExisting: !!opts.skipExisting });
  }

  if (rowsToProcess.length === 0) {
    log_('INFO', 'Nothing to do', { reason: opts.skipExisting ? 'all rows have values' : 'rowsOverride empty' });
    SpreadsheetApp.getUi().alert('No rows to process.');
    endTimer_(planTimer, 'plan-rows (nothing to do)');
    endTimer_(execTimer, 'total-execution (nothing to do)');
    return;
  }

  var spans = toContiguousSpans_(rowsToProcess);
  endTimer_(planTimer, 'plan-rows', { spans: spans.length });

  // Prepare temp sheet and ensure minimal columns for spills.
  var tempTimer = startTimer_('prepare-temp-sheet');
  var temp = getOrCreateHiddenTempSheet_(ss, cfg.TEMP_SHEET);
  ensureTempColumns_(temp, 4);
  endTimer_(tempTimer, 'prepare-temp-sheet', { tempSheet: cfg.TEMP_SHEET });

  var totalProcessed = 0;

  outer:
  for (var s = 0; s < spans.length; s++) {
    var span = spans[s]; // {start, end}
    log_('INFO', 'Span start', { spanIndex: s + 1, start: span.start, end: span.end, spanRows: (span.end - span.start + 1) });

    for (var start = span.start; start <= span.end; start += cfg.BATCH_SIZE) {
      var end = Math.min(start + cfg.BATCH_SIZE - 1, span.end);
      var batchSize = end - start + 1;

      // Time slicing: stop early before we hit the hard cap; user can re-run to continue.
      var elapsed = Date.now() - startExec;
      if (elapsed > cfg.MAX_EXECUTION_MS) {
        log_('WARN', 'Stopping early to avoid hard timeout', { processed: totalProcessed, elapsed_ms: elapsed });
        SpreadsheetApp.getUi().alert('Paused to avoid the time limit. Processed ' + totalProcessed + ' row(s). Run again to continue.');
        break outer;
      }

      var batchTimer = startTimer_('batch ' + start + '..' + end);
      log_('INFO', 'Batch begin', { startRow: start, endRow: end, batchSize: batchSize });

      // 1) Prepare temp for this batch: clear A..E in used rows and write item names as Plain text.
      var prepTimer = startTimer_('batch-prep-temp');
      var batchItemsRangeMain = sh.getRange(start, cfg.COL_ITEMS, batchSize, 1);
      var batchItems = batchItemsRangeMain.getValues();
      var tempAreaToClear = temp.getRange(1, 1, batchSize, 5); // A..E for current batch height
      tempAreaToClear.clearContent();
      var tempInput = temp.getRange(1, 1, batchSize, 1);
      tempInput.setNumberFormat('@'); // ensure leading apostrophes are preserved as characters
      tempInput.setValues(batchItems);
      endTimer_(prepTimer, 'batch-prep-temp', {
        mainA1: batchItemsRangeMain.getA1Notation(),
        tempInputA1: tempInput.getA1Notation(),
        tempClearedA1: tempAreaToClear.getA1Notation()
      });

      // 2) Write JANICE_PRICER formulas (3-arg) for both markets on temp (only top-left cells).
      var formulaTimer = startTimer_('batch-write-formulas');
      for (var m = 0; m < cfg.MARKETS.length; m++) {
        var mk = cfg.MARKETS[m];
        var destStartColInTemp = 2 + (m * 2); // B for Jita, D for Amarr
        var destTopLeft = temp.getRange(1, destStartColInTemp);
        var itemsA1 = temp.getRange(1, 1, batchSize, 1).getA1Notation();
        var formula = '=JANICE_PRICER(' + itemsA1 + ',"' + cfg.FIELDS_ARG + '","' + mk.market + '")';
        destTopLeft.setFormula(formula);
        log_('DEBUG', 'Set formula', { market: mk.market, destA1: destTopLeft.getA1Notation(), itemsA1: itemsA1 });
      }
      endTimer_(formulaTimer, 'batch-write-formulas');

      // 3) Wait for calculation (bounded per batch).
      var waitTimer = startTimer_('batch-wait-calc');
      var tempDestRanges = [];
      for (var m2 = 0; m2 < cfg.MARKETS.length; m2++) {
        tempDestRanges.push(temp.getRange(1, 2 + (m2 * 2), batchSize, 2)); // B:C or D:E
      }
      waitForCalculationOrTimeout_(tempDestRanges, cfg.CALC_MAX_WAIT_MS, cfg.CALC_POLL_INTERVAL_MS);
      endTimer_(waitTimer, 'batch-wait-calc', {
        destA1s: tempDestRanges.map(function(r){ return r.getA1Notation(); })
      });

      // Snapshot results once (avoid repeated getValues calls).
      var snapTimer = startTimer_('batch-snapshot');
      var marketMatrices = [];
      for (var m3 = 0; m3 < cfg.MARKETS.length; m3++) {
        var mat = temp.getRange(1, 2 + (m3 * 2), batchSize, 2).getValues();
        marketMatrices.push(mat);
      }
      endTimer_(snapTimer, 'batch-snapshot', { markets: cfg.MARKETS.map(function(m){ return m.label; }) });

      // 4) Retry unresolved rows — pass 1: toggle a single leading apostrophe.
      var retry1Timer = startTimer_('batch-retry-pass1-toggle-leading');
      var unresolvedIdxs = collectErrorRowIndexesFromMatrices_(marketMatrices);
      log_('DEBUG', 'Unresolved count after first calc', { unresolvedCount: unresolvedIdxs.length });
      if (unresolvedIdxs.length > 0) {
        var changed1 = adjustBatchItemsToggleLeading_(batchItems, unresolvedIdxs);
        log_('DEBUG', 'Toggle-leading applied', { changedRows: changed1 ? unresolvedIdxs.length : 0 });
        if (changed1) {
          tempInput.setNumberFormat('@');
          tempInput.setValues(batchItems);

          // Reapply formulas once to recalc
          for (var m4 = 0; m4 < cfg.MARKETS.length; m4++) {
            var itemsA1b = temp.getRange(1, 1, batchSize, 1).getA1Notation();
            var destA1b = temp.getRange(1, 2 + (m4 * 2)).getA1Notation();
            var formulaB = '=JANICE_PRICER(' + itemsA1b + ',"' + cfg.FIELDS_ARG + '","' + cfg.MARKETS[m4].market + '")';
            temp.getRange(destA1b).setFormula(formulaB);
          }
          waitForCalculationOrTimeout_(tempDestRanges, cfg.CALC_MAX_WAIT_MS, cfg.CALC_POLL_INTERVAL_MS);

          // Refresh matrices and recompute unresolvedIdxs
          marketMatrices = [];
          for (var m5 = 0; m5 < cfg.MARKETS.length; m5++) {
            marketMatrices.push(temp.getRange(1, 2 + (m5 * 2), batchSize, 2).getValues());
          }
          unresolvedIdxs = collectErrorRowIndexesFromMatrices_(marketMatrices);
          log_('DEBUG', 'Unresolved count after retry pass 1', { unresolvedCount: unresolvedIdxs.length });
        }
      }
      endTimer_(retry1Timer, 'batch-retry-pass1-toggle-leading');

      // 5) Retry unresolved rows — pass 2: normalize curly quotes to straight apostrophes.
      var retry2Timer = startTimer_('batch-retry-pass2-normalize-curly-quotes');
      if (unresolvedIdxs.length > 0) {
        var changed2 = adjustBatchItemsNormalizeQuotes_(batchItems, unresolvedIdxs);
        log_('DEBUG', 'Normalize-curly-quotes applied', { changedRows: changed2 ? unresolvedIdxs.length : 0 });
        if (changed2) {
          tempInput.setNumberFormat('@');
          tempInput.setValues(batchItems);
          for (var m6 = 0; m6 < cfg.MARKETS.length; m6++) {
            var itemsA1c = temp.getRange(1, 1, batchSize, 1).getA1Notation();
            var destA1c = temp.getRange(1, 2 + (m6 * 2)).getA1Notation();
            var formulaC = '=JANICE_PRICER(' + itemsA1c + ',"' + cfg.FIELDS_ARG + '","' + cfg.MARKETS[m6].market + '")';
            temp.getRange(destA1c).setFormula(formulaC);
          }
          waitForCalculationOrTimeout_(tempDestRanges, cfg.CALC_MAX_WAIT_MS, cfg.CALC_POLL_INTERVAL_MS);

          // Refresh matrices again
          marketMatrices = [];
          for (var m7 = 0; m7 < cfg.MARKETS.length; m7++) {
            marketMatrices.push(temp.getRange(1, 2 + (m7 * 2), batchSize, 2).getValues());
          }
          var remaining = collectErrorRowIndexesFromMatrices_(marketMatrices).length;
          log_('DEBUG', 'Unresolved count after retry pass 2', { unresolvedCount: remaining });
        }
      }
      endTimer_(retry2Timer, 'batch-retry-pass2-normalize-curly-quotes');

      // 6) Optionally blank error-like strings and write to main in one shot per market.
      var writeTimer = startTimer_('batch-write-main');
      for (var w = 0; w < cfg.MARKETS.length; w++) {
        var out = marketMatrices[w];
        if (cfg.TREAT_ERROR_TEXT_AS_EMPTY_ON_WRITE) {
          out = sanitizeErrorStringsToBlank_(out);
        }
        var destRange = sh.getRange(start, cfg.MARKETS[w].destStartCol, batchSize, 2);
        destRange.setValues(out);
        log_('DEBUG', 'Wrote values to main', {
          market: cfg.MARKETS[w].label,
          destA1: destRange.getA1Notation(),
          rows: batchSize
        });
      }
      endTimer_(writeTimer, 'batch-write-main');

      totalProcessed += batchSize;
      log_('INFO', 'Batch end (values only)', { startRow: start, endRow: end, processedSoFar: totalProcessed });

      // 7) Light clean for next batch
      var cleanTimer = startTimer_('batch-clean-temp');
      temp.getRange(1, 2, batchSize, 4).clearContent();
      endTimer_(cleanTimer, 'batch-clean-temp');

      endTimer_(batchTimer, 'batch ' + start + '..' + end);
      Utilities.sleep(60); // tiny pause can improve stability on very large sheets
    }
  }

  log_('INFO', 'Run complete', { processed: totalProcessed });
  SpreadsheetApp.getUi().alert('Processed ' + totalProcessed + ' row(s). Run again to continue, or use Force All/Selection as needed.');
  endTimer_(execTimer, 'total-execution');
}

/* ==========================
 * ROW SELECTION + PLANNING
 * ========================== */

/**
 * If skipExisting is true, include only rows with blanks or obvious error text in B:E.
 * Otherwise, include every row in [startRow..endRow].
 * WHY: Avoid recomputing clean rows for speed.
 */
function computeRowsToProcess_(sh, startRow, endRow, skipExisting) {
  var t = startTimer_('computeRowsToProcess_');
  var rows = [];
  if (!skipExisting) {
    for (var r = startRow; r <= endRow; r++) rows.push(r);
    endTimer_(t, 'computeRowsToProcess_', { rows: rows.length, skippedExisting: false });
    return rows;
  }
  var numRows = endRow - startRow + 1;
  var beRange = sh.getRange(startRow, 2, numRows, 4); // B..E
  var be = beRange.getValues();
  for (var i = 0; i < numRows; i++) {
    var rowVals = be[i];
    var needs = false;
    for (var c = 0; c < rowVals.length; c++) {
      var v = rowVals[c];
      if (v === '' || v === null) { needs = true; break; }
      var s = String(v);
      if (s.indexOf('#N/A') === 0 || s.indexOf('#ERROR') === 0 || s.indexOf('#VALUE') === 0 || s.indexOf('#REF') === 0) {
        needs = true; break;
      }
    }
    if (needs) rows.push(startRow + i);
  }
  endTimer_(t, 'computeRowsToProcess_', { rows: rows.length, skippedExisting: true, scanA1: beRange.getA1Notation() });
  return rows;
}

/**
 * Clip and dedupe explicit row list, ensuring it falls within [startRow..endRow], sorted asc.
 */
function sanitizeRowsOverride_(rows, startRow, endRow) {
  var t = startTimer_('sanitizeRowsOverride_');
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i] | 0;
    if (r >= startRow && r <= endRow) seen[r] = true;
  }
  var out = [];
  for (var k in seen) out.push(parseInt(k, 10));
  out.sort(function(a, b){ return a - b; });
  endTimer_(t, 'sanitizeRowsOverride_', { rows: out.length, first: out[0], last: out[out.length - 1] });
  return out;
}

/** Convert a set of row numbers into contiguous spans [{start, end}, ...]. */
function toContiguousSpans_(rows) {
  var t = startTimer_('toContiguousSpans_');
  if (rows.length === 0) {
    endTimer_(t, 'toContiguousSpans_', { spans: 0 });
    return [];
  }
  var spans = [];
  var s = rows[0], prev = rows[0];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i] === prev + 1) {
      prev = rows[i];
    } else {
      spans.push({ start: s, end: prev });
      s = rows[i]; prev = rows[i];
    }
  }
  spans.push({ start: s, end: prev });
  endTimer_(t, 'toContiguousSpans_', { spans: spans.length });
  return spans;
}

/* ==========================
 * TEMP SHEET + CALC HELPERS
 * ========================== */

/**
 * Get or create a hidden temp sheet. It stays hidden and is reused run-to-run.
 * WHY: We evaluate JANICE_PRICER off-sheet to keep the main sheet formula-free.
 */
function getOrCreateHiddenTempSheet_(ss, name) {
  var t = startTimer_('getOrCreateHiddenTempSheet_');
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.hideSheet();
    if (sh.getMaxColumns() < 6) sh.insertColumnsAfter(1, 6 - sh.getMaxColumns());
    if (sh.getMaxRows() < 1000) sh.insertRowsAfter(1, 1000 - sh.getMaxRows());
    log_('INFO', 'Created hidden temp sheet', { name: name, cols: sh.getMaxColumns(), rows: sh.getMaxRows() });
  } else if (!sh.isSheetHidden()) {
    sh.hideSheet();
    log_('DEBUG', 'Hid existing temp sheet', { name: name });
  } else {
    log_('DEBUG', 'Reusing existing hidden temp sheet', { name: name, cols: sh.getMaxColumns(), rows: sh.getMaxRows() });
  }
  endTimer_(t, 'getOrCreateHiddenTempSheet_', { name: name });
  return sh;
}

/**
 * Ensure the temp sheet has at least minCols columns (two markets x 2 cols = 4).
 */
function ensureTempColumns_(sheet, minCols) {
  var t = startTimer_('ensureTempColumns_');
  var cols = sheet.getMaxColumns();
  if (cols < minCols) {
    sheet.insertColumnsAfter(cols, (minCols - cols));
    log_('INFO', 'Expanded temp sheet columns', { from: cols, to: sheet.getMaxColumns() });
  }
  endTimer_(t, 'ensureTempColumns_', { minCols: minCols, currentCols: sheet.getMaxColumns() });
}

/**
 * Wait until all provided ranges look calculated or we hit the per-batch timeout.
 * "Calculated" = non-blank and not an obvious error string.
 */
function waitForCalculationOrTimeout_(ranges, maxWaitMs, intervalMs) {
  var t = startTimer_('waitForCalculationOrTimeout_');
  var start = Date.now();
  var loops = 0;
  while (true) {
    if (allRangesLookCalculated_(ranges)) break;
    if (Date.now() - start > maxWaitMs) {
      log_('WARN', 'Per-batch calc wait timed out', { waited_ms: Date.now() - start, polls: loops });
      break;
    }
    Utilities.sleep(intervalMs);
    loops++;
  }
  endTimer_(t, 'waitForCalculationOrTimeout_', { polls: loops, waited_ms: Date.now() - start });
}

/**
 * Heuristic: every cell in every range must be non-blank and not an obvious error string.
 */
function allRangesLookCalculated_(ranges) {
  for (var i = 0; i < ranges.length; i++) {
    var vals = ranges[i].getValues();
    for (var r = 0; r < vals.length; r++) {
      for (var c = 0; c < vals[r].length; c++) {
        var v = vals[r][c];
        if (v === '' || v === null) return false;
        var s = String(v);
        if (s.indexOf('#N/A') === 0 || s.indexOf('#ERROR') === 0 || s.indexOf('#VALUE') === 0 || s.indexOf('#REF') === 0) {
          return false;
        }
      }
    }
  }
  return true;
}

/* ==========================
 * ERROR + VALUE HELPERS
 * ========================== */

/**
 * From an array of matrices (one per market), return 0-based row indexes that are unresolved.
 * "Unresolved" means any market returned blank or an error-like string for that row.
 */
function collectErrorRowIndexesFromMatrices_(matrices) {
  var idxs = [];
  var height = matrices[0].length;
  for (var r = 0; r < height; r++) {
    var hasErr = false;
    for (var m = 0; m < matrices.length; m++) {
      var v1 = matrices[m][r][0];
      var v2 = matrices[m][r][1];
      if (isErrorish_(v1) || isErrorish_(v2)) { hasErr = true; break; }
    }
    if (hasErr) idxs.push(r);
  }
  return idxs;
}

function isErrorish_(v) {
  if (v === '' || v === null) return true;
  var s = String(v);
  return (s.indexOf('#N/A') === 0 || s.indexOf('#ERROR') === 0 || s.indexOf('#VALUE') === 0 || s.indexOf('#REF') === 0);
}

/** Replace error-looking strings with blanks so the main sheet stays clean. */
function sanitizeErrorStringsToBlank_(matrix) {
  for (var r = 0; r < matrix.length; r++) {
    for (var c = 0; c < matrix[r].length; c++) {
      if (isErrorish_(matrix[r][c])) matrix[r][c] = '';
    }
  }
  return matrix;
}

/* ==========================
 * NAME RETRY HELPERS
 * ========================== */

/**
 * Retry pass 1: toggle a single leading apostrophe on unresolved names ONLY.
 * - If the name starts with "'", remove the first one.
 * - If it does not, add one at the start.
 * WHY: Some sources match with a literal leading apostrophe; others match without it.
 * NOTE: This adjusts the temporary batch copy only; your main sheet remains unchanged.
 */
function adjustBatchItemsToggleLeading_(batchItems, idxs) {
  var changed = false;
  for (var i = 0; i < idxs.length; i++) {
    var r = idxs[i];
    var name = String(batchItems[r][0] || '');
    if (name.length === 0) continue;
    if (name.charAt(0) === "'") {
      batchItems[r][0] = name.slice(1); // strip one leading apostrophe
      changed = true;
      log_('DEBUG', 'Stripped leading apostrophe for retry', { rowInBatch: r, original: name, updated: batchItems[r][0] });
    } else {
      batchItems[r][0] = "'" + name;    // add one leading apostrophe
      changed = true;
      log_('DEBUG', 'Added leading apostrophe for retry', { rowInBatch: r, original: name, updated: batchItems[r][0] });
    }
  }
  return changed;
}

/**
 * Retry pass 2: normalize curly single quotes to straight ASCII apostrophes.
 * Example: “‘Wetu’ Mobile Depot” -> "'Wetu' Mobile Depot".
 * NOTE: Temp batch only; main sheet is not modified by this.
 */
function adjustBatchItemsNormalizeQuotes_(batchItems, idxs) {
  var changed = false;
  for (var i = 0; i < idxs.length; i++) {
    var r = idxs[i];
    var name = String(batchItems[r][0] || '');
    var norm = name.replace(/[‘’]/g, "'");
    if (norm !== name) {
      batchItems[r][0] = norm;
      changed = true;
      log_('DEBUG', 'Normalized curly quotes for retry', { rowInBatch: r, original: name, updated: norm });
    }
  }
  return changed;
}


























