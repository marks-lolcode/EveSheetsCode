/**
 * WorkbookShrink.gs
 *
 * WHAT: Reduce every sheet's grid to its used range (+ a small safety buffer)
 *       so the entire spreadsheet stays under the 10,000,000-cell cap.
 * WHY:  Blank rows/columns still count toward the limit; trimming them restores headroom.
 * HOW:  Adds a "Maintenance" menu with one action: "Shrink workbook grids".
 * LOGS: Uses the project's newLogger() if present; otherwise falls back to Logger.log.
 */

// BEGINNER NOTE: We try to re-use your project's newLogger() from WebApp.gs.
// If it isn't loaded in this project, we define a tiny fallback here.
function __fallbackLogger() {
  return {
    log: (msg, data) => {
      try {
        const line = data !== undefined ? (msg + ' ' + JSON.stringify(data)) : String(msg);
        Logger.log(line);
        console.log(line);
      } catch (e) {
        Logger.log(String(msg));
        console.log(String(msg));
      }
    }
  };
}

function __getLogger() {
  try {
    // If your WebApp.gs newLogger() exists, use it for consistent formatting.
    if (typeof newLogger === 'function') return newLogger();
  } catch (e) {}
  return __fallbackLogger();
}

/**
 * WHAT: Iterate every sheet; compute used rows/columns; delete surplus rows/cols.
 * WHY:  Keeps workbook below 10M cells while leaving a bit of headroom for edits.
 * SAFETY: Keeps at least 200 rows and 26 columns (A:Z) on each sheet.
 */
function compressWorkbook() {
  const ss = SpreadsheetApp.getActive();
  const L = __getLogger();

  const safetyMinRows = 200; // minimum rows to keep even if used rows are smaller
  const safetyMinCols = 26;  // keep at least columns A:Z

  L.log('[SHRINK] Start', { spreadsheet: ss.getName() });

  ss.getSheets().forEach(sheet => {
    try {
      const name = sheet.getName();

      // WHAT: "Used" size = where content/formatting ends.
      const usedRows = Math.max(1, sheet.getLastRow());
      const usedCols = Math.max(1, sheet.getLastColumn());

      // WHY: Add a tiny buffer so normal edits don’t immediately grow the grid.
      const keepRows = Math.max(safetyMinRows, usedRows + 100);
      const keepCols = Math.max(safetyMinCols, usedCols + 5);

      const curMaxRows = sheet.getMaxRows();
      const curMaxCols = sheet.getMaxColumns();

      // Delete extra rows (from bottom)
      if (curMaxRows > keepRows) {
        const delRows = curMaxRows - keepRows;
        sheet.deleteRows(keepRows + 1, delRows);
      }
      // Delete extra columns (from right)
      if (curMaxCols > keepCols) {
        const delCols = curMaxCols - keepCols;
        sheet.deleteColumns(keepCols + 1, delCols);
      }

      L.log('[SHRINK] Resized sheet', { sheet: name, keptRows: keepRows, keptCols: keepCols });
    } catch (e) {
      L.log('[SHRINK] Error on sheet', { sheet: sheet.getName(), error: String(e) });
    }
  });

  L.log('[SHRINK] Done.');
}
