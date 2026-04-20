/**
 * AllInv_Move_RenameSwap
 * Builds a new sheet from AllInventory__STAGING, then renames NEW->AllInventory and deletes OLD.
 * This avoids resizing the existing AllInventory grid and prevents giant recalcs.
 */
function AllInv_Move_RenameSwap() {
  const L = (typeof newLogger === 'function')
    ? newLogger()
    : { log: (m, d) => console.log(d === undefined ? String(m) : (String(m) + ' ' + JSON.stringify(d))) };

  const STAGING = 'AllInventory__STAGING';
  const LIVE    = 'AllInventory';
  const NEW     = 'AllInventory__NEW';
  const OLD     = 'AllInventory__OLD';

  L.log('[MoveSwap] Start');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const staging = ss.getSheetByName(STAGING);
  if (!staging) {
    L.log('[MoveSwap] ERROR: staging sheet missing', { STAGING });
    throw new Error('Staging sheet not found: ' + STAGING);
  }

  const rows = staging.getLastRow();
  const cols = staging.getLastColumn();
  if (!rows || !cols) {
    L.log('[MoveSwap] Staging empty; aborting to protect live sheet.');
    return { ok: false, reason: 'empty_staging' };
  }

  const values = staging.getRange(1, 1, rows, cols).getValues();
  L.log('[MoveSwap] Read staging', { rows, cols });

  // Clean up any prior NEW/OLD to avoid name conflicts.
  const priorNew = ss.getSheetByName(NEW);
  if (priorNew) { ss.deleteSheet(priorNew); L.log('[MoveSwap] Deleted prior NEW'); }
  const priorOld = ss.getSheetByName(OLD);
  if (priorOld) { ss.deleteSheet(priorOld); L.log('[MoveSwap] Deleted prior OLD'); }

  // Build NEW from scratch (no dependencies pointing to it yet → no heavy recalc).
  const newSheet = ss.insertSheet(NEW);
  // Size exactly to data block
  if (newSheet.getMaxRows() < rows) newSheet.insertRowsAfter(newSheet.getMaxRows(), rows - newSheet.getMaxRows());
  if (newSheet.getMaxColumns() < cols) newSheet.insertColumnsAfter(newSheet.getMaxColumns(), cols - newSheet.getMaxColumns());
  if (newSheet.getMaxRows() > rows) newSheet.deleteRows(rows + 1, newSheet.getMaxRows() - rows);
  if (newSheet.getMaxColumns() > cols) newSheet.deleteColumns(cols + 1, newSheet.getMaxColumns() - cols);

  newSheet.getRange(1, 1, rows, cols).setValues(values);
  SpreadsheetApp.flush();
  L.log('[MoveSwap] Wrote NEW data', { rows, cols });

  // Swap names atomically
  const live = ss.getSheetByName(LIVE);
  if (live) { live.setName(OLD); L.log('[MoveSwap] LIVE → OLD'); }
  newSheet.setName(LIVE);
  L.log('[MoveSwap] NEW → LIVE');

  // Remove OLD (or comment this out to keep a rollback)
  const oldNow = ss.getSheetByName(OLD);
  if (oldNow) { ss.deleteSheet(oldNow); L.log('[MoveSwap] Deleted OLD'); }

  // Keep staging hidden
  try { staging.hideSheet(); } catch (_) {}

  L.log('[MoveSwap] Done');
  return { ok: true, rows, cols, method: 'rename_swap' };
}


function _test_AllInv_Move_RenameSwap() {
  const r = AllInv_Move_RenameSwap();
  Logger.log(JSON.stringify(r));
}