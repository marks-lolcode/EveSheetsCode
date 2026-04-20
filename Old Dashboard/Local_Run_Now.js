// ==========================================================
// FILE: Local_Run_Now.gs
// WHAT: Local shim you can run from the editor to perform
//       the legacy single-shot import with consistent logs.
// WHY:  Convenience: run and read logs inside Apps Script.
// ==========================================================

function UpdatePIFullnessNow() {
  // WHAT: Calls the core single-shot import using the same logger style.
  // WHY:  Quick run inside the project; mirrors web app behavior.
  const L = newLogger();
  console.log('[LOCAL] UpdatePIFullnessNow() start');

  try {
    const result = updatePIFullness_(L);
    const imported = (result && result.rows) ? result.rows : 0;

    SpreadsheetApp.getActive().toast('PI import complete: ' + imported + ' rows', 'PI Tools', 5);
    console.log('[LOCAL] Result: ' + JSON.stringify(result));

    const lines = L.dump();
    if (lines && lines.length) {
      const block = lines.join('\n');
      Logger.log('[LOCAL] Detailed logs:\n' + block);
      console.log('[LOCAL] Detailed logs:\n' + block);
    }
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    SpreadsheetApp.getActive().toast('updatePIFullness failed: ' + msg, 'PI Tools', 7);
    console.error('[LOCAL] ERROR: ' + msg);
    Logger.log('[LOCAL] ERROR: ' + msg);
    throw err;
  } finally {
    console.log('[LOCAL] UpdatePIFullnessNow() end');
  }
}
