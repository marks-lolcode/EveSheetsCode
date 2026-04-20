// ==========================================================
// FILE: PI_Menu.gs
// WHAT: Adds a "PI Import" menu with One-Shot and Chunked
//       runners that call the core workers in WebApp_Core.
// WHY:  Lets you run from the Sheet UI with clear logs.
// ==========================================================

/* =========================
   === LOGGER (UI) ========
   ========================= */

function _getLoggerUI() {
  // WHAT: A minimal logger for UI runs that writes to both logs.
  const buf = [];
  return {
    log: (msg, data) => {
      const line = (data === undefined) ? String(msg) : `${msg} ${_safeJsonUI(data)}`;
      buf.push(line);
      Logger.log(line);
      console.log(line);
    },
    dump: () => buf.slice()
  };
}
function _safeJsonUI(v){ try { return JSON.stringify(v); } catch(e){ return String(v); } }

/* =========================
   === ONE-SHOT RUNNER ====
   ========================= */

function piImport_RunOneShot() {
  // WHAT: Call the core single-shot worker. Good for small CSVs.
  const L = _getLoggerUI();
  const t0 = Date.now();
  SpreadsheetApp.getActive().toast('PI import (one-shot) started…', 'PI Import', 5);
  L.log('[PI] One-shot import start');

  try {
    const result = updatePIFullness_(newLogger()); // reuse core logger format
    const ms = Date.now() - t0;
    L.log('[PI] One-shot import complete', { ms, result });
    SpreadsheetApp.getActive().toast(`Done (one-shot). Rows: ${(result && result.rows) || 'n/a'}`, 'PI Import', 8);
  } catch (err) {
    L.log('[PI] One-shot import FAILED', { error: String(err) });
    SpreadsheetApp.getActive().toast('Import failed – see Executions → Logs', 'PI Import', 8);
    throw err;
  }
}

/* =========================
   === CHUNKED RUNNER  ====
   ========================= */

const UI_CHUNK_ROWS = 1000;          // You can tune this if needed
const UI_INTER_CHUNK_SLEEP_MS = 150; // Small breather between chunks

function piImport_RunChunked() {
  // WHAT: Loop the core chunked worker until done.
  const L = _getLoggerUI();
  const t0 = Date.now();
  let startIndex = 0;
  let cleared = false;
  let totalWritten = 0;
  let loops = 0;
  const maxLoops = 10000;

  SpreadsheetApp.getActive().toast('PI import (chunked) started…', 'PI Import', 5);
  L.log('[PI] Chunked import start', { UI_CHUNK_ROWS });

  try {
    for (; loops < maxLoops; loops++) {
      const body = { startIndex, chunkRows: UI_CHUNK_ROWS, clear: !cleared };
      L.log('[PI] Calling updatePIFullnessChunked_', body);

      const r = updatePIFullnessChunked_(newLogger(), body) || {};
      const wrote = Number(r.wrote || 0);
      const next  = (r.nextIndex !== undefined) ? Number(r.nextIndex) : (startIndex + wrote);
      const done  = (r.done === true);

      totalWritten += wrote;
      if (!cleared) cleared = true;

      L.log('[PI] Chunk result', { loop: loops + 1, wrote, nextIndex: next, done });

      SpreadsheetApp.getActive().toast(`Chunk ${loops + 1}: +${wrote}`, 'PI Import', 3);

      if (done) break;
      if (wrote <= 0 && next <= startIndex) {
        L.log('[PI] No forward progress; stopping defensively', { wrote, startIndex, next });
        break;
      }

      startIndex = next;
      Utilities.sleep(UI_INTER_CHUNK_SLEEP_MS);
    }

    const ms = Date.now() - t0;
    L.log('[PI] Chunked import complete', { loops: loops + 1, totalWritten, elapsed_ms: ms });
    SpreadsheetApp.getActive().toast(`Done (chunked). Rows written: ${totalWritten}`, 'PI Import', 8);

  } catch (err) {
    L.log('[PI] Chunked import FAILED', { error: String(err) });
    SpreadsheetApp.getActive().toast('Import failed – see Executions → Logs', 'PI Import', 8);
    throw err;
  }
}
