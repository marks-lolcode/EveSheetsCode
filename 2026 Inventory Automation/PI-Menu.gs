// ==========================================================
// FILE: PI-Menu.gs
// WHAT: "PI Factory" spreadsheet menu + thin UI runners. Each runner toasts,
//        calls a core worker in PI-Factory.gs (or refreshAssets_ in Assets-GESI.gs)
//        with a fresh logger, and reports via toast. Errors go to Executions->Logs.
//
// NOTE: Schematic upload is NOT a menu item - it is the manual PC-side
//       `Refresh-Inventory.ps1 -Sde` run (recipes change ~once per expansion).
// ==========================================================

// Menu grouped by cadence. Submenus tell you HOW OFTEN + in what ORDER to run.
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // DAILY — run top to bottom. Each step feeds the next.
  const daily = ui.createMenu('Daily  (run in order)')
    .addItem('▶ Run ALL daily (1-6)', 'pi_RunDaily')
    .addSeparator()
    .addItem('1. Refresh Assets', 'pi_RefreshAssets')
    .addItem('2. Refresh Cost Accounting (Wallet→WAC)', 'pi_RefreshCostAccounting')
    .addItem('3. Build PI Profit', 'pi_BuildProfit')
    .addItem('4. Build PI Inventory Grid', 'pi_BuildInventory')
    .addItem('5. Build Per-Char Buy List', 'pi_BuildBuyListByChar')
    .addItem('5b. Build Buy List By Item', 'pi_BuildBuyListByItem')
    .addItem('6. Check Sell Floor', 'pi_CheckSellFloor');

  // PERIODIC — when planet layout / what each char produces changes.
  const periodic = ui.createMenu('Periodic  (when production changes)')
    .addItem('Refresh Planet Production', 'pi_RefreshPlanetProd')
    .addItem('Suggest Planet Swaps', 'pi_SuggestSwaps')
    .addItem('Build Input Usage Lists', 'pi_BuildUsageLists')
    .addItem('Build Buy List (manual plan)', 'pi_BuildBuyList')
    .addItem('Refresh Colonies (finish forecast)', 'pi_RefreshColonies');

  // AFTER SDE UPDATE — only after Refresh-Inventory.ps1 -Sde changes recipes/volumes.
  const sde = ui.createMenu('After SDE update')
    .addItem('Build Run Sizes', 'pi_BuildRunSizes')
    .addItem('Rebuild PI Profit', 'pi_BuildProfit');

  // SETUP — rare: new alts, skill changes, new locations, illiquid flags.
  const setup = ui.createMenu('Setup / occasional')
    .addItem('Refresh Character List', 'pi_RefreshCharList')
    .addItem('Refresh Config from ESI', 'pi_RefreshConfig')
    .addItem('List PI Locations', 'pi_ListLocations');

  ui.createMenu('PI Factory')
    .addSubMenu(daily)
    .addSubMenu(periodic)
    .addSubMenu(sde)
    .addSubMenu(setup)
    .addSeparator()
    .addItem('Start Batch (ad-hoc)', 'pi_StartBatch')
    .addToUi();
}

/* =========================
   === RUNNER PLUMBING   ===
   ========================= */

function pi_toast_(msg, secs) { SpreadsheetApp.getActive().toast(msg, 'PI Factory', secs || 5); }

// Wrap a core worker: clear START/END markers in Executions->Logs (which action,
// when it began, when it finished + elapsed), plus start/result toasts.
function pi_run_(label, fn) {
  const t0 = Date.now();
  const L = newLogger();
  const fnName = (fn && fn.name) ? fn.name : '(anonymous)';
  L.log('======== START: ' + label + ' [' + fnName + '] ========');
  pi_toast_(label + ' started...', 5);
  try {
    const result = fn(L);
    const ms = Date.now() - t0;
    L.log('======== END: ' + label + ' [' + fnName + '] (' + ms + 'ms) ========', result || {});
    pi_toast_(label + ' done (' + ms + 'ms): ' + pi_summary_(result), 8);
    return result;
  } catch (err) {
    const ms = Date.now() - t0;
    L.log('======== FAILED: ' + label + ' [' + fnName + '] (' + ms + 'ms) ========', { error: String(err) });
    pi_toast_(label + ' FAILED - see Executions -> Logs', 8);
    throw err;
  }
}

function pi_summary_(o) {
  if (!o || typeof o !== 'object') return String(o);
  return Object.keys(o).map(k => k + '=' + o[k]).join(', ');
}

/* =========================
   === RUNNERS           ===
   ========================= */

function pi_RefreshAssets()  { pi_run_('Refresh Assets', refreshAssets_); }
function pi_RefreshCharList() { pi_run_('Refresh Character List', refreshCharacterList_); }
function pi_RefreshConfig()  { pi_run_('Refresh Config from ESI', refreshConfigFromEsi_); }
function pi_ListLocations()  { pi_run_('List PI Locations', listPiLocations_); }
function pi_BuildProfit()    { pi_run_('Build PI Profit', buildPiProfit_); }
function pi_BuildInventory() { pi_run_('Build PI Inventory Grid', buildPiInventoryGrid_); }
function pi_BuildRunSizes()  { pi_run_('Build Run Sizes', buildRunSizes_); }
function pi_RefreshPlanetProd() { pi_run_('Refresh Planet Production', buildPiPlanetProduction_); }
function pi_SuggestSwaps()   { pi_run_('Suggest Planet Swaps', suggestReplacements_); }
function pi_BuildBuyList()   { pi_run_('Build Buy List', buildPiBuyList_); }
function pi_BuildBuyListByChar() { pi_run_('Build Per-Char Buy List', buildPiBuyListByChar_); }
function pi_BuildBuyListByItem() { pi_run_('Build Buy List By Item', buildPiBuyListByItem_); }
function pi_BuildUsageLists() { pi_run_('Build Input Usage Lists', buildInputUsageLists_); }

// Run the 6 daily steps in order. Continue-on-error (a failed step is logged and
// skipped, later steps still run on last-good sheet data). One shared logger, clear
// per-step + overall START/END markers. WARNING: heavy chain — on a consumer Google
// account the 6-min script limit can bite; if it times out, run steps individually.
function pi_RunDaily() {
  const steps = [
    ['1. Refresh Assets', refreshAssets_],
    ['2. Refresh Cost Accounting', refreshCostAccounting_],
    ['3. Build PI Profit', buildPiProfit_],
    ['4. Build PI Inventory Grid', buildPiInventoryGrid_],
    ['5. Build Per-Char Buy List', buildPiBuyListByChar_],
    ['5b. Build Buy List By Item', buildPiBuyListByItem_],
    ['6. Check Sell Floor', checkSellFloor_],
  ];
  const L = newLogger();
  const t0 = Date.now();
  L.log('######## DAILY RUN START ########');
  pi_toast_('Daily run started (1-6)...', 5);
  const summary = [];
  let failed = 0;
  for (let i = 0; i < steps.length; i++) {
    const label = steps[i][0], fn = steps[i][1];
    const st = Date.now();
    L.log('---- step ' + (i + 1) + '/' + steps.length + ' START: ' + label + ' ----');
    pi_toast_('Daily ' + (i + 1) + '/' + steps.length + ': ' + label, 4);
    try {
      const r = fn(L);
      L.log('---- step ' + (i + 1) + ' OK: ' + label + ' (' + (Date.now() - st) + 'ms) ----', r || {});
      summary.push(label + ' OK');
    } catch (e) {
      failed++;
      L.log('---- step ' + (i + 1) + ' FAIL: ' + label + ' (' + (Date.now() - st) + 'ms) ----', { error: String(e) });
      summary.push(label + ' FAIL');
    }
  }
  const total = Date.now() - t0;
  L.log('######## DAILY RUN END (' + total + 'ms, ' + failed + ' failed) ########', { summary });
  pi_toast_('Daily done (' + Math.round(total / 1000) + 's, ' + failed + ' failed). See Executions->Logs.', 10);
  return { total_ms: total, failed, summary };
}

function pi_StartBatch() {
  const ui = SpreadsheetApp.getUi();
  const pr = ui.prompt('Start Batch', 'Product (schematic name):', ui.ButtonSet.OK_CANCEL);
  if (pr.getSelectedButton() !== ui.Button.OK) return;
  const product = pr.getResponseText().trim();
  if (!product) { pi_toast_('No product entered.', 5); return; }
  const cr = ui.prompt('Start Batch', 'Cycle count for "' + product + '":', ui.ButtonSet.OK_CANCEL);
  if (cr.getSelectedButton() !== ui.Button.OK) return;
  const count = Number(cr.getResponseText().trim()) || 1;
  pi_run_('Start Batch', L => recordBatch_(L, product, count));
}

function pi_CheckSellFloor() { pi_run_('Check Sell Floor', checkSellFloor_); }
function pi_RefreshCostAccounting() { pi_run_('Refresh Cost Accounting', refreshCostAccounting_); }
function pi_RefreshColonies() { pi_run_('Refresh Colonies', buildPiColonies_); }
