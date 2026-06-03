// ==========================================================
// FILE: Assets-GESI.gs
// WHAT: Server-side EVE inventory refresh for the FRESH sheet.
//        Pulls assets for ALL authorized characters via GESI, resolves each
//        item's real location (walking container/ship chains), enriches with
//        static SDE data + Jita prices (Fuzzwork), then atomic-swaps a lean
//        table into the live AllInventory tab.
//
// DEPLOY (clasp): see apps-script/CLASP-SETUP.md.
//   - GESI added as a library (identifier `GESI`) — already in appsscript.json.
//   - Script Property `GESI_TOKEN` = shared secret (also $env:GESI_TOKEN on PC).
//   - Web App: Execute as ME (GESI-authorized owner), access ANYONE_ANONYMOUS.
//
// ROUTES (POST JSON, all require { token }):
//   fn=ping            -> health check
//   fn=refreshAssets   -> full GESI pull + resolve + enrich + stage + swap
//   fn=beginSde        -> clear a staging tab (start of /sde upload)  { target }
//   fn=importSdeRows   -> append a chunk of rows  { target, startIndex, values }
//   fn=commitSde       -> swap staging -> live  { target }
//
// SDE upload targets (built PC-side by Refresh-Inventory.ps1 -Sde):
//   types     -> SDE_Types     (type_id, type_name, group_name, category_name, volume, repackaged_volume)
//   systems   -> SDE_Systems   (system_id, system_name, region_id, security)
//   regions   -> SDE_Regions   (region_id, region_name)
//   stations  -> SDE_Stations  (station_id, system_id)
// ==========================================================

/* =========================
   === CONFIG CONSTANTS  ===
   ========================= */

const BUILD_TAG = 'gesi-assets-2026-06-03';

// Shared secret. Stored in Script Properties (key 'GESI_TOKEN'), NOT in source.
function getToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('GESI_TOKEN');
  if (!t) throw new Error("Script Property 'GESI_TOKEN' is not set (Project Settings -> Script Properties).");
  return t;
}

const SS = SpreadsheetApp.getActiveSpreadsheet();

// Live asset tab + hidden staging twin.
const ASSET_SHEET_NAME   = 'AllInventory';
const ASSET_STAGING_NAME = 'AllInventory__staging';

// Persistent cache of resolved station/structure names (so we don't refetch).
const LOCNAME_SHEET_NAME = 'SDE_LocNames'; // loc_id, name, system_id, kind

// SDE upload targets -> { live, staging } tab names.
const SDE_TARGETS = {
  types:    { live: 'SDE_Types',    staging: 'SDE_Types__staging' },
  systems:  { live: 'SDE_Systems',  staging: 'SDE_Systems__staging' },
  regions:  { live: 'SDE_Regions',  staging: 'SDE_Regions__staging' },
  stations: { live: 'SDE_Stations', staging: 'SDE_Stations__staging' },
};

// GESI endpoints (ESI operationIds).
const EP_ASSETS    = 'characters_character_assets';
const EP_STRUCTURE = 'universe_structures_structure';

// Market data: Fuzzwork aggregates for The Forge (Jita) region.
const FUZZWORK_URL   = 'https://market.fuzzwork.co.uk/aggregates/';
const FORGE_REGION   = 10000002;
const PRICE_CHUNK    = 200;   // type_ids per Fuzzwork call

const RETURN_LOGS = true;

/* =========================
   === LOGGING + RESPONSE ===
   ========================= */

function newLogger() {
  const buf = [];
  return {
    log: (msg, data) => {
      try {
        const line = (data === undefined) ? String(msg) : `${msg} ${JSON.stringify(data)}`;
        buf.push(line); console.log(line);
      } catch (e) { const line = `${msg} [unserializable]`; buf.push(line); console.log(line); }
    },
    dump: () => buf.slice(),
  };
}

function json_(ok, fn, result, L, startedMs) {
  const ms = startedMs ? (Date.now() - startedMs) : 0;
  const out = { ok: !!ok, fn: String(fn || ''), ms, result: result || {} };
  if (RETURN_LOGS && L) out.logs = L.dump();
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* =========================
   === SHEET HELPERS     ===
   ========================= */

function getOrCreateSheet_(name, makeHidden, L) {
  let sh = SS.getSheetByName(name);
  if (!sh) { sh = SS.insertSheet(name); L && L.log('[getOrCreateSheet_]', { created: name }); }
  if (makeHidden && !sh.isSheetHidden()) { sh.hideSheet(); L && L.log('[getOrCreateSheet_]', { hid: name }); }
  return sh;
}

function clearSheet_(sheet, L) { sheet.clearContents(); L && L.log('[clearSheet_]', { sheet: sheet.getName() }); }

function ensureCapacity_(sheet, neededRows, neededCols, L) {
  const curRows = sheet.getMaxRows(), curCols = sheet.getMaxColumns();
  if (curRows < neededRows) sheet.insertRowsAfter(curRows, neededRows - curRows);
  if (curCols < neededCols) sheet.insertColumnsAfter(curCols, neededCols - curCols);
}

function resizeGridExactly_(sheet, rows, cols, L) {
  ensureCapacity_(sheet, rows, cols, L);
  const mr = sheet.getMaxRows(), mc = sheet.getMaxColumns();
  if (mr > rows) sheet.deleteRows(rows + 1, mr - rows);
  if (mc > cols) sheet.deleteColumns(cols + 1, mc - cols);
}

function usedSize_(sheet) {
  return { rows: Math.max(1, sheet.getLastRow()), cols: Math.max(1, sheet.getLastColumn()) };
}

function swapStagingToLive_(stagingName, liveName, L) {
  const staging = getOrCreateSheet_(stagingName, true, L);
  const live = getOrCreateSheet_(liveName, false, L);
  const used = usedSize_(staging);
  const values = (used.rows > 0 && used.cols > 0)
    ? staging.getRange(1, 1, used.rows, used.cols).getValues() : [[]];
  clearSheet_(live, L);
  resizeGridExactly_(live, used.rows, used.cols, L);
  if (used.rows > 0 && used.cols > 0) live.getRange(1, 1, used.rows, used.cols).setValues(values);
  SpreadsheetApp.flush();
  L.log('[swap] committed', { liveName, rows: used.rows, cols: used.cols });
  return { ok: true, rows: used.rows, cols: used.cols };
}

function writeAllToStaging_(stagingName, values, L) {
  const staging = getOrCreateSheet_(stagingName, true, L);
  clearSheet_(staging, L);
  if (!values || !values.length) { SpreadsheetApp.flush(); return { rows: 0, cols: 0 }; }
  const rows = values.length, cols = values[0].length || 1;
  ensureCapacity_(staging, rows, cols, L);
  const BLOCK = 2000;
  let w = 0;
  while (w < rows) {
    const take = Math.min(BLOCK, rows - w);
    staging.getRange(1 + w, 1, take, cols).setValues(values.slice(w, w + take));
    w += take;
  }
  SpreadsheetApp.flush();
  L.log('[writeAllToStaging_]', { stagingName, rows, cols });
  return { rows, cols };
}

// Read a live tab into a Map keyed by the first column (string), value = full row.
function loadTabRows_(name) {
  const sh = SS.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return { header: [], rows: [] };
  const data = sh.getDataRange().getValues();
  return { header: data[0].map(h => String(h).trim().toLowerCase()), rows: data.slice(1) };
}

/* =========================
   === SDE / UNIVERSE MAPS ===
   ========================= */

// SDE_Types -> Map(type_id -> {name, group, category, volume})
function loadTypeMap_(L) {
  const { header, rows } = loadTabRows_('SDE_Types');
  const map = new Map();
  if (!rows.length) { L.log('[loadTypeMap_] SDE_Types empty'); return map; }
  const c = { id: header.indexOf('type_id'), name: header.indexOf('type_name'),
              grp: header.indexOf('group_name'), cat: header.indexOf('category_name'), vol: header.indexOf('volume') };
  rows.forEach(r => {
    const id = Number(r[c.id]); if (!id) return;
    map.set(id, { name: r[c.name], group: r[c.grp], category: r[c.cat], volume: Number(r[c.vol]) || 0 });
  });
  L.log('[loadTypeMap_]', { types: map.size });
  return map;
}

// SDE_Systems -> Map(system_id -> {name, regionId, security}); SDE_Regions -> Map(region_id -> name)
function loadUniverse_(L) {
  const sys = new Map(), reg = new Map(), sta = new Map();
  let t = loadTabRows_('SDE_Systems');
  if (t.rows.length) {
    const c = { id: t.header.indexOf('system_id'), name: t.header.indexOf('system_name'),
                rid: t.header.indexOf('region_id'), sec: t.header.indexOf('security') };
    t.rows.forEach(r => { const id = Number(r[c.id]); if (id) sys.set(id, { name: r[c.name], regionId: Number(r[c.rid]) || 0, security: Number(r[c.sec]) }); });
  }
  t = loadTabRows_('SDE_Regions');
  if (t.rows.length) {
    const c = { id: t.header.indexOf('region_id'), name: t.header.indexOf('region_name') };
    t.rows.forEach(r => { const id = Number(r[c.id]); if (id) reg.set(id, r[c.name]); });
  }
  t = loadTabRows_('SDE_Stations');
  if (t.rows.length) {
    const c = { id: t.header.indexOf('station_id'), sid: t.header.indexOf('system_id') };
    t.rows.forEach(r => { const id = Number(r[c.id]); if (id) sta.set(id, Number(r[c.sid]) || 0); });
  }
  L.log('[loadUniverse_]', { systems: sys.size, regions: reg.size, stations: sta.size });
  return { sys, reg, sta };
}

/* =========================
   === LOCATION-NAME CACHE ===
   Persisted in SDE_LocNames so station/structure names are fetched once.
   ========================= */

function loadLocNameCache_(L) {
  const sh = getOrCreateSheet_(LOCNAME_SHEET_NAME, true, L);
  const map = new Map();
  if (sh.getLastRow() >= 2) {
    const data = sh.getDataRange().getValues(); // loc_id, name, system_id, kind
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0]); if (!id) continue;
      map.set(id, { name: data[i][1], systemId: Number(data[i][2]) || 0, kind: data[i][3] });
    }
  }
  L.log('[loadLocNameCache_]', { cached: map.size });
  return map;
}

function saveLocNameCache_(map, L) {
  const sh = getOrCreateSheet_(LOCNAME_SHEET_NAME, true, L);
  const out = [['loc_id', 'name', 'system_id', 'kind']];
  map.forEach((v, k) => out.push([k, v.name, v.systemId, v.kind]));
  clearSheet_(sh, L);
  resizeGridExactly_(sh, out.length, 4, L);
  sh.getRange(1, 1, out.length, 4).setValues(out);
  SpreadsheetApp.flush();
}

// Fetch a public NPC-station name from ESI (no auth).
function fetchStationName_(stationId) {
  const url = 'https://esi.evetech.net/latest/universe/stations/' + stationId + '/?datasource=tranquility';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const o = JSON.parse(res.getContentText());
  return { name: o.name, systemId: o.system_id || 0, kind: 'station' };
}

// Fetch a player-structure name via GESI (needs read_structures scope + docking access).
function fetchStructure_(structureId, charName, L) {
  try {
    const rows = GESI.invoke(EP_STRUCTURE, { structure_id: Number(structureId), name: charName, show_column_headings: true });
    if (!rows || rows.length < 2) return null;
    const h = rows[0].map(x => String(x).trim().toLowerCase());
    const nameI = h.indexOf('name'), sysI = h.indexOf('solar_system_id');
    return { name: rows[1][nameI], systemId: Number(rows[1][sysI]) || 0, kind: 'structure' };
  } catch (e) { L.log('[fetchStructure_] failed', { structureId, error: String(e) }); return null; }
}

/* =========================
   === MARKET PRICES     ===
   ========================= */

// Fuzzwork aggregates -> Map(type_id -> { buy, sell }) using the 5%/95% percentile.
function loadJitaPrices_(typeIds, L) {
  const map = new Map();
  const ids = Array.from(new Set(typeIds.filter(Boolean)));
  for (let i = 0; i < ids.length; i += PRICE_CHUNK) {
    const chunk = ids.slice(i, i + PRICE_CHUNK);
    const url = FUZZWORK_URL + '?region=' + FORGE_REGION + '&types=' + chunk.join(',');
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) { L.log('[loadJitaPrices_] http', { code: res.getResponseCode(), at: i }); continue; }
      const o = JSON.parse(res.getContentText());
      chunk.forEach(id => {
        const e = o[id]; if (!e) return;
        const buy = e.buy ? (Number(e.buy.percentile) || Number(e.buy.max) || 0) : 0;
        const sell = e.sell ? (Number(e.sell.percentile) || Number(e.sell.min) || 0) : 0;
        map.set(Number(id), { buy, sell });
      });
    } catch (err) { L.log('[loadJitaPrices_] error', { at: i, error: String(err) }); }
  }
  L.log('[loadJitaPrices_]', { priced: map.size, requested: ids.length });
  return map;
}

/* =========================
   === ASSETS REFRESH    ===
   ========================= */

const ASSET_HEADER = [
  'owner', 'item_id', 'type_id', 'type_name', 'group_name', 'category_name',
  'quantity', 'singleton', 'location_flag', 'location_id', 'location_name',
  'system', 'region', 'security', 'volume', 'total_volume',
  'jita_buy', 'jita_sell', 'value',
];

function idx_(header, name) { return header.indexOf(name); }

function refreshAssets_(L) {
  const names = GESI.getAuthenticatedCharacterNames();
  if (!names || !names.length) throw new Error('No authorized characters in GESI.');
  L.log('[refreshAssets_] characters', { count: names.length });

  const raw = GESI.invokeMultiple(EP_ASSETS, names);
  if (!raw || raw.length < 2) throw new Error('GESI returned no asset rows.');
  const H = raw[0].map(h => String(h).trim().toLowerCase());
  const ci = {
    owner: (idx_(H, 'character_name') >= 0 ? idx_(H, 'character_name') : (idx_(H, 'character') >= 0 ? idx_(H, 'character') : idx_(H, 'name'))),
    item: idx_(H, 'item_id'), type: idx_(H, 'type_id'), qty: idx_(H, 'quantity'),
    locId: idx_(H, 'location_id'), locType: idx_(H, 'location_type'), locFlag: idx_(H, 'location_flag'),
    singleton: idx_(H, 'is_singleton'),
  };
  if (ci.type < 0 || ci.item < 0 || ci.locId < 0) throw new Error('assets header missing item_id/type_id/location_id: ' + H.join(','));
  const assets = raw.slice(1);
  L.log('[refreshAssets_] rows', { rows: assets.length, header: H });

  // Index every owned item so we can walk container/ship chains.
  const byItem = new Map();
  assets.forEach(r => byItem.set(String(r[ci.item]), r));

  // Enrichment sources.
  const types = loadTypeMap_(L);
  const uni = loadUniverse_(L);
  const locCache = loadLocNameCache_(L);
  const prices = loadJitaPrices_(assets.map(r => Number(r[ci.type])), L);

  // Resolve the ROOT location_id of an asset by following item->parent links.
  function resolveRoot(r) {
    let cur = r, depth = 0;
    while (depth++ < 64) {
      const lt = ci.locType >= 0 ? String(cur[ci.locType] || '') : '';
      const lid = String(cur[ci.locId]);
      const parent = byItem.get(lid);
      const isItem = lt === 'item' || (!!parent && lt !== 'station' && lt !== 'solar_system');
      if (isItem && parent) { cur = parent; continue; }
      return { id: Number(cur[ci.locId]), type: lt };
    }
    return { id: Number(r[ci.locId]), type: 'other' };
  }

  // Resolve a root location_id -> { name, systemId } using SDE + caches + ESI.
  let newCacheEntries = 0;
  function resolveLocation(rootId, rootType, ownerName) {
    // NPC station from SDE.
    if (uni.sta.has(rootId)) {
      const sysId = uni.sta.get(rootId);
      let nm = locCache.get(String(rootId));
      if (!nm) {
        const f = fetchStationName_(rootId);
        nm = f || { name: 'Station ' + rootId, systemId: sysId, kind: 'station' };
        locCache.set(String(rootId), nm); newCacheEntries++;
      }
      return { name: nm.name, systemId: nm.systemId || sysId };
    }
    // Solar system directly (item floating in space).
    if (rootType === 'solar_system' || uni.sys.has(rootId)) {
      const s = uni.sys.get(rootId);
      return { name: s ? s.name : ('System ' + rootId), systemId: rootId };
    }
    // Player structure (or unknown large id) -> ESI via GESI, cached.
    let nm = locCache.get(String(rootId));
    if (!nm) {
      const f = fetchStructure_(rootId, ownerName, L);
      nm = f || { name: 'Structure ' + rootId, systemId: 0, kind: 'structure' };
      locCache.set(String(rootId), nm); newCacheEntries++;
    }
    return { name: nm.name, systemId: nm.systemId };
  }

  const out = [ASSET_HEADER.slice()];
  assets.forEach(r => {
    const typeId = Number(r[ci.type]);
    const qty = ci.qty >= 0 ? (Number(r[ci.qty]) || 0) : 0;
    const owner = ci.owner >= 0 ? r[ci.owner] : '';
    const meta = types.get(typeId) || { name: '', group: '', category: '', volume: 0 };

    const root = resolveRoot(r);
    const loc = resolveLocation(root.id, root.type, owner);
    const s = loc.systemId ? uni.sys.get(loc.systemId) : null;
    const regionName = s ? (uni.reg.get(s.regionId) || '') : '';

    const p = prices.get(typeId) || { buy: 0, sell: 0 };
    out.push([
      owner, Number(r[ci.item]), typeId, meta.name, meta.group, meta.category,
      qty, ci.singleton >= 0 ? r[ci.singleton] : '', ci.locFlag >= 0 ? r[ci.locFlag] : '',
      root.id, loc.name, s ? s.name : '', regionName, s ? s.security : '',
      meta.volume, meta.volume * qty, p.buy, p.sell, p.sell * qty,
    ]);
  });

  if (newCacheEntries > 0) { saveLocNameCache_(locCache, L); L.log('[refreshAssets_] locname cache grew', { added: newCacheEntries }); }

  writeAllToStaging_(ASSET_STAGING_NAME, out, L);
  const swapped = swapStagingToLive_(ASSET_STAGING_NAME, ASSET_SHEET_NAME, L);

  try { const stamp = SS.getRangeByName('allinventoryimport'); if (stamp) stamp.setValue(new Date()); } catch (e) {}

  return { ok: true, characters: names.length, rows: out.length - 1, priced: prices.size, newLocations: newCacheEntries, swapped };
}

/* =========================
   === SDE UPLOAD (PC)   ===
   ========================= */

function sdeTarget_(body) {
  const key = String((body && body.target) || 'types');
  const t = SDE_TARGETS[key];
  if (!t) throw new Error('Unknown SDE target: ' + key);
  return t;
}

function importSdeRows_(L, body) {
  const t = sdeTarget_(body);
  const staging = getOrCreateSheet_(t.staging, true, L);
  const startIndex = Math.max(0, Number(body.startIndex || 0));
  const values = body && body.values;
  if (!values || !values.length || !Array.isArray(values[0])) throw new Error('Missing/invalid "values" (2D array).');
  const neededRows = startIndex + values.length, neededCols = values[0].length || 1;
  ensureCapacity_(staging, neededRows, neededCols, L);
  staging.getRange(startIndex + 1, 1, values.length, neededCols).setValues(values);
  SpreadsheetApp.flush();
  return { ok: true, target: body.target || 'types', wrote: values.length, nextIndex: startIndex + values.length };
}

/* =========================
   === HTTP HANDLERS     ===
   ========================= */

function doGet(e) {
  const L = newLogger();
  const fn = (e && e.parameter && (e.parameter.fn || e.parameter.op)) || 'ping';
  return json_(true, fn, { pong: true, build: BUILD_TAG, now: new Date().toISOString(), method: 'GET' }, L);
}

/**
 * Run ONCE from the editor (select authorizeOnce -> Run) to grant the OAuth
 * scopes the web app needs, so the deployed /exec works for anonymous callers.
 */
function authorizeOnce() {
  const hasToken = !!PropertiesService.getScriptProperties().getProperty('GESI_TOKEN');
  SpreadsheetApp.getActiveSpreadsheet().getName();
  UrlFetchApp.fetch('https://esi.evetech.net/latest/status/');
  let chars = [];
  try { chars = GESI.getAuthenticatedCharacterNames(); } catch (e) {}
  console.log('authorizeOnce ok', { tokenSet: hasToken, gesiCharacters: chars.length });
  return { ok: true, tokenSet: hasToken, gesiCharacters: chars.length };
}

function doPost(e) {
  const L = newLogger();
  const started = Date.now();
  let body = {};
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    body = JSON.parse(raw);
  } catch (err) { return json_(false, 'parseError', { error: 'Invalid JSON body' }, L, started); }

  if (!body || body.token !== getToken_()) return json_(false, 'forbidden', { error: 'Forbidden' }, L, started);

  const fn = body.fn || body.op || 'ping';
  try {
    switch (fn) {
      case 'ping':
        return json_(true, fn, { pong: true, build: BUILD_TAG, now: new Date().toISOString(), method: 'POST' }, L, started);
      case 'refreshAssets':
        return json_(true, fn, refreshAssets_(L), L, started);
      case 'beginSde': {
        const t = sdeTarget_(body);
        clearSheet_(getOrCreateSheet_(t.staging, true, L), L); SpreadsheetApp.flush();
        return json_(true, fn, { ok: true, target: body.target || 'types' }, L, started);
      }
      case 'importSdeRows':
        return json_(true, fn, importSdeRows_(L, body), L, started);
      case 'commitSde': {
        const t = sdeTarget_(body);
        return json_(true, fn, swapStagingToLive_(t.staging, t.live, L), L, started);
      }
      default:
        return json_(false, fn, { error: 'Unknown function' }, L, started);
    }
  } catch (err) {
    return json_(false, fn, { error: String(err && err.stack || err) }, L, started);
  }
}
