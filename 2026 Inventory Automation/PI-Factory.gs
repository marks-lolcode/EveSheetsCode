// ==========================================================
// FILE: PI-Factory.gs
// WHAT: PI manufacturing profit calculator + buy list + cost-basis sell floor.
//        Buys low-tier PI on the market, runs factory planets to convert it to
//        the next tier (P2->P3, P3->P4), sells the output. Ranks schematics by
//        profit/hr after all costs (inputs + POCO customs + Jita sales/broker).
//
// SHARES the Assets-GESI.gs project (one Apps Script project = one global scope),
// so it reuses, with no import:
//   loadTabRows_, loadTypeMap_, loadJitaPrices_, getOrCreateSheet_, clearSheet_,
//   resizeGridExactly_, ensureCapacity_, writeAllToStaging_, swapStagingToLive_,
//   newLogger, SS, ASSET_SHEET_NAME.
//
// DATA SOURCES:
//   SDE_Schematics  -> recipes (uploaded PC-side by Refresh-Inventory.ps1 -Sde).
//   SDE_Types       -> names + group_name (tier).
//   Fuzzwork        -> live Jita buy/sell (loadJitaPrices_).
//   ESI markets/prices -> CCP base_value for POCO customs (approximation).
//   CONFIG          -> global taxes/fees + PI knobs (shared with mfg/trading).
//
// PHASES: 1 calculator (PIProfit), 2 buy list (PIBuyList), 3 colonies (future),
//         4 cost-basis sell floor (PICostBasis). Phase 3 is a stub below.
// ==========================================================

/* =========================
   === CONSTANTS         ===
   ========================= */

const SCHEMATICS_TAB  = 'SDE_Schematics';
const BASEPRICES_TAB  = 'SDE_BasePrices';
const CONFIG_TAB      = 'CONFIG';
const PI_FILTER_TAB   = 'PILocationFilter';
const PIPROFIT_TAB    = 'PIProfit';
const PIPLAN_TAB      = 'PIFactoryPlan';
const PIBUYLIST_TAB   = 'PIBuyList';
const PICOSTBASIS_TAB = 'PICostBasis';
const PIINVENTORY_TAB    = 'PIInventory';
const PIRUNSIZES_TAB     = 'PIRunSizes';
const PIILLIQUID_TAB     = 'PIIlliquid';
const PIPLANETPROD_TAB   = 'PIPlanetProduction';
const PIREPLACE_TAB      = 'PIReplace';
const PIBUYLISTCHAR_TAB  = 'PIBuyListByChar';
const PIBUYLISTITEM_TAB  = 'PIBuyListByItem';
const PIINPUTSCHAR_TAB   = 'PIInputsByChar';
const PIINPUTSMAT_TAB    = 'PIInputsByMaterial';

// CONFIG character section (below the config block, after a blank spacer gap).
const CHAR_SECTION_MARKER = '>> CHARACTERS (mark include for PI steps)';
const CHAR_SUBHEADER      = ['character', 'include'];
const CONFIG_SPACER_ROWS  = 15; // blank rows left for future config keys

// Accounting / Broker Relations skill typeIDs (for ESI tax/fee derivation).
const SKILL_ACCOUNTING       = 16622;
const SKILL_BROKER_RELATIONS = 3446;

const CONFIG_HEADER = ['key', 'manual_value', 'esi_value', 'use_esi', 'source', 'last_updated', 'note'];
// Seeded once if CONFIG is empty. use_esi=FALSE => manual_value wins.
const CONFIG_DEFAULTS = [
  ['sales_tax_rate',    0.045,    '', false, 'manual', '', 'Jita sales tax. ESI-derivable from Accounting (16622): 4.5% x (1 - 0.11 x lvl)'],
  ['broker_fee_rate',   0.03,     '', false, 'manual', '', 'Jita broker fee. ESI-approx from Broker Relations (3446): 3% - 0.3% x lvl'],
  ['price_region_id',   10000002, '', false, 'manual', '', 'The Forge (Jita)'],
  ['selling_character', 'Sidewaze', '', false, 'manual', '', 'Buy/sell char (id 2118597159). Drives wallet pull + ESI tax/fee. Must match the GESI character name.'],
  ['poco_tax_rate',     0.05,     '', false, 'manual', '', 'Single POCO tax, applied to BOTH import (inputs) and export (output)'],
  ['input_price_side',  'buy',    '', false, 'manual', '', 'buy|sell - price basis for inputs'],
  ['output_price_side', 'sell',   '', false, 'manual', '', 'buy|sell - price basis for outputs'],
  ['min_margin_pct',    0.05,     '', false, 'manual', '', 'Min margin fraction (0.05 = 5%) for Phase 4 sell-floor target (PIProfit no longer filters). Pure buy->sell PI runs ~2-16%; raise if self-extracting P0.'],
  ['launchpad_m3',      10000,    '', false, 'manual', '', 'Launchpad storage capacity (m3). Drives run sizing in PIRunSizes.'],
  ['p4_launchpads_per_planet', 2, '', false, 'manual', '', 'P4 launchpads (= runs) per planet. P3 uses one launchpad per input.'],
  ['broker_faction_id', 500001,   '', false, 'manual', '', 'Sell-hub owner FACTION for broker-fee standings refine (500001 = Caldari State / Jita).'],
  ['broker_corp_id',    1000035,  '', false, 'manual', '', 'Sell-hub owner NPC CORP for broker-fee standings refine (1000035 = Caldari Navy / Jita 4-4).'],
  ['buylist_target_tier_only', 'TRUE', '', false, 'manual', '', 'PIBuyListByItem: TRUE = only count planets whose schematic tier is in buylist_target_tiers. FALSE = all producing planets.'],
  ['buylist_target_tiers', 'P3,P4', '', false, 'manual', '', 'Comma-separated tiers kept by the by-item buy list when buylist_target_tier_only=TRUE.'],
];

const PI_FILTER_HEADER  = ['include_systems', 'include_structures', 'include_containers', 'include_container_types', 'include_flags']; // cols A..E
const PI_FILTER_HELPERS = ['found_systems', 'found_structures', 'found_containers', 'found_container_types', 'found_flags'];          // cols G..K
const PI_FILTER_HELPER_COL = 7; // G (col F left as a gap between include lists and helpers)
const PIPLAN_HEADER     = ['product', 'slots', 'daily_target', 'char_planet_label'];
const PICOSTBASIS_HEADER = ['batch_id', 'date_started', 'product', 'output_qty', 'input_cost_total',
  'customs_cost', 'all_in_cost', 'break_even_unit', 'target_floor_unit', 'cost_source',
  'market_input_cost', 'price_variance', 'status', 'notes'];

/* =========================
   === SMALL HELPERS     ===
   ========================= */

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function round4(n) { return Math.round((Number(n) || 0) * 10000) / 10000; }
function norm_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }

function pickPrice_(entry, side) { entry = entry || { buy: 0, sell: 0 }; return side === 'buy' ? (entry.buy || 0) : (entry.sell || 0); }

// Replace any existing basic filter on a tab with one covering its used range.
// (swapStagingToLive_ rewrites the live sheet each run, dropping the old filter.)
function setSheetFilter_(tabName, L) {
  const sh = SS.getSheetByName(tabName);
  if (!sh) return;
  const existing = sh.getFilter(); if (existing) existing.remove();
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr > 1 && lc > 0) sh.getRange(1, 1, lr, lc).createFilter();
  L && L.log('[setSheetFilter_]', { tab: tabName, rows: lr, cols: lc });
}
function toBool_(v) { return v === true || String(v == null ? '' : v).trim().toUpperCase() === 'TRUE'; }

// AllInventory column indices (incl owner + location_flag for filtering/grouping).
function invCols_(header) {
  return { type: header.indexOf('type_id'), qty: header.indexOf('quantity'),
           system: header.indexOf('system'), location: header.indexOf('location_name'),
           container: header.indexOf('container_name'), containerType: header.indexOf('container_type'),
           flag: header.indexOf('location_flag'), owner: header.indexOf('owner') };
}

// Price side for a type given its role. Illiquid items are conservative both ways:
// as an INPUT we buy off sell orders ('sell'); as our OUTPUT we dump into buy
// orders ('buy'). Normal items use the configured side per role.
function effectivePriceSide_(typeId, role, illiquid, cfg) {
  if (illiquid && illiquid.has(Number(typeId))) return role === 'input' ? 'sell' : 'buy';
  return role === 'input' ? cfgStr_(cfg, 'input_price_side', 'buy') : cfgStr_(cfg, 'output_price_side', 'sell');
}

// Output type's group_name embeds the tier verbatim (Basic..Tier 1 .. Advanced..Tier 4).
function tierOf_(groupName) {
  const g = String(groupName || '');
  if (g.indexOf('Tier 4') >= 0) return 'P4';
  if (g.indexOf('Tier 3') >= 0) return 'P3';
  if (g.indexOf('Tier 2') >= 0) return 'P2';
  if (g.indexOf('Tier 1') >= 0) return 'P1';
  return null;
}

/* =========================
   === CONFIG (global)   ===
   ========================= */

function ensureConfig_(L) {
  const sh = getOrCreateSheet_(CONFIG_TAB, false, L);
  if (sh.getLastRow() >= 2) return sh;
  const out = [CONFIG_HEADER.slice()].concat(CONFIG_DEFAULTS.map(r => r.slice()));
  resizeGridExactly_(sh, out.length, CONFIG_HEADER.length, L);
  sh.getRange(1, 1, out.length, CONFIG_HEADER.length).setValues(out);
  SpreadsheetApp.flush();
  L && L.log('[ensureConfig_] seeded', { rows: CONFIG_DEFAULTS.length });
  return sh;
}

// CONFIG -> Map(key -> effective value). use_esi ? esi_value : manual_value.
function loadConfig_() {
  const { header, rows } = loadTabRows_(CONFIG_TAB);
  const map = new Map();
  if (!rows.length) return map;
  const c = { key: header.indexOf('key'), man: header.indexOf('manual_value'),
              esi: header.indexOf('esi_value'), use: header.indexOf('use_esi') };
  for (let i = 0; i < rows.length; i++) {
    const key = String(rows[i][c.key] || '').trim();
    if (key === CHAR_SECTION_MARKER) break; // character list lives below the config block
    if (!key) continue;
    const useEsi = toBool_(rows[i][c.use]);
    map.set(key, useEsi ? rows[i][c.esi] : rows[i][c.man]);
  }
  return map;
}

/* =========================
   === CHARACTER LIST (CONFIG) ===
   A vertical, manually-curated include-list of authenticated characters, kept on
   the CONFIG sheet below a blank spacer gap. Gates every per-character PI step.
   ========================= */

// Row (1-based) of the CHARACTERS marker in CONFIG col A, or -1 if not present.
function findCharMarkerRow_(sh) {
  const last = sh.getLastRow();
  if (last < 1) return -1;
  const colA = sh.getRange(1, 1, last, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim() === CHAR_SECTION_MARKER) return i + 1;
  }
  return -1;
}

// Ensure the CHARACTERS marker + subheader exist (after ~15 blank spacer rows).
// Returns { sh, markerRow }.
function ensureCharacterSection_(L) {
  const sh = getOrCreateSheet_(CONFIG_TAB, false, L);
  ensureConfig_(L);
  let markerRow = findCharMarkerRow_(sh);
  if (markerRow < 0) {
    markerRow = sh.getLastRow() + 1 + CONFIG_SPACER_ROWS;
    ensureCapacity_(sh, markerRow + 1, Math.max(2, sh.getMaxColumns()), L);
    sh.getRange(markerRow, 1).setValue(CHAR_SECTION_MARKER);
    sh.getRange(markerRow + 1, 1, 1, CHAR_SUBHEADER.length).setValues([CHAR_SUBHEADER]);
    SpreadsheetApp.flush();
    L && L.log('[ensureCharacterSection_] created', { markerRow });
  }
  return { sh, markerRow };
}

// Menu: Refresh Character List. Pull authed GESI characters and merge into the
// CONFIG CHARACTERS block: preserve existing include marks, append new chars
// (default include=TRUE), keep vanished chars but force include=FALSE. Idempotent.
function refreshCharacterList_(L) {
  const names = GESI.getAuthenticatedCharacterNames();
  if (!names || !names.length) throw new Error('No authorized characters in GESI.');
  const { sh, markerRow } = ensureCharacterSection_(L);
  const firstDataRow = markerRow + 2;
  const last = sh.getLastRow();

  const existing = new Map(); // name -> include bool
  if (last >= firstDataRow) {
    const vals = sh.getRange(firstDataRow, 1, last - firstDataRow + 1, 2).getValues();
    vals.forEach(r => { const n = String(r[0] || '').trim(); if (n) existing.set(n, toBool_(r[1])); });
  }

  const seen = new Set();
  const rowsOut = [];
  names.slice().sort((a, b) => String(a).localeCompare(String(b))).forEach(n => {
    seen.add(n);
    rowsOut.push([n, existing.has(n) ? existing.get(n) : true]);
  });
  existing.forEach((inc, n) => { if (!seen.has(n)) rowsOut.push([n, false]); }); // vanished

  if (last >= firstDataRow) sh.getRange(firstDataRow, 1, last - firstDataRow + 1, 2).clearContent();
  ensureCapacity_(sh, firstDataRow + rowsOut.length - 1, 2, L);
  if (rowsOut.length) {
    sh.getRange(firstDataRow, 1, rowsOut.length, 2).setValues(rowsOut);
    sh.getRange(firstDataRow, 2, rowsOut.length, 1).insertCheckboxes();
  }
  SpreadsheetApp.flush();
  L && L.log('[refreshCharacterList_]', { authenticated: names.length, listed: rowsOut.length });
  return { authenticated: names.length, listed: rowsOut.length, included: rowsOut.filter(r => r[1]).length };
}

// Names with include=TRUE in the CONFIG CHARACTERS block. Gates per-char steps.
function loadIncludedCharacters_() {
  const sh = SS.getSheetByName(CONFIG_TAB);
  if (!sh) return [];
  const markerRow = findCharMarkerRow_(sh);
  if (markerRow < 0) return [];
  const firstDataRow = markerRow + 2;
  const last = sh.getLastRow();
  if (last < firstDataRow) return [];
  const vals = sh.getRange(firstDataRow, 1, last - firstDataRow + 1, 2).getValues();
  const out = [];
  vals.forEach(r => { const n = String(r[0] || '').trim(); if (n && toBool_(r[1])) out.push(n); });
  return out;
}

// Generic single-key reader (shared by every sheet family). Returns Number when numeric.
function getConfig_(key, dflt) {
  const v = loadConfig_().get(key);
  if (v === undefined || v === null || v === '') return dflt;
  const n = Number(v);
  return isNaN(n) ? v : n;
}

function cfgNum_(cfg, key, dflt) {
  const v = cfg.get(key);
  const n = Number(v);
  return (v === '' || v === undefined || v === null || isNaN(n)) ? dflt : n;
}
function cfgStr_(cfg, key, dflt) {
  const v = cfg.get(key);
  return (v === undefined || v === null || v === '') ? (dflt || '') : String(v).trim();
}

// Menu: Refresh Config from ESI. Manual trigger only. Fills esi_value for
// sales_tax_rate / broker_fee_rate from the selling char's skills. The user
// still toggles use_esi=TRUE to adopt them.
function refreshConfigFromEsi_(L) {
  ensureConfig_(L);
  const cfg = loadConfig_();
  const charName = cfgStr_(cfg, 'selling_character', '');
  if (!charName) throw new Error('Set selling_character in CONFIG before Refresh Config from ESI.');

  const skills = fetchSkillLevels_(charName, L);
  const acct = skills.get(SKILL_ACCOUNTING) || 0;
  const brk  = skills.get(SKILL_BROKER_RELATIONS) || 0;
  const salesTax  = 0.045 * (1 - 0.11 * acct);

  // Broker fee = 3% - 0.3%*BrokerRelations - 0.03%*factionStanding - 0.02%*corpStanding
  // (standings toward the sell-hub owner). Approximation; structure/hub owners vary.
  const factionId = cfgNum_(cfg, 'broker_faction_id', 0);
  const corpId    = cfgNum_(cfg, 'broker_corp_id', 0);
  const st = fetchStandings_(charName, L);
  const facStand = factionId ? (st.get(factionId) || 0) : 0;
  const corpStand = corpId ? (st.get(corpId) || 0) : 0;
  const brokerFee = Math.max(0, 0.03 - 0.003 * brk - 0.0003 * facStand - 0.0002 * corpStand);

  writeConfigEsi_({ sales_tax_rate: salesTax, broker_fee_rate: brokerFee }, 'ESI:skills+standings', L);
  L && L.log('[refreshConfigFromEsi_]', { charName, acct, brk, facStand, corpStand, salesTax, brokerFee });
  return { selling_character: charName, accounting: acct, broker_relations: brk,
           faction_standing: round2(facStand), corp_standing: round2(corpStand),
           sales_tax_rate: round4(salesTax), broker_fee_rate: round4(brokerFee) };
}

// characters_character_standings -> Map(from_id -> standing). Covers faction + NPC corp.
function fetchStandings_(charName, L) {
  const map = new Map();
  try {
    const rows = GESI.invoke('characters_character_standings', { name: charName, show_column_headings: true });
    if (rows && rows.length > 1) {
      const h = rows[0].map(x => String(x).trim().toLowerCase());
      const idI = h.indexOf('from_id'), stI = h.indexOf('standing');
      if (idI >= 0 && stI >= 0) {
        for (let i = 1; i < rows.length; i++) map.set(Number(rows[i][idI]), Number(rows[i][stI]) || 0);
      }
    }
  } catch (e) { L && L.log('[fetchStandings_] failed', { error: String(e) }); }
  return map;
}

function fetchSkillLevels_(charName, L) {
  const map = new Map();
  try {
    const rows = GESI.invoke('characters_character_skills', { name: charName, show_column_headings: true });
    if (rows && rows.length > 1) {
      const h = rows[0].map(x => String(x).trim().toLowerCase());
      const sidI = h.indexOf('skill_id');
      let lvlI = h.indexOf('trained_skill_level');
      if (lvlI < 0) lvlI = h.indexOf('active_skill_level');
      if (lvlI < 0) lvlI = h.indexOf('skill_level');
      if (sidI >= 0 && lvlI >= 0) {
        for (let i = 1; i < rows.length; i++) map.set(Number(rows[i][sidI]), Number(rows[i][lvlI]) || 0);
      }
    }
  } catch (e) { L && L.log('[fetchSkillLevels_] failed', { error: String(e) }); }
  return map;
}

// Write esi_value + source + last_updated for the given keys (does NOT toggle use_esi).
function writeConfigEsi_(updates, source, L) {
  const sh = ensureConfig_(L);
  const data = sh.getDataRange().getValues();
  const h = data[0].map(x => String(x).trim().toLowerCase());
  const c = { key: h.indexOf('key'), esi: h.indexOf('esi_value'), src: h.indexOf('source'), upd: h.indexOf('last_updated') };
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][c.key] || '').trim();
    if (Object.prototype.hasOwnProperty.call(updates, key)) {
      if (c.esi >= 0) sh.getRange(i + 1, c.esi + 1).setValue(updates[key]);
      if (c.src >= 0) sh.getRange(i + 1, c.src + 1).setValue(source);
      if (c.upd >= 0) sh.getRange(i + 1, c.upd + 1).setValue(now);
    }
  }
  SpreadsheetApp.flush();
}

/* =========================
   === SCHEMATICS        ===
   ========================= */

// SDE_Schematics (long) -> [{ id, name, cycle, output:{type,qty,name,group}, inputs:[{type,qty,name}], tier }]
// Keeps only schematics whose OUTPUT tier is P3 or P4.
function loadSchematics_(L) {
  const { header, rows } = loadTabRows_(SCHEMATICS_TAB);
  if (!rows.length) { L && L.log('[loadSchematics_] SDE_Schematics empty'); return []; }
  const c = { id: header.indexOf('schematic_id'), name: header.indexOf('schematic_name'),
              cycle: header.indexOf('cycle_time_sec'), type: header.indexOf('type_id'),
              qty: header.indexOf('quantity'), inp: header.indexOf('is_input') };
  const types = loadTypeMap_(L);
  const map = new Map();
  rows.forEach(r => {
    const sid = Number(r[c.id]); if (!sid) return;
    let s = map.get(sid);
    if (!s) { s = { id: sid, name: String(r[c.name] || ''), cycle: Number(r[c.cycle]) || 0, output: null, inputs: [], tier: null }; map.set(sid, s); }
    const tid = Number(r[c.type]); const q = Number(r[c.qty]) || 0;
    const isInput = r[c.inp] === 1 || r[c.inp] === true || String(r[c.inp]) === '1';
    const meta = types.get(tid) || { name: '', group: '', volume: 0 };
    if (isInput) s.inputs.push({ type: tid, qty: q, name: meta.name, vol: meta.volume || 0 });
    else s.output = { type: tid, qty: q, name: meta.name, group: meta.group, vol: meta.volume || 0 };
  });
  const out = [];
  map.forEach(s => {
    if (!s.output) return;
    s.tier = tierOf_(s.output.group);
    if (s.tier === 'P3' || s.tier === 'P4') out.push(s);
  });
  L && L.log('[loadSchematics_]', { total: map.size, kept: out.length });
  return out;
}

/* =========================
   === ILLIQUID ITEMS    ===
   Thin-market PI. Priced conservatively in every calc: as an INPUT we acquire off
   sell orders ('sell'); as our OUTPUT we dump into buy orders ('buy'). See
   effectivePriceSide_.
   ========================= */

const PIILLIQUID_HEADER = ['illiquid_item'];

function ensureIlliquidTab_(L) {
  const sh = getOrCreateSheet_(PIILLIQUID_TAB, false, L);
  if (sh.getLastRow() >= 1 && String(sh.getRange(1, 1).getValue()).trim()) return sh;
  sh.getRange(1, 1, 1, PIILLIQUID_HEADER.length).setValues([PIILLIQUID_HEADER]);
  SpreadsheetApp.flush();
  return sh;
}

const PIILLIQUID_ROWS = 500;

// All Planetary Commodities (P0..P4) names from SDE_Types, sorted. Dropdown source.
function loadPiCommodityNames_(L) {
  const types = loadTypeMap_(L);
  const names = [];
  types.forEach(m => { if (String(m.category || '') === 'Planetary Commodities' && m.name) names.push(m.name); });
  return names.sort();
}

// Put a Planetary-Commodities dropdown on PIIlliquid col A. Called from List PI
// Locations (Setup) so it refreshes alongside the location-filter dropdowns.
function applyIlliquidDropdown_(L) {
  const sh = ensureIlliquidTab_(L);
  ensureCapacity_(sh, PIILLIQUID_ROWS + 1, 1, L);
  const names = loadPiCommodityNames_(L);
  const range = sh.getRange(2, 1, PIILLIQUID_ROWS, 1);
  if (names.length) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(true).build();
    range.setDataValidation(rule);
  } else {
    range.clearDataValidations();
  }
  SpreadsheetApp.flush();
  L && L.log('[applyIlliquidDropdown_]', { commodities: names.length });
  return names.length;
}

// PIIlliquid item names -> Set(type_id), resolved via SDE_Types.
function loadIlliquidSet_(L) {
  ensureIlliquidTab_(L);
  const set = new Set();
  const { rows } = loadTabRows_(PIILLIQUID_TAB);
  if (!rows.length) return set;
  const types = loadTypeMap_(L);
  const nameToId = new Map();
  types.forEach((m, id) => nameToId.set(norm_(m.name), id));
  rows.forEach(r => { const id = nameToId.get(norm_(r[0])); if (id) set.add(id); });
  L && L.log('[loadIlliquidSet_]', { listed: rows.length, resolved: set.size });
  return set;
}

/* =========================
   === PLANET NAMES / NUMBERS ===
   In-game planet number (the Roman numeral, e.g. "Tanoo III" -> 3) for sorting.
   ESI /universe/names/ (no auth), cached in SDE_PlanetNames.
   ========================= */

const PLANETNAMES_TAB = 'SDE_PlanetNames';

function loadPlanetNameCache_(L) {
  const sh = getOrCreateSheet_(PLANETNAMES_TAB, true, L);
  const map = new Map();
  if (sh.getLastRow() >= 2) {
    const data = sh.getDataRange().getValues(); // planet_id, name
    for (let i = 1; i < data.length; i++) { const id = String(data[i][0]); if (id) map.set(id, data[i][1]); }
  }
  return map;
}
function savePlanetNameCache_(map, L) {
  const sh = getOrCreateSheet_(PLANETNAMES_TAB, true, L);
  const out = [['planet_id', 'name']];
  map.forEach((v, k) => out.push([k, v]));
  clearSheet_(sh, L);
  resizeGridExactly_(sh, out.length, 2, L);
  sh.getRange(1, 1, out.length, 2).setValues(out);
  SpreadsheetApp.flush();
}

// Resolve planet_ids -> names via ESI /universe/names/ (batched <=1000), cached.
function resolvePlanetNames_(ids, L) {
  const want = Array.from(new Set((ids || []).map(Number).filter(Boolean)));
  const cache = loadPlanetNameCache_(L);
  const missing = want.filter(id => !cache.has(String(id)));
  let added = 0;
  const url = 'https://esi.evetech.net/latest/universe/names/?datasource=tranquility';
  for (let i = 0; i < missing.length; i += 1000) {
    const chunk = missing.slice(i, i + 1000);
    try {
      const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json',
        payload: JSON.stringify(chunk), muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        JSON.parse(res.getContentText()).forEach(o => { if (o && o.id) { cache.set(String(o.id), o.name || ''); added++; } });
      } else L && L.log('[resolvePlanetNames_] http', { code: res.getResponseCode() });
    } catch (e) { L && L.log('[resolvePlanetNames_] error', { error: String(e) }); }
  }
  if (added > 0) savePlanetNameCache_(cache, L);
  return cache;
}

function romanToInt_(s) {
  const map = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  s = String(s).toUpperCase();
  let tot = 0, prev = 0;
  for (let i = s.length - 1; i >= 0; i--) { const v = map[s[i]] || 0; if (v < prev) tot -= v; else { tot += v; prev = v; } }
  return tot;
}
// "Tanoo III" -> 3. 0 if no trailing Roman numeral.
function planetNumberFromName_(name) {
  const m = String(name || '').trim().match(/\b([IVXLCDM]+)\s*$/i);
  return m ? romanToInt_(m[1]) : 0;
}

/* =========================
   === BASE PRICES (POCO) ===
   ========================= */

// ESI /markets/prices/ average_price as the customs base_value. APPROXIMATION:
// CCP's published POCO base value is not exposed; average_price is the closest
// no-auth source. Persists the needed subset to SDE_BasePrices.
function fetchBasePrices_(neededIds, L) {
  const url = 'https://esi.evetech.net/latest/markets/prices/?datasource=tranquility';
  const map = new Map();
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const arr = JSON.parse(res.getContentText());
      arr.forEach(o => { const id = Number(o.type_id); if (id) map.set(id, Number(o.average_price) || Number(o.adjusted_price) || 0); });
    } else L && L.log('[fetchBasePrices_] http', { code: res.getResponseCode() });
  } catch (e) { L && L.log('[fetchBasePrices_] error', { error: String(e) }); }

  const need = neededIds ? new Set(Array.from(neededIds).map(Number)) : null;
  const out = [['type_id', 'base_value']];
  map.forEach((v, k) => { if (!need || need.has(k)) out.push([k, v]); });
  const sh = getOrCreateSheet_(BASEPRICES_TAB, true, L);
  clearSheet_(sh, L);
  resizeGridExactly_(sh, out.length, 2, L);
  sh.getRange(1, 1, out.length, 2).setValues(out);
  SpreadsheetApp.flush();
  L && L.log('[fetchBasePrices_]', { fetched: map.size, persisted: out.length - 1 });
  return map;
}

/* =========================
   === PHASE 1: PROFIT   ===
   ========================= */

// Menu: Build PI Profit. Lists ALL P3/P4 schematics, sorted by profit/hr desc
// after all costs (no margin filter — unprofitable rows are kept so they're visible).
// opts (optional, from the headless POST route): { sample } adds a top-5 diagnostic.
function buildPiProfit_(L, opts) {
  ensureConfig_(L);
  const cfg = loadConfig_();
  const poco      = cfgNum_(cfg, 'poco_tax_rate', 0);
  const salesTax  = cfgNum_(cfg, 'sales_tax_rate', 0);
  const broker    = cfgNum_(cfg, 'broker_fee_rate', 0);
  const ill       = loadIlliquidSet_(L);

  const schems = loadSchematics_(L);
  if (!schems.length) throw new Error('No P3/P4 schematics. Upload SDE_Schematics (Refresh-Inventory.ps1 -Sde) first.');

  const idSet = new Set();
  schems.forEach(s => { idSet.add(s.output.type); s.inputs.forEach(i => idSet.add(i.type)); });
  const prices = loadJitaPrices_(Array.from(idSet), L);
  const base   = fetchBasePrices_(idSet, L);

  const calc = schems.map(s => {
    const inputCost  = s.inputs.reduce((a, i) => a + i.qty * pickPrice_(prices.get(i.type), effectivePriceSide_(i.type, 'input', ill, cfg)), 0);
    const baseIn     = s.inputs.reduce((a, i) => a + i.qty * (base.get(i.type) || 0), 0);
    const baseOut    = s.output.qty * (base.get(s.output.type) || 0);
    const customsIn  = poco * baseIn;
    const customsOut = poco * baseOut;
    const revenue    = s.output.qty * pickPrice_(prices.get(s.output.type), effectivePriceSide_(s.output.type, 'output', ill, cfg));
    const salesBroker = revenue * (salesTax + broker);
    const net        = revenue - inputCost - customsIn - customsOut - salesBroker;
    const costBasis  = inputCost + customsIn + customsOut;
    const margin     = costBasis > 0 ? net / costBasis : 0;
    const profitHr   = s.cycle > 0 ? net / (s.cycle / 3600) : 0;
    return { s, inputCost, customsIn, customsOut, revenue, salesBroker, net, margin, profitHr };
  });

  const kept = calc.slice().sort((a, b) => b.profitHr - a.profitHr);

  const out = [['schematic_name', 'tier', 'output_name', 'output_qty', 'cycle_time_sec',
    'input_cost', 'customs_in', 'customs_out', 'sell_revenue', 'sales_broker_cost',
    'net_profit', 'profit_per_hr', 'margin_pct']];
  kept.forEach(r => out.push([
    r.s.name, r.s.tier, r.s.output.name, r.s.output.qty, r.s.cycle,
    round2(r.inputCost), round2(r.customsIn), round2(r.customsOut), round2(r.revenue), round2(r.salesBroker),
    round2(r.net), round2(r.profitHr), round4(r.margin),
  ]));
  writeAllToStaging_(PIPROFIT_TAB + '__staging', out, L);
  swapStagingToLive_(PIPROFIT_TAB + '__staging', PIPROFIT_TAB, L);
  L && L.log('[buildPiProfit_]', { schematics: schems.length, ranked: kept.length });
  const res = { schematics: schems.length, ranked: kept.length };
  if (opts && opts.sample) {
    res.top = calc.sort((a, b) => b.profitHr - a.profitHr).slice(0, 5).map(r => ({
      name: r.s.name, tier: r.s.tier, profit_hr: round2(r.profitHr), margin: round4(r.margin),
      revenue: round2(r.revenue), input_cost: round2(r.inputCost),
      customs: round2(r.customsIn + r.customsOut), fees: round2(r.salesBroker), net: round2(r.net),
    }));
  }
  return res;
}

/* =========================
   === LOCATION FILTER   ===
   ========================= */

function ensurePiFilterTab_(L) {
  const sh = getOrCreateSheet_(PI_FILTER_TAB, false, L);
  ensureCapacity_(sh, 2, PI_FILTER_HELPER_COL + PI_FILTER_HELPERS.length, L);
  // Always (re)write the label rows so an older sheet gains include_container_types.
  sh.getRange(1, 1, 1, PI_FILTER_HEADER.length).setValues([PI_FILTER_HEADER]);
  sh.getRange(1, PI_FILTER_HELPER_COL, 1, PI_FILTER_HELPERS.length).setValues([PI_FILTER_HELPERS]);
  SpreadsheetApp.flush();
  return sh;
}

// Five include-lists (display names, lowercased): system / structure / container name /
// container type / location_flag (ship + specialty holds). All empty = include all.
function loadLocationFilter_() {
  const sh = SS.getSheetByName(PI_FILTER_TAB);
  const f = { systems: [], structures: [], containers: [], containerTypes: [], flags: [] };
  if (!sh || sh.getLastRow() < 2) return f;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, PI_FILTER_HEADER.length).getValues();
  data.forEach(r => {
    const a = norm_(r[0]); if (a) f.systems.push(a);
    const b = norm_(r[1]); if (b) f.structures.push(b);
    const c = norm_(r[2]); if (c) f.containers.push(c);
    const d = norm_(r[3]); if (d) f.containerTypes.push(d);
    const e = norm_(r[4]); if (e) f.flags.push(e);
  });
  return f;
}

// OR-union: if every dimension is empty, include everything. Otherwise keep a row if
// its value matches ANY filled cell in ANY dimension (a union of all listed places).
function rowPassesFilter_(row, filter, cols) {
  const dims = [
    [filter.systems, cols.system], [filter.structures, cols.location],
    [filter.containers, cols.container], [filter.containerTypes, cols.containerType],
    [filter.flags, cols.flag],
  ];
  let any = false;
  for (let i = 0; i < dims.length; i++) {
    const list = dims[i][0], col = dims[i][1];
    if (!list.length) continue;
    any = true;
    if (col >= 0 && list.indexOf(norm_(row[col])) >= 0) return true;
  }
  return !any; // no constraints set => include all
}

// Menu: List PI Locations. Distinct names from AllInventory -> filter helper area (cols E..G).
function listPiLocations_(L) {
  const sh = ensurePiFilterTab_(L);
  const { header, rows } = loadTabRows_(ASSET_SHEET_NAME);
  if (!rows.length) throw new Error('AllInventory empty. Run Refresh Assets first.');
  const ic = invCols_(header);
  const sys = new Set(), loc = new Set(), con = new Set(), cty = new Set(), flg = new Set();
  rows.forEach(r => {
    if (ic.system >= 0 && r[ic.system]) sys.add(String(r[ic.system]));
    if (ic.location >= 0 && r[ic.location]) loc.add(String(r[ic.location]));
    if (ic.container >= 0 && r[ic.container]) con.add(String(r[ic.container]));
    if (ic.containerType >= 0 && r[ic.containerType]) cty.add(String(r[ic.containerType]));
    if (ic.flag >= 0 && r[ic.flag]) flg.add(String(r[ic.flag]));
  });
  const cols = [Array.from(sys).sort(), Array.from(loc).sort(), Array.from(con).sort(),
                Array.from(cty).sort(), Array.from(flg).sort()];
  const n = cols.length; // 5
  const maxLen = Math.max.apply(null, cols.map(a => a.length).concat([0]));

  if (sh.getMaxRows() > 1) sh.getRange(2, PI_FILTER_HELPER_COL, sh.getMaxRows() - 1, n).clearContent();
  if (maxLen > 0) {
    ensureCapacity_(sh, maxLen + 1, PI_FILTER_HELPER_COL + n, L);
    const block = [];
    for (let i = 0; i < maxLen; i++) block.push(cols.map(a => a[i] || ''));
    sh.getRange(2, PI_FILTER_HELPER_COL, block.length, n).setValues(block);
  }

  // Dropdowns on each include column (A..E) sourced from its found_* helper (G..K).
  const ROWS = 200;
  ensureCapacity_(sh, ROWS + 1, PI_FILTER_HELPER_COL + n, L);
  for (let c = 0; c < n; c++) {
    const srcLen = cols[c].length;
    const incRange = sh.getRange(2, c + 1, ROWS, 1);
    if (srcLen > 0) {
      const src = sh.getRange(2, PI_FILTER_HELPER_COL + c, srcLen, 1);
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(src, true).setAllowInvalid(true).build();
      incRange.setDataValidation(rule);
    } else {
      incRange.clearDataValidations();
    }
  }
  SpreadsheetApp.flush();
  const counts = { systems: cols[0].length, structures: cols[1].length, containers: cols[2].length, container_types: cols[3].length, flags: cols[4].length };
  counts.illiquid_options = applyIlliquidDropdown_(L); // refresh PIIlliquid dropdown too
  L && L.log('[listPiLocations_]', counts);
  return counts;
}

/* =========================
   === PHASE 2: BUY LIST ===
   ========================= */

function ensurePlanTab_(L) {
  const sh = getOrCreateSheet_(PIPLAN_TAB, false, L);
  if (sh.getLastRow() >= 1 && String(sh.getRange(1, 1).getValue()).trim()) return sh;
  sh.getRange(1, 1, 1, PIPLAN_HEADER.length).setValues([PIPLAN_HEADER]);
  SpreadsheetApp.flush();
  return sh;
}

// Menu: Build Buy List. Demand (per cycle x slots, or scaled to daily_target)
// minus filtered on-hand, priced at input_price_side.
function buildPiBuyList_(L) {
  ensurePlanTab_(L); ensurePiFilterTab_(L); ensureConfig_(L);
  const cfg = loadConfig_();
  const ill = loadIlliquidSet_(L);

  const plan = loadTabRows_(PIPLAN_TAB);
  if (!plan.rows.length) throw new Error('PIFactoryPlan empty. Add products (and slots or daily_target) first.');
  const ph = { prod: plan.header.indexOf('product'), slots: plan.header.indexOf('slots'), daily: plan.header.indexOf('daily_target') };

  const schems = loadSchematics_(L);
  const byName = new Map(); schems.forEach(s => byName.set(norm_(s.name), s));

  // Aggregate per-input demand across all planned products.
  const demand = new Map(); // type_id -> { qty, name }
  plan.rows.forEach(r => {
    const name = norm_(r[ph.prod]); if (!name) return;
    const s = byName.get(name);
    if (!s) { L && L.log('[buildPiBuyList_] unknown product', { name }); return; }
    const slots = ph.slots >= 0 ? Number(r[ph.slots]) || 0 : 0;
    const daily = ph.daily >= 0 ? Number(r[ph.daily]) || 0 : 0;
    let factor = 0;
    if (slots > 0) factor = slots;                                  // one cycle's inputs x slots
    else if (daily > 0 && s.output.qty > 0) factor = daily / s.output.qty; // cycles to hit daily output
    if (factor <= 0) return;
    s.inputs.forEach(i => {
      const cur = demand.get(i.type) || { qty: 0, name: i.name };
      cur.qty += i.qty * factor;
      demand.set(i.type, cur);
    });
  });
  if (!demand.size) throw new Error('No demand. Set slots or daily_target on PIFactoryPlan rows.');

  // On-hand from AllInventory, filtered by PILocationFilter.
  const inv = loadTabRows_(ASSET_SHEET_NAME);
  const ic = invCols_(inv.header);
  const filter = loadLocationFilter_();
  const onHand = new Map();
  inv.rows.forEach(r => {
    if (!rowPassesFilter_(r, filter, ic)) return;
    const tid = Number(r[ic.type]); if (!demand.has(tid)) return;
    onHand.set(tid, (onHand.get(tid) || 0) + (Number(r[ic.qty]) || 0));
  });

  const ids = Array.from(demand.keys());
  const prices = loadJitaPrices_(ids, L);
  ids.sort((a, b) => String(demand.get(a).name).localeCompare(String(demand.get(b).name)));

  const out = [['input_name', 'type_id', 'qty_needed', 'qty_on_hand', 'qty_to_buy', 'unit_price', 'line_cost']];
  ids.forEach(tid => {
    const d = demand.get(tid);
    const need = Math.ceil(d.qty);
    const have = onHand.get(tid) || 0;
    const buy = Math.max(0, need - have);
    const unit = pickPrice_(prices.get(tid), effectivePriceSide_(tid, 'input', ill, cfg));
    out.push([d.name, tid, need, have, buy, round2(unit), round2(buy * unit)]);
  });
  writeAllToStaging_(PIBUYLIST_TAB + '__staging', out, L);
  swapStagingToLive_(PIBUYLIST_TAB + '__staging', PIBUYLIST_TAB, L);
  L && L.log('[buildPiBuyList_]', { inputs: ids.length });
  return { inputs: ids.length };
}

/* =========================
   === PHASE 4: COST BASIS ===
   ========================= */

function ensureCostBasisTab_(L) {
  const sh = getOrCreateSheet_(PICOSTBASIS_TAB, false, L);
  if (sh.getLastRow() >= 1 && String(sh.getRange(1, 1).getValue()).trim()) return sh;
  sh.getRange(1, 1, 1, PICOSTBASIS_HEADER.length).setValues([PICOSTBASIS_HEADER]);
  SpreadsheetApp.flush();
  return sh;
}

// Menu: Start Batch. Snapshots current input buy cost + customs + projected
// sales/broker for product x count cycles; appends locked break-even/floor.
function recordBatch_(L, product, count) {
  ensureConfig_(L);
  const sh = ensureCostBasisTab_(L);
  const cfg = loadConfig_();
  const poco      = cfgNum_(cfg, 'poco_tax_rate', 0);
  const salesTax  = cfgNum_(cfg, 'sales_tax_rate', 0);
  const broker    = cfgNum_(cfg, 'broker_fee_rate', 0);
  const minMargin = cfgNum_(cfg, 'min_margin_pct', 0);
  const ill       = loadIlliquidSet_(L);

  const schems = loadSchematics_(L);
  const s = schems.find(x => norm_(x.name) === norm_(product));
  if (!s) throw new Error('Unknown product: ' + product);
  const cnt = Math.max(1, Number(count) || 1);

  const ids = new Set([s.output.type]); s.inputs.forEach(i => ids.add(i.type));
  const prices = loadJitaPrices_(Array.from(ids), L);
  const base   = fetchBasePrices_(ids, L);

  // Cost basis: prefer actual InventoryWAC average per input (true paid cost),
  // fall back to the market price side where there is no on-hand WAC data.
  const wac = loadInventoryWac_(L);
  let usedWac = false;
  let inputCost = 0, inputCostMkt = 0;
  s.inputs.forEach(i => {
    const q = i.qty * cnt;
    const mkt = pickPrice_(prices.get(i.type), effectivePriceSide_(i.type, 'input', ill, cfg));
    const w = wac.get(i.type);
    const unit = (w && w.avg > 0) ? w.avg : mkt;
    if (w && w.avg > 0) usedWac = true;
    inputCost += q * unit;
    inputCostMkt += q * mkt;
  });
  const costSource    = usedWac ? 'WAC' : 'market';
  const priceVariance = inputCostMkt - inputCost; // + = on-hand stock cheaper than market (favorable)

  const customsIn  = poco * s.inputs.reduce((a, i) => a + i.qty * cnt * (base.get(i.type) || 0), 0);
  const outQty     = s.output.qty * cnt;
  const customsOut = poco * outQty * (base.get(s.output.type) || 0);
  const revenue    = outQty * pickPrice_(prices.get(s.output.type), effectivePriceSide_(s.output.type, 'output', ill, cfg));
  const salesBroker = revenue * (salesTax + broker);
  const customs    = customsIn + customsOut;
  const allIn      = inputCost + customs + salesBroker;
  const breakEven  = outQty > 0 ? allIn / outQty : 0;
  const target     = breakEven * (1 + minMargin);
  const batchId    = 'B' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');

  sh.appendRow([batchId, new Date(), s.name, outQty, round2(inputCost), round2(customs),
    round2(allIn), round2(breakEven), round2(target), costSource, round2(inputCostMkt),
    round2(priceVariance), 'OPEN', '']);
  SpreadsheetApp.flush();
  L && L.log('[recordBatch_]', { batchId, product: s.name, outQty, breakEven, costSource, priceVariance: round2(priceVariance) });
  return { batchId, product: s.name, outQty, breakEven: round2(breakEven), target: round2(target),
           cost_source: costSource, price_variance: round2(priceVariance) };
}

// Menu: Check Sell Floor. Flags open batches BELOW BREAK-EVEN / BELOW TARGET / OK
// vs current market sell. Close a batch by setting its status to SOLD.
function checkSellFloor_(L) {
  const sh = SS.getSheetByName(PICOSTBASIS_TAB);
  if (!sh || sh.getLastRow() < 2) { L && L.log('[checkSellFloor_] no batches'); return { batches: 0, flagged: 0 }; }
  ensureConfig_(L);
  const cfg = loadConfig_();
  const ill = loadIlliquidSet_(L);

  const data = sh.getDataRange().getValues();
  const h = data[0].map(x => String(x).trim().toLowerCase());
  const c = { prod: h.indexOf('product'), be: h.indexOf('break_even_unit'), tg: h.indexOf('target_floor_unit'),
              st: h.indexOf('status'), nt: h.indexOf('notes') };

  const schems = loadSchematics_(L);
  const byName = new Map(); schems.forEach(s => byName.set(norm_(s.name), s));
  const ids = new Set();
  for (let i = 1; i < data.length; i++) { const s = byName.get(norm_(data[i][c.prod])); if (s) ids.add(s.output.type); }
  const prices = loadJitaPrices_(Array.from(ids), L);

  let flagged = 0;
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][c.st] || '').toUpperCase();
    if (status === 'SOLD' || status === 'CLOSED') continue;
    const s = byName.get(norm_(data[i][c.prod])); if (!s) continue;
    const cur = pickPrice_(prices.get(s.output.type), effectivePriceSide_(s.output.type, 'output', ill, cfg));
    const be = Number(data[i][c.be]) || 0, tg = Number(data[i][c.tg]) || 0;
    let flag = 'OK';
    if (cur < be) flag = 'BELOW BREAK-EVEN';
    else if (cur < tg) flag = 'BELOW TARGET';
    if (flag !== 'OK') flagged++;
    sh.getRange(i + 1, c.st + 1).setValue(flag === 'OK' ? 'OPEN' : flag);
    if (c.nt >= 0) sh.getRange(i + 1, c.nt + 1).setValue('mkt ' + round2(cur) + ' @ ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd HH:mm'));
  }
  SpreadsheetApp.flush();
  L && L.log('[checkSellFloor_]', { batches: data.length - 1, flagged });
  return { batches: data.length - 1, flagged };
}

/* =========================
   === PHASE 3: COLONIES ===
   Live colony monitoring. Menu-driven (runs as the GESI-authorized owner), so no
   web-app route is needed. Reads each char's planets -> factory pins' schematic_id
   + last_cycle_start/expiry_time -> PIColonies finish-time forecast.
   CAVEAT: ESI colony contents only refresh when the player touches the colony
   in-game. Finish time (cycle/expiry fixed at install) is reliable; live current
   contents are not. The caveat is written as a note on PIColonies!A1.
   ========================= */

const PICOLONIES_TAB = 'PIColonies';
const COLONY_STALENESS_NOTE =
  'Forecast only. ESI colony data refreshes when you touch the colony in-game, so ' +
  'finish times are reliable (cycle/expiry fixed at install) but live contents are not.';

// All schematics (no tier filter) -> Map(schematic_id -> { name, cycle, output type_id }).
// Factories at any tier (P1..P4) appear in colonies, so this is unfiltered.
function loadSchematicIndex_(L) {
  const { header, rows } = loadTabRows_(SCHEMATICS_TAB);
  const map = new Map();
  if (!rows.length) return map;
  const c = { id: header.indexOf('schematic_id'), name: header.indexOf('schematic_name'),
              cycle: header.indexOf('cycle_time_sec'), type: header.indexOf('type_id'),
              inp: header.indexOf('is_input') };
  const types = loadTypeMap_(L);
  rows.forEach(r => {
    const sid = Number(r[c.id]); if (!sid) return;
    let s = map.get(sid);
    if (!s) { s = { name: String(r[c.name] || ''), cycle: Number(r[c.cycle]) || 0, output: 0, outputName: '', tier: null }; map.set(sid, s); }
    const isInput = r[c.inp] === 1 || r[c.inp] === true || String(r[c.inp]) === '1';
    if (!isInput) {
      const tid = Number(r[c.type]);
      s.output = tid;
      const meta = types.get(tid);
      s.outputName = meta ? meta.name : '';
      s.tier = tierOf_(meta ? meta.group : '');
    }
  });
  return map;
}

// GESI's planet-detail endpoint name varies by library version. Probe candidates
// once, cache the working one, reuse. (planets LIST is characters_character_planets;
// the DETAIL operationId differs across GESI releases.)
let PLANET_DETAIL_EP_ = null;
const PLANET_DETAIL_EP_CANDIDATES = [
  'characters_character_planets_planet',
  'characters_character_planet',
  'characters_character_planets_planet_id',
  'characters_character_id_planet',
];
function invokePlanetDetail_(name, planetId, L) {
  const params = { name, planet_id: planetId, show_column_headings: true };
  if (PLANET_DETAIL_EP_) return GESI.invoke(PLANET_DETAIL_EP_, params);
  let lastErr = null;
  for (let i = 0; i < PLANET_DETAIL_EP_CANDIDATES.length; i++) {
    const ep = PLANET_DETAIL_EP_CANDIDATES[i];
    try {
      const r = GESI.invoke(ep, params);
      PLANET_DETAIL_EP_ = ep;
      L && L.log('[invokePlanetDetail_] using endpoint', { ep });
      return r;
    } catch (e) { lastErr = e; }
  }
  throw new Error('No working planet-detail endpoint. Tried [' +
    PLANET_DETAIL_EP_CANDIDATES.join(', ') + ']. Last error: ' + String(lastErr));
}

// characters_character_planets_planet returns ONE row of 3 columns (links, pins,
// routes), each a JSON-string array. Parse the pins array -> [{pin_id, type_id,
// schematic_id?, last_cycle_start?, expiry_time?, contents?}, ...].
function parsePlanetPins_(detail) {
  if (!detail || detail.length < 2) return [];
  const dh = detail[0].map(x => String(x).trim().toLowerCase());
  const pinsIdx = dh.indexOf('pins');
  if (pinsIdx < 0) return [];
  const cell = detail[1][pinsIdx];
  if (cell == null || cell === '') return [];
  if (Array.isArray(cell)) return cell; // some GESI versions hand back a parsed array
  try { return JSON.parse(String(cell)); } catch (e) { return []; }
}

// GESI/ESI timestamps may arrive as a Date or an ISO string. Normalize to Date|null.
function parseEsiDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// Menu: Refresh Colonies. Scans every authorized char's planets, groups factory
// pins by schematic per planet, and writes PIColonies with the next finish time.
function buildPiColonies_(L) {
  const names = GESI.getAuthenticatedCharacterNames();
  if (!names || !names.length) throw new Error('No authorized characters in GESI.');
  const schemIdx = loadSchematicIndex_(L);
  if (!schemIdx.size) L && L.log('[buildPiColonies_] SDE_Schematics empty - cycle times unknown');
  const uni = loadUniverse_(L);

  const out = [['character', 'system', 'planet_id', 'planet_type', 'schematic_name',
    'factory_count', 'next_finish_time', 'program_expiry']];
  let planetsScanned = 0, colonyRows = 0;

  names.forEach(name => {
    let planets;
    try { planets = GESI.invoke('characters_character_planets', { name, show_column_headings: true }); }
    catch (e) { L && L.log('[buildPiColonies_] planets failed', { name, error: String(e) }); return; }
    if (!planets || planets.length < 2) return;
    const ph = planets[0].map(x => String(x).trim().toLowerCase());
    const pcol = { pid: ph.indexOf('planet_id'), ptype: ph.indexOf('planet_type'), sys: ph.indexOf('solar_system_id') };

    for (let r = 1; r < planets.length; r++) {
      const planetId = Number(planets[r][pcol.pid]); if (!planetId) continue;
      planetsScanned++;
      const sysId = pcol.sys >= 0 ? Number(planets[r][pcol.sys]) || 0 : 0;
      const sysMeta = sysId ? uni.sys.get(sysId) : null;
      const sysName = sysMeta ? sysMeta.name : ('System ' + sysId);
      const ptype = pcol.ptype >= 0 ? planets[r][pcol.ptype] : '';

      let detail;
      try { detail = invokePlanetDetail_(name, planetId, L); }
      catch (e) { L && L.log('[buildPiColonies_] planet detail failed', { name, planetId, error: String(e) }); continue; }
      // Group factory pins (those carrying a schematic_id) by schematic.
      const groups = new Map(); // schematic_id -> { count, nextFinish, expiry }
      parsePlanetPins_(detail).forEach(pin => {
        const sid = Number(pin.schematic_id) || 0;
        if (!sid) return; // non-factory pin (extractor head / storage / launchpad)
        const meta = schemIdx.get(sid);
        const cycle = meta ? meta.cycle : 0;
        const lcs = parseEsiDate_(pin.last_cycle_start);
        const exp = parseEsiDate_(pin.expiry_time);
        let finish = null;
        if (lcs && cycle) finish = new Date(lcs.getTime() + cycle * 1000); // next cycle completes
        else if (exp) finish = exp;
        const g = groups.get(sid) || { count: 0, nextFinish: null, expiry: null };
        g.count++;
        if (finish && (!g.nextFinish || finish < g.nextFinish)) g.nextFinish = finish;
        if (exp && (!g.expiry || exp < g.expiry)) g.expiry = exp;
        groups.set(sid, g);
      });

      groups.forEach((g, sid) => {
        const meta = schemIdx.get(sid);
        out.push([name, sysName, planetId, ptype, meta ? meta.name : ('Schematic ' + sid),
          g.count, g.nextFinish || '', g.expiry || '']);
        colonyRows++;
      });
    }
  });

  writeAllToStaging_(PICOLONIES_TAB + '__staging', out, L);
  swapStagingToLive_(PICOLONIES_TAB + '__staging', PICOLONIES_TAB, L);
  try { SS.getSheetByName(PICOLONIES_TAB).getRange(1, 1).setNote(COLONY_STALENESS_NOTE); } catch (e) {}
  L && L.log('[buildPiColonies_]', { characters: names.length, planets: planetsScanned, rows: colonyRows });
  return { characters: names.length, planets: planetsScanned, colony_rows: colonyRows };
}

/* =========================
   === PI INVENTORY GRID (#3) ===
   Grid of PI on-hand: rows = items used by P3/P4 schematics (inputs+outputs),
   columns = INCLUDED characters, cell = qty held in approved locations.
   ========================= */

// Menu: Build PI Inventory Grid.
function buildPiInventoryGrid_(L) {
  const chars = loadIncludedCharacters_();
  if (!chars.length) throw new Error('No included characters. Run Refresh Character List and mark include=TRUE.');
  ensurePiFilterTab_(L);
  const schems = loadSchematics_(L);
  if (!schems.length) throw new Error('No P3/P4 schematics. Upload SDE_Schematics first.');

  const typeName = new Map(); // type_id -> item name
  schems.forEach(s => { typeName.set(s.output.type, s.output.name); s.inputs.forEach(i => typeName.set(i.type, i.name)); });

  const inv = loadTabRows_(ASSET_SHEET_NAME);
  if (!inv.rows.length) throw new Error('AllInventory empty. Run Refresh Assets first.');
  const ic = invCols_(inv.header);
  const filter = loadLocationFilter_();
  const charSet = new Set(chars);

  const grid = new Map(); // type_id -> Map(char -> qty)
  typeName.forEach((_, tid) => grid.set(tid, new Map()));
  inv.rows.forEach(r => {
    if (!rowPassesFilter_(r, filter, ic)) return;
    const tid = Number(r[ic.type]); if (!grid.has(tid)) return;
    const own = String(ic.owner >= 0 ? r[ic.owner] : '').trim(); if (!charSet.has(own)) return;
    const m = grid.get(tid);
    m.set(own, (m.get(own) || 0) + (Number(r[ic.qty]) || 0));
  });

  const tids = Array.from(typeName.keys()).sort((a, b) => String(typeName.get(a)).localeCompare(String(typeName.get(b))));
  const out = [['pi_item', 'type_id'].concat(chars).concat(['total'])];
  tids.forEach(tid => {
    const m = grid.get(tid);
    let tot = 0;
    const row = [typeName.get(tid), tid];
    chars.forEach(ch => { const q = m.get(ch) || 0; tot += q; row.push(q); });
    row.push(tot);
    out.push(row);
  });
  writeAllToStaging_(PIINVENTORY_TAB + '__staging', out, L);
  swapStagingToLive_(PIINVENTORY_TAB + '__staging', PIINVENTORY_TAB, L);
  L && L.log('[buildPiInventoryGrid_]', { items: tids.length, characters: chars.length });
  return { items: tids.length, characters: chars.length };
}

/* =========================
   === RUN SIZES (#4)    ===
   Launchpad-capacity run sizing per P3/P4 schematic. P3: one launchpad per input,
   cycles limited by the tightest input (output spread across the input launchpads).
   P4: size by output filling one launchpad; runs_per_planet = p4_launchpads_per_planet.
   Units derive from live SDE volume; the user's worked examples (e.g. 13333 P2 -> P3,
   1200 P3 -> 200 P4) are a sanity check, not hardcoded.
   ========================= */

// Map(schematic_id -> { id, name, tier, outputName, outputType, outputPerRun,
//   cycles, lpUsed, runsPerPlanet, inputs:[{type,name,unitsPerRun}] }). P3/P4 only.
function computeRunSizes_(L) {
  const cfg = loadConfig_();
  const lp   = cfgNum_(cfg, 'launchpad_m3', 10000);
  const p4lp = cfgNum_(cfg, 'p4_launchpads_per_planet', 2);
  const schems = loadSchematics_(L);
  const map = new Map();
  schems.forEach(s => {
    let cycles = 0, lpUsed = 0, runsPerPlanet = 1;
    if (s.tier === 'P3') {
      let c = Infinity;
      s.inputs.forEach(i => {
        const cap = i.vol > 0 ? Math.floor(lp / i.vol) : 0;       // units of this input per launchpad
        const cc  = i.qty > 0 ? Math.floor(cap / i.qty) : 0;      // cycles that many units support
        if (cc < c) c = cc;
      });
      cycles = isFinite(c) ? c : 0;
      lpUsed = s.inputs.length;
      runsPerPlanet = 1;
    } else { // P4: fill one launchpad with output
      const outPer = s.output.vol > 0 ? Math.floor(lp / s.output.vol) : 0;
      cycles = s.output.qty > 0 ? Math.floor(outPer / s.output.qty) : 0;
      lpUsed = 1;
      runsPerPlanet = p4lp;
    }
    const inputs = s.inputs.map(i => ({ type: i.type, name: i.name, unitsPerRun: cycles * i.qty }));
    map.set(s.id, { id: s.id, name: s.name, tier: s.tier, outputName: s.output.name, outputType: s.output.type,
                    outputPerRun: cycles * s.output.qty, cycles, lpUsed, runsPerPlanet, inputs });
  });
  return map;
}

// Menu: Build Run Sizes. Writes PIRunSizes (one row per input per schematic).
function buildRunSizes_(L) {
  const runs = computeRunSizes_(L);
  if (!runs.size) throw new Error('No P3/P4 schematics. Upload SDE_Schematics first.');
  const out = [['schematic', 'tier', 'output_name', 'output_per_run', 'launchpads_used',
    'runs_per_planet', 'cycles_per_run', 'input_name', 'input_units_per_run']];
  const list = Array.from(runs.values()).sort((a, b) =>
    (a.tier + a.name).localeCompare(b.tier + b.name));
  list.forEach(r => {
    r.inputs.forEach(i => {
      out.push([r.name, r.tier, r.outputName, r.outputPerRun, r.lpUsed, r.runsPerPlanet, r.cycles, i.name, i.unitsPerRun]);
    });
  });
  writeAllToStaging_(PIRUNSIZES_TAB + '__staging', out, L);
  swapStagingToLive_(PIRUNSIZES_TAB + '__staging', PIRUNSIZES_TAB, L);
  L && L.log('[buildRunSizes_]', { schematics: runs.size, rows: out.length - 1 });
  return { schematics: runs.size, rows: out.length - 1 };
}

/* =========================
   === PLANET PRODUCTION (#7a) ===
   Auto-derived map of which INCLUDED character produces what, on which planet,
   from live factory-pin schematic_ids. Source of truth for per-char buy lists.
   ========================= */

// Menu: Refresh Planet Production. One row per (planet, schematic) for included chars.
function buildPiPlanetProduction_(L) {
  const included = loadIncludedCharacters_();
  if (!included.length) throw new Error('No included characters. Run Refresh Character List and mark include=TRUE.');
  const names = (GESI.getAuthenticatedCharacterNames() || []).filter(n => included.indexOf(n) >= 0);
  if (!names.length) throw new Error('No included characters are authenticated in GESI.');

  const idx = loadSchematicIndex_(L);   // all tiers, with name + tier
  const runs = computeRunSizes_(L);      // P3/P4 run sizes
  const uni = loadUniverse_(L);

  const recs = [];                 // {name, sysName, planetId, ptype, tier, schematic, cnt, outPer, runsPer}
  const planetIds = new Set();
  let planetsScanned = 0;

  names.forEach(name => {
    let planets;
    try { planets = GESI.invoke('characters_character_planets', { name, show_column_headings: true }); }
    catch (e) { L && L.log('[buildPiPlanetProduction_] planets failed', { name, error: String(e) }); return; }
    if (!planets || planets.length < 2) return;
    const ph = planets[0].map(x => String(x).trim().toLowerCase());
    const pcol = { pid: ph.indexOf('planet_id'), ptype: ph.indexOf('planet_type'), sys: ph.indexOf('solar_system_id') };

    for (let r = 1; r < planets.length; r++) {
      const planetId = Number(planets[r][pcol.pid]); if (!planetId) continue;
      planetsScanned++;
      planetIds.add(planetId);
      const sysId = pcol.sys >= 0 ? Number(planets[r][pcol.sys]) || 0 : 0;
      const sysMeta = sysId ? uni.sys.get(sysId) : null;
      const sysName = sysMeta ? sysMeta.name : ('System ' + sysId);
      const ptype = pcol.ptype >= 0 ? planets[r][pcol.ptype] : '';

      let detail;
      try { detail = invokePlanetDetail_(name, planetId, L); }
      catch (e) { L && L.log('[buildPiPlanetProduction_] planet detail failed', { name, planetId, error: String(e) }); continue; }

      const counts = new Map(); // schematic_id -> factory pin count
      parsePlanetPins_(detail).forEach(pin => {
        const sid = Number(pin.schematic_id) || 0;
        if (!sid) return; // non-factory pin
        counts.set(sid, (counts.get(sid) || 0) + 1);
      });
      counts.forEach((cnt, sid) => {
        const meta = idx.get(sid);
        const run = runs.get(sid);
        recs.push({ name, sysName, planetId, ptype, tier: meta ? (meta.tier || '') : '',
          schematic: meta ? meta.name : ('Schematic ' + sid), cnt,
          outPer: run ? run.outputPerRun : '', runsPer: run ? run.runsPerPlanet : '' });
      });
    }
  });

  // Resolve in-game planet names/numbers once for all scanned planets.
  const pnames = resolvePlanetNames_(Array.from(planetIds), L);

  const out = [['character', 'system', 'planet_id', 'planet_name', 'planet_number', 'planet_type',
    'tier', 'schematic_name', 'factory_count', 'output_per_run', 'runs_per_planet']];
  recs.forEach(r => {
    const pname = pnames.get(String(r.planetId)) || '';
    out.push([r.name, r.sysName, r.planetId, pname, planetNumberFromName_(pname), r.ptype,
      r.tier, r.schematic, r.cnt, r.outPer, r.runsPer]);
  });
  const prodRows = recs.length;

  writeAllToStaging_(PIPLANETPROD_TAB + '__staging', out, L);
  swapStagingToLive_(PIPLANETPROD_TAB + '__staging', PIPLANETPROD_TAB, L);
  try { SS.getSheetByName(PIPLANETPROD_TAB).getRange(1, 1).setNote(COLONY_STALENESS_NOTE); } catch (e) {}
  L && L.log('[buildPiPlanetProduction_]', { characters: names.length, planets: planetsScanned, rows: prodRows });
  return { characters: names.length, planets: planetsScanned, production_rows: prodRows };
}

/* =========================
   === REPLACEMENT SUGGESTIONS (#7b) ===
   For each producing planet, suggest the best same-tier schematic by profit/hr
   (same tier keeps the factory layout). Reads PIPlanetProduction + PIProfit.
   ========================= */

// Menu: Suggest Planet Swaps.
function suggestReplacements_(L) {
  const prod = loadTabRows_(PIPLANETPROD_TAB);
  if (!prod.rows.length) throw new Error('PIPlanetProduction empty. Run Refresh Planet Production first.');
  const prof = loadTabRows_(PIPROFIT_TAB);
  if (!prof.rows.length) throw new Error('PIProfit empty. Run Build PI Profit first.');

  const pc = { name: prof.header.indexOf('schematic_name'), tier: prof.header.indexOf('tier'), phr: prof.header.indexOf('profit_per_hr') };
  const profByName = new Map();           // norm name -> { profitHr, tier }
  const bestByTier = new Map();           // tier -> { name, profitHr }
  prof.rows.forEach(r => {
    const nm = String(r[pc.name] || ''); if (!nm) return;
    const tier = String(r[pc.tier] || '');
    const phr = Number(r[pc.phr]) || 0;
    profByName.set(norm_(nm), { profitHr: phr, tier });
    const b = bestByTier.get(tier);
    if (!b || phr > b.profitHr) bestByTier.set(tier, { name: nm, profitHr: phr });
  });

  const dc = { ch: prod.header.indexOf('character'), sys: prod.header.indexOf('system'),
               pid: prod.header.indexOf('planet_id'), pname: prod.header.indexOf('planet_name'),
               pnum: prod.header.indexOf('planet_number'), tier: prod.header.indexOf('tier'),
               name: prod.header.indexOf('schematic_name') };
  const recs = [];
  prod.rows.forEach(r => {
    const curName = String(r[dc.name] || ''); if (!curName) return;
    const tier = String(r[dc.tier] || '');
    const best = bestByTier.get(tier);
    if (!best) return;
    const cur = profByName.get(norm_(curName));
    const curHr = cur ? cur.profitHr : 0;
    const delta = best.profitHr - curHr;
    if (norm_(best.name) === norm_(curName) || delta <= 0) return; // already best (or no gain)
    const pnum = dc.pnum >= 0 ? (Number(r[dc.pnum]) || 0) : 0;
    recs.push({ ch: String(r[dc.ch] || ''), sys: r[dc.sys], pid: r[dc.pid],
      pname: dc.pname >= 0 ? r[dc.pname] : '', pnum, tier, curName, curHr,
      best: best.name, bestHr: best.profitHr, delta });
  });
  // Sort by character, then in-game planet number.
  recs.sort((a, b) => a.ch.localeCompare(b.ch) || a.pnum - b.pnum);

  const out = [['character', 'system', 'planet_id', 'planet_name', 'planet_number', 'tier',
    'current_schematic', 'current_profit_hr', 'suggested_schematic', 'suggested_profit_hr', 'delta_per_hr']];
  recs.forEach(r => out.push([r.ch, r.sys, r.pid, r.pname, r.pnum, r.tier,
    r.curName, round2(r.curHr), r.best, round2(r.bestHr), round2(r.delta)]));
  const suggestions = recs.length;
  writeAllToStaging_(PIREPLACE_TAB + '__staging', out, L);
  swapStagingToLive_(PIREPLACE_TAB + '__staging', PIREPLACE_TAB, L);
  L && L.log('[suggestReplacements_]', { planets: prod.rows.length, suggestions });
  return { planets: prod.rows.length, suggestions };
}

/* =========================
   === PER-CHARACTER BUY LIST (#8) ===
   For each included char, demand = their PIPlanetProduction x run-size inputs,
   minus that char's filtered on-hand. Priced via effectivePriceSide_.
   ========================= */

// Build Map(norm schematic name -> run) from computeRunSizes_, plus the raw map.
function runsByName_(L) {
  const runs = computeRunSizes_(L);
  const byName = new Map();
  runs.forEach(r => byName.set(norm_(r.name), r));
  return { runs, byName };
}

// Per-char on-hand for the given type set: Map(owner -> Map(type_id -> qty)), filtered.
function onHandByChar_(typeSet, L) {
  const inv = loadTabRows_(ASSET_SHEET_NAME);
  const ic = invCols_(inv.header);
  const filter = loadLocationFilter_();
  const byChar = new Map();
  inv.rows.forEach(r => {
    if (!rowPassesFilter_(r, filter, ic)) return;
    const tid = Number(r[ic.type]); if (typeSet && !typeSet.has(tid)) return;
    const own = String(ic.owner >= 0 ? r[ic.owner] : '').trim(); if (!own) return;
    let m = byChar.get(own); if (!m) { m = new Map(); byChar.set(own, m); }
    m.set(tid, (m.get(tid) || 0) + (Number(r[ic.qty]) || 0));
  });
  return byChar;
}

// Aggregate per-(char,input) demand from PIPlanetProduction x run sizes.
// Returns { demand: Map(char -> Map(type -> {qty,name})), outputByKey, prodRows }.
function planetInputDemand_(L, opts) {
  opts = opts || {};
  const tiers = opts.tiers || null; // Set of tier strings (e.g. {'P3','P4'}); null = no planet-tier filter
  const prod = loadTabRows_(PIPLANETPROD_TAB);
  if (!prod.rows.length) throw new Error('PIPlanetProduction empty. Run Refresh Planet Production first.');
  const dc = { ch: prod.header.indexOf('character'), name: prod.header.indexOf('schematic_name'), tier: prod.header.indexOf('tier') };
  const { byName } = runsByName_(L);
  const demand = new Map(); // char -> Map(type -> {qty, name})
  let usedRows = 0;
  prod.rows.forEach(r => {
    const ch = String(r[dc.ch] || '').trim(); if (!ch) return;
    if (tiers && dc.tier >= 0 && !tiers.has(String(r[dc.tier] || '').trim().toUpperCase())) return;
    const run = byName.get(norm_(String(r[dc.name] || '')));
    if (!run) return; // non P3/P4 (no run size)
    usedRows++;
    let m = demand.get(ch); if (!m) { m = new Map(); demand.set(ch, m); }
    run.inputs.forEach(i => {
      const cur = m.get(i.type) || { qty: 0, name: i.name, outputs: new Set() };
      cur.qty += i.unitsPerRun * run.runsPerPlanet; // one planet row = runsPerPlanet runs
      if (run.outputName) cur.outputs.add(run.outputName);
      m.set(i.type, cur);
    });
  });
  return { demand, prodRows: usedRows };
}

// Shared: per-(char,type) to-buy lines = planet demand (optionally tier-filtered)
// minus that char's filtered on-hand. Returns { lines, typeSet, chars, prodRows }.
function perCharBuyLines_(L, opts) {
  const { demand, prodRows } = planetInputDemand_(L, opts);
  if (!demand.size) throw new Error('No P3/P4 production found in PIPlanetProduction.');
  const typeSet = new Set();
  demand.forEach(m => m.forEach((_, tid) => typeSet.add(tid)));
  const onHand = onHandByChar_(typeSet, L);
  const chars = Array.from(demand.keys()).sort((a, b) => a.localeCompare(b));
  const lines = [];
  chars.forEach(ch => {
    const m = demand.get(ch);
    const have = onHand.get(ch) || new Map();
    const tids = Array.from(m.keys()).sort((a, b) => String(m.get(a).name).localeCompare(String(m.get(b).name)));
    tids.forEach(tid => {
      const d = m.get(tid);
      const need = Math.ceil(d.qty);
      const onhand = have.get(tid) || 0;
      const output = d.outputs ? Array.from(d.outputs).sort().join(', ') : '';
      lines.push({ ch, tid, name: d.name, output, need, onhand, buy: Math.max(0, need - onhand) });
    });
  });
  return { lines, typeSet, chars, prodRows };
}

// Menu: Build Per-Char Buy List. Prices at Jita sell; m3 = qty_to_buy x unit volume.
function buildPiBuyListByChar_(L) {
  ensurePiFilterTab_(L); ensureConfig_(L);
  const { lines, typeSet, chars, prodRows } = perCharBuyLines_(L, {});
  const prices = loadJitaPrices_(Array.from(typeSet), L);
  const types = loadTypeMap_(L);

  lines.sort((a, b) => a.ch.localeCompare(b.ch) || String(a.name).localeCompare(String(b.name)));
  const out = [['output_name', 'character', 'input_name', 'type_id', 'qty_needed', 'qty_on_hand', 'qty_to_buy', 'unit_price', 'm3', 'line_cost']];
  lines.forEach(ln => {
    const unit = pickPrice_(prices.get(ln.tid), 'sell');
    const vol = (types.get(ln.tid) || {}).volume || 0;
    out.push([ln.output, ln.ch, ln.name, ln.tid, ln.need, ln.onhand, ln.buy, round2(unit), round2(ln.buy * vol), round2(ln.buy * unit)]);
  });
  writeAllToStaging_(PIBUYLISTCHAR_TAB + '__staging', out, L);
  swapStagingToLive_(PIBUYLISTCHAR_TAB + '__staging', PIBUYLISTCHAR_TAB, L);
  setSheetFilter_(PIBUYLISTCHAR_TAB, L);
  L && L.log('[buildPiBuyListByChar_]', { characters: chars.length, lines: lines.length, production_rows: prodRows });
  return { characters: chars.length, lines: lines.length };
}

// Menu: Build Buy List By Item. Aggregates qty_to_buy across all chars into one
// line per item (Jita sell price, m3, cost) + a TOTAL ISK / TOTAL M3 block.
// CONFIG buylist_target_tier_only=TRUE restricts to planets in buylist_target_tiers.
function buildPiBuyListByItem_(L) {
  ensurePiFilterTab_(L); ensureConfig_(L);
  const cfg = loadConfig_();
  const tierOnly = cfgStr_(cfg, 'buylist_target_tier_only', 'TRUE').toUpperCase() !== 'FALSE';
  const tiers = tierOnly
    ? new Set(cfgStr_(cfg, 'buylist_target_tiers', 'P3,P4').split(',').map(s => s.trim().toUpperCase()).filter(Boolean))
    : null;
  const { lines, typeSet } = perCharBuyLines_(L, { tiers });
  const prices = loadJitaPrices_(Array.from(typeSet), L);
  const types = loadTypeMap_(L);

  const agg = new Map(); // type_id -> { name, tier, qty }
  lines.forEach(ln => {
    if (ln.buy <= 0) return;
    let a = agg.get(ln.tid);
    if (!a) { const meta = types.get(ln.tid) || {}; a = { name: ln.name, tier: tierOf_(meta.group) || '', qty: 0 }; agg.set(ln.tid, a); }
    a.qty += ln.buy;
  });

  const out = [['input_name', 'type_id', 'tier', 'quantity', 'm3', 'unit_price', 'cost']];
  let totalIsk = 0, totalM3 = 0;
  const tids = Array.from(agg.keys()).sort((x, y) => String(agg.get(x).name).localeCompare(String(agg.get(y).name)));
  tids.forEach(tid => {
    const a = agg.get(tid);
    const unit = pickPrice_(prices.get(tid), 'sell');
    const vol = (types.get(tid) || {}).volume || 0;
    const m3 = a.qty * vol, cost = a.qty * unit;
    totalIsk += cost; totalM3 += m3;
    out.push([a.name, tid, a.tier, a.qty, round2(m3), round2(unit), round2(cost)]);
  });
  // Totals block (each row padded to header width for setValues).
  out.push(['', '', '', '', '', '', '']);
  out.push(['TOTAL M3', '', '', '', round2(totalM3), '', '']);
  out.push(['TOTAL ISK (Jita Sell)', '', '', '', '', '', round2(totalIsk)]);

  writeAllToStaging_(PIBUYLISTITEM_TAB + '__staging', out, L);
  swapStagingToLive_(PIBUYLISTITEM_TAB + '__staging', PIBUYLISTITEM_TAB, L);
  L && L.log('[buildPiBuyListByItem_]', { items: agg.size, total_isk: round2(totalIsk), total_m3: round2(totalM3), tier_only: tierOnly });
  return { items: agg.size, total_isk: round2(totalIsk), total_m3: round2(totalM3) };
}

/* =========================
   === INPUT USAGE LISTS (#9) ===
   Two views of consumption (character, input_material, num_inputs, output) from
   PIPlanetProduction x run sizes: one sorted by character, one by input material.
   ========================= */

// Menu: Build Input Usage Lists.
function buildInputUsageLists_(L) {
  const prod = loadTabRows_(PIPLANETPROD_TAB);
  if (!prod.rows.length) throw new Error('PIPlanetProduction empty. Run Refresh Planet Production first.');
  const dc = { ch: prod.header.indexOf('character'), name: prod.header.indexOf('schematic_name') };
  const { byName } = runsByName_(L);

  const agg = new Map(); // char|output|input -> { char, input, output, num }
  prod.rows.forEach(r => {
    const ch = String(r[dc.ch] || '').trim(); if (!ch) return;
    const run = byName.get(norm_(String(r[dc.name] || '')));
    if (!run) return;
    run.inputs.forEach(i => {
      const key = ch + '|' + run.outputName + '|' + i.name;
      const cur = agg.get(key) || { char: ch, input: i.name, output: run.outputName, num: 0 };
      cur.num += i.unitsPerRun * run.runsPerPlanet;
      agg.set(key, cur);
    });
  });
  const rows = Array.from(agg.values());
  const header = ['character', 'input_material', 'num_inputs', 'output'];

  const byCharRows = rows.slice().sort((a, b) =>
    a.char.localeCompare(b.char) || a.input.localeCompare(b.input) || a.output.localeCompare(b.output));
  const byMatRows = rows.slice().sort((a, b) =>
    a.input.localeCompare(b.input) || a.char.localeCompare(b.char) || a.output.localeCompare(b.output));
  const toGrid = list => [header].concat(list.map(r => [r.char, r.input, r.num, r.output]));

  writeAllToStaging_(PIINPUTSCHAR_TAB + '__staging', toGrid(byCharRows), L);
  swapStagingToLive_(PIINPUTSCHAR_TAB + '__staging', PIINPUTSCHAR_TAB, L);
  writeAllToStaging_(PIINPUTSMAT_TAB + '__staging', toGrid(byMatRows), L);
  swapStagingToLive_(PIINPUTSMAT_TAB + '__staging', PIINPUTSMAT_TAB, L);
  L && L.log('[buildInputUsageLists_]', { rows: rows.length });
  return { rows: rows.length };
}
