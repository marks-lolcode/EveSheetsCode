/* ===========================================================
 * Janice_menu_updated.gs.js
 * Menu-only Janice price refresh (batch, cached, deduped, sorted)
 * =========================================================== */

/**
 * This file is designed to run ONLY from a spreadsheet menu item.
 * It does not provide custom functions (which are prone to timeouts).
 *
 * Named ranges expected:
 * - JaniceCommodityName : range containing item names or type IDs (one per cell)
 * - JaniceOutput        : a single-cell anchor (top-left) for where to write results
 *
 * Script Properties expected:
 * - JANICE_API_KEY      : your Janice API key
 */

/** Toggle debug logging here (true = very chatty logs). */
const JANICE_DEBUG = false;

/** How many items to send per API request (keeps payload sizes reasonable). */
const JANICE_CHUNK_SIZE = 200;

/** Cache TTL in seconds (6 hours). */
const JANICE_CACHE_TTL_SECONDS = 6 * 60 * 60;

// We cache per-item (small) results to avoid CacheService value-size limits.
// Key is stable across runs: marketId + normalized item key.
const JANICE_ITEM_CACHE_PREFIX = "janice_item_v2";

function buildItemCacheKey_(marketId, itemKey) {
  const normalized = String(itemKey).trim().replace(/\s+/g, " ");
  return `${JANICE_ITEM_CACHE_PREFIX}|m:${marketId}|i:${normalized}`;
}

function cacheGetPrice_(cache, marketId, itemKey) {
  const key = buildItemCacheKey_(marketId, itemKey);
  const raw = cache.get(key);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    // If the cache entry is corrupt for any reason, remove it and treat as a miss.
    cache.remove(key);
    return null;
  }
}

function cachePutPrice_(cache, marketId, itemKey, buy, sell) {
  const key = buildItemCacheKey_(marketId, itemKey);
  const payload = JSON.stringify({ buy: buy, sell: sell, cachedAt: Date.now() });

  // Defensive: avoid putting unexpectedly large payloads.
  if (payload.length > 80000) return;

  cache.put(key, payload, JANICE_CACHE_TTL_SECONDS);
}


/** Janice base URL. */
const JANICE_API_URL = "https://janice.e-351.com/api/rest/v2";






//*********************************************************************************************************

/**
 * Refresh Janice commodity prices for multiple markets (e.g., Jita and Amarr).
 * Fetches current prices from the Janice API (no cache), logs all missing data,
 * and outputs cleanly starting from row 2 (header stays in row 1).
 */
function refreshJaniceCommodityPrices() {
  const runId = new Date().toISOString();
  const log = makeLogger_("Janice", runId);
  log.info("START refreshJaniceCommodityPrices");

  const apiKey = getRequiredScriptProperty_("JANICE_API_KEY");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const JANICE_MARKETS = {
    "jita": 2,
    "r1o-gn": 3,
    "perimeter": 4,
    "jitameter": 5,
    "npc": 6,
    "mj-5f9": 114,
    "amarr": 115
  };

  let rawItems = readNamedRangeAsList_(ss, "JaniceCommodityName", log);

  // ✅ Strictly skip only actual header match, not first data row
  rawItems = rawItems.filter((v, i) => String(v).toLowerCase() !== "janicecommodityname" && String(v).trim() !== "");

  const items = rawItems.map(i => String(i).trim()).filter(v => v);
  if (items.length === 0) {
    log.warn("No items found in named range JaniceCommodityName. Nothing to do.");
    return;
  }

  const dedupeResult = dedupePreserveOrder_(items);
  log.info(
    "Loaded items. raw=%s deduped=%s duplicatesRemoved=%s",
    items.length,
    dedupeResult.unique.length,
    dedupeResult.duplicatesRemoved
  );

  const marketsToFetch = [
    { key: "jita", id: JANICE_MARKETS["jita"] },
    { key: "amarr", id: JANICE_MARKETS["amarr"] }
  ];

  const allPrices = new Map();
  const missingMap = new Map();
  dedupeResult.unique.forEach(item => {
    allPrices.set(String(item), {});
    missingMap.set(String(item), []);
  });

  marketsToFetch.forEach(({ key: marketKey, id: marketId }) => {
    const chunks = chunkArray_(dedupeResult.unique, JANICE_CHUNK_SIZE);
    log.info("Fetching %s items for %s in %s chunks", dedupeResult.unique.length, marketKey, chunks.length);

    let totalMissing = 0;

    chunks.forEach((chunkItems, idx) => {
      const chunkInputs = chunkItems.map(i => String(i).trim());
      log.debug("Sending chunk %s/%s to %s: %s", idx + 1, chunks.length, marketKey, JSON.stringify(chunkInputs));

      let chunkData;
      try {
        chunkData = janicePricer_(apiKey, marketId, chunkInputs, log);
      } catch (e) {
        log.error("Chunk %s/%s for %s failed: %s", idx + 1, chunks.length, marketKey, e);
        chunkInputs.forEach(name => {
          const entry = missingMap.get(name);
          if (entry && !entry.includes(marketKey)) entry.push(marketKey);
        });
        return; // Skip this chunk on failure
      }

      const resultByName = new Map();
      chunkData.forEach(d => {
        if (d?.itemType?.name) {
          resultByName.set(String(d.itemType.name).trim(), d);
        }
      });

      chunkItems.forEach(name => {
        name = String(name).trim();
        const rowObj = resultByName.get(name);
        const entry = allPrices.get(name) || {};

        if (!rowObj || !rowObj.immediatePrices) {
          totalMissing++;
          log.warn("No price found for '%s' in %s", name, marketKey);
          entry[marketKey] = { buy: "", sell: "" };
          const missing = missingMap.get(name);
          if (missing && !missing.includes(marketKey)) missing.push(marketKey);
        } else {
          const buy = safeNumber_(rowObj.immediatePrices?.buyPrice);
          const sell = safeNumber_(rowObj.immediatePrices?.sellPrice);
          entry[marketKey] = { buy, sell, source: "api" };
        }

        allPrices.set(name, entry);
      });
    });

    log.info("Missing prices for %s: %s items", marketKey, totalMissing);
  });

  const priceRows = [];
  let fullyMissingRows = 0;

  dedupeResult.unique.forEach(item => {
    const name = String(item).trim();
    const entry = allPrices.get(name) || {};
    const jita = entry.jita || { buy: "", sell: "" };
    const amarr = entry.amarr || { buy: "", sell: "" };

    if (!jita.buy && !jita.sell && !amarr.buy && !amarr.sell) {
      log.warn("No prices found for item across all markets: '%s'", name);
      fullyMissingRows++;
      return; // Skip adding to output, but keep for MissingItems
    }

    priceRows.push([name, jita.buy, jita.sell, amarr.buy, amarr.sell]);
  });

  priceRows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  const outAnchor = ss.getRangeByName("JaniceOutput");
  if (!outAnchor) throw new Error("Named range 'JaniceOutput' not found.");

  const sheet = outAnchor.getSheet();
  const startRow = outAnchor.getRow();
  const startCol = outAnchor.getColumn();

  const maxRowsToClear = sheet.getMaxRows() - startRow + 1;
  sheet.getRange(startRow, startCol, maxRowsToClear, 5).clearContent();

  if (priceRows.length > 0) {
    // ✅ Only write directly from startRow + 1, ensuring no skipped rows
    sheet.getRange(startRow + 1, startCol, priceRows.length, 5).setValues(priceRows);
    sheet.getRange(startRow + 1, startCol, priceRows.length, 5)
      .sort({ column: startCol, ascending: true });
  }

  let missingSheet = ss.getSheetByName("MissingItems");
  if (!missingSheet) missingSheet = ss.insertSheet("MissingItems");
  else missingSheet.clear();

  const missingRows = Array.from(missingMap.entries())
    .filter(([_, markets]) => markets.length > 0)
    .map(([item, markets]) => [item, markets.join(", ")]);

  if (missingRows.length) {
    missingSheet.getRange(1, 1, 1, 2).setValues([["Missing Item", "Market"]]);
    missingSheet.getRange(2, 1, missingRows.length, 2).setValues(missingRows);
  }

  log.info("DONE refreshJaniceCommodityPrices. rowsWritten=%s", priceRows.length);
  log.info("Rows with no price data at all: %s", fullyMissingRows);
} // end


//*********************************************************************************************************


// ✅ Function name remains unchanged: `refreshJaniceCommodityPrices`
// If your menu item points to this function, no change is needed.

// ✅ Function name remains unchanged: `refreshJaniceCommodityPrices`
// If your menu item points to this function, no change is needed.


/**
 * Calls Janice /pricer for a set of items in a single chunk.
 * Uses caching to avoid repeated calls for the same (market + payload).
 *
 * Janice expects:
 * - POST
 * - contentType: text/plain
 * - payload: items separated by newlines
 * - header: X-ApiKey
 *
 * Returns: parsed JSON (must be an array).
 */
function janicePricer_(apiKey, marketId, items, log) {
  const url = JANICE_API_URL + "/pricer?market=" + encodeURIComponent(String(marketId));
  const payload = items.join("\n");

  const options = {
    method: "post",
    contentType: "text/plain",
    headers: { "X-ApiKey": apiKey },
    payload: payload,
    muteHttpExceptions: true
  };

  // Build a stable cache key. We do NOT include any cache buster.
  const cacheKey = buildCacheKey_(url, options);
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (cached) {
    if (JANICE_DEBUG) log.debug("Cache hit. key=%s bytes=%s", cacheKey, cached.length);
    const parsed = JSON.parse(cached);
    if (!Array.isArray(parsed)) {
      log.error("Cached Janice response was not an array. key=%s", cacheKey);
      throw new Error("Cached Janice response was not an array.");
    }
    return parsed;
  }

  const response = fetchWithRetry_(url, options, log);
  const bodyText = response.getContentText() || "";

  // Janice errors (like HTTP 400) often return an HTML string or a JSON object.
  // We only accept an array response here.
  let json;
  try {
    json = JSON.parse(bodyText);
  } catch (e) {
    log.error("Janice returned non-JSON response. status=%s bodySnippet=%s", response.getResponseCode(), bodyText.slice(0, 300));
    throw new Error("Janice returned non-JSON response (status " + response.getResponseCode() + ").");
  }

  if (!Array.isArray(json)) {
    log.error(
      "Janice response was not an array. status=%s typeof=%s keys=%s bodySnippet=%s",
      response.getResponseCode(),
      typeof json,
      (json && typeof json === "object") ? Object.keys(json).join(",") : "(n/a)",
      bodyText.slice(0, 500)
    );
    throw new Error("Janice response was not an array. Check Logs for details.");
  }

  // NOTE:
  // We intentionally do NOT cache the full array response here.
  // Large payloads can exceed CacheService value-size limits and throw:
  //   Exception: Argument too large: value
  // Instead, we cache per-item (small) buy/sell results in the caller.

  return json;
}

/**
 * Fetch wrapper that retries ONLY on retryable status codes (429/5xx),
 * with exponential backoff. No random sleep jitter on normal requests.
 */
function fetchWithRetry_(url, options, log) {
  const maxAttempts = 5;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();

    if (code >= 200 && code < 300) {
      return resp;
    }

    const retryable = (code === 429) || (code >= 500 && code < 600);
    const snippet = (resp.getContentText() || "").slice(0, 300);

    log.error(
      "HTTP error. attempt=%s code=%s retryable=%s bodySnippet=%s",
      attempt,
      code,
      retryable,
      snippet
    );

    if (!retryable) {
      throw new Error("Janice HTTP error " + code + ": " + snippet);
    }

    // Exponential backoff: 500ms, 1000ms, 2000ms, 4000ms...
    const sleepMs = 500 * Math.pow(2, attempt - 1);
    Utilities.sleep(sleepMs);
  }

  throw new Error("Janice request failed after max retry attempts.");
}

/**
 * Reads a named range and returns a flat, trimmed list of non-empty values.
 */
function readNamedRangeAsList_(ss, rangeName, log) {
  const range = ss.getRangeByName(rangeName);
  if (!range) {
    throw new Error("Named range '" + rangeName + "' not found.");
  }

  const values = range.getValues();
  const out = [];

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const v = values[r][c];
      if (v === null || v === undefined) continue;
      const s = String(v).trim();
      if (s.length === 0) continue;
      out.push(s);
    }
  }

  if (JANICE_DEBUG) log.debug("readNamedRangeAsList_ range=%s rows=%s cols=%s out=%s", rangeName, values.length, values[0] ? values[0].length : 0, out.length);
  return out;
}

/**
 * De-duplicate while preserving original order.
 */
function dedupePreserveOrder_(list) {
  const seen = new Set();
  const unique = [];
  let duplicatesRemoved = 0;

  for (let i = 0; i < list.length; i++) {
    const v = String(list[i]);
    if (seen.has(v)) {
      duplicatesRemoved++;
      continue;
    }
    seen.add(v);
    unique.push(v);
  }

  return { unique, duplicatesRemoved };
}

/**
 * Splits an array into equally sized chunks.
 */
function chunkArray_(arr, chunkSize) {
  const out = [];
  for (let i = 0; i < arr.length; i += chunkSize) {
    out.push(arr.slice(i, i + chunkSize));
  }
  return out;
}

/**
 * Converts value to a number where possible; otherwise returns blank string.
 * This keeps Sheets output clean.
 */
function safeNumber_(v) {
  if (v === null || v === undefined || v === "") return "";
  const n = Number(v);
  if (Number.isFinite(n)) return n;
  return "";
}

/**
 * Creates a stable cache key from request parts.
 * We hash large strings to keep keys short.
 */
function buildCacheKey_(url, options) {
  const parts = [
    "u=" + url,
    "m=" + String(options.method || "get").toLowerCase(),
    "ct=" + String(options.contentType || ""),
    "h=" + JSON.stringify(options.headers || {}),
    "p=" + String(options.payload || "")
  ].join("|");

  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, parts, Utilities.Charset.UTF_8);
  return "janice:" + digest.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

/**
 * Gets a required Script Property.
 */
function getRequiredScriptProperty_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v || String(v).trim().length === 0) {
    throw new Error("Missing Script Property: " + key);
  }
  return v;
}

/**
 * Minimal structured logger with levels.
 */
function makeLogger_(prefix, runId) {
  function fmt_(level, msg, args) {
    const ts = new Date().toISOString();
    let out = "[" + prefix + "][" + ts + "][" + level + "][" + runId + "] " + msg;
    if (args && args.length) {
      // Simple %s replacement to keep logs readable for volunteers.
      args.forEach(a => { out = out.replace("%s", String(a)); });
    }
    return out;
  }

  return {
    info: function (msg) { Logger.log(fmt_("INFO", msg, Array.prototype.slice.call(arguments, 1))); },
    warn: function (msg) { Logger.log(fmt_("WARN", msg, Array.prototype.slice.call(arguments, 1))); },
    error: function (msg) { Logger.log(fmt_("ERROR", msg, Array.prototype.slice.call(arguments, 1))); },
    debug: function (msg) { if (JANICE_DEBUG) Logger.log(fmt_("DEBUG", msg, Array.prototype.slice.call(arguments, 1))); }
  };
}

/* ===========================================================
 * End of file: Janice_menu_updated.gs.js
 * =========================================================== */