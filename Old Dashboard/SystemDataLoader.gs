/**
 * EVE SYSTEM LOOKUP - ON DEMAND
 * 
 * Provides on-demand fetching of system details (name, region, etc.)
 * with local caching to avoid repeated API calls.
 * 
 * Usage in sheets: =getSystemName(30000001) or =getSystemInfo(30000001, "region_id")
 */

// Cache object to store fetched systems (persists during script execution)
const systemCache = {};

/**
 * Get system name by system ID
 * Fetches from cache if available, otherwise calls GESI and caches result
 * 
 * @param {number} systemId - The EVE system ID
 * @returns {string} System name (e.g., "Jita")
 */
function getSystemName(systemId) {
  try {
    if (!systemId || isNaN(systemId)) {
      return "INVALID_ID";
    }
    
    const info = getSystemInfo(systemId, "name");
    return info || "UNKNOWN";
  } catch (error) {
    Logger.log(`ERROR in getSystemName(${systemId}): ${error.message}`);
    return "ERROR";
  }
}


/**
 * Get system region ID by system ID
 * 
 * @param {number} systemId - The EVE system ID
 * @returns {number} Region ID
 */
function getSystemRegionId(systemId) {
  try {
    if (!systemId || isNaN(systemId)) {
      return "";
    }
    
    return getSystemInfo(systemId, "region_id");
  } catch (error) {
    Logger.log(`ERROR in getSystemRegionId(${systemId}): ${error.message}`);
    return "";
  }
}


/**
 * Get system constellation ID by system ID
 * 
 * @param {number} systemId - The EVE system ID
 * @returns {number} Constellation ID
 */
function getSystemConstellationId(systemId) {
  try {
    if (!systemId || isNaN(systemId)) {
      return "";
    }
    
    return getSystemInfo(systemId, "constellation_id");
  } catch (error) {
    Logger.log(`ERROR in getSystemConstellationId(${systemId}): ${error.message}`);
    return "";
  }
}


/**
 * Get any system info field by system ID
 * 
 * CORE FUNCTION: Manages caching and GESI API calls
 * 
 * @param {number} systemId - The EVE system ID
 * @param {string} field - The field to retrieve (e.g., "name", "region_id", "security_status")
 * @returns {*} The requested field value, or empty string if not found
 */
function getSystemInfo(systemId, field) {
  try {
    if (!systemId || isNaN(systemId)) {
      throw new Error("Invalid systemId");
    }
    
    if (!field || typeof field !== 'string') {
      throw new Error("Invalid field name");
    }
    
    // Check cache first
    if (systemCache[systemId]) {
      Logger.log(`[CACHE HIT] System ${systemId}`);
      const value = systemCache[systemId][field];
      return value !== undefined ? value : "";
    }
    
    // Not in cache - fetch from GESI
    Logger.log(`[CACHE MISS] Fetching system ${systemId} from GESI...`);
    const systemDetail = GESI.universe_systems_system(systemId, "en", false, "v4");
    
    if (!systemDetail || systemDetail.length === 0) {
      throw new Error(`No data returned for system ${systemId}`);
    }
    
    // systemDetail is a 1D array when show_column_headings=false
    // Parse it into an object using known field order
    const systemObj = parseSystemDetail(systemDetail);
    
    // Store in cache
    systemCache[systemId] = systemObj;
    Logger.log(`[CACHED] System ${systemId}: ${systemObj.name}`);
    
    // Return requested field
    const value = systemObj[field];
    return value !== undefined ? value : "";
    
  } catch (error) {
    Logger.log(`ERROR in getSystemInfo(${systemId}, ${field}): ${error.message}`);
    return "";
  }
}


/**
 * HELPER: Parse GESI system detail array into object
 * 
 * Converts the 1D array returned by universe_systems_system() 
 * into a named object for easier field access
 * 
 * Field order from ESI v4:
 * 0:system_id, 1:name, 2:security_status, 3:constellation_id,
 * 4:region_id, 5:star_id, 6:starbases, 7:stations, 8:planets,
 * 9:asteroids, 10:security_class
 * 
 * @param {Array} detailArray - Raw array from universe_systems_system()
 * @returns {Object} Parsed system object with named fields
 */
function parseSystemDetail(detailArray) {
  const fieldNames = [
    "system_id",
    "name",
    "security_status",
    "constellation_id",
    "region_id",
    "star_id",
    "starbases",
    "stations",
    "planets",
    "asteroids",
    "security_class"
  ];
  
  const systemObj = {};
  
  for (let i = 0; i < fieldNames.length && i < detailArray.length; i++) {
    systemObj[fieldNames[i]] = detailArray[i];
  }
  
  return systemObj;
}


/**
 * CACHE MANAGEMENT: Clear the system cache
 * 
 * Use this if you need to force a refresh of cached systems
 * Note: Cache is automatically cleared when script ends
 */
function clearSystemCache() {
  const cacheSize = Object.keys(systemCache).length;
  for (const key in systemCache) {
    delete systemCache[key];
  }
  Logger.log(`Cleared system cache (${cacheSize} entries removed)`);
}


/**
 * CACHE MANAGEMENT: Show cache statistics
 */
function getSystemCacheStats() {
  const cacheSize = Object.keys(systemCache).length;
  Logger.log(`System cache contains ${cacheSize} entries`);
  return cacheSize;
}