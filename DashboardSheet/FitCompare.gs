/**
 * EVE Online Ship Fit Comparison Tool
 */

// Global cache for SDE lookups
let sdeCache = {};

function loadSdeCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sdeSheet = ss.getSheetByName('SDE_invTypes');
  
  if (!sdeSheet) return;
  
  Logger.log('Loading SDE cache...');
  const data = sdeSheet.getRange('A2:C60000').getValues();  // Changed from 10000 to 60000
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) {
      sdeCache[data[i][0]] = data[i][2];
    }
  }
  
  Logger.log(`Cached ${Object.keys(sdeCache).length} type IDs`);
}

function getItemNameFromTypeId(typeId) {
  if (sdeCache[typeId]) {
    return sdeCache[typeId];
  }
  return `[TypeID ${typeId}]`;
}

function getSystemForLocation(locationId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fitData = ss.getSheetByName('FitData');
  
  // Check cache first (AB2:AD500)
  const cacheData = fitData.getRange('AB2:AD500').getValues();
  for (let i = 0; i < cacheData.length; i++) {
    if (cacheData[i][0] == locationId && cacheData[i][2]) {
      Logger.log(`✓ Cache hit for location ${locationId}: ${cacheData[i][2]}`);
      return {
        locationId: locationId,
systemId: cacheData[i][1],
        systemName: cacheData[i][2],
        locationType: 'cached'
      };
    }
  }
  
  let systemId = null;
  let systemName = null;
  let locationType = null;
  
  // Station IDs: 60,000,000 to 69,999,999
  if (locationId >= 60000000 && locationId <= 69999999) {
    Logger.log(`Checking STATION for location ${locationId}...`);
    try {
      const stationData = GESI.universe_stations_station(locationId, true);
      if (stationData && stationData.length > 1 && stationData[1][1]) {
        systemId = stationData[1][3];
        systemName = stationData[1][1];
        locationType = 'station';
        Logger.log(`✓ Found station: ${systemName} (System: ${systemId})`);
        
        // Cache it
fitData.appendRow([locationId, systemId, systemName]);      }
    } catch (e) {
      Logger.log(`✗ Station lookup failed for ${locationId}: ${e.message}`);
    }
  }
  // Structure IDs: 1,000,000,000 and up
  else if (locationId >= 1000000000) {
    Logger.log(`Checking STRUCTURE for location ${locationId}...`);
    try {
      const structData = GESI.universe_structures_structure(locationId, '', true);
      if (structData && structData.length > 1 && structData[1][0]) {
        systemId = structData[1][3];   // solar_system_id
        systemName = structData[1][0]; // name
        locationType = 'structure';
        Logger.log(`✓ Found structure: ${systemName} (System: ${systemId})`);
        
        // Cache it
fitData.appendRow([locationId, systemId, systemName]);
     }
    } catch (e) {
      Logger.log(`✗ Structure lookup failed for ${locationId}: ${e.message}`);
    }
  }
  // FleetHangar or other container locations that need to be resolved
  // Will be handled by asset chain following in pullFitInventory()
  else {
    Logger.log(`ℹ Location ${locationId} requires asset chain resolution (FleetHangar or container)`);
  }
  
  // Return result
  if (systemName) {
    return { locationId, systemId, systemName, locationType };
  }
  
  Logger.log(`✗ Failed to resolve ${locationId}`);
  return { locationId, systemId: null, systemName: `[Unknown ${locationId}]`, locationType: 'unknown' };
}

function pullFitInventory() {
  loadSdeCache();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fitCompare = ss.getSheetByName('FitCompare');
  const fitData = ss.getSheetByName('FitData');
  
  // Load ship type IDs (groupID 25-26 are ships)
  const shipTypeIds = new Set();
  const sdeSheet = ss.getSheetByName('SDE_invTypes');
  if (sdeSheet) {
    const typeData = sdeSheet.getRange('A2:D60000').getValues();
    for (let i = 0; i < typeData.length; i++) {
      if (typeData[i][1] >= 25 && typeData[i][1] <= 26) {
        shipTypeIds.add(typeData[i][0]);
      }
    }
    Logger.log(`Loaded ${shipTypeIds.size} ship type IDs`);
  }
  
  // Get selected characters
  const selectedCharsRange = fitCompare.getRange('K3:L11');
  const selectedCharsData = selectedCharsRange.getValues();
  
  const selectedChars = [];
  for (let i = 0; i < selectedCharsData.length; i++) {
    if (selectedCharsData[i][0] && selectedCharsData[i][1] === true) {
      selectedChars.push(selectedCharsData[i][0]);
    }
  }
  
  if (selectedChars.length === 0) {
    Logger.log('No characters selected.');
    return;
  }
  
  // Setup
  const locationCache = {};
  const rows = [];
  const locationFlags = ['Unlocked', 'Locked', 'Hangar', 'Deliveries', 'Cargo', 'CapsuleerDeliveries', 'InfrastructureHangar', 'FleetHangar', 'ShipHangar'];
  const skipFlags = ['AutoFit', 'Fitting'];
  
  let totalItems = 0;
  let uniqueLocations = 0;
  
  try {
    // Fetch inventory for each character
    for (const charName of selectedChars) {
      Logger.log(`Fetching assets for ${charName}...`);
      const assets = GESI.characters_character_assets(charName, false);
      
      if (!assets || assets.length < 2) {
        Logger.log(`No assets for ${charName}`);
        continue;
      }
      
      Logger.log(`${charName} has ${assets.length - 1} items total`);
      
      // Build set of fitted ship itemIds (ships that have items inside them)
      const fittedShips = new Set();
      for (let i = 1; i < assets.length; i++) {
        const locationId = assets[i][4];
        // Check if this locationId is a ship's itemId (has children)
        for (let j = 1; j < assets.length; j++) {
          if (assets[j][2] === locationId && shipTypeIds.has(assets[j][7])) {
            fittedShips.add(locationId);
            break;
          }
        }
      }
      Logger.log(`Found ${fittedShips.size} fitted ships`);
      
      // Build a map of item_id -> asset for quick lookup
      const assetMap = {};
      for (let i = 1; i < assets.length; i++) {
        const asset = assets[i];
        const itemId = asset[2]; // item_id is column 2
        assetMap[itemId] = asset;
      }
      
      // Process each asset
      for (let i = 1; i < assets.length; i++) {
        const asset = assets[i];
        const locationFlag = asset[3];
        const locationId = asset[4];
        const quantity = asset[6];
        const typeId = asset[7];
        
        // Filter by location flags and skip fitted items
        if (!locationFlags.includes(locationFlag) || skipFlags.includes(locationFlag)) {
          continue;
        }

        // Skip fitted ships (Hangar ships that have items fitted inside)
        if (shipTypeIds.has(typeId) && locationFlag === 'Hangar' && fittedShips.has(locationId)) {
          Logger.log(`Skipping fitted ship ${getItemNameFromTypeId(typeId)} at ${locationId}`);
          continue;
        }
        
        totalItems++;
        
        // Resolve location chain: if locationId points to a container/ship, follow the chain
        let finalLocationId = locationId;
        
        // Check if this location_id is actually a container/ship (item in assets)
        if (assetMap[locationId]) {
          let currentId = locationId;
          // Follow the chain up until we reach a station/structure (not in assetMap)
          while (assetMap[currentId]) {
            const containerAsset = assetMap[currentId];
            currentId = containerAsset[4]; // Get the container's location_id
            if (!assetMap[currentId]) {
              // Found the top-level location (station/structure)
              finalLocationId = currentId;
              break;
            }
          }
        }
        
        // Look up location only once per unique locationId
        if (!locationCache[finalLocationId]) {
          Logger.log(`Looking up location ${finalLocationId} (item in flag: ${locationFlag})...`);
          locationCache[finalLocationId] = getSystemForLocation(finalLocationId);
          uniqueLocations++;
        }
        
        const systemInfo = locationCache[finalLocationId];
        const itemName = getItemNameFromTypeId(typeId);
        
        rows.push([
          charName,
          typeId,
          itemName,
          quantity,
          finalLocationId,
          locationFlag,
          systemInfo.locationType || '',
          systemInfo.systemId || '',
          systemInfo.systemName || '[Unknown]'
        ]);
      }
    }
    
    Logger.log(`Total items: ${totalItems}, Unique locations: ${uniqueLocations}`);
    Logger.log(`DEBUG: rows.length = ${rows.length}`);
    if (rows.length > 0) {
      Logger.log(`DEBUG: rows[0] = ${JSON.stringify(rows[0])}`);
    }
    
    if (rows.length === 0) {
      Logger.log('No items found in selected locations.');
      return;
    }
    
    // Clear previous data
    Logger.log(`Clearing S2:AA2000...`);
    fitData.getRange('S2:AA2000').clearContent();
    
    // Write headers
    const headers = [['Character', 'TypeID', 'ItemName', 'Qty', 'LocationID', 'LocationFlag', 'LocationType', 'SystemID', 'SystemName']];
    fitData.getRange('S2:AA2').setValues(headers);
    Logger.log('Headers written');
    
    Logger.log(`Writing ${rows.length} rows to sheet in batches...`);
    
    // Batch writes in chunks of 500
    const batchSize = 500;
    for (let i = 0; i < rows.length; i += batchSize) {
      const batch = rows.slice(i, i + batchSize);
      const startRow = 3 + i;
      const endRow = startRow + batch.length - 1;
      const rangeRef = `S${startRow}:AA${endRow}`;
      Logger.log(`Writing batch: ${rangeRef} (${batch.length} rows)`);
      
      fitData.getRange(rangeRef).setValues(batch);
      Logger.log(`✓ Wrote ${rangeRef}`);
    }
    
    Logger.log(`Complete. ${rows.length} items from ${uniqueLocations} unique locations.`);
    
  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    Logger.log(`Error: ${error.message}`);
  }
}

/**
 * POPULATE LOCATION FILTER
 * 
 * After pulling inventory, call this to populate FitCompare location filter (M3:N50)
 * with deduplicated, alphabetically sorted unique station/structure names
 * 
 * Reads from inventory data (S2:AA2000), extracts column AA (SystemName which contains station/structure names),
 * deduplicates, sorts, and writes to M3:M50
 * Also sets all checkboxes (N3:N50) to TRUE by default
 */
function populateLocationFilter() {
  try {
    Logger.log('=== POPULATE LOCATION FILTER START ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fitCompare = ss.getSheetByName('FitCompare');
    const fitData = ss.getSheetByName('FitData');
    
    // Read from SystemCache (AB2:AD500) column AD = SystemName
    Logger.log('Reading system cache...');
    const cacheRange = fitData.getRange('AD3:AD500');
    const cacheData = cacheRange.getValues();
    
    // Get unique locations
    const uniqueLocations = [];
    const seen = new Set();
    
    for (const row of cacheData) {
      if (row && row[0] && typeof row[0] === 'string') {
        const location = row[0].trim();
        if (location.length > 0 && !location.includes('[Unknown') && !seen.has(location)) {
          uniqueLocations.push(location);
          seen.add(location);
        }
      }
    }
    
    uniqueLocations.sort();
    
    if (uniqueLocations.length === 0) {
      Logger.log('⚠ No valid locations found');
      return;
    }
    
    // Write to FitCompare M3:N50
    fitCompare.getRange('M3:N50').clearContent();
    
    const filterData = [];
    for (const location of uniqueLocations) {
      filterData.push([location, true]);
    }
    
    const rangeRef = `M3:N${2 + filterData.length}`;
    fitCompare.getRange(rangeRef).setValues(filterData);
    
    Logger.log(`✓ Populated with ${filterData.length} locations`);
    
  } catch (error) {
    Logger.log(`❌ ERROR: ${error.message}`);
    throw error;
  }
}

/**
 * RUN FIT COMPARISON WORKFLOW
 * 
 * Main orchestration function for the complete fit comparison process:
 * 1. Read target fit from FitCompare sheet (A3:H55)
 * 2. Read current fit from FitCompare sheet (A60:H112) - optional
 * 3. Parse both fits using EFT parser
 * 4. Compare to determine what's needed
 * 5. Read selected systems from location filter (M3:N50)
 * 6. Read inventory data from FitData (S2:AA2000)
 * 7. Calculate gaps for selected systems only
 * 8. Populate buy list (A116:B250)
 * 
 * Call this via "Sheet Tools" menu after:
 * - Pasting target fit in A3:H55
 * - (Optional) Pasting current fit in A60:H112
 * - Running pullFitInventory() to populate inventory data
 * - Checking desired systems in location filter (M3:N50)
 */
function runFitComparison() {
  try {
    Logger.log('\n' + '='.repeat(80));
    Logger.log('STARTING FIT COMPARISON WORKFLOW');
    Logger.log('='.repeat(80) + '\n');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fitCompare = ss.getSheetByName('FitCompare');
    const fitData = ss.getSheetByName('FitData');
    
    if (!fitCompare || !fitData) {
      throw new Error('Required sheets not found: FitCompare and/or FitData');
    }
    
    // STEP 1: Read target fit from FitCompare (A3:H55)
    Logger.log('STEP 1: Reading target fit (A3:H55)...');
    const targetFitText = readFitFromSheet(fitCompare, 3, 55);
    if (!targetFitText) {
      throw new Error('No target fit found in A3:H55. Paste EFT format fit.');
    }
    Logger.log(`✓ Target fit read (${targetFitText.length} chars)`);
    
    // STEP 2: Read current fit from FitCompare (A60:H112)
    Logger.log('\nSTEP 2: Reading current fit (A60:H112, optional)...');
   const currentFitText = readFitFromSheet(fitCompare, 60, 112);
    if (currentFitText) {
      Logger.log(`✓ Current fit read (${currentFitText.length} chars)`);
    } else {
      Logger.log('ℹ No current fit found (will treat as empty)');
    }
    
    // STEP 3: Parse both fits
    Logger.log('\nSTEP 3: Parsing fits...');
    const targetFit = parseEFT(targetFitText);
    if (!targetFit.ship) {
      throw new Error('Failed to parse target fit. Check EFT format: [ShipType, FitName]');
    }
    Logger.log(`✓ Target fit parsed: ${targetFit.ship} (${targetFit.items.length} items)`);
    
    const currentFit = currentFitText ? parseEFT(currentFitText) : { ship: null, fitName: null, items: [] };
    Logger.log(`✓ Current fit parsed: ${currentFit.ship || '[empty]'} (${currentFit.items.length} items)`);
    
// After parsing fits, log removals/additions
const transitionList = buildFitTransitionList(targetFit, currentFit);
logRemovalAdditions(fitCompare, transitionList);

    // STEP 4: Compare fits
    Logger.log('\nSTEP 4: Comparing fits...');
    const neededItems = compareFits(targetFit, currentFit);
    Logger.log(`✓ ${neededItems.length} items needed`);
    
    // STEP 5: Get FitQuantity multiplier
    Logger.log('\nSTEP 5: Reading FitQuantity (L18)...');
    const fitQuantity = getFitQuantity(fitCompare);
    
    // STEP 6: Get selected systems from location filter
    Logger.log('\nSTEP 6: Reading location filter (M3:N50)...');
    const selectedSystems = getSelectedSystems(fitCompare);
    if (selectedSystems.length === 0) {
      Logger.log('⚠ WARNING: No systems selected in location filter. Buy list will be empty.');
    } else {
      Logger.log(`✓ ${selectedSystems.length} systems selected: ${selectedSystems.join(', ')}`);
    }
    
// STEP 7: Reading inventory data (S2:AA2000)...
Logger.log('\nSTEP 7: Reading inventory data (S2:AA2000)...');
const invRange = fitData.getRange('S2:AA2000');
const inventory = invRange.getValues().filter(row => row[0]); // Filter out empty rows
Logger.log(`✓ Inventory: ${inventory.length} items across all locations`);

// STEP 8: Calculating gaps with FitQuantity multiplier...
Logger.log('\nSTEP 8: Calculating gaps with FitQuantity multiplier...');
const gapResult = calculateGapWithSources(
  neededItems,
  inventory,
  selectedSystems,
  fitQuantity
);

Logger.log(`DEBUG: gapResult = ${JSON.stringify(gapResult)}`);
Logger.log(`DEBUG: gapResult.gap = ${JSON.stringify(gapResult.gap)}`);
Logger.log(`DEBUG: gapResult.gap type = ${typeof gapResult.gap}`);

const gapAnalysis = gapResult.gap;
const sources = gapResult.sources;

// STEP 9: Populate buy list
Logger.log('\nSTEP 9: Populating buy list (A116:B250)...');
populateBuyList(fitCompare, gapAnalysis);

// STEP 10: Populate inventory sources
Logger.log('\nSTEP 10: Populating inventory sources (D116:G250)...');
populateInventorySources(fitCompare, gapAnalysis, sources);
    
    // SUMMARY
    const itemsToBuy = gapAnalysis.filter(item => item.qtyToBuy > 0);
    const totalUnits = itemsToBuy.reduce((sum, item) => sum + item.qtyToBuy, 0);
    
    Logger.log('\n' + '='.repeat(80));
    Logger.log('FIT COMPARISON COMPLETE');
    Logger.log('='.repeat(80));
    Logger.log(`Target Ship: ${targetFit.ship}`);
    Logger.log(`Fit Quantity: ${fitQuantity}x`);
    Logger.log(`Items Needed: ${neededItems.length}`);
    Logger.log(`Items to Buy: ${itemsToBuy.length}`);
    Logger.log(`Total Units to Buy: ${totalUnits}`);
    Logger.log('='.repeat(80) + '\n');
    
  } catch (error) {
    Logger.log(`\n❌ ERROR: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    throw error;
  }
}

/**
 * Read fit text from a range on the sheet
 * Concatenates all non-empty cells in the range into a single string
 * 
 * @param {Sheet} sheet - The sheet to read from
 * @param {string} rangeRef - Range reference like 'A3:H55'
 * @returns {string} Concatenated fit text, or empty string if range is empty
 */
function readFitFromSheet(sheet, startRow, endRow) {
  const range = sheet.getRange(startRow, 1, (endRow - startRow + 1), 1);
  const values = range.getValues();
  
  let fitText = '';
  for (const row of values) {
    const cell = row[0];
    if (cell && cell.toString().trim()) {
      fitText += cell + '\n';
    }
  }
  
  return fitText;
}

/**
 * Get list of selected systems from location filter
 * Reads M3:N50 where column M = system names, column N = checkboxes (TRUE/FALSE)
 * Supports up to 48 unique systems
 * 
 * @param {Sheet} sheet - FitCompare sheet
 * @returns {array} Array of system names that are checked
 */
function getSelectedSystems(sheet) {
  const range = sheet.getRange('M3:N50');
  const values = range.getValues();
  
  const selected = [];
  for (let row = 0; row < values.length; row++) {
    const systemName = values[row][0];
    const isChecked = values[row][1];
    
    // Only include if system name exists and checkbox is TRUE
    if (systemName && typeof systemName === 'string' && systemName.trim().length > 0 && isChecked === true) {
      selected.push(systemName.trim());
    }
  }
  
  return selected;
}

/**
 * Populate buy list in FitCompare sheet (A116:B250)
 * Writes item names and quantities to buy
 * Only includes items with qtyToBuy > 0 (sorted by qty descending)
 * 
 * @param {Sheet} sheet - FitCompare sheet
 * @param {array} gapAnalysis - Array from calculateGap()
 */
function populateBuyList(sheet, gapAnalysis) {
  // Clear column B (remove ADD/REMOVE markers)
  sheet.getRange('B117:B250').clearContent();
  
  // Clear existing buy list data
  sheet.getRange('A117:B250').clearContent();
  
  // Build output array: only items with qtyToBuy > 0
  const buyItems = gapAnalysis.filter(item => item.qtyToBuy > 0);
  
  if (buyItems.length === 0) {
    Logger.log('No items to buy.');
    return;
  }
  
  const output = [];
  for (const item of buyItems) {
    output.push([
      item.name,
      item.qtyToBuy
    ]);
  }
  
  const rangeRef = `A117:B${116 + output.length}`;
  Logger.log(`Writing buy list to ${rangeRef} (${output.length} items)`);
  
  sheet.getRange(rangeRef).setValues(output);
}

function logRemovalAdditions(sheet, transitionList) {
  Logger.log('\n=== REMOVAL/ADDITION LIST ===');
  
  // Log removals
  Logger.log('REMOVE from current fit:');
  for (const item of transitionList.removals) {
    Logger.log(`  ✗ ${item.name} (qty: ${item.qty})`);
  }
  
  // Log additions
  Logger.log('ADD to target fit:');
  for (const item of transitionList.additions.concat(transitionList.replacements)) {
    Logger.log(`  + ${item.name} (qty: ${item.qty})`);
  }
}

function clearLocationFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fitCompare = ss.getSheetByName('FitCompare');
  
  const range = fitCompare.getRange('N3:N50');
  range.clearContent();
  
  Logger.log('✓ Location filter cleared');
}

function populateInventorySources(sheet, gapAnalysis, sources) {
  Logger.log(`DEBUG: gapAnalysis.length = ${gapAnalysis.length}`);
  Logger.log(`DEBUG: sources keys = ${Object.keys(sources)}`);
  
  const output = [];
  
  for (const item of gapAnalysis) {
    const hasSource = sources[item.name] ? true : false;
    Logger.log(`Item: ${item.name}, qtyToBuy: ${item.qtyToBuy}, hasSource: ${hasSource}`);
    
    if (item.qtyToBuy > 0 && sources[item.name]) {
      for (const source of sources[item.name]) {
        output.push([
          item.name,
          source.qty,
          source.character,
          source.location
        ]);
      }
    }
  }
  
  sheet.getRange('D115:G250').clearContent();
  sheet.getRange('D116:D116').setValue('Item');
  sheet.getRange('E116:E116').setValue('Qty');
  sheet.getRange('F116:F116').setValue('Character');
  sheet.getRange('G116:G116').setValue('Location');
  
  if (output.length > 0) {
    sheet.getRange(`D117:G${116 + output.length}`).setValues(output);
    Logger.log(`✓ Populated ${output.length} inventory source rows`);
  }
}

function populateRemovalAdditions() {
  try {
    Logger.log('=== POPULATE REMOVAL/ADDITIONS ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fitCompare = ss.getSheetByName('FitCompare');
    const fitData = ss.getSheetByName('FitData');
    
    // Read target fit (A3:H55)
    const targetRange = fitCompare.getRange('A3:H55');
    const targetData = targetRange.getValues();
    
    // Read current fit (A60:H112)
    const currentRange = fitCompare.getRange('A60:H112');
    const currentData = currentRange.getValues();
    
    // Get item names from target
    const targetItems = [];
    for (let row = 0; row < targetData.length; row++) {
      for (let col = 0; col < targetData[row].length; col++) {
        const cell = targetData[row][col];
        if (cell && typeof cell === 'string' && cell.trim().length > 0 && !cell.includes('[')) {
          targetItems.push({ name: cell.trim().toLowerCase(), row: row + 3 });
        }
      }
    }
    
    // Get item names from current
    const currentItems = [];
    for (let row = 0; row < currentData.length; row++) {
      for (let col = 0; col < currentData[row].length; col++) {
        const cell = currentData[row][col];
        if (cell && typeof cell === 'string' && cell.trim().length > 0 && !cell.includes('[')) {
          currentItems.push({ name: cell.trim().toLowerCase(), row: row + 60 });
        }
      }
    }
    
    // Mark REMOVE in current fit (items not in target)
    for (const item of currentItems) {
      const inTarget = targetItems.find(t => t.name === item.name);
      if (!inTarget) {
        fitCompare.getRange(`B${item.row}`).setValue('REMOVE');
        Logger.log(`REMOVE: ${item.name}`);
      }
    }
    
    // Mark ADD in target fit (items not in current)
    for (const item of targetItems) {
      const inCurrent = currentItems.find(c => c.name === item.name);
      if (!inCurrent) {
        fitCompare.getRange(`B${item.row}`).setValue('ADD');
        Logger.log(`ADD: ${item.name}`);
      }
    }
    
    Logger.log('✓ Marked removals and additions');
    
  } catch (error) {
    Logger.log(`❌ ERROR: ${error.message}`);
    throw error;
  }
}