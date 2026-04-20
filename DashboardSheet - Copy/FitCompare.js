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
    if (cacheData[i][0] == locationId) {
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
  
  try {
    if (locationId >= 60000000 && locationId <= 69999999) {
      // Station
      const stationData = GESI.universe_stations_station(locationId, true);
      if (stationData && stationData.length > 1) {
        systemId = stationData[1][3];
        systemName = stationData[1][1];
        locationType = 'station';
      }
    } else if (locationId >= 100000000) {
      // Structure
      const structData = GESI.universe_structures_structure(locationId, '', true);
      if (structData && structData.length > 1) {
        systemId = structData[1][3];  // solar_system_id (correct)
        systemName = structData[1][0]; // name is at [1][0], not [1][1]
        locationType = 'structure';
      }
    }
  } catch (e) {
    Logger.log(`Location lookup failed for ${locationId}: ${e.message}`);
    systemName = `[Unknown ${locationId}]`;
  }
  
  // Cache it
  const lastRow = cacheData.filter(row => row[0]).length + 2;
  fitData.getRange(`AB${lastRow}:AD${lastRow}`).setValues([[locationId, systemId, systemName]]);
  
  return {
    locationId: locationId,
    systemId: systemId,
    systemName: systemName,
    locationType: locationType
  };
}

function pullFitInventory() {
  loadSdeCache();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fitCompare = ss.getSheetByName('FitCompare');
  const fitData = ss.getSheetByName('FitData');
  
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
    SpreadsheetApp.getUi().alert('No characters selected.');
    return;
  }
  
  // Setup
  const locationCache = {};
  const rows = [];
  const locationFlags = ['Unlocked', 'Locked', 'Hangar', 'Deliveries', 'Cargo', 'CapsuleerDeliveries', 'InfrastructureHangar', 'FleetHangar'];
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
      SpreadsheetApp.getUi().alert('No items found in selected locations.');
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
    SpreadsheetApp.getUi().alert(`Loaded ${rows.length} items from ${uniqueLocations} locations.`);
    
  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}