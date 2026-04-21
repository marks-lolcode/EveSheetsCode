/**
 * EVE ONLINE FIT COMPARISON - ENHANCED VERSION
 * 
 * Extended workflow:
 * 1. Parse target and current fits
 * 2. Compare to find: removals, additions, replacements
 * 3. Calculate gaps for selected systems with FitQuantity multiplier
 * 4. Track inventory sources (character + location) for each item
 * 5. Output removal list, needed items, and where to find them
 */

/**
 * Build a removal/addition list when comparing fits
 * Shows what needs to be removed from current fit and what needs to be added
 * 
 * @param {object} targetFit - Parsed target fit
 * @param {object} currentFit - Parsed current fit
 * @returns {object} {removals: [], additions: [], replacements: []}
 */
function buildFitTransitionList(targetFit, currentFit) {
  Logger.log('=== BUILD FIT TRANSITION LIST ===');
  
  const result = {
    removals: [],      // Items in current but not in target (or different qty)
    additions: [],     // Items in target but not in current (or different qty)
    replacements: []   // Items where current qty < target qty
  };
  
  // Build maps: {itemName: qty}
  const targetMap = {};
  const currentMap = {};
  
  for (const item of targetFit.items) {
    const key = item.name.toLowerCase();
    targetMap[key] = (targetMap[key] || 0) + item.qty;
  }
  
  for (const item of currentFit.items) {
    const key = item.name.toLowerCase();
    currentMap[key] = (currentMap[key] || 0) + item.qty;
  }
  
  // Find items only in current (removals)
  for (const [itemKey, currentQty] of Object.entries(currentMap)) {
    const targetQty = targetMap[itemKey] || 0;
    if (targetQty === 0) {
      // Item completely removed
      result.removals.push({
        name: findOriginalName(itemKey, currentFit.items),
        qty: currentQty,
        reason: 'completely_removed'
      });
    } else if (currentQty > targetQty) {
      // Item quantity reduced
      result.removals.push({
        name: findOriginalName(itemKey, currentFit.items),
        qty: currentQty - targetQty,
        reason: 'qty_reduced'
      });
    }
  }
  
  // Find items only in target or increased qty (additions)
  for (const [itemKey, targetQty] of Object.entries(targetMap)) {
    const currentQty = currentMap[itemKey] || 0;
    if (currentQty === 0) {
      // New item
      result.additions.push({
        name: findOriginalName(itemKey, targetFit.items),
        qty: targetQty,
        reason: 'new_item'
      });
    } else if (targetQty > currentQty) {
      // Item quantity increased
      result.replacements.push({
        name: findOriginalName(itemKey, targetFit.items),
        qty: targetQty - currentQty,
        reason: 'qty_increased'
      });
    }
  }
  
  Logger.log(`Removals: ${result.removals.length}, Additions: ${result.additions.length}, Replacements: ${result.replacements.length}`);
  Logger.log('=== END FIT TRANSITION LIST ===\n');
  
  return result;
}

/**
 * Helper: Find original item name from items array by lowercase key
 */
function findOriginalName(key, itemsArray) {
  for (const item of itemsArray) {
    if (item.name.toLowerCase() === key) {
      return item.name;
    }
  }
  return `[Unknown: ${key}]`;
}

/**
 * Calculate gap with FitQuantity multiplier and track inventory sources
 * 
 * @param {array} neededItems - Array from compareFits()
 * @param {array} inventory - InventoryPull data (S2:AA2000)
 * @param {array} selectedSystems - Selected systems from filter
 * @param {number} fitQuantity - Number of fits to build (from FitQuantity range)
 * @returns {object} {gap: [], sources: {}}
 *   gap = array of {name, qtyNeeded, qtyOnHand, qtyToBuy}
 *   sources = {itemName: [{character, location, qty}, ...]}
 */
function calculateGapWithSources(neededItems, inventory, selectedSystems, fitQuantity) {
  Logger.log('=== CALCULATE GAP WITH SOURCES START ===');
  Logger.log(`Needed items: ${neededItems.length}, Fit Quantity: ${fitQuantity}`);
  
  if (!fitQuantity || fitQuantity < 1) {
    fitQuantity = 1;
    Logger.log('ℹ FitQuantity invalid, defaulting to 1');
  }
  
  let gap = [];
 let sources = {};
  
  // Build inventory map with sources: {itemName: [{character, location, qty}, ...]}
  const invMap = {};
  
  for (const row of inventory) {
    if (!row || row.length < 9) continue;
    
    // [Character, TypeID, ItemName, Qty, LocationID, LocationFlag, LocationType, SystemID, SystemName]
    const character = row[0];
    const itemName = row[2];
    const qty = row[3];
    const location = row[8];  // SystemName (station/structure name)
    const systemName = row[8];
    
    // Only count if system is selected and location is not "[Unknown]"
    if (systemName && selectedSystems.includes(systemName) && !systemName.includes('[Unknown')) {
      if (!invMap[itemName]) {
        invMap[itemName] = [];
      }
      invMap[itemName].push({
        character: character,
        location: location,
        qty: qty
      });
    }
  }
  
  Logger.log(`Inventory items in selected systems: ${Object.keys(invMap).length}`);
  
  // Calculate gaps with multiplier
  for (const item of neededItems) {
    const qtyNeeded = item.qty * fitQuantity;  // Multiply by number of fits
    
    // Sum available quantity from all sources
    let qtyOnHand = 0;
    if (invMap[item.name]) {
      qtyOnHand = invMap[item.name].reduce((sum, source) => sum + source.qty, 0);
      // Store sources for this item
      sources[item.name] = invMap[item.name];
    }
    
    const qtyToBuy = Math.max(0, qtyNeeded - qtyOnHand);
    
    gap.push({
      name: item.name,
      qtyNeeded: qtyNeeded,
      qtyOnHand: qtyOnHand,
      qtyToBuy: qtyToBuy,
      reason: item.reason
    });
    
    Logger.log(`${item.name}: need ${qtyNeeded} (x${fitQuantity}), have ${qtyOnHand}, buy ${qtyToBuy}`);
  }
  
  // Sort
  gap.sort((a, b) => {
    if ((a.qtyToBuy > 0) !== (b.qtyToBuy > 0)) return b.qtyToBuy - a.qtyToBuy;
    if (a.qtyToBuy !== b.qtyToBuy) return b.qtyToBuy - a.qtyToBuy;
    return a.name.localeCompare(b.name);
  });
  
  Logger.log(`Gap rows: ${gap.length}`);
  Logger.log('=== CALCULATE GAP WITH SOURCES END ===\n');
  
  return { gap, sources };
}

/**
 * Get FitQuantity value from Constants sheet or default
 * Reads named range "FitQuantity" which should be in FitCompare L18
 * 
 * @param {Sheet} sheet - FitCompare sheet
 * @returns {number} Number of fits (default 1)
 */
function getFitQuantity(sheet) {
  try {
    const quantityRange = sheet.getRange('L18');
    const value = quantityRange.getValue();
    const qty = parseInt(value, 10);
    
    if (isNaN(qty) || qty < 1) {
      Logger.log(`ℹ FitQuantity invalid (${value}), defaulting to 1`);
      return 1;
    }
    
    Logger.log(`✓ FitQuantity: ${qty}`);
    return qty;
  } catch (e) {
    Logger.log(`⚠ Could not read FitQuantity: ${e.message}, defaulting to 1`);
    return 1;
  }
}

/**
 * Get unique, deduplicated, alphabetically sorted locations from inventory
 * Properly excludes header row and [Unknown] locations
 * 
 * @param {array} inventory - InventoryPull data
 * @returns {array} Sorted unique location names
 */
function getUniqueLocations(inventory) {
  Logger.log('=== GET UNIQUE LOCATIONS ===');
  
  const locationSet = new Set();
  let processed = 0;
  let skipped = 0;
  
  for (const row of inventory) {
    if (!row || row.length < 9) {
      skipped++;
      continue;
    }
    
    const locationName = row[8];  // SystemName column
    
    // Skip empty, header row, and [Unknown] locations
    if (!locationName || typeof locationName !== 'string' || locationName.length === 0) {
      skipped++;
      continue;
    }
    
    const trimmed = locationName.trim();
    
    // Skip header row label
    if (trimmed === 'SystemName') {
      Logger.log('ℹ Skipped header row (SystemName)');
      skipped++;
      continue;
    }
    
    // Skip [Unknown] locations
    if (trimmed.includes('[Unknown')) {
      skipped++;
      continue;
    }
    
    locationSet.add(trimmed);
    processed++;
  }
  
  const uniqueLocations = Array.from(locationSet).sort();
  
  Logger.log(`Processed: ${processed}, Skipped: ${skipped}, Unique: ${uniqueLocations.length}`);
  for (let i = 0; i < Math.min(3, uniqueLocations.length); i++) {
    Logger.log(`  ${i+1}. ${uniqueLocations[i]}`);
  }
  Logger.log('=== END GET UNIQUE LOCATIONS ===\n');
  
  return uniqueLocations;
}
