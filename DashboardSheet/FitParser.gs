/**
 * EVE ONLINE EFT FIT PARSER
 * 
 * Parses EVE Fitting Tool (EFT) format text into structured data.
 * EFT format is the standard clipboard format used in EVE Online.
 * 
 * See: https://wiki.eveuniversity.org/Fitting_Format
 */

/**
 * Parse EFT format text and extract all fitted items
 * 
 * @param {string} eftText - Raw EFT format text from clipboard
 * @returns {object} Structured fit data with ship and items array
 * 
 * @example
 * const fit = parseEFT(eftText);
 * // Returns: {
 * //   ship: "Heron Navy Issue",
 * //   fitName: "Deepflow Rift Dredger",
 * //   items: [
 * //     {name: "Inertial Stabilizers II", qty: 1, type: "low"},
 * //     {name: "Antimatter Charge S", qty: 8, type: "cargo"},
 * //   ]
 * // }
 */
function parseEFT(eftText) {
  Logger.log('=== EFT PARSER START ===');
  
  if (!eftText || typeof eftText !== 'string') {
    Logger.log('ERROR: Invalid EFT text input');
    return { ship: null, fitName: null, items: [] };
  }
  
  // Split into lines and trim
  const lines = eftText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
  
  if (lines.length === 0) {
    Logger.log('ERROR: Empty EFT text');
    return { ship: null, fitName: null, items: [] };
  }
  
  // Parse header: [Ship, Fit Name]
  const headerMatch = lines[0].match(/^\[([^\]]+),\s*(.+)\]$/);
  if (!headerMatch) {
    Logger.log(`ERROR: Invalid header format: ${lines[0]}`);
    return { ship: null, fitName: null, items: [] };
  }
  
  const ship = headerMatch[1].trim();
  const fitName = headerMatch[2].trim();
  Logger.log(`✓ Parsed ship: ${ship}, fit: ${fitName}`);
  
  const items = [];
  
  // Process each line after header
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    
    // Empty line signals section change - skip
    if (line === '') {
      continue;
    }
    
    // Skip empty slot markers
    if (line.includes('[Empty')) {
      continue;
    }
    
    // Parse item line
    const parsed = parseItemLine(line);
    
    if (parsed) {
      items.push(parsed);
      Logger.log(`  ✓ ${parsed.name} (qty: ${parsed.qty}, type: ${parsed.type})`);
    }
  }
  
  Logger.log(`✓ Total items parsed: ${items.length}`);
  Logger.log('=== EFT PARSER END ===\n');
  
  return { ship, fitName, items };
}

/**
 * Parse a single item line from EFT format
 * 
 * Handles formats like:
 * - "125mm Railgun I" (module)
 * - "125mm Railgun I /offline" (offline module - suffix ignored)
 * - "Antimatter Charge S x42" (cargo with quantity)
 * - "Warrior II x2" (drones with quantity)
 * 
 * @param {string} line - Single item line
 * @returns {object|null} Parsed item {name, qty, type} or null if invalid
 */
function parseItemLine(line) {
  // Remove /offline suffix (in-game ignores it anyway)
  line = line.replace(/\s*\/offline\s*$/i, '').trim();
  
  // Check for quantity suffix "x42"
  const quantityMatch = line.match(/^(.+?)\s+x(\d+)\s*$/);
  let itemName = line;
  let qty = 1;
  let type = 'module';
  
  if (quantityMatch) {
    itemName = quantityMatch[1].trim();
    qty = parseInt(quantityMatch[2], 10);
    
    // Items with quantity suffix are typically drones or cargo
    // Drones: Warrior, Hobgoblin, etc. (contain common drone keywords)
    // Otherwise: cargo
    if (itemName.match(/warrior|hobgoblin|valkyrie|infiltrator|antimatter|iridium|thorium/i)) {
      type = 'drone_or_charge';
    } else {
      type = 'cargo';
    }
  } else {
    type = 'module';
  }
  
  // Validate item name
  if (!itemName || itemName.length === 0 || itemName.match(/^\[/)) {
    return null;
  }
  
  return {
    name: itemName,
    qty: qty,
    type: type
  };
}

/**
 * Compare target fit against current fit
 * Returns list of items that differ (need to be replaced or added)
 * 
 * @param {object} targetFit - Parsed target fit from parseEFT()
 * @param {object} currentFit - Parsed current fit from parseEFT() (can be empty)
 * @returns {array} Array of items that need to be purchased
 * 
 * @example
 * const needed = compareFits(targetFit, currentFit);
 * // Returns: [
 * //   {name: "Heron Navy Issue", qty: 1, reason: "ship"},
 * //   {name: "Inertial Stabilizers II", qty: 1, reason: "new"},
 * // ]
 */
function compareFits(targetFit, currentFit) {
  Logger.log('=== COMPARE FITS START ===');
  
  const needed = [];
  
  // Always include target ship
  if (targetFit.ship) {
    needed.push({
      name: targetFit.ship,
      qty: 1,
      reason: 'ship'
    });
    Logger.log(`✓ Added ship: ${targetFit.ship}`);
  }
  
  // Build current fit inventory map {itemName: qty}
  const currentInventory = {};
  if (currentFit && currentFit.items) {
    for (const item of currentFit.items) {
      const key = item.name.toLowerCase();
      currentInventory[key] = (currentInventory[key] || 0) + item.qty;
    }
  }
  
  Logger.log(`Current fit has ${Object.keys(currentInventory).length} unique item types`);
  
  // Compare each target item against current
  for (const targetItem of targetFit.items) {
    const key = targetItem.name.toLowerCase();
    const currentQty = currentInventory[key] || 0;
    const neededQty = targetItem.qty - currentQty;
    
    if (neededQty > 0) {
      needed.push({
        name: targetItem.name,
        qty: neededQty,
        reason: currentQty > 0 ? 'replacement' : 'new'
      });
      Logger.log(`✓ ${targetItem.name}: need ${neededQty} (have ${currentQty})`);
    } else if (neededQty === 0) {
      Logger.log(`✓ ${targetItem.name}: complete (have ${currentQty})`);
    }
  }
  
  Logger.log(`Total unique items needed: ${needed.length}`);
  Logger.log('=== COMPARE FITS END ===\n');
  
  return needed;
}

/**
 * Calculate gap between needed items and inventory (per system filter)
 * 
 * Filters inventory to only count items in selected systems.
 * Items are only counted once per location (not duplicated across regions).
 * 
 * @param {array} neededItems - Array from compareFits()
 * @param {array} inventory - Data from InventoryPull range (columns S-AA)
 *                           Columns: S=Character, T=TypeID, U=ItemName, V=Qty, W=LocationID, X=LocationFlag, Y=LocationType, Z=SystemID, AA=SystemName
 * @param {array} selectedSystems - Array of system names that are checked in location filter
 * @returns {array} Array of {name, qtyNeeded, qtyOnHand, qtyToBuy, reason}
 */
function calculateGap(neededItems, inventory, selectedSystems) {
  Logger.log('=== CALCULATE GAP START ===');
  Logger.log(`Needed items: ${neededItems.length}, Inventory rows: ${inventory.length}`);
  Logger.log(`Selected systems: ${selectedSystems.join(', ')}`);
  
  const gap = [];
  
  // Build inventory map {itemName: qtyOnHand} filtered by selected systems
  const invMap = {};
  
  for (const row of inventory) {
    if (!row || row.length === 0) continue;
    
    // InventoryPull columns: [Character, TypeID, ItemName, Qty, LocationID, LocationFlag, LocationType, SystemID, SystemName]
    //                        [0,         1,       2,        3,   4,          5,            6,            7,        8]
    const itemName = row[2];      // ItemName
    const qty = row[3];           // Qty
    const systemName = row[8];    // SystemName
    
    // Only count if system is selected
    if (systemName && selectedSystems.includes(systemName)) {
      invMap[itemName] = (invMap[itemName] || 0) + qty;
    }
  }
  
  Logger.log(`Inventory items in selected systems: ${Object.keys(invMap).length}`);
  
  // Calculate gap for each needed item
  for (const item of neededItems) {
    const qtyOnHand = invMap[item.name] || 0;
    const qtyToBuy = Math.max(0, item.qty - qtyOnHand);
    
    gap.push({
      name: item.name,
      qtyNeeded: item.qty,
      qtyOnHand: qtyOnHand,
      qtyToBuy: qtyToBuy,
      reason: item.reason
    });
    
    Logger.log(`${item.name}: need ${item.qty}, have ${qtyOnHand}, buy ${qtyToBuy}`);
  }
  
  // Sort: items to buy first (by qty descending), then by name
  gap.sort((a, b) => {
    if ((a.qtyToBuy > 0) !== (b.qtyToBuy > 0)) {
      return b.qtyToBuy - a.qtyToBuy;
    }
    if (a.qtyToBuy !== b.qtyToBuy) {
      return b.qtyToBuy - a.qtyToBuy;
    }
    return a.name.localeCompare(b.name);
  });
  
  Logger.log(`Total gap rows: ${gap.length}`);
  Logger.log('=== CALCULATE GAP END ===\n');
  
  return gap;
}