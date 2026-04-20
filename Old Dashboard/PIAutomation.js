function calculateProfit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const sheet = ss.getSheetByName("ProfitCalculation") || ss.insertSheet("ProfitCalculation");
  const errorSheet = ss.getSheetByName("ErrorLog") || ss.insertSheet("ErrorLog");

  // Clear previous data
  sheet.clear();
  sheet.appendRow(["Item", "Percentage Profit", "Profit", "Output Quantity", "Input Quantity"]);

  errorSheet.clear();
  errorSheet.appendRow(["Timestamp", "Item", "Error Message"]);

  try {
    // Load data from named ranges
    const tierList = getNamedRangeData("PITierList");
    const tiers = getNamedRangeData("PITiers");
    const requirements = getNamedRangeData("PIRequirements");
    const prices = getNamedRangeData("PIPrices");

    // Debugging: Log retrieved data
    console.log("tierList:", JSON.stringify(tierList));
    console.log("tiers:", JSON.stringify(tiers));
    console.log("requirements:", JSON.stringify(requirements));
    console.log("prices:", JSON.stringify(prices));

    // Create lookup objects
    const priceMap = createLookup(prices, 0, 1);
    const tierMap = createLookup(tierList, 0, 1);
    const tierDataMap = createLookup(tiers, 0, [1, 2]); // Maps Tier -> [Component Multiplier, Output Quantity]
    const requirementMap = createLookupMultiple(requirements, 0); // Maps Item -> List of Components

    console.log("priceMap:", JSON.stringify(priceMap));
    console.log("tierMap:", JSON.stringify(tierMap));
    console.log("tierDataMap:", JSON.stringify(tierDataMap));
    console.log("requirementMap:", JSON.stringify(requirementMap));

    let results = [];

    // Process each item in PITierList
    for (let i = 0; i < tierList.length; i++) {
      try {
        let item = tierList[i][0];
        if (!item) {
          logError(errorSheet, "Unknown", "Empty item name in PITierList.");
          continue;
        }

        let tier = tierMap[item];
        if (tier === undefined || !tierDataMap[tier]) {
          logError(errorSheet, item, "Tier not found in PITiers.");
          continue;
        }

        let [componentMultiplier, outputQuantity] = tierDataMap[tier];
        if (!outputQuantity || !componentMultiplier) {
          logError(errorSheet, item, "Invalid componentMultiplier or outputQuantity in PITiers.");
          continue;
        }

        if (!priceMap[item]) {
          logError(errorSheet, item, "Item price not found in PIPrices.");
          continue;
        }

        let sellPrice = priceMap[item] * outputQuantity;
        let components = requirementMap[item] || [];
        let totalComponentCost = 0;
        let componentCosts = {};
        let componentQuantities = {};

// Calculate component costs and quantities
components.forEach(component => {
  try {
    let componentName = component[1];
    if (!componentName) {
      logError(errorSheet, item, "Empty component name in PIRequirements.");
      return;
    }
    if (!priceMap[componentName]) {
      logError(errorSheet, componentName, "Component price not found in PIPrices.");
      return;
    }
    let componentPrice = priceMap[componentName];

    // FIX: Component Multiplier already accounts for scaling, do NOT multiply by outputQuantity
    let totalComponentNeeded = componentMultiplier; 
    let cost = componentPrice * totalComponentNeeded;

    totalComponentCost += cost;
    componentCosts[componentName] = cost;
    componentQuantities[componentName] = totalComponentNeeded;
  } catch (componentError) {
    logError(errorSheet, component[1] || "Unknown Component", componentError.message);
  }
});


        // Calculate profit
        let profit = sellPrice - totalComponentCost;
        let percentageProfit = totalComponentCost > 0 ? (profit / totalComponentCost) * 100 : 0;

        // Prepare row data
        let row = [item, percentageProfit.toFixed(2) + "%", `$${profit.toFixed(2)}`, outputQuantity];

        // Add input quantity as a separate column
        row.push(Object.values(componentQuantities).join(", "));

        // Add each component's cost as a separate column
        Object.keys(componentCosts).forEach(component => {
          row.push(`${component} x${componentQuantities[component]} ($${componentCosts[component].toFixed(2)})`);
        });

        results.push(row);
      } catch (itemError) {
        logError(errorSheet, tierList[i][0] || "Unknown Item", itemError.message);
      }
    }

// Output results to the sheet
if (results.length > 0) {
  let columnHeaders = ["Item", "Percentage Profit", "Profit", "Output Quantity",
                       "Component 1", "Input Quantity 1",
                       "Component 2", "Input Quantity 2",
                       "Component 3", "Input Quantity 3"];

  // Clear and write new headers
  sheet.clear();
  sheet.appendRow(columnHeaders);

  results.forEach(row => {
    let formattedRow = row.slice(0, 4); // Base columns: Item, Percentage Profit, Profit, Output Quantity

    let componentData = row.slice(4); // Component information
    let componentPairs = []; // Stores [component name, quantity] pairs

    for (let i = 0; i < componentData.length; i++) {
      // Remove any parenthetical currency value (e.g., " ($149096200.00)")
      let cleanedData = componentData[i].replace(/\s?\(\$[\d,\.]+\)/g, '').trim();
      
      let componentInfo = cleanedData.split(" x"); // Extract "Component Name" and "Quantity"
      if (componentInfo.length === 2) {
        let componentName = componentInfo[0].trim(); // Component name
        let componentQuantity = componentInfo[1].trim(); // Component quantity
        componentPairs.push([componentName, componentQuantity]); // Add to list
      }
    }

    // Ensure exactly 3 components are represented
    for (let i = 0; i < 3; i++) {
      if (i < componentPairs.length) {
        formattedRow.push(componentPairs[i][0]); // Component name
        formattedRow.push(componentPairs[i][1]); // Quantity only, no price or currency
      } else {
        formattedRow.push(""); // Leave blank for missing components
        formattedRow.push(""); // Leave blank for missing input quantities
      }
    }

    sheet.appendRow(formattedRow);
  });

  // Get all data from the sheet (including headers)
  let range = sheet.getDataRange();
  let values = range.getValues();

  // Sort data by the "Profit" column (index 2) in descending order
  values.sort((a, b) => b[2] - a[2]); // b[2] and a[2] are the profit columns

  // Write the sorted data back to the sheet
  range.setValues(values);

} else {
  logError(errorSheet, "General", "No data was processed. Check input sheets.");
}

  } catch (generalError) {
    logError(errorSheet, "General", generalError.message);
  }
}

/**
 * Retrieves data from a named range in Google Sheets.
 */
function getNamedRangeData(rangeName) {
  try {
    let range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
    if (!range) {
      console.log(`❌ Named range "${rangeName}" NOT FOUND`);
      return [];
    }
    let data = range.getValues();
    return data.length > 1 ? data.slice(1) : []; // Remove headers
  } catch (error) {
    console.log(`❌ Error retrieving named range "${rangeName}": ${error.message}`);
    return [];
  }
}

/**
 * Helper function to create a lookup object from a 2D array
 */
function createLookup(data, keyIndex, valueIndex) {
  let map = {};
  data.forEach(row => {
    if (row[keyIndex] !== "") {
      map[row[keyIndex]] = Array.isArray(valueIndex) 
        ? valueIndex.map(index => row[index]) 
        : row[valueIndex];
    }
  });
  return map;
}

/**
 * Helper function to create a lookup object where multiple rows can have the same key
 */
function createLookupMultiple(data, keyIndex) {
  let map = {};
  data.forEach(row => {
    if (!map[row[keyIndex]]) {
      map[row[keyIndex]] = [];
    }
    map[row[keyIndex]].push(row);
  });
  return map;
}

/**
 * Logs errors into the "ErrorLog" sheet with a timestamp
 */
function logError(sheet, item, message) {
  console.log(`Error: ${item} - ${message}`);
  sheet.appendRow([new Date().toISOString(), item, message]);
}



/////////////////////////////////////////////////////////////////////////////////

function updatePIComponentsNeeded() {
  const profitThreshold = 20000000; // Profit threshold for processing items
  
  // Get the PIProfitCalc range
  const profitCalcRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("PIProfitCalc").getValues(); 
  Logger.log("ProfitCalc Data: " + JSON.stringify(profitCalcRange)); // Log the data from PIProfitCalc
  
  // Get the PIComponentsNeeded range
  const componentsNeededRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("PIComponentsNeeded");
  
  // Prepare an array to store the data to be written to PIComponentsNeeded
  let componentsNeededData = [];
  
  // Loop through each row of the PIProfitCalc data
  profitCalcRange.forEach((row, index) => {
    if (index === 0) return; // Skip the header row
    
    let itemName = row[0]; // Item name (column 1)
    let profit = row[2]; // Profit amount (column 3)
    
    // Only process items with a profit greater than the threshold
    if (profit > profitThreshold) {
      let componentData = [];
      
      // Process components (columns 5, 7, 9 for component names and 6, 8, 10 for quantities)
      for (let i = 4; i <= 8; i += 2) { // Component name columns: 5, 7, 9
        let componentName = row[i];
        let componentQuantity = row[i + 1]; // Corresponding quantity columns: 6, 8, 10
        
        if (componentName && componentQuantity) {
          componentData.push([componentName, -componentQuantity]); // Add component with negative quantity
        }
      }
      
      // Add the item and components data to the componentsNeededData array
      componentData.forEach(component => {
        // Ensure 6 columns: component name, negative quantity, empty character column, empty placeholder column, item name
        componentsNeededData.push([
          component[0],               // Component name (column 1)
          component[1],               // Negative input quantity (column 2)
          "",                         // Empty column 3 (for character)
          "",                         // Empty column 4 (placeholder)
          "",                         // Empty column 5 (placeholder)
          itemName                    // Item name (column 6)
        ]);
      });
      
      Logger.log("Item: " + itemName + " has profit: " + profit + " and components: " + JSON.stringify(componentData));
    }
  });
  
  Logger.log("Components Needed Data: " + JSON.stringify(componentsNeededData));
  
  // Write the components needed data to PIComponentsNeeded range if we have data
  if (componentsNeededData.length > 0) {
    // Resize the range to match the number of rows of data
    componentsNeededRange.clear(); // Clear previous data
    
    // Resize the range to match the number of rows of data
    componentsNeededRange.offset(0, 0, componentsNeededData.length, 6).setValues(componentsNeededData); // Use 6 columns
  } else {
    Logger.log("No data to populate in PIComponentsNeeded.");
  }
}
