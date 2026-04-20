function updateMakeList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Active Spreadsheet retrieved.");
  
  // Get the named ranges using the correct name for PITopItems.
  var topItemsRange = ss.getRangeByName("PITopItems");
  var makeListRange = ss.getRangeByName("PIMakeList");
  Logger.log("Named ranges retrieved: PITopItems and PIMakeList.");
  
  // Get the values from both ranges
  var topItemsData = topItemsRange.getValues(); // 2D array for PITopItems
  var makeListData = makeListRange.getValues(); // 2D array for PIMakeList
  Logger.log("Retrieved " + topItemsData.length + " rows from PITopItems.");
  Logger.log("Retrieved " + makeListData.length + " rows from PIMakeList.");
  
  // Build an array of text values from column 3 of PITopItems (index 2)
  var topItemValues = topItemsData.map(function(row) {
    return row[2];
  });
  Logger.log("Extracted values from column 3 of PITopItems: " + topItemValues);
  
  // Loop through each row in PIMakeList
  for (var i = 0; i < makeListData.length; i++) {
    var listCol6 = makeListData[i][5]; // Column 6
    var listCol4 = makeListData[i][3]; // Column 4
    var listCol5 = makeListData[i][4]; // Column 5

    Logger.log("Processing row " + (i + 1) + ": Col6=" + listCol6 + ", Col4=" + listCol4 + ", Col5=" + listCol5);
    
    if (topItemValues.indexOf(listCol6) !== -1) {
      Logger.log("Row " + (i + 1) + ": Col6 value '" + listCol6 + "' found in PITopItems.");
      // The value in column 6 matches one of the values from PITopItems.
      // If column 4 is empty, move value from column 5 to column 4.
      if (!listCol4) {
        makeListData[i][3] = listCol5;
        makeListData[i][4] = "";
        Logger.log("Row " + (i + 1) + ": Moved value from Col5 to Col4.");
      } else {
        Logger.log("Row " + (i + 1) + ": Col4 already has a value, no move needed.");
      }
    } else {
      Logger.log("Row " + (i + 1) + ": Col6 value '" + listCol6 + "' NOT found in PITopItems.");
      // The value in column 6 does NOT match any from PITopItems.
      // If column 4 has a value, move that value to column 5.
      if (listCol4) {
        makeListData[i][4] = listCol4;
        makeListData[i][3] = "";
        Logger.log("Row " + (i + 1) + ": Moved value from Col4 to Col5.");
      } else {
        Logger.log("Row " + (i + 1) + ": Col4 is empty, no move needed.");
      }
    }
  }
  
  // Write the updated data back to the PIMakeList range.
  makeListRange.setValues(makeListData);
  Logger.log("Updated data written back to PIMakeList.");
}


function generateCleanComponentList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("PIComponentList");
  const outputSheet = ss.getSheetByName("ProcessedOutput") || ss.insertSheet("ProcessedOutput");

  outputSheet.clear(); // Clean output sheet

  const data = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 2).getValues();

  let output = [["Item", "Component", "Quantity"]];
  
  data.forEach(([item, components], rowIndex) => {
    if (!item || !components) {
      console.log(`Skipping row ${rowIndex + 2}: Missing item or components.`);
      return;
    }
    
    // Regular expression to find all (Component Name + Quantity + Cost) sets
    const pattern = /(.+?) (\d+) ([\d.]+)/g;
    let match;
    
    while ((match = pattern.exec(components)) !== null) {
      const componentName = match[1].trim();
      const quantity = match[2];
      // const cost = match[3]; // Ignored, per your instructions
      
      console.log(`Row ${rowIndex + 2}: Item=${item}, Component=${componentName}, Quantity=${quantity}`);
      output.push([item, componentName, quantity]);
    }
  });

  outputSheet.getRange(1, 1, output.length, 3).setValues(output);
}
