function processToonOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Retrieve authenticated character names and call the orders function.
  var results = GESI.invokeMultiple("characters_character_orders", GESI.getAuthenticatedCharacterNames());
  
  // Get the sheet "Outstanding" and clear its contents.
  var outstandingSheet = ss.getSheetByName("Outstanding");
  outstandingSheet.clearContents();
  
  // If we have results, paste them into the "Outstanding" sheet.
  if (results && results.length > 0) {
    // Assume results is an array of arrays.
    var numRows = results.length;
    var numCols = results[0].length;
    outstandingSheet.getRange(1, 1, numRows, numCols).setValues(results);
  }
  

 /* - moved this to an in-cell formula 
  // Get the total number of rows. Assuming the first row is a header.
  var lastRow = outstandingSheet.getLastRow();
  if (lastRow < 2) return; // Exit if there's no data rows
  
  // Get all the data from the "Outstanding" sheet.
  var dataRange = outstandingSheet.getRange(1, 1, lastRow, outstandingSheet.getLastColumn());
  var data = dataRange.getValues();
  
  var totalProductSum = 0;
  
  // Loop through each data row (skip header row at index 0)
  for (var j = 1; j < data.length; j++) {
    // Column 9 (index 8) multiplied by Column 13 (index 12)
    var col9Val = parseFloat(data[j][8]) || 0;
    var col13Val = parseFloat(data[j][12]) || 0;
    var lineProduct = col9Val * col13Val;
    Logger.log("Row " + (j + 1) + " product: " + lineProduct);
    totalProductSum += lineProduct;
  }
  
  Logger.log("Total sum of line products: " + totalProductSum);
  
  // Set the total product sum into the named range "OutstandingSell".
  var targetRange = ss.getRangeByName("OutstandingSell");
  targetRange.setValue(totalProductSum);
*/
}
