/**
 * HELPER: Clear sheet or range of all data
 * 
 * Clears all content from a sheet or named range.
 * Accepts either a Sheet object or a Range object.
 * 
 * @param {Sheet|Range} sheetOrRange - The sheet or range to clear
 * @throws {Error} If object is invalid
 */
function clearSheetOrRange(sheetOrRange) {
  try {
    if (!sheetOrRange) {
      throw new Error("sheetOrRange is null or undefined");
    }
    
    // Check if it's a Range object (has getSheet method)
    if (typeof sheetOrRange.getSheet === 'function') {
      // It's a Range
      Logger.log(`Clearing range: ${sheetOrRange.getA1Notation()}`);
      sheetOrRange.clearContent();
    } 
    // Check if it's a Sheet object (has getDataRange method)
    else if (typeof sheetOrRange.getDataRange === 'function') {
      // It's a Sheet
      const sheetName = sheetOrRange.getName();
      Logger.log(`Clearing sheet: ${sheetName}`);
      sheetOrRange.getDataRange().clearContent();
    } 
    else {
      throw new Error("Input must be a Sheet or Range object");
    }
    
    Logger.log("Clear operation completed successfully");
  } catch (error) {
    Logger.log(`ERROR in clearSheetOrRange: ${error.message}`);
    throw error;
  }
}


/**
 * HELPER: Get existing sheet or create new one
 * 
 * Checks if a sheet with given name exists.
 * If it does, returns the sheet.
 * If not, creates a new sheet with that name and returns it.
 * 
 * @param {string} sheetName - The name of the sheet to find or create
 * @returns {Sheet} The existing or newly created sheet
 * @throws {Error} If spreadsheet is unavailable
 */
function getOrCreateSheet(sheetName) {
  try {
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error("sheetName must be a non-empty string");
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error("No active spreadsheet found");
    }
    
    // Try to get existing sheet
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (sheet) {
      Logger.log(`Sheet found: ${sheetName}`);
      return sheet;
    }
    
    // Sheet doesn't exist, create it
    Logger.log(`Sheet not found. Creating new sheet: ${sheetName}`);
    sheet = spreadsheet.insertSheet(sheetName);
    Logger.log(`Sheet created successfully: ${sheetName}`);
    
    return sheet;
    
  } catch (error) {
    Logger.log(`ERROR in getOrCreateSheet: ${error.message}`);
    throw error;
  }
}


/**
 * HELPER: Clear SystemCache data rows (keep header)
 * 
 * Clears all data from SystemCache (AB3:AD) in FitData sheet.
 * Preserves the header row (AB2:AD2).
 * 
 * @throws {Error} If FitData sheet not found
 */
function clearInventoryCache() {
 clearSheetOrRange(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FitData').getRange('AB3:AD2000'))
}

function dummy() {
  // Do nothing - used for menu headers only
}