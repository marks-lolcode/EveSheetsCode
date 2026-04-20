/** -------
 * 
 * 
 */
function clearContentsOnly(clearsheet = "Buy", clearrange = "OwnedPI") {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(clearsheet);
  var range = sheet.getRange(clearrange);
  range.clearContent();
}