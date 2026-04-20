function updateDashboardWithTop7() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("PI Flip");
  var dashboardSheet = ss.getSheetByName("Dashboard");
  var dataRange = sourceSheet.getRange("PIFlip");
  var data = dataRange.getValues();
  
  // Check that the range is not empty
  if (data.length < 2) {
    Logger.log("No data found in PIFlip range.");
    return;
  }
  
  // Get values from column 6 and filter for those over 20,000,000
  var filteredData = data.filter(function(row) {
    return row[5] > 20000000; // Column 6 (index 5)
  });
  
  // Sort by 6th column in descending order
  filteredData.sort(function(a, b) {
    return b[5] - a[5];
  });
  
  // Get the top 7 rows
  var top7 = filteredData.slice(0, 7);
  
  // Extract required columns: 1, 2, 3, 4, 6, 7, 8, 9, 10, 13
  var output = top7.map(function(row) {
    return [
      row[0],  // Column 1
      row[1],  // Column 2
      row[2],  // Column 3
      row[3],  // Column 4
      row[5],  // Column 6
      row[6],  // Column 7
      row[7],  // Column 8
      row[8],  // Column 9
      row[9],  // Column 10
      row[12]  // Column 13
    ];
  });
  
  // Sort output by Column 6 (index 4) in descending order
  output.sort(function(a, b) {
    return b[4] - a[4];
  });
  
  // Clear the target range in Dashboard
  dashboardSheet.getRange("A13:J19").clearContent();
  
  // Paste the values into Dashboard A13:J19
  if (output.length > 0) {
    dashboardSheet.getRange(13, 1, output.length, output[0].length).setValues(output);
  }
  
  Logger.log("Top 7 values over 20,000,000 have been updated in 'Dashboard' A13:J19.");
}


