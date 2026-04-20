function runUpdates() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Script is running...", 10); // Display for 5 seconds
  //Set the updating range
  updateStatusRange("UPDATING");

  // Run the Janice update function
  updateJanice();
  
  // Wait for 10 seconds to ensure the PROCESSINGDATA sheet is fully populated.
  Utilities.sleep(10000);
  
  // Update the dashboard with the top 7 items.
 // updateDashboardWithTop7();

  //Update the outstanding orders values.
  processToonOrders()

  //Force Dashboard to update formulas
  updateDashboard();

  //Update the make list on Buy tab
  updateMakeList()
  
  //Set the updating range
  updateStatusRange("COMPLETE");
  
  //Update the date and time
  updateTime();

  Utilities.sleep(2000);

  //Update the time since updated
  updateSinceUpdate();

 SpreadsheetApp.getActiveSpreadsheet().toast("Script is complete...", 10); // Display for 5 seconds
}

function updateDashboard() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  var range = sheet.getRange("DashboardUpdate");
  
  // Toggle checkbox value
  var currentValue = range.getValue();
  var toggledValue = !currentValue;
  range.setValue(toggledValue);
  SpreadsheetApp.flush();
}

function updateDBMoney() {
  try {
    Logger.log("Starting updateDBMoney...");
    processToonOrders();
    updateDashboard();
    Logger.log("Finished updateDBMoney.");
  } catch (error) {
    Logger.log("Error in updateDBMoney: " + error);
  }
}


function updateStatusRange(processPoint) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRangeByName("UPDATING");
  var currentValue = range.getValue();  // Get the cell's value
  Logger.log(currentValue);
  Logger.log(processPoint);
  switch(processPoint) {
    case "UPDATING": 
      range.setValue("UPDATING");
      range.setFontColor("black");
      range.setBackground("yellow");
      break;
    case "COMPLETE":
      range.setValue("COMPLETE");
      range.setFontColor("black");
      range.setBackground("green");
      break;
    default:
      range.setValue("!!ERROR!!");
      range.setFontColor("black");
      range.setBackground("red");
      break;
  }
  
  range.setFontWeight("bold");
  range.setHorizontalAlignment("center");
}

function updateTime() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the range by its name "UpdatedTime"
  var range = spreadsheet.getRangeByName("UpdatedTime");
  
  // Set the range value to the current date and time
  range.setValue(new Date());
}

function updateSinceUpdate() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the date/time value from the named range "UpdatedTime"
  var updatedTimeRange = ss.getRangeByName("UpdatedTime");
  var updatedTime = updatedTimeRange.getValue();
  
  // Get the current date/time
  var now = new Date();
  
  // Check if the retrieved value is a valid date
  if (!(updatedTime instanceof Date)) {
    ss.getRangeByName("SinceUpdate").setValue("Invalid date in UpdatedTime");
    return;
  }
  
  // Calculate the difference in milliseconds
  var diffMs = now - updatedTime;
  
  // Convert difference into minutes first
  var totalMinutes = Math.floor(diffMs / (1000 * 60));
  
  // Calculate days, hours, and minutes
  var days = Math.floor(totalMinutes / (60 * 24));
  var hours = Math.floor((totalMinutes - days * 24 * 60) / 60);
  var minutes = totalMinutes % 60;
  
  // Create a human-readable string
  var result = days + " days, " + hours + " hours, " + minutes + " minutes ago";
  
  // Set the result into the named range "SinceUpdate"
  ss.getRangeByName("SinceUpdate").setValue(result);
}

