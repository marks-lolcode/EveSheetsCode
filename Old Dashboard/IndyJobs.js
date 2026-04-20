function processData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const itemRange = sheet.getRangeByName('ItemsSheet2');  // Named range for ItemsSheet2
  const activityRange = sheet.getRangeByName('IndJobTypes');  // Named range for IndJobTypes
  const stationRange = sheet.getRangeByName('StationIDs');  // Named range for StationIDs
  const outputSheet = sheet.getSheetByName('IndyJobsParsed');  // Output sheet

  // Get jobs data from the getJobs2 function
  const data = getJobs2();  // Changed to getJobs2()
  const processedData = [];

  // Loop through each row of the data returned by getJobs2, starting from row 1 (skip the first row)
  for (let i = 1; i < data.length; i++) {  // Start from 1 to skip the first row
    const row = data[i];
    const activityId = row[0];  // Column 1 (activity_id)
    const itemId = row[16];  // Column 17 (ItemIDs)
    const dateValue = new Date(row[8]);  // Column 9 (Date)
    const facilityId = row[9];  // Column 10 (facility_id)
    const timeDifference = calculateTimeDifference(dateValue);

    // Cross-referencing activity_id using named range "IndJobTypes"
    const activity = crossReferenceByNamedRange(activityRange, activityId, "activityId", 'stringToNumber');
    // Cross-referencing item_id using the ItemsSheet2 named range
    const item = crossReferenceByNamedRange(itemRange, itemId, "itemId", 'stringToNumber');
    // Cross-referencing facility_id using the StationIDs named range
    const facility = crossReferenceByNamedRange(stationRange, facilityId, "facilityId", 'textToText');

    // Determine status based on date value
    const status = (dateValue < new Date()) ? 'Done' : 'Active';

    // Prepare the processed row
    const processedRow = [
      activity,  // Column 1 (from IndJobTypes)
      row[1],     // Column 2 (unchanged)
      row[2],     // Column 3 (unchanged)
      item,       // Column 17 (from ItemsSheet2)
      row[4],     // Column 5 (unchanged)
      row[5],     // Column 6 (unchanged)
      row[6],     // Column 7 (unchanged)
      row[7],     // Column 8 (unchanged)
      status,     // Column 21 ("Done" or "Active")
      facility,   // Column 10 (from StationIDs)
      row[22],    // Column 23 (unchanged)
      timeDifference  // Time difference for Column 23
    ];

    processedData.push(processedRow);
  }

  // Write the processed data back to the IndyJobsParsed sheet
  outputSheet.getRange(2, 1, processedData.length, processedData[0].length).setValues(processedData);
}

// Function to retrieve the jobs data (now using getJobs2)
function getJobs2() {
  // Sample structure for the jobs data, each row is an array of 23 columns
  return [
    ["Header1", "Header2", "Header3", "Header4", "Header5", "Header6", "Header7", "Header8", "Header9", "Header10", "Header11", "Header12", "Header13", "Header14", "Header15", "Header16", "Header17", "Header18", "Header19", "Header20", "Header21", "Header22", "Header23"],  // Row 0 (header)
    ["101", "value2", "value3", "type1", "value5", "value6", "value7", "value8", "2025-01-09T21:06:40Z", "facility1", "value11", "value12", "value13", "value14", "value15", "value16", "101", "value18", "value19", "value20", "Done", "facility2", "value22", "value23"]
    // Add more rows as needed
  ];
}

// Function to cross-reference using a named range (with multiple rows)
function crossReferenceByNamedRange(range, value, fieldName, typeConversion) {
  const data = range.getValues();  // Get data from the named range
  Logger.log(`Searching for ${fieldName}: '${value}'`);  // Log the value being searched
  
  // Convert the input value to the appropriate type based on the typeConversion parameter
  const convertedValue = handleTypeConversion(value, typeConversion);
  Logger.log(`Converted ${fieldName}: '${convertedValue}'`);  // Log the converted input value

  // Loop through all rows in the named range to find a match
  for (let i = 0; i < data.length; i++) {
    const namedRangeValue = handleTypeConversion(data[i][0], typeConversion);  // Compare with the first column
    Logger.log(`Row ${i}: Comparing '${namedRangeValue}' with '${convertedValue}'`);

    // Compare converted values
    if (namedRangeValue == convertedValue) {
      Logger.log(`Match found for ${fieldName}: ${data[i][1]}`);  // Log the result from the second column
      return data[i][1];  // Return the corresponding value from the second column (index 1)
    }
  }
  Logger.log(`No match found for ${fieldName}: '${value}'`);  // Log if no match is found
  return 'Not Found';  // Return 'Not Found' if no match
}

// Function to handle type conversions based on the `typeConversion` parameter
function handleTypeConversion(value, typeConversion) {
  // Log the value being processed to help with debugging
  Logger.log(`Handling value: ${value} (Type: ${typeof value})`);
  
  // Ensure the value is a string before calling .trim()
  if (typeof value !== 'string') {
    value = String(value);  // Convert value to string if it's not already a string
  }

  // Trim any extra spaces
  value = value.trim();
  
  switch (typeConversion) {
    case 'stringToNumber':
      return Number(value);  // Convert string to number
    case 'textToText':
      return value;  // No conversion needed, just return the text as it is
    default:
      return value;  // Default case (return the value unchanged)
  }
}

// Function to calculate the time difference in a human-readable format
function calculateTimeDifference(dateValue) {
  const now = new Date();
  const diffInMs = dateValue - now;
  const diffInMinutes = Math.floor(diffInMs / (1000 * 60));
  const diffInHours = Math.floor(diffInMinutes / 60);
  const diffInDays = Math.floor(diffInHours / 24);

  if (diffInDays > 0) {
    return `${diffInDays} days`;
  } else if (diffInHours > 0) {
    return `${diffInHours} hours`;
  } else if (diffInMinutes > 0) {
    return `${diffInMinutes} minutes`;
  } else {
    return 'Past';  // If it's in the past
  }
}
