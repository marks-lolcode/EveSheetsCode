//GESI Inventory for all characters.
function getAllInventory() {
  return GESI.invokeMultiple("characters_character_assets", GESI.getAuthenticatedCharacterNames());
}




/**
 * Processes inventory data retrieved from getAllInventory(),
 * cross-references type_id with named ranges, and writes the data to the "AllItems" sheet.
 */
function processInventoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("AllItems");

  // Create the "AllItems" sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet("AllItems");
    Logger.log("Created new sheet: AllItems");
  } else {
    // Clear previous contents entirely including headers
    sheet.clearContents();
    Logger.log("Cleared all contents from AllItems sheet.");
  }

  // Retrieve inventory data using the existing getAllInventory function
  const data = getAllInventory();
  Logger.log(`Retrieved inventory data with ${data.length} rows.`);

  // Extract headers and body
  const headers = data[0];
  const body = data.slice(1);

  // Load named ranges: invItemID (IDs), invItemName (Names), InvGroupID (Group IDs)
  const idRange = ss.getRangeByName("invItemID");
  const nameRange = ss.getRangeByName("invItemName");
  const groupRange = ss.getRangeByName("InvGroupID");

  // Load group to category and group name mappings
  const groupIdRange = ss.getRangeByName("sdeGroupID");
  const categoryIdRange = ss.getRangeByName("sdeCategoryID");
  const groupNameRange = ss.getRangeByName("sdeGroupName");

  if (!idRange || !nameRange || !groupRange || !groupIdRange || !categoryIdRange || !groupNameRange) {
    throw new Error("One or more named ranges are missing.");
  }

  const ids = idRange.getValues().flat();
  const names = nameRange.getValues().flat();
  const groups = groupRange.getValues().flat();
  const sdeGroupIds = groupIdRange.getValues().flat();
  const sdeCategoryIds = categoryIdRange.getValues().flat();
  const sdeGroupNames = groupNameRange.getValues().flat();

  const idToName = new Map();
  const idToGroup = new Map();
  const groupToCategory = new Map();
  const groupToName = new Map();

  // Build mappings from type_id to item name and group ID
  for (let i = 0; i < ids.length; i++) {
    idToName.set(ids[i], names[i]);
    idToGroup.set(ids[i], groups[i]);
  }

  // Build mappings from group ID to category ID and group name
  for (let i = 0; i < sdeGroupIds.length; i++) {
    groupToCategory.set(sdeGroupIds[i], sdeCategoryIds[i]);
    groupToName.set(sdeGroupIds[i], sdeGroupNames[i]);
  }

  Logger.log(`Built type_id, group_id, category, and group_name mappings.`);

  // Append item name, group ID, category ID, and group name as the last columns
  const output = body.map(row => {
    const typeId = row[7]; // type_id is expected at index 7
    const itemName = idToName.get(typeId) || "Unknown";
    const groupId = idToGroup.get(typeId) || "Unknown";
    const categoryId = groupToCategory.get(groupId) || "Unknown";
    const groupName = groupToName.get(groupId) || "Unknown";
    return [...row, itemName, groupId, categoryId, groupName];
  });

  Logger.log(`Prepared output with additional metadata. Writing ${output.length} rows to sheet.`);

  // Append new headers
  const extendedHeaders = [...headers, "item_name", "group_id", "category", "group_name"];
  sheet.getRange(1, 1, 1, extendedHeaders.length).setValues([extendedHeaders]);

  // Write full data below headers
  if (output.length > 0) {
    sheet.getRange(2, 1, output.length, output[0].length).setValues(output);
    Logger.log("Inventory data written to AllItems sheet with updated headers.");
  } else {
    Logger.log("No inventory data to write.");
  }
} 























//Load PI Inventory onto a spreadsheet
function inventoryToSheet(){
  
  //Get the PI inventory
  const updateInv = getPIInventory();
  
  //Create or empty the worksheet
  const invWorksheet = createOrClearSheet("InventoryData");
  
  //Set the range to start at the top and end at the end of the data
  const destinationRange = invWorksheet.getRange(1, 1, updateInv.length, updateInv[0].length);
  
  //Populate the range
  destinationRange.setValues(updateInv);
  
  //Clear any blank columns and cells
  deleteBlankColumnsAndCollumns(invWorksheet);
}

//Imports the assets for a character
function inventoryImport(charName) {
  
  //Get the assets
  var assets = GESI.characters_character_assets(charName);
  console.log ({assets});
  
  // Check if the assets are empty
  if (assets.length === 0) {
    console.log("No assets found for character " + charName);
    return null;
  }
  
  return assets;
}

//Get the PI Container ID from the Constants spreadsheet tab
function getPIContainerID (rangeName) {
  
  //Set the spreadsheet
  const s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
  
  //Get the PI Container name from the constants tab
  const piContainer= s.getRange(rangeName).getValues();
  console.log({piContainer});
  
  return piContainer;
}

function doesthiswork() {
var sw = "Sidewaze";
var charInitialss = {"Sidewaze": "SW", "LeftKnuckle": "LK", "Irard Geik Fidard": "IGF", "Agate West": "AW"};
console.log ({sw});
console.log (charInitialss.Sidewaze); 
console.log (charInitialss[sw]); 
console.log ({charInitialss});
} 


//Get invetory of PI items from identified containers for identified characters.
function getPIInventory(){
  var arrayPI = new Array();
  const invChars = ["Sidewaze","Agate West", "LeftKnuckle"]; //,"Irard Geik Fidard"

  //List of initials for characters
  const charInitial = {
   "Sidewaze": "SW",
   "LeftKnuckle": "LK",
   "Irard Geik Fidard": "IGF",
   "Agate West": "AW"};

  //List of Spreadsheet range names for PIID on Constants sheet
  const piContainerIDNames = {
   "Sidewaze": "SidewazePIID",
   "LeftKnuckle": "LeftKnucklePIID",
   "Irard Geik Fidard": "IrardPIID",
   "Agate West": "AgatePIID"};

  console.log("InvCHars Length: "+invChars.length);
  for (let n = 0; n < invChars.length; n++) {

    console.log("Char initial: "+charInitial[invChars[n]]);

    //Get the constant name
    var piContainerIDName = piContainerIDNames[invChars[n]];
    console.log({piContainerIDName});
    //Get the Container ID from the constants spreadsheet
    searchValueArray = getPIContainerID(piContainerIDName);
    console.log({searchValueArray});

    //Set the value to a number
    var searchValue = Number(searchValueArray[0]);
    console.log({searchValue});

    //Import the inventory for the character
    var inventory = inventoryImport(invChars[n]);
    console.log("Inventory Length: "+inventory.length);
    //Loop through the inventory looking for things in the PI container
    for (let i = 0; i < inventory.length; i++) {
      if(inventory[i][4] === searchValue){
        console.log("Found one: "+inventory[i][7]+" "+inventory[i][6]);
 
        //Look up the commodity name
        var commodityName = aLookup(inventory[i][7],"SDE_invTypes","ItemsSheet",2);
        console.log({commodityName});
 
        //Look up commodity weight
        var commodityWeight = Number(aLookup(inventory[i][7],"SDE_invTypes","ItemsSheet",3));
        console.log({commodityWeight});
 
        //Get description by the weight
        var commodityDescription = commodityTypeByWeight(commodityWeight);
        console.log({commodityDescription});
 
        //Add the item to the array to get pushed to the spreadsheet
        arrayPI.push(
          [commodityName,
          inventory[i][6],
          commodityDescription,
          inventory[i][7],
          commodityWeight,
          "",
          "",
          "",
          "",
          "",
          "",
          charInitial[invChars[n]]]);
        console.log({arrayPI});
      }
    }
  }
  console.log({arrayPI});
 
  return arrayPI;
}