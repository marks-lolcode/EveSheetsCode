/**
 * Determine what the commodity type is based on mass for Planetary Interaction materials
 * 
 * @param {"1.5"}                   itemMass  The mass of the item
 * @returns                         The commodity type
 * @customfunction
 */
function commodityTypeByWeight(itemMass=50) {
   const s = SpreadsheetApp.getActive(); 
   console.log("itemMass: "+itemMass);
   console.log(s.getRangeByName("Tier1Weight").getValue());
  switch(itemMass) {
    case s.getRangeByName("Tier0Weight").getValue():
      return "Planet Organic - Raw Resource";
      break;
    case s.getRangeByName("Tier1Weight").getValue():
      return "Basic Commodities - Tier 1";
      break;
    case s.getRangeByName("Tier2Weight").getValue():
      return "Refined Commodities - Tier 2";
      break;
    case s.getRangeByName("Tier3Weight").getValue():
      return "Specialized Commodities - Tier 3";
      break;
    case s.getRangeByName("Tier4Weight").getValue():
      return "Advanced Commodities - Tier 4";
      break;
    default:
      throw new error("Mass does not match any known PI values");
  } 
}

/**
 * Look up Commodity Name or ID or Volume must supply ID number for volume
 * 
 * @param {"Damage Control II or 108"}     luValue  Item to look up
 * @param {"'Name' or 'ID' or 'Volume'"}   nameOrID  What you want returned the name or the item ID or the volume
 * @returns                                The time the reference was last changed.
 * @customfunction
 */
  function comLookup(luValue = "Ibis", nameOrID = "ID") {
    console.log("luvalue: "+luValue);
    console.log("nameOrID: "+nameOrID);
  
    switch(nameOrID) {
      case "ID":
        var colToReturn = 0;
        var colToLookup = 2;
        break;
      case "Name":
        var colToReturn = 2
        var colToLookup = 0
        break;
      case "Volume":
        var colToReturn = 3
        var colToLookup = 0
        break;
      default:
        throw "comLookup: Need to enter what you are looking up, Name or ID";
    }
    console.log("colToReturn: "+colToReturn);
    console.log("colToLookup: "+colToLookup);
      var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SDE_invTypes");     
    
      var data =  s.getRange("ItemsSheet").getValues();
      console.log(data);
      var searchValue = luValue;
      console.log("searchValue: "+searchValue);
      var dataList = data.map(x => x[colToLookup]);
      console.log(dataList);
      var index = dataList.indexOf(searchValue);
    
      if (index === -1) {
        console.log("index: "+index);
        var foundValue = false;
        throw new Error('Value -'+searchValue+'- not found');
      } else {
      console.log("index: "+index);
      console.log("data[index][colToReturn]: "+data[index][colToReturn]);
          var foundValue = data[index][colToReturn];
      console.log("foundValue: "+foundValue);
      }
       return foundValue;
  }
  