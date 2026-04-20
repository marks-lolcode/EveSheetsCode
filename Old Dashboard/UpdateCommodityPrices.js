/**
* Run the updated commodities for multiple locations and buy and sell types
* 
* @customfunction
* Author: Mark Sauntry
* Modified by Mark Sauntry 04/28/2023
*
*/
function allCommodityUpdate() {
  //List of market hubs
  const market_hub = ["Jita","Amarr"];
  //List of order types
  const order_type = ["buy","sell"]

  //Cycle through each market hub and order type to update prices
  for (var m = 0; m < market_hub.length; m++) {
    for (var o = 0; o < order_type.length; o++){
      let getprices = commodityPriceUpdateS(market_hub[m],order_type[o]);
    }
  }
}




/**
* Get updated commodities prices
*
* @param {string} string Jita, Amarr, Dodixie, Rens, Hek, Defaults to Jita.
* @return result for each type_id. This can be configured differently.
* @customfunction
* Author: Mark Sauntry
* Modified by Mark Sauntry 04/28/2023
*
*/
function commodityPriceUpdateS(market_hub = "Jita",order_type = "sell"){
  //Set the starting column number based on market hub
  const startColNbr = {
    "Jita": 3,
    "Amarr": 5
    };
  
  console.log("market_hub: "+market_hub);
  
  //Get the spreadsheet CommoditiesIDs
  const s = SpreadsheetApp.getActive().getSheetByName("CommoditiesIDs");
  
  //Sort by Commodity ID because Fuzzworks returns values sorted by id this automatically
  s.sort(2, true);
  const commodityNewValues = new Array();
  
  //Get the Commodity IDs
  let comIDswithnulls = s.getRange("CommodityID").getValues();
  
  //Remove null values from the Com IDs (empty rows at the end)
  let comIDs = comIDswithnulls.filter(String)
  
  //Get the length (number) of Com IDs
  let lastRow = comIDs.length;
  
  //Set the number of items to send to Fuzzworks at a time, too many will cause errors
  const size = 250
  
  //Split array up into smaller arrays  
  const chunks = chunker(comIDs.filter(Number), size);
  console.log("chunks:");
  console.log(chunks);
  

  //Find out how many chunks there are
  let numChunks = chunks.length;
  console.log ("numChunks: "+numChunks);

  //Itterate through the chunks and add the records to the commList
  for (var i = 0; i < numChunks; i++) {  
    console.log ("i: "+i);
  
    //Find how many records are in the chunk
    const chunkPartLength = chunks[i].length;
    console.log ("chunkPartLength: "+chunkPartLength);
    let commList = "";
  
    //Itterate through the chunk and add records into the commList with a comma in between to fit Fuzzworks format
    for (var j = 0; j < chunkPartLength; j++) {
      console.log ("j: "+j);
      console.log ("commListLength: "+commList.length);
  
      //If this is the first record do not add a comma, after that put a comma before appending
      if (commList.length <1) {
        commList += chunks[i][j];
      }  else {
        commList += ","+chunks[i][j];
      }
    }
   console.log ("commList: "+commList);



  let commodityVal = [[]];
  let commodityValAll = [[]];
  
  //Itterate through the chunks
  for (var k = 0; k < chunks.length; k++) {
    
    //Get Fuzzworks price for the items in the chunk
    commodityVal = fuzzPriceDataByHub(chunks[k],market_hub,order_type);
    console.log ("commodityVal:");
    console.log (commodityVal);

    //Combine the values into a single array
    commodityValAll.push(...commodityVal);

    console.log ("commodityValAll:");
    console.log (commodityValAll);
    }
  const newArr = [];
  while(commodityValAll.length) newArr.push(commodityValAll.splice(0,1));
  console.log (newArr);

  const newArrTwo = [];
  newArrTwo.push(newArr.splice(0,1));

  let startColumnNbr = startColumn(market_hub,order_type);
  let cvalength = newArr.length;
  console.log ("cvalength: "+cvalength);
  let startNumber = 2;
  var cell = s.getRange(startNumber,startColNbr[market_hub],cvalength,1);
  cell.setValues(newArr);
  console.log ("End");

  }
}

/**
* Chunk an array into pieces
*
* @param {array} input array.
* @param {integer} size of chunks
* @return result for each type_id. This can be configured differently.
* @customfunction
* Author: Mark Sauntry
* Modified by Mark Sauntry 04/28/2023
*
*/
function chunker (inputArray, size) {
  const chunks = [];
  const items = inputArray.slice();
  while (items.length) chunks.push(items.splice(0, size));
  return chunks;
}

/**
* Get the starting column based on market hub and order type
*
* @param {string} market_hub
* @param {string} order_type
* @return result for each type_id. This can be configured differently.
* @customfunction
* Author: Mark Sauntry
* Modified by Mark Sauntry 04/28/2023
*
*/
function startColumn(market_hub = "Jita",order_type = "sell") {
  //Assign the starting column based on the station system entered
  let startColumn;
  switch (market_hub) {
    case "Jita":
      startColumn = 3;
      break;
    case "Amarr":
      startColumn = 5;
      break;
    default:
      logger.log("Invalid Region");
      alertMessage("Station must be Jita or Amarr");
      return;
  }
  console.log("startColumn: "+startColumn);
  //Check to see if the order type is Buy, if so add one to the start column
  if (order_type === "buy") {
    startColumn++;
  }
  console.log("startColumn: "+startColumn);
  return startColumn;
}