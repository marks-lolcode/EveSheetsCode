/**
 * Get the profit for manufacturing an item
 * 
 * @param {"Damage Control II"}     itemName  Item to look up
 * @param {".9"}                    materialEff    Material Efficiency default .9
 * @returns                         The amount and percentage of profit for an item
 * @customfunction
 */
function bpProfit(itemName = "Proton S",materialEff = .9){
//Get BP ID to cross reference materials
const bpID = getbpID(itemName); //881
console.log ("bpID: "+bpID);
if (bpID == null || bpID == 0) {
  throw new Error("BP ID check failed")
}
//Get BP Info (name and how many are made from the blueprint)
const bpQuantity = getBPInfo(bpID);
console.log("bpQuantity: "+bpQuantity)   //100
if (bpQuantity == null || bpQuantity.length == 0) {
  throw new Error("BP ID check failed")
}
//Get BP materials
let bpMats = getMats(bpID);
console.log("bpMats: "+bpMats);
if (bpMats == null || bpMats.length == 0) {
  throw new Error("BP materials failed")
}

/**
 * [ [ 58, 34, 'Tritanium' ],
  [ 23, 35, 'Pyerite' ],
  [ 1, 37, 'Isogen' ] ]
 */

//Get materials cost
const matsCost = getMatsCost(bpMats);
console.log ("matsCost: "+matsCost);
if (matsCost == null || matsCost.length == 0) {
  throw new Error("BP materials cost failed")
}
//Get product sale amount
const bpSell = getBPSale(itemName);
console.log("bpSell: "+bpSell);
if (bpSell == null || bpSell.length == 0) {
  throw new Error("BP materials cost failed")
}
//Return results
let profitamt = (bpSell*bpQuantity) - matsCost;
console.log ("profitamt: "+profitamt);
let profitpct = (bpSell*bpQuantity) / matsCost;
console.log ("profitpct: "+profitpct);
let result = new Array();
result.push([profitamt],[profitpct]);
console.log("result: "+result);
return result;
}

/**
 * Get the BP ID to cross reference materials
 * 
 * @param {"Damage Control II"}     itemName  Item to look up
 * @returns                         Array with name and ID number for BP
 * @customfunction
 */
function getbpID(itemName="Proton S") {
//Look up the Item ID for the product
  let itemID = comLookup(itemName+" Blueprint","ID");
  console.log("itemID: ");
  console.log(itemID); //180
  return itemID;
}



/**
 * Get the BP quantity produced
 * 
 * @param {id of BP}                    bpID  The ID of the BP
 * @returns                             Array with quantity the BP produces
 * @customfunction
 */
function getBPInfo(bpID=881) {
  let bpInfo = new Array();
  const s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
  //Get the Products sheet from the SDE Import
  //Products sheet is BP ID, Activity ID, Product Type ID, Quantity
  let data =  s.getRange("Products").getValues();
    
  let dataFiltered = data.filter(function (manufacturing) {
    return manufacturing[1] === 1;
  }); 
  console.log(dataFiltered);
  
  let dataList = dataFiltered.map(x => x[0])
  console.log(dataList);
    let index = dataList.indexOf(bpID);
  
    if (index === -1) {
      console.log("index: "+index);
      throw new Error('Value not found');
      return;
    } else {
      console.log("index: "+index);
      console.log("data[index][0]: "+data[index][0]);
        //Return the number of items the BP makes
        bpInfo.push(dataFiltered[index][3]);
        console.log(bpInfo);
      return bpInfo;
}
}  
  
/**
 * Get the items needed to make a BP
 * 
 * @param {id of BP}                    bpID  The ID of the BP
 * @returns                             Array with quantity the BP produces
 * @customfunction
 */
function getMats(bpID=881){
  let manufacturingInfos = new Array();
  console.log("bpid: "+bpID);
  const s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
  //Get info from Materials spreadsheet from SDE import
  //typeID	activityID	materialTypeID	quantity
  let data = s.getRange("Materials").getValues();
  console.log("data: "+data);
  //Cycle through the Materials array
  for (let i = 0; i < data.length; i++) {
    //Look for the bpID to match the item ID and for the activity to be 1 (Manufacturing)
    //console.log("typeID and Activity: "+data[i][0]+" - "+data[i][1]);
    if (data[i][0] === bpID && data[i][1] === 1) {
      let dataName = comLookup(data[i][2],"Name");
      console.log("dataName: "+dataName);
      manufacturingInfos.push([data[i][3],data[i][2],dataName]);
      console.log("manufacturingInfos:");
      console.log(manufacturingInfos);
    }
  }
/**
 *  34	58
    35	23
    37	1
 */
console.log("manufacturingInfos:");
console.log(manufacturingInfos);
  return manufacturingInfos;
}

/**
 * Get the cost of the materials
 * 
 * @param {list of materials}           bpMats list of the materials
 * @returns                             Array with quantity the BP produces
 * @customfunction
 */
function getMatsCost(bpMats = [ [ 58, 34, 'Tritanium' ],[ 23, 35, 'Pyerite' ],[ 1, 37, 'Isogen' ] ]){
 let totalCost = 0
 const s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //Get info from Janice costs spreadsheet
  //JaniceCommodityName	JaniceJitaBuy	JaniceJitaSell	JaniceAmarrBuy	JaniceAmarrSell
  let data = s.getRange("JaniceCommodities").getValues();
    for (let i = 0; i < bpMats.length; i++) {
    for (let j = 0; j < data.length; j++) {
      console.log("data and mats: "+data[j][0]+" - "+bpMats[i][2]);
      if (data[j][0] === bpMats[i][2]) {
        console.log("bpMats - data: "+bpMats[i][0]+" - "+data[j][1]); 
      let itemcost = bpMats[i][0]*data[j][1];
      totalCost = totalCost + itemcost;
      console.log("totalCost: "+totalCost);
    }
    }
  }
  console.log("totalCost: "+totalCost);
  return totalCost;
}


/**
 * Get the sale amount of the product
 * 
 * @param {item name}                   itemName  list of the materials
 * @returns                             Array with quantity the BP produces
 * @customfunction
 */
function getBPSale(itemName){
 let totalSell = 0
 let s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //Get info from Janice costs spreadsheet
  //JaniceCommodityName	JaniceJitaBuy	JaniceJitaSell	JaniceAmarrBuy	JaniceAmarrSell
  var data = s.getRange("JaniceCommodities").getValues();
  for (j = 0; j < data.length; j++)
    if (data[j][0] === itemName) {
      totalSell =+ data[j][2];
    }
  return totalSell;
}













/** OLD CODE */






function getBP(itemName) {
  var idQuantity = new Array();
  //Look up the Item ID for the item
  var itemID = comLookup(itemName,"ID");
  console.log("itemID: ");
  console.log(itemID);
  if (itemID === false) {
    //var addCom = addCommodity(itemName);
    //var itemID = comLookup(itemName,"ID");
    return idQuantity;
  }
    var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
    
    var data =  s.getRange("Products").getValues();
    
    var dataFiltered = data.filter(function (manufacturing) {
      return manufacturing[1] === 1;
    }); 
    
    var dataList = dataFiltered.map(x => x[2])
    
      var index = dataList.indexOf(itemID);
    
      if (index === -1) {
        console.log("index: "+index);
        throw new Error('Value not found');
        return idQuantity;
      } else {
      console.log("index: "+index);
      console.log("data[index][0]: "+data[index][0]);
        idQuantity.push(itemID,dataFiltered[index][0],dataFiltered[index][3]);
        console.log(idQuantity);
        //var bpid = data[index][0];
      //console.log("foundValue: "+bpid);
       //return bpid;
       return idQuantity;
    }
  }








/**
 * Get the profit for manufacturing an item
 * 
 * @param {"Damage Control II"}     itemName  Item to look up
 * @param {".9"}                    matEff    Material Efficiency default .9
 * @returns                         The time the reference was last changed.
 * @customfunction
 */
function mProfit(itemName = "Proton S",matEff = .9){
  //Get the BP Info
  var manufacturingInfo = getBP(itemName);
  console.log("manufacturingInfo: ");
  console.log(manufacturingInfo);
  //Check if there was a result
  if (manufacturingInfo.length === 0) {
    console.log("Could not find item");
    throw Error("Could not find item");
    return;
  }
  //Get the items needed to complete a BP
  var manufacturingItems = getItems(manufacturingInfo[1]);
  console.log("manufacturingItems: ");
  console.log(manufacturingItems);
  //Get the cost of the items
  var totalCost = getCost(manufacturingItems,matEff);
  console.log("Cost: "+totalCost);
  //Get the value of the item
  var totalValue = getValue(itemName,manufacturingInfo[2])
    console.log("Total Value: "+totalValue);
    if (totalValue === 0){
      throw Error ("Item Not Found: "+manufacturingInfo[0]);
      //var addCom = addCommodity(itemName);
      //var totalValue = getValue(manufacturingInfo)
      return;
    }
  console.log("Value: "+totalValue);
  var profit = totalValue - totalCost;
  console.log("profit: "+profit);
  var profitPct = profit / totalCost;
  console.log("Profit %"+profitPct);
  var result = [[profit,profitPct]];
  return result;
}

//Get the BP name and number from an item name
function getBP(itemName) {
  var idQuantity = new Array();
  //Look up the Item ID for the item
  var itemID = comLookup(itemName,"ID");
  console.log("itemID: ");
  console.log(itemID);
  if (itemID === false) {
    //var addCom = addCommodity(itemName);
    //var itemID = comLookup(itemName,"ID");
    return idQuantity;
  }
     var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
    
    var data =  s.getRange("Products").getValues();
    
    var dataFiltered = data.filter(function (manufacturing) {
      return manufacturing[1] === 1;
    }); 
    
    var dataList = dataFiltered.map(x => x[2])
    
      var index = dataList.indexOf(itemID);
    
      if (index === -1) {
        console.log("index: "+index);
        throw new Error('Value not found');
        return idQuantity;
      } else {
      console.log("index: "+index);
      console.log("data[index][0]: "+data[index][0]);
        idQuantity.push(itemID,dataFiltered[index][0],dataFiltered[index][3]);
        console.log(idQuantity);
        //var bpid = data[index][0];
      //console.log("foundValue: "+bpid);
       //return bpid;
       return idQuantity;
    }
  }

//Get the items needed to make a BP
function getItems(bpid){
  var manufacturingInfos = new Array();
  console.log("bpid: "+bpid);
  var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
  var data =  s.getRange("Materials").getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === bpid && data[i][1] === 1) {
      dataName = comLookup(data[i][3],"Name");
      manufacturingInfos.push([data[i][2],data[i][3]],dataName) 
      console.log("manufacturingInfos:");
      console.log(manufacturingInfos);
    }
  }
  return manufacturingInfos;
}

function getCost(itemList,matEff = .9){
  var itemCost = 0;
  var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
  var data =  s.getRange("JaniceCommodities").getValues();
  for(i = 0; i < itemList.length; i++){
    for (var j = 0; j < data.length; j++) {
      if (data[j][1] === itemList[i][0] ) {
        console.log("dataj3: "+data[j][1]+" itemList1: "+itemList[i][1]);
        itemCost = (itemCost + (Number(data[j][2]) * Math.ceil(Number(itemList[i][1]) * matEff)));
        console.log("Item cost: "+itemCost);
      }
    }
  }
  return itemCost;
}

function getValue(itemName,quantity){
  var itemValue = 0;
  var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
  var data =  s.getRange("JaniceCommodities").getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === itemName)  {
      itemValue = Number(data[i][2]) * Number(quantity);
      console.log("ItemValue: "+itemValue)
      break;
    }
  }
      return itemValue;
}



