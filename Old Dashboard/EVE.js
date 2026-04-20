function togglePolymers(tf=false) {
  var cBoxes = ["B5","B7","B13","B15","B21","B23","B29","B31","B37","B39","B45","B47","B53","B55","B61","B63","B69","B71"]
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("Polymer");
  for (var i = 0; i < cBoxes.length; i++){
    spreadsheet.getRange(cBoxes[i]).activate();
    spreadsheet.getCurrentCell().setValue(tf);
  }
}

function togglePolymersOff() {
  togglePolymers(false);
}
  
function togglePolymersOn() {
  togglePolymers(true);
}
  
  
/**
 * Add a commodity item to the end of the list
 * 
 * @param {"Damage Control II"}     comName  Item to look up
  * @returns                        Nothing
 * @customfunction
 */function addCommodity(comName = false) {
  const s = SpreadsheetApp.getActive()
  const sheet = s.getSheetByName('CommoditiesIDs')
  if (comName === false){
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter New Commodity Name', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      var buttonStatus = true;
      var newComName = response.getResponseText();
    } else {
      var buttonStatus = false;
    }
  } else {
    var buttonStatus = true;
    var newComName = comName;
  }

  if (buttonStatus === true) {
    //Check to see if the item already exists in the sheet
    var aVals = s.getRange("A1:A").getValues();
    const includesMultiDimension = (arr, str) =>
    JSON.stringify(arr).includes(str);
    var alreadyExists = (includesMultiDimension(aVals, newComName));
    console.log ("alreadyExists: "+alreadyExists);
    if (alreadyExists === true) {
      SpreadsheetApp.getUi().alert("This commodity already exists");
      return
    } else {
      const lr = sheet.getLastRow;
      //Get the Commodity ID
      var comID = comLookup(newComName,"ID");      
      var newRow = new Array;
      //Get Jita value
      var commodityValJita = commodityValueFW(60003760,comID);
      console.log ("commodityValJita: "+commodityValJita);
      var jitaSell = commodityValJita[0][0];
      var jitaBuy = commodityValJita[0][1];
      //Get Amarr Value
      var commodityValAmarr = commodityValueFW(60008494,comID);
      console.log ("commodityValAmarr: "+commodityValAmarr);
      var amarrSell = commodityValAmarr[0][0];
      var amarrBuy = commodityValAmarr[0][1];
      newRow.push(newComName,comID,jitaSell,jitaBuy,amarrSell,amarrBuy);
      console.log(newRow);
      //Add to end of sheet
      sheet.appendRow(newRow);

/*    changed to single row push  
      var Bvals = s.getRange("A1:A").getValues();
      var lastRow = Bvals.filter(String).length;
      console.log ("lastRow ="+lastRow);
      var row = lastRow;
      var col = 2; //Set to CommodityID column
      console.log ("Row: "+row);
      console.log ("comID: "+comID);
      var cell = sheet.getRange(row,col);
      console.log ("cell0: " + comID);
      cell.setValue(comID);
      var cell = sheet.getRange(row,col+1);
      console.log ("cell1: " + cell);
      cell.setValue(sell);
      var cell = sheet.getRange(row,col+2);
      console.log ("cell2: " + cell);
      cell.setValue(buy);
        var cell = sheet.getRange(row,col+3);
      console.log ("cell1: "+cell);
      cell.setValue(sell);
      var cell = sheet.getRange(row,col+4);
      console.log ("cell2: "+cell);
      cell.setValue(buy);
*/
    }
  }
};
  
  function SHEETNAME() {
    let activeSheet = SpreadsheetApp.getActiveSheet();
    return activeSheet.getName();
  }
  
  function alertMessage(showme) {
   console.log("showme: "+showme);
    SpreadsheetApp.getUi().alert(showme);
  }
  
function quicktest(){
   var s = SpreadsheetApp.getActive().getSheetByName("CommoditiesIDs");
   var test = s.getLastRow();
   console.log ("test ="+test);
   test = 3;
    for (var i = 2; i < test; i++) {
    console.log ("i: "+i);
    var row = i;
    var col= 2 ;
    var value = s.getRange(row, col).getValue();
    console.log ("value= "+value);
    var quicktestval = commodityValueFW(60003760,value);
    console.log ("quicktestval: "+quicktestval);
    var jitaSell = quicktestval[0][0];
    var jitaBuy = quicktestval[0][1];
    var cell = s.getRange(row,col+1);
    console.log ("cell: "+cell);
    console.log ("quicktesteval[0]: "+quicktestval[0][0]);
    cell.setValue(jitaSell);
    var cell = s.getRange(row,col+2);
    console.log ("cell: "+cell);
    cell.setValue(jitaBuy);
    } 
}


  function initCounters(){
    var scriptPrp = PropertiesService.getScriptProperties()
    scriptPrp.setProperty('jitaCounter', '0');
    scriptPrp.setProperty('jitaMax', '0');
    scriptPrp.setProperty('amarrCounter', '0');
    scriptPrp.setProperty('amarrMax', '0');
    scriptPrp.setProperty('JitaUpdateCounter', '0');
    scriptPrp.setProperty('AmarrUpdateCounter', '0');
    
    Logger.log('go!')
    }
    

  function updateJitaPrices(){
    var updatJita = updatePrices("Jita");
/*
    const s = SpreadsheetApp.getActive().getSheetByName("CommoditiesIDs");
    //Check a checkbox on the sheet to override the 24 hour minimum, if True do not stop if fail time check.
    var timeOverride = s.getRange(1,15);
    var chkboxOverride = timeOverride.isChecked();
    //var chkboxOverride = timeOverride.getValues();
    Logger.log("chkboxOverride: "+chkboxOverride);
    console.log("chkboxOverride: "+chkboxOverride);
    var updateTime = updateTimeCheck();
    Logger.log("updateTime: "+updateTime);
    if (chkboxOverride === false && updateTime === true) {
            Logger.log("Not time to update yet");
            alertMessage("Not time to update yet");
        } else {
            var scriptPrp = PropertiesService.getScriptProperties();
            var counter = scriptPrp.getProperty('jitaCounter')*1;
            Logger.log("jitaCounter: "+counter);
            if(counter < 2) {
                counter = 2;
            }
            const prog = s.getRange(3,12);
            Logger.log("jitaCounter min 2: "+counter);
                //var counterMax = scriptPrp.getProperty('jitaMax')  
                //Logger.log("jitaMax: "+counterMax);
            //Determine the total rows with a value in Column B
            var Bvals = s.getRange("B1:B").getValues();
            var lastRow = Bvals.filter(String).length;
            Logger.log ("lastRow ="+lastRow);
            if(counter>=lastRow+2){
                Logger.Error("Already Completed - counter/lastrow: "+counter+" - "+lastRow);
            } else {
                for (var i = counter; i < lastRow+1; i++) {
                    var currentNumber =  i + "/" + lastRow +"  "+Math.floor((i/lastRow)*100)+"%";
                    console.log(currentNumber);
                    if (i % 10 === 0){
                      s.getRange('JitaUpdateCounter').setValue(currentNumber);
                      //s.getRange(3,13).setValue(currentNumber);
                    }
                    console.log ("------i: "+i);
                    var row = i;
                    var col= 2; //Set to CommodityID column
                    var comID = s.getRange(row, col).getValue();
                    console.log ("value= "+comID);
                    var commodityVal = commodityValueFW(60003760,comID);
                    console.log ("commodityVal: "+commodityVal);
                    var sell = commodityVal[0][0];
                    var buy = commodityVal[0][1];
                    var cell = s.getRange(row,col+1);
                    console.log ("cell1: "+cell);
                    cell.setValue(sell);
                    var cell = s.getRange(row,col+2);
                    console.log ("cell2: "+cell);
                    cell.setValue(buy);
                    scriptPrp.setProperty('jitaCounter', i);
                }
                var timeUpdate = updateEndTime();
                scriptPrp.setProperty('jitaCounter', '0');
                s.getRange(3,13).setValue("COMPLETE");
            }
        }
*/
}


function updateEndTime() {
    const s = SpreadsheetApp.getActive().getSheetByName("CommoditiesIDs");
    var endTime = s.getRange(1,16);
    var d = new Date();
    endTime.setValue(d);
    return true;
}

//Checks the last time Jita Update was run and returns True if less than 24 hours
function updateTimeCheck() {
    const MILLIS_PER_DAY = 1000 * 60 * 60 * 24
    Logger.log("MILLIS_PER_DAY: "+MILLIS_PER_DAY);
    console.log("MILLIS_PER_DAY: "+MILLIS_PER_DAY);
    const s = SpreadsheetApp.getActive().getSheetByName("CommoditiesIDs");
    var sheetTime = s.getRange(1,16);
    var finishTime = new Date(sheetTime.getValues());
    Logger.log("Finish Time: "+finishTime);
    console.log("Finish Time: "+finishTime);
    var d = new Date();
    Logger.log("Now: "+d);
    var timeDiff = (d.getTime()-finishTime.getTime())*1;
    Logger.log("Difference: "+timeDiff);
    console.log("Now: "+d);
    console.log("Difference: "+timeDiff);
    if(timeDiff>MILLIS_PER_DAY){
      return false; //false means ok to run
    } else {
      return true; //true means do not run
    }
}



function updateAmarrPrices(){

    var updatAmarr = updatePrices("Amarr");

/*    var s = SpreadsheetApp.getActive().getSheetByName("CommoditiesIDs");
    //var lastRow = s.getLastRow();
    var scriptPrp = PropertiesService.getScriptProperties();
    var counter = scriptPrp.getProperty('amarrCounter')*1;
    Logger.log("amarrCounter: "+counter);
    if(counter < 2) {
        counter = 2;
    }
    Logger.log("amarrCounter min 2: "+counter);
    var Bvals = s.getRange("B1:B").getValues();
    var lastRow = Bvals.filter(String).length;
    console.log ("lastRow ="+lastRow);
    for (var i = counter; i < lastRow+1; i++) {
        var currentNumber =  i + "/" + lastRow +"  "+Math.floor((i/lastRow)*100)+"%";
        console.log(currentNumber);
        console.log(i % 10)
        if (i % 10 === 0){
          s.getRange('AmarrUpdateCounter').setValue(currentNumber);
         // s.getRange(4,13).setValue(currentNumber);
        }
        console.log ("------i: "+i);
        var row = i;
        var col= 2; //Set to CommodityID column
        var comID = s.getRange(row, col).getValue();
        console.log ("value= "+comID);
        var quicktestval = commodityValueFW(60008494,comID);
        console.log ("quicktestval: "+quicktestval);
        var sell = quicktestval[0][0];
        var buy = quicktestval[0][1];
        var cell = s.getRange(row,col+3);
        console.log ("cell1: "+cell);
        cell.setValue(sell);
        var cell = s.getRange(row,col+4);
        console.log ("cell2: "+cell);
        cell.setValue(buy);
        scriptPrp.setProperty('amarrCounter', i);
    }
        scriptPrp.setProperty('amarrCounter', '0');
        s.getRange(4,13).setValue("COMPLETE");
*/
}


  //Core function, takes the input and runs the other scripts to get the output
  function commodityValueEM(region, commodityID) {
    var amt = [];
      console.log ("region: "+region);
      console.log ("commodityID: "+commodityID);
 //     if(commodityID.map) {
 //       return input.map(commodityID);
 //     } else {
        var sysTrans = getSysInfo(region); //Splits the region name and transaction type
        var regionName = sysTrans[0]; 
        var transactionType = sysTrans[1]
        console.log ("sysTrans: "+sysTrans);
        console.log ("regionName: "+regionName)
        console.log ("transactionType: "+transactionType)
        var regID = aLookup(regionName,"SystemIDs","Regions",1); //Gets the system ID from the region name
        console.log ("sysID: "+regID);
        amt = getEMdata(regID,commodityID,transactionType); //Gets the value of the commodity for the transaction type
        console.log ("amt: "+amt);
        return amt;
//      }
}
  
 //Split out the system name and transaction type
  function getSysInfo(sysNameType="The Forge Sell") {
      const nameArray = sysNameType.split(" ");
      console.log("nameArray: "+nameArray);
      console.log("nameArray length: "+nameArray.length);
      var namelen = nameArray.length;
      console.log("namelen: "+namelen);
      switch(namelen) {
        case 2:
          var sysName = nameArray[0];
          var transType = nameArray[1];
          break;
        case 3:
          var sysName = nameArray[0]+" "+nameArray[1];
          var transType = nameArray[2];
          break;
        default:
          throw new Error('Problem splitting name');
      }
      console.log("sysName: "+sysName);
      console.log("transType: "+transType);
      
      return [sysName,transType];
  }
  
function qtcomLookup() {
console.log(comLookup("Ibis","ID"));
}







  //Look up data in a range
  function aLookup(luValue, luSheet, rangeName, colToReturn) {
     
      console.log("luvalue: "+luValue);
      console.log("lusheet: "+luSheet);
      console.log("rangename: "+rangeName);
      console.log("coltoreturn: "+colToReturn);
    
      var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();     
    
      var data =  s.getRange(rangeName).getValues();
    
      var searchValue = luValue;
      console.log("searchValue: "+searchValue);
      var dataList = data.map(x => x[0])
    
      var index = dataList.indexOf(searchValue);
    
      if (index === -1) {
        console.log("index: "+index);
        throw new Error('Value not found');
    
      } else {
      console.log("index: "+index);
      console.log("data[index][colToReturn]: "+data[index][colToReturn]);
          var foundValue = data[index][colToReturn];
      console.log("foundValue: "+foundValue);
       return foundValue;
      }
    }
  
//Use the ImportJSON script to get the data from EveMarketer
function getEMdata(sysID="10000002", itemID, buysell="Sell") {
  var conurl = "http://evetycoon.com/api/v1/market/stats/"+sysID+"/"+itemID;
  console.log("conurl: "+conurl);
  var comVal = ImportJSON(conurl);
  console.log("comVal: "+comVal);
  console.log("BA5: "+comVal[1][8]);
  console.log("SA5: "+comVal[1][9]);
  var rows = [];
  var rows = [[comVal[1][9],comVal[1][8]]];
  console.log("rows: "+rows);
  return rows;
}

//Use the ImportJSON script to get the data from Fuzzworks
function getFWdata(sysID="60003760", itemID) {
  var conurl = "https://market.fuzzwork.co.uk/aggregates/?station="+sysID+"&types="+itemID;
  console.log("conurl: "+conurl);
  var comVal = ImportJSON(conurl);
  console.log("comVal: "+comVal);
  console.log("BPer: "+comVal[1][7]);
  console.log("SPer: "+comVal[1][15]);
  var rows = [];
//Values are in [1], Percentile Buy is 7, Percentile Sell is 15 
  var rows = [[comVal[1][15],comVal[1][7]]];
  console.log(rows);
  return rows;
}
/** Fuzzworks return values, numbers added for reference. Aditional items repeat order
1 45998	
buy	
0 weightedAverage	"4118163.179916318"
1 max	"4399000.0"
2 min	"2500000.0"
3 stddev	"475516.590813116"
4 median	"4225000.0"
5 volume	"1434.0"
6 orderCount	"22"
7 percentile	"4396884.239888424"
sell	
8 weightedAverage	"5934742.208135235"
9 max	"19990000.0"
10 min	"4672000.0"
11 stddev	"1843797.2697951498"
12 median	"5384500.0"
13 volume	"3786.0"
14 orderCount	"122"
15 percentile	"4838020.602218701"
 */





function qc() {
  var qcTest = inventoryImport('Sidewaze');
  console.log("qcTest: "+qcTest);
}





function Quick(){
  var commodityWeight = aLookup(9838,"SDE_invTypes","ItemsSheet",3);
  console.log("commodityWeight: "+commodityWeight);
}





function arraytest(){
var test = [["SW"],[9836],["Superconductors"],[1.5],["Tech 1"]];
test.push(["LK"],[9836],["Superconductors"],[1.5],["Tech 1"]);
console.log(test);
}




//Check a checkbox on the sheet to override the 24 hour minimum, if True do not stop if fail time check.
function timeToUpdate(s) {
  var timeOverride = s.getRange(1,15);
    var chkboxOverride = timeOverride.isChecked();
    //var chkboxOverride = timeOverride.getValues();
    Logger.log("chkboxOverride: "+chkboxOverride);
    console.log("chkboxOverride: "+chkboxOverride);
    var updateTime = updateTimeCheck();
    Logger.log("updateTime: "+updateTime);
    if (chkboxOverride === false && updateTime === true) {
            Logger.log("Not time to update yet");
            alertMessage("Not time to update yet");
            return false;
    } else {
      return true;
    }
  
}

//
function scriptCounter (counterName) {
  console.log ("counterName: " + counterName);
  var scriptPrp = PropertiesService.getScriptProperties();
  var counter = scriptPrp.getProperty(counterName)*1;
  Logger.log(counterName+ ": " + counter);
  return counter;
}


function upTest() {
  updatePrices("The Forge");
}



function updatePrices(station = "Amarr"){
  const s = SpreadsheetApp.getActive().getSheetByName("CommoditiesIDs");
  const scriptPrp = PropertiesService.getScriptProperties();
  const commodityNewValues = new Array();
  var counterRange = "";
  var startColumn = 7;
  var region = 60003760;
  switch (station) {
    case "Jita":
      counterRange = "JitaUpdateCounter";
      startColumn = 3;
      region = 60003760;
      break;
    case "Amarr":
      counterRange = "AmarrUpdateCounter";
      startColumn = 5;
      region = 60008494;
      break;
    default:
      logger.log("Invalid Region");
      alertMessage("Station must be Jita or Amarr");
      return;
  }
  if (station == "Amarr" || timeToUpdate(s) == true) {
    var counter = scriptCounter(counterRange);      
    if(counter < 1) {
      counter = 1;
  } else {
    return;
  }
  Logger.log(region + " counter min 1: "+counter);
  console.log("counterRange: "+counterRange);
  console.log("startColumn: "+startColumn);
  console.log("region: "+region);
  //Get the value of the update counter
  const prog = s.getRange(counterRange).getValue;
  //Determine the total rows with a value in Column B
  var comIDs = s.getRange("CommodityID").getValues();
  var lastRow = comIDs.filter(String).length;
  Logger.log ("lastRow ="+lastRow);
  if(counter>=lastRow+2){
    Logger.Error("Already Completed - counter/lastrow: "+counter+" - "+lastRow);
    } else {
      //Loop through the Commodity IDs
      for (var i = counter; i < lastRow+1; i++) {
        console.log ("------i: "+i);
        var comID = comIDs[i];
        console.log ("value= "+comID);
        var commodityVal = commodityValueFW(region,comID);
        console.log (commodityVal);
        var sell = commodityVal[0][0];
        var buy = commodityVal[0][1];
        commodityNewValues[i] = [sell,buy];
        console.log (commodityNewValues);
        }
//        if (i % 370 === 0){
          var j = scriptPrp.getProperty(counterRange);
          if (j == 0) {
            var startNumber = 2;
          } else {
              var startNumber = j;
          }
          console.log ("startNumber: "+ startNumber)
          var endNumber = startNumber + commodityNewValues.length;
          var rangeNumber = endNumber - startNumber - 1;
          var slicedResults = [];
          slicedResults = commodityNewValues.slice(commodityNewValues[startNumber],commodityNewValues.length)
          console.log(slicedResults);
          var slicedResultsClean = [];
          slicedResultsClean =  slicedResults.filter(e =>  e);
          console.log(slicedResultsClean);
//          for (var k = 0; k < slicedResultsClean.length; k++) {
//            console.log ("k: "+k);
            console.log ("startColumn: "+startColumn);
            console.log ("SR length: "+slicedResultsClean.length)
            console.log ("SR1 length: "+slicedResultsClean[1].length)
            var cell = s.getRange(startNumber,startColumn,slicedResultsClean.length,slicedResultsClean[1].length);
            console.log ("cell: "+cell);
            cell.setValues(slicedResultsClean);
//          }
          //Update the status cell
//          var currentNumber =  i + "/" + lastRow +"  "+Math.floor((i/lastRow)*100)+"%";
//          console.log(currentNumber);
//         s.getRange(counterRange).setValue(currentNumber);

//      }
    }
      var timeUpdate = updateEndTime();
      scriptPrp.setProperty(counterRange, '0');
      s.getRange(counterRange).setValue("COMPLETE");
  }
}

//Core function, takes the input and runs the other scripts to get the output
function commodityValueFW(region=60003760, commodityID) {
  var amt = [];
  console.log ("region: "+region);
  console.log ("commodityID: "+commodityID);
  amt = getFWdata(region,commodityID); //Gets the Buy and Sell values for the commodity
  console.log (amt);
  return amt;
}