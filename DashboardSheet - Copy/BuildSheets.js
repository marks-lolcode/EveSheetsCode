function buildFitCompareSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create sheets if they don't exist
  const fitCompare = getOrCreateSheet('FitCompare');
  const fitData = getOrCreateSheet('FitData');
  fitData.hideSheet(); // Hide working sheet
  
  Logger.log('=== Building FitCompare Sheet ===');
  buildFitCompareLayout(fitCompare);
  
  Logger.log('=== Building FitData Sheet ===');
  buildFitDataLayout(fitData);
  
  Logger.log('=== Creating named ranges ===');
  createNamedRanges(ss, fitCompare, fitData);
  
  Logger.log('✓ Sheet build complete.');
}

function buildFitCompareLayout(sheet) {
  sheet.clearContents();
  
  // Set column widths
  sheet.setColumnWidth(1, 120); // A
  sheet.setColumnWidths(2, 7, 100); // B-H
  sheet.setColumnWidth(11, 140); // K
  sheet.setColumnWidth(12, 100); // L
  sheet.setColumnWidth(13, 140); // M
  
  // ===== SECTION 1: TARGET FIT =====
  sheet.getRange('A1').setValue('TARGET FIT').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A2').setValue('Paste EFT format fit:').setFontStyle('italic');
  sheet.getRange('A3:H55').setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  
  // ===== SECTION 2: CURRENT FIT =====
  sheet.getRange('A58').setValue('CURRENT FIT (Optional)').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A59').setValue('Paste EFT format fit:').setFontStyle('italic');
  sheet.getRange('A60:H112').setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  
  // ===== SECTION 3: CHARACTER SELECTION =====
  sheet.getRange('K1').setValue('SELECT CHARACTERS').setFontWeight('bold').setFontSize(12);
  sheet.getRange('K2').setValue('Character Name');
  sheet.getRange('L2').setValue('Include?');
  
  // Character list (pulled from ToonName range)
  for (let row = 3; row <= 11; row++) {
    const cellRef = `K${row}`;
    sheet.getRange(cellRef).setFormula(`=IFERROR(INDEX(ToonName,${row-2}),"")`);
    
    // Add data validation for checkbox
    const checkboxRange = sheet.getRange(`L${row}`);
    const rule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    checkboxRange.setDataValidation(rule);
  }
  
  // ===== SECTION 4: FIT QUANTITY =====
  sheet.getRange('K18').setValue('Number of Fits:').setFontStyle('italic');
  sheet.getRange('L18').setValue(1).setNumberFormat('0');
  
  // ===== SECTION 5: PULL INVENTORY BUTTON =====
  sheet.getRange('K20:L20').setBackground('#4285F4').setFontColor('white');
  sheet.getRange('K20').setValue('Pull Inventory').setFontWeight('bold').setHorizontalAlignment('center');
  
  // ===== SECTION 6: LOCATION FILTER =====
  sheet.getRange('M1').setValue('LOCATION FILTER').setFontWeight('bold').setFontSize(12);
  sheet.getRange('M2').setValue('Systems with Inventory:').setFontStyle('italic');
  
  // Dynamic location list with checkboxes
  for (let row = 3; row <= 20; row++) {
    const sysCell = `M${row}`;
    sheet.getRange(sysCell).setFormula(`=IFERROR(INDEX(UniqueSystems,${row-2}),"")`);
    
    const checkboxRange = sheet.getRange(`N${row}`);
    const rule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    checkboxRange.setDataValidation(rule);
    checkboxRange.setValue(true);
  }
  
  // ===== SECTION 7: BUY LIST OUTPUT =====
  sheet.setRowHeight(115, 25);
  sheet.getRange('A115').setValue('BUY LIST').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A116').setValue('Item Name').setFontWeight('bold').setBackground('#e8e8e8');
  sheet.getRange('B116').setValue('Qty to Buy').setFontWeight('bold').setBackground('#e8e8e8');
  
  // Buy list formulas
  for (let row = 117; row <= 250; row++) {
    const rowOffset = row - 117;
    sheet.getRange(`A${row}`).setFormula(
      `=IFERROR(INDEX(FitData!$AH$2:$AH$1000,SMALL(IF(FitData!$AJ$2:$AJ$1000>0,ROW(FitData!$AJ$2:$AJ$1000)-1),${rowOffset})),"")`
    );
    sheet.getRange(`B${row}`).setFormula(
      `=IF(A${row}="","",IFERROR(INDEX(FitData!$AJ$2:$AJ$1000,SMALL(IF(FitData!$AJ$2:$AJ$1000>0,ROW(FitData!$AJ$2:$AJ$1000)-1),${rowOffset})),0))`
    );
  }
}

function buildFitDataLayout(sheet) {
  sheet.clearContents();
  
  // ===== SECTION A: SELECTED CHARACTERS =====
  sheet.getRange('A1').setValue('SelectedCharacters');
  sheet.getRange('B1').setValue('Include');
  
  for (let row = 2; row <= 10; row++) {
    sheet.getRange(`A${row}`).setFormula(`=FitCompare!K${row+1}`);
    sheet.getRange(`B${row}`).setFormula(`=FitCompare!L${row+1}`);
  }
  
  // ===== SECTION B: INVENTORY PULL =====
  sheet.getRange('S1').setValue('InventoryPull').setFontWeight('bold');
  const invHeaders = [['Character', 'TypeID', 'ItemName', 'Qty', 'LocationID', 'LocationFlag', 'LocationType', 'SystemID', 'SystemName']];
  sheet.getRange('S2:AA2').setValues(invHeaders).setBackground('#e8e8e8').setFontWeight('bold');
  
  // ===== SECTION C: SYSTEM CACHE =====
  sheet.getRange('AB1').setValue('SystemCache').setFontWeight('bold');
  const cacheHeaders = [['LocationID', 'SystemID', 'SystemName']];
  sheet.getRange('AB2:AD2').setValues(cacheHeaders).setBackground('#e8e8e8').setFontWeight('bold');
  
  // ===== SECTION D: UNIQUE SYSTEMS =====
  sheet.getRange('AF1').setValue('UniqueSystems').setFontWeight('bold');
  sheet.getRange('AF2').setValue('SystemName');
  sheet.getRange('AF3').setFormula(
    '=IFERROR(UNIQUE(FILTER(S:S,S:S<>"",S:S<>"SystemName")),"")'
  );
  
  // ===== SECTION E: GAP CALCULATION =====
  sheet.getRange('AH1').setValue('GapCalc').setFontWeight('bold');
  const gapHeaders = [['TypeID', 'ItemName', 'QtyNeeded', 'QtyOnHand', 'QtyToBuy']];
  sheet.getRange('AH2:AL2').setValues(gapHeaders).setBackground('#e8e8e8').setFontWeight('bold');
}

function createNamedRanges(ss, fitCompare, fitData) {
  function deleteNamedRange(name) {
    try {
      ss.getRangeByName(name).deleteName();
    } catch (e) {}
  }
  
  deleteNamedRange('SelectedCharacters');
  ss.setNamedRange('SelectedCharacters', fitData.getRange('A2:B10'));
  
  deleteNamedRange('UniqueSystems');
  ss.setNamedRange('UniqueSystems', fitData.getRange('AF3:AF50'));
  
  deleteNamedRange('InventoryPull');
  ss.setNamedRange('InventoryPull', fitData.getRange('S2:AA500'));
  
  deleteNamedRange('SystemCache');
  ss.setNamedRange('SystemCache', fitData.getRange('AB2:AD500'));
  
  deleteNamedRange('GapCalc');
  ss.setNamedRange('GapCalc', fitData.getRange('AH2:AL1000'));
  
  Logger.log('Named ranges created');
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}