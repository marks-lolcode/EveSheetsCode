/**
 * Selects the named range "FitBuy".
 * WHAT: Activates the sheet that contains the named range,
 *       unhides it if necessary, and selects the exact range.
 * WHY: Useful for quick navigation or preparing to copy.
 */
function selectFitBuy() {
  const logPrefix = '[selectFitBuy]';
  const ss = SpreadsheetApp.getActive();

  Logger.log(`${logPrefix} Starting…`);

  const name = 'FitBuy';
  const range = ss.getRangeByName(name);

  if (!range) {
    Logger.log(`${logPrefix} Named range not found: "${name}"`);
    alertSafe(`Named range not found: "${name}"`);
    return;
  }

  const sheet = range.getSheet();
  const a1 = range.getA1Notation();
  Logger.log(`${logPrefix} Found ${name} → ${sheet.getName()}!${a1}`);

  // If the sheet is hidden, show it so the selection is visible.
  try {
    if (sheet.isSheetHidden && sheet.isSheetHidden()) {
      Logger.log(`${logPrefix} Sheet is hidden; showing it.`);
      sheet.showSheet();
    }
  } catch (e) {
    // Older environments may not have isSheetHidden(); ignore safely.
    Logger.log(`${logPrefix} isSheetHidden() not available or failed: ${e}`);
    try { sheet.showSheet(); } catch (_) {}
  }

  // Activate the sheet and select the named range.
  ss.setActiveSheet(sheet);
  sheet.setActiveRange(range);

  // Force UI update.
  SpreadsheetApp.flush();

  Logger.log(`${logPrefix} Selected ${sheet.getName()}!${a1}`);
}




/**
 * Opens the dialog and instructs it to load the named range "FitBuy".
 * WHY: Clipboard writes must happen client-side; we prep data on the server.
 */
function showCopyDialog_FitBuy() {
  const t = HtmlService.createTemplateFromFile('copy_to_clipboard');
  t.initialName = 'FitBuy'; // hard-coded named range
  const html = t.evaluate().setWidth(520).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Copy Named Range: FitBuy');
}

/**
 * Returns a NAMED RANGE as TSV/CSV.
 * @param {string} rangeName - Named range to fetch (e.g., "FitBuy").
 * @param {string} format - "tsv" or "csv".
 * @return {{text:string, meta:{a1:string,rows:number,cols:number,format:string}}}
 * WHAT: Uses display values so the copied text matches what users see in the sheet.
 */
function getRangeAsTextByName(rangeName, format) {
  const ss = SpreadsheetApp.getActive();
  const name = String(rangeName || '').trim();
  if (!name) throw new Error('No named range provided.');

  const nr = ss.getRangeByName(name);
  if (!nr) {
    Logger.log(`[getRangeAsTextByName] Not found: "${name}"`);
    throw new Error(`Named range not found: "${name}"`);
  }

  const a1 = nr.getA1Notation();
  const values = nr.getDisplayValues(); // preserves visible formatting
  const rows = values.length;
  const cols = rows ? values[0].length : 0;
  const fmt = (String(format || 'tsv').toLowerCase() === 'csv') ? 'csv' : 'tsv';

  Logger.log(`[getRangeAsTextByName] ${name} → ${a1} (${rows}x${cols}) fmt=${fmt}`);

  const text = (fmt === 'tsv')
    ? values.map(r => r.join('\t')).join('\n')
    : values.map(r => r.map(csvQuote_).join(',')).join('\n');

  return { text, meta: { a1: `${name} → ${a1}`, rows, cols, format: fmt } };
}

/**
 * Quotes a single field for CSV.
 * RULES: If a field contains comma, quote, or newline, wrap in quotes and double any quotes.
 */
function csvQuote_(field) {
  const s = String(field ?? '');
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

/**
 * Quotes a single field for CSV.
 * RULES:
 * - If field contains comma, quote, or newline, wrap in quotes.
 * - Inside quotes, double any quotes.
 */
function csvQuote_(field) {
  const s = String(field ?? '');
  if (/[",\n]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}


/**
 * Creates a blank sheet or resets Existing sheet
 * @param {*} sheetName  Expected String, Name of Sheet (tab name)
 * @returns Blank Sheet
 */
function createOrClearSheet(sheetName) {
  console.time("createOrClearSdeSheet({sheetName:" + sheetName + "}})");
  if (sheetName === null || sheetName === "")
    throw "sheet name is required;";
  let activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  let workSheet = activeSheet.getSheetByName(sheetName);

  //found the Sheet, Clear it and Move on
  if (workSheet != null) {
    workSheet.clearContents();

    console.timeEnd("createOrClearSdeSheet({sheetName:" + sheetName + "}})");
    return workSheet;
  }
  //assume new sheet
  workSheet = activeSheet.insertSheet();
  workSheet.setName(sheetName);
  deleteBlankColumnsAndCollumns(workSheet);
  console.timeEnd("createOrClearSdeSheet({sheetName:" + sheetName + "}})");
  return workSheet;
}

function  deleteBlankColumnsAndCollumns(workSheet) {
  if(workSheet == null)
    throw ("workSheet not defined")
  let maxColumns = workSheet.getMaxColumns();
  let lastColumns = workSheet.getLastColumn();
  let maxRows = workSheet.getMaxColumns();
  let lastRows = workSheet.getLastColumn();
 
 if (maxRows - lastRows == 0 && maxColumns - lastColumns == 0)   return;

  const columnsReset = 2;
  const rowsReset = 2;
   
  if (lastColumns < columnsReset) { //save 2 columns on a new sheet
    lastColumns = columnsReset;
  }
  if(maxColumns - lastColumns != 0){
    workSheet.deleteColumns(lastColumns + 1, maxColumns - lastColumns);
  } 
  if (lastRows < rowsReset) { //save 2 columns on a new sheet
    lastRows = rowsReset;
  }
  if(maxRows - lastRows != 0){
    workSheet.deleteRows(lastRows + 1, maxRows - lastRows);
  }
  
}
