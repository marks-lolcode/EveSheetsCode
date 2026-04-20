function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Sheet Tools')
      .addItem('Update SDE Data', 'importSDE')
      .addSeparator()
      .addItem('Update Inventory','inventoryToSheet')
      //.addItem('Update PI Make List','updateDashboardWithTop7')
      .addItem('Update Dashboard','runUpdates')
      .addSeparator()
      .addItem('Update Janice','runPriceFillBatches')
      .addItem("Refresh Prices", "refreshJaniceCommodityPrices")
      .addSeparator()
      .addItem('Run (One Shot)', 'piImport_RunOneShot')
      .addItem('Run (Chunked)',  'piImport_RunChunked')
      .addSeparator()
      .addItem("Filter Launchpads and Storage", "filterInventoryByTypeFirst")
      .addSeparator()
      .addSeparator()
      .addItem("Compare Fits", "compareAndWriteColumns")
      .addSeparator()
      .addItem("Load and Process Inventory", "updatePIFullness")
      .addSeparator()
      .addItem('Self-Test: Config & Access', 'SelfTest_ConfigAndAccess')
      .addItem('Import Fuzzworks Pricing Data', 'Import_FromExcel_PreserveFormatting')
      .addSeparator()
      .addItem('Fit Compare Pull Inventory', 'pullFitInventory')
      .addToUi();
  ui.createMenu('Maintenance')
    .addItem('Update SDE Data', 'importSDE')
    .addItem('Shrink workbook grids', 'compressWorkbook')
    .addToUi();
  //var run = runUpdates();
}
