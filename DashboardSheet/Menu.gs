function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  /**
   * Sheet Tools Menu
   * 
   * - Pull Inventory: Fetches assets from selected characters via GESI
   *   (Characters selected in FitCompare K3:L11)
   * 
   * - Populate Location Filter: After pulling inventory, consolidates unique 
   *   locations, sorts alphabetically, and populates M3:N50 with dedup'd locations
   * 
   * - Run Fit Comparison: Main fit comparison workflow
   *   1. Parse target fit from A3:H55
   *   2. Parse current fit from A60:H112 (optional)
   *   3. Compare fits to find what's needed
   *   4. Filter inventory by selected systems (M3:N50)
   *   5. Calculate gaps (what to buy)
   *   6. Populate buy list (A116:B250)
   */
  ui.createMenu('Sheet Tools')
       .addItem('--- MAINTENANCE ---', 'dummy') // Acts as a header/separator with words
        .addItem('01 - Clear Inventory Cache', 'clearInventoryCache')      
       .addSeparator()
       .addItem('--- PERFORM COMPARE ---', 'dummy') // Acts as a header/separator with words
       .addSeparator()       
        .addItem('02 - Pull Inventory (Selected Characters)', 'pullFitInventory')
        .addItem('03 - Populate Location Filter', 'populateLocationFilter')
        .addItem('04 - (optional) Clear Location Filter', 'clearLocationFilter')
        .addItem('05 - Run Fit Comparison', 'runFitComparison')
        .addItem('06 - Fit Compare Pull Inventory', 'pullFitInventory')
        .addItem('07 - Populate Removal/Additions', 'populateRemovalAdditions')
      .addToUi();
}