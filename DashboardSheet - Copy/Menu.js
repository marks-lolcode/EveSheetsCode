function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Sheet Tools')
      .addItem('Fit Compare Pull Inventory', 'pullFitInventory')
      .addToUi();
}
