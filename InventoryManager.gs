function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Inventory Manager")
    .addItem("Open Dashboard", "showDashboardSidebar")
    .addItem("Add New Item", "showAddItemSidebar")
    .addItem("Scanner Sort", "scannerSort")
    .addToUi();
}

function showDashboardSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Dashboard")
    .setTitle("Inventory Dashboard");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAddItemSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("AddItemSidebar")
    .setTitle("Add New Item");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getCategories() {
  return [
    "Alcohol Free - Bottles", "Bottles", "Cans", "Cocktail Ingredients", "Consumables",
    "Draught", "Mixers", "Post Mix - Vimto", "Soft Drink - Juice",
    "O'Donnell Moonshine - 700ml", "O'Donnell Moonshine 50 ml - Single Serve",
    "O'Donnell Moonshine - 20 ml - Single Serve", "Spirits", "Wines"
  ];
}

function addItemToSheet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory");
  if (!sheet) return "Inventory sheet not found.";

  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1).setValue(data.category);
  sheet.getRange(nextRow, 2).setValue(data.item);
  sheet.getRange(nextRow, 3).setValue(data.holdingStock);
  sheet.getRange(nextRow, 4).setValue(data.current);
  sheet.getRange(nextRow, 5).setValue(data.need);

  return `Added: ${data.item} (${data.category}) with ${data.holdingStock} units`;
}

function duplicateInventorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventory");
  if (!sheet) return;

  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  const name = `Stocktake - ${timestamp}`;
  sheet.copyTo(ss).setName(name);
  SpreadsheetApp.getUi().alert(`New stocktake sheet "${name}" created.`);
}

function createScannerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd-HHmm");
  const name = `Scanner - ${timestamp}`;
  const sheet = ss.insertSheet(name);
  sheet.getRange("A1").setValue("Barcode");
  return sheet;
}

function matchAndCountBarcodes(scannerSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productSheet = ss.getSheetByName("Product Data");
  if (!productSheet) {
    SpreadsheetApp.getUi().alert("Product Data sheet not found.");
    return;
  }

  const productBarcodes = productSheet.getRange("A2:A" + productSheet.getLastRow()).getValues().flat();
  const scannedBarcodes = scannerSheet.getRange("A2:A" + scannerSheet.getLastRow()).getValues().flat();
  const counts = {};

  scannedBarcodes.forEach(code => {
    if (!code) return;
    if (productBarcodes.includes(code)) {
      counts[code] = (counts[code] || 0) + 1;
    }
  });

  const output = scannedBarcodes.map(code => (code && counts[code] ? counts[code] : ""));
  scannerSheet.getRange(2, 4, output.length, 1).setValues(output.map(v => [v]));
  SpreadsheetApp.getUi().alert("Scanner data processed and matched counts written to Column D.");
}

function scannerSort() {
  const scannerSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  matchAndCountBarcodes(scannerSheet);
}

function countUniqueReferences() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange("A:A").getValues().flat().filter(String);
  const counts = {};

  data.forEach(ref => {
    counts[ref] = (counts[ref] || 0) + 1;
  });

  const uniqueRefs = Object.keys(counts);
  const outputB = uniqueRefs.map(ref => [ref]);
  const outputC = uniqueRefs.map(ref => [counts[ref]]);

  sheet.getRange("B:C").clearContent();
  sheet.getRange("B1").setValue("Unique Reference");
  sheet.getRange("C1").setValue("Count");
  sheet.getRange(2, 2, outputB.length, 1).setValues(outputB);
  sheet.getRange(2, 3, outputC.length, 1).setValues(outputC);
}
