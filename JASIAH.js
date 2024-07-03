// CONSTANTS (things that will not change)
const NOW = new Date();
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const UI = SpreadsheetApp.getUi();

// these are the names of the spreadsheets. referenced throughout the code. can be changed if needed.
const SHOPIFY_INVENTORY = "shopify_inventory";
const JACOBS_STATS = "Jacob's Stats";
const SKU_LIST_SHEET = "SKU LISTS";

// other helpful variables
var currentDayOfTheWeek = getCurrentDayOfWeek();
var currentDaySheet; // "MONDAY_INVENTORY"
var currentDaySheetExport; // "MONDAY_INVENTORY_EXPORT"
var doingTodaysInventory = true;
var dayMap = {
  'm': 'MONDAY', 'mon': 'MONDAY', 'monday': 'MONDAY',
  't': 'TUESDAY', 'tue': 'TUESDAY', 'tuesday': 'TUESDAY',
  'w': 'WEDNESDAY', 'wed': 'WEDNESDAY', 'wednesday': 'WEDNESDAY',
  'th': 'THURSDAY', 'thu': 'THURSDAY', 'thursday': 'THURSDAY',
  'f': 'FRIDAY', 'fri': 'FRIDAY', 'friday': 'FRIDAY'
}


/**
 * 
 * execute()
 * 
 * main driver. duplicates sheet, deletes any non-valid SKUs, and cleans everything up.
 * 
 */
function execute() {
  // check for any SKUs that were scanned, but not inside today's SKU list
  var shouldContinue = doubleCheckSKUs();

  if (shouldContinue == true) {
    start();
    filterSKUs();
    prettify();
  }
}

/**
 * 
 * exportSheet()
 * 
 * duplicates the inventory sheet and formats it so that we can import it back into Shopify.
 * 
 */
function exportSheet() {
  // get all the sheets
  var allSheets = SPREADSHEET.getSheets();

  // find the day of the week we are currently taking inventory for
  for (var i = 0; i < allSheets.length; i++) {
    var currentSheet = allSheets[i];
    if (currentSheet.getName().endsWith("_STOCKY")) {
      currentDayOfTheWeek = currentSheet.getName().split('_')[0];
    }
  }

  // duplicate the current day sheet, rename to ${currentDaySheet}_EXPORT
  duplicateSheet(`${currentDayOfTheWeek}_INVENTORY`, `${currentDayOfTheWeek}_INVENTORY_EXPORT`);

  var exportSheet = SPREADSHEET.getSheetByName(`${currentDayOfTheWeek}_INVENTORY_EXPORT`);

  // copy column N, and paste it into column P.
  var source = exportSheet.getRange("N:N");
  var columnN = source.getValues();
  var target = exportSheet.getRange('P:P');
  target.setValues(columnN);

  // delete column L:O
  exportSheet.deleteColumn(15);
  exportSheet.deleteColumn(14);
  exportSheet.deleteColumn(13); 
  exportSheet.deleteColumn(12); 

  // rename cell L1 to 'Main Office'
  exportSheet.getRange("L1").setValue("Main Office");

  // save the result of the formula for later if doing today's inventory
  if (doingTodaysInventory == true) {
    getTodayScore();
  }
}


/**
 * 
 * start()
 * 
 * duplicate the shopify inventory sheet & rename it to today's day of the week _inventory
 */
function start() {
  duplicateSheet(SHOPIFY_INVENTORY,`${currentDayOfTheWeek}_INVENTORY`);
  currentDaySheet = SPREADSHEET.getSheetByName(`${currentDayOfTheWeek}_INVENTORY`);
}


/**
 * 
 * filterSKUs()
 * 
 * get the SKUs we want based on the day and then delete any rows that are not in that list
 * @param: skuList, array = optional list of SKUs that we are checking against.
 * 
 */
function filterSKUs(skuList) {
  var currentSKUs;
  var rowsToKeep = [];
  
  if (!skuList) { //if the user did not provide a skuList, get it based on today's date
    currentSKUs = getSKUList(dayToNumber(currentDayOfTheWeek));
  }
  else {
    currentSKUs = skuList; 
  }

  // get all of the data from the sheet
  var sheetData = currentDaySheet.getDataRange().getValues();

  // make a list of all the SKUs we want to keep
  rowsToKeep.push(sheetData[0]); //keep the first row

  // go row by row and check which rows should be kept
  for (var i = 1; i < sheetData.length; i++) {
    var currentRow = sheetData[i];
    var sku = currentRow[8]; //8 is the column with the SKU which is what we are filtering for

    if (currentSKUs.includes(sku)) {
      rowsToKeep.push(currentRow);
    }
  }

  // clear the sheet
  currentDaySheet.clear();

  // add all the rows back that we want to keep
  currentDaySheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
}


/**
 * 
 * prettify()
 * 
 * add live hats, stocky total, & totals columns
 * hide columns, remove blank lines, fix formulas, sort the sheet, and apply conditional formatting
 * 
 */
function prettify() {
  // insert 3 columns for live hats, stocky totals, & total
  currentDaySheet.insertColumns(12 ,3);

  // give column headers where needed
  currentDaySheet.getRange("L1").setValue("Live Hats").setFontWeight("bold");
  currentDaySheet.getRange("M1").setValue("Stocky Total").setFontWeight("bold");
  currentDaySheet.getRange("N1").setValue("Total").setFontWeight("bold");

  // Hide columns
  hideSpecifiedColumns("J:K");
  hideSpecifiedColumns("A:H");

  // update the formula to contain the correct date
  insertFormulas();

  // sort the sheet alphabetically by SKU
  sortSheetBySKU();

  // apply conditional formatting
  conditionalFormatting();
}


/**
 * duplicateSheet()
 * 
 * @param: oldSheetName, String = what the old sheet (one to be duplicated) is called
 * @param: newSheetName, String = what the new sheet will be called. dynamically made based on day.
 * 
 * Duplicates old sheet, and renames it to newSheetName
 * 
 */
function duplicateSheet(oldSheetName, newSheetName) {
  var sheet = SPREADSHEET.getSheetByName(oldSheetName);
 
  if (sheet) {
    var copiedSheet = sheet.copyTo(SPREADSHEET);
    copiedSheet.setName(newSheetName);
  }
  else {
    throw new Error(`Cannot find sheet with the name ${oldSheetName}.`);
  }
}


/**
 * 
 * getSKUList()
 * 
 * @param: column, int = the column that we will be pulling SKUs for (monday = 1, tuesday = 2...)
 * @return: list of SKUs for said day
 * 
 * get the list of 'valid' SKUs so that we can delete any SKUs not in this list later
 * 
 */
function getSKUList(column) {
  // Access the sheet with SKUs for each day
  var sheet = SPREADSHEET.getSheetByName(SKU_LIST_SHEET);

  if (!sheet) {
    throw new Error ("There is no sheet named SKU LISTS. Cannot filter by SKU.")
  }

  var skuList = getNonEmptyValues(sheet, column);

  // Remove the first item in the array (will be MONDAY/TUESDAY/WEDNESDAY/THURSDAY/FRIDAY)
  skuList.shift();

  return skuList;
}


/**
* insertFormulas()
*
* insert stocky lookup formula and sum formula
*/
function insertFormulas() {
  var lastRow = currentDaySheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var stockyTotalsCell = currentDaySheet.getRange("M" + i);
    var sumCell = currentDaySheet.getRange("N" + i);

    var stockyFormula = `=IFNA(VLOOKUP(I${i}, ${currentDayOfTheWeek}_STOCKY!C:P, 13, FALSE) * 25, 0 )`;
    stockyTotalsCell.setFormula(stockyFormula);

    var sumFormula = `=SUM(L${i},M${i})`;
    sumCell.setFormula(sumFormula);
  }
}


/**
 * 
 * dayToNumber()
 * 
 * convert a day of the week to a number
 * @return: int number noted below
 * monday = 1, tuesday = 2, wednesday = 3, thursday = 4, friday = 5
 * 
 */
function dayToNumber(day) {
  switch (day) {
    case 'MONDAY':
      return 1;
    case 'TUESDAY':
      return 2;
    case 'WEDNESDAY':
      return 3;
    case 'THURSDAY':
      return 4;
    case 'FRIDAY':
      return 5;
  }
}


/**
 * 
 * getCurrentDayOfWeek()
 * 
 * @return: string of day of the week
 * 
 * gets the current day of the week based on time
 * 
 * 
 */
function getCurrentDayOfWeek() {
  // Get the correct time zone
  var timeZone = Session.getScriptTimeZone();
  
  // Format the date to get the day of the week in uppercase
  var dayOfWeek = Utilities.formatDate(NOW, timeZone, 'EEEE').toUpperCase();
  
  return dayOfWeek;
}


/**
 * 
 * hideSpecifiedColumns()
 * 
 * @param: range, string = the range of columns you want to hide, A:H
 * 
 * custom hide columns method to make things look nicer :)
 * 
 */
function hideSpecifiedColumns(range) {
  var hiddenRange = currentDaySheet.getRange(range);
  currentDaySheet.hideColumn(hiddenRange);
}


/**
 * 
 * sortSheetBySKU()
 * 
 * sort the newly made sheet alphabetically for easy access to live counting hats
 * 
 */
function sortSheetBySKU() {
  var range = currentDaySheet.getRange(2, 1, currentDaySheet.getLastRow() - 1, currentDaySheet.getLastColumn());
  range.sort({column: 9, ascending: true});
}


/**
 * 
 * conditionalFormatting()
 * 
 * applies conditional formats for two parameters:
 * 1. if shopify inventory < our inventory mark it yellow
 * 2. if shopify inventory < our inventory by 25 or more mark it red
 * 
 */
function conditionalFormatting() {
  var range = currentDaySheet.getRange("O2:O");
  range.clearFormat();

  var yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=O2 < N2")
    .setBackground("#FFFF00")
    .setRanges([range])
    .build();
  
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=(N2 - O2) >= 25")
    .setBackground("#FF0000")
    .setRanges([range])
    .build();

  var rules = currentDaySheet.getConditionalFormatRules();
  rules.push(redRule);
  rules.push(yellowRule);
  currentDaySheet.setConditionalFormatRules(rules);
}


/**
 * 
 * onOpen()
 * 
 * adds UI element to run execute() from within google sheets
 * 
 */
function onOpen() {
  // Create a custom menu
  UI.createMenu('JASIAH')
      .addItem("Take today's inventory", 'execute')
      .addItem("Take multiple days of inventory", "takeMultipleDaysOfInventory")
      .addItem('Export inventory', 'exportSheet')
      .addItem('Take custom day inventory (M/T/W/TH/F)', 'takeCustomInventory')
      .addItem('Take inventory of everything', 'doEverything')
      .addItem('Clear sheet', 'clear')
      .addToUi();
}




/**
 * 
 * takeCustomInventory()
 * 
 * takes inventory based on a day that the user inputs inside google sheets
 * 
 */
function takeCustomInventory() {
  doingTodaysInventory = false;

  var response = UI.prompt("Which day of the week would you like to take inventory for? \n (e.g monday: m, mon, or monday)");

  // Check if the user clicked OK
  if (response.getSelectedButton() == UI.Button.OK) {
    var userInput = response.getResponseText().toLowerCase();

    // if the input is inside the dictionary, move forward
    if(dayMap[userInput]) {
      currentDayOfTheWeek = dayMap[userInput];
      execute();
    }
    else {
      throw new Error (`There is no such day as ${response}... check for typos?`);
    }
  }
}


/**
 * 
 * clear()
 * 
 * deletes any non-standard sheets (stocky, shopify, exports, etc)
 * 
 */
function clear() {
  var response = UI.alert("Are you sure you want to delete any non-standard sheets? (You can probably undo this)", UI.ButtonSet.YES_NO);

  if (response == UI.Button.YES) {
    var allSheets = SPREADSHEET.getSheets();

    allSheets.forEach(sheet => {
      var name = sheet.getName();
      if (name.endsWith("_STOCKY") || name.endsWith("_INVENTORY") || name.endsWith("_EXPORT") || name === SHOPIFY_INVENTORY) {
        SPREADSHEET.deleteSheet(sheet);
      }
    });
  }
}


/**
 * 
 * getNonEmtpyValues()
 * 
 * @param: sheet, string = the sheet you want to get the values from
 * @param: column, int = the column you want to get all values for
 * @return: list of non-empty values
 * 
 * 
 */
function getNonEmptyValues(sheet, column) {
  var lastRow = sheet.getRange(sheet.getMaxRows(), column).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  var range = sheet.getRange(1, column, lastRow);
  var values = range.getValues();

  var flatValues = values.map(function(row) {
    return row[0];
  })

  return flatValues;
}


/**
 * 
 * getTodayScore()
 * 
 * inserts inventory score (see docs for formula) into the sheet
 * 
 */
function getTodayScore() {
  // open the jacob stats sheet
  var jacobStats = SPREADSHEET.getSheetByName(JACOBS_STATS);

  if (!jacobStats) {
    throw new Error ("No sheet named 'Jacob's Stats'");
  }

  // insert a row after row 1
  jacobStats.insertRowBefore(2);

  // in A2, insert today's date
  var formattedDate = Utilities.formatDate(NOW, Session.getScriptTimeZone(), "MM/dd/yy");
  jacobStats.getRange("A2").setValue(formattedDate);

  // in B2, insert today's DOTW
  jacobStats.getRange("B2").setValue(currentDayOfTheWeek);

  // in C2, insert the custom formula based on the day of the week
  jacobStats.getRange("C2").setFormula(`=CONCATENATE(ROUND((SUM(${currentDayOfTheWeek}_INVENTORY!N:N) / SUM(${currentDayOfTheWeek}_INVENTORY!O:O)) * 100, 2), "%")`);
  var formulaResultValue = jacobStats.getRange("C2").getValue();
  jacobStats.getRange("C2").setValue(formulaResultValue);
}

/**
 * 
 * getAllSKUs()
 * 
 * creates a list of every SKU in the SKU LIST sheet when taking all of inventory
 * 
 * @return: the list of every SKU in the SKU list
 * 
 */
function getAllSKUs() {
  // get every SKU from the the SKU list
  var skuListSheet = SPREADSHEET.getSheetByName(SKU_LIST_SHEET);

  if (!skuListSheet) {
    throw new Error ("No sheet named 'SKU LISTS'");
  }

  // Get all of the SKUs from all the columns
  var allSKUS = [];

  for (var i = 1; i <= 5; i++) {
    var currentRange = getNonEmptyValues(skuListSheet, i);
    currentRange.shift();
    allSKUS.push(currentRange);
  }

  var allSKUS = allSKUS.flat();

  return allSKUS;
}


/**
 * 
 * doEverything()
 * 
 * takes inventory for every SKU in SKU List
 * 
 */
function doEverything() {
  var allSKUs = getAllSKUs();

  var shouldContinue = doubleCheckSKUs(allSKUs);

  if (shouldContinue == true) {
    start();

    filterSKUs(allSKUs);

    prettify(); 
  }
}


/**
 * 
 * doubleCheckSKUs(listOfSKUs = null)
 * 
 * alert the user of all the SKUs that were scanned but not inside SKU list for the day
 * if allSKUs is provided, it will use that list instead of the day's SKU list
 * 
 */
function doubleCheckSKUs(listOfSKUs = null) {
  var stockySheet = SPREADSHEET.getSheetByName(`${currentDayOfTheWeek}_STOCKY`);

  if (!stockySheet) {
    throw new Error ("JASIAH cannot find the imported Stocky data.");
  }

  var scannedSKUs = getNonEmptyValues(stockySheet, 3);

  scannedSKUs.shift(); // remove the first item (will be a header)

  // if the listOfSKUs was provided (we are multiple days of inventory or all of inventory), use that. If not, get today's list of SKUs
  var comparisonSKUs = listOfSKUs ? listOfSKUs : getSKUList(dayToNumber(currentDayOfTheWeek));

  var missingSKUs = [];
  for (var i = 0; i < scannedSKUs.length; i++) {
    var currentSKU = scannedSKUs[i];
    if (comparisonSKUs.indexOf(currentSKU) === -1) {
      missingSKUs.push(currentSKU);
    }
  }

  if (missingSKUs.length == 0) {
    UI.alert("All scanned SKUs are in the SKU list :)");
    return true;
  }
  else {
    var response = UI.alert(`Below are all SKUs that were scanned today, but not inside the SKU list: \n\n\n ${missingSKUs} \n\n\n Do you want to continue anyway?`, UI.ButtonSet.YES_NO);
    if (response == UI.Button.YES) {
      return true;
    }
    else {
      return false;
    }
  }
}

function takeMultipleDaysOfInventory() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('selectMultipleDays')
      .setWidth(600)
      .setHeight(125);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Which Days?');
}

function handleSelectedOptions(days) {
  if (days.length > 0) {
    var bigSKUList = [];

    for (var i = 0; i < days.length; i++) {
      var currentNumber = dayToNumber(days[i])
      var currentSKUList = getSKUList(currentNumber);
      bigSKUList.push(currentSKUList);
    }

    var bigSKUList = bigSKUList.flat();

    // check for any SKUs that were scanned, but not inside today's SKU list
    var shouldContinue = doubleCheckSKUs(bigSKUList);

    if (shouldContinue == true) {
      start();
      filterSKUs(bigSKUList);
      prettify();
    }
  } else {
    UI.alert('No options selected...');
  }
}
