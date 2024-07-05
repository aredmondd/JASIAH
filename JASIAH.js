// CONSTANTS (things that will not change)
const NOW = new Date();
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const UI = SpreadsheetApp.getUi();

// these are the names of the spreadsheets. referenced throughout the code. can be changed if needed.
const SHOPIFY_INVENTORY = "shopify_inventory";
const JACOBS_STATS = "Jacob's Stats";
const SKU_LIST_SHEET = "SKU LISTS";

// other helpful variables
let currentDayOfTheWeek = getCurrentDayOfWeek();
let currentDaySheet; // something like... "MONDAY_INVENTORY"
let currentDaySheetExport; // something like... "MONDAY_INVENTORY_EXPORT"
let doingTodaysInventory = true; 
let dayMap = {
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
 * @params: None
 * @returns: None
 * 
 * main driver. duplicates sheet, deletes any non-valid SKUs, and cleans everything up.
 * 
 */
function execute() {
  // check for any SKUs that were scanned, but not inside today's SKU list
  let shouldContinue = doubleCheckSKUs();

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
 * @params: None
 * @returns: None
 * 
 * duplicates the inventory sheet and formats it so that we can import it back into Shopify.
 * 
 */
function exportSheet() {
  // get all the sheets
  let allSheets = SPREADSHEET.getSheets();

  // find the day of the week we are currently taking inventory for
  for (let i = 0; i < allSheets.length; i++) {
    let currentSheet = allSheets[i];
    if (currentSheet.getName().endsWith("_STOCKY")) {
      currentDayOfTheWeek = currentSheet.getName().split('_')[0];
    }
  }

  // duplicate the current day sheet, rename to ${currentDaySheet}_EXPORT
  duplicateSheet(`${currentDayOfTheWeek}_INVENTORY`, `${currentDayOfTheWeek}_INVENTORY_EXPORT`);

  let exportSheet = SPREADSHEET.getSheetByName(`${currentDayOfTheWeek}_INVENTORY_EXPORT`);

  // copy column N, and paste it into column P.
  let source = exportSheet.getRange("N:N");
  let columnN = source.getValues();
  let target = exportSheet.getRange('P:P');
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
 * @params: None
 * @returns: None
 * 
 * duplicate the shopify inventory sheet & rename it to today's day of the week _inventory
 * 
 */
function start() {
  duplicateSheet(SHOPIFY_INVENTORY,`${currentDayOfTheWeek}_INVENTORY`);
  currentDaySheet = SPREADSHEET.getSheetByName(`${currentDayOfTheWeek}_INVENTORY`);
}


/**
 * 
 * filterSKUs()
 * 
 * @param: skuList, array = optional list of SKUs that we are checking against.
 * @returns: None
 * 
 * get the SKUs we want based on the day and then delete any rows that are not in that list
 *
 */
function filterSKUs(skuList) {
  let currentSKUs;
  let rowsToKeep = [];
  
  if (!skuList) { //if the user did not provide a skuList, get it based on today's date
    currentSKUs = getSKUList(dayToNumber(currentDayOfTheWeek));
  }
  else {
    currentSKUs = skuList; 
  }

  // get all of the data from the sheet
  let sheetData = currentDaySheet.getDataRange().getValues();

  // make a list of all the SKUs we want to keep
  rowsToKeep.push(sheetData[0]); //keep the first row

  // go row by row and check which rows should be kept
  for (let i = 1; i < sheetData.length; i++) {
    let currentRow = sheetData[i];
    let sku = currentRow[8]; //8 is the column with the SKU which is what we are filtering for

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
 * @params: None
 * @returns: None
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
  let sheet = SPREADSHEET.getSheetByName(oldSheetName);
 
  if (sheet) {
    let copiedSheet = sheet.copyTo(SPREADSHEET);
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
 * @returns: skuList, Array = list of SKUs for said day
 * 
 * get the list of 'valid' SKUs so that we can delete any SKUs not in this list later
 * 
 */
function getSKUList(column) {
  // Access the sheet with SKUs for each day
  let sheet = SPREADSHEET.getSheetByName(SKU_LIST_SHEET);

  if (!sheet) {
    throw new Error ("There is no sheet named SKU LISTS. Cannot filter by SKU.")
  }

  let skuList = getNonEmptyValues(sheet, column);

  // Remove the first item in the array (will be MONDAY/TUESDAY/WEDNESDAY/THURSDAY/FRIDAY)
  skuList.shift();

  return skuList;
}


/**
* insertFormulas()
*
* @params: None
* @returns: None
* 
* insert stocky lookup formula and sum formula
* 
*/
function insertFormulas() {
  let lastRow = currentDaySheet.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    let stockyTotalsCell = currentDaySheet.getRange("M" + i);
    let sumCell = currentDaySheet.getRange("N" + i);

    let stockyFormula = `=IFNA(VLOOKUP(I${i}, ${currentDayOfTheWeek}_STOCKY!C:P, 13, FALSE) * 25, 0 )`;
    stockyTotalsCell.setFormula(stockyFormula);

    let sumFormula = `=SUM(L${i},M${i})`;
    sumCell.setFormula(sumFormula);
  }
}


/**
 * 
 * dayToNumber()
 * 
 * @params: None
 * @returns: int =  number noted below
 * 
 * convert a day of the week to a number
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
 * @params: None
 * @returns: dayOfWeek, string = string of day of the week
 * 
 * gets the current day of the week based on time
 * 
 */
function getCurrentDayOfWeek() {
  // Get the correct time zone
  let timeZone = Session.getScriptTimeZone();
  
  // Format the date to get the day of the week in uppercase
  let dayOfWeek = Utilities.formatDate(NOW, timeZone, 'EEEE').toUpperCase();
  
  return dayOfWeek;
}


/**
 * 
 * hideSpecifiedColumns()
 * 
 * @param: range, string = the range of columns you want to hide, A:H
 * @returns: None
 * 
 * custom hide columns method to make things look nicer :)
 * 
 */
function hideSpecifiedColumns(range) {
  let hiddenRange = currentDaySheet.getRange(range);
  currentDaySheet.hideColumn(hiddenRange);
}


/**
 * 
 * sortSheetBySKU()
 * 
 * @params: None
 * @returns: None
 * 
 * sort the newly made sheet alphabetically for easy access to live counting hats
 * 
 */
function sortSheetBySKU() {
  let range = currentDaySheet.getRange(2, 1, currentDaySheet.getLastRow() - 1, currentDaySheet.getLastColumn());
  range.sort({column: 9, ascending: true});
}


/**
 * 
 * conditionalFormatting()
 * 
 * @params: None
 * @returns: None
 * 
 * applies conditional formats for two parameters:
 * 1. if shopify inventory < our inventory mark it yellow
 * 2. if shopify inventory < our inventory by 25 or more mark it red
 * 
 */
function conditionalFormatting() {
  let range = currentDaySheet.getRange("O2:O");
  range.clearFormat();

  let yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=O2 < N2")
    .setBackground("#FFFF00")
    .setRanges([range])
    .build();
  
  let redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=(N2 - O2) >= 25")
    .setBackground("#FF0000")
    .setRanges([range])
    .build();

  let rules = currentDaySheet.getConditionalFormatRules();
  rules.push(redRule);
  rules.push(yellowRule);
  currentDaySheet.setConditionalFormatRules(rules);
}


/**
 * 
 * onOpen()
 * 
 * @params: None
 * @returns: None
 * 
 * adds UI element to run execute() from within google sheets
 * 
 */
function onOpen() {
  UI.createMenu('JASIAH')
      .addItem("Take today's inventory", 'execute')
      .addItem("Take multiple days of inventory", "takeMultipleDaysOfInventory")
      .addItem('Take custom day inventory (M/T/W/TH/F)', 'takeCustomInventory')
      .addItem('Take inventory of everything', 'doEverything')
      .addItem('Export inventory', 'exportSheet')
      .addItem('Clear sheet', 'clear')
      .addToUi();
}


/**
 * 
 * takeCustomInventory()
 * 
 * @params: None
 * @returns: None
 * 
 * takes inventory based on a day that the user inputs inside google sheets
 * 
 */
function takeCustomInventory() {
  doingTodaysInventory = false;

  let response = UI.prompt("Which day of the week would you like to take inventory for? \n (e.g monday: m, mon, or monday)");

  // Check if the user clicked OK
  if (response.getSelectedButton() == UI.Button.OK) {
    let userInput = response.getResponseText().toLowerCase();

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
 * @params: None
 * @returns: None
 * 
 * deletes any non-standard sheets (stocky, shopify, exports, etc)
 * 
 */
function clear() {
  let response = UI.alert("Are you sure you want to delete any non-standard sheets? (You can probably undo this)", UI.ButtonSet.YES_NO);

  if (response == UI.Button.YES) {
    let allSheets = SPREADSHEET.getSheets();

    allSheets.forEach(sheet => {
      let name = sheet.getName();
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
 * @returns: flatValues, Array = list of non-empty values
 * 
 * 
 */
function getNonEmptyValues(sheet, column) {
  let lastRow = sheet.getRange(sheet.getMaxRows(), column).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  let range = sheet.getRange(1, column, lastRow);
  let values = range.getValues();

  let flatValues = values.map(function(row) {
    return row[0];
  })

  return flatValues;
}


/**
 * 
 * getTodayScore()
 * 
 * @params: None
 * @returns: None
 * 
 * inserts inventory score (see docs for formula) into the sheet
 * 
 */
function getTodayScore() {
  // open the jacob stats sheet
  let jacobStats = SPREADSHEET.getSheetByName(JACOBS_STATS);

  if (!jacobStats) {
    throw new Error ("No sheet named 'Jacob's Stats'");
  }

  // insert a row after row 1
  jacobStats.insertRowBefore(2);

  // in A2, insert today's date
  let formattedDate = Utilities.formatDate(NOW, Session.getScriptTimeZone(), "MM/dd/yy");
  jacobStats.getRange("A2").setValue(formattedDate);

  // in B2, insert today's DOTW
  jacobStats.getRange("B2").setValue(currentDayOfTheWeek);

  // in C2, insert the custom formula based on the day of the week
  jacobStats.getRange("C2").setFormula(`=CONCATENATE(ROUND((SUM(${currentDayOfTheWeek}_INVENTORY!N:N) / SUM(${currentDayOfTheWeek}_INVENTORY!O:O)) * 100, 2), "%")`);
  let formulaResultValue = jacobStats.getRange("C2").getValue();
  jacobStats.getRange("C2").setValue(formulaResultValue);
}

/**
 * 
 * getAllSKUs()
 * 
 * @params: None
 * @returns: allSKUs, Array = list of every SKU in the SKU list
 * 
 * creates a list of every SKU in the SKU LIST sheet when taking all of inventory
 * 
 * 
 */
function getAllSKUs() {
  // get every SKU from the the SKU list
  let skuListSheet = SPREADSHEET.getSheetByName(SKU_LIST_SHEET);

  if (!skuListSheet) {
    throw new Error ("No sheet named 'SKU LISTS'");
  }

  // Get all of the SKUs from all the columns
  let allSKUS = [];

  for (let i = 1; i <= 5; i++) {
    let currentRange = getNonEmptyValues(skuListSheet, i);
    currentRange.shift();
    allSKUS.push(currentRange);
  }

  allSKUS = allSKUS.flat();

  return allSKUS;
}


/**
 * 
 * doEverything()
 * 
 * @params: None
 * @returns: None
 * 
 * takes inventory for every SKU in SKU List
 * 
 */
function doEverything() {
  let allSKUs = getAllSKUs();

  let shouldContinue = doubleCheckSKUs(allSKUs);

  if (shouldContinue == true) {
    start();

    filterSKUs(allSKUs);

    prettify(); 
  }
}


/**
 * 
 * doubleCheckSKUs()
 * 
 * @params: listOfSKUs, Array = optional parameter if we aren't getting SKU list based on current day
 * @returns: boolean = If user wants to continue or not
 * 
 * alert the user of all the SKUs that were scanned but not inside SKU list for the day
 * if listOfSKUs is provided, it will use that list instead of the day's SKU list
 * 
 */
function doubleCheckSKUs(listOfSKUs = null) {
  let stockySheet = SPREADSHEET.getSheetByName(`${currentDayOfTheWeek}_STOCKY`);

  if (!stockySheet) {
    throw new Error ("JASIAH cannot find the imported Stocky data.");
  }

  let scannedSKUs = getNonEmptyValues(stockySheet, 3);

  scannedSKUs.shift(); // remove the first item (will be a header)

  // if the listOfSKUs was provided (we are multiple days of inventory or all of inventory), use that. If not, get today's list of SKUs
  let comparisonSKUs = listOfSKUs ? listOfSKUs : getSKUList(dayToNumber(currentDayOfTheWeek));

  let missingSKUs = [];
  for (let i = 0; i < scannedSKUs.length; i++) {
    let currentSKU = scannedSKUs[i];
    if (comparisonSKUs.indexOf(currentSKU) === -1) {
      missingSKUs.push(currentSKU);
    }
  }

  if (missingSKUs.length == 0) {
    UI.alert("All scanned SKUs are in the SKU list :)");
    return true;
  }
  else {
    let response = UI.alert(`Below are all SKUs that were scanned today, but not inside the SKU list: \n\n\n ${missingSKUs} \n\n\n Do you want to continue anyway?`, UI.ButtonSet.YES_NO);
    if (response == UI.Button.YES) {
      return true;
    }
    else {
      return false;
    }
  }
}

/**
 * 
 * takeMultipleDaysOfInventory()
 * 
 * @params: None
 * @returns: None
 * 
 * create custom HTML dialogue box to input which days of the week we are taking inventory for
 * 
 */
function takeMultipleDaysOfInventory() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('selectMultipleDays')
      .setWidth(600)
      .setHeight(125);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Which Days?');
}


/**
 * 
 * handleSelectedOptions()
 * 
 * @params: None
 * @returns: None
 * 
 * Gather list of SKUs based on what days were input
 * 
 */
function handleSelectedOptions(days) {
  if (days.length > 0) {
    let bigSKUList = [];

    for (let i = 0; i < days.length; i++) {
      let currentNumber = dayToNumber(days[i])
      let currentSKUList = getSKUList(currentNumber);
      bigSKUList.push(currentSKUList);
    }

    bigSKUList = bigSKUList.flat();

    // check for any SKUs that were scanned, but not inside today's SKU list
    let shouldContinue = doubleCheckSKUs(bigSKUList);

    if (shouldContinue == true) {
      start();
      filterSKUs(bigSKUList);
      prettify();
    }
  } else {
    UI.alert('No options selected...');
  }
}


