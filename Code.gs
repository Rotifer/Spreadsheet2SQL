/*
Create a new spreadsheet, add a sheet to it with user-specified name and write the SQL statements to it.
*/
function writeSqlToSheet(createTableSql, insertSqls, spreadsheetName, sheetName) {
  var newSpreadsheet = SpreadsheetApp.create(spreadsheetName),
      newSheet = newSpreadsheet.insertSheet(sheetName),
      rowCount = insertSqls.length,
      twoDArrayToWrite = [],
      lastCellAddr,
      insertSqlOutputRngAddr,
      insertSqlOutputRng,
      sheet1;
  //Delete default sheet1 from the new spreadsheet
  sheet1 = newSpreadsheet.getSheetByName('Sheet1');
  newSpreadsheet.deleteSheet(sheet1);
  //Write the "CREATE TABLE statement to cell A1
  newSheet.getRange('A1').setValue(createTableSql);
  // map returns a one-dimensional array
  twoDArrayToWrite = insertSqls.map(function(sql) { return [sql]; });
  lastCellAddr = newSheet.getRange('A2').offset(twoDArrayToWrite.length - 1, 0).getA1Notation();
  // Define the range to which the "INSERT" SQL statements are to be written.
  insertSqlOutputRngAddr = 'A2:' + lastCellAddr;
  insertSqlOutputRng = newSheet.getRange(insertSqlOutputRngAddr);
  //Write "INSERTs" to range defined above
  insertSqlOutputRng.setValues(twoDArrayToWrite);
}

/*
Takes the options passed from the input form and uses them to call methods defined in the "SPREADSHEET2SQL"
module. Gets the data back from the module methods and calls function "writeSqlToSheet()" to write
the output to a spreadsheet.
*/
function runMod(tableName, rdbmsName, newSpreadsheetName, newSheetName) {
  var rng = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PasteTableDataHere').getDataRange();
  var mod = SPREADSHEET2SQL;
  mod.setRange(rng);
  mod.setHeaderRow();
  var createTableSql = mod.makeCreateTableSql(tableName, rdbmsName);
  var insertSqls = mod.makeInsertSql(tableName);
  writeSqlToSheet(createTableSql, insertSqls, newSpreadsheetName, newSheetName)
}

/*
Use the GUI code defined in file "index.html" to create a form that is then displayed in the spreadsheet.
*/
function displayForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      html = HtmlService.createHtmlOutputFromFile('index').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  ss.show(html);
}

/*
Extract the user input values from the form using the element names defined in the HTML.
Pass these options to "runMod()". The call is wrapped in a "try ... catch", on success,
a message box is displayed. If there is an error, the error message is written to the
script log, check the "View->Logs" to see the error.
*/
function getValuesFromForm(form){
  var tableName, 
      rdbmsName, 
      newSpreadsheetName, 
      newSheetName;
  tableName = form.table_name;
  rdbmsName = form.rdbms_name;
  newSpreadsheetName = form.new_spreadsheet_name;
  newSheetName = form.new_sheet_name;
  //Logger.log([tableName, rdbmsName, newSpreadsheetName, newSheetName]);
  try {
    runMod(tableName, rdbmsName, newSpreadsheetName, newSheetName);
    Browser.msgBox('Done!')
  } catch(e) {
    Logger.log(e);
  }
}

/*
The "onOpen()" trigger addes the menu that is used to display the form.
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Create SQL')
      .addItem('SQL Create Form', 'displayForm')
      .addToUi();
}
