function onOpen(e) {
  var ui = SpreadsheetApp.getActive().getUi();
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var detail = active.getSheetByName('Invoice');
  ui.createMenu('Update')
      .addItem('Update Line Item', 'setRow')
      .addToUi();
}

function setRow() {
  var sheet = SpreadsheetApp.getActive();
  var tab = sheet.getSheetByName('Billing Log');
  var incomingtab = sheet.getSheetByName('Incoming Line Items');  
  var row = incomingtab.getLastRow();
  var conditionOne = incomingtab.getRange(row,1,1,1).getValue();
  if  (conditionOne != "Created Date" ){   
    var exchangesheet = sheet.getSheetByName('Exchange Rates');
    var currency = incomingtab.getRange(row,8,1,1).getValue();
    var duedate = incomingtab.getRange(row,14,1,1).getValue();
    var month = incomingtab.getRange(row,15,1,1).getValue();  
    var rowFinder = exchangesheet.createTextFinder(currency);
    var rowNum = rowFinder.findNext().getRow();
    var colFinder = exchangesheet.createTextFinder(month);
    var colNum = colFinder.findNext().getColumn();
    var exchangeRate = exchangesheet.getRange(rowNum,colNum,1,1).getValue();
    var exchangeRateDest = incomingtab.getRange(row,13,1,1);
    var dueDateDest = incomingtab.getRange(row,14,1,1);
    var euroAmount = '=RC[-1]/RC[3]';
    var euroAmountDest = incomingtab.getRange(row,10,1,1);
    var taxAmount = '=RC[-1]*RC[-4]';
    var taxAmountDest = incomingtab.getRange(row,11,1,1);
    var bruttoAmount = '=RC[-2]+RC[-1]';
    var bruttoAmountDest = incomingtab.getRange(row,12,1,1);
        exchangeRateDest.setValue(exchangeRate);
        euroAmountDest.setValue(euroAmount);
        taxAmountDest.setValue(taxAmount);
        bruttoAmountDest.setValue(bruttoAmount);
        dueDateDest.setValue(duedate);
           finaliseRow();
    }
}

function finaliseRow() {
  var active = SpreadsheetApp.getActive();
  var billingLog = active.getSheetByName('Billing Log');
  var incomingLog = active.getSheetByName('Incoming Line Items');
  var number = incomingLog.getRange(incomingLog.getLastRow(),2,1,1).getValue();
  var numberFinder = billingLog.createTextFinder(number);
  var updateRow = numberFinder.findNext().getRow();
  var updateRange = billingLog.getRange(updateRow,1,1,14);
  var incomingRow = incomingLog.getRange(incomingLog.getLastRow(),1,1,14).getValues();
      updateRange.setValues(incomingRow);
      cleanIncomingLog();
}

function cleanIncomingLog() {
  var activeSheet = SpreadsheetApp.getActive();
  var finalLog = activeSheet.getSheetByName('Billing Log');
  var incomingRow = activeSheet.getSheetByName('Incoming Line Items');
  var logTest = incomingRow.getRange(1,2,1,1).getValue();
  var lastRow = incomingRow.getLastRow();
  var incomingTest = incomingRow.getRange(lastRow,2,1,1).getValue();
      if ( incomingTest != logTest) {
          incomingRow.deleteRow(lastRow);
          }
}     