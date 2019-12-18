function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var detail = active.getSheetByName('Invoice');
  ui.createMenu('Billing')
      .addItem('Process Payment', 'showSidebar')
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
  var infoSheet = active.getSheetByName('Client Info');
  var number = incomingLog.getRange(incomingLog.getLastRow(),2,1,1).getValue();
  var numberFinder = billingLog.createTextFinder(number);
  var updateRow = numberFinder.findNext().getRow();
  var updateRange = billingLog.getRange(updateRow,1,1,14);
  var incomingRow = incomingLog.getRange(incomingLog.getLastRow(),1,1,14).getValues();
  var name = incomingLog.getRange(incomingLog.getLastRow(),3,1,1).getValue();
  var nameFinder = infoSheet.createTextFinder(name);
  var addRow = nameFinder.findNext().getRow();
  var geo = infoSheet.getRange(addRow,6,1,2).getValues();
  var geoRange = billingLog.getRange(updateRow,20,1,2);
      geoRange.setValues(geo); 
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

function showSidebar() {
  var list = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Reconciliations')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(list);
}

function addPayment(number,amount,currency,date){
  var sheet = SpreadsheetApp.getActive();
  var payments = sheet.getSheetByName('Reconciliations');
  var exchangesheet = sheet.getSheetByName('Exchange Rates');
  var row = payments.getLastRow()+1;
  var today = new Date();
  const monthNames = ["January", "February", "March", "April", "May", "June","July", "August", "September", "October", "November", "December"];
  var month = monthNames[today.getMonth()-1]
  var numberRange = payments.getRange(row,1,1,1);
  var amountRange = payments.getRange(row,2,1,1);
  var euroRange = payments.getRange(row,3,1,1);
  var rateRange = payments.getRange(row,4,1,1);
  var dateRange = payments.getRange(row,5,1,1);
  var diffRange = payments.getRange(row,6,1,1);
  var rowFinder = exchangesheet.createTextFinder(currency);
  var rowNum = rowFinder.findNext().getRow();
  var colFinder = exchangesheet.createTextFinder(month);
  var colNum = colFinder.findNext().getColumn();
  var rate = exchangesheet.getRange(rowNum,colNum,1,1).getValue();
  var formula = "=RC[-1]/RC[1]";
     numberRange.setValue(number);
     amountRange.setValue(amount);
     euroRange.setValue(formula);
     rateRange.setValue(rate);
     dateRange.setValue(date);
     processPayment();
}

function processPayment(){
  var sheet = SpreadsheetApp.getActive();
  var payments = sheet.getSheetByName('Reconciliations');
  var log = sheet.getSheetByName('Billing Log');
  var number = payments.getRange(payments.getLastRow(),1,1,1).getValue();
  var numberFinder = log.createTextFinder(number);
  var updateRow = numberFinder.findNext().getRow();
  var updateRange = log.getRange(updateRow,15,1,5);
  var incomingRow = payments.getRange(payments.getLastRow(),2,1,5).getValues();
      updateRange.setValues(incomingRow);
      cleanPayments();
}

function cleanPayments(){
  var sheet = SpreadsheetApp.getActive();
  var payments = sheet.getSheetByName('Reconciliations');
  var cleanRange = payments.getLastRow();
    payments.deleteRow(cleanRange);
}

function showCurrent(){
  var sheet = SpreadsheetApp.getActive();
  var status = sheet.getSheetByName('Reconciliation Status');
  var resultOne = status.getRange(2,2,1,1).getValue();
  var resultTwo = status.getRange(3,2,1,1).getValue();
  var resultThree = status.getRange(4,2,1,1).getValue();
  var current = HtmlService.createHtmlOutput('<table style="width:100%"><tr><th>Due Within</th><th>Amount</th></tr><tr><td>14 Days</td><td>'+ resultOne +'</td></tr><tr><td>30 Days</td><td>'+ resultTwo +'</td></tr><tr><td>60 Days</td><td>'+ resultThree +'</td></tr></table>')
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(current, 'Currently Due');
}

function showOverdue(){
}

function showRecent(){
}