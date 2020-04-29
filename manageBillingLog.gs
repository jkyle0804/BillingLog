function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var detail = active.getSheetByName('Invoice');
  ui.createMenu('Billing')
         .addItem('Update Incoming Items', 'setRow')
      .addSubMenu(ui.createMenu('Payments')
         .addItem('Single Payment', 'showSidebar')
         .addItem('Multiple Payments','paymentsLoop'))
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
    var difference = '=RC[-5]-RC[-2]';
    var differenceDest =incomingtab.getRange(row,17,1,1);
    var name = incomingLog.getRange(incomingLog.getLastRow(),3,1,1).getValue();
    var nameFinder = infoSheet.createTextFinder(name);
    var addRow = nameFinder.findNext().getRow();
    var geo = infoSheet.getRange(addRow,6,1,2).getValues();
    var geoRange = billingLog.getRange(updateRow,20,1,2);
    var year = number.substring(0,4);
    var yearRange = billingLog.getRange(updateRow,23,1,1);
    var accName = incomingLog.getRange(incomingLog.getLastRow(),3,1,1).getValue();
    var nameFinder = infoSheet.createTextFinder(accName);
    var accountRow = nameFinder.findNext().getRow();
    var accNumber = infoSheet.getRange(accountRow,3,1,1).getValue();
    var newNumber = billingLog.getRange(updateRow,22,1,1);
        geoRange.setValues(geo);
        yearRange.setValue(year);
        newNumber.setValue(accNumber);
        exchangeRateDest.setValue(exchangeRate);
        euroAmountDest.setValue(euroAmount);
        taxAmountDest.setValue(taxAmount);
        bruttoAmountDest.setValue(bruttoAmount);
        dueDateDest.setValue(duedate);
        differenceDest.setValue(difference);
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
  var updateRange = billingLog.getRange(updateRow,1,1,24);
  var incomingRow = incomingLog.getRange(incomingLog.getLastRow(),1,1,24).getValues();
  var removeRange = billingLog.getRange(updateRow,1,1,1);
  var month = incomingLog.getRange(incomingLog.getLastRow(),15,1,1).getValue();
  var moveMonth = incomingLog.getRange(incomingLog.getLastRow(),21,1,1);
      removeRange.clearDataValidations();
      moveMonth.setValue(month); 
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
  var monthCalc = monthNames[today.getMonth()-1]
  if (monthCalc == null){
    var month = 'December';
  }
  else {
    month = monthCalc;
  }
  var numberRange = payments.getRange(row,1,1,1);
  var amountRange = payments.getRange(row,2,1,1);
  var euroRange = payments.getRange(row,3,1,1);
  var rateRange = payments.getRange(row,4,1,1);
  var dateRange = payments.getRange(row,5,1,1);
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

function paymentsLoop(){
  var ui = SpreadsheetApp.getUi();
  var warning = ui.alert('Have you enetered the payment details in the Reconciliations tab?', ui.ButtonSet.YES_NO);
     if (warning == ui.Button.YES) {
     var sheet = SpreadsheetApp.getActive();
     var payments = sheet.getSheetByName('Reconciliations');
     var test = payments.getRange(1,1,1,1).getValue();
         for (var i=0; i<payments.length; i++){
              if (test != "INVOICE NUMBER"){
                processPayment();
    }
  }
}
  else if (warning == ui.Button.NO) {
    ui.alert('You must prepare the payments list before you can process multiple payments.');
  }
}

function processPayment(){
  var sheet = SpreadsheetApp.getActive();
  var payments = sheet.getSheetByName('Reconciliations');
  var log = sheet.getSheetByName('Billing Log');
  var number = payments.getRange(payments.getLastRow(),1,1,1).getValue();
  var numberFinder = log.createTextFinder(number);
  var updateRow = numberFinder.findNext().getRow();
  var updateRange = log.getRange(updateRow,15,1,4);
  var incomingRow = payments.getRange(payments.getLastRow(),2,1,4).getValues();
  var difference = '"=RC[-7]-RC[-3]';
  var diffRange = log.getRange(updateRow,18,1,1);
      updateRange.setValues(incomingRow);
      diffRange.setValue(difference);
      cleanPayments();
}

function cleanPayments(){
  var sheet = SpreadsheetApp.getActive();
  var payments = sheet.getSheetByName('Reconciliations');
  var cleanRange = payments.getLastRow();
    payments.deleteRow(cleanRange);
}

function portNumber(){
  var log = SpreadsheetApp.getActive().getActiveSheet();
  var approval = log.getActiveCell().getRow();
  var response = log.getRange(approval,1,1,1).getValue();
  var number = log.getRange(approval,2,1,1).getValue();
  var doc = log.getRange(approval,4,1,1).getValue();
  if (response == 'create'){
    var invoiceFile = SpreadsheetApp.openById(doc);
    var newNumber = invoiceFile.getSheetByName('Calculations').getRange(16,6,1,1);
    var url = "https://docs.google.com/spreadsheets/d/"+doc;
    var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
    var userInterface = HtmlService.createHtmlOutput(html);
        newNumber.setValue(number);
        SpreadsheetApp.getUi().showModalDialog(userInterface, "Opening Invoice File");    
  }
}