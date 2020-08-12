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
  var infotab = sheet.getSheetByName('Client Info'); 
  var row = incomingtab.getLastRow();
  var test = incomingtab.getRange(1,1,1,1).getValue();
  var testRow = incomingtab.getRange(row,1,1,1).getValue();
    if (test != testRow){  
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
          phaseTwo();
    }
}

function phaseTwo() {
  var sheet = SpreadsheetApp.getActive();
  var incomingtab = sheet.getSheetByName('Incoming Line Items'); 
  var infotab = sheet.getSheetByName('Client Info'); 
  var row = incomingtab.getLastRow();
  var accName = incomingtab.getRange(row,3,1,1).getValue();
  var nameFinder = infotab.createTextFinder(accName);
  var accountRow = nameFinder.findNext().getRow();
  var geo = infotab.getRange(accountRow,6,1,2).getValues();
  var geoRange = incomingtab.getRange(row,18,1,2);
  var number = incomingtab.getRange(row,2,1,1).getValue();
  var year = number.substring(0,4);
  var yearRange = incomingtab.getRange(row,22,1,1);
  var accNumber = infotab.getRange(accountRow,3,1,1).getValue();
  var newNumber = incomingtab.getRange(row,20,1,1);
  var prefix = infotab.getRange(accountRow,4,1,1).getValue();
  var prefixRange = incomingtab.getRange(row,24,1,1);
  var accountExecutive = infotab.getRange(accountRow,8,1,1).getValue();
  var accountExecRange = incomingtab.getRange(row,25,1,1);
  var month = incomingtab.getRange(row,15,1,1).getValue();
  var moveMonth = incomingtab.getRange(row,21,1,1);
        prefixRange.setValue(prefix);
        accountExecRange.setValue(accountExecutive);
        geoRange.setValues(geo);
        yearRange.setValue(year);
        newNumber.setValue(accNumber);
        moveMonth.setValue(month);
        finaliseRow();
}

function finaliseRow() {
  var active = SpreadsheetApp.getActive();
  var tab = active.getSheetByName('Billing Log');
  var incomingtab = active.getSheetByName('Incoming Line Items');
  var infoSheet = active.getSheetByName('Client Info');
  var lastRow = incomingtab.getLastRow();
  var number = incomingtab.getRange(lastRow,2,1,1).getValue();
  var numberFinder = tab.createTextFinder(number);
  var updateRow = numberFinder.findNext().getRow();
  var updateRange = tab.getRange(updateRow,1,1,24);
  var incomingRow = incomingtab.getRange(lastRow,1,1,24).getValues();
  var removeRange = tab.getRange(updateRow,1,1,1);
  var servicePeriod = '=CONCATENATE(RC[-2]," ",RC[-1])';
  var serviceRange = incomingtab.getRange(lastRow,23,1,1);
  var difference = '=RC[-5]-RC[-2]';
  var differenceDest = incomingtab.getRange(lastRow,17,1,1);
  var paidAmount = '';
  var paidDest = incomingtab.getRange(lastRow,15,1,1);
      paidDest.setValue(paidAmount);
      differenceDest.setValue(difference);
      serviceRange.setValue(servicePeriod);
      removeRange.clearDataValidations();
      updateRange.setValues(incomingRow);
      copyToNNNBL();
}

function copyToNNNBL() {
  var sheet = SpreadsheetApp.getActive();
  var newItems = sheet.getSheetByName('Incoming Line Items');
  var newRow = newItems.getRange(newItems.getLastRow(),1,1,27).getValues();
  var NNNBL = SpreadsheetApp.openById('1CgitHYvXGAUVBsci_Pyz0caPDww7Wf4JJbM29h6ZDsM');
  var dest = NNNBL.getSheetByName('Incoming Line Items');
  var copy = dest.getRange(dest.getLastRow()+1,2,1,27);
      copy.setValues(newRow);
      cleanIncomingLog();
}

function cleanIncomingLog() {
  var activeSheet = SpreadsheetApp.getActive();
  var finalLog = activeSheet.getSheetByName('Billing Log');
  var incomingRow = activeSheet.getSheetByName('Incoming Line Items');
  var logTest = incomingRow.getRange(1,2,1,1).getValue();
  var lastRow = incomingRow.getLastRow();
  var incomingTest = incomingRow.getRange(lastRow,2,1,1).getValue();
      incomingRow.deleteRow(lastRow);
      setRow();
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