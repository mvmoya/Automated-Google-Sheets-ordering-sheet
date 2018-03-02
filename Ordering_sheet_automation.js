
/*
* Welcome to the Heintz Lab Ordering sheet code!
* 
* The first function (onEdit()) has multiple jobs:
* It creates a new empty row at the top of the sheet (with the right formatting) after a new order is completed in the new order row.
* (completed = the "Order Requested" column is filled in with a date by the user). The first column will say "Write new order in this row".
* It also changes the formatting of the newly placed order row to match the rest of the sheet.
* DEPRECATED**The onEdit() function also listens and checks for the same catalog order in the sheet, so it will warn you with a dialog box if you're about to
* make a duplicate order.
* It's also supposed to sort orders by the "Order Requested" date, in case someone puts in an old order, but (I think) that hasn't been working...**END
* 
* The second function (sendNewOrder()) checks the current date and time every 2 hours and, if it's a weekday between 8 and 5, it sends Ivarine
* an email telling her if there are unplaced orders in the order sheet (unplaced = "order placed" box has no date in it).
* This function also checks what day of the month it is. If it's the first of the month, it automatically creates a new sheet for the new month,
* calls it "Current month", and archives the old sheet with the month number/year.
* Any orders that were not placed before the end of the month are migrated onto the new sheet automatically so they don't get forgotten.
* The new sheet should have all the same formatting as the old sheet, and already have the new order line written in, as well as the total money
* function already working at the bottom of the sheet.
*/


function onEdit() {
  /* Upon an edit, get the active spreadsheet and check if the order date column has been inputted. If it has, proceed with making new row */
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current month");
  if (sheet.getRange(2, 11).getValue() != "[Enter date]") {
    
    var order_date = sheet.getRange(2, 11);  

    sheet.insertRowAfter(1);
            
    order_date.setValue("[Enter date]");
    
    var low_row = sheet.getRange(4, 1, 1, 18);
    low_row.copyFormatToRange(sheet, 1, 18, 3, 3);
    
    var old_row = sheet.getRange(3, 1, 1, 18);
    var new_row = sheet.getRange(2, 1, 1, 18);     
    var new_cell = sheet.getRange(2, 1);    
    

    low_row.copyFormatToRange(sheet, 1, 18, 2, 2);
    sheet.getRange(2, 7).setFormula("=E2*F2");
    new_cell.setValue("[Write one new order at a time in this row]");
    new_cell.setFontSize(11);
    new_cell.setFontWeight("bold");
    new_cell.setFontColor("#E60000");
    
    order_date.setFontSize(11);
    order_date.setFontWeight("bold");
    order_date.setFontColor("#E60000");
    old_row.setFontSize(10);
    old_row.setFontColor("#000000");
    old_row.setFontWeight("normal");
    sheet.getRange(3, 8).setFontColor("#1155cc");
    
  }
  /*else if (sheet.getRange(2, 3).isBlank() == false && sheet.getRange(2, 4).isBlank() == true) {
    var catalognum = sheet.getRange(2, 3).getValue();
    var keepGoing = true;
    var rownum = 3;
    var orderRow = sheet.getRange(2, 1, 1, 11);
    var new_cell = sheet.getRange(2, 1);
    while(keepGoing == true) {
      if (sheet.getRange(rownum, 3).getValue() == catalognum) {
        var user = sheet.getRange(rownum, 9).getValue();
        var date = sheet.getRange(rownum, 11).getValue().toDateString();
        var result = Browser.msgBox(('Note!\n'+user+' placed an order with this catalog number on '+date), Browser.Buttons.OK);
        keepGoing = false;
        if (result == "no") {
          orderRow.clear({contentsOnly: true});
          sheet.getRange(2, 7).setFormula("=E2*F2");
          new_cell.setValue("[Write new order in this row]");
          break;
        }
         else{
           keepGoing = false;
          }
      }
      else if (sheet.getRange(rownum, 3).isBlank() == true && sheet.getRange((rownum+1), 3).isBlank() == true) {
          keepGoing = false;
        }
      else{
        rownum += 1;
        keepGoing = true;
        };
      };
      }*/
  
  else{
    return 0
  };
  
};

/* function timeTrigger() {
  ScriptApp.newTrigger('sendNewOrder')
  .timeBased()
  .everyHours(2)
  .create(); 
} */

function sendNewOrder() {
 var nowH = new Date().getHours();
 var nowD = new Date().getDay();
 var nowDate = new Date().getDate();
  if (nowH > 8 && nowH < 17 && nowD > 0 && nowD < 6){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Current month");
    var keepGoing = true;
    var num = 3;
    while (keepGoing == true) {
      if (sheet.getRange(num, 12, 1).isBlank() == true && sheet.getRange(num, 1, 1).isBlank() == false) {
        num += 1;
      }
      else if (sheet.getRange(num, 12, 1).isBlank() == false || sheet.getRange(num, 1, 1).isBlank() == true){
        keepGoing = false;
      }
    }
    if (num > 3){
      var email_body = "The following orders in the sheet need to be placed: \n \n";
      var htmlTableBody="";
      for (i = 3; i < num; i++) {
        var quantity = sheet.getRange(i, 5, 1).getValue().toString();
        var item_name = sheet.getRange(i, 1, 1).getValue().toString();
        var cat_number = sheet.getRange(i, 3, 1).getValue().toString();
        var price = sheet.getRange(i, 6, 1).getValue().toString();
        var vendor = sheet.getRange(i, 2, 1).getValue().toString();
        var webpage = sheet.getRange(i, 8, 1).getValue().toString();
        var orderer = sheet.getRange(i, 9, 1).getValue().toString();
        
        email_body = email_body + (quantity + "\t " + item_name + "\t (" + cat_number + ")\t $" + price + "\t " + vendor + "\t " + webpage + "\t ordered by " + orderer + "\n");
        htmlTableBody = htmlTableBody + "<tr><td>"+item_name+"</td><td>"+vendor+"</td><td>"+cat_number+"</td><td>"+quantity+"</td><td>$"+price+"</td><td>"+webpage+"</td></tr>";
      }
      email_body = email_body + '\nRemember to write the "Order placed" date and grant into the sheet:\t' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
      var htmlEmailBody = "<!DOCTYPE html><html><head><title>The order email</title><style>table{border:2px solid black;}th{border:1px solid black; font-family:Verdana; font-size:10px; text-align:left;}td{border:1px solid black; font-family:Verdana; font-size:10px; text-align:left; word-wrap:break-word;}</style></head><body><p>The following orders still need to be placed:</p><table><th>Item</th><th>Vendor</th><th>Catalog #</th><th>Qty</th><th>Unit price</th><th>Webpage</th>"+htmlTableBody+"</table><p>Remember to write the 'Order placed' date and grant into the sheet: "+SpreadsheetApp.getActiveSpreadsheet().getUrl()+"</p></body></html>";
      //Browser.msgBox(email_body, Browser.Buttons.OK_CANCEL);//
      MailApp.sendEmail("irose@rockefeller.edu", "Orders that need to be placed", email_body, {htmlBody: htmlEmailBody});
      
    }
    else{
      return 0;
    }
    }
  else if (nowDate == 1 && nowH > 1 && nowH < 3){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var oldsheet = ss.getSheetByName("Current month");
    var toprows = oldsheet.getRange(1, 1, 2, 18);
    var totalrow = oldsheet.getRange(oldsheet.getMaxRows(), 6, 1, 2);
    var oldMonth = new Date().getMonth();
    var thisMonth = new Date().getMonth() + 1;
    var thisYear = new Date().getFullYear();
    var oldYear = new Date(). getFullYear() - 1;
    var keepGoing = true;
    var num = 3;
    oldsheet.getRange(2, 1).setValue('   This sheet is no longer active. Please go to "Current month" tab.');
    while (keepGoing == true) {
      if (oldsheet.getRange(num, 12).isBlank() == true && oldsheet.getRange(num, 1).isBlank() == false) {
        num += 1;
      }
      else {
        keepGoing = false;
      }
    }
    var rows = num - 3
    if (rows != 0){
      var unplacedOrders = oldsheet.getRange(3, 1, rows, 18);
    }
    if (oldMonth === 0){
      oldsheet.setName("12/"+oldYear);
    }
    else{
      oldsheet.setName(oldMonth+"/"+thisYear);
    }
    
    ss.insertSheet("Current month", 0);
    var newsheet = ss.getSheetByName("Current month");
    var newtoprows = newsheet.getRange(1,1,2,18);
    toprows.copyTo(newtoprows);
    newsheet.getRange(2,1).setValue("[Write one new order at a time in this row]");
    totalrow.copyTo(newsheet.getRange(num, 6, 1, 2));
    
    if (rows != 0){
      unplacedOrders.copyTo(newsheet.getRange(3, 1, rows, 18));    
    }
    else {
      var whoohoo=0;
    }    
    newsheet.getRange(2, 11).setValue('[Enter date]');
    newsheet.setRowHeight(2, 26);
    newsheet.setRowHeight(1, 26);
    newsheet.setColumnWidth(1, 300);
    newsheet.setColumnWidth(3, 110);
    newsheet.setColumnWidth(4, 80);
    newsheet.setColumnWidth(5, 40);
    newsheet.setColumnWidth(6, 80);
    newsheet.setColumnWidth(7, 90);
    newsheet.setColumnWidth(8, 135);
    newsheet.setTabColor("#75e667");
    oldsheet.setTabColor("#e60000");
    
    newsheet.setFrozenRows(1);
    
    var protection = newsheet.getRange("A1:R1").protect().setDescription("Header cannot be changed");
    protection.addEditor("mvmoya10@gmail.com");
    
  }; 
};

/*function sendHTMLemail() {
  var item= "Anti-FOXP4 antibody produced in rabbit";
  var link="http://www.sigmaaldrich.com/catalog/product/sigma/hpa007176?lang=en&region=US";
  var catalogNum= "2721634";
  var email_body= "Yo";
  var row= "<tr><td>"+item+"</td><td>Sigma-Aldrich</td><td>"+catalogNum+"</td><td>1</td><td>$240.00</td><td>"+link+"</td></tr>";
  var emailHtml= "<!DOCTYPE html><html><head><title>The order email</title><style>table{border:2px solid black;}th{border:1px solid black; font-family:Verdana; font-size:10px; text-align:left;}td{border:1px solid black; font-family:Verdana; font-size:10px; text-align:left; word-wrap:break-word;}</style></head><body><p>The following orders still need to be placed:</p><table><th>Item</th><th>Vendor</th><th>Catalog #</th><th>Qty</th><th>Price</th><th>Webpage</th>"+row+row+"</table><p>Be sure to mark the orders as placed:</p></body></html>";
  MailApp.sendEmail("mmoya@rockefeller.edu", "Orders that need to be placed", email_body, {htmlBody: emailHtml});
};*/                                                                                  
