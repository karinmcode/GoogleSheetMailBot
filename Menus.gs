function onOpen() {
// The onOpen function is executed automatically every time a Spreadsheet is loaded
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var menuEntries = [];
   // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is
   // executed.
   menuEntries.push({name: "Send emails", functionName: "sendEmails"});
   menuEntries.push({name: "GetEmailHistory", functionName: "getEmailHistory"});

   ss.addMenu("Scripts", menuEntries);

  // Go to last entry
  var s = ss.getSheetByName("Dentists");
  var lastrow = s.getLastRow();
  s.setActiveRange(s.getRange(lastrow, 1));
 }

function onEdit(e){

  autogenerate_email_from_website(e);
}
  
