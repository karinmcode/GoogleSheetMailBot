/**
 * Sends an email using data from the active Google Sheet and the active row.
 * Avoids sending duplicate emails by checking the 'Msg status' column.
 * 
 * Requirements for code to work:
 * - sheet must have columns with the following headers : Email,	Subject,	Message,	Msg status
 * - make sure you have the necessary permissions (for sending emails and accessing the sheet's data) in the Google Apps Script environment
 */

function sendEmails(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getSheetValues(1,1,1,sheet.getLastColumn())[0];

  var col4msg_status = headers.indexOf('Msg status') + 1;
  var col4msg = headers.indexOf('Message') + 1;
  var col4email = headers.indexOf("Email") + 1;
  var col4subjet = headers.indexOf("Subject") + 1;

  // Get the current selected range.
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var endRow = startRow + range.getNumRows() - 1;

  for (var row = startRow; row <= endRow; row++) {
    var msg_status = sheet.getSheetValues(row, col4msg_status, 1, 1)[0][0];
    var subject = sheet.getSheetValues(row, col4subjet, 1, 1)[0][0];
    var message = sheet.getSheetValues(row, col4msg, 1, 1)[0][0];
    var emailAddress = sheet.getSheetValues(row, col4email, 1, 1)[0][0];

    var condForSendingEmail = (msg_status == "");

    if (condForSendingEmail) {
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(row, col4msg_status).setValue(new Date);
    }
  }

  // The rest of your formatting code remains unchanged:
  sheet.getRange(2, col4msg_status, sheet.getLastRow()-1, 1).setNumberFormat("dd/MM/yyyy HH:mm");
  sheet.getRange(2, col4msg_status, sheet.getLastRow()-1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(2, col4msg_status, sheet.getLastRow()-1, 1).setFontSize(6);

  sheet.getRange(2, col4msg, sheet.getLastRow()-1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(2, col4msg, sheet.getLastRow()-1, 1).setFontSize(6);

  sheet.getRange(2, col4subjet, sheet.getLastRow()-1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(2, col4subjet, sheet.getLastRow()-1, 1).setFontSize(6);

  sheet.getRange(2, col4email, sheet.getLastRow()-1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(2, col4email, sheet.getLastRow()-1, 1).setFontSize(6);
  sheet.getRange(2, col4email, sheet.getLastRow()-1, 1).setFontColor("black");

  SpreadsheetApp.flush();
}




function getEmailHistory() {
  
  //var email = "nevena.rockov@hotmail.com";
  
  // add info in new sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // get email
  var s = ss.getSheetByName("Dentists");
  var row = s.getActiveCell().getLastRow();
  var emailAddress = s.getSheetValues(row, 4, 1, 1);//data[3];
  var email = emailAddress[0];
  
  try {
    ss.deleteSheet(ss.getSheetByName(email));
  }
  catch(e){
  }
  
  
  var sh = ss.insertSheet();
  sh.setName(email);
  sh.setTabColor("ff0000");
  sh.appendRow(["From" , "To" , "Date" , "Subject" , "Body"])
  sh.setFrozenRows(1)
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 600);

  
  
  var threads = GmailApp.search('to:' + email + "OR from:" + email);
  var threads = GmailApp.search('to:' + email + " OR from:" + email);
  var A = new Array; // From, To, Date, Obj, Body
  for (it=0; it<threads.length ; it++){
  var thread=threads[it];

    var msg = thread.getMessages();
    for (im =0;im<thread.getMessageCount();im++){
      
      var d = msg[im].getDate();
      var txt = msg[im].getPlainBody();
      Logger.log(txt);
      var t=cleanBody(txt);
      var values = [ msg[im].getFrom() , msg[im].getTo() , d , msg[im].getSubject() ,  t[0]];
      sh.appendRow(values);
      var r = sh.getRange(sh.getLastRow(), 1,1,5);
      r.setWraps([[true, true, true, true , true]])
      sh.setRowHeight(sh.getLastRow(), 150);
      r.setVerticalAlignment("Top")


    }
    
  }
  
  sh.sort(3, true);
  
  for (irow = 2;irow<=sh.getLastRow();irow++){
      sh.setRowHeight(irow, 250)
      for (icol =1; icol<=2;icol++){
      var r = sh.getRange(irow, icol);
    var val = r.getValue();
      if ((val == "quaiduchevalblancla@gmail.com" ) || (val == '"quaiduchevalblancla@gmail.com" <quaiduchevalblancla@gmail.com>' )){
        r.setBackgroundRGB(50, 255, 50)
      }
  }
  }
  
}




function autogenerate_email_from_website(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Fetch headers (assuming headers are in the first row)
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Locate the columns by their header names
  var WEBSITE_COLUMN = headers.indexOf("Website") + 1;  // +1 because arrays are 0-indexed but sheets are 1-indexed
  var EMAIL_COLUMN = headers.indexOf("Email") + 1;

  // Check if edited cell is in the "Website" column
  if (range.getColumn() == WEBSITE_COLUMN) {
    var website = range.getValue();
    
    // Removing 'http://' or 'https://'
    website = website.replace(/^(http:\/\/|https:\/\/)/, '');
    
    // Strip out paths after the domain
    website = website.split('/')[0];

    // Check if corresponding "Email" cell is empty
    var correspondingEmailCell = sheet.getRange(range.getRow(), EMAIL_COLUMN);
    if (correspondingEmailCell.getValue() !== "") {
      return;  // Exit the function if the "Email" cell is not empty
    }

    // Validation checks
    if (!website || !isValidWebsite(website)) {
      return;  // Exit the function if website is empty or invalid
    }

    // Split the website domain by "." and extract the second last part
    var domainParts = website.split(".");
    
    // Check if there are at least 2 parts
    if (domainParts.length >= 2) {
      var secondLastPart = domainParts[domainParts.length - 2];
      var lastPart = domainParts[domainParts.length - 1]; // to keep the original extension
      
      // Construct the email address
      var email = "info@" + secondLastPart + "." + lastPart;
      
      // Set the email address in the "Email" column of the same row
      sheet.getRange(range.getRow(), EMAIL_COLUMN).setValue(email);
    }
  }
}
