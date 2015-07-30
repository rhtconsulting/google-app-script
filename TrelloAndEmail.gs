/*
 * Global variables. 
 * Some (emailAddress) are specific to the environment (e.g. TEST vs PROD).
 * Some (EMAIL_SENT) have values that are specific to the spreadsheet design -- keep them in sync with the spreadsheet.
 */
  var EMAIL_SENT = "EMAIL_SENT";
  var SFDC_OPPORTUNITY_NUMBER = "SFDC Opportunity Number";
  var ACCOUNT_NAME = "Account Name";
  var emailAddress = '';
  //var emailAddress = '';
  var trelloAddress = '';


 

function Initialize() {
  var triggers = ScriptApp.getProjectTriggers();
 
  for(var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  ScriptApp.newTrigger("sendEmails")
  .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
  .onFormSubmit()
  .create();
}

/**
 * Send one email for each line in the spreadsheet that does not contain the value "EMAIL_SENT" in the column with the same name.
 */
function sendEmails() {

  var subject = "Incoming Consulting Opportunity - ACTION REQUIRED";  
  var message = "<p class=\"p1\"><b>RHC West - Consulting Opportunity Request Form Submission</b></p>\n\n";
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var values = rows.getValues();
  
  var headerRow = values[0];
  var indexOfEMAIL_SENT = -1;
  var indexOfSFDC_OPPORTUNITY_NUMBER = -1;
  var indexOfACCOUNT_NAME = -1;
  var accountName;
  var oppNum;
  
  for(var i = 0; i < rows.getNumColumns(); i++) {
    if(EMAIL_SENT == headerRow[i]) {
      indexOfEMAIL_SENT = i;
    }
    if(SFDC_OPPORTUNITY_NUMBER == headerRow[i]) {
       indexOfSFDC_OPPORTUNITY_NUMBER = i; // 2
    }
    if(ACCOUNT_NAME == headerRow[i]) {
      indexOfACCOUNT_NAME = i;
    }
  }
  if(-1 == indexOfEMAIL_SENT || -1 == indexOfSFDC_OPPORTUNITY_NUMBER) {
    Logger.log("Required column not found.");
    // TODO throw an error here
  }

  // Find all rows that haven't yet been sent in emails, and send them.
  
  for (var i = 1; i < rows.getLastRow(); i++) {
    var rowValues = values[i];  
    if(values[i][indexOfEMAIL_SENT] != EMAIL_SENT) {
      for(var j = 0; j < rows.getLastColumn(); j++) {
        if(headerRow[j] != EMAIL_SENT) {
          message += "<p class=\"p2\"><b>" + headerRow[j] + "</b></p>\n" + values[i][j] + "\n";
          // message += headerRow[j] + ": " + values[i][j] + "\n"; // plain-text
        }
        if(indexOfSFDC_OPPORTUNITY_NUMBER == j) {
            oppNum = values[i][j];
        }
        if(indexOfACCOUNT_NAME == j) {
          accountName = values[i][j];
        }
      }    
      subject += " - " + accountName + " - " + oppNum;
      Logger.log("Sending this email: " + message);
      
      // Send to rhc-west Alias
      // MailApp.sendEmail(emailAddress, subject, message); // plain-text
      var emailContent = {"to":emailAddress,"subject":subject,"htmlBody":message}; // HTML formatted, to alias
      MailApp.sendEmail(emailContent);
      
      // Send to Trello
      subject = accountName + " - " + oppNum;
      emailContent = {"to":trelloAddress,"subject":subject,"htmlBody":message}; // HTML formatted, to Trello board
      MailApp.sendEmail(emailContent);
        
      // Clear the message
      message = "";
      // now add a new cell to the end of the row to mark the row 'EMAIL_SENT'
      var cell = rows.getCell(i+1,indexOfEMAIL_SENT+1).setValue(EMAIL_SENT);
    }
  }
}


/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Send Emails",
    functionName : "sendEmails"
  }];
  spreadsheet.addMenu("Script Center Menu", entries);
};

