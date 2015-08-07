function formSubmitReply(e) {
  sendEmailToRmo(e);
}

function sendEmailToRmo(e){
  var itemResponses = e.response.getItemResponses();
  var emailAddress = "asoderbe@redhat.com, cwestove@redhat.com, jevans@redhat.com";
  var cc = e.response.getRespondentEmail();
  var subject = "";
  var customerName = "";
  var oppNum = ""
  var emailBody = "";
  
  
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    var itemTitle = itemResponse.getItem().getTitle();
    var itemValue = itemResponse.getResponse(); 
    
    Logger.log( 'Response to the question "%s" was "%s"', itemTitle, itemValue );
    
    emailBody +=  "<p class=\"p2\"><b>" + itemTitle + "</b></p>\n" + itemValue + "\n"; 
    
    if ( itemTitle === "Customer Name" ){
      customerName = itemValue;
    }
    if ( itemTitle === "Salesforce Opp #" ){
      oppNum = itemValue; 
    }
   }
   
   subject = "New Staffing Request - ACTION REQUIRED - " + customerName + " - " + oppNum;
  
  var email = {"to":emailAddress, "cc":cc,"subject":subject,"htmlBody":emailBody}; 
  MailApp.sendEmail(email);
}
