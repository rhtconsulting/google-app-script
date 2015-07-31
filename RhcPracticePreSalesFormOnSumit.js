function formSubmitReply(e) {
  sendEmailToPractice(e);
}

function sendEmailToPractice(e){
  var itemResponses = e.response.getItemResponses();
  var emailAddress = "nhopman@redhat.com, gborsuk@redhat.com, tanderso@redhat.com, jmelhorn@redhat.com, vchebolu@redhat.com, tquinn@redhat.com, akarpe@redhat.com, jwaldman@redhat.com, jholmes@redhat.com";
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
   
   subject = "New PreSales Request - ACTION REQUIRED - " + customerName + " - " + oppNum;
  
  var email = {"to":emailAddress, "cc":cc,"subject":subject,"htmlBody":emailBody}; 
  MailApp.sendEmail(email);
}
