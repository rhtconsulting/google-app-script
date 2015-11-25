//This code creates a PDF and e-mails it and stores it to a shared team Google Drive Folder
//This code also creates a new menu in the responses sheet to allow for a custom pull request of the PDF. That only gets sent via e-mail

//DISCLAIMER: This code is very ugly. I wrote it as fast as possible to get the job done, but it 
// could definitely use some cleaning up.

var docTemplate = "1DmwYsFwu6J0a4Q_hUZLry3E5VBZMUmePW9uSVHB9uNI";
var docName = "Enablement Participant Scorecard"

//This function creates the menu item on the Responses sheet
function onOpen() 
{
  SpreadsheetApp.getUi() 
      .createMenu('Create Report')
      .addItem('Create PDF', 'openDialog')
      .addToUi();
}

//This is the function that creates and emails the pdf when the form is submitted
function onFormSubmit(e)
{
  var email = e.values[4] + ", jwaldman@redhat.com"; //"jwaldman@redhat.com";
  var timestamp = e.values[0];
  var date = e.values[1];
  var technology = e.values[2];
  var name = e.values[3];
  var preFeed = e.values[5];
  var bootFeed = e.values[6];
  var HWFeed = e.values[7];
  var shadowFeed = e.values[8];
  var stackMonitoring = e.values[9];
  var stackNetworking = e.values[10];
  var stackStorage = e.values[11];
  var stackPlatform = e.values[12];
  var cloudForms = e.values[13];
  var openShift = e.values[14];
  var mobile = e.values[15];
  var fuse = e.values[16];
  var amq = e.values[17];
  var bpms = e.values[18];
  var brms = e.values[19];
  
  
  var copyId = DriveApp.getFileById(docTemplate).makeCopy(docName + ' for ' + name).getId();
  var copyDoc = DocumentApp.openById(copyId);
  var copyBody = copyDoc.getActiveSection();
  
  copyBody.replaceText('keyName', name);
  copyBody.replaceText('keyDate', date);
  copyBody.replaceText('keyTechnology', technology);
  copyBody.replaceText('keyPreworkFeedback', preFeed);
  copyBody.replaceText('keyBootcampFeedback', bootFeed);
  copyBody.replaceText('keyHWFeedback', HWFeed);
  copyBody.replaceText('keyShadowFeedback', shadowFeed);
  
  var skills = "";
  for(var i = 9; i < 20; i++)
  {
    if(e.values[i] != null && e.values[i] != "")
    {
      var tech = "";
      switch(i)
      {
        case 9: tech = "RHEL OpenStack - Monitoring";
          break;
        case 10: tech = "RHEL OpenStack - Networking";
          break;
        case 11: tech = "RHEL OpenStack - Storage";
          break;
        case 12: tech = "RHEL OpenStack - Platform";
          break;
        case 13: tech = "Red Hat CloudForms";
          break;
        case 14: tech = "Red Hat OpenShift";
          break;
        case 15: tech = "Red Hat Mobile Application Platform";
          break;
        case 16: tech = "JBoss Fuse";
          break;
        case 17: tech = "JBoss A-MQ";
          break;
        case 18: tech = "JBoss BPM Suite";
          break;
        case 19: tech = "JBoss BRMS";
          break;
        default: tech = "";
          
      }
      skills = skills + tech + ": " + e.values[i] + "\n";
    }
  }
  copyBody.replaceText('keySkillLevel', skills);
  
  copyDoc.saveAndClose();
  var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");
  
  var subject = "Enablement Scorecard";
  var body = "Please find the " + technology + " Enablement Report for " + name + " attached.";
  MailApp.sendEmail(email, subject, body, {htmlBody: body, attachments: pdf});
  
  var folder = DriveApp.getFolderById("0B4bc9lqZ4LL4WUYyTUtVeEQ0Z0E");
  var f = DriveApp.createFile(pdf);
  folder.addFile(f);
  DriveApp.getFileById(copyId).setTrashed(true);
  
}

//This function opens the window to choose a previously submitted response and create a PDF from that
function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Create Scorecard Report');
  //getList();
}

//This function is to populate the dropdown list for the previously submitted responses
function getList()
{
  var form = FormApp.openById("1ArBMLb19Gx5CYmlXRlaaXmDh5DTZ4_eZuT-W8FxfJLY");
  var formResponses = form.getResponses();
  var temp = "";
  var entries = [];
  for (var i = 0; i < formResponses.length; i++) 
  {
    var title = "";
    var response = "";
    var formResponse = formResponses[i];
    var itemResponses = formResponse.getItemResponses();
    var entry = "";
    var who;
    var when;
    var what;
    for (var j = 0; j < itemResponses.length; j++)
    {
      var itemResponse = itemResponses[j];
      title = itemResponse.getItem().getTitle();
      response = itemResponse.getResponse();
      switch(title)
      {
        case "Who is being scored?":
          who = response;
          break;
        case "What was the start date of the enablement program?":
          when = response;
          break;
        case "Which Emerging Technology Enablement Program did the participant take?":
          what = response;
          break;
      }
      entry = who + " " + what + " " + when;
      entries[i] = entry;
    }
  }
  //Logger.log(temp);
  return entries;
}

//Based on the selection from the dropdown, this creates the pdf and e-mails it
function processForm(formObject) {
  Logger.log("Entering Process");
  var selection = formObject.mySelect;
  var form = FormApp.openById("1ArBMLb19Gx5CYmlXRlaaXmDh5DTZ4_eZuT-W8FxfJLY");
  var formResponses = form.getResponses();
  var formResponse = formResponses[selection];
  var itemResponses = formResponse.getItemResponses();
  
  var name = "";
  var email = Session.getActiveUser().getEmail() + ", jwaldman@redhat.com";
  var date = "";
  var technology;
  var skills = "";
  var preFeed;
  var bootFeed;
  var HWFeed;
  var shadowFeed;
  
  for (var i = 0; i < itemResponses.length; i++) 
  {
    var itemResponse = itemResponses[i];
    title = itemResponse.getItem().getTitle();
    response = itemResponse.getResponse();
    switch(title)
    {
      case "Who is being scored?":
        name = response;
        break;
      case "What was the start date of the enablement program?":
        date = response;
        break;
      case "Which Emerging Technology Enablement Program did the participant take?":
        technology = response;
        break;
      case "RHEL OpenStack Platform":
        skills = skills + title + ": " + response + "\n";
        break;
      case "RHEL OpenStack - Monitoring":
        skills = skills + title + ": " + response + "\n";
        break;
      case "RHEL OpenStack - Networking":
        skills = skills + title + ": " + response + "\n";
        break;
      case "RHEL OpenStack - Storage":
        skills = skills + title + ": " + response + "\n";
        break;
      case "Red Hat CloudForms":
        skills = skills + title + ": " + response + "\n";
        break;
      case "Red Hat OpenShift":
        skills = skills + title + ": " + response + "\n";
        break;
      case "Red Hat Mobile Application Platform (RH MAP)":
        skills = skills + title + ": " + response + "\n";
        break;
      case "JBoss Fuse":
        skills = skills + title + ": " + response + "\n";
        break;
      case "JBoss A-MQ":
        skills = skills + title + ": " + response + "\n";
        break;
      case "JBoss BPM Suite":
        skills = skills + title + ": " + response + "\n";
        break;
      case "JBoss BRMS":
        skills = skills + title + ": " + response + "\n";
        break;
      case "Prework Feedback":
        preFeed = response;
        break;
      case "Bootcamp Feedback":
        bootFeed = response;
        break;
      case "Homework Project Feedback":
        HWFeed = response;
        break;
      case "Shadowing Feedback":
        shadowFeed = response;
        break;   
    }
  }
  var copyId = DriveApp.getFileById(docTemplate).makeCopy(docName + ' for ' + name).getId();
  var copyDoc = DocumentApp.openById(copyId);
  var copyBody = copyDoc.getActiveSection();
  
  copyBody.replaceText('keyName', name);
  copyBody.replaceText('keyDate', date);
  copyBody.replaceText('keyTechnology', technology);
  copyBody.replaceText('keyPreworkFeedback', preFeed);
  copyBody.replaceText('keyBootcampFeedback', bootFeed);
  copyBody.replaceText('keyHWFeedback', HWFeed);
  copyBody.replaceText('keyShadowFeedback', shadowFeed);
  copyBody.replaceText('keySkillLevel', skills);
  
  copyDoc.saveAndClose();
  var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");
  
  var subject = "Enablement Scorecard";
  var body = "Please find the " + technology + " Enablement Report for " + name + " attached.";
  MailApp.sendEmail(email, subject, body, {htmlBody: body, attachments: pdf});
  DriveApp.getFileById(copyId).setTrashed(true);
  
}
