//Global settings
const debugMode = false; //Set this to true to send emails to the replyTo address not the one in the sheet
const replyTo = "tzarr@salesforce.com";
const productName = "Salesforce Order Management";
const eventName = "How to Deliver";
const partnerCommunity = "https://partners.salesforce.com/";
const b2bLexChatterGroup = "https://partners.salesforce.com/_ui/core/chatter/groups/GroupProfilePage?g=0F93A000000LovT";
const somChatterGroup = "https://partners.salesforce.com/_ui/core/chatter/groups/GroupProfilePage?g=0F93A000000HZ5o";

//Map all columns by index as it makes getting older data when needed much easier
const emailColumnIndex = 1;
const partnerCompanyNameColumnIndex = 2;
const firstNameColumnIndex = 3;
const lastNameColumnIndex = 4;
const countryColumnIndex = 5;
const activeOpportunityColumnIndex = 6;
const prerequisitesCompleteColumnIndex = 7;
const requestedTrainingDateColumnIndex = 8;

function onFormSubmit(e) {
  //Get the row of submitted data by names and other means where needed
  let submission = { 
    emailAddress: e.values[emailColumnIndex]
    , partnerCompanyName: e.values[partnerCompanyNameColumnIndex]
    , requestedTrainingDate : e.values[requestedTrainingDateColumnIndex]
    , firstName: e.values[firstNameColumnIndex]
    , lastName: e.values[lastNameColumnIndex]
    , country: e.values[countryColumnIndex]
    , activeOpportunity: e.values[activeOpportunityColumnIndex]
    , prerequisitesComplete: e.values[prerequisitesCompleteColumnIndex]
  };
  
  //Avoid mistakes or spamming when first assigning this code to your form using the debugMode variable
  if(debugMode){
    submission.emailAddress = replyTo;  
  }
  
  //Check if we have a record on file for the email address and same session time
  let alreadyRegisteredResult = isDuplicateRegistration(submission.emailAddress, submission.requestedTrainingDate);
  
  if(alreadyRegisteredResult.alreadyRegistered === true){
    if(submission.requestedTrainingDate === 'Self-Enablement (Partner Learning Camp Migration)')
    {
      MailApp.sendEmail({
      to: submission.emailAddress
      , replyTo: replyTo
      , subject: "Duplicate registration request for " + productName + ": " + eventName
      , htmlBody: "Dear " + submission.firstName + " " + submission.lastName + ", <br /><br />Your request for <b>" + productName + ": " + eventName + "</b> via "
      + submission.requestedTrainingDate + " has been received.<br /><br /><b><u>Please Note:</u></b> You can start consuming the content immediately. There is no approval process for self-enablement.<br /><br />"
      + "The registration information captured follows:<br /><br />"
      + makeRegistrationTable(submission) + "<br /><br />"
      + "If you do not have access to the <a href=\"" + partnerCommunity + "\">Partner Community</a> or appropriate  <a href=\"" + somChatterGroup + "\">Chatter Group</a> you should work with your Partner Admin to establish that as soon as possible."
    });
    }
    else if(submission.requestedTrainingDate === 'Office Hours (Partner Learning Camp Migration)')
    {
      MailApp.sendEmail({
      to: submission.emailAddress
      , replyTo: replyTo
      , subject: "Duplicate registration request for " + productName + ": " + eventName
      , htmlBody: "Dear " + submission.firstName + " " + submission.lastName + ", <br /><br />Your request for <b>" + productName + ": " + eventName + "</b> "
      + submission.requestedTrainingDate + " has been received.<br /><br /><b><u>Please Note:</u></b><br /><br />"
      + "The registration information captured follows:<br /><br />"
      + makeRegistrationTable(submission) + "<br /><br />"
      + "Please allow 48 hours for us to respond. If you do not hear back in 48 hours please send an email to Tom Zarr (tzarr@salesforce.com) including your time zone and available windows for a 1 hour appointment.<br /><br />"
      + "If you do not have access to the <a href=\"" + partnerCommunity + "\">Partner Community</a> or appropriate  <a href=\"" + somChatterGroup + "\">Chatter Group</a> you should work with your Partner Admin to establish that as soon as possible."
      });
    }
    else{
      MailApp.sendEmail({
        to: submission.emailAddress
        , replyTo: replyTo
        , subject: "Duplicate registration request for " + productName + ": " + eventName
        , htmlBody: "Salutations " + submission.firstName + " " + submission.lastName + ", <br /><br />Your request for <b>" + productName + ": " + eventName + "</b> training on " + submission.requestedTrainingDate + " appears to be a duplicate.<br/><br />Please see the prior registration submission below for your records:<br /><br />"
        + makeRegistrationTable(alreadyRegisteredResult.priorData) + "<br /><br />"
        + "<b><u>Please Note:</u></b> You still need to be approved for the session before you can attend.<br /><br />"
        + "Please allow us 2-3 business days to approve your request and respond. You will receive a follow-up communication regarding approval status and if approved, more information about environments and any related pre-work about 1 week prior to the session.<br /><br />"
        + "If you do not have access to the <a href=\"" + partnerCommunity + "\">Partner Community</a> or appropriate <a href=\"" + somChatterGroup + "\">Chatter Group</a> you should work with your Partner Admin to establish that as soon as possible."
      });
      Logger.log("Sent Duplicate registration request for " + productName + ": " + eventName + " to <" + submission.emailAddress + ">");
    }
  }
  else{
    //NOTE: If your form allows editing, you will get an exception for not having a recipient here when editing - so let's not allow editing ;)
    //Send Ack email to the person registering
    if(submission.requestedTrainingDate === 'Self-Enablement (Partner Learning Camp Migration)')
    {
      MailApp.sendEmail({
        to: submission.emailAddress
        , replyTo: replyTo
        , subject: "Registration request for " + productName + ": " + eventName + " self-enablement training acknowledged"
        , htmlBody: "Dear " + submission.firstName + " " + submission.lastName + ", <br /><br />Your request for <b>" + productName + ": " + eventName + "</b> via "
        + submission.requestedTrainingDate + " has been received.<br /><br /><b><u>Please Note:</u></b> You can start consuming the content immediately. There is no approval process for self-enablement.<br /><br />"
        + "The registration information captured follows:<br /><br />"
        + makeRegistrationTable(submission) + "<br /><br />"
        + "If you do not have access to the <a href=\"" + partnerCommunity + "\">Partner Community</a> or appropriate  <a href=\"" + somChatterGroup + "\">Chatter Group</a> you should work with your Partner Admin to establish that as soon as possible."
      });
    }
    else if(submission.requestedTrainingDate === 'Office Hours (Partner Learning Camp Migration)')
    {
      MailApp.sendEmail({
      to: submission.emailAddress
      , replyTo: replyTo
      , subject: "Registration request for " + productName + ": " + eventName + " office hours acknowledged"
      , htmlBody: "Dear " + submission.firstName + " " + submission.lastName + ", <br /><br />Your request for <b>" + productName + ": " + eventName + "</b> "
      + submission.requestedTrainingDate + " has been received.<br /><br /><b><u>Please Note:</u></b><br /><br />"
      + "The registration information captured follows:<br /><br />"
      + makeRegistrationTable(submission) + "<br /><br />"
      + "Please allow 48 hours for us to respond. If you do not hear back in 48 hours please send an email to Tom Zarr (tzarr@salesforce.com) including your time zone and available windows for a 1 hour appointment.<br /><br />"
      + "If you do not have access to the <a href=\"" + partnerCommunity + "\">Partner Community</a> or appropriate  <a href=\"" + somChatterGroup + "\">Chatter Group</a> you should work with your Partner Admin to establish that as soon as possible."
      });
    }
    else
    {
      MailApp.sendEmail({
        to: submission.emailAddress
        , replyTo: replyTo
        , subject: "Registration request for " + productName + ": " + eventName + " training acknowledged"
        , htmlBody: "Dear " + submission.firstName + " " + submission.lastName + ", <br /><br />Your request for <b>" + productName + ": " + eventName + "</b> training on "
        + submission.requestedTrainingDate + " has been received.<br /><br /><b><u>Please Note</u>:</b> You still need to be approved for the session before you can attend.<br /><br />"
        + "The registration information captured follows:<br /><br />"
        + makeRegistrationTable(submission) + "<br /><br />"
        + "Please allow us 7 days prior to the event's commencement to approve your request and respond. You will receive a follow-up communication regarding approval status and if approved, more information about environments and any related pre-work prior to the session.<br /><br />"
        + "If you do not have access to the <a href=\"" + partnerCommunity + "\">Partner Community</a> or appropriate  <a href=\"" + somChatterGroup + "\">Chatter Group</a> you should work with your Partner Admin to establish that as soon as possible."
      });
    }
    Logger.log("Sent Registration acknowledgement for " + productName + ": " + eventName + " to <" + submission.emailAddress + ">");
  }
}

//Make fungible table of name-value pairs in html
function makeRegistrationTable(data)
{
  let registrationTable = "";
  registrationTable += "<b>Training Type</b>: " + productName + ": " + eventName  + "<br />";
  registrationTable += "<b>Requested Date</b>: " + data.requestedTrainingDate + "<br />";
  registrationTable += "<b>First</b>: " + data.firstName + "<br />";
  registrationTable += "<b>Last</b>: " + data.lastName + "<br />";
  registrationTable += "<b>Country</b>: " + data.country + "<br />";
  registrationTable += "<b>Partner Company Name</b>: " + data.partnerCompanyName + "<br />";
  registrationTable += "<b>Active Opportunity</b>: " + data.activeOpportunity + "<br />";
  registrationTable += "<b>Prerequisites Completed</b>: " + data.prerequisitesComplete + "<br />";
  return registrationTable;
}

//Detect a duplicate based on the email address and training slot
function isDuplicateRegistration(email, requestedTrainingDate) {
  //Logger.log("isDuplicateRegistration(" + email + "," + requestedTrainingDate + ")");
  //The current submission does not get stopped so we need to just track it and kill the real dup looking at any matches added beyond a length of 1
  let hits = [];
  let sheet = SpreadsheetApp.getActiveSheet();
  let data  = sheet.getDataRange().getValues();
  let i = 0;
        
  for (i = 0; i < data.length; i++) {
    if (data[i][emailColumnIndex].toString().toLowerCase() === email.toLowerCase() && data[i][requestedTrainingDateColumnIndex].toString().toLowerCase() === requestedTrainingDate.toLowerCase()){
      Logger.log("Pushing: " + JSON.stringify(data[i]));  
      hits.push(data[i]);
    }
  }
  
  //Return enriched results we can use  
  if(hits.length > 1)
  {
    Logger.log("Located duplicate email: <" + email + "> with time slot: " + requestedTrainingDate.toLowerCase());
    //Remove the newest addition using the tracking array (like it never happened)
    sheet.deleteRow(i);
    //Get the data for the original (old) submission into something we can use
    let originalRow = hits[hits.length - 2];
    
    let priorRegistration = {
      partnerCompanyName: originalRow[partnerCompanyNameColumnIndex]
      , requestedTrainingDate : originalRow[requestedTrainingDateColumnIndex]
      , firstName: originalRow[firstNameColumnIndex]
      , lastName: originalRow[lastNameColumnIndex]
      , country: originalRow[countryColumnIndex]
      , activeOpportunity: originalRow[activeOpportunityColumnIndex]
      , prerequisitesComplete: originalRow[prerequisitesCompleteColumnIndex]
    };
    
    return JSON.parse("{\"alreadyRegistered\": true, \"priorData\": " + JSON.stringify(priorRegistration) +"}")
  }
  else
  {
    return JSON.parse("{\"alreadyRegistered\": false, \"priorData\": null}");
  }
}
