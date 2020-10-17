//
//Variables
//

var label_search = "Vacation Emails"; //Label in Gmail. Used to tag an email as needing to be checked (there are some additional restrictions too: email must be less than 1 day old, in inbox, and unread, etc)
var label_archive = "Vacation Archive"; //Label in Gmail that reprensents an email message that has received a vacation email.
var vacation_from_email = "chris75898@mywork.com"; //Email address that sends the vacation email
var email_domain = "@mywork.com"; //Only emails in the this domain (format: @domain.com) will be checked. All other domains will be allowed.
//List of email addresses to NOT send vacation emails too. These are important people
var allowed_internal_email_addresses = ["cschulte", "cschulte+vm", "john.doe"]
var googleSheetId = "Google Sheet ID"; //A google sheet that can record the vacation emails
//The vacation email to send - Text Version
var vacationEmailText = "Thank you for your email. I am currently working on other projects and will not be checking email. Please put in a help desk ticket. Chris Schulte";;
//The vacation email to send - HTML version
var vacationEmailHTML = "Thank you for your email.<br />I am currently working on other projects and will not be checking email.<br />Please put in a help desk ticket.<br />Chris Schulte";;

/**
 * This is the start function.
 * It is called on a timer.
 */
function mainFunction() 
{
    //Pull the two labels needed for processing
    var vacationEmailLabel = GmailApp.getUserLabelByName(label_search);
    var vacationEmailArchive = GmailApp.getUserLabelByName(label_archive);


    //Search inbox for emails:
    // -> to me, that are unread, and in my inbox
    // -> that are also newer than 1d (new today)
    // -> that also contain the label_search label
    var threads = GmailApp.search("to:me is:unread in:inbox newer_than:1d label:" + label_search.toLowerCase().replace(" ", "-"));
  
    //cycle through all of the email threads
    for (var i=0; i<threads.length; i++)
    {
        var currentThread = threads[i];

        //process the thread
        var processedThread = ProcessEmailThread(currentThread);

        //If any recipients on list, send them vacation emails
        if (processedThread.sendMessagesTo.length > 0)
          sendEmail(processedThread);
    
        //if the thread should be archived
        if (processedThread.archiveEmail)
        {
            currentThread.moveToArchive(); //archive the thread
            currentThread.addLabel(vacationEmailArchive) //add archive label 
        }
        else
        {
            currentThread.moveToInbox(); //move to inbox
            currentThread.removeLabel(vacationEmailLabel); //remove label to prevent future processing
            currentThread.removeLabel(vacationEmailArchive); //this email shouldn't be archived
        }
    }  
}

/**
 * The function that actuall sends emails
 * @param processedThread
 */
function sendEmail(processedThread)
{
  //keep track of the messages that were sent
  var messagesSent = {};
  
  //open the spreadsheet to record the archived emails
  if (googleSheetId)
    var sheet = SpreadsheetApp.openById(googleSheetId).getActiveSheet();

  //cycle through all of the emails that need to be sent
  for(var i=0; i<processedThread.sendMessagesTo.length; i++)
  {
    var currentMessageTo = processedThread.sendMessagesTo[i];

    //check for duplicates
    if (currentMessageTo.from in messagesSent)
      continue;
    messagesSent[currentMessageTo.from] = true;

    //send email
    currentMessageTo.message.message.reply(vacationEmailText, {htmlBody: vacationEmailHTML, from: vacation_from_email})
    //write to sheet
    if (sheet)
      sheet.appendRow([currentMessageTo.from, currentMessageTo.message.message.getSubject()]);
  }
}

/**
 * Functions looks through a single email thread to see:
 * 1) who should receive a reply to email
 * 2) if the thread should be archived
 * 
 * @param emailThread 
 */
function ProcessEmailThread(emailThread)
{
    //return data that contains:
    // -> a list of people who should receive a vacation email
    // -> a boolean to determine if email should be archived
    var returnData = {sendMessagesTo: [], archiveEmail: true};
    var returnData_noChange = {sendMessagesTo: [], archiveEmail: false};

    var alreadyReceivedVM = []; //an array of email addresses that have already received an auto-reply
    var unreadEmails = [];
    
    //pull all of the messages from the thread
    var allMessages = emailThread.getMessages();

    //cycle through each message in thread
    for (var i=0; i<allMessages.length; i++)
    {
        var currentMessage = allMessages[i];

        var messageObject = 
        {
            from: cleanEmail(currentMessage.getFrom().toLowerCase()),
            to: cleanEmail(currentMessage.getTo().toLowerCase()),
            message: currentMessage,
            isUnread: currentMessage.isUnread()
        };
        
        //if the current email is FROM the address that sends vacation emails
        //then this must be an auto reply
        if (messageObject.from == vacation_from_email)
            alreadyReceivedVM.push(messageObject.from);
    
        //if FROM an allowed email address, then whole thread is valid
        if (allowed_internal_email_addresses.filter(p => (p + email_domain) == messageObject.from).length > 0)
          return returnData_noChange;

        //if FROM someone not in organization domain, then whole thread is valid
        if (!messageObject.from.endsWith(email_domain))
          return returnData_noChange;

        //if the message is unread and they haven't received an email yet, 
        if (messageObject.isUnread && !alreadyReceivedVM.includes(messageObject.from))
        {
          //add address to email list
          returnData.sendMessagesTo.push({from: currentEmail.from, message: currentEmail});
          //record that this person has received an email
          alreadyReceivedVM.push(currentEmail.from);
        }
    }
  return returnData;
}

/**
 * Removes any extra bits from an email
 * @param emailString The origional email
 */
function cleanEmail(emailString)
{
  if (emailString.indexOf("<") === -1)
    return emailString;
    
  return emailString.substring(emailString.indexOf("<") + 1, emailString.indexOf(">"));
}
/**
 * Adds the EndsWith function to strings
 */
if (!String.prototype.endsWith) 
{
    String.prototype.endsWith = function(search, this_len) 
    {
        if (this_len === undefined || this_len > this.length) 
			this_len = this.length;
		return this.substring(this_len - search.length, this_len) === search;
	};
}