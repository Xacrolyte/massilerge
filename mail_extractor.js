//Search parameters - Change subject or add keywords 
var SEARCH_QUERY = "label:inbox is:unread from:me to:me";

//Find emails
function getEmails_(q) 
{
    var emails = [];
    
    //Number of searches
    var threads = GmailApp.search(q);
    for (var i in threads) 
    {        
        //Number of messages
        var msgs = threads[i].getMessages();
        Logger.log(threads[i] + "THREAD")
        for (var j in msgs) 
        {
            emails.push([msgs[j].getBody()
                .replace(/<.*?>/g, '\n')
                .replace(/^\s*\n/gm, '')
                .replace(/^\s*/gm, '')
                .replace(/\s*\n/gm, '\n')]);
        }
    }
    return emails;
}

//Append data
function appendData_(sheet, array2d) 
{
    //Paste values/emails 
    sheet.getRange(sheet.getLastRow() + 1, 1, array2d.length, array2d[0].length).setValues(array2d);
}

//Send Email
function sendEmail_()
{
    //Sheet with Emails    
    var textEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmailDump").getRange(1,1).getValue();
    var textEmailString = textEmail.toString();
    var posTopicString = textEmailString.search("Topic");
    var posDateString = textEmailString.search("Date");
    var subString = textEmailString.substring(posTopicString, posDateString+20);
    
    //Sheet with Email IDS = EmailIDS
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmailIDS").activate();
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lr = ss.getLastRow();
    var count = 0;
    for(var i=2; i<=lr; i++)
    {
       //Sheet with Email IDS 
       //Column Email
       var currentEmail = ss.getRange(i, 2).getValue();
       //Column Name
       var currentName = ss.getRange(i, 1).getValue();
       
       var messageBody = subString;
       var subjectLine = "Reminder: " + " is DUE";
       
       count += 1;
       
       Logger.log(count);
       Logger.log(currentName);
       Logger.log(currentEmail);
       Logger.log(subjectLine); 
       Logger.log(messageBody);
         
       MailApp.sendEmail(currentEmail, subjectLine, messageBody);  
    }
 } 

//Clear Page
function clear_() 
{  
      var i = 0;
      var name;
      var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
      var numer = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
      
      for(i=0; i<numer; i++)
      {
            name = sheets[i].getName();
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).activate();
            var spreadsheet = SpreadsheetApp.getActive();
            var sheet = spreadsheet.getActiveSheet();
            
            if(name == "EmailDump")
            {
                  sheet.clear();
            }  
      }
}


//Save email
function saveEmails() 
{
    //Clear
    clear_();
    
    //Find emails
    var array2d = getEmails_(SEARCH_QUERY);
    
    //If found
    if (array2d) 
    {
        //Logger.log(array2d.length() + " STOP");
        //Saved in a spread sheet
        appendData_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmailDump").activate(), array2d);
        Logger.log("FOUND")
    }
    
    sendEmail_();
}