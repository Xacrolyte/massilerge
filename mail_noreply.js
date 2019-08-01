function emailNotify()
{

  //Basic Initialization
  //Change name to the name of your Spreadsheet
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hypercare").activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var count  = 0;
  
  //Basic Variables
  var i = 0;
  var m=0;
  var name;
  var sheets = SpreadsheetApp.getActive().getSheets();
  var numer = SpreadsheetApp.getActive().getNumSheets();
  
  //Check for Sheet Existence
  for(i=0; i<numer; i++)
  {
      name = sheets[i].getName();      
      if(name == "Template")
      {
          m = 1;
          break;
      }  
    
      else 
      {
          continue;
      }
  }
  
  //If sheet Exists
  if(m == 1)
  {
    SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName('Template'), true);
  }
  else
  {
    //Set message accordingt to your Needs.
    SpreadsheetApp.getActive().insertSheet(SpreadsheetApp.getActive().getActiveSheet().getIndex() + 1).activate();
    SpreadsheetApp.getActive().getActiveSheet().setName('Template');
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1, 1).setValue("Hello {name},\n\nThis is a Reminder that the {application} raised by {raise} with GIS stakeholder {GIS}, is DUE in {days} days.\nNature of Issue: {noi}\n\nBest Regards,\nXacrolyte");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(2, 2).setValue("Hello {name},\n\nThis is a Reminder that the {application} raised by {raise} with GIS stakeholder {GIS}, is DUE in {days} days.\nNature of Issue: {noi}\n\nBest Regards,\nXacrolyte");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(3, 3).setValue("Hello {name},\n\nThis is a Reminder that the {application} assigned to {assign}, raised by {raise} with GIS stakeholder {GIS} is DUE in {days} days.\nNature of Issue: {noi}\n\nBest Regards,\nXacrolyte");
  }
 
  //Template Init
  var templateTextBuild = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1, 1).getValue();
  var templateZTextBuild = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(2, 2).getValue();
  var templateY1TextBuild = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(3, 3).getValue();
  var templateY2TextBuild = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(3, 3).getValue();
  
  
  //Qouta var and check  
  var qoutaLeft = MailApp.getRemainingDailyQuota();
  //Logger.log("Dailt qouta left = " + qoutaLeft);
  
  //If qouta is not available
  if((lr-1)>qoutaLeft)
  {
    Browser.msgBox('You have ' + qoutaLeft + ' mail qouta left and you are trying to send '  + (lr-1) + ' mails');
  }
  
  //If qouta is available
  else
  {
    
    //Sends Mails to all
    for(var i=2; i<=lr; i++)
    {
     
     //Raised On Date
     var raisedDate = ss.getRange(i, 4).getValue();
     //Logger.log(raisedDate);
     
     //Planned Closure Date
     var plannedCDate = ss.getRange(i, 8).getValue();
     //Logger.log(plannedCDate);
     
     //Actual Closure Date
     var actualCDate = ss.getRange(i, 9).getValue();
     //Logger.log(actualCDate + "ACTUAL");
     
     //Current Date
     var now = new Date();
     
     //Difference
     var diff = plannedCDate - now;
     var dayDiff = parseInt(diff * 0.00000001157407);
     //Logger.log(diff);
     //Logger.log(dayDiff);
     
     //If Planned-Closure Date is approaching
	if(dayDiff>0 && actualCDate == "")
	{
       Logger.log("Number of days passed = " + dayDiff);
       
       //Column Assigned To Email
       var currentAssignedEmail = ss.getRange(i, 10).getValue();
       //Column Assigned To
       var currentAssigned = ss.getRange(i, 10).getValue();
       
       //Column Application Name
       var currentApplication = ss.getRange(i, 1).getValue();
       
       //Column Raised By Email
       var currentRaisedNameEmail = ss.getRange(i, 2).getValue();
       //Column Raised By
       var currentRaisedName = ss.getRange(i, 2).getValue();
       
       //Column GIS Stakeholder Email
       var currentGISNameEmail = ss.getRange(i, 3).getValue();
       //Column GIS Stakeholder
       var currentGISName = ss.getRange(i, 3).getValue();
       
       //Column Nature of Issue
       var currentNOI = ss.getRange(i, 6).getValue();
       
       //Column Issue Decription
       var currentDesc = ss.getRange(i, 5).getValue();
       
       var messageBody = templateTextBuild.replace("{name}", currentAssigned)
       .replace("{application}",currentApplication)
       .replace("{raise}",currentRaisedName)
       .replace("{GIS}",currentGISName)
       .replace("{days}",dayDiff)
       .replace("{noi}",currentNOI);
       
       var subjectLine = "Reminder: "+ currentApplication + " is DUE";
       
       count += 1;
       
       Logger.log(count);
       Logger.log(subjectLine); 
       Logger.log(messageBody);
         
       MailApp.sendEmail(currentAssignmentEmail, subjectLine, messageBody); 
     }//End If 
     
     else if(dayDiff <= 0 && actualCDate == "" )
     {
       
       Logger.log("Number of days passed = " + dayDiff);
       
       //Column Assigned To Email
       var currentAssignedEmail = ss.getRange(i, 10).getValue();
       //Column Assigned To
       var currentAssigned = ss.getRange(i, 10).getValue();
       
       //Column Application Name
       var currentApplication = ss.getRange(i, 1).getValue();
       
       //Column Raised By Email
       var currentRaisedNameEmail = ss.getRange(i, 2).getValue();
       //Column Raised By
       var currentRaisedName = ss.getRange(i, 2).getValue();
       
       //Column GIS Stakeholder Email
       var currentGISNameEmail = ss.getRange(i, 3).getValue();
       //Column GIS Stakeholder
       var currentGISName = ss.getRange(i, 3).getValue();
       
       //Column Nature of Issue
       var currentNOI = ss.getRange(i, 6).getValue();
       
       //Column Issue Decription
       var currentDesc = ss.getRange(i, 5).getValue();
       
       //Mail to Assigned  
       var messageBody = templateZTextBuild.replace("{name}", currentAssigned)
       .replace("{application}",currentApplication)
       .replace("{raise}",currentRaisedName)
       .replace("{GIS}",currentGISName)
       .replace("{days}",dayDiff)
       .replace("{noi}",currentNOI);
       var subjectLine = "Reminder: "+ currentApplication + " is DUE";
       count += 1;
       
       Logger.log(count);
       Logger.log(subjectLine); 
       Logger.log(messageBody);

       MailApp.sendEmail(currentAssignedEmail, subjectLine, messageBody); 
       
       //Mail to GIS Stakeholder
       var messageBody2 = templateY2TextBuild.replace("{name}", currentGISName)
       .replace("{application}",currentApplication)
       .replace("{assign}", currentAssigned)
       .replace("{raise}",currentRaisedName)
       .replace("{GIS}",currentGISName)
       .replace("{days}",dayDiff)
       .replace("{noi}",currentNOI);
       
       Logger.log(count);
       Logger.log(subjectLine); 
       Logger.log(messageBody2);
       
       MailApp.sendEmail(currentGISNameEmail, subjectLine, messageBody1); 
       
       //Mail to Rasied Name
       var messageBody1 = templateY1TextBuild.replace("{name}", currentRaisedName)
       .replace("{application}",currentApplication)
       .replace("{assign}", currentAssigned)
       .replace("{raise}",currentRaisedName)
       .replace("{GIS}",currentGISName)
       .replace("{days}",dayDiff)
       .replace("{noi}",currentNOI);
       
       Logger.log(count);
       Logger.log(subjectLine); 
       Logger.log(messageBody1);
       
       //MailApp.sendEmail(currentRaisedNameEmail, subjectLine, messageBody2); 
       
     }

    }
    
  }

}