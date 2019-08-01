function sendEmail() 
{
  
  //Basic Initialization
  //Change name to the name of your Spreadsheet
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails").activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = ss.getLastRow();
  var count = 0;
  
  //Color Variables
  var color = "#ccc";
  var color1 = "#2fb726";
  var color2 = "#FFFF00";

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
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1, 1).setValue("Hello {name},\n\nThis is a Reminder that your {class} is DUE for build in {days} days.\n\nBest Regards,\nXacrolyte");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1, 2).setValue("Hello {name},\n\nThis is a Reminder that your {class} is OVERDUE for build in {days} days.\n\nBest Regards,\nXacrolyte");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(2, 1).setValue("Hello {name},\n\nThis is a Reminder that your {class} is finished and you MAY start the UAT.\n\nBest Regards,\nXacrolyte");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(2, 2).setValue("Hello {name},\n\nThis is a Reminder that your {class} is finished and waiting UAT for {days} days.\n\nBest Regards,\nXacrolyte");
  }

  
  //Template initialization
  var templateTextBuild = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1, 1).getValue();
  var templateTextOBuild = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1, 2).getValue();
  var templateTextMUAT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(2, 1).getValue();
  var templateTextUAT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(2, 2).getValue();
  
  //Qouta var and check  
  var qoutaLeft = MailApp.getRemainingDailyQuota();
  Logger.log("Dailt qouta left = " + qoutaLeft);

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
      
      //State Var
      var currentState = ss.getRange(i, 5).getValue();
      
      //St-Date var
      var stDate = ss.getRange(i, 11).getValue();
      var stformatDate = new Date(stDate);
      
      //End-Date var
      var endDate = ss.getRange(i, 12).getValue();
      var endformatDate = new Date(endDate);
      
      //Current-Date var
      var now = new Date();
      
      //Comp-Date var
      var compDate = ss.getRange(i, 13).getValue();
      var compformatDate = new Date(compDate);

      //UATComp-Date var
      var UATcompDate = ss.getRange(i, 14).getValue();
      var UATcompformatDate = new Date(UATcompDate);
      
      //Difference in Date   
      var diff = (endformatDate.valueOf() - now.valueOf());
      var dayDiff = parseInt(diff * 0.00000001157407);
      
      if (currentState == "Complete")
      {
        var setColor1 = ss.getRange(i, 5).setBackground(color1);
      }
      //If Comp-Date is not Filled
      else if (compDate == "" && UATcompDate=="") 
      {
        //If End-Date is Approaching
        if (endformatDate.valueOf() > now.valueOf()) 
        {
          
          //If End-Date approaching in 7 days
          if (dayDiff<8 && dayDiff>0) //DUE
          {
            
            Logger.log("Number of days left = " + dayDiff);
        
            var currentEmail = ss.getRange(i, 24).getValue();
          
            var currentClassTitle = ss.getRange(i, 2).getValue();
            var currentName = ss.getRange(i, 23).getValue();
          
            var messageBody = templateTextBuild.replace("{name}", currentName).replace("{class}",currentClassTitle).replace("{days}",dayDiff);
            var subjectLine = "Reminder: "+ currentClassTitle + " DUE";
            count += 1;
            Logger.log(count);
            Logger.log(subjectLine); 
            Logger.log(messageBody);
            
            var setColor1 = ss.getRange(i, 5).setBackground(color2);
              
            MailApp.sendEmail(currentEmail, subjectLine, messageBody);    
        
          }//dayDiff if

        }//End-Date if

        else if (endformatDate.valueOf() <= now.valueOf())//OVERDUE
        {  
          Logger.log("Number of days added to delay = " +dayDiff);
      
          var currentEmail = ss.getRange(i, 24).getValue();
        
          var currentClassTitle = ss.getRange(i, 2).getValue();
          var currentName = ss.getRange(i, 23).getValue();
        
          var messageBody = templateTextOBuild.replace("{name}", currentName).replace("{class}",currentClassTitle).replace("{enddate}",endformatDate).replace("{days}",dayDiff);
          var subjectLine = "Reminder: "+ currentClassTitle + " OVERDUE";
        
          count += 1;
          Logger.log(count);
          Logger.log(subjectLine); 
          Logger.log(messageBody);
            
          MailApp.sendEmail(currentEmail, subjectLine, messageBody);  
            
        }//End-Date else

      }//Comp-Date if

      //If Comp-Date is Filled
      else 
      {

        //If UATcomp-Date is not present
        if(UATcompDate == "")
        {

          //If compDate has passed endDate
          if (compDate.valueOf() >= endDate.valueOf()) //UAT
          {

            Logger.log("Number of days added to delay = " + dayDiff);
      
            var currentEmail = ss.getRange(i, 22).getValue();
          
            var currentClassTitle = ss.getRange(i, 2).getValue();
            var currentName = ss.getRange(i, 23).getValue();
          
            var messageBody = templateTextUAT.replace("{name}", currentName).replace("{class}",currentClassTitle).replace("{days}",(dayDiff*(-1)));
            var subjectLine = "Reminder: "+ currentClassTitle + " UAT";
          
            count += 1;
            Logger.log(count);
            Logger.log(subjectLine); 
            Logger.log(messageBody);
              
            MailApp.sendEmail(currentEmail, subjectLine, messageBody);

          }//Comp-Date present-if

          else //MAY_UAT
          {

            Logger.log("Number of days left = " +dayDiff);
      
            var currentEmail = ss.getRange(i, 22).getValue();
          
            var currentClassTitle = ss.getRange(i, 2).getValue();
            var currentName = ss.getRange(i, 23).getValue();
          
            var messageBody = templateTextMUAT.replace("{name}", currentName).replace("{class}",currentClassTitle).replace("{days}",(dayDiff*(-1)));
            var subjectLine = "Reminder: "+ currentClassTitle + " MAY-UAT";
          
            count += 1;
            Logger.log(count);
            Logger.log(subjectLine); 
            Logger.log(messageBody);
              
            MailApp.sendEmail(currentEmail, subjectLine, messageBody);

          }//Comp-Date present-else

        }//UATcomp-Date if
        
        else 
        {
          
          var setColor = ss.getRange(i, 5).setBackground(color);
          
        }
      
      }//Comp-date else

    }//For-loop        

  }//Qouta

}


