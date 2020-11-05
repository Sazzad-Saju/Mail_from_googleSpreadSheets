/*Sending Mail From Google SpreadSheet
This code would mention receiver's names in emails, 
That's why "List" sheet must have names for receivers
Code by Sazzad Saju, CSE, HSTU*/


function OpenSheet(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();  //get spreadsheet
  var sheets = ss.getSheets();                     //get the two sheets
  for( var n in sheets){
    if (name==sheets[n].getName()){                //name== List, It'll get only the List sheet, then return the sheet
      return sheets[n];
    }
  }
  return sheets[0];
}

function getMessage(name){
  var sheet = OpenSheet("Mail");
  var oldSheet = SpreadsheetApp.getActiveSheet()
  SpreadsheetApp.setActiveSheet(sheet)                //open and activate "Mail" sheet
  
  var dataRange = sheet.getRange(2,1,1,2)
  data = dataRange.getValues()
  //Logger.log(data)
  var msg = data[0][0]                              //get's mail
  //Logger.log(msg)
  msg = msg.replace("%FIRST%",name)                 //mention each name
  //Logger.log(msg)
  var sub = data[0][1]                              //get's subject
  SpreadsheetApp.setActiveSheet(oldSheet)           //reactivate List sheet
  return{                                           //return as an attribute of an object
    msg,
    sub
  };
}  

function sendNow(){
  var sheet = OpenSheet("List")
  //Logger.log(sheet.getName())                        //show the sheet name "Mail"  
  var rows = sheet.getLastRow()
  //Logger.log(rows)                                   //number of rows (here 4)
  var columns = sheet.getLastColumn()
  //Logger.log(columns)                                  //number of column (here 2)
  var dataRange = sheet.getRange(2,1,rows-1,columns)   //row 2, column 1, up to, last rows, last columns)
  var data = dataRange.getValues();
  //Logger.log(data)                                    //(name,email).... all data shown up
  //Logger.log(data[2][1])   //[0,0] first name, [0,1] first email, [1,0] second name, [1,1] second email, [2,0] third name, [2,1] third email
  for(i in data){
    var name = data[i][0]               //one by one names to var name
    var to = data[i][1]                 //one by one emails to var to
    //GmailApp.sendEmail(to,"Mass Mailing Test","Hi, this is a demonstration. Thanks") //This will also send this mail to all
    //getMessage(name)
    var mail = getMessage(name)          //gets mail and subject to object mail
    //Logger.log(mail.sub, mail.msg)
    GmailApp.sendEmail(to,mail.sub,mail.msg)
  }
}