/*Sending Mail From Google SpreadSheet
This code wouldn't mention receiver's names in emails, 
"List" sheet just needed all of the collection of email addresses
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

function getMessage(){
  var sheet = OpenSheet("Mail");
  var oldSheet = SpreadsheetApp.getActiveSheet()
  SpreadsheetApp.setActiveSheet(sheet)                //open and activate "Mail" sheet
  
  var dataRange = sheet.getRange(1,1,1,2)
  data = dataRange.getValues()
  //Logger.log(data)
  var msg = data[0][0]                              //get's mail
  //Logger.log(msg)
  var sub = data[0][1]                              //get's subject
  SpreadsheetApp.setActiveSheet(oldSheet)
  //return msg,sub
  return{
    msg,
    sub
  };
}  

function sendNow(){
  var sheet = OpenSheet("List")
  //Logger.log(sheet.getName())                        //show the sheet name "List"  
  var rows = sheet.getLastRow()
  //Logger.log(rows)                                   //number of rows (here 4)
  var columns = sheet.getLastColumn()
  //Logger.log(columns)                                  //number of column (here 2)
  var dataRange = sheet.getRange(1,1,rows,columns)   //row 2, column 1, up to, last rows, last columns)
  var data = dataRange.getValues();
  for(i in data){
    var to = data[i][0]
    var mail = getMessage()
    //Logger.log(mail.sub, mail.msg)                //show subject and mail
    GmailApp.sendEmail(to,mail.sub,mail.msg)
  }
}
