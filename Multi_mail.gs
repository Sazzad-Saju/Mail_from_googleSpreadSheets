function OpenSheet(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for( var n in sheets){
    if (name==sheets[n].getName()){
      return sheets[n];
    }
  }
  return sheets[0];
}

function getMessage(name){
  var sheet = OpenSheet("Mail");
  var oldSheet = SpreadsheetApp.getActiveSheet()
  SpreadsheetApp.setActiveSheet(sheet)  
  var dataRange = sheet.getRange(2,1,1,2)
  data = dataRange.getValues()
  var msg = data[0][0]
  msg = msg.replace("%FIRST%",name)
  var sub = data[0][1]
  SpreadsheetApp.setActiveSheet(oldSheet)
  return{
    msg,
    sub
  };
}  

function sendNow(){
  var sheet = OpenSheet("List")
  var rows = sheet.getLastRow()
  var columns = sheet.getLastColumn()
  var dataRange = sheet.getRange(2,1,rows-1,columns)
  var data = dataRange.getValues();
  for(i in data){
    var name = data[i][0]
    var to = data[i][1]
    var mail = getMessage(name)
    GmailApp.sendEmail(to,mail.sub,mail.msg)
  }
}