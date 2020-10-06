function onOpen(){
}

function onEdit(){
  createTrigger();
  checkDeadlineDates();
}


function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; 
  var numRows = sheet.getLastRow();
  var numCol = sheet.getLastColumn();
  var date = new Date(new Date().getFullYear(),new Date().getMonth() , new Date().getDate());
  var dataRange = sheet.getRange(startRow, 1, numRows-1, numCol);
  var data = dataRange.getValues();
  var html =     HtmlService.createTemplateFromFile('Index');
  var template = html.evaluate().getContent();
  
  for (var i in data) {
    var row = data[i];
    var renewaldate = new Date(row[4]);
    var dateinvoice = new Date(row[2]);
    var template = html.evaluate().getContent()
    .replace(/{client}/g, row[0])
    .replace(/{type}/g, row[1])
    .replace(/{renewaldate}/g, (renewaldate.getDate()  + "/" + (renewaldate.getMonth() + 1)+ "/" + renewaldate.getFullYear()))
    .replace(/{contactemail}/g, row[6])
    .replace(/{value}/g, row[3])
    .replace(/{dateinvoice}/g, (dateinvoice.getDate()  + "/" + (dateinvoice.getMonth() + 1)+ "/" + dateinvoice.getFullYear()));
    var rowColour = parseInt(i)+2;
    var dateObj = new Date(new Date().getFullYear(),new Date().getMonth() , new Date().getDate());
    dateObj.setDate(date.getDate() + 7);
    if (row[4].valueOf() === dateObj.valueOf()){
      var emailAddress = row[7];
      var sunumRowsject = row[0]+' Support Will Expire Soon';
    
      MailApp.sendEmail({
         to: emailAddress,
         subject: sunumRowsject,
         htmlBody: template
       });
      sheet.getRange(rowColour, 1, 1, numCol).setBackground("#4285f4");
   }
  }
}

function deadline(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; 
  var numRows = sheet.getLastRow();
  var numCol = sheet.getLastColumn();
  var date = new Date(new Date().getFullYear(),new Date().getMonth() , new Date().getDate());
  var dataRange = sheet.getRange(startRow, 1, numRows-1, numCol);
  var data = dataRange.getValues();
  for (var i in data) {
    var row = data[i];
    var rowColour = parseInt(i)+2;
    if (row[4].valueOf() === date.valueOf()){
      sheet.getRange(rowColour, 1, 1, numCol).setBackground("red");
   }
  }
}

function trigger(){
  sendEmails();
  deadline();
}


function createTrigger(){
   var triggers = ScriptApp.getProjectTriggers();
   var dates = getDates();
   
   for (var i = 0; i < triggers.length; i++) {
     ScriptApp.deleteTrigger(triggers[i]);
  }

   ScriptApp.newTrigger('trigger')
              .timeBased()
              .everyDays(1)
              .atHour(9)
              .create();
              
}

function getDates(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var numCols = sheet.getLastColumn();
  var numRows = sheet.getLastRow(); 
  var dates=[];
   var dataRange = sheet.getRange(2, 1, numRows-1, numCols);
  var data = dataRange.getValues();
  for (var i in data) {
    var row = data[i];
    dates.push(new Date(row[5]));
  }
  return dates;
}

function checkDeadlineDates(){
  var dates = getDates();
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; 
  var numRows = sheet.getLastRow();
  var numCol = sheet.getLastColumn();
  var dataRange = sheet.getRange(startRow, 1, numRows-1, numCol);
  var data = dataRange.getValues();
  var date = new Date(new Date().getFullYear(),new Date().getMonth() , new Date().getDate());
  for (var i in data) {
    var row = data[i];
    var rowColour = parseInt(i)+2;
    var rowSpec = sheet.getRange(rowColour, 1, 1, numCol);
   if ((row[4].valueOf() > date.valueOf()) && (sheet.getRange(rowColour, 1, 1, numCol).getBackground() == "#ff0000")){
     Logger.log("set Green");
      rowSpec.setBackground("green");
      sheet.moveRows(rowSpec, 2);
   }
  }
}
