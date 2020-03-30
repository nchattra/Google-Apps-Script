/*
Automatically send an e-mail with an attached file
***Use with Google spreadsheet***
Column 2: e-mail address
Column 3: subject
Column 4: message
Column 5: attached file name

Example:
Google Form: https://forms.gle/Grp5wj9xAaAVzBBYA (Generate Google Sheet)
Google Sheet: https://docs.google.com/spreadsheets/d/1lM7uDhCXeHW24guTrMFuctGU-f5bKAp-FeYK2o6Z6LA/edit?usp=sharing

Binding GoogleSheet with GoogleApp:
In an active spreadsheet, click 'Tools' -> 'Script Editor'
Replace existing code with the one below

Original Author: Nithiwadee Thaicharoen
Misc Author: Narupon Chattrapiban
*/

function logStudentInfo() {
  var EMAIL_SENT = "EMAIL_SENT";
  var startRow = 2;  // First row of data to process, skip header
  var startCol = 2;  // First col of data to process, skip header
  var numRows = 2;   // Number of rows to process       
  var numCols = 5;   // Number of cols to process  
  var statuscol = startCol + numCols; // Column of 'EMAIL_SENT' status
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols)
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var emailAddress = data[i][0];  // First column
    var subject = data[i][1];       // Second column
    var message = data[i][2];       // Third column
    var attachedFile = data[i][3];  // Fourth column
    // Fifth column, Check if blank, send email.  After email is sent add EMAIL_SENT in this cell
    var emailSent = data[i][4];
    Logger.log('LOGGER');
    Logger.log('Email: ' + emailAddress);
    Logger.log('Subj: ' + subject);
    Logger.log('Mess: ' + message);
    Logger.log('Attch: ' + attachedFile);
    Logger.log('Status: ' + emailSent);
    
    if (emailSent != EMAIL_SENT)  {  // Prevents sending duplicates  
    var folders = DriveApp.getFoldersByName('googleapp00');  // Replace 'googleapp00' with an appropriate folder name
    while(folders.hasNext()){
      Logger.log('Folder: ' + folders.next().getName());
      var files = DriveApp.getFilesByName(attachedFile+'.pdf');
      while(files.hasNext()){
        var file = files.next();
        Logger.log('File: ' + file.getName());
        MailApp.sendEmail(emailAddress, subject, message, {attachments: [file.getBlob()],name: 'Nithiwadee Thaicharoen'})
      }
    }
    sheet.getRange(startRow + i, statuscol).setValue(EMAIL_SENT);
    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();
    }
  }
}

function onOpen() {
 var ss = SpreadsheetApp.getActiveSpreadsheet(),
     options = [
      {name:"Send Mail", functionName:"logStudentInfo"},
     ];
 ss.addMenu("Email Sender", options);
}
