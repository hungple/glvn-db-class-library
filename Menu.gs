/**
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Create HK1 Report Cards",
    functionName : "createReports_1"
  },
  {
    name : "Create and Email HK1 Report Cards",
    functionName : "emailReports_1"
  },
  {
    name : "Create HK2 Report Cards",
    functionName : "createReports_2"
  },
  {
    name : "Create and Email HK2 Report Cards",
    functionName : "emailReports_2"
  }
  ];
  sheet.addMenu("GLVN", entries);
};


function getReportCardTemplateId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  return sheet.getRange("B2:B2").getCell(1, 1).getValue();
}


function getReportCardFolderId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
  return sheet.getRange("B3:B3").getCell(1, 1).getValue();
}




function createReports_1() {
  createDoc(false, false);
}

function emailReports_1() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to email to the report cards to the parents?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    createDoc(false, true);
  }
}

function createReports_2() {
  createDoc(true, false);
}

function emailReports_2() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
      'Warning!!!',
      'Do you want to email to the report cards to the parents?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    createDoc(true, true);
  }
}

function createDoc(isHK2, sendEmail) {

  var DECIMAL_COL_LEN = 5; //IMPORTANT: any point column must has name with lenght = 5. For example: Part1 or Hwrk1

  var idCol          = 1;
  // column 2 is saintName which is used for honor certificates
  var fNameCol       = 3;
  var lNameCol       = 4;
  var pEMailCol      = 5;
  var totalPointsCol = 6;
  var actionCol      = 7;
  var pointBeginCol  = 8;
  
  var colNames = [];
  var colPoints = [];
  var colPointsMax = [];
  
  var folerId     = getReportCardFolderId();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("grades");
  var range = sheet.getRange(1, 1, 50, 30); //row, col, numRows, numCols
  //////////////////////////////////////////////////////////////////////////////
  
  var cName = range.getCell(1, 1).getValue();
  var tName = range.getCell(1, 2).getValue();
  var tmpName     = "HK1-Report-Card";
  if(isHK2) {
    tmpName     = "HK2-Report-Card";
  }
  
  // iterate through all column names
  for (var cellCol = pointBeginCol; ; cellCol++) {
    var colPointMax = range.getCell(1, cellCol).getValue();
    var colName = range.getCell(2, cellCol).getValue();
    if(colName == "") { break; }
    colNames[cellCol-pointBeginCol] = colName;
    colPointsMax[cellCol-pointBeginCol] = colPointMax;
  }
  
  Logger.log(colNames);
  
  var halfCol = colNames.length/2;

  var docName, id, fName, lName, email, action;
  
  // iterate through all rows in the range
  for (var cellRow = 3; ; cellRow++) {
    
    id = range.getCell(cellRow, idCol).getValue(); 
    if(id == "") { break; }
    
    fName = range.getCell(cellRow, fNameCol).getValue(); 
    lName = range.getCell(cellRow, lNameCol).getValue(); 
    email = range.getCell(cellRow, pEMailCol).getValue().trim();
    action = range.getCell(cellRow, actionCol).getValue();
    
    Logger.log(fName + ' ' + lName);
    
    if(action == "x") {
    
      if(email == "" || email.length < 6) {
        docName = '--no-email-' + cName + '-' + fName + '-' + lName + '-' + id + '-' + tmpName;
      }
      else {
        docName = cName + '-' + fName + '-' + lName + '-' + id + '-' + tmpName;
      }
    
      // iterate throught all cell/points
      for (var cCol = 0; cCol<colNames.length; cCol++) {
        colPoints[cCol] = range.getCell(cellRow, cCol+8).getValue();
      }      
      Logger.log(colPoints);
      
      var formId      = getReportCardTemplateId();
      Logger.log(formId);
      
      // Get document template, copy it as a new temp doc, and save the Doc’s id
      var copyId = DriveApp.getFileById(formId).makeCopy(docName).getId();

      // Open the temporary document
      var copyDoc = DocumentApp.openById(copyId);

      Logger.log(copyDoc.getName());

      // Get the document’s body section
      var copyBody = copyDoc.getActiveSection();

      // Replace place holder keys,in our google doc template
      copyBody.replaceText('@cname@', cName);
      copyBody.replaceText('@tname@', tName);
      copyBody.replaceText('@sname@', fName + ' ' + lName);
      
      
      
      // HK1 - fill in data for HK1
      var hk1Total = 0;
      for (var i = 0; i<halfCol-1; i++) {
        if(colNames[i].length == DECIMAL_COL_LEN) {
          copyBody.replaceText('@' + colNames[i] + '@', getLetterGrade(colPoints[i],colPointsMax[i]));
          hk1Total = hk1Total + parseFloat(colPoints[i]);
        }
        else {
          copyBody.replaceText('@' + colNames[i] + '@', colPoints[i]);
        }
      }
      copyBody.replaceText('@Total1@', getLetterGrade(hk1Total, 50));
      var c1 = colPoints[halfCol-1];
      if(c1.length < 70) {
        c1 = c1 + "\n\n";
      }
      else if(c1.length < 140) {
        c1 = c1 + "\n";
      }
      
      copyBody.replaceText('@Comment1@', c1);

      // HK2
      var hk2Total = 0;
      if(isHK2) { // fill in data for HK2
        for (var i = halfCol; i<colNames.length-1; i++) {
          if(colNames[i].length == DECIMAL_COL_LEN) {
            copyBody.replaceText('@' + colNames[i] + '@', getLetterGrade(colPoints[i],colPointsMax[i]));
            hk2Total = hk2Total + parseFloat(colPoints[i]);
          }
          else {
            copyBody.replaceText('@' + colNames[i] + '@', colPoints[i]);
          }
        }
        copyBody.replaceText('@Total2@', getLetterGrade(hk2Total, 50));
        copyBody.replaceText('@text1@', "Nhận xét của Giáo Lý Viên - Teacher's Comment:");
        var c2 = colPoints[colNames.length-1];
        if(c2.length < 70) {
          c2 = c2 + "\n\n";
        }
        else if(c2.length < 140) {
          c2 = c2 + "\n";
        }
        
        copyBody.replaceText('@Comment2@', c2);
        copyBody.replaceText('@text2@', "Chử ký của Giáo Lý Viên - Teacher's Signature:_______________________________");
        
        // fill in data for Yearly Total
        for (var i = 0; i<halfCol-1; i++) {
          if(colNames[i].length == DECIMAL_COL_LEN) {
            copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', getLetterGrade(colPoints[i]+colPoints[i+halfCol],colPointsMax[i]*2));
          }
          else {
            copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', (colPoints[i]+colPoints[i+halfCol]));
          }
        }
        copyBody.replaceText('@Total3@', getLetterGrade(hk1Total + hk2Total, 100));
      }
      else { // fill in '-' for HK2
        for (var i = halfCol; i<colNames.length-1; i++) {
          copyBody.replaceText('@' + colNames[i] + '@', '-');
        }
        copyBody.replaceText('@Total2@', '-');
        copyBody.replaceText('@text1@', '');
        copyBody.replaceText('@Comment2@', "\n\n\n");
        copyBody.replaceText('@text2@', '');
        
        // fill in '-' for Yearly Total
        for (var i = 0; i<halfCol-1; i++) {
          copyBody.replaceText('@' + colNames[i].substring(0,colNames[i].length-1) + '3@', '-');
        }
        copyBody.replaceText('@Total3@', '-');
      }
      

      
      // Save and close the temporary document
      copyDoc.saveAndClose();

      // Convert temporary document to PDF
      var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");
 
      // Delete temp file
      DriveApp.getFileById(copyId).setTrashed(true);

      // Delete old file
      var files = DriveApp.getFolderById(folerId).getFilesByName(docName + ".pdf");
      while (files.hasNext()) {
        var file = files.next();
        if(file.getOwner().getEmail() == Session.getActiveUser()) {
          file.setTrashed(true); 
        }        
      }

      // Save pdf
      DriveApp.getFolderById(folerId).createFile(pdf);

      // Send email
      if(sendEmail == true && email != "" && email.length > 5) {
        // Attach PDF and send the email
        var subject = docName;
        var body = "Mến chào quí Phụ Huynh,<br>Xin phụ huynh xem phiếu báo điểm đính kèm. Xin cám ơn.<br>Ban GLVN.";
        //email = "hle007@yahoo.com";
        MailApp.sendEmail(email, subject, body, {htmlBody: body, attachments: pdf});
      }
    }
  }
}


function getLetterGrade(points, maxPoint) {
  if(points/maxPoint*100 > 89.99) {
    return "A";
  }
  else if(points/maxPoint*100 > 79.99) {
    return "B";
  }
  else if(points/maxPoint*100 > 69.99) {
    return "C";
  }
  else {
    return "D";
  }
}