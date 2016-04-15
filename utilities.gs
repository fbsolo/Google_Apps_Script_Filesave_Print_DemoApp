//  Article
//
//    http://bitvectors.blogspot.com/2016/04/page1.html
//
//  references and describes this application . . .

function returnScratchpadFileCollection() {

  // Called by Code.gs

  // For "this user" and "this app instance", build a unique
  // scratchpad filename based on date-time stamp . . .

///////////////////////////////
//
//// This machinery involved placing scratchpad files in
//// a defined scratchpad file folder, itself a child of
//// the root. The scratchpad file use the email address
//// of the active user as part of its name. I had to
//// drop this idea because the machinery won't work
//// if multiple, different Google accounts use the app.
//// Specifically, this machinery will not dependably
//// catch the email address of non-owner users . . .
//
//  var scratchpadFolderCollection = DriveApp.getFoldersByName('scratchpadFolder');
//
//// Create the scratchpad folder
//// if it does not exist . . .
//
//  while (scratchpadFolderCollection.hasNext()) {
//    var scratchpadFolder = scratchpadFolderCollection.next();
//  }
//
//  if (scratchpadFolder === undefined) {
//    scratchpadFolder = DriveApp.createFolder('scratchpadFolder');
//  }
//
//// Clean out any scratchpad files found in
//// an existing scratchpad folder . . .
//
//  var scratchpadFileName = "SCRATCHPAD_" + Session.getActiveUser().getEmail();
//  var scratchpadFileCollection = scratchpadFolder.getFilesByName(scratchpadFileName);
//
//  while (scratchpadFileCollection.hasNext()) {
//    var scratchpadFile = scratchpadFileCollection.next();
//    Drive.Files.remove(scratchpadFile.getId());
//  };
//
//// Build a new scratchpad file in
//// the scratchpad folder . . .
//
//  var scratchpadFolderId = scratchpadFolder.getId();
//  var scratchpadFileName = "SCRATCHPAD_" + Session.getActiveUser().getEmail();
//  var scratchpadFile = {
//                         "title": scratchpadFileName,
//                         "mimeType": "application/vnd.google-apps.spreadsheet",
//                         "parents": [
//                                      {
//                                        "id": scratchpadFolderId
//                                      }
//                                    ]
//                       };
//
//  Drive.Files.insert(scratchpadFile);  
//
///////////////////////////////

  var dateVal = new Date();
  var scratchpadFileName = "SCRATCHPAD_" + dateVal.getTime();

  // Define the root folder in the app owner's Drive
  // as the "rootFolder". The app will place the
  // scratchpad sheets in that folder . . .

  var rootFolder = DriveApp.getFolderById(DriveApp.getRootFolder().getId());
  var scratchpadFileCollection = rootFolder.getFilesByName(scratchpadFileName);
  var scratchpadFile;

  while (scratchpadFileCollection.hasNext()) {
    scratchpadFile = scratchpadFileCollection.next();
  };

  scratchpadFile = {
                     "title": scratchpadFileName,
                     "mimeType": "application/vnd.google-apps.spreadsheet",
                     "parents": [
                                  {
                                    "id": rootFolder.getId()
                                  }
                                ]
                   };

  var returnedMetadata = Drive.Files.insert(scratchpadFile);

  // This delay will let the folder / file
  // creation process finish out . . .

  Utilities.sleep(5000);

  var JSONstructure = JSON.parse(returnedMetadata);

  // Extract the new file ID from JSONstructure and return it . . .

  return JSONstructure['id'];
}

function formatScratchpadSpreadsheet(localScratchpadFileId) {

  // Called by Code.gs

  //  Set up the header cells and spreadsheet formatting.
  //  This way, the app can build the spreadsheet itself
  //  with minimal developer / user involvement . . .

  //  This function will place the headerArray array
  //  values in spreadsheet cell range A2:G2. Use \n
  //  as a line break for cell text . . .

  //  If needed, this next line could clean out the
  //  cells
  //
  //    A4:G107
  //
  //  where the application writes but we'll leave
  //  them for now . . .

  //  scratchpadSpreadsheet.getRange('A4:G107').setValue(' ');. . .

  var scratchpadSpreadsheet = SpreadsheetApp.openById(localScratchpadFileId);
  var headerArray = [
                      "First Zip\nDigit",
                      "Employee\nCount",
                      "Q1 Payroll\n(1 = $ 1K)",
                      "Total Annual Payroll\n(1 = $ 1K)",
                      "Total Establishment\nCount",
                      "Function for\nSelect Clause",
                      "Number of\nQuantiles"
                    ];

  //  These lines set cell values, format the sheet, etc.
  //  For cell range A3:G3 the background color #c9daf8
  //  draws a light blue . . .  

  scratchpadSpreadsheet.getRange("A1:G2").setFontFamily("TimesNewRoman");
  scratchpadSpreadsheet.getRange("A1:G4").setHorizontalAlignment("center");
  scratchpadSpreadsheet.getRange("B6").setWrap(true);
  scratchpadSpreadsheet.getRange("C6").setWrap(true);
  scratchpadSpreadsheet.getRange("B6:C106").setHorizontalAlignment("right");
  scratchpadSpreadsheet.getRange("A1").setValue("CENSUS BUREAU\nCOMPLETE ZIP CODE\nTOTALS FILE");
  scratchpadSpreadsheet.getRange("A1:G1").merge();
  scratchpadSpreadsheet.getRange("A1:G2").setFontWeight("bold");
  scratchpadSpreadsheet.getRange("A1:G1").setFontSize(24);
  scratchpadSpreadsheet.getRange("A2:G2").setFontSize(18);
  scratchpadSpreadsheet.getRange("A3:G3").setBackground("#c9daf8");
  scratchpadSpreadsheet.getRange("A3:G3").merge();
  scratchpadSpreadsheet.getRange('A4:G107').setFontFamily("Arial");
  scratchpadSpreadsheet.getRange("A4:G107").setFontSize(10);

  //  The setValues function takes an array as a parameter, but it
  //  wants a multi-dimensional array. To do this, wrap the array
  //  it will get inside an array of its own . . .

  scratchpadSpreadsheet.getRange("A2:G2").setValues([headerArray]);

  // Resize the column width values manually because
  //
  //   scratchpadSpreadsheet.autoResizeColumn(i)
  //
  // resizes the columns with too much extra space.
  // The merged cells
  //
  //   A1:G1
  //
  // in the onOpen function (trigger) might have
  // something to do with this . . .

  var widthArr = [126, 126, 150, 229, 218, 290, 127];

  for (var i = 0; i < widthArr.length; i++) {
    scratchpadSpreadsheet.setColumnWidth((i + 1), widthArr[i]);
  }
}

function removeEscapeCharacters (headerColString) {

  // Called by Code.gs

  //  This function removes the escaped characters
  //  from the column aliases . . .

  headerColString = headerColString.replace(/_/g, ' ');
  headerColString = headerColString.replace(/x24/g, '$');
  headerColString = headerColString.replace(/x28/g, '(');
  headerColString = headerColString.replace(/x29/g, ')');
  headerColString = headerColString.replace(/x3d/g, '=');
  headerColString = headerColString.replace(/x5e/g, '^');

  return headerColString;
}

function returnScratchpadFile() {

  var scratchpadSpreadsheetId = PropertiesService.getScriptProperties().getProperty('scratchpadSpreadsheetId');
  return SpreadsheetApp.openById(scratchpadSpreadsheetId);
}