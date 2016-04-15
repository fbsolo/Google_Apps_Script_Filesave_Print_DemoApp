//  Article
//
//    http://bitvectors.blogspot.com/2016/04/page1.html
//
//  references and describes this application . . .

function doGet(){

  // This called function handles scratch pad file creation . . .

  var scratchpadSpreadsheetId = returnScratchpadFileCollection();

  PropertiesService.getScriptProperties().setProperty('scratchpadSpreadsheetId', scratchpadSpreadsheetId);

  // Through the createHtmlOutputFromFile() function,
  // multiple users can use the application . . .

  return HtmlService.createHtmlOutputFromFile("BigQueryDemoApp.html").setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function returnFormParams(queryString, arrayParam){

  var scratchpadSpreadsheet = returnScratchpadFile();
  var returnParam = ' ';
  
  scratchpadSpreadsheet.getRange("A4:G4").setValues([arrayParam.slice(0, 7)]);
  returnParam = runQuery(queryString);

  return returnParam;
}

function runQuery(queryString){

  var scratchpadSpreadsheet = returnScratchpadFile();

  //  If the user already ran the app with the quantiles function
  //  and the quantiles function returned an empty result set,
  //  runQuery will gray out cell B7 to help illustrate the
  //  situation. Therefore, initialize the cell B7 background
  //  color . . .

  scratchpadSpreadsheet.getRange('B7').setBackground("white");

  // Replace this value with the project ID placed
  // in the Google Developers Console project . . .

  var projectId = PropertiesService.getScriptProperties().getProperty('projectId');  
  var request = {query: queryString};
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;

  // Get the entire result set . . .

  var rows = queryResults.rows;

  while (queryResults.pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken: queryResults.pageToken
    });
    rows = rows.concat(queryResults.rows);
  }

  //  The headers array has the result
  //  set column names . . .

  var headers = queryResults.schema.fields.map(function(field) {
    return field.name;
  });
  
  //  The result set data will go into a
  //  nested array that will work as an
  //  array of arrays. The inner array(s)
  //  will have two values . . .

  var dataArray = [[]];

  //  In dataArray, dataArray[0][0] has the function
  //  name and dataArray[0][1] has the calculated
  //  value BigQuery returned . . .

  dataArray[0][0] = removeEscapeCharacters(headers[0]);

  if (headers.length == 1) {

    //  The headers[] array has one element, so the user picked a single-value
    //  result set function. If BigQuery returned NULL for the chosen parameters,
    //  place an information string in dataArray[0][1]; otherwise, place the
    //  returned non-null value in dataArray[0][1] . . .

    dataArray[0][1] = (rows[0].f[0].v === null) ? "No value calculated for chosen parameters" : rows[0].f[0].v;
    scratchpadSpreadsheet.getRange('G4').setBackground("lightgray");

  } else if (headers.length == 2) {

    //  The user picked a two-column result
    //  set - specifically, the quantiles
    //  function . . .

    dataArray[0][1] = removeEscapeCharacters(headers[1]);

    if (rows.length < 2) {

      //  The BigQuery quantiles function returned
      //  zero data rows for the parameters, so
      //  first, gray out cell B7 as a visual guide,
      //  and build a two-cell array that explains
      //  everything . . .

      scratchpadSpreadsheet.getRange('B7').setBackground("lightgray");

      var tempArray = new Array(2);

      tempArray[0] = " ";
      tempArray[1] = "No quantile values calculated for chosen parameters";

      //  The slice() method guarantees that tempArray[] will
      //  have the new values from the sourcing rows[] array.
      //  Without slice(), the push method will push arrays
      //  referenced by the last tempArray it pushed in this
      //  loop . . .

      dataArray.push(tempArray.slice());
    } else {

      //  Array tempArray will hold the quantile / quantile value pairs . . .

      var tempArray = new Array(2);

      for (var i = 0; i < rows.length; i++) {
        tempArray[0] = rows[i].f[0].v;
        tempArray[1] = rows[i].f[1].v;

        //  The slice() method guarantees that tempArray[] will
        //  have the new values from the sourcing rows[] array.
        //  Without slice(), the push method will push arrays
        //  referenced by the last tempArray it pushed in this
        //  loop . . .

        dataArray.push(tempArray.slice());
      }
    }
  }

  //  The dataArray array now has all the result set data. Nested loops could certainly
  //  place the dataArray array values in the spreadsheet cells, but it would take forever.
  //  Instead, place the entire assembled dataArray array in the spreadsheet, at the target
  //  location all at once. This will boost the speed.
  //
  //  The getRange function specifies the target location which starts at B6, extends to
  //  column C, and down to the row matching the length of dataArray[0]. Here, dataArray[0].length
  //  is the column length - AKA the number of quantile / quantile value pairs. Add 5 (five)
  //  because targetRange has five blank rows above it . . .

  var targetRange = "B6:C" + (dataArray.length + 5); 

  scratchpadSpreadsheet.getRange(targetRange).setValues(dataArray);

  return dataArray;
}

function funcSaveSheet(firstBlankRow) {

  // Place the header text / etc. in the scratchpad spreadsheet.
  // Do this in this function instead of returnFormParams() because
  // at this point, the app workflow will save the scratchpad spreadsheet.
  // If the call happens in returnFormParams(), the app workflow could
  // delete the file before the user saves it, wasting the call . . .

  var scratchpadSpreadsheet = returnScratchpadFile();

  formatScratchpadSpreadsheet(PropertiesService.getScriptProperties().getProperty('scratchpadSpreadsheetId'));

  // Enable the Drive API
  //
  // 1) in the script editor
  //
  //    Resources -> Advanced Google Services
  //
  // 2) Google Developers Console
  //
  //    https://console.developers.google.com

  var i, j;
  var response;

  // This app will only need the single / first component
  // sheet of the spreadsheet file. Therefore, hardwire
  // the sheet variable as that component sheet . . .
  
  var sheet = scratchpadSpreadsheet.getSheets()[0];
  var url_ext;

  // Get the scratchpad spreadsheet URL, removing the trailing 'edit' . . .

  var url = scratchpadSpreadsheet.getUrl().replace(/edit$/,'');

  var options = {
    headers: {
      'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
    }
  }

  // A Google Sheet defaults to 1000 rows. Starting from
  // parameter firstBlankRow, hide the rest of the rows
  // down to the end / bottom of the sheet. This way,
  // they won't show up in the PDF as a possibly blank
  // sheet . . .
  
  sheet.hideRows(firstBlankRow, (1000 - firstBlankRow));
  
  // These parameters configure the sheet for export as a PDF . . .
  
  url_ext = 'export?exportFormat=pdf&format=pdf'   // PDF format
  + '&gid=' + sheet.getSheetId()   // sheet ID
  // optional parameters
  + '&size=letter'      // set paper size
  + '&portrait=true'    // set the orientation
  + '&fitw=true'        // set fit to width, false for actual size
  + '&sheetnames=false&printtitle=false&pagenumbers=false'  // headers and footers - all off
  + '&gridlines=false'  // gridlines - off
  + '&fzr=false';       // no repeated row headers (frozen rows) for all pages

  // Place the content of each sheet into a variable . . .

  response = UrlFetchApp.fetch(url + url_ext, options);

  // The return uses these functions to correctly format
  // the file content as a finished PDF file . . .

  return Utilities.base64Encode(response.getBlob().getBytes());
}