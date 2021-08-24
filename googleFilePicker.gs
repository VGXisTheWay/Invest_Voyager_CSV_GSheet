function googleDriveFilePicker() {
  response = areYouSureClearSheet();
  if (response == true){
    var html = HtmlService.createHtmlOutputFromFile('googleFolderPicker.html')
        .setWidth(600)
        .setHeight(425)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Voyager CSV File');
  }
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
/*
function del_importCSVFromWeb(url) {
  Logger.log("running importCSVFromWeb")
  // Provide the full URL of the CSV file.
  var csvUrl = "https://drive.google.com/file/d/1SD9sNj-pGQRPOUN5jEz6nhKNn89MZOgU";
  var csvContent = UrlFetchApp.fetch(csvUrl).getContentText();
  //var csvData = Utilities.parseCsv(csvContent);
  Logger.log(csvContent)
  buildGainsSheet(csvContent)
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
}

function del_importCSVFromDrive(filename) {
  var sheet = SpreadsheetApp.getActiveSheet();
  //var fileName = promptUserForInput("Please enter the name of the CSV file to import from Google Drive:");
  var files = findFilesInDrive(filename);
  if(files.length === 0) {
    displayToastAlert("No files with name \"" + fileName + "\" were found in Google Drive.");
    return;
  } else if(files.length > 1) {
    displayToastAlert("Multiple files with name " + fileName +" were found. This program does not support picking the right file yet.");
    return;
  }
  var file = files[0];
  //var contents = file.getBlob().getDataAsString();
  //var csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ",");
  csvData = importCSV(file.getBlob(), sheet)
  buildVoyagerCSVSheet(csvData);
}

//Returns files in Google Drive that have a certain name.
function del_findFilesInDrive(filename) {
  var files = DriveApp.getFilesByName(filename);
  var result = [];
  while(files.hasNext())
    result.push(files.next());
  return result;
}*/
