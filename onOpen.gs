function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('#VGXHeroes')
      .addItem('Import Voyager CSV from GMail', 'importVoyagerCSVgmail')
      .addItem('Import Voyager CSV from Google Drive', 'googleDriveFilePicker')
      .addSeparator()
      .addItem('VGXHeroes.com', 'VGXHeroes.com')
      .addItem('VGXHeroes Discord', 'VGXHeroesDiscord')
      .addItem('VGXHeroes Twiter', 'VGXHeroesTwitter')
      .addToUi();

  buildVoyagerInterestSheet();
  getHeroImageIDs();
  SpreadsheetApp.flush();
}

function getHeroImageIDs() { //save image blobs to global for fast fetching later
  if (PropertiesService.getScriptProperties().getProperty('img0') != undefined){
    var folder = DriveApp.getFolderById("1pwmbq7WmHkVCsxj9m3FkpfXUEn-D5mJv");
    var files = folder.getFiles();
    var fileIDarray = new Array;
    var c=0;

    while(files.hasNext()) {
      var s = files.next();
      fileIDarray[c] = s.getId();
      PropertiesService.getScriptProperties().setProperty('img'+String(c), fileIDarray[c]);
      c=c+1;
    }
  }
}

function loadImageBytes(id){
  var bytes = DriveApp.getFileById(id).getBlob().getBytes();
  //Logger.log(bytes.length)
  var base64 = Utilities.base64Encode(bytes);
  //Logger.log(base64.length)
  return base64
}
