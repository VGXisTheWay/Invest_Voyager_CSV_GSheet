function displayImgModal(imageCount){
  var output = HtmlService.createHtmlOutput();
  var html= ""
  var i = 0
  while (i < imageCount){
    blob = imgBlobs[i];
    html="<p style='text-align:center;'> <img src='data:image/jpg;base64,"+ blob + "' /></p>";
    output.setContent(html);
    output.setWidth(600);
    output.setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(output, 'title');
    i += 1;
    Utilities.sleep(1500)
  }
  return html;
}

function displayHeroImg(imgNum){
  var output = HtmlService.createHtmlOutput();
  var html= ""
  var imgID = PropertiesService.getScriptProperties().getProperty('img'+String(imgNum));

  var blob = loadImageBytes(imgID);
  html="<p style='display:flex; justify-content:center;'> <img src='data:image/jpg;base64,"+ blob + "' style='max-width:95vw; max-height:85vh; width:auto;height:auto;'/></p>";
  //output.setContent(html);
  //output.setWidth(600);
  //output.setHeight(600);

  return html; //output;
}

function getHeroImageIDs() { //save image blobs to global for fast fetching later
  if (PropertiesService.getScriptProperties().getProperty('img0') != undefined){
    var folder = DriveApp.getFolderById("1pwmbq7WmHkVCsxj9m3FkpfXUEn-D5mJv");
    var files = folder.getFiles();
    var fileIDarray = new Array;
    var c=0;
    var imgIDs = [];

    while ( files.hasNext() ){
      imgIDs.push(files.next().getId());
    }

    while (c<250){
      PropertiesService.getScriptProperties().setProperty('img'+String(c), imgIDs[Math.floor(Math.random() * (250 - 1 + 1)) + 1]);
      c += 1;
    }

    Logger.log(c);
  }
}

function loadImageBytes(id){
  var bytes = DriveApp.getFileById(id).getBlob().getBytes();
  //Logger.log(bytes.length)
  var base64 = Utilities.base64Encode(bytes);
  //Logger.log(base64.length)
  return base64
}

function randomIntFromInterval(min, max) { // min and max included
  return Math.floor(Math.random() * (max - min + 1) + min)
}
