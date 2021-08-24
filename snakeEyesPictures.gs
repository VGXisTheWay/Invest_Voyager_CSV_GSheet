function displayImgModal(imageCount){
  var output = HtmlService.createHtmlOutput();
  var html= ""
  var i = 0
  while (i < imageCount){ //for ([img, blob] of Object.entries(imgBlobs)){
    blob = imgBlobs[i]; //PropertiesService.getScriptProperties().getProperty('img'+String(i));
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

  var blob = loadImageBytes(imgID)//imgBlobs[imageNum]; //PropertiesService.getScriptProperties().getProperty('img'+String(i));
  html="<p style='display:flex; justify-content:center;'> <img src='data:image/jpg;base64,"+ blob + "' style='max-width:95vw; max-height:85vh; width:auto;height:auto;'/></p>";
  //output.setContent(html);
  //output.setWidth(600);
  //output.setHeight(600);

  return html; //output;
}
