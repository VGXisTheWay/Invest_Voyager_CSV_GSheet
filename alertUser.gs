function areYouSureClearSheet(){
  var buttonSet = SpreadsheetApp.getUi().ButtonSet;
  response = alertUser(
    "⚠️ Are you sure?",
    "This action will erase the active sheet!",
    buttonSet.YES_NO_CANCEL
  );
  return response
}

function alertUser(title = "Alert!", alertmessage, button_layout = SpreadsheetApp.getUi().ButtonSet.OK){
  // Display a dialog box with a message and "Yes" and "No" buttons. The user can also close the
  // dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(title, alertmessage, button_layout);//ui.ButtonSet.button_layout);

  // Process the user's response.
  if (response == ui.Button.YES) {
    Logger.log('The user clicked "Yes."');
    return true
  } else {
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
    return false
  }
}

//Displays an alert as a Toast message
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "⚠️ Alert");
}

function htmlPopUp(message, title){
  var output = HtmlService.createHtmlOutput();
  output.setContent(message);
  output.setWidth(600);
  output.setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(output, title);
}
