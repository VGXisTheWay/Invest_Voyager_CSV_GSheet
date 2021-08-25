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
  buildCoinURLsSheet();
  SpreadsheetApp.flush();
}
