function importVoyagerCSVgmail(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var buttonSet = SpreadsheetApp.getUi().ButtonSet;
  response = areYouSureClearSheet();
  if (response == true){
    try{
      var threads = GmailApp.search('is:starred from:"taxes@investvoyager.com"');
      var message = threads[0].getMessages()[0];
    }
    catch{
      response = alertUser(
        '⚠️ ERROR',
        'Unable to find email from "taxes@investvoyager.com"' +
        '\r\nBe sure the email is starred and in your GMail inbox' +
        '\r\n\r\nWould you like to use dummy data instead?',
        buttonSet.YES_NO_CANCEL
      )
      if (response == true){
        csvData = importDummyVoyagerCSV(sheet)
        buildVoyagerCSVSheet(csvData);
        return;
      }
      else{
        return;
      }
    }
    var attachment = message.getAttachments()[0];
    console.log(threads[0].getFirstMessageSubject());
    console.log(message.getBody());
    console.log(attachment.getContentType());

    response = alertUser(
      "Is this the correct email?",
      "From: " + message.getFrom() +
      "\r\nReceived: " + message.getDate() +
      "\r\nSubject: " + threads[0].getFirstMessageSubject() +
      "\r\nAttachement: " + attachment.getName(),
      buttonSet.YES_NO_CANCEL
    );

    if (response == true){
      csvData = importCSV(attachment, sheet);
      buildVoyagerCSVSheet(csvData);
    }
    else if (response == false){
      response = alertUser(
        'ERROR',
        'Be sure the email from "taxes@investvoyager.com" is starred and in your GMail inbox' +
        '\r\n\r\nWould you like to use dummy data instead?',
        buttonSet.YES_NO_CANCEL
      )
      if (response == true){
        csvData = importDummyVoyagerCSV(sheet)
        buildVoyagerCSVSheet(csvData);
      }
    }
  }
}

function importDummyVoyagerCSV(sheet){
  csvData = dummyVoyagerCSV();
  Logger.log(csvData)
  dummySheet = createSheet('Voyager CSV')
  //sheet.clearContents().clearFormats();
  //dummySheet.getRange("A1").setValue("DUMMY DATA");
  dummySheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  resizeAllColumns();
  //sheet.setName('DUMMY Voyager CSV');
  return csvData
}

function buildVoyagerCSVSheet(data) {
  //var sheet = SpreadsheetApp.getActiveSheet();
  //var transactions = {};
  //var transaction_id = '';
  var base_asset = '';
  var i = 0;

  var s = SpreadsheetApp.getActiveSpreadsheet();
  var voyager_CSV_Sheet = s.getSheetByName('Voyager CSV');
  var drng = voyager_CSV_Sheet.getDataRange();
  var rng = voyager_CSV_Sheet.getRange(2,1, drng.getLastRow()-1,drng.getLastColumn());
  var rngA = rng.getValues();//Array of input values
  var transactions ={};

  /**
   * Voyager default format should have 9 columns:
   * transaction_date, transaction_id, transaction_direction, transaction_type, base_asset, quote_asset, quantity, net_amount, price
   */
  /*
    var transaction_date = data[i][0];
    var transaction_id = data[i][1];
    var transaction_direction = data[i][2];
    var transaction_type = data[i][3];
    var base_asset = data[i][4];
    var quote_asset = data[i][5];
    var quantity = ata[i][6];
    var net_amount = data[i][7];
    var price = data[i][8];
    */

  var output = HtmlService.createHtmlOutput();
  htmlPopUp('<b>Much work. Ready soon... ' +
              '<br><br>' +
              displayHeroImg(randomIntFromInterval(0,250)),
              'Processing...'
  );

  transactions = voyager_csv_sheet_to_dictionary(true);

  htmlPopUp('<b>Processed ' +
    String(rngA.length) + "/" + String(rngA.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Building Gains Sheet...'
  );
  buildGainsSheet(transactions);
  htmlPopUp('<b>Processed ' +
    String(rngA.length) + "/" + String(rngA.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Building Current Market Sheet...'
  );
  buildCurrentMarketSheet(transactions);
  htmlPopUp('<b>Processed ' +
    String(rngA.length) + "/" + String(rngA.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Building Coin Forecast Sheet...'
  );
  buildCoinForecastSheet(transactions);
  htmlPopUp('<b>Processed ' +
    String(rngA.length) + "/" + String(rngA.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Building Gains Forecast Table...'
  );
  buildGainsForecastTable();

  htmlPopUp('<b>Processed ' +
    String(rngA.length) + "/" + String(rngA.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Calculating Coin Totals...'
  );
  var total_bought = 0
  transactions['total_interest'] = 0;
  transactions['total_sold_in_USD'] = 0
  var status = "<b>Coin Totals:</b><br>"

  for([coin, dict] of Object.entries(transactions)){
    var totals = countTotalQty(transactions[coin]);
    transactions[coin]['TotalQty'] = totals['quantity'];
    transactions[coin]['net_amount'] = totals['net_amount'];
    transactions['total_sold_in_USD'] = transactions['total_sold_in_USD'] + totals['total_sold_in_USD'];
    Logger.log(coin + " Total: " + totals['quantity']);
    if (coin != "USD" && coin != "total_interest" && coin != 'total_sold_in_USD'){
      Logger.log(dict)
      var status = status + "<br>" + coin + ": " + totals['quantity']
      total_bought = total_bought + totals['total_bought']
    }
    transactions['total_interest'] = transactions['total_interest'] + totals['interest'];
  }

  Logger.log("Total USD:" + String(transactions["USD"]['TotalQty']) + ", total sold:" + String(transactions['total_sold_in_USD']) + ", bought:" + String(total_bought));
  var usdTotal = transactions["USD"]['TotalQty'] + transactions['total_sold_in_USD'] - total_bought;
  usdTotal = usdTotal.toFixed(2);

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gains').activate(); //Activate Gains Sheet

  status = status + "<br>USD: " + String(usdTotal);
  status = status + "<br><br>TOTAL INTEREST: " + String(transactions['total_interest']);
  output.setContent(status);
  SpreadsheetApp.getUi().showModalDialog(output, 'Work Complete');
}

function importCSV(attachment, sheet){
  var csvData = Utilities.parseCsv(attachment.getDataAsString(), ",");
  var voyager_CSV_Sheet = createSheet('Voyager CSV')
  voyager_CSV_Sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  resizeAllColumns();
  return csvData
}

function voyager_csv_sheet_to_dictionary(showHeroPictures=false){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var sht = s.getSheetByName('Voyager CSV');
  var drng = sht.getDataRange();
  var rng = sht.getRange(2,1, drng.getLastRow()-1,drng.getLastColumn());
  var rngA = rng.getValues();//Array of input values
  var transactions ={};
  var i = 0;
  var total_quantity = 0;

  var x = 0
  for([transaction, dict] of Object.entries(rngA)){ //build dictionary of coins
    var transaction_date = dict[0];
    var transaction_id = dict[1];
    var transaction_direction = dict[2];
    var transaction_type = dict[3];
    var base_asset = dict[4];
    var quote_asset = dict[5];
    var quantity = dict[6];
    var net_amount = dict[7];
    var price = dict[8];

    if (showHeroPictures == true){
      if (x % 75 === 0){ //display progress every X transactions
        htmlPopUp('<b>Processed Transaction ' +
                    String(x) + "/" + String(rngA.length-1) +
                    '<br><br>' +
                    displayHeroImg(randomIntFromInterval(0,250)),
                    'Processing...'
                  );
      }
    }

    if (transactions[base_asset] == undefined){
      transactions[base_asset] = {};
    }
    if (transactions[base_asset][0] == undefined){
      i = 0
      transactions[base_asset][i] = {};
    }
    else{
      i = Object.keys(transactions[base_asset]).length;
    }
    transactions[base_asset][i] = {};
    transactions[base_asset][i] = {};
    transactions[base_asset][i]['transaction_date'] = transaction_date;
    transactions[base_asset][i]['transaction_id'] = transaction_id;
    transactions[base_asset][i]['transaction_direction'] = transaction_direction;
    transactions[base_asset][i]['transaction_type'] = transaction_type;
    transactions[base_asset][i]['base_asset'] = base_asset;
    transactions[base_asset][i]['quote_asset'] = quote_asset;
    transactions[base_asset][i]['quantity'] = quantity;
    transactions[base_asset][i]['net_amount'] = net_amount;
    transactions[base_asset][i]['price'] = price;

    if (transactions[base_asset][i-1] == undefined){ //check if there was a previous transaction for this coin
      transactions[base_asset][i]['total_quantity'] = quantity;
    }
    else {
      if (transaction_direction == "Buy" || transaction_direction == "deposit"){ //add previous qty to this one and record new total
        transactions[base_asset][i]['total_quantity'] = transactions[base_asset][i-1]['total_quantity'] + quantity;
        if(base_asset == "VET"){
          Logger.log(base_asset + " Transaction: " + String(i) + " Buy: " + String(quantity) + " Adding to: " + String(transactions[base_asset][i-1]['total_quantity']))
          Logger.log("Total Qty: " + String(transactions[base_asset][i]['total_quantity']))
        }
      }
      else if (transaction_direction == "Sell"){ //subtract previous qty from this one and record new total
        transactions[base_asset][i]['total_quantity'] = transactions[base_asset][i-1]['total_quantity'] - quantity;
        if(base_asset == "VET"){
          Logger.log(base_asset + " Transaction: " + String(i) + " Sell: " + String(quantity) + " Subtracting from: " + String(transactions[base_asset][i-1]['total_quantity']))
          Logger.log("Total Qty: " + String(transactions[base_asset][i]['total_quantity']))
        }
      }
    }
    x+=1;
  }
  return transactions
}

function importCSVFromWeb(url) {
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

function importCSVFromDrive(filename) {
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
function findFilesInDrive(filename) {
  var files = DriveApp.getFilesByName(filename);
  var result = [];
  while(files.hasNext())
    result.push(files.next());
  return result;
}

function countTotalQty(coin){
  var totals = {};
  totals['quantity'] = 0;
  totals['net_amount'] = 0;
  totals['interest'] = 0;
  totals['total_sold_in_USD'] = 0;
  totals['total_bought'] = 0;

  for([transaction, value] of Object.entries(coin)){
    Logger.log(transaction);
    Logger.log(value);
    var coin = value['base_asset'];
    var buy_sell_deposit = value['transaction_direction'];
    var transaction_type = value['transaction_type'];
    Logger.log(buy_sell_deposit);

    if (buy_sell_deposit == "Buy"){
      totals['quantity'] = totals['quantity'] + value['quantity'];
      totals['total_bought'] = totals['total_bought'] + value['net_amount'].round(2);
    }
    else if (buy_sell_deposit == "deposit"){
      totals['quantity'] = totals['quantity'] + value['quantity'];

      if (transaction_type != "INTEREST" && transaction_type != "ADMIN"){
        totals['net_amount'] = totals['net_amount'] + value['net_amount'];
      }
      else if (transaction_type == "INTEREST"){
        totals['interest'] = totals['interest'] + value['net_amount'];
      }
    }
    else if (buy_sell_deposit == "Sell"){
      totals['quantity'] = totals['quantity'] - value['quantity'];
      totals['net_amount'] = totals['net_amount'] - value['net_amount'];
      totals['total_sold_in_USD'] = totals['total_sold_in_USD'] + value['net_amount']
    }
  }
  return totals
}

Number.prototype.round = function(places) {
  return +(Math.round(this + "e+" + places)  + "e-" + places);
}

function rowToDict(sheet, rownumber) {
  var columns = sheet.getRange(1,1,1, sheet.getMaxColumns()).getValues()[0];
  var data = sheet.getDataRange().getValues()[rownumber];
  var dict_data = {};
  for (var keys in columns) {
    var key = columns[keys];
    dict_data[key] = data[keys];
  }
  Logger.log(dict_data)
  return dict_data;
}

function resizeAllColumns () {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var firstColumn = dataRange.getColumn();
  var lastColumn = dataRange.getLastColumn();
  sheet.autoResizeColumns(firstColumn, lastColumn);
}

function dummyVoyagerCSV() {
  dummyCSV =  'transaction_date,transaction_id,transaction_direction,transaction_type,base_asset,quote_asset,quantity,net_amount,price' +
              '\r\n2019-01-01 01:00:00.000000+00:00,USD123456,Buy,TRADE,USD,USD,63025,1.00,1.00' +
              '\r\n2020-01-01 01:00:00.000000+00:00,BTC1234567,Buy,TRADE,BTC,USD,0.01,450.00,45000.00' +
              '\r\n2020-02-01 01:00:00.000000+00:00,BTC7891011,Buy,TRADE,BTC,USD,0.5,15000.00,30000.00' +
              '\r\n2020-01-02 01:00:00.000000+00:00,ADA1234567,Buy,TRADE,ADA,USD,1000,1500.00,1.50' +
              '\r\n2020-02-02 01:00:00.000000+00:00,ADA2468012,Buy,TRADE,ADA,USD,100,125.00,1.25' +
              '\r\n2020-01-03 01:00:00.000000+00:00,STEVE42691,Buy,TRADE,VGX,USD,1000,2000.00,2.00' +
              '\r\n2020-02-04 01:00:00.000000+00:00,STEVE2MOON,Buy,TRADE,VGX,USD,15000,33750.00,2.25' +
              '\r\n2020-02-04 01:10:00.000000+00:00,STMX123456,Buy,TRADE,STMX,USD,400000,25000.00,0.025' +
              '\r\n2020-03-01 01:10:00.000000+00:00,VGXINT1234,deposit,INTEREST,VGX,N/A,25,87.50,3.50' +
              '\r\n2020-03-02 01:15:00.000000+00:00,DOT1234545,Buy,TRADE,DOT,USD,10,200.00,20.00';


  dummyData = Utilities.parseCsv(dummyCSV, ",");
  Logger.log(dummyData);
  return dummyData;
}
