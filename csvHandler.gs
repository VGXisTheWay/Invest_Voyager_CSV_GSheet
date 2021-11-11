function importVoyagerCSVgmail(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var buttonSet = SpreadsheetApp.getUi().ButtonSet;
  var csvData = {};
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
        csvData = importDummyVoyagerCSV(sheet);
        csvData.splice(0,1); //remove headers row
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
      csvData.splice(0,1); //remove headers row
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
        csvData.splice(0,1); //remove headers row
        buildVoyagerCSVSheet(csvData);
      }
    }
  }
}

function importDummyVoyagerCSV(sheet){
  csvData = dummyVoyagerCSV(100); //imports 100 rows of random dummy CSV data
  Logger.log(csvData)
  dummySheet = createSheet('Voyager CSV')
  dummySheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  resizeAllColumns();
  return csvData
}

function buildVoyagerCSVSheet(data) {
  //var s = SpreadsheetApp.getActiveSpreadsheet();
  //var voyager_CSV_Sheet = s.getSheetByName('Voyager CSV');
  //var drng = voyager_CSV_Sheet.getDataRange();
  //var rng = voyager_CSV_Sheet.getRange(2,1, drng.getLastRow()-1,drng.getLastColumn());
  //var rngA = rng.getValues();//Array of input values
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

  transactions = voyager_csv_sheet_to_dictionary(data);//, true);

  htmlPopUp('<b>Processed ' +
    String(data.length) + "/" + String(data.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Building Gains Sheet...'
  );
  buildGainsSheet(transactions);
  htmlPopUp('<b>Processed ' +
    String(data.length) + "/" + String(data.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Building Current Market Sheet...'
  );
  buildCurrentMarketSheet(transactions);
  htmlPopUp('<b>Processed ' +
    String(data.length) + "/" + String(data.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Building Coin Forecast Sheet...'
  );
  buildCoinForecastSheet(transactions);
  htmlPopUp('<b>Processed ' +
    String(data.length) + "/" + String(data.length) +
    ' Transactions </b><br><br>' +
    displayHeroImg(randomIntFromInterval(0,250)),
    'Building Gains Forecast Table...'
  );
  buildGainsForecastTable();

  htmlPopUp('<b>Processed ' +
    String(data.length) + "/" + String(data.length) +
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

function voyager_csv_sheet_to_dictionary(data="", showHeroPictures=false){
  var transactions ={};
  var dataLength = 0;
  var i = 0;
  var x = 0
  data = build_dict();

  test().then(results => {
     // array of results in order here
     console.log(transactions);
     return transactions
  }).catch(err => {
      console.log(err);
  });
  return transactions

  function test(){
    let promises = [];
    for (let i=0; i < data.length; i++){
      promises.push(process_transactions(data, i));
    }

    return Promise.all(promises);
  }

  function build_dict(){
    //Build dictionary with headers as keys from Voyager CSV sheet
    //This will allows us to work with different formats based on headers vs column position.
    //Get the currently active sheet
    var s = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = s.getSheetByName('Voyager CSV');
    //Get the number of rows and columns which contain some content
    var [rows, columns] = [sheet.getLastRow(), sheet.getLastColumn()];
    //Get the data contained in those rows and columns as a 2 dimensional array
    var data = sheet.getRange(1, 1, rows, columns).getValues();

    // I modified below script.
    var header = data[0];
    data.shift();
    var convertedData = data.map(function(row) {
      return header.reduce(function(o, h, i) {
        o[h] = row[i];
        return o;
      }, {});
    });
    return convertedData
  }

  async function process_transactions(data, i){
    //TODO Recognize which format CSV is in Voyager CSV
      dataLength = data.length;
      var transaction_date = data[i]['transaction_date'];
      var transaction_id = data[i]['transaction_id'];
      var transaction_direction = data[i]['transaction_direction'];
      var transaction_type = data[i]['transaction_type'];
      var base_asset = data[i]['base_asset'];
      var quote_asset = data[i]['quote_asset'];
      var quantity = Number(data[i]['quantity']);
      var net_amount = Number(data[i]['net_amount']);
      var price = Number(data[i]['price']);

      if (showHeroPictures == true){
        if (x % Math.floor(dataLength/5) === 0){ //displays a new hero image every ~1/5 of the iterations
          htmlPopUp('<b>Processed Transaction ' +
                      String(x) + "/" + String(data.length) +
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
      return transactions
  }
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

/** Builds Voyager CSV sheet with DUMMY data
 *
 * @param {number} numberOfRows Number of rows of data to generate
 * @customfunction
 */
function dummyVoyagerCSV(numberOfRows = 10) {
  var date = new Date();
  var startDate = new Date(2018,0,1);
  var endDate = new Date();
  var transactionID = "";
  var transactionDirection = ['Buy', 'Sell', 'deposit'];
  var randomTransactionDirection = "";
  var transactionType = ['TRADE','INTEREST','BANK','REWARD', 'ADMIN'];
  var randomTransactionType = "";
  var baseAssets = ['ADA', 'BTC', 'ETH', 'SOL', 'STMX', 'DOT', 'USD', 'USDC', 'VET', 'VGX'];
  var qty = 0;
  var netAmt = 0;
  var price = 0;
  var dummyHeaderCSV =  'transaction_date,transaction_id,transaction_direction,transaction_type,base_asset,quote_asset,quantity,net_amount,price';
  var dummyDataCSV = "";

  for(let i =0; i < numberOfRows; i++){
    date = new Date(+startDate + Math.random() * (endDate - startDate)).toISOString();
    transactionID = makeid(10);
    let randomBaseAsset = baseAssets[Math.floor(Math.random() * baseAssets.length)];

    if (randomBaseAsset != 'USD'){
      randomTransactionDirection = transactionDirection[Math.floor(Math.random() * transactionDirection.length)];
      randomTransactionType = transactionType[Math.floor(Math.random() * transactionType.length)];
    } else {
      randomTransactionDirection = 'deposit';
      randomTransactionType = 'bank';
    }

    if (randomTransactionDirection == 'Sell' && randomBaseAsset != 'BTC'){ //reduce odds of having more sold than bought
      qty = (Math.random() * 5) + .1;
    } else if (randomTransactionDirection == 'Sell' && randomBaseAsset == 'BTC'){ //reduce odds of having more sold than bought
      qty = (Math.random() * .01) + .001;
    } else if (randomBaseAsset == 'USD') {
      qty = (Math.random() * 1000) + 1;
    } else if (randomBaseAsset == 'BTC') {
      qty = (Math.random() * 1) + .001;
    } else if (randomBaseAsset == 'ETH') {
      qty = (Math.random() * 50) + .05;
    } else {
      qty = (Math.random() * 1000) + 1;
    }

    if (randomBaseAsset == 'ADA'){
      price = (Math.random() * 3) + 0.10;
    } else if (randomBaseAsset == 'BTC'){
      price = (Math.random() * 60000) + 10000;
    } else if (randomBaseAsset == 'ETH'){
      price = (Math.random() * 3850.50) + 1805.25;
    } else if (randomBaseAsset == 'SOL'){
      price = (Math.random() * 205.13) + 20.38;
    } else if (randomBaseAsset == 'STMX'){
      price = (Math.random() * 0.09) + 0.001;
    } else if (randomBaseAsset == 'DOT'){
      price = (Math.random() * 35) + 15;
    } else if (randomBaseAsset == 'USD'){
      price = 1;
    } else if (randomBaseAsset == 'USDC'){
      price = 1;
    } else if (randomBaseAsset == 'VET'){
      price = (Math.random() * .14) + .05608;
    } else if (randomBaseAsset == 'VGX'){
      price = (Math.random() * 4) + 0.01;
    }

    dummyDataCSV = dummyDataCSV + '\r\n' + date.toString() + ',' + transactionID + ',' + randomTransactionDirection + ',' + randomTransactionType + "," + randomBaseAsset + ',USD,' + String(qty) + ',' + String(qty*price) + ',' + String(price);
  }

  dummyDataCSV = dummyHeaderCSV + dummyDataCSV;
  Logger.log(dummyDataCSV)
  dummyData = Utilities.parseCsv(dummyDataCSV, ",");
  Logger.log(dummyData);
  return dummyData;
}

function randomDate(start, end, startHour, endHour) {
  var date = new Date(+start + Math.random() * (end - start));
  var hour = startHour + Math.random() * (endHour - startHour) | 0;
  date.setHours(hour);
  return date;
}

function makeid(length) {
    var result           = '';
    var characters       = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
      result += characters.charAt(Math.floor(Math.random() *
 charactersLength));
   }
   return result;
}
