function buildGainsSheet(transactions){
  //transactions = voyager_csv_sheet_to_dictionary(false); //use for debugging only
  var gainsSheet = createSheet('Gains');
  var activeRange = SpreadsheetApp.getActiveRange();
  var gainsHeaders = ["Coin", "Current Interest", "Quantity", '=CONCATENATE(TEXT(NOW(),"MMM")," Ave Daily Qty")', "Expected Interest Income", "Ave Cost Per Share", "Total Cost", "7-Day Price Graph", "Current Price", "Current Value", "$ Gain", "% Gain", "Total Interest Earned", "Refresh Data"];
  var headerRange = gainsSheet.getRange(1,1,1,gainsHeaders.length);
  var i = 0;
  var transactionsVGX_length = 0;
  for (header in gainsHeaders){
    gainsSheet.getRange(1,i+1).setValue(gainsHeaders[i]);
    i += 1;
  }

  headerRange.setBackgroundRGB(119,136,153);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setWrap(true);

  voyagerInterest = new Array;

  var row = 2;
  gainsSheet.getRange(row,14).insertCheckboxes(false); //checkbox acts as toggle which refreshes current market importHTML() formulas
  for([coin, value] of Object.entries(transactions)){
    if (coin != "USD"){
      gainsSheet.getRange(row,1).setValue(coin);
      //Interest
      if(["BTC", "ETH", "USDC"].includes(coin)){
        gainsSheet.getRange(row,2).setFormula("=IFS(INDEX($A:$C,MATCH(\"VGX\",$A:$A,0),3)>20000,SUBSTITUTE(INDEX('Voyager Interest'!$B:$C,MATCH(\"*" + coin + "*\",'Voyager Interest'!$C:$C,0),1),\"*\",\"\")+1.5%, INDEX($A:$C,MATCH(\"VGX\",$A:$A,0),3)>5000,SUBSTITUTE(INDEX('Voyager Interest'!$B:$C,MATCH(\"*" + coin + "*\",'Voyager Interest'!$C:$C,0),1),\"*\",\"\")+1.0%, INDEX($A:$C,MATCH(\"VGX\",$A:$A,0),3)>500,SUBSTITUTE(INDEX('Voyager Interest'!$B:$C,MATCH(\"*" + coin + "*\",'Voyager Interest'!$C:$C,0),1),\"*\",\"\")+0.5%)");
        transactionsVGX_length = Object.keys(transactions['VGX']).length-1;
        if(5000 > transactions['VGX'][transactionsVGX_length]['total_quantity'] > 500){
          gainsSheet.getRange("B"+String(row)).setNote('Congrats Adventurer! 0.5% BOOST').setFontColor('purple');
        }
        else if(20000 > transactions['VGX'][transactionsVGX_length]['total_quantity'] > 5000){
          gainsSheet.getRange("B"+String(row)).setNote('Congrats Explorer! 1% BOOST').setFontColor('purple');
        }
        else if(transactions['VGX'][transactionsVGX_length]['total_quantity']>20000){
          gainsSheet.getRange("B"+String(row)).setNote('Congrats Navigator! 1.5% BOOST').setFontColor('purple');
        }
      }
      else{
        gainsSheet.getRange(row,2).setFormula("=SUBSTITUTE(INDEX('Voyager Interest'!$B:$C,MATCH(" + '"*' + coin + '*"' + ",'Voyager Interest'!$C:$C,0),1)," + '"*"' + ',"")');
      }
      gainsSheet.getRange(row,2).setNumberFormat("#0.00%");
      //Quantity
      gainsSheet.getRange(row,3).setFormula("=SUM(SUMIFS('Voyager CSV'!$G:$G,'Voyager CSV'!$E:$E," + '"*' + coin + '*"' + ",'Voyager CSV'!$C:$C," +'"Buy"' + ")+SUMIFS('Voyager CSV'!$G:$G,'Voyager CSV'!$E:$E," + '"*' + coin + '*"' + ",'Voyager CSV'!$C:$C," + '"deposit"' + ")-SUMIFS('Voyager CSV'!$G:$G,'Voyager CSV'!$E:$E," + '"*' + coin + '*"' + ",'Voyager CSV'!$C:$C," + '"Sell"))');
      //Average Daily
      gainsSheet.getRange(row,4).setFormula('=getAveBal($A'+row+')');
      //Expected Interest
      gainsSheet.getRange(row,5).setFormula('=IF(C'+row+'<>"#N/A",$D'+row+'*$B'+row+'/12)');
      //Ave Cost Per Coin
      gainsSheet.getRange(row,6).setFormula("=SUMIFS('Voyager CSV'!$H:$H,'Voyager CSV'!$E:$E," + '"' + coin + '"' + ",'Voyager CSV'!$D:$D," + '"<>REWARD"' + ")/SUMIFS('Voyager CSV'!$G:$G,'Voyager CSV'!$E:$E," + '"' + coin + '"' + ", 'Voyager CSV'!$D:$D," + '"<>REWARD")');
      //Total Cost
      gainsSheet.getRange(row,7).setFormula('$C'+row+'*$F'+row);
      //7-Day Price Graph
      gainsSheet.getRange(row,8).setFormula('=SPARKLINE(CRYPTOFINANCE("' + coin + '", "sparkline"))');
      //Current Price
      gainsSheet.getRange(row,9).setFormula("=INDEX('Current Market'!$A:$C,MATCH(" + '"' + coin + '"' + ",'Current Market'!$A:$A,0),3)").setNumberFormat("$#,##0.00;$(#,##0.00)");
      //Current Value
      gainsSheet.getRange(row,10).setFormula('$C'+row+'*$I'+row);
      //$ Gain
      gainsSheet.getRange(row,11).setFormula('($C'+row+'*$I'+row+')-($C'+row+'*$F'+row+')');
      //Percent Gain
      gainsSheet.getRange(row,12).setFormula('($C'+row+'*$I'+row+')/($C'+row+'*$F'+row+')-1');
      gainsSheet.getRange(row,12).setNumberFormat("##%");
      //Total Interest Earned
      gainsSheet.getRange(row,13).setFormula("=SUMIFS('Voyager CSV'!$H:$H,'Voyager CSV'!$E:$E," + '"' + coin + '"'  + ",'Voyager CSV'!$D:$D," + '"=INTEREST")');
      var lastGainsTableRow = row
      row += 1;
    }
  }
  //Set cells that show column totals
  //Sum Expected Interest
  i=2; //start at row 2 because row 1 is headers
  expectedInterest = "=0"
  while (i<row){
    if (gainsSheet.getRange(i,5).getValue() != "#N/A"){ //skips coins that don't earn interest
      expectedInterest = expectedInterest + "+(E" + String(i) + "* I" + String(i) + ")"; //add (multiply Ave Daily by Current Price)
    }
    i+=1;
  }
  setAlternatingColors("$A2:$M" + String(row-1));
  gainsSheet.getRange(row,5).setFormula(expectedInterest).setFontWeight('bold').setFontSize(12).setNumberFormat("$#,##0.00;$(#,##0.00)");
  //Sum Total Cost Column
  gainsSheet.getRange(row,7).setFormula("=SUMIF(G2:G" +String(row-1)+",\"<>#DIV/0!\",G2:G"+String(row-1)+")").setFontWeight('bold').setFontSize(12);
  gainsSheet.getRange("G:G").setNumberFormat("$#,##0.00;$(#,##0.00)");
  //Sum Total Current Value Column
  gainsSheet.getRange(row,10).setFormula("=SUMIF(J2:J" +String(row-1)+",\"<>#DIV/0!\",J2:J"+String(row-1)+")").setFontWeight('bold').setFontSize(12);
  gainsSheet.getRange("J:J").setNumberFormat("$#,##0.00;$(#,##0.00)");
  //Sum Total $ Gain Column
  gainsSheet.getRange(row,11).setFormula("=SUMIF(K2:K" +String(row-1)+",\"<>#DIV/0!\",K2:K"+String(row-1)+")").setFontWeight('bold').setFontSize(12);
  gainsSheet.getRange("K:K").setNumberFormat("$#,##0.00;$(#,##0.00)");
  //Sum Total Interest Earned
  gainsSheet.getRange(row,13).setFormula("=SUMIF(M2:M" +String(row-1)+",\"<>#DIV/0!\",M2:M"+String(row-1)+")").setFontWeight('bold').setFontSize(12);
  gainsSheet.getRange("M:M").setNumberFormat("$#,##0.00;$(#,##0.00)");
  //Set Interest column to center
  gainsSheet.getRange("B:B").setHorizontalAlignment('center');
  row+=3

  //TODO Conditional Formatting Rules

  //Forcast Table
  gainsSheet.getRange(row,1).setValue("Forecast Table").setFontWeight('bold'); //Forecast Table Title
  row+=1;
  var forecastHeaders = ["Coin", "Date", "Forecasted Price", 'Current Qty', "Total Cost", "Forecasted Value", "$ Gain", "% Gain", "Total Value"];
  var forecastHeaderRange = gainsSheet.getRange(row,1,1,forecastHeaders.length);
  i = 0;
  for (header in forecastHeaders){
    gainsSheet.getRange(row,i+1).setValue(forecastHeaders[i]);
    i += 1;
  }
  with (forecastHeaderRange){
    setBackgroundRGB(119,136,153);
    setFontWeight('bold');
    setHorizontalAlignment('center');
    setWrap(true);
  }
  row+=1;

  var forecastCoins = getCoinsWithForecast();
  var forecastDates = {Dec_2021:"+1,4), \"*\",\"\")", Dec_2022:"+2,4), \"*\",\"\")", Dec_2023:"+3,4), \"*\",\"\")", Dec_2024:"+4,4), \"*\",\"\")", Dec_2025:"+5,4), \"*\",\"\")"};
  var yearStartRow = "";

  for([date, value] of Object.entries(forecastDates)){
    gainsSheet.getRange(row,1).setValue(date).setFontWeight('bold');
    row+=1;
    yearStartRow = String(row-1); //record first row of this year forecast
    for (i in forecastCoins){
      buildForcastedGainsTable(gainsSheet, date, forecastCoins[i], value, String(row), String(lastGainsTableRow));
      row+=1;
    }
    gainsSheet.getRange(row-1,9).setFormula("=SUM($F"+yearStartRow+":$F"+String(row-1)+")").setNumberFormat("$#,##0.00;$(#,##0.00)").setFontWeight('bold');
    row+=1;
    setAlternatingColors("$A" + yearStartRow + ":$I" + String(row-2));
  }
  resizeAllColumns();
}

function setAlternatingColors(range) {
  Logger.log(range.toString());
  var bandingTheme = [SpreadsheetApp.BandingTheme.BLUE, SpreadsheetApp.BandingTheme.BROWN, SpreadsheetApp.BandingTheme.CYAN,SpreadsheetApp.BandingTheme.GREEN,SpreadsheetApp.BandingTheme.GREY,SpreadsheetApp.BandingTheme.INDIGO,SpreadsheetApp.BandingTheme.LIGHT_GREEN,SpreadsheetApp.BandingTheme.LIGHT_GREY,SpreadsheetApp.BandingTheme.ORANGE,SpreadsheetApp.BandingTheme.PINK,SpreadsheetApp.BandingTheme.TEAL,SpreadsheetApp.BandingTheme.YELLOW]
  let randomBandingTheme = bandingTheme[Math.floor(Math.random() * bandingTheme.length)];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getRange(range);
  // first remove any existing alternating colors in range to prevent error "Exception: You cannot add alternating background colors to a range that already has alternating background colors."
  range.getBandings().forEach(banding => banding.remove());
  // apply alternate background colors
  range.applyRowBanding(randomBandingTheme, false, false);
}

function getCoinsWithForecast(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var values = sheet.getSheetByName('Coin Forecast').getDataRange().getValues();
  var cellValues = [];
  var i=0;
  for(n=0;n<values.length;++n){
    cellValues[i] = values[n][0];
    i+=1;
  }
  cellValues = cellValues.filter(Boolean);
  return cellValues;
}

function buildForcastedGainsTable(sheet, date, coin, dateFormula, row, lastGainsTableRow){
  sheet.getRange(row,1).setValue(coin); //coin name
  sheet.getRange(row,2).setValue(date); //forecast date
  sheet.getRange(row,3).setFormula("=SUBSTITUTE(INDEX('Coin Forecast'!$A:$E,MATCH(\""+coin+"\",'Coin Forecast'!$A:$A,0)" + dateFormula); //Formula gets forcasted price for date
  sheet.getRange(row,4).setFormula("=VLOOKUP(\""+coin+"\",$A$2:$M$"+lastGainsTableRow+", 3, False)"); //Qty
  sheet.getRange(row,5).setFormula("=VLOOKUP(\""+coin+"\",$A$2:$M$"+lastGainsTableRow+", 7, False)").setNumberFormat("$#,##0.00;$(#,##0.00)"); //Total Cost
  sheet.getRange(row,6).setFormula("C"+row+"*D"+row).setNumberFormat("$#,##0.00;$(#,##0.00)"); //Forecasted Value (Forcasted Price * Qty)
  sheet.getRange(row,7).setFormula("F"+row+"-E"+row).setNumberFormat("$#,##0.00;$(#,##0.00)"); //$ Gains (Forecasted Value - Total Cost)
  sheet.getRange(row,8).setFormula("F"+row+"/E"+row+"*1").setNumberFormat("#0.00%"); //% Gains (Forecasted Value / Total Cost * 1)
}
// Calculate average balance for this month
function getAveBal(sym){
  //var sym = "BTT";
  var currentMonth = new Date().getMonth();
  var date = new Date();
  var lastDayOfThisMonth = new Date(date.getFullYear(), currentMonth+1, 0).getDate();
  var today = date.getDate();
  var dayTotals = [];
  var aveDailyBal;
  var i = 0;
  var x = 0;
  var thisTransDate;
  var transactions = voyager_csv_sheet_to_dictionary(false);
  var transaction;
  var transaction_value;
  var coin;
  var dict = transactions[sym];
  var total_quantity = 0;
  /*
  function orderBySubKey( input, key ) {
    return Object.values( input ).map( value => value ).sort( (a, b) => a[key] - b[key] );
  }
  var newdict = orderBySubKey(dict,'transaction_date')*/
  for ([transaction, transaction_value] of Object.entries(dict)){
    if (transaction_value != 'total_quantity'){
      transaction = parseInt(transaction);
      thisTransDate = new Date(transaction_value['transaction_date']);
      total_quantity = total_quantity + transaction_value['quantity'];
      if (thisTransDate.getMonth() == currentMonth){
        x+=1;
        if (thisTransDate.getDate() == 1){ //Transaction ocurred on the 1st so yesterday's total is today's.
          if (dict[transaction - 1] != undefined){ //make sure this isn't the first time this coin has been seen
            console.log("Transaction ocurred on 1st. Setting to yesterdays total: " + String(dict[transaction - 1]['total_quantity']));
            dayTotals[thisTransDate.getDate()] = dict[transaction - 1]['total_quantity'];
            console.log("Day: " + String(thisTransDate.getDate()+1) + " Total: " + String(dict[transaction]['total_quantity']));
            dayTotals[thisTransDate.getDate()+1] = dict[transaction]['total_quantity']; //set the next daysTotals to this transaction total quantity
            continue;
          }
          else {  //first ever transaction for this coin
            console.log("Day: " + String(thisTransDate.getDate()) + " Total: 0"); //this transaction doesn't take affect until the next day
            dayTotals[thisTransDate.getDate()] = 0;
            //console.log("Day: " + String(thisTransDate.getDate()+1) + " Total: " + String(dict[transaction]['total_quantity']));
            //dayTotals[thisTransDate.getDate()+1] = dict[transaction]['total_quantity'];
            i = thisTransDate.getDate()+1;
            continue;
          }
        }
        else {
          if (x>1){
            console.log("Day: " + String(thisTransDate.getDate()) + " Total: " + String(dict[transaction - 1]['total_quantity'])); //this transaction doesn't take affect until the next day
            dayTotals[thisTransDate.getDate()] = dict[transaction - 1]['total_quantity'];
            console.log("Day: " + String(thisTransDate.getDate()+1) + " Total: " + String(dict[transaction]['total_quantity']));
            dayTotals[thisTransDate.getDate()+1] = dict[transaction]['total_quantity'];
            i = thisTransDate.getDate()+1
          }
          else {
            i = 1;
            while (i <= thisTransDate.getDate()){ //today){
              console.log("Day: " + String(i) + " Total: " + String(dict[transaction - 1]['total_quantity']));
              dayTotals[i] = dict[transaction - 1]['total_quantity']; //total_quantity;
              i++;
            }
          }
          //Logger.log(Object.keys(dict).length)
          if(transaction != Object.keys(dict).length - 1){
            continue
          }
        }
        //i = thisTransDate.getDate()
        /*
        while (i <= thisTransDate.getDate()){ //today){
          dayTotals[i] = dayTotals[i-1]//total_quantity;
          i++
        }
        */
      }
    }
  }
  i=1
  while (i <= lastDayOfThisMonth){ // thisTransDate.getDate()){ //today){
    if (dayTotals[i] == undefined){
      if (dayTotals[i-1] != undefined){
        console.log("Day: " + String(i) + " Total: " + String(dayTotals[i-1]));
        dayTotals[i] = dayTotals[i-1]//total_quantity;
      }
    }
    i++
  }

  if (dayTotals.length == 0){ //No transactions found for the month. Get latest total quantity and carry it forward
    console.log("No trades this month for " + sym + ". Carrying forward total: " + String(dict[transaction]['total_quantity']));
    aveDailyBal = dict[transaction]['total_quantity'];
  }
  else if (dayTotals.length == 1){
    console.log("Monthly ave balance for " + sym + " calculated as: " + String(dayTotals[1]))
    aveDailyBal = dayTotals[1]
  }
  else{
    console.log("Monthly ave balance for " + sym + " calculated as: " + String(dayTotals.reduce((a, b) => a + b, 0) / (dayTotals.length)))
    aveDailyBal = dayTotals.reduce((a, b) => a + b, 0) / (dayTotals.length)
  }

  return aveDailyBal;
}

/** Creates a string of random letters at a set length.
 *
 * @param {number} len The total number of random letters in the string.
 * @param {number} num What type of random number 0. Alphabet with Upper and Lower. 1.Alphanumeric 2. Alphanumeric + characters
 * @return an array of random letters
 * @customfunction
 */
function RANDALPHA(len, num) {
  var text = "";

  //Check if numbers
  if(typeof len !== 'number' ||  typeof num !== 'number'){return text = "NaN"};

  var charString = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789/+";
  var charStringRange
  switch (num){
     case 0:
       //Alphabet with upper and lower case
       charStringRange = charString.substr(0,52);
       break;
     case 1:
       //Alphanumeric
       charStringRange = charString.substr(0,62);
       break;
     case 2:
       //Alphanumeric + characters
       charStringRange = charString;
       break;
     default:
       //error reporting
       return text = "Error: Type choice > 2"

  }
  //
  for (var i = 0; i < len; i++)
    text += charStringRange.charAt(Math.floor(Math.random() * charStringRange.length));

  return text;
}

/**
* Imports JSON data to your spreadsheet
* @param url URL of your JSON data as string
* @param xpath simplified xpath as string
* @customfunction
*/
function importJSON(url,xpath){
  count = 0;
  maxTries = 15;
  while(true) {
    try{
      // /rates/EUR
      var res = UrlFetchApp.fetch(url);
      var content = res.getContentText();
      var json = JSON.parse(content);

      var patharray = xpath.split(".");
      //Logger.log(patharray);

      for(var i=0;i<patharray.length;i++){
      json = json[patharray[i]];
    }

    //Logger.log(typeof(json));

    if(typeof(json) === "undefined"){
      return "Node Not Available";
    }
    else if(typeof(json) === "object"){
      var tempArr = [];

      for(var obj in json){
        tempArr.push([obj,json[obj]]);
      }
      return tempArr;
    }
      else if(typeof(json) !== "object") {
        return json;
      }
    }
    catch(err){
      //return "Error getting data";
      Utilities.sleep(100);
      if (++count == maxTries) throw err;
    }
  }
}

function buildCurrentMarket(coins){
  var currentMarket = createSheet('Current Market')
  var row = 1;

  for (coin in coins){
    if (coin != 'total_interest' && coin != 'USD'){
      Logger.log(coin);
      url = rowOfCoin(coin, 'current market');
      currentMarket.getRange('B' + String(row)).setValue('=IMPORTHTML("'+url+'","table", 1, \'Gains\'!$N$2)');
      currentMarket.getRange('A' + String(row)).setValue(coin);
      row = row+8;
    }
  }
  resizeAllColumns();
}

function buildCoinForecast(coins){
  coinForecast = createSheet('Coin Forecast')
  var row = 1;

  for (coin in coins){
    if (coin != 'total_interest' && coin != 'USD'){
      Logger.log(coin);
      var url = rowOfCoin(coin, 'forecast');
      if(url != "N/A"){
        //coinForecast.getRange('B' + String(row)).setValue(url);
        //coinForecast.getRange('A' + String(row)).setValue(coin);
        //row = row+2;
      //}
      //else{
        coinForecast.getRange('B' + String(row)).setValue('=IMPORTHTML("'+url+'","table", 1)');
        coinForecast.getRange('A' + String(row)).setValue(coin);
        row = row+15;
      }
    }
  }
  resizeAllColumns();
}

function buildVoyagerInterestSheet(){
  voyagerInterest = createSheet('Voyager Interest')

  formula = '=IMPORTHTML("https://rewards.investvoyager.com/interest/","table", 1)';
  voyagerInterest.getRange('A1').setValue(formula);

  resizeAllColumns();
}

function rowOfCoin(coin, column){
  //coin = "BTC"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Coin URLs');
  var data = sheet.getDataRange().getValues();
  //var coinRow = sheet.getRange("A:A").getValue();
  for(var i = 0; i<data.length;i++){
    if(data[i][0] == coin){ //[1] because column B
      //Logger.log((i+1))
      if (column == 'current market'){
        return data[i][1];
      }
      else if (column == 'forecast'){
        if (data[i][2] != undefined){
          return data[i][2];
        }
        else{
          return undefined
        }
      }
    }
  }
}

function createSheet(sheetName){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSheet = sheet.getSheetByName(sheetName);
  if (newSheet != null){
    newSheet.clearContents().clearFormats();
  }
  else{
    newSheet = sheet.insertSheet();
    newSheet.setName(sheetName);
  }

  newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  newSheet.activate();
  return newSheet
}

function buildCoinURLsSheet(){
  var coinURLs =
    {AAVE:{coinMarketCap:'https://coinmarketcap.com/currencies/aave/', coinPriceForcast:'https://coinpriceforecast.com/aave'},
    ADA:{coinMarketCap:'https://coinmarketcap.com/currencies/cardano/', coinPriceForcast:'https://coinpriceforecast.com/cardano-forecast-2020-2025-2030'},
    ALGO:{coinMarketCap:'https://coinmarketcap.com/currencies/algorand/', coinPriceForcast:'https://coinpriceforecast.com/algorand'},
    ATOM:{coinMarketCap:'https://coinmarketcap.com/currencies/cosmos/', coinPriceForcast:''},
    AVAX:{coinMarketCap:'https://coinmarketcap.com/currencies/avalanche/', coinPriceForcast:'https://coinpriceforecast.com/avalanche'},
    BAND:{coinMarketCap:'https://coinmarketcap.com/currencies/band-protocol/', coinPriceForcast:''},
    BAT:{coinMarketCap:'https://coinmarketcap.com/currencies/basic-attention-token/', coinPriceForcast:''},
    BCH:{coinMarketCap:'https://coinmarketcap.com/currencies/bitcoin-cash/', coinPriceForcast:'https://coinpriceforecast.com/bitcoin-cash-forecast-2020-2025-2030'},
    BSV:{coinMarketCap:'https://coinmarketcap.com/currencies/bitcoin-sv/', coinPriceForcast:'https://coinpriceforecast.com/bsv'},
    BTC:{coinMarketCap:'https://coinmarketcap.com/currencies/bitcoin/', coinPriceForcast:'https://coinpriceforecast.com/bitcoin-forecast-2020-2025-2030'},
    BTT:{coinMarketCap:'https://coinmarketcap.com/currencies/bittorrent/', coinPriceForcast:'https://coinpriceforecast.com/btt'},
    CELO:{coinMarketCap:'https://coinmarketcap.com/currencies/celo/', coinPriceForcast:'https://coinpriceforecast.com/celo'},
    CHZ:{coinMarketCap:'https://coinmarketcap.com/currencies/chiliz/', coinPriceForcast:'https://coinpriceforecast.com/chiliz'},
    CKB:{coinMarketCap:'https://coinmarketcap.com/currencies/nervos-network/', coinPriceForcast:''},
    COMP:{coinMarketCap:'https://coinmarketcap.com/currencies/compound/', coinPriceForcast:'https://coinpriceforecast.com/compound'},
    DAI:{coinMarketCap:'https://coinmarketcap.com/currencies/multi-collateral-dai/', coinPriceForcast:''},
    DASH:{coinMarketCap:'https://coinmarketcap.com/currencies/dash/', coinPriceForcast:'https://coinpriceforecast.com/dash-forecast-2020-2025-2030'},
    DGB:{coinMarketCap:'https://coinmarketcap.com/currencies/digibyte/', coinPriceForcast:''},
    DOGE:{coinMarketCap:'https://coinmarketcap.com/currencies/dogecoin/', coinPriceForcast:'https://coinpriceforecast.com/dogecoin'},
    DOT:{coinMarketCap:'https://coinmarketcap.com/currencies/polkadot-new/', coinPriceForcast:'https://coinpriceforecast.com/dot'},
    EGLD:{coinMarketCap:'https://coinmarketcap.com/currencies/elrond-egld/', coinPriceForcast:''},
    ENJ:{coinMarketCap:'https://coinmarketcap.com/currencies/enjin-coin/', coinPriceForcast:''},
    EOS:{coinMarketCap:'https://coinmarketcap.com/currencies/eos/', coinPriceForcast:'https://coinpriceforecast.com/eos'},
    ETC:{coinMarketCap:'https://coinmarketcap.com/currencies/ethereum-classic/', coinPriceForcast:'https://coinpriceforecast.com/ethereum-classic-forecast-2020-2025-2030'},
    ETH:{coinMarketCap:'https://coinmarketcap.com/currencies/ethereum/', coinPriceForcast:'https://coinpriceforecast.com/ethereum-forecast-2020-2025-2030'},
    FIL:{coinMarketCap:'https://coinmarketcap.com/currencies/filecoin/', coinPriceForcast:'https://coinpriceforecast.com/filecoin'},
    GLM:{coinMarketCap:'https://coinmarketcap.com/currencies/golem-network-tokens/', coinPriceForcast:''},
    GRT:{coinMarketCap:'https://coinmarketcap.com/currencies/the-graph/', coinPriceForcast:'https://coinpriceforecast.com/grt'},
    HBAR:{coinMarketCap:'https://coinmarketcap.com/currencies/hedera-hashgraph/', coinPriceForcast:''},
    ICX:{coinMarketCap:'https://coinmarketcap.com/currencies/icon/', coinPriceForcast:''},
    IOT:{coinMarketCap:'https://coinmarketcap.com/currencies/iot-chain/', coinPriceForcast:''},
    KNC:{coinMarketCap:'https://coinmarketcap.com/currencies/kyber-network-crystal-legacy/', coinPriceForcast:''},
    LINK:{coinMarketCap:'https://coinmarketcap.com/currencies/chainlink/', coinPriceForcast:'https://coinpriceforecast.com/chainlink'},
    LTC:{coinMarketCap:'https://coinmarketcap.com/currencies/litecoin/', coinPriceForcast:'https://coinpriceforecast.com/litecoin-forecast-2020-2025-2030'},
    LUNA:{coinMarketCap:'https://coinmarketcap.com/currencies/terra-luna/', coinPriceForcast:'https://coinpriceforecast.com/terra'},
    MANA:{coinMarketCap:'https://coinmarketcap.com/currencies/decentraland/', coinPriceForcast:'https://coinpriceforecast.com/mana'},
    MATIC:{coinMarketCap:'https://coinmarketcap.com/currencies/polygon/', coinPriceForcast:'https://coinpriceforecast.com/polygon'},
    MKR:{coinMarketCap:'https://coinmarketcap.com/currencies/maker/', coinPriceForcast:''},
    NEO:{coinMarketCap:'https://coinmarketcap.com/currencies/neo/', coinPriceForcast:'https://coinpriceforecast.com/neo-forecast-2020-2025-2030'},
    OCEAN:{coinMarketCap:'https://coinmarketcap.com/currencies/ocean-protocol/', coinPriceForcast:''},
    OMG:{coinMarketCap:'https://coinmarketcap.com/currencies/omg/', coinPriceForcast:''},
    ONT:{coinMarketCap:'https://coinmarketcap.com/currencies/ontology/', coinPriceForcast:'https://coinpriceforecast.com/ontology'},
    OXT:{coinMarketCap:'https://coinmarketcap.com/currencies/orchid/', coinPriceForcast:''},
    QTUM:{coinMarketCap:'https://coinmarketcap.com/currencies/qtum/', coinPriceForcast:''},
    SHIB:{coinMarketCap:'https://coinmarketcap.com/currencies/shiba-inu/', coinPriceForcast:'https://coinpriceforecast.com/shib'},
    SRM:{coinMarketCap:'https://coinmarketcap.com/currencies/serum/', coinPriceForcast:''},
    STMX:{coinMarketCap:'https://coinmarketcap.com/currencies/stormx/', coinPriceForcast:''},
    SUSHI:{coinMarketCap:'https://coinmarketcap.com/currencies/sushiswap/', coinPriceForcast:'https://coinpriceforecast.com/sushi'},
    TRX:{coinMarketCap:'https://coinmarketcap.com/currencies/tron/', coinPriceForcast:'https://coinpriceforecast.com/tron'},
    TUSD:{coinMarketCap:'https://coinmarketcap.com/currencies/trueusd/', coinPriceForcast:''},
    UMA:{coinMarketCap:'https://coinmarketcap.com/currencies/uma/', coinPriceForcast:''},
    UNI:{coinMarketCap:'https://coinmarketcap.com/currencies/uniswap/', coinPriceForcast:'https://coinpriceforecast.com/uniswap'},
    USDC:{coinMarketCap:'https://coinmarketcap.com/currencies/usd-coin/', coinPriceForcast:''},
    USDT:{coinMarketCap:'https://coinmarketcap.com/currencies/tether/', coinPriceForcast:''},
    VET:{coinMarketCap:'https://coinmarketcap.com/currencies/vechain/', coinPriceForcast:'https://coinpriceforecast.com/vechain'},
    VGX:{coinMarketCap:'https://coinmarketcap.com/currencies/voyager-token/', coinPriceForcast:''},
    XLM:{coinMarketCap:'https://coinmarketcap.com/currencies/stellar/', coinPriceForcast:'https://coinpriceforecast.com/stellar-forecast-2020-2025-2030'},
    XMR:{coinMarketCap:'https://coinmarketcap.com/currencies/monero/', coinPriceForcast:'https://coinpriceforecast.com/monero-forecast-2020-2025-2030'},
    XTZ:{coinMarketCap:'https://coinmarketcap.com/currencies/tezos/', coinPriceForcast:'https://coinpriceforecast.com/tezos'},
    XVG:{coinMarketCap:'https://coinmarketcap.com/currencies/verge/', coinPriceForcast:''},
    YFI:{coinMarketCap:'https://coinmarketcap.com/currencies/yearn-finance/', coinPriceForcast:'https://coinpriceforecast.com/yfi'},
    ZEC:{coinMarketCap:'https://coinmarketcap.com/currencies/zcash/', coinPriceForcast:'https://coinpriceforecast.com/zcash-forecast-2020-2025-2030'},
    ZRX:{coinMarketCap:'https://coinmarketcap.com/currencies/0x/', coinPriceForcast:''}
    };
  var coinURLsSheet = createSheet("Coin URLs");
  coinURLsSheet.getRange('B1').setValue('CoinMarketCap URLs');
  coinURLsSheet.getRange('C1').setValue('CoinPriceForecast URLs');

  var row = 2
  for ([coin, url] of Object.entries(coinURLs)){
    coinURLsSheet.getRange('A' + String(row)).setValue(coin);
    coinURLsSheet.getRange('B' + String(row)).setValue(url['coinMarketCap']);
    if (url['coinPriceForcast'] != ''){
      coinURLsSheet.getRange('C' + String(row)).setValue(url['coinPriceForcast']);
    }
    else{
      coinURLsSheet.getRange('C' + String(row)).setValue('N/A');
    }
    row = row+1;
  }
  resizeAllColumns();
}
