function buildGainsSheet(){
  gainsSheet = createSheet('2 Gains')
  gainsSheet.getRange('A1').setValue('Value');
}

function getAveBalDict(sym){
  //var sym = "DOT/USD"
  var mth = new Date().getMonth()
  var date = new Date();
  var lastDayOfThisMonth = new Date(date.getFullYear(), mth+1, 0).getDate();
  var today = date.getDate();
  var dayTotals = []
  var aveDailyBal
  var i = 0
  var x = 1
  var daysSkipped = 1
  var thisTransDate
  var nextTransDate
  //var s = SpreadsheetApp.getActiveSpreadsheet();
  //var sht = s.getSheetByName('Transaction History')
  //var drng = sht.getDataRange();
  //var rng = sht.getRange(2,1, drng.getLastRow()-1,drng.getLastColumn());
  //var rngA = rng.getValues();//Array of input values


  for(var i = 0; i < rngA.length; i++){  //iterate through all transactions

    if (rngA[i][1].indexOf(sym)>-1 && new Date(rngA[i][2]).getMonth() == mth){ //symbol and transaction month match what was requested

      console.log("Row: " + i + ", " + rngA[i][1] + ", " + new Date(rngA[i][2]) + ", " + rngA[i][5])
      thisTransDate = new Date(rngA[i][2]).getDate()

      if (thisTransDate == new Date(rngA[i+1][2]).getDate()){  //Check if this day has multiple transactions
        console.log("Next transaction occurs on the same day. Skipping this iteration...")
        daysSkipped++
        continue
      } else if(thisTransDate == new Date(rngA[i+1][2]).getDate() && thisTransDate == 1){ //Transaction ocurred on the 1st so yesterdays total is todays.
        if (rngA[i-1][1] == sym){
          console.log("Transaction ocurred on 1st. Setting to yesterdays total: " + String(rngA[i-1][5]-rngA[i-1][4]))
          dayTotals[0] = rngA[i-1][5]-rngA[i-1][4]
        } else{
          console.log("Day: " + String(x) + " Total: 0")
          dayTotals[0] = 0
        }
        continue
      } else if (thisTransDate == 1 && rngA[i-1][1] == sym){  //transactions don't take affect until the following day. so this date gets yesterday's total
        if (rngA[i-1][1] == sym){
          console.log("Day: " + String(x) + " Total: " + rngA[i-1][5])
          dayTotals[i] = rngA[i-1][5]
        } else{
          console.log("Day: " + String(x) + " Total: 0")
          dayTotals[i] = 0
        }

        x++
      }

      nextTransDate = new Date(rngA[i+1][2]).getDate()

      if (rngA[i-daysSkipped][1] != sym && x == 1){ //this is the first transaction ever made for this symbol and every day's total quantity through today is zero.
          console.log("First ever transaction for " + sym + ". Setting preceeding days total to zero.")
        while (x <= thisTransDate){
          console.log("Day: " + String(x) + " Total: 0")
          dayTotals[x] = 0
          x++
        }
      }

      if (rngA[i+1][1].indexOf(sym) == -1 || new Date(rngA[i+1][2]).getMonth() != mth){  //Check if this is the last transaction for the month
        while (x <= lastDayOfThisMonth && x <= today - 1){
          if (rngA[i-1][1] == sym || x > 1){  //Make sure the last transaction was the same symbol.
            console.log("Day: " + String(x) + " Total: " + rngA[i][5])
            dayTotals[x] = rngA[i][5]
          } else{ //Otherwise, this is the first transaction ever made for this symbol and today's total quantity is zero.
            console.log("Day: " + String(x) + " Total: 0")
            dayTotals[x] = 0
          }
          x++
        }
      }
      else{  //there's more transactions for this symbol in this month
        while ( x <= nextTransDate){  //only process dates leading up to the day of next transaction
        console.log("Day: " + String(x) + " Total: " + rngA[i][5])
        dayTotals[x] = rngA[i][5]
        x++
        }
      }
    }
  }

  if (dayTotals.length == 0){ //No transactions found for the month. Get latest total quantity and carry it forward
    i = rngA.length-1
    while (rngA[i][1].indexOf(sym)<0){
      i--
      dayTotals[0] = rngA[i][5]
    }
  }

  if (dayTotals.length == 1){
    console.log("Monthly ave balance calculate as: " + String(dayTotals[0]))
    aveDailyBal = dayTotals[0]
  }
  else{
    console.log("Monthly ave balance calculate as: " + String(dayTotals.reduce((a, b) => a + b, 0) / (x-1)))
    aveDailyBal = dayTotals.reduce((a, b) => a + b, 0) / (x-1)
  }

  return aveDailyBal;
}

// Calculate average balance for this month
function getAveBal(sym){
  //var sym = "DOT/USD"
  var mth = new Date().getMonth()
  var date = new Date();
  var lastDayOfThisMonth = new Date(date.getFullYear(), mth+1, 0).getDate();
  var today = date.getDate();
  var dayTotals = []
  var aveDailyBal
  var i = 0
  var x = 1
  var daysSkipped = 1
  var thisTransDate
  var nextTransDate
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var sht = s.getSheetByName('Transaction History')
  var drng = sht.getDataRange();
  var rng = sht.getRange(2,1, drng.getLastRow()-1,drng.getLastColumn());
  var rngA = rng.getValues();//Array of input values


  for(var i = 0; i < rngA.length; i++){  //iterate through all transactions

    if (rngA[i][1].indexOf(sym)>-1 && new Date(rngA[i][2]).getMonth() == mth){ //symbol and transaction month match what was requested

      console.log("Row: " + i + ", " + rngA[i][1] + ", " + new Date(rngA[i][2]) + ", " + rngA[i][5])
      thisTransDate = new Date(rngA[i][2]).getDate()

      if (thisTransDate == new Date(rngA[i+1][2]).getDate()){  //Check if this day has multiple transactions
        console.log("Next transaction occurs on the same day. Skipping this iteration...")
        daysSkipped++
        continue
      } else if(thisTransDate == new Date(rngA[i+1][2]).getDate() && thisTransDate == 1){ //Transaction ocurred on the 1st so yesterdays total is todays.
        if (rngA[i-1][1] == sym){
          console.log("Transaction ocurred on 1st. Setting to yesterdays total: " + String(rngA[i-1][5]-rngA[i-1][4]))
          dayTotals[0] = rngA[i-1][5]-rngA[i-1][4]
        } else{
          console.log("Day: " + String(x) + " Total: 0")
          dayTotals[0] = 0
        }
        continue
      } else if (thisTransDate == 1 && rngA[i-1][1] == sym){  //transactions don't take affect until the following day. so this date gets yesterday's total
        if (rngA[i-1][1] == sym){
          console.log("Day: " + String(x) + " Total: " + rngA[i-1][5])
          dayTotals[i] = rngA[i-1][5]
        } else{
          console.log("Day: " + String(x) + " Total: 0")
          dayTotals[i] = 0
        }

        x++
      }

      nextTransDate = new Date(rngA[i+1][2]).getDate()

      if (rngA[i-daysSkipped][1] != sym && x == 1){ //this is the first transaction ever made for this symbol and every day's total quantity through today is zero.
          console.log("First ever transaction for " + sym + ". Setting preceeding days total to zero.")
        while (x <= thisTransDate){
          console.log("Day: " + String(x) + " Total: 0")
          dayTotals[x] = 0
          x++
        }
      }

      if (rngA[i+1][1].indexOf(sym) == -1 || new Date(rngA[i+1][2]).getMonth() != mth){  //Check if this is the last transaction for the month
        while (x <= lastDayOfThisMonth && x <= today - 1){
          if (rngA[i-1][1] == sym || x > 1){  //Make sure the last transaction was the same symbol.
            console.log("Day: " + String(x) + " Total: " + rngA[i][5])
            dayTotals[x] = rngA[i][5]
          } else{ //Otherwise, this is the first transaction ever made for this symbol and today's total quantity is zero.
            console.log("Day: " + String(x) + " Total: 0")
            dayTotals[x] = 0
          }
          x++
        }
      }
      else{  //there's more transactions for this symbol in this month
        while ( x <= nextTransDate){  //only process dates leading up to the day of next transaction
        console.log("Day: " + String(x) + " Total: " + rngA[i][5])
        dayTotals[x] = rngA[i][5]
        x++
        }
      }
    }
  }

  if (dayTotals.length == 0){ //No transactions found for the month. Get latest total quantity and carry it forward
    i = rngA.length-1
    while (rngA[i][1].indexOf(sym)<0){
      i--
      dayTotals[0] = rngA[i][5]
    }
  }



  if (dayTotals.length == 1){
    console.log("Monthly ave balance calculate as: " + String(dayTotals[0]))
    aveDailyBal = dayTotals[0]
  }
  else{
    console.log("Monthly ave balance calculate as: " + String(dayTotals.reduce((a, b) => a + b, 0) / (x-1)))
    aveDailyBal = dayTotals.reduce((a, b) => a + b, 0) / (x-1)
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



function gainsForumulas(coin){
  voyagerInterest = new Array;

  var row = 4
  for([transaction, value] of Object.entries(coin)){
    voyagerInterest[coin] = '=SUBSTITUTE(INDEX("Voyager Interest!"$B:$C,MATCH("*' + coin + '*","Voyager Interest"!$C:$C,0),1), "*","")';
    qty[coin] = '=SUM(SUMIFS("Voyager CSV"!$G:$G,"Voyager CSV"!$E:$E,"'+coin+'","Voyager CSV"!$C:$C,"Buy")+SUMIFS("Voyager CSV"!$G:$G,"Voyager CSV"!$E:$E,"'+coin+'","Voyager CSV"!$C:$C,"deposit"))';
    expectedInterest[coin] = '=IF(C'+row+'<>"#N/A",$E'+row+'*$C'+row+'/12)'
    aveCostPerCoin[coin] = '=SUMIFS("Voyager CSV"!$H:$H,"Voyager CSV"!$E:$E,"'+coin+'","Voyager CSV"!$D:$D,"<>REWARD")/SUMIFS("Voyager CSV"!$G:$G,"Voyager CSV"!$E:$E,""+coin+"", "Voyager CSV"!$D:$D, "<>REWARD")';
    totalCost[coin] = '$D'+row+'*$G'+row;
    currentPrice[coin] = '=INDEX("Current Market"!$A:$C,MATCH("'+coin+'","Current Market"!$A:$A,0),3)';
    currentValue[coin] = '$D'+row+'*$J'+row;
    gain[coin] = '($D'+row+'*$J'+row+')-($D'+row+'*$G'+row+')';
    percentGain[coin] = '($D'+row+'*$J'+row+')/($D'+row+'*$G'+row+')-1';
    totalInterestEarned[coin] = '=SUMIFS("Voyager CSV"!$H:$H,"Voyager CSV"!$E:$E,"'+coin+'","Voyager CSV"!$D:$D,"=INTEREST")';
    forcastValue[coin]["2021"] = '=SUBSTITUTE(INDEX("Coin Forecast"!$A:$E,MATCH("'+coin+'","Coin Forecast"!$A:$A,0)+1,4), "*","")'; //+2,4 for 2022 etc.
    row += 1;
  }
}

function buildCurrentMarket(coins){
  var currentMarket = createSheet('Current Market')
  var row = 1;

  for (coin in coins){
    if (coin != 'total_interest' && coin != 'USD'){
      Logger.log(coin);
      url = rowOfCoin(coin, 'current market');
      currentMarket.getRange('B' + String(row)).setValue('=IMPORTHTML("'+url+'","table", 1)');
      currentMarket.getRange('A' + String(row)).setValue(coin);
      row = row+8;
    }
    resizeAllColumns();
  }
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
  var coinURLsSheet = createSheet("2 Coin URLs");
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
