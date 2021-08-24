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
  currentMarket = createSheet('Current Market')
  row = 1;

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
  row = 1;

  for (coin in coins){
    if (coin != 'total_interest' && coin != 'USD'){
      Logger.log(coin);
      url = rowOfCoin(coin, 'forecast');
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
