/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
//function doGet(request) {
//  return HtmlService.createTemplateFromFile('Index')
//      .evaluate();
//}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
/**
 * Get the URL for the Google Apps Script running as a WebApp.
 */
function getScriptUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}

/**
 * Get "home page", or a requested page.
 * Expects a 'page' parameter in querystring.
 *
 * @param {event} e Event passed to doGet, with querystring
 * @returns {String/html} Html to be served
 */
function doGet(e) {
  Logger.log( Utilities.jsonStringify(e) );
  if (!e.parameter.page) {
    // When no specific page requested, return "home page"
    return HtmlService.createTemplateFromFile('Index').evaluate();
  }
  // else, use page parameter to pick an html file from the script
  return HtmlService.createTemplateFromFile(e.parameter['page']).evaluate();
}
function isFloat(n){
    return Number(n) === n && n % 1 !== 0;
}
function get2018Trades()
{
  var sheetName = '2018';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  Logger.log(sheet);
  var range = sheet.getRange('B2:K400').getValues();
  for (var i = 0; i < range.length; i++) {
    for (var j = 0; j < range[i].length; j++) {
      //range[i][j] = range[i][j].toString();
      if(range[i][j].getMonth) range[i][j] = (range[i][j].getMonth()+1) + '/' +range[i][j].getDate()+ '/'+range[i][j].getYear();
      if (isFloat(range[i][j]))range[i][j] = parseFloat(Math.round(range[i][j] * 100) / 100).toFixed(2);

    }
  }
   return range;
}

function testFunction()
{
  var sheetName = 'Cur';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  Logger.log(sheet);
  //var range = sheet.getRange('E3:J65').getValues();
  //var range = sheet.getRange('C3:D65').getValues();
  var range = sheet.getRange('B2:N65').getValues();
  for (var i = 0; i < range.length; i++) {
    for (var j = 0; j < range[i].length; j++) {
      //range[i][j] = range[i][j].toString();
      if(range[i][j].getMonth) range[i][j] = (range[i][j].getMonth()+1) + '/' +range[i][j].getDate();
      if (isFloat(range[i][j]))range[i][j] = parseFloat(Math.round(range[i][j] * 100) / 100).toFixed(2);

    }
  }
   return range;
}
function getCurrentPositions()
{
  var sheetName = 'Cur';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  Logger.log(sheet);
  var range = sheet.getRange('B2:U90').getValues();
  Logger.log(range);

  return range;
}

  /**
  * @Description - Gets the midprice of a two option spread where one strike is short (sold) and the other is long (bought)
  *
  * @param optionList (array[object]) - array of objects containing, among other things, the bid and ask price of each option
  * @param shortStrike (int) - integer representing shortStrike
  * @param longStrike (int) - integer representing longStrike
  *
  * @return (array[float]) - returns a two element array of [mid price, last price] of an option spread
  */
  function getTwoOptionMidPrice(options, type, shortStrike, longStrike)
  {
      var i, shortMidPrice, longMidPrice;
      var shortLastPrice, longLastPrice;
      for (i = 0; i < options.length; i++)
      {
        if (options[i].option_type.toLowerCase() !== type) continue;
        if (options[i].strike == shortStrike)
        {
          shortLastPrice = options[i].last;
          shortMidPrice = (options[i].bid + options[i].ask)/2;
        }
        if (options[i].strike == longStrike)
        {
          longLastPrice = options[i].last;   
          longMidPrice = (options[i].bid + options[i].ask)/2;   
        }
      }
      return [(shortMidPrice - longMidPrice).toFixed(2), (shortLastPrice - longLastPrice).toFixed(2)];
  }
  /**
  * @Description - Gets the midprice of a put spread
  *
  * @param puts (array[object]) - array of objects representing put options containing, among other things, the bid and ask price of each option
  * @param shortStrike (int) - integer representing shortStrike
  * @param longStrike (int) - integer representing longStrike
  *
  * @return (array[float]) - returns a two element array of [mid price, last price] of a put spread
  */
  function getPutSpreadMidPrice(options, shortStrike, longStrike)
  {
    return getTwoOptionMidPrice(options, "put", shortStrike, longStrike);
  }

  /**
  * @Description - Gets the mid price of a call spread
  *
  * @param calls (array[object]) - array of objects representing call options containing, among other things, the bid and ask price of each option
  * @param shortStrike (int) - integer representing shortStrike
  * @param longStrike (int) - integer representing longStrike
  *
  * @return (array[float]) - returns a two element array of [mid price, last price] of a call spread
  */
  function getCallSpreadMidPrice(options, shortStrike, longStrike)
  {
    return getTwoOptionMidPrice(options, "call", shortStrike, longStrike);
  }
  
  /**
  * @Description - Gets the mid price of an Iron Condor
  *
  * @param puts (array[object]) - array of objects representing call options containing, among other things, the bid and ask price of each option
  * @param calls (array[object]) - array of objects representing call options containing, among other things, the bid and ask price of each option
  *
  * @param shortStrike (int) - integer representing put shortStrike
  * @param longStrike (int) - integer representing put longStrike
  *
  * @param shortStrike2 (int) - integer representing call shortStrike
  * @param longStrike2 (int) - integer representing call longStrike
  *
  * @return (array[float]) - returns a two element array of [mid price, last price] of an Iron Condor spread
  */
  function getICSpreadMidPrice(options, shortStrike, longStrike, shortStrike2, longStrike2)
  {
    var putSpread = getPutSpreadMidPrice(options, shortStrike, longStrike);
    var callSpread = getCallSpreadMidPrice(options, shortStrike2, longStrike2);
    return [(+putSpread[0] + +callSpread[0]).toFixed(2), (+putSpread[1] + +callSpread[1]).toFixed(2)];
  }

  /**
  * @Description - Gets the mid price of an Iron Butterfly
  *
  * @param puts (array[object]) - array of objects representing call options containing, among other things, the bid and ask price of each option
  * @param calls (array[object]) - array of objects representing call options containing, among other things, the bid and ask price of each option
  *
  * @param shortStrike (int) - integer representing put shortStrike
  * @param longStrike (int) - integer representing put longStrike
  *
  * @param shortStrike2 (int) - integer representing call shortStrike
  * @param longStrike2 (int) - integer representing call longStrike
  *
  * @return (array[float]) - returns a two element array of [mid price, last price] of an Iron Butterfly spread
  */
  function getIBSpreadMidPrice(options, shortStrike, longStrike, shortStrike2, longStrike2)
  {
    return getICSpreadMidPrice(options, shortStrike, longStrike, shortStrike2, longStrike2);
  }

  /**
  * @Description - Gets the mid price of an four types of option strategies: Put, Call, Iron Condor, and Iron Butterfly Strategies
  *
  * @param parsed (object) - a JSON object returning multiple options all expiring on the same date
  * @param spreadType (text) - string representing the type of spread: "Put", "Call", "IC", "IB"
  *
  * @param shortStrike (int) - integer representing put shortStrike
  * @param longStrike (int) - integer representing put longStrike
  *
  * @param shortStrike2 (int) - integer representing call shortStrike
  * @param longStrike2 (int) - integer representing call longStrike
  *
  * @return (array[float]) - returns a two element array of [mid price, last price] of an option spread
  */
  function spreadMidCalculator(parsed, spreadType, shortStrike, longStrike, shortStrike2, longStrike2)
  {
    var options = parsed.options.option;
    var spreadMidPrice;
    if (spreadType === "Put") spreadMidPrice = getPutSpreadMidPrice(options, shortStrike, longStrike);
    else if (spreadType === "Call") spreadMidPrice = getCallSpreadMidPrice(options, shortStrike, longStrike);
    else if (spreadType === "IC") spreadMidPrice = getICSpreadMidPrice(options, shortStrike, longStrike, shortStrike2, longStrike2);
    else if (spreadType === "IB") spreadMidPrice = getIBSpreadMidPrice(options, shortStrike, longStrike, shortStrike2, longStrike2);
    return spreadMidPrice;
  }


function getSpreadPriceFromTradier(ticker, spreadType, exDate, shortStrike, longStrike, shortStrike2, longStrike2, numberOfContracts, credit, midCellId, lastCellId, pandLCellId) {
  var returnArray = ['Not Responding', 'Not Responding', 'Not Responding', midCellId, lastCellId, pandLCellId];


  var queryString = "https://sandbox.tradier.com/v1/markets/options/chains?symbol="+ticker+"&expiration="+convertDateForTradier(exDate);
  var options = {
        headers : {
          "Authorization" : "Bearer PDcI9Z8ztnqwfnsVFkFULUfX2YGB",
          "Accept" : "application/json" 
        }
      };
  Utilities.sleep(Math.random() * 1000);
  var result = UrlFetchApp.fetch(queryString, options);
  if (result.getResponseCode() == 200) {

    var json = result.getContentText();
    var parsed = JSON.parse(json);

    midPrice = spreadMidCalculator(parsed, spreadType, shortStrike, longStrike, shortStrike2, longStrike2);
    returnArray[0] = midPrice[0];
    returnArray[1] = midPrice[1];
    returnArray[2] = (((credit - midPrice[0])*100)*numberOfContracts).toFixed(2);
  }  
  return returnArray;
}
function convertDateForTradier(exDate)
{
  var split = exDate.split('/');
  return split[2]+'-'+split[0]+'-'+split[1];
}
function testGET() {
  var queryString = "https://sandbox.tradier.com/v1/markets/options/chains?symbol=adbe&expiration=2018-11-02";
  
  var url = queryString;
  
  var options =
      {
        headers :
        {
          "Authorization" : "Bearer PDcI9Z8ztnqwfnsVFkFULUfX2YGB",
          "Accept" : "application/json" 
        }
      };
    
  var result = UrlFetchApp.fetch(url, options);
  var json = result.getContentText();
  var data = JSON.parse(json);
  
  return JSON.stringify(data.options.option[0]);
    

  if (result.getResponseCode() == 200) {
      return JSON.parse(result);
    
  }  
  else return "error";
}