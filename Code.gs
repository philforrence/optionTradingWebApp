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

function getTwoOptionMidPrice(optionList, shortStrike, longStrike)
{
    var i, longMidPrice, shortMidPrice;
    for (i = 0; i < optionList.length; i++)
    {
      if (optionList[i].strike == shortStrike)
        shortMidPrice = (optionList[i].bid + optionList[i].ask)/2;
      if (optionList[i].strike == longStrike)
        longMidPrice = (optionList[i].bid + optionList[i].ask)/2;    
    }
    return shortMidPrice - longMidPrice;
}
function getPutSpreadMidPrice(puts, shortStrike, longStrike)
{
  return getTwoOptionMidPrice(puts, shortStrike, longStrike);
}
function getCallSpreadMidPrice(calls, shortStrike, longStrike)
{
  return getTwoOptionMidPrice(calls, shortStrike, longStrike);
}
function getICSpreadMidPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2)
{
  return getPutSpreadMidPrice(puts, shortStrike, longStrike) + getCallSpreadMidPrice(calls, shortStrike2, longStrike2);
}
function getIBSpreadMidPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2)
{
  return getICSpreadMidPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2);
}
function spreadMidCalculator(parsed, spreadType, shortStrike, longStrike, shortStrike2, longStrike2)
{
  var puts = parsed.optionChain.result[0].options[0].puts;
  var calls = parsed.optionChain.result[0].options[0].calls;
  var spreadMidPrice = 0;
  if (spreadType == "Put") spreadMidPrice = getPutSpreadMidPrice(puts, shortStrike, longStrike);
  else if (spreadType == "Call") spreadMidPrice = getCallSpreadMidPrice(calls, shortStrike, longStrike);
  else if (spreadType == "IC") spreadMidPrice = getICSpreadMidPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2);
  else if (spreadType == "IB") spreadMidPrice = getIBSpreadMidPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2);
  return spreadMidPrice;
}

function getTwoOptionLastPrice(optionList, shortStrike, longStrike)
{
    var i, longLastPrice, shortLastPrice;
    for (i = 0; i < optionList.length; i++)
    {
      if (optionList[i].strike == shortStrike)
        shortLastPrice = optionList[i].lastPrice;
      if (optionList[i].strike == longStrike)
        longLastPrice = optionList[i].lastPrice;   
    }
    return shortLastPrice - longLastPrice;
}
function getPutSpreadLastPrice(puts, shortStrike, longStrike)
{
  return getTwoOptionLastPrice(puts, shortStrike, longStrike);
}
function getCallSpreadLastPrice(calls, shortStrike, longStrike)
{
  return getTwoOptionLastPrice(calls, shortStrike, longStrike);
}
function getICSpreadLastPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2)
{
  return getPutSpreadLastPrice(puts, shortStrike, longStrike) + getCallSpreadLastPrice(calls, shortStrike2, longStrike2);
}
function getIBSpreadLastPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2)
{
  return getICSpreadLastPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2);
}
function spreadLastCalculator(parsed, spreadType, shortStrike, longStrike, shortStrike2, longStrike2)
{
  var puts = parsed.optionChain.result[0].options[0].puts;
  var calls = parsed.optionChain.result[0].options[0].calls;
  var spreadLastPrice = 0;
  if (spreadType == "Put") spreadLastPrice = getPutSpreadLastPrice(puts, shortStrike, longStrike);
  else if (spreadType == "Call") spreadLastPrice = getCallSpreadLastPrice(calls, shortStrike, longStrike);
  else if (spreadType == "IC") spreadLastPrice = getICSpreadLastPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2);
  else if (spreadType == "IB") spreadLastPrice = getIBSpreadLastPrice(puts, calls, shortStrike, longStrike, shortStrike2, longStrike2);
  return spreadLastPrice;
}
function getSpreadPrice(ticker, spreadType, exDate, shortStrike, longStrike, shortStrike2, longStrike2, type)
{
  var queryString = "https://query1.finance.yahoo.com/v7/finance/options/" + ticker+"?date="+exDate;
  var data = UrlFetchApp.fetch(queryString);
  
  
  var jsonData = data.getContentText();
  var parsed = JSON.parse(jsonData);
  
  var returnSpreadPrice = "hello";
  if (type=="Last") returnSpreadPrice = spreadLastCalculator(parsed, spreadType, shortStrike, longStrike, shortStrike2, longStrike2);
  else if (type == "Mid") returnSpreadPrice = spreadMidCalculator(parsed, spreadType, shortStrike, longStrike, shortStrike2, longStrike2);
  
  if (isNaN(returnSpreadPrice)) returnSpreadPrice = 0;
  return returnSpreadPrice;

}
