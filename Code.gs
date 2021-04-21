/**
 * Source: github.com/sergioloc
 */

var titles;

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Export ARB",
    functionName : "exportARB"
  }];
  sheet.addMenu("Export", entries);
};

// triggers parsing and displays results in a text area inside a custom modal window
function exportARB() {
  var languages = makeJson(SpreadsheetApp.getActiveSheet().getDataRange());
  var textHtml = "";
  for (var i = 0; i < languages.length ; ++i) {
    textHtml += "<h1 style=\"font-family:courier;\">" + titles[i] + "</h1> <textarea style='width:100%;' rows='20'>" + languages[i] + "</textarea>"
  }
  displayText_(textHtml);
};

function makeJson(dataRange) {
  var charSep = '"';
  titles = Array();
  var keys = Array(), final = Array();
  var result = "", thisName = "", thisData = "";

  var frozenRows = SpreadsheetApp.getActiveSheet().getFrozenRows();
  var frozenColumns = SpreadsheetApp.getActiveSheet().getFrozenColumns();

  var dataRangeArray = dataRange.getValues();
  var dataWidth = dataRange.getWidth() - frozenColumns;
  var dataHeight = dataRange.getHeight() - frozenRows;
  
  for (var i = 0; i < dataHeight ; ++i) {
      keys.push(dataRangeArray[i+frozenRows][frozenColumns-1]);
  }

  for (var i = 0; i < dataWidth ; ++i) {
      titles.push(dataRangeArray[frozenRows-1][i+frozenColumns]);
  }

  for (var h = 0; h < dataWidth ; ++h) {
    
    result = '{\n';
    
    for (var i = 0; i < dataHeight; ++i) {
      thisName = keys[i];
      thisData = dataRangeArray[i+frozenRows][h+frozenColumns];

      // add name 
      result += '   ' + charSep + thisName + charSep + ': '
        
      // add data
      result += charSep + jsonEscape(thisData) + charSep + ',\n';
    }
    //remove last comma and space
    result = result.slice(0,-2);
    result += '\n}';
    final.push(result);
  }
  
  return final;
    
}

function jsonEscape(str)  {
  if (typeof str === "string" && str !== "") {
    return str.replace(/\n/g, "<br/>").replace(/\r/g, "<br/>").replace(/\t/g, "\\t");
  } else {
    return str;
  }
}

function displayText_(textHtml) {
  var output = HtmlService.createHtmlOutput(textHtml);
  output.setWidth(1000)
  output.setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(output, 'ARB exported');
}