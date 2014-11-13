var spreadsheetId = ""
var spreadsheetUrl = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit";
var spreadsheet = null;
var sheetName = "";
var sheet = null;
var type = "";

/**
 *
 */
function authorize() {
}

/**
 *
 */
function doGet(e) {
  var result = "";
  sheetName = e.parameter.sheet;
  type = e.parameter.type;
  
  
  spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  if (spreadsheet != null) {
    if (sheetName==null) {
      return HtmlService.createTemplateFromFile('index').evaluate();
    }
    else {
      sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet != null) {
        return HtmlService.createTemplateFromFile('form').evaluate();
      }
      else{
        msg = "Error: I cannot access the requested sheet, dew."
        Logger.log(msg)
        return HtmlService.createTemplateFromFile('error').evaluate();
      }
    }
  }
  else {
    msg = "Error: I cannot access the active spreadsheet, dew."
    Logger.log(msg)
    return HtmlService.createTemplateFromFile('error').evaluate();
  }
}

/**
*
*/
function retrieveSheetsList() {
  var result = [];
  var sheets = spreadsheet.getSheets();
  for (var sheetIndex in sheets) {
    var sheet = sheets[sheetIndex];
    result.push(sheet.getName());
  }
  return result;
}

/**
*
*/
function retrieveSheetFormTest() {
  spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  sheet = spreadsheet.getSheetByName("Hoja 1");
  return  retrieveSheetForm();
}

/**
*
*/
function retrieveSheetForm() {
  var result = [];
  if (sheet != null) {
    var headers = getHeadings(this.sheet, 0, undefined);
    for (var headerIndex in headers["column"]) {
      var value = headers["column"][headerIndex]["value"];
      var note = headers["column"][headerIndex]["note"];
      var resultObj = {};
      try {
        var noteParsed = JSON.parse(note);
        
        resultObj["columnname"] = value;
        resultObj["title"] = noteParsed.title || value;
        resultObj["description"] = noteParsed.description;
        resultObj["type"] = noteParsed.type;
        resultObj["defvalue"] = noteParsed.defvalue || "";
        resultObj["mandatory"] = noteParsed.mandatory;
        resultObj["items"] = noteParsed.items;
        /*
        {"Type":"", #Puede ser "autoincrement" que ademas de hidden pilla la anterior y suma uno, o "autodate" que pone la fecha actual inmodificable, autotime autodatetime
        # o si no number, text, textarea, checkbox, date, time, select
        "Title":"",
        "Description":"",
        "DefValue":"",
        "Mandatory":""
        "Items":[] Siempre que el tipo sea select se tiene en cuenta
        }
        */
        result.push(resultObj);
      }
      catch(err) {
      }
    }
  }
  return result;
}

/**
*
*/
function getNewIncrementInSheet(columnname) {
  var result = 0;
  if (sheet != null) {
    var columnNumber = getColumnByName(sheet, columnname);
    var range = sheet.getRange(1, columnNumber, sheet.getMaxRows(), 1);
    var values = range.getValues();
    
    for (var index in values) {
      if (result < parseInt(values[index])) result = parseInt(values[index])
    }
  }
  return result + 1;
}

/**
*
*/
function processForm(formObject) {
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheetName = formObject["sheet"];
  var sheet = spreadsheet.getSheetByName(sheetName);
  var rowNumber = getLastRow(sheet) + 1;
  
  // First loop to check for corrections
  for (var propertyName in formObject) {
    var columnNumber = getColumnByName(sheet, propertyName);
    if (columnNumber) {
      var range = sheet.getRange(rowNumber, columnNumber, 1, 1);
      var note = sheet.getRange(1, columnNumber, 1, 1).getNote();
      var value = formObject[propertyName];
      try {
        var noteParsed = JSON.parse(note);
      }
      catch(err) {
      }
      
      if (noteParsed) {
        var title = noteParsed.title;
        var mandatory = noteParsed.mandatory;
        
        if (mandatory == "yes" && (value == undefined || value == "")) {
          msg = "Mandatory not satisfied in " + title;
          Logger.log(msg)
          throw msg
        }
      }
    }
  }
  
  // Second loop, everything ok!
  for (var propertyName in formObject) {
    var columnNumber = getColumnByName(sheet, propertyName);
    if (columnNumber) {
      var range = sheet.getRange(rowNumber, columnNumber, 1, 1);
      var note = sheet.getRange(1, columnNumber, 1, 1).getNote();
      var value = formObject[propertyName];
      try {
        var noteParsed = JSON.parse(note);
      }
      catch(err) {
      }
      
      if (noteParsed) {
        var title = noteParsed.title;
        var description = noteParsed.description;
        var type = noteParsed.type;
        var defvalue = noteParsed.defvalue;
        var mandatory = noteParsed.mandatory;
        var items = noteParsed.items;
        
        if (mandatory == "yes" && (value == undefined || value == "")) {
          msg = "Mandatory not satisfied in " + propertyName;
          Logger.log(msg)
          throw msg
        }
        else {
          range.setValue(formObject[propertyName]);
          if (type == "autoincrement") {
          } 
          if (type == "autodate") {
          } 
          if (type == "autotime") {
          } 
          if (type == "autodatetime") { 
          } 
          if (type == "number") { 
          }
          if (type == "text") { 
          } 
          if (type == "textarea") {
          } 
          if (type == "checkbox") {
          } 
          if (type == "date") { 
          } 
          if (type == "time") {
          } 
          if (type == "select") { 
          } 
          Logger.log("Inserting " + propertyName + " in row " + rowNumber + " and column " + columnNumber + " contents " + value)
        }
      }
    }
  }
}
