/********************************************
             Spreadsheet related
********************************************/


/////////////FETCH ///////////////

/**
* 
*/ 
function filterSheetsByName(sheets, name) {
  var result = new Array();
  for (var index=0; index<sheets.length; index++) {
    var sheet = sheets[index];
    if (sheet != null && sheet.getName().indexOf(name) != -1) {
      result.push(sheet);
    }
  }
  return result
}

/**
* Recupera el numero de columna cuyo identificador es el num_esimo de nombre name
* Ojo que coje tanto mayusculas como minusculas!
*/ 
function getColumnByName(sheet, name, num) {
  if (sheet != null) {
    //Fetch all at once
    var firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()
    for (var i=0; i<firstRow[0].length; i=i+1){
      if (firstRow[0][i].toLowerCase() == name.toLowerCase()) {
        if (num == 1 || num == undefined) {
          return i + 1
        }
        else {
          num = num - 1
        }  
      }
    }
  }
  return null
}

/*
* Ojo que coje tanto mayusculas como minusculas!
*/ 
function getRowByName(sheet, name, num) {
  if (sheet != null) {
    //Fetch all at once
    var firstColumn = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues()
    for (var i=0; i<firstColumn.length; i=i+1){
      if (firstColumn[i][0].toLowerCase() == name.toLowerCase()) {
        if (num == 1 || num == undefined) {
          return i + 1
        }
        else {
          num = num - 1
        }  
      }
    }
  }
  return null
}

/*
* 
*/ 
function isRowEmpty(sheet, rowNum) {
  if (sheet != null) {
    //Fetch all at once
    var row = sheet.getRange(rowNum, 1, 1, sheet.getMaxColumns())
    var rangeValues = row.getValues()
    var rangeFormulas = row.getFormulas()
    for (var column=0; column<sheet.getMaxColumns(); column = column + 1){
      if (rangeValues[0][column]!="" || rangeFormulas[0][column]!="") {
        return false 
      }
    }
    return true
  }
}

/*
* 
*/ 
function isColumnEmpty(sheet, columnNum) {
  if (sheet != null) {
    //Fetch all at once
    var column = sheet.getRange(1, columnNum, sheet.getMaxRows(), 1)
    var rangeValues = column.getValues()
    var rangeFormulas = column.getFormulas()
    for (var row=0; row<sheet.getMaxRows(); row = row + 1){
      if (rangeValues[row][0]!="" || rangeFormulas[row][0]!="") {
        return false 
      }
    }
    return true
  }
}


/*
* This returns the first empty row (without contents)
* Ex. Sheet with 10 rows, but only the first 5 filled. Returns 
*/ 
function getLastRow(sheet) {
  if (sheet != null) {
    //Fetch all at once
    var maxColumns = sheet.getMaxColumns()
    var maxRows = sheet.getMaxRows()
    var range = sheet.getRange(1, 1, maxRows, maxColumns)
    var rangeValues = range.getValues()
    var rangeFormulas = range.getFormulas()
    for (var row=maxRows-1; row>=0; row = row - 1){
      for (var column=0; column<maxColumns; column = column + 1){
        if (rangeValues[row][column]!="" || rangeFormulas[row][column]!="") {
          return row+1 
        }
      }
    }
  }
  return null
}

/*
* 
*/ 
function getFirstRowWithoutContents(sheet) {
  return getLastRowWithContents(sheet) + 1;
}

/*
* 
*/ 
function getLastRowWithContents(sheet) {
  return getLastRow(sheet);
}

/*
* 
*/ 
function getFirstColumnWithoutContents(sheet) {
  return getLastColumnWithContents(sheet) + 1;
}

/*
* 
*/ 
function getLastColumnWithContents(sheet) {
  return getLastColumn(sheet);
}

/*
* 
*/ 
function getTheVeryLastRow(sheet) {
  if (sheet != null) {
    return sheet.getMaxRows()
  }
  return null;
}

/*
* 
*/ 
function getTheVeryLastColumn(sheet) {
  if (sheet != null) {
    return sheet.getMaxColumns()
  }
  return null;
}

/**
* Retrieve headings of columns and rows, to the maximum or the defined number
*/ 
function getHeadings(sheet, numRows, numColumns) {
  var outputSheetIndex = new Array();
  outputSheetIndex["row"] = new Array();
  outputSheetIndex["column"] = new Array();
  if (sheet != null) {
    var lastRow = getLastRow(sheet)
    var lastColumn = sheet.getLastColumn()
    if (numRows != undefined) {
      lastRow = numRows
    }
    if (numColumns != undefined) {
      lastColumn = numColumns
    }
    
    for (var row = 1; row <= lastRow; row = row + 1) {
      //TODO: Improve this by fetching all the values at once
      var obj = {}
      obj["value"] = sheet.getRange(row, 1, 1, 1).getValue();
      obj["note"] = sheet.getRange(row, 1, 1, 1).getNote();
      outputSheetIndex["row"][row] = obj;
    }
    for (var column = 1; column <= lastColumn; column = column + 1) {
      var obj = {}
      obj["value"] = sheet.getRange(1, column, 1, 1).getValue();
      obj["note"] = sheet.getRange(1, column, 1, 1).getNote();
      outputSheetIndex["column"][column] = obj;
    }    
    
    return outputSheetIndex;
  }
  else return null
}



/////////////CLEAR ///////////////

/**
* Wipe this whole spreadsheet (removes all sheets but the last)
*/
function wipeThisSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear
  wipeSpreadsheet(spreadsheet);
}

/**
* Wipe the whole spreadsheet (removes all sheets but the last)
*/
function wipeSpreadsheet(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  
  // Delete all sheets but one
  while (sheets.length > 1) {
    var sheet = sheets.pop();  
    Logger.log("Deleting sheet" + sheet.getName());
    spreadsheet.deleteSheet(sheet);
  }
  
  // Delete last sheet contents
  var lastSheet = sheets.pop();
  wipeSheet(lastSheet);
}

/**
* Clear the whole sheet (alias of removeAllData)
*/
function wipeSheet(sheet) {
  removeAllData(sheet);
}

/*
* This clear all data but the cells remains
*/ 
function clearAllData(sheet) {
  if (sheet != null) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
  }
  return null
}

/*
* This leaves only one empty cell
*/ 
function removeAllData(sheet) {
  if (sheet != null) {
    var lastRow = sheet.getMaxRows();
    var lastColumn = sheet.getMaxColumns();
    if (lastRow > 1) sheet.deleteRows(1,lastRow - 1);
    if (lastColumn > 1) sheet.deleteColumns(1, lastColumn - 1);
    sheet.getRange(1, 1).clear();
  }
  return null
}

/*
* This removes the sheet from the spreadsheet
*/ 
function removeSheet(sheet) {
  if (sheet != null) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet)
      SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet()
    } catch (err){}
  }
  return null
}

/**
* 
*/ 
function removeSheetByName(sheetName) {
  return removeSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName))
}


/////////////ADDITIONS ///////////////


/*
*  This sets a single value
*/ 
function setSingleValue(sheet, row, column, value) {
  if (sheet != null) {
    sheet.getRange(row, column, 1, 1).setValue(value)
    return true
  }
  else {
    return false
  }
}

/*
*  This adds a row at the very end
*/ 
function addRowAtTheEnd(sheet, name, visualOpts) {
  if (sheet != null) {
    
    var maxRows = sheet.getMaxRows();
    sheet.insertRowAfter(maxRows)
    sheet.getRange(maxRows + 1, 1, 1, 1).setValue(name)
    
    if  (visualOpts != undefined) {      
      sheet
      .getRange(maxRows + 1,
                1, 
                1,
                sheet.getMaxColumns())
      .setFontColor(visualOpts["fontColor"])
      .setBackground(visualOpts["backgroundColor"])
    }
    
    return getLastRow(sheet)
  }
  else {
    return null
  }
}

/*
*  This adds a column at the very end
*/ 
function addColumnAtTheEnd(sheet, name) {
  if (sheet != null) {
    var maxColumns = sheet.getMaxColumns()
    sheet.insertColumnAfter(maxColumns)
    sheet.getRange(1, maxColumns + 1, 1, 1).setValue(name)
    return getLastColumn(sheet)
  }
  else {
    return null
  }
}

/**
* Set headings of columns and/or rows
*/ 
function setHeadings(sheet, rowHeadings, columnHeadings, visualOpts) {
  if (sheet != null) {
    
    wipeSheet(sheet);
    
    var backgroundColor = "white"
    var fontColor = "black"
    
    if  (visualOpts != undefined) {
      backgroundColor = visualOpts["backgroundColor"]
      fontColor = visualOpts["fontColor"]
    }
    
    if (rowHeadings != undefined) {
      for (var row = 0; row < rowHeadings.length; row = row + 1) {
        var value = rowHeadings[row]
        sheet.getRange(row + 1, 1, 1, 1).setValue(value).setFontColor(fontColor).setBackground(backgroundColor)
      }
    }
    if (columnHeadings != undefined) {
      for (var column = 0; column < columnHeadings.length; column = column + 1) {
        var value = columnHeadings[column]
        sheet.getRange(1, column + 1, 1, 1).setValue(value).setFontColor(fontColor).setBackground(backgroundColor)
      }    
    }

      
    return sheet;
  }
  else return null
}



///extra//

/*
*  This transforms the reference to abs or relative. Accepted types:
*   1 To absolute -> $A$1
*   2 To mixed relative column -> A$1
*   3 To mixed relative row -> $A1
*   4 To relative -> A1
*
*/
function transformA1Notation(input, type) {
  try {
    
    if (input != null && type != null) {
      var splittedInput=input.split(":");
      //Logger.log(splittedInput.length);
      
      if (splittedInput.length == 1 ) {
        var firstPart = splittedInput[0]
        
        var firstcolumnPatt = new RegExp("[A-Za-z]+");
        var firstcolumn = firstcolumnPatt.exec(firstPart)[0]
        var firstrowPatt = new RegExp("[0-9]+");
        var firstrow = firstrowPatt.exec(firstPart)[0]
        
        //Logger.log(firstPart);
        //Logger.log(firstcolumn);
        //Logger.log(firstrow);
        
        // To relative -> A1
        if (type == 4) return firstcolumn+firstrow
        // To absolute -> $A$1
        if (type == 1) return "$"+firstcolumn+"$"+firstrow
        // To mixed relative row -> $A1
        if (type == 3) return "$"+firstcolumn+firstrow
        // To mixed relative column -> A$1
        if (type == 2) return firstcolumn+"$"+firstrow
        
      }
      if (splittedInput.length == 2 ) {
        var firstPart = splittedInput[0]
        var lastPart = splittedInput[1]
        
        var firstcolumnPatt = new RegExp("[A-Za-z]+");
        var firstcolumn = firstcolumnPatt.exec(firstPart)[0]
        var firstrowPatt = new RegExp("[0-9]+");
        var firstrow = firstrowPatt.exec(firstPart)[0]
        var lastcolumnPatt = new RegExp("[A-Za-z]+");
        var lastcolumn  = lastcolumnPatt.exec(lastPart)[0]
        var lastrowPatt = new RegExp("[0-9]+");
        var lastrow = lastrowPatt.exec(lastPart)[0]
        
        //Logger.log(firstPart);
        //Logger.log(lastPart);
        //Logger.log(firstcolumn);
        //Logger.log(firstrow);
        //Logger.log(lastcolumn);
        //Logger.log(lastrow);
        
        // To relative -> A1
        if (type == 4) return firstcolumn+firstrow+":"+lastcolumn+lastrow
        // To absolute -> $A$1
        if (type == 1) return "$"+firstcolumn+"$"+firstrow+":"+"$"+lastcolumn+"$"+lastrow
        // To mixed relative row -> $A1
        if (type == 3) return "$"+firstcolumn+firstrow+":"+"$"+lastcolumn+lastrow
        // To mixed relative column -> A$1
        if (type == 2) return firstcolumn+"$"+firstrow+":"+lastcolumn+"$"+lastrow
        
      }
    }
    
    return null
    
  }catch(exception) {
    throw "Error in transformA1Notation with input " + input
  }
}

/**
* Deletes everything and resizes
*/ 
function resize(sheet, numRows, numColumns) {
  if (sheet != null) {
    removeAllData(sheet);
    sheet.insertRowsAfter(1, numRows - 1);
    sheet.insertColumnsAfter(1, numColumns - 1);
  }
  return null
}




//  function colorAll() {
//    var sheet = SpreadsheetApp.getActiveSheet();
//    var startRow = 2;
//    var endRow = sheet.getLastRow();
//
//    for (var r = startRow; r <= endRow; r++) {
//      colorRow(r);
//    }
//  }
//
//  function colorRow(r){
//    var sheet = SpreadsheetApp.getActiveSheet();
//    var dataRange = sheet.getRange(r, 1, 1, 3);
// 
//    var data = dataRange.getValues();
//    var row = data[0];
// 
//    if(row[0] === ""){
//      dataRange.setBackgroundRGB(255, 255, 255);
//    }else if(row[0] > 0){
//      dataRange.setBackgroundRGB(192, 255, 192);
//    }else{
//      dataRange.setBackgroundRGB(255, 192, 192);
//    }
//
//    SpreadsheetApp.flush(); 
//  }
//
//  function onEdit(event)
//  {
//    colorRow(event.source.getActiveRange().getRowIndex());
//  }
