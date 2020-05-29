var epicLinkColumn = "P";
var epicsSheetName = "Epics";

var tmpSheetName = "JiraImportTMP";
var finalSheetName = "JiraImport";

var sprintsSheetName = "Sprints";



function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Import Data",
    functionName : "importJira"
  }];
  sheet.addMenu("Jira", entries);
};

function sheetClose(name) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var tmpSheet = activeSpreadsheet.getSheetByName(name);

    if (tmpSheet != null) {
        activeSpreadsheet.deleteSheet(tmpSheet);
    }
}

function sheetOpen(name) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var tmpSheet = activeSpreadsheet.getSheetByName(name);
  
  if (tmpSheet != null) {
    return tmpSheet;
  }
  
  tmpSheet = activeSpreadsheet.insertSheet();
  tmpSheet.setName(name);

  return tmpSheet;
}

function reopenSheet(name) {
  sheetClose(name)
  return sheetOpen(name);
}

function insertTasks(tmpSheet) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sprintsSheet = activeSpreadsheet.getSheetByName(sprintsSheetName);
  
  var i = 2;
  var cell = tmpSheet.getRange(tmpSheet.getLastRow()+1, 1);
  cell.setFormula("QUERY(JIRA(\"project = CRM AND resolution = Done AND resolved >= \" & TEXT(Sprints!A"+i+",\"YYYY-MM-DD\") & \" AND resolved <= \" & TEXT(Sprints!B"+i+",\"YYYY-MM-DD\") & \" ORDER BY lastViewed DESC\"), \"select '\" & TEXT(Sprints!A"+i+",\"YYYY-MM-DD\") & \"','\" & TEXT(Sprints!B"+i+",\"YYYY-MM-DD\") & \"',Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9,Col10,Col11,Col12,Col13,Col14,Col15,Col16,Col17 LABEL '\" & TEXT(Sprints!A"+i+",\"YYYY-MM-DD\") & \"' 'SprintStart','\" & TEXT(Sprints!B"+i+",\"YYYY-MM-DD\") & \"' 'SprintEnd'\", -1)");
  
  for(var i = 3; i <= sprintsSheet.getLastRow(); i++) {
    var cell = tmpSheet.getRange(tmpSheet.getLastRow()+1, 1);
    cell.setFormula("QUERY(QUERY(JIRA(\"project = CRM AND resolution = Done AND resolved >= \" & TEXT(Sprints!A"+i+",\"YYYY-MM-DD\") & \" AND resolved <= \" & TEXT(Sprints!B"+i+",\"YYYY-MM-DD\") & \" ORDER BY lastViewed DESC\"), \"select '\" & TEXT(Sprints!A"+i+",\"YYYY-MM-DD\") & \"','\" & TEXT(Sprints!B"+i+",\"YYYY-MM-DD\") & \"',Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9,Col10,Col11,Col12,Col13,Col14,Col15,Col16,Col17\", -1),\"select * offset 1\", 0)");
  }
}

function getEpicsData() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(epicsSheetName);
  var range = sheet.getRange(2,1,sheet.getLastRow()-1, sheet.getLastColumn());
  var rangeValues = range.getDisplayValues();
  return rangeValues;
}

function cloneGoogleSheet(sourceSheet, targetSheet) {
  var sourceRange = sourceSheet.getDataRange();
  var sourceA1Range = sourceRange.getA1Notation();
  var sourceData = sourceRange.getValues();


  targetSheet.clear({contentsOnly: true});
  targetSheet.getRange(sourceA1Range).setValues(sourceData);
};

function replaceEpics(sheet) {
  var epicsData = getEpicsData()
  
  epicsData.forEach(function(row){
    var oldText = row[0];
    var newText = row[1];
    var textFinder = sheet.getRange(epicLinkColumn + "2:" + epicLinkColumn + sheet.getLastRow()).createTextFinder(oldText);
    textFinder.matchCase(false);
    textFinder.matchEntireCell(true);
    textFinder.replaceAllWith("["+oldText+"] "+newText);
  });
}

function importJira() {
  var tmpSheet = reopenSheet(tmpSheetName);
  var finalSheet = sheetOpen(finalSheetName);
  
  insertTasks(tmpSheet);
  cloneGoogleSheet(tmpSheet, finalSheet);
  
  replaceEpics(finalSheet);
  
  sheetClose(tmpSheetName);
}

