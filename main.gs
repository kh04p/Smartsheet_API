function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu("Smartsheet API")
  .addItem("Import CHS data","get_report_chs")
  .addItem("Import effort data","get_sheet_efforts")
  .addItem("Import project data","get_report_projects")
  .addItem("Import intake data","get_sheet_intake")
  .addToUi();
}

function get_report_chs() {
  get_report("000000000000000","Team Report");
}

function get_report_projects() {
  get_report("000000000000000","Projects Report");
}

function get_sheet_efforts() {
  get_sheet("000000000000000","Efforts Report");
}

function get_sheet_intake() {
  get_sheet("000000000000000","Intake Report");
}

function get_report(id,sheet_name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheet_name);
  sheet.clear();

  var uri = "https://api.smartsheet.com/2.0/reports/" + id + "?level=3&include=objectValue&pageSize=1000";

  var header = {
    "Authorization": "Bearer XXXXXXXXXXXXXXXXXX",
    "Content-Type":"application/json"
  }

  var params = {
    "method":"GET",
    "headers":header,
    muteHttpsExceptions:true
  };

  var response = UrlFetchApp.fetch(uri,params);
  var result = JSON.parse(response.getContentText());

  Logger.log(result);

  var output_array = [];

  var cols = result.columns;
  //Logger.log(cols);
  var cols_array = [];
  for (var i = 0; i < cols.length; i++) {
    cols_array.push(cols[i].title);
  }

  output_array.push(cols_array);

  rowloop:
  for (var i = 0; i < result.rows.length; i++) {
    var current_row = i + 1;
    Logger.log("ROW " + current_row);
    var row = result.rows[i].cells;
    Logger.log(row);
    var row_array = [];

    cellloop:
    for (var j = 0; j < row.length; j++) {
      var value = row[j].displayValue;
      if (value == undefined) {
        value = row[j].value;
      }
      row_array.push(value);
    }

    var item_name = row_array[0];
    //Logger.log("Item name: " + item_name);
    var team_owner = row_array[1];
    //Logger.log("Team owner: " + team_owner);
    var item_num = row_array[3];
    //Logger.log("Item #: " + item_num);

    if (team_owner == undefined  && item_num == undefined && item_name.toString().indexOf(">>>CHS Solutioning Phase") == -1 && item_name.toString().indexOf("SA7") == -1 && item_name.toString().indexOf("SA2") == -1) {
      row_array[0] = row_array[11];
    } else if (item_name.toString().indexOf(">>>CHS - Solutioning Phase") != -1 || item_name.toString().indexOf(">>>CHS Solutioning Phase") != -1 || item_name.toString().indexOf("Solutioning - R&D (CHS)") != -1 || item_name.toString().indexOf("SA7") != -1 || item_name.toString().indexOf("SA2") != -1) {
      row_array[0] = row_array[11];  
    }
    
    output_array.push(row_array);
  }

  sheet.getRange(1,1,output_array.length,cols.length).setValues(output_array);
  sheet.moveColumns(sheet.getRange("N1:N"),3);
  sheet.getDataRange().removeDuplicates([1,3,4])

  var today = new Date();
  var today = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss");
  sheet.getRange(1,sheet.getDataRange().getLastColumn() + 1).setValue("Last Refreshed");
  sheet.getRange(2,sheet.getDataRange().getLastColumn(), sheet.getDataRange().getLastRow()-1,1).setValue(today);

  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,sheet.getDataRange().getLastColumn()).setFontWeight('bold');
}

function get_sheet(id,sheet_name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheet_name);
  sheet.clear();

  var uri = "https://api.smartsheet.com/2.0/sheets/" + id;
  var header = {
    "Authorization": "Bearer XXXXXXXXXXXXXXXXXX",
    "Content-Type":"application/json"
  }

  var params = {
    "method":"GET",
    "headers":header,
    muteHttpsExceptions:true
  };

  var response = UrlFetchApp.fetch(uri,params);
  var result = JSON.parse(response.getContentText());

  Logger.log(result);

  var output_array = [];

  var cols = result.columns;
  //Logger.log(cols);
  var cols_array = [];
  for (var i = 0; i < cols.length; i++) {
    Logger.log("Cols length: " + cols.length);
    cols_array.push(cols[i].title);
  }

  output_array.push(cols_array);

  rowloop:
  for (var i = 0; i < result.rows.length; i++) {
    var current_row = i + 1;
    Logger.log("ROW " + current_row);
    Logger.log("Row length: " + result.rows.length);
    var row = result.rows[i].cells;
    Logger.log(row);
    var row_array = [];

    cellloop:
    for (var j = 0; j < row.length; j++) {
      var value = row[j].displayValue;
      if (value == undefined) {
        value = row[j].value;
      }
      row_array.push(value);
    }

    var item_name = row_array[0];
    //Logger.log("Item name: " + item_name);
    var team_owner = row_array[1];
    //Logger.log("Team owner: " + team_owner);
    var item_num = row_array[3];
    //Logger.log("Item #: " + item_num);

    if (team_owner == undefined && item_num == undefined && item_name == undefined) {
      continue rowloop;
    }
    else if (team_owner == undefined  && item_num == undefined && item_name.toString().indexOf(">>>CHS Solutioning Phase") == -1 && item_name.toString().indexOf("SA7") == -1 && item_name.toString().indexOf("SA2") == -1) {
      row_array[0] = row_array[11];
    } else if (item_name.toString().indexOf(">>>CHS - Solutioning Phase") != -1 || item_name.toString().indexOf(">>>CHS Solutioning Phase") != -1 || item_name.toString().indexOf("Solutioning - R&D (CHS)") != -1 || item_name.toString().indexOf("SA7") != -1 || item_name.toString().indexOf("SA2") != -1) {
      row_array[0] = row_array[11];  
    }
    
    output_array.push(row_array);
  }

  sheet.getRange(1,1,output_array.length,cols.length).setValues(output_array);
  sheet.moveColumns(sheet.getRange("N1:N"),3);
  sheet.getDataRange().removeDuplicates([1,2,6]);

  var today = new Date();
  var today = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy HH:mm:ss");
  sheet.getRange(1,sheet.getDataRange().getLastColumn() + 1).setValue("Last Refreshed");
  sheet.getRange(2,sheet.getDataRange().getLastColumn(), sheet.getDataRange().getLastRow() - 1,1).setValue(today);

  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,sheet.getDataRange().getLastColumn()).setFontWeight('bold');
}
