var DATATABLE_KEY = 'dt_';

function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('What-If Analysis 2')
    .addItem('Create Data Table', 'create_')
    .addItem('Refresh Data Tables', 'refresh_')
    .addItem('Help', 'help_')
    .addToUi();
  
  // initialize document state for datatables
  PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY) || "{}");
}

function help_() {
  SpreadsheetApp.getUi().alert("Selected range must be at least 2x2: input values on left column (and top row, in the case of a 2D datatable), the model output at the top-left, and table values to the bottom-right.");
}

function create_() {
  var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY));
  var ui = SpreadsheetApp.getUi();
  var dt_range = SpreadsheetApp.getActiveRange();
  var config = false;
  
  if (dt_range.getNumColumns() < 2) {
    help_();
  } else {
    if (dt_range.getNumColumns() > 2) {
      // 2D data-table: row and column inputs
      var result_rowinput = ui.prompt("Specify Model Row Input", 'Specify the row input cell\nFor example, enter "A2" to set cell A2 with the values in the top row.', ui.ButtonSet.OK_CANCEL);
      if(!result_rowinput) return;
      var result_colinput = ui.prompt("Specify Model Column Input", 'Specify the column input cell\nFor example, enter "A4" to set cell A4 with the values in the left column.', ui.ButtonSet.OK_CANCEL);
      if(!result_rowinput) return;
      
      var output2d = dt_range.getCell(1,1);
      var rowinput = SpreadsheetApp.getActiveSpreadsheet().getRange(result_rowinput.getResponseText());
      var colinput = SpreadsheetApp.getActiveSpreadsheet().getRange(result_colinput.getResponseText());
      config = { "range": dt_range.getA1Notation(), "output": output2d.getA1Notation(), "rowinput": rowinput.getA1Notation(), "colinput": colinput.getA1Notation() };
    } else { //numColumn == 2
      // column inputs only
      var result_input = ui.prompt('Specify Model Input', 'Specify the (column) input cell.\nFor example, enter "A2" to set cell A2 with the values in the left column.', ui.ButtonSet.OK_CANCEL);
      if(!result_input) return;

      var input = SpreadsheetApp.getActiveSpreadsheet().getRange(result_input.getResponseText());
      var output = dt_range.getCell(1,2);
      config = { "range": dt_range.getA1Notation(), "output": output.getA1Notation(), "rowinput": null, "colinput": input.getA1Notation() };
    }
    // actually do the work now:
    datatables_(config);
    
    // save named range and property to be able to refresh data
    var key = dt_range.getA1Notation();
    var name = "DataTable_" + key.replace(/[^A-Z0-9]/g,"");
    SpreadsheetApp.getActive().setNamedRange(name, dt_range);
    dt_[key] = config;
    PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, JSON.stringify(dt_));
  }
}

function refresh_() {
  var sheet = SpreadsheetApp.getActive();
  var dt_ = JSON.parse(PropertiesService.getDocumentProperties().getProperty(DATATABLE_KEY));
  var namedRanges = sheet.getNamedRanges();
  var keys = Object.keys(dt_);

  var i, range, name, key, config;
  for (i = 0; i < keys.length; i++) {
    key = keys[i];
    config = dt_[key];
    range = sheet.getRange(config.range);

    namedIndex = namedRanges.indexOf(function(element) {
      return element.getRange().getA1Notation() = this;// "this" is the range
    }, range);

    if(namedIndex > -1) {
      name = namedRanges[namedIndex].getName();
      SpreadsheetApp.getUi().alert("Reevaluating datatable " + name + " at " + key);
      datatables_(config);
    } else {
      delete dt_[key];
      SpreadsheetApp.getUi().alert("Range configuration: " + key + " does not have correponding named range. The configuration has been removed.");
    }
  }
  if(keys.length == 0)
    SpreadsheetApp.getUi().alert("There are no DataTable to refresh!");
  PropertiesService.getDocumentProperties().setProperty(DATATABLE_KEY, JSON.stringify(dt_));
}

function datatables_(config) {
  var s = SpreadsheetApp.getActive();
  var dt_range = s.getRange(config.range);
  if (!config.rowinput) {
    var input = s.getRange(config.colinput);
    var original = input.getValue();
    var output = s.getRange(config.output);
    var numRow = dt_range.getNumRows();
    for (var i = 2; i <= numRow; i++) { 
      input.setValue(dt_range.getCell(i, 1).getValue());
      dt_range.getCell(i, 2).setValue(output.getValue());
    }
    input.setValue(original);
  } else {
    // 2D
    var colinput = s.getRange(config.colinput);
    var rowinput = s.getRange(config.rowinput);
    var colOriginal = colinput.getValue();
    var rowOriginal = rowinput.getValue();
    var output = s.getRange(config.output);
    var numRow = dt_range.getNumRows();
    var numCol = dt_range.getNumColumns();
    for (var i = 2; i <= numRow; i++) {
      for (var j = 2; j <= numCol; j++) {
        colinput.setValue(dt_range.getCell(i, 1).getValue());
        rowinput.setValue(dt_range.getCell(1, j).getValue());
        dt_range.getCell(i, j).setValue(output.getValue());
      }
    }
    colinput.setValue(colOriginal);
    rowinput.setValue(rowOriginal);
  }
}
