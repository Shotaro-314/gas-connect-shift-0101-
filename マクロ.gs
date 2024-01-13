

function master_footprint_delete() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C12').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('ManagementSheet'), true);
  spreadsheet.getRange('EE3:EE102').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};