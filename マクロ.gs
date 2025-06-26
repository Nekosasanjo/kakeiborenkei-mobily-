function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2:A3').activate()
  .breakApart();
  spreadsheet.getRange('A5:A6').activate()
  .breakApart();
  spreadsheet.getRange('B6:D6').activate()
  .breakApart();
  spreadsheet.getRange('E5').activate();
  spreadsheet.getRange('B6').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('B3:D3').activate()
  .breakApart();
  spreadsheet.getRange('E2').activate();
  spreadsheet.getRange('B3').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('3:3').activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('5:5').activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('A15').activate();
};