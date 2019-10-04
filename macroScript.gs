/** @OnlyCurrentDoc */

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C2').activate()
  .setFormula('=MID(A2&" "&A2,FIND(" ",A2)+1,LEN(A2))');
  spreadsheet.getActiveRangeList().clearFormat();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C1:C2'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  var destinationRange = spreadsheet.getActiveRange().offset(0, 0, 234);
  spreadsheet.getActiveRange().autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getCurrentCell().offset(232, 0).activate();
  spreadsheet.getCurrentCell().offset(0, -2, 2, 1).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn() - 2, sheet.getMaxRows(), 1).activate();
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn() + 2, sheet.getMaxRows(), 1).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn() + 2, sheet.getMaxRows(), 1).activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow(), 1, 1, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getCurrentCell().activate();
  spreadsheet.getCurrentCell().setValue('Prezime i Ime');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('Broj Stola');
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(0, -1, 1, 2).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRangeList().setFontSize(12)
  .setFontWeight('bold');
  spreadsheet.getActiveSheet().setColumnWidth(1, 152);
  currentCell = spreadsheet.getCurrentCell().offset(208, -1);
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn() - 1, sheet.getMaxRows(), 1).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveSheet().setFrozenRows(0);
  currentCell = spreadsheet.getCurrentCell();
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn(), sheet.getMaxRows(), 2).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: spreadsheet.getActiveRange().getColumn(), ascending: true});
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn() + 2, sheet.getMaxRows(), 1).activate();
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn() - 2, sheet.getMaxRows(), 2).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn() + 2, sheet.getMaxRows(), 1).activate();
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, spreadsheet.getCurrentCell().getColumn() - 4, sheet.getMaxRows(), 2).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(1, -2).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getCurrentCell().offset(52, -2, 1, 2).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(1, 2).activate();
  spreadsheet.getCurrentCell().offset(52, -2, 183, 2).moveTo(spreadsheet.getActiveRange());
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(1, 1, 1, 2).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(1, 2).activate();
  spreadsheet.getCurrentCell().offset(52, -2, 131, 2).moveTo(spreadsheet.getActiveRange());
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(1, 1).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getCurrentCell().offset(0, 4, 79, 2).moveTo(spreadsheet.getActiveRange());
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getCurrentCell().offset(-52, 0).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getCurrentCell().offset(106, 0, 1, 2).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(0, 2).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getCurrentCell().offset(53, -2, 26, 2).moveTo(spreadsheet.getActiveRange());
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(1, 1).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().offset(0, 0, 1, 2).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(1, 2).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getCurrentCell().offset(53, -2, 895, 2).moveTo(spreadsheet.getActiveRange());
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(27, -3, 1, 6).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getCurrentCell().offset(0, 2, 1, 2).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getCurrentCell().offset(-14, 0).activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getActiveRange().getDataRegion().activate();
  currentCell.activateAsCurrentCell();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getActiveRange()])
  .whenCellNotEmpty()
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getActiveRange()])
  .whenFormulaSatisfied('=ISEVEN()')
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getActiveRange()])
  .whenFormulaSatisfied('=ISEVEN(ROW())')
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getActiveRange()])
  .whenFormulaSatisfied('=ISEVEN(ROW())')
  .setBackground('#CFE2F3')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  spreadsheet.getCurrentCell().offset(-4, -1).activate();
};

function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D20').activate();
  spreadsheet.getActiveSheet().setFrozenRows(0);
  spreadsheet.getRange('B:B').activate();
  spreadsheet.getActiveSheet().moveColumns(spreadsheet.getRange('B:B'), 1);
  spreadsheet.getRange('C2').activate()
  .setFormula('=MID(A2&" "&A2,FIND(" ",A2)+1,LEN(A2))');
  spreadsheet.getActiveRangeList().clearFormat();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C1:C2'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C2:C450'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C2:C450').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getRange('A:A').activate();
  spreadsheet.getRange('C:C').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('C:C').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue('Prezime i Ime');
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('Broj Stola');
  spreadsheet.getRange('A1:B1').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B1'));
  spreadsheet.getActiveRangeList().setFontSize(12)
  .setFontWeight('bold');
  spreadsheet.getRange('A2').activate();
  spreadsheet.getActiveSheet().setColumnWidth(1, 147);
  spreadsheet.getActiveSheet().autoResizeColumns(1, 1);
  spreadsheet.getActiveSheet().setColumnWidth(1, 156);
  spreadsheet.getActiveSheet().setColumnWidth(1, 146);
  spreadsheet.getActiveSheet().setColumnWidth(2, 117);
  spreadsheet.getRange('A:B').activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 1, ascending: true});
  spreadsheet.getRange('C:C').activate();
  spreadsheet.getRange('A:B').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('E:E').activate();
  spreadsheet.getRange('A:B').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('C2').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getRange('A1').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('A56:B56').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getRange('C2').activate();
  spreadsheet.getRange('A56:B451').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('C56:D56').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getRange('E2').activate();
  spreadsheet.getRange('C56:D397').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('D2').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('E56').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('A56').activate();
  spreadsheet.getRange('E56:F343').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('A55').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('A106:B106').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('A112').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('C56').activate();
  spreadsheet.getRange('A112:B343').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('B56').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('C112:D112').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getRange('E1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('E56').activate();
  spreadsheet.getRange('C112:D287').moveTo(spreadsheet.getActiveRange());
  spreadsheet.getRange('D56').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('C67').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('A111').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('A111:F111').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('C111:D111').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('A111').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.PREVIOUS).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('A111:F111').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).activate();
  currentCell.activateAsCurrentCell();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:F111')])
  .whenCellNotEmpty()
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:F111')])
  .whenFormulaSatisfied('=ISEVEN()')
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:F111')])
  .whenFormulaSatisfied('=ISEVEN(ROW())')
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:F111')])
  .whenFormulaSatisfied('=ISEVEN(ROW())')
  .setBackground('#CFE2F3')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  spreadsheet.getRange('E112').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
};
