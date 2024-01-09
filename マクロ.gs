function copy_and_sync() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('シフト修正'), true);
  
  //すべてコピー(値のみ)
  spreadsheet.getRange('A10').activate();
  spreadsheet.getRange('\'シフト作成\'!A10:CF109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  //算出列のコピー元生成(関数のコピー)
  spreadsheet.getRange('E111').activate();
  spreadsheet.getRange('\'シフト作成\'!E10:CF10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  //算出列のコピー
  spreadsheet.getRange('H10').activate();
  spreadsheet.getRange('H111:I111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('M10').activate();
  spreadsheet.getRange('M111:N111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('R10').activate();
  spreadsheet.getRange('R111:S111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('W10').activate();
  spreadsheet.getRange('W111:X111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AB10').activate();
  spreadsheet.getRange('AB111:AC111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AG10').activate();
  spreadsheet.getRange('AG111:AH111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AL10').activate();
  spreadsheet.getRange('AL111:AM111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AQ10').activate();
  spreadsheet.getRange('AQ111:AR111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AV10').activate();
  spreadsheet.getRange('AV111:AW111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BA10').activate();
  spreadsheet.getRange('BA111:BB111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BF10').activate();
  spreadsheet.getRange('BF111:BG111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BK10').activate();
  spreadsheet.getRange('BK111:BL111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BP10').activate();
  spreadsheet.getRange('BP111:BQ111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BU10').activate();
  spreadsheet.getRange('BU111:BV111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BZ10').activate();
  spreadsheet.getRange('BZ111:CA111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('CE10').activate();
  spreadsheet.getRange('CE111:CF111').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // 値のみコピーした不要なデータを削除
  spreadsheet.getRange('H11:I109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('M11:N109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('R11:S109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('W11:X109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('AB11:AC109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('AG11:AH109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('AL11:AM109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('AQ11:AR109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('AV11:AW109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('BA11:BB109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('BF11:BG109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('BK11:BL109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('BP11:BQ109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('BU11:BV109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('BZ11:CA109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('CE11:CF109').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
  //日付同期
  spreadsheet.getRange('A2').activate();
  spreadsheet.getRange('\'シフト作成\'!2:2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  // 人員配置予定をコピー
  spreadsheet.getRange('E4').activate();
  spreadsheet.getRange('\'シフト作成\'!E4:E5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('J4').activate();
  spreadsheet.getRange('\'シフト作成\'!J4:J5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('O4').activate();
  spreadsheet.getRange('\'シフト作成\'!O4:O5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('T4').activate();
  spreadsheet.getRange('\'シフト作成\'!T4:T5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('Y4').activate();
  spreadsheet.getRange('\'シフト作成\'!Y4:Y5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AD4').activate();
  spreadsheet.getRange('\'シフト作成\'!AD4:AD5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AI4').activate();
  spreadsheet.getRange('\'シフト作成\'!AI4:AI5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AN4').activate();
  spreadsheet.getRange('\'シフト作成\'!AN4:AN5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AS4').activate();
  spreadsheet.getRange('\'シフト作成\'!AS4:AS5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AX4').activate();
  spreadsheet.getRange('\'シフト作成\'!AX4:AX5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BC4').activate();
  spreadsheet.getRange('\'シフト作成\'!BC4:BC5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BH4').activate();
  spreadsheet.getRange('\'シフト作成\'!BH4:BH5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BM4').activate();
  spreadsheet.getRange('\'シフト作成\'!BM4:BM5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BR4').activate();
  spreadsheet.getRange('\'シフト作成\'!BR4:BR5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BW4').activate();
  spreadsheet.getRange('\'シフト作成\'!BW4:BW5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('CB4').activate();
  spreadsheet.getRange('\'シフト作成\'!CB4:CB5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};


function calc_col_hide() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H:I').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('H2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('M:N').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('M2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('R:S').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('R2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
 
  spreadsheet.getRange('W:X').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('W2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('AB:AC').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AB2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('AG:AH').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AG2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('AL:AM').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AL2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('AQ:AR').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AQ2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('AV:AW').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AV2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('BA:BB').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('BA2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('BF:BG').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('BF2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('BK:BL').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('BK2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('BP:BQ').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('BP2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('BU:BV').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('BU2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('BZ:CA').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('BZ2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  spreadsheet.getRange('CE:CF').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('CE2'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

};

function draft() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('シフト修正'), true);

  spreadsheet.getRange('I10').activate();
  spreadsheet.getRange('\'シフト作成\'!I10:I109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('N10').activate();
  spreadsheet.getRange('\'シフト作成\'!N10:N109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('S10').activate();
  spreadsheet.getRange('\'シフト作成\'!S10:S109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('X10').activate();
  spreadsheet.getRange('\'シフト作成\'!X10:X109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('AC10').activate();
  spreadsheet.getRange('\'シフト作成\'!AC10:AC109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('AH10').activate();
  spreadsheet.getRange('\'シフト作成\'!AH10:AH109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('AM10').activate();
  spreadsheet.getRange('\'シフト作成\'!AM10:AM109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('AR10').activate();
  spreadsheet.getRange('\'シフト作成\'!AR10:AR109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('AW10').activate();
  spreadsheet.getRange('\'シフト作成\'!AW10:AW109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('BB10').activate();
  spreadsheet.getRange('\'シフト作成\'!BB10:BB109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('BG10').activate();
  spreadsheet.getRange('\'シフト作成\'!BG10:BG109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('BL10').activate();
  spreadsheet.getRange('\'シフト作成\'!BL10:BL109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('BQ10').activate();
  spreadsheet.getRange('\'シフト作成\'!BQ10:BQ109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('BV10').activate();
  spreadsheet.getRange('\'シフト作成\'!BV10:BV109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('CA10').activate();
  spreadsheet.getRange('\'シフト作成\'!CA10:CA109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('CF10').activate();
  spreadsheet.getRange('\'シフト作成\'!CF10:CF109').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};