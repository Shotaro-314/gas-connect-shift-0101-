function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('HappyShift');
  menu.addItem('収集期間更新','form_sync');
  menu.addItem('従業員マスタ同期','master_sync');
  menu.addItem('PDFエクスポート','pdf_export');
  menu.addItem('シフト作成画面クリア','shift_create_clear');
  menu.addSeparator();
  menu.addSubMenu(
      ui.createMenu("開発用")
      .addItem("copy_and_sync", "copy_and_sync")
      .addItem("calc-hide","calc_col_hide")
      .addItem("draft","draft")
  );
  //menu.addItem('hide',ms_sheet_hide);
  //menu.addItem('unhide',ms_sheet_unhide);
  menu.addToUi();
}