function onOpen() {
  let ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
  let menu = ui.createMenu('HappyShift');        // Uiクラスからメニューを作成する
  menu.addItem('収集期間更新','form_sync');
  menu.addItem('従業員マスタ同期','master_sync');
  menu.addItem('PDFエクスポート','pdf_export');
  menu.addItem('シフト作成画面クリア','shift_create_clear');
  //menu.addItem('hide',ms_sheet_hide);
  //menu.addItem('unhide',ms_sheet_unhide);
  menu.addToUi();                            // メニューをUiクラスに追加する
}