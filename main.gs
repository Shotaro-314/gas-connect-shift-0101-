function pdf_export(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ss_id = ss.getId();
  let sh_id = ss.getSheetByName("確定シフト").getSheetId();

  let sheet = ss.getSheetByName('ManagementSheet');
  let target_name = sheet.getRange("S4").getValue();

  let file_name = ""+target_name+""+now_time+"";
  createPdf(drive_id, ss_id, sh_id, file_name);
}

function master_sync(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ManagementSheet");
  let name_array = null;
  let name_col = 1; //氏名データが入っていた列の指定
  let name_first_row = 3; //　最初の行の指定

  let last_row = sheet.getRange(sheet.getMaxRows(),name_col).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();//指定列の最終行の取得

  last_row = last_row - 2;  //シートの最終行と氏名データの最終行との差を解消

  if("氏名" == sheet.getRange("A2").getValue()){
    name_array = sheet.getRange(name_first_row,name_col, last_row).getValues();
  }

  //配列の長さを取得、行の調整(+1)
  let namearray_length = 1;
  namearray_length = name_array.length　+ 2;
  

  // Googleフォームのプルダウン内の値を上書きする処理
  let form = FormApp.openById(form_id);

  // 質問項目がプルダウンのもののみ取得
  let items = form.getItems(FormApp.ItemType.LIST);

  items.forEach(function(item){
    // 質問項目が氏名を含むものに対して、内容を反映する
    if(item.getTitle().match(/氏名.*$/)){
      let listItemQuestion = item.asListItem();
      let choices = [];

      name_array.forEach(function(name){
        if(name != ""){
          choices.push(listItemQuestion.createChoice(name));
        }
      });
      // プルダウンの選択肢を上書き
      listItemQuestion.setChoices(choices);
    }
  });
  //footprint

  sheet.getRange("AP5").setValue(now_time); //同期時間
  sheet.getRange('EE3:EE'+namearray_length).setValues(name_array);  //最終同期配列の格納

}

function form_sync(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ManagementSheet");

  let title_range = "AG3:AG18";
  let remarks_range = "AH3:AH18";
  let description_range ="AH21"

  let title_array = null;
  let remarks_array = null;
  let target_term = sheet.getRange("S4").getValue();
  let target = null;

  title_array = sheet.getRange(title_range).getValues();
  remarks_array = sheet.getRange(remarks_range).getValues();
  let form_description = sheet.getRange(description_range).getValue();

  // console.log(title_array);
  // console.log(remarks_array);

  let form = FormApp.openById(form_id);
  let form_items = form.getItems();

  target = sheet.getRange("S4").getValues();
  form_items[3].asListItem().setChoiceValues(target).setRequired(true);

  form.setDescription(form_description);
  
  for(let i = 4;i <= 19; i++ ){
    form_items[i].setTitle(title_array[i-4]);
    form_items[i].setHelpText(remarks_array[i-4]);
  }

  //footprint
  sheet.getRange("AG3:AG18").copyTo(sheet.getRange("AL3:AL18"),{contentsOnly:true});
  sheet.getRange("AH3:AH18").copyTo(sheet.getRange("AM3:AM18"),{contentsOnly:true});

  sheet.getRange("AP2").setValue(now_time);
  sheet.getRange("AP3").setValue(target_term);

}

function shift_create_clear(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("シフト作成");

  //注意画面の作成
  let chk_result = Browser.msgBox("作成画面をクリアします。よろしいですか？",Browser.Buttons.OK_CANCEL);
  if (chk_result =="ok"){
    sheet.getRangeList(['F10:G109', 'K10:L109', 'P10:Q109', 'U10:V109', 'Z10:AA109', 'AE10:AF109', 'AJ10:AK109', 'AO10:AP109', 'AT10:AU109', 'AY10:AZ109', 'BD10:BE109', 'BI10:BJ109', 'BN10:BO109', 'BS10:BT109', 'BX10:BY109', 'CC10:CD109']).activate()
    .clear({contentsOnly: true, skipFilteredRows: true});
  }else{
    Browser.msgBox("キャンセルされました。");
  }
}

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

  //人員配置算出欄のコピー
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