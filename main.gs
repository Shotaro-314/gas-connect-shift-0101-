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

  if("氏名" == sheet.getRange("A2").getValue()){
    name_array = sheet.getRange(name_first_row,name_col, last_row).getValues();
  }
  // Googleフォームのプルダウン内の値を上書きする処理
  let form = FormApp.openById(form_id);

  // 質問項目がプルダウンのもののみ取得
  let items = form.getItems(FormApp.ItemType.LIST);

  items.forEach(function(item){
    // 質問項目が「氏名」を含むものに対して、スプレッドシートの内容を反映する
    if(item.getTitle().match(/氏名.*$/)){
      let listItemQuestion = item.asListItem();
      let choices = [];

      name_array.forEach(function(name){
        if(name != ""){
          choices.push(listItemQuestion.createChoice(name));
        }
      });
      // プルダウンの選択肢を上書きする
      listItemQuestion.setChoices(choices);
    }
  });
  //footprint

  sheet.getRange("AP5").setValue(now_time);

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

