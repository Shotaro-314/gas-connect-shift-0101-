function createPdf(folderId, ssId, shId, fileName){
  //PDFを作成するためのベースとなるURL
  let baseUrl = "https://docs.google.com/spreadsheets/d/"
          +  ssId
          + "/export?gid="
          + shId;
 
  //PDFのオプションを指定
  let pdfOptions = "&exportFormat=pdf&format=pdf"
              + "&size=A4" //用紙サイズ (A4)
              + "&portrait=false"  //用紙の向き true: 縦向き / false: 横向き
              + "&fitw=true"  //ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
              + "&top_margin=0.50" //上の余白
              + "&right_margin=0.50" //右の余白
              + "&bottom_margin=0.50" //下の余白
              + "&left_margin=0.50" //左の余白
              + "&horizontal_alignment=CENTER" //水平方向の位置
              + "&vertical_alignment=CENTER" //垂直方向の位置
              + "&printtitle=false" //スプレッドシート名の表示有無
              + "&sheetnames=false" //シート名の表示有無
              + "&gridlines=true" //グリッドラインの表示有無
              + "&fzr=true" //固定行の表示有無
              + "&fzc=false" //固定列の表示有無;

  //PDFを作成するためのURL
  let url = baseUrl + pdfOptions;

  //アクセストークンを取得する
  let token = ScriptApp.getOAuthToken();

  //headersにアクセストークンを格納する
  let options = {
    headers: {
        'Authorization': 'Bearer ' +  token
    }
  };
 
  //PDFを作成する
  let blob = UrlFetchApp.fetch(url, options).getBlob().setName(fileName + '.pdf');

  //PDFの保存先フォルダー
  //フォルダーIDは引数のfolderIdを使用します
  let folder = DriveApp.getFolderById(folderId);

  //PDFを指定したフォルダに保存する
  folder.createFile(blob);
}