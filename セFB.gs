/*************************************************************************/
// メール送信(合否状況用：行ごとにメール件名変更可能　サブルーチンはメール送信gs参照）
/*************************************************************************/
function Manual_SendMail_centerFB(){
  
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 各種項目の設定列・行を取得
  var aryCol = new Array();      
  aryCol[0] = mySheet.getRange("e3").getValue();　//メアド
  aryCol[1] = 1;                      //開始項目 (シートの列すべて）
  aryCol[2] = mySheet.getLastColumn();//終了項目 (シートの列すべて）…配列で回すのでそんな遅くならないと思う
  aryCol[3] = mySheet.getRange("e4").getValue();  //メール件名の入力列
  aryCol[4] = mySheet.getRange("e5").getValue();    //※の入力列
  aryCol[5] = mySheet.getRange("e6").getValue();   //送信タイムスタンプ列    

  var chkRow = mySheet.getRange("C3").getValue();   // 手動送信チェック行
  var outRow = mySheet.getRange("C4").getValue();; 　  // 送信対象開始行
  
  var shTitle = mySheet.getRange("c5").getValue(); //件名（タイトルセルが空白の場合使用）  


  mySheet.getRange("C12:F310").copyTo(mySheet.getRange("C12:F310"),{contentsOnly:true});      // メールアドレス等取れないところをべた張り

  
  //引数(1)起動シート　(2)メアド+開始列+終了列+件名+※+送信日 (3)送信項目✔列 (4)明細開始行 (5)代表件名
  //
  sendmailMain3(mySheet,aryCol,chkRow,outRow,shTitle);

  
  mySheet.getRange("C12:F310").clearContent(); //べたばりしたところクリア
  
}

