/*************************************************************************/
/* メール送信（通信生課題フォーム返信用（シート名に拠らず起動））
/*************************************************************************/
function Manual_SendMail_tsushinkadai(){
  
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSheet();                  // シートを取得
  var rowSheet=mySheet.getDataRange().getLastRow();             // 最終行を取得
  var sheetName=mySheet.getName();　　　　　　　　　　　　　　　// シート名を取得
  
  // 各種データをセット
  var strFrom=mySheet.getRange("B3").getValue(); //FROMメアド
  var strSubject=mySheet.getRange("B6").getValue(); //メール件名
  
  // 各種項目の設定列数を取得
  var aryCol = new Array();      
  aryCol[0] = mySheet.getRange("M2").getValue(); //生徒番号
  aryCol[1] = mySheet.getRange("M3").getValue();     //生徒氏名
  aryCol[2] = mySheet.getRange("M4").getValue(); //生徒Mail
  aryCol[3] = mySheet.getRange("M5").getValue();      //メール出力項目FROM
  aryCol[4] = mySheet.getRange("M6").getValue();      //メール出力項目to
  
  // 出力項目の確認
  var aryCheck = new Array();
  var aryField = new Array();
  var chkRow = 9;   // 手動送信チェック行
  var c = 0;        // カウント用
  for(var i = aryCol[3]; i <= aryCol[4]; i++){    
    aryCheck[c]=mySheet.getRange(chkRow, i).getValue();        // 出力確認
    aryField[c]=mySheet.getRange(chkRow + 1, i).getValue();    // 項目名
    c = c + 1;
  }
  
  // メール送信
  var outRow = 11;   // データ開始行
  for(var i = outRow; i <= rowSheet; i++){
    
    var chkSend=mySheet.getRange(i,2).getValue();              // 送信チェック      
    if(chkSend == "※") {        
      
      // 本文データを取得
      var strNo=mySheet.getRange(i,aryCol[0]).getValue();      // 生徒番号
      var strName=mySheet.getRange(i,aryCol[1]).getValue();    // 氏名
      var strTo=mySheet.getRange(i,aryCol[2]).getValue();      // メールアドレス
      var aryBody = new Array();                               // 総評　等
      var c = 0;
      for(var j = aryCol[3]; j <= aryCol[4]; j++){    
        aryBody[c]=mySheet.getRange(i,j).getValue();
        c = c + 1;
      }
      
      // 本文作成
      var attachmentFiles = new Array();
      var strBody= "";

      for(var k = 0; k <= aryCol[4] - aryCol[3]; k++){    
    
        if(aryCheck[k] == "✔︎") {
          if(aryField[k] == "添付1") {
            if(aryBody[k] != "") {
              var attachment1 = DriveApp.getFileById(aryBody[k]).getBlob();
              attachmentFiles.push({fileName:attachment1.getName(), mimeType: attachment1.getContentType(), content:attachment1.getBytes()});
            }
          } else if(aryField[k] == "添付2") {
            if(aryBody[k] != "") {
              var attachment2 = DriveApp.getFileById(aryBody[k]).getBlob();
              attachmentFiles.push({fileName:attachment2.getName(), mimeType: attachment2.getContentType(), content:attachment2.getBytes()});
            }
          } else if(aryField[k] == "添付3") {
            if(aryBody[k] != "") {
              var attachment3 = DriveApp.getFileById(aryBody[k]).getBlob();
              attachmentFiles.push({fileName:attachment3.getName(), mimeType: attachment3.getContentType(), content:attachment3.getBytes()});
            }
          } else {
            strBody = strBody + "【" + aryField[k] + "】\n" + " " + aryBody[k] + "\n\n"
          }
        }
      }
      
      mySheet.getRange(i,3).setValue(new Date());
      mySheet.getRange(i,2).setValue("");
      
      //送信前のスリープ 1sec
      Utilities.sleep(1000);
      
      // メールを送信（添付ファイルがある場合とない場合で処理分け）
      if (attachmentFiles.length > 0) {
        MailApp.sendEmail(strTo, strSubject, strBody, {attachments:attachmentFiles});
      } else {
        MailApp.sendEmail(strTo, strSubject, strBody);
      }
      
      // 配列を初期化
      attachmentFiles.length = 0;
      
    }
  }      
  
}
