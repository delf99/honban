/*************************************************************************/
/* 手動メール送信　
/*************************************************************************/
function Manual_SendMail_kakomon1(){
 
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSheet();                  // シートを取得
  var rowSheet=mySheet.getDataRange().getLastRow();             // 最終行を取得
  var sheetName=mySheet.getName();　　　　　　　　　　　　　　　// シート名を取得
  
  if(sheetName=="私立過去問"||sheetName=="ｾﾝﾀｰ過去問"){
  
    // 各種データをセット
    var strFrom=mySheet.getRange("B3").getValue(); 
    var strSubject=mySheet.getRange("B6").getValue();
    
    // 各種項目の設定列数を取得
    var aryCol = new Array();      
    aryCol[0] = mySheet.getRange("J2").getValue();
    aryCol[1] = mySheet.getRange("J3").getValue();
    aryCol[2] = mySheet.getRange("J4").getValue();
    aryCol[3] = mySheet.getRange("J5").getValue();
    aryCol[4] = mySheet.getRange("J6").getValue();
    
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
        var strBody= "";   //本文の最初（""内）を好きに変えてOK（生徒番号と名前を表示したい時は、"生徒番号" + strNo + "　" + strName + " 様\n\n"を入力）
        for(var k = 0; k <= aryCol[4] - aryCol[3]; k++){    
          if(aryCheck[k] == "✔︎") {
            if(aryField[k] == "文頭") {
              strBody = strBody + aryBody[k] + "\n\n"
              break;
            }
          }
        }
        
        for(var k = 0; k <= aryCol[4] - aryCol[3]; k++){    
          if(aryCheck[k] == "✔︎") {
            if(aryField[k] == "文頭") {
              // 何もしない
            } else {
              strBody = strBody + "【" + aryField[k] + "】\n" + " " + aryBody[k] + "\n\n"
            }
          }
        }

        //送信前のスリープ 1sec
        Utilities.sleep(1000);
        
        MailApp.sendEmail(strTo, strSubject, strBody);
        
        mySheet.getRange(i,3).setValue(new Date())
        mySheet.getRange(i,2).setValue("")
        
      }
    }      
  }
}
