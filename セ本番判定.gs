function hantei(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("セ本番判定");
  var lastrow=ss.getDataRange().getLastRow();             // 最終行を取得  
  var startrow = 16;
  var arrGrade = ss.getRange("D1:D"+lastrow).getValues();
  var arrSend = ss.getRange("Az1:Az"+lastrow).getValues(); //送信日列（空白以外は上書きしないように）
  
  var colA = ss.getRange("bc1").getColumn(); 
  var colB = ss.getRange("bd1").getColumn();   
  var colC = ss.getRange("be1").getColumn();   
  var colD = ss.getRange("bf1").getColumn();   
  var colE = ss.getRange("bg1").getColumn();    

  var colA2 = ss.getRange("bi1").getColumn(); 
  var colB2 = ss.getRange("bj1").getColumn();   
  var colC2 = ss.getRange("bk1").getColumn();   
  var colD2 = ss.getRange("bl1").getColumn();   
  var colE2 = ss.getRange("bm1").getColumn();     

  var colmae = ss.getRange("Bn1").getColumn(); 
  var colato = ss.getRange("Bo1").getColumn();   
  
  var formularow = 15;
  
  ss.getRange(16,colA,300,5).clear();
  ss.getRange(16,colA2,300,5).clear();
  ss.getRange(16,colmae,300,2).clear();
  
  ss.getRange("BB9").setValue(new Date());
  
  for(var i=startrow;i<=lastrow;i++){
    // 受験生以外と送信済の行はなにもしない
    if(arrGrade[i-1][0] == '' || arrGrade[i-1][0] == '高1' || arrGrade[i-1][0] == '高2' || arrGrade[i-1][0] == '中3' || arrSend[i-1][0] != ''){
    }else{
      ss.getRange(formularow,colA).copyTo(ss.getRange(i,colA), SpreadsheetApp.CopyPasteType.PASTE_FORMULA); 
      ss.getRange(formularow,colB).copyTo(ss.getRange(i,colB), SpreadsheetApp.CopyPasteType.PASTE_FORMULA); 
      ss.getRange(formularow,colC).copyTo(ss.getRange(i,colC), SpreadsheetApp.CopyPasteType.PASTE_FORMULA); 
      ss.getRange(formularow,colD).copyTo(ss.getRange(i,colD), SpreadsheetApp.CopyPasteType.PASTE_FORMULA); 
      ss.getRange(formularow,colE).copyTo(ss.getRange(i,colE), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);       
      
      ss.getRange(formularow,colA2).copyTo(ss.getRange(i,colA2), SpreadsheetApp.CopyPasteType.PASTE_FORMULA); 
      ss.getRange(formularow,colB2).copyTo(ss.getRange(i,colB2), SpreadsheetApp.CopyPasteType.PASTE_FORMULA); 
      ss.getRange(formularow,colC2).copyTo(ss.getRange(i,colC2), SpreadsheetApp.CopyPasteType.PASTE_FORMULA); 
      ss.getRange(formularow,colD2).copyTo(ss.getRange(i,colD2), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);       
      ss.getRange(formularow,colE2).copyTo(ss.getRange(i,colE2), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);         
      
      ss.getRange(formularow,colmae).copyTo(ss.getRange(i,colmae), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);               
      ss.getRange(formularow,colato).copyTo(ss.getRange(i,colato), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);                     
    }  
  }  
  
  
  ss.getRange(startrow,colA,lastrow-startrow+1,5).setValues(ss.getRange(startrow,colA,lastrow-startrow+1,5).getValues());
  ss.getRange(startrow,colA2,lastrow-startrow+1,5).setValues(ss.getRange(startrow,colA2,lastrow-startrow+1,5).getValues());  
  ss.getRange(startrow,colmae,lastrow-startrow+1,2).setValues(ss.getRange(startrow,colmae,lastrow-startrow+1,2).getValues());    
   Browser.msgBox('owari');
}


/*************************************************************************/
/* 手動メール送信
/*************************************************************************/
function Manual_SendMail_center_hantei(){
  
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // メールセット用各種データをセット
  var strTitle = mySheet.getRange("c5").getValue(); //送信タイトル
  
  // 各種項目の設定列数を取得
  var aryCol = new Array();      
  aryCol[0] = mySheet.getRange("e3").getValue();　//メアド
  aryCol[1] = mySheet.getRange("e4").getValue();//開始項目
  aryCol[2] = mySheet.getRange("e5").getValue();//終了項目
  
  // 出力項目の確認
  var chkRow = 12;   // 手動送信チェック行
   
  // メール送信
  var outRow = 16; 
  var chkcol = mySheet.getRange("c3").getValue();    //※の入力列
  var timecol = mySheet.getRange("c4").getValue();    //送信タイムスタンプ列    

//  mySheet.getRange("C12:F310").copyTo(mySheet.getRange("C12:F310"),{contentsOnly:true});      // メールアドレス等取れないところをべた張り
 
  //引数(1)起動シート　(2)メアド開始終了配列 (3)送信項目✔列 (4)明細開始行 (5)※列 (6)タイムスタンプ列　(7)件名
  sendmailMain(mySheet,aryCol,chkRow,outRow,chkcol,timecol,strTitle);

//  mySheet.getRange("C12:F310").clearContent(); //べたばりしたところクリア
  
}

