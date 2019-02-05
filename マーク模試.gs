/*************************************************************************/
/* 手動メール送信　2016/08/24 作成 */
/*************************************************************************/
function Manual_SendMail4(){
 
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSheet();                  // シートを取得
  var rowSheet=mySheet.getDataRange().getLastRow();             // 最終行を取得
  
  var sheetName=mySheet.getName();　　　　　　　　　　　　　　　// シート名を取得
  
  if(sheetName=="マーク模試"){

    // 各種データをセット
    var strFrom=mySheet.getRange("送信元4").getValue(); 
    var strSubject=mySheet.getRange("手動タイトル4").getValue();
    
    // 各種項目の設定列数を取得
    var aryCol = new Array();      
    aryCol[0] = mySheet.getRange("生徒番号4").getValue();
    aryCol[1] = mySheet.getRange("氏名4").getValue();
    aryCol[2] = mySheet.getRange("メールアドレス4").getValue();
    aryCol[3] = mySheet.getRange("開始項目4").getValue();
    aryCol[4] = mySheet.getRange("終了項目4").getValue();
    
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
            if(aryField[k] == "文頭") {
              strBody = strBody + aryBody[k] + "\n\n"
              break;
            }
          }
        }
        
        for(var k = 0; k <= aryCol[4] - aryCol[3]; k++){    
          //Browser.msgBox("aryCheck:" + aryCheck[k]); 
          //Browser.msgBox("aryBody:" + aryBody[k]);
          if(aryCheck[k] == "✔︎") {
            if(aryField[k] == "文頭") {
              // 何もしない
            } else if(aryField[k] == "添付1") {
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
        
        //var lastCol=mySheet.getLastColumn()    // 送信完了日の列数（C列）
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
        
        //ドキュメントの内容をログに表示
        //Logger.log(strBody);
      }
    }      
  }
}

//返却後2名の（正しい）点数-偏差値の範囲を選択（2行2列）
//→偏差値の理論値を求め、外れる偏差値で報告してきた人は色を付ける（N5、P6に数式が入っていること）
//±1の範囲のずれはOK（Q5セルに入力）
function hensachiCheck() {
  var spl = SpreadsheetApp.getActiveSpreadsheet();
  var　sh = spl.getActiveSheet();
  // 点1　偏差値1
  // 点2　偏差値2
  var inp1=sh.getActiveRange().getValues();
  var wk1 = sh.getRange("N5").getFormula();

  var moshiNm = sh.getRange("q3").getValue();
  
  if (sh.getRange("N5").getFormula()!="=MINVERSE(N3:O4)" ||sh.getRange("P5").getFormula()!="=MMULT(N5:O6, P3:P4)"){
    return;
  }  
  try{
    sh.getRange("N3").setValue(inp1[0][0]); //点数1  
    sh.getRange("p3").setValue(inp1[0][1]); //偏差値1  
    sh.getRange("N4").setValue(inp1[1][0]); //点数2  
    sh.getRange("p4").setValue(inp1[1][1]); //偏差値2
    //ねんのためちょっとねる 
    Utilities.sleep(100);
    //Q5,Q6に逆行列MMULT→行列の積(MMULT) で求めた解を四捨五入した値
    var a = sh.getRange("P5").getValue();
    var b = sh.getRange("P6").getValue();
    var gosa = sh.getRange("Q5").getValue();  //前後でずれててもOKの範囲
    
    var ans;
    var col = sh.getActiveRange().getColumn();
    var lastrow = sh.getLastRow();
    
    var arrpt = sh.getRange(12,col,lastrow-12+1,1).getValues();
    var arrhen = sh.getRange(12,col+1,lastrow-12+1,1).getValues();
    var arrsaiten = sh.getRange(12,10,lastrow-12+1,1).getValues();
    var arrmoshiNm = sh.getRange(12,9,lastrow-12+1,1).getValues();
    
    for(var i=0;i<arrpt.length;i++){
      
      //返却後で理論値±OK誤差以外の偏差値は赤背景
      if(arrpt[i]!='' && (arrhen[i] < arrpt[i]*a+b-gosa || arrhen[i] > arrpt[i]*a+b+gosa) && arrsaiten[i]=='返却後'&& arrmoshiNm[i]==moshiNm){
        sh.getRange(12+i,col+1,1,1).setBackground("lightpink");
        sh.getRange(12+i,col+1,1,1).setNote(Math.round((arrpt[i]*a+b)*100)/100);
      }else{
        sh.getRange(12+i,col+1,1,1).setBackground("white");         
      }   
    }  
  }catch(e){
    Browser.msgBox("選択範囲を確認してください。処理中断");
    return;
  } 
}

//行数追加時修正個所：①arr1の範囲(to)　②arr3の範囲(from) ③"データなし"setvalueの範囲(+176⇒arr3 のfromの行数）　④arr1のsetValue(全データセット)の範囲
//
//模試返却後マージ
function tip1(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("マーク模試");  // シートを取得
  var _ = Underscore.load();
  var arr1 =  sheet.getRange("D12:AS166").getValues();//[2018/02/14 21:37:53,28,藤沢,通塾..],[2018/02/14 21:37:53,31,佐野,.],[2018/02/14 21:37:53,50..],
  var arr2 = _.zip.apply(_,arr1);//[（タイムスタンプ）],[28,31,50,53,...],[藤沢,佐野,生田,..],[通塾,..],[..,]
  var lastRow = sheet.getLastRow();
  var arr3 = sheet.getRange("D167:AS").getValues();//(返却後のデータ）[2018/02/14 21:37:53,28,藤沢,通塾..],[2018/02/14 21:37:53,31,佐野,.],[2018/02/14 21:37:53,50..],
  
  for (i=0;i<arr3.length;i++){
    var stu =arr3[i][1]; //生徒番号
    if (arr2[1].indexOf(stu)==-1){ sheet.getRange(i+167,41).setValue("データなし"); continue;}
    var arrno =arr2[1].indexOf(stu); //生徒番号の配列の番号（indexOf：最初に見つけた要素のインデックス番号）
    var tip =arr3[i][8]-arr1[arrno][8];//自己採点と返却後の差 
    arr1[arrno]=arr3[i];  //返却後にデータ入れ替え
    arr1[arrno][38]=tip;  //自己採点との差を入力
  }
  sheet.getRange("D12:AS166").setValues(arr1);  //全データセット
  sheet.getRange("G12:H").clearContent(); //arrayformulaとのバッティング解除
  sheet.getRange("K12:M").clearContent();     //得点率～英数理得点率
  sheet.getRange("Z12:Z").clearContent();     //国語
  sheet.getRange("AO12:AO").clearContent();   //mail
}





