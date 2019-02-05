/* メニュー */
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menus = [{name: 'マーク模試マージ', functionName: 'tip1'},
               {name: '記述模試マージ', functionName: 'tip2'},
               {name: '模試マージ(増田)', functionName: 'tip_work'}
               
              ];

  ss.addMenu('メニュー', menus);
}

/*************************************************************************/
/* 手動メール送信　2016/08/24 作成 */
/*************************************************************************/
function Manual_SendMail(){
 
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSheet();                  // シートを取得
  var rowSheet=mySheet.getDataRange().getLastRow();             // 最終行を取得
  var sheetName=mySheet.getName();　　　　　　　　　　　　　　　// シート名を取得
  
  if(sheetName=="記述模試"){
  
    // 各種データをセット
    var strFrom=mySheet.getRange("送信元").getValue(); 
    var strSubject=mySheet.getRange("手動タイトル").getValue();
    
    // 各種項目の設定列数を取得
    var aryCol = new Array();      
    aryCol[0] = mySheet.getRange("生徒番号").getValue();
    aryCol[1] = mySheet.getRange("氏名").getValue();
    aryCol[2] = mySheet.getRange("メールアドレス").getValue();
    aryCol[3] = mySheet.getRange("開始項目").getValue();
    aryCol[4] = mySheet.getRange("終了項目").getValue();
    
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
        mySheet.getRange(i,3).setValue(new Date())
        mySheet.getRange(i,2).setValue("")
        
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


//模試返却後マージ
function tip2(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("記述模試");  // シートを取得
  var _ = Underscore.load();
  var arr1 =  sheet.getRange("D12:AG257").getValues();//[2018/02/14 21:37:53,28,藤沢,通塾..],[2018/02/14 21:37:53,31,佐野,.],[2018/02/14 21:37:53,50..],
  var arr2 = _.zip.apply(_,arr1);//[（タイムスタンプ）],[28,31,50,53,...],[藤沢,佐野,生田,..],[通塾,..],[..,]
  var lastRow = sheet.getLastRow();
  var arr3 = sheet.getRange("D258:AG").getValues();//(返却後のデータ）[2018/02/14 21:37:53,28,藤沢,通塾..],[2018/02/14 21:37:53,31,佐野,.],[2018/02/14 21:37:53,50..],
  
  for (i=0;i<arr3.length;i++){
    var stu =arr3[i][1]; //生徒番号
    if(arr3[i][6]=="自己採点"){continue;}
    if (arr2[1].indexOf(stu)==-1){ sheet.getRange(i+258,29).setValue("データなし"); continue;} //getRangeのi+ には返却後の開始行をいれる
    var arrno =arr2[1].indexOf(stu); //生徒番号の配列の番号（indexOf：最初に見つけた要素のインデックス番号）
    var tip =arr3[i][8]-arr1[arrno][8];//自己採点と返却後の差 
    arr1[arrno]=arr3[i];  //返却後にデータ入れ替え
    arr1[arrno][26]=tip;  //自己採点との差を入力
  }
  sheet.getRange("D12:AG257").setValues(arr1);  //全データセット
  sheet.getRange("G12:H").clearContent(); //arrayformulaとのバッティング解除
  sheet.getRange("K12:M").clearContent(); 
  sheet.getRange("AB12:AB").clearContent(); //mail
}

//返却読み取り
function tip3(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var score = sheet.getRange("J12:J").getValues();
  for (i=0;i<score.length;i++){
    if (score[i] == '返却後'){
     
    }
  }


}


//模試返却後マージ
function tip_work(){
  
  var _ = Underscore.load();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetNm = sheet.getSheetName();
  if(sheetNm != "記述模試" && sheetNm != "マーク模試"){return;}
 
  var ans = Browser.inputBox("マージ元となる自己採点の最終行を入力してください（この行の模試名を対象に返却後マージします）",Browser.Buttons.OK);
  
  if(ans > 12){
    var wkrow = Number(ans)+1;
    if(sheetNm=="記述模試"){
      var rng1 = "D12:AG"+ans;
      var rng2 = "D"+wkrow+":AG";
      var col1 = 30;   //AD列（自己採点の列）
      var col2 = 26;
    }else{ 
      //マーク模試
      var rng1 = "D12:AS"+ans;
      var rng2 = "D"+wkrow+":AS";
      var col1 = 42;      //AP列（自己採点の列）
      var col2 = 38;      
    }      
  }else{  return;}
  
  var targetTest=sheet.getRange("I"+ans).getValue();
  
  var arr1 =  sheet.getRange(rng1).getValues();//[2018/02/14 21:37:53,28,藤沢,通塾..],[2018/02/14 21:37:53,31,佐野,.],[2018/02/14 21:37:53,50..],
  var arr2 = _.zip.apply(_,arr1);//[（タイムスタンプ）],[28,31,50,53,...],[藤沢,佐野,生田,..],[通塾,..],[..,]
  var lastRow = sheet.getLastRow();
  var arr3 = sheet.getRange(rng2).getValues();//(返却後のデータ）[2018/02/14 21:37:53,28,藤沢,通塾..],[2018/02/14 21:37:53,31,佐野,.],[2018/02/14 21:37:53,50..],
  var j,arrno;
  for (var i=0;i<arr3.length;i++){
    var stu =arr3[i][1]; //生徒番号
    if(arr3[i][6]=="自己採点" || arr3[i][5]!=targetTest ){continue;}

    if (arr2[1].indexOf(stu)==-1){ 
      sheet.getRange(i+wkrow,col1).setValue("データなし"); 
      continue;
    } //getRangeのi+ には返却後の開始行をいれる
    
    j=0;
    while(arr2[1].indexOf(stu,j) != -1){
      arrno =arr2[1].indexOf(stu,j); //生徒番号の配列の番号（indexOf：最初に見つけた要素のインデックス番号）
      if(arr1[arrno][5] !=targetTest){
        j= arrno+1;
      }else{
        var tip =arr3[i][8]-arr1[arrno][8];//自己採点と返却後の差 
        arr1[arrno]=arr3[i];  //返却後にデータ入れ替え
        arr1[arrno][col2]=tip;  //自己採点との差を入力
        j=arr2[1].length;
      }  
    }    
  }
  
  sheet.getRange(rng1).setValues(arr1);  //全データセット
  sheet.getRange("G12:H").clearContent(); //arrayformulaとのバッティング解除
  sheet.getRange("K12:M").clearContent(); 
  
  if(sheetNm=="記述模試"){
    sheet.getRange("AB12:AB").clearContent(); //mail
  }else{
    sheet.getRange("Z12:Z").clearContent();     //国語
    sheet.getRange("AO12:AO").clearContent();   //mail
  }    
  
}




