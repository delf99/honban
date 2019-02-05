//メール送信
function MailGouhi() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  var data=sheet.getRange("A9:N").getValues();
  var data1=sheet.getRange("A9:A").getValues();//チェック欄
  var data2=sheet.getRange("B9:B").getValues();//送信日時
  var subject=sheet.getRange("C5").getValue();
  var body = "";
  var to ="";
  for(i=0;i<data.length;i++){
    if (data[i][0]=="1"){  
      to=data[i][5];
       for(j=0;j<data[0].length;j++){
            if(data[0][j]=="1"){
              body +=data[1][j]+"\n"+data[i][j]+"\n"+"\n";
           }
           }GmailApp.sendEmail(
       to,
       subject,
       body,
        {from:"aceacademy@delf.co.jp"}
       );
      
      data1[i][0]="";
      data2[i][0]=new Date();
      body="";

    }
  }
  sheet.getRange("A9:A").setValues(data1);
  sheet.getRange("B9:B").setValues(data2);
      
     
    }    


/*************************************************************************/
// 合否フォーム送信用スクリプト
/*************************************************************************/
function Manual_SendMail_gouhiFB(){
  
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("合否");
  var lastrow = mySheet.getLastRow();
  // 各種項目の設定列・行を取得
  var aryCol = new Array();      
  aryCol[0] = mySheet.getRange("e3").getValue();　//メアド
  aryCol[1] = mySheet.getRange("e4").getValue();//開始項目
  aryCol[2] = mySheet.getRange("e5").getValue();//終了項目

  var chkcol = mySheet.getRange("e6").getValue();    //※の入力列
  var timecol = mySheet.getRange("e7").getValue();    //送信タイムスタンプ列    
  var chkRow = mySheet.getRange("C3").getValue();   // 手動送信チェック行
  var outRow = mySheet.getRange("C4").getValue();; 　  // 送信対象開始行
  var strTitle = mySheet.getRange("c5").getValue(); //送信タイトル  
　//---  
  
  mySheet.getRange(12, 6,lastrow-11,1).copyTo(mySheet.getRange(12, 6,lastrow-11,1),{contentsOnly:true});      // メールアドレス等取れないところをべた張り 
  
  //引数(1)起動シート　(2)メアド開始終了配列 (3)送信項目✔列 (4)明細開始行 (5)※列 (6)タイムスタンプ列　(7)件名
  //
  sendmailMain(mySheet,aryCol,chkRow,outRow,chkcol,timecol,strTitle);
  
  mySheet.getRange(12, 6,lastrow-11,1).clearContent(); //べたばりしたところクリア  

}

//送信済（B列に値有）で不合格の行を非表示にする
function hideRowFugoukaku() {
 var sh=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var lastrow=sh.getLastRow();
  
  var arr1 =  sh.getRange(1,2,lastrow,1).getValues(); //タイムスタンプ列の値を取得
  var arr2 =  sh.getRange(1,8,lastrow,1).getValues(); //合否結果列の値を取得

  for(var i=lastrow;i>=12;i--){
    if(sh.isRowHiddenByUser(i)){
      //既に非表示の行はスルー
    }else{  
      if(arr1[i-1][0] != '' && arr2[i-1][0].match(/不合格/)){
        sh.hideRow(sh.getRange(i,1));
      }
    }  
  }  
}