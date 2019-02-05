//メール送信練習
function MailPractice() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  var data=sheet.getRange("A13:AH").getValues();
  var data1=sheet.getRange("A13:A").getValues();//チェック欄
  var data2=sheet.getRange("B13:B").getValues();//送信日時
  var subject=sheet.getRange("C5").getValue();
  var body = "";
  var to ="";
  var array=[];
  for(i=0;i<data.length;i++){
    if (data[i][0]=="1"){  
      to=data[i][6];
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
              array.push(i+13);
      
    }else if(data[i][0]=="2"){
      to=data[i][6];
      for(j=0;j<data[0].length;j++){
            if(data[0][j]=="1"){
              body +=data[1][j]+"\n"+data[i][j]+"\n"+"\n";
            }
      }
      GmailApp.sendEmail(
                to,
                subject,
                body,
        {from:"aceacademy@delf.co.jp"}
              );
              data1[i][0]="";
              data2[i][0]=new Date();
              body="";
      
    
    }
  }for(var p in array){
   sheet.hideRows(array[p]);
  }
  
  
  sheet.getRange("A13:A").setValues(data1);
  sheet.getRange("B13:B").setValues(data2);
      
     
    }    

















function Manual_SendMail_shiritsuFB(){
  
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

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

  var rowSheet=mySheet.getDataRange().getLastRow();             // 最終行を取得  
  mySheet.getRange(outRow, aryCol[0],rowSheet-outRow+1,1).copyTo(mySheet.getRange(outRow, aryCol[0],rowSheet-outRow+1,1),{contentsOnly:true});      // メールアドレスべた張り 
  
  //引数(1)起動シート　(2)メアド開始終了配列 (3)送信項目✔列 (4)明細開始行 (5)※列 (6)タイムスタンプ列　(7)件名
  //
  sendmailMain(mySheet,aryCol,chkRow,outRow,chkcol,timecol,strTitle);
  
  mySheet.getRange(outRow, aryCol[0],rowSheet-outRow+1,1).clearContent(); //べたばりしたところクリア  

}

function setsoudanBGcolor() {
 var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("私本番");
 var lastrow=sh.getLastRow();
var lastcol=sh.getLastColumn();
  
  var arr1 =  sh.getRange(14,1,1,lastcol).getValues();

  if(arr1[0].indexOf("相談内容") >-1){
    if(sh.getRange(lastrow, arr1[0].indexOf("相談内容")+1).getValue() != ''){
       sh.getRange(lastrow, arr1[0].indexOf("相談内容")+1).setBackground("powderblue");
    }  
  }  
}

//送信済（B列に値有）の行を非表示にする
function hideRow() {
 var sh=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var lastrow=sh.getLastRow();
  
  var arr1 =  sh.getRange(1,2,lastrow,1).getValues(); //タイムスタンプ列の値を取得

  for(var i=lastrow;i>=16;i--){
    if(sh.isRowHiddenByUser(i)){
      //既に非表示の行はスルー
    }else{  
      if(arr1[i-1][0] != ''){
        sh.hideRow(sh.getRange(i,1));
      }
    }  
  }  
}