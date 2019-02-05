// カンマ区切り大学名文字列の差分取得。
// cel1にあってcel2にない大学名を取得する
//   cel1:○大学,△大学,×大学,★大学
//   cel2:△大学,×大学,□大学,
//   return:○大学,★大学
function mydiff2(cel1,cel2) {
  var ary1 = cel1.split(",");
  var ary2 = cel2.split(",");  
  var _ = Underscore.load();
//  var ary3=_.difference(_.sortBy(ary1),_.sortBy(ary2));
  var ary3=_.uniq(_.difference(ary1,ary2));
  return ary3.join(",");
}

/*************************************************************************/
// メール送信(私受験校＞最終報告確認）　送信用スクリプト
/*************************************************************************/
function Manual_SendMail_saisyukakunin(){
  
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 各種項目の設定列・行を取得
  var aryCol = new Array();      
  aryCol[0] = mySheet.getRange("dk3").getValue();　//メアド
  aryCol[1] = mySheet.getRange("dk4").getValue();//開始項目
  aryCol[2] = mySheet.getRange("dk5").getValue();//終了項目

  var chkcol = mySheet.getRange("dk6").getValue();    //※の入力列
  var timecol = mySheet.getRange("dk7").getValue();    //送信タイムスタンプ列    
  var chkRow = mySheet.getRange("di3").getValue();   // 手動送信チェック行
  var outRow = mySheet.getRange("di4").getValue();; 　  // 送信対象開始行
  var strTitle = mySheet.getRange("di5").getValue(); //送信タイトル  
　//---  
  
  //引数(1)起動シート　(2)メアド開始終了配列 (3)送信項目※列 (4)明細開始行 (5)※列 (6)タイムスタンプ列　(7)件名
  //
  sendmailMain2(mySheet,aryCol,chkRow,outRow,chkcol,timecol,strTitle);

}
///*************************************************************************/
///* 私受験校用　推奨校送信スクリプトはメール送信.gsの汎用を使用
///*************************************************************************/
//function Manual_SendMail_jukenko_shiritsu(){
//  
//  // スプレッドシートのシートを取得と準備
//  var mySheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("私受験校");
//  
//  // メールセット用各種データをセット
//  var strTitle = mySheet.getRange("dd9").getValue(); //送信タイトル
//  
//  // 各種項目の設定列数を取得
//  var aryCol = new Array();      
//  aryCol[0] = mySheet.getRange("e3").getValue();　//メアド
//  aryCol[1] = mySheet.getRange("e4").getValue();//開始項目
//  aryCol[2] = mySheet.getRange("e5").getValue();//終了項目
//  
//  // 出力項目の確認
//  var chkRow = 9;   // 手動送信チェック行
//   
//  // メール送信
//  var outRow = 12; 
//  var chkcol = mySheet.getRange("DD1").getColumn();    //※の入力列
//  var timecol = mySheet.getRange("de1").getColumn();    //送信タイムスタンプ列    
//  
//  //引数(1)起動シート　(2)メアド開始終了配列 (3)送信項目✔列 (4)明細開始行 (5)※列 (6)タイムスタンプ列　(7)件名
//  sendmailMain(mySheet,aryCol,chkRow,outRow,chkcol,timecol,strTitle);
//
//}