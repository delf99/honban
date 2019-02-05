//*************************************************************************
function Manual_SendMail_kakomon(){
  
  // スプレッドシートのシートを取得と準備
  var mySheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("過去問");
    
  // メールセット用各種データをセット
  var strTitle = mySheet.getRange("dz9").getValue(); //送信タイトル
  
  // 各種項目の設定列数を取得
  var aryCol = new Array();      
  aryCol[0] = mySheet.getRange("e3").getValue();　//メアド
  aryCol[1] = mySheet.getRange("e4").getValue();//開始項目
  aryCol[2] = mySheet.getRange("e5").getValue();//終了項目
  
  // 出力項目の確認
  var chkRow = 9;   // 手動送信チェック行
   
  // メール送信
  var outRow = 13; 
  var chkcol = mySheet.getRange("dz1").getColumn();    //※の入力列
  var timecol = mySheet.getRange("ea1").getColumn();    //送信タイムスタンプ列    
  
  //引数(1)起動シート　(2)メアド開始終了配列 (3)送信項目✔列 (4)明細開始行 (5)※列 (6)タイムスタンプ列　(7)件名
  //セ本番(2019判定用).gs シート参照
  sendmailMain(mySheet,aryCol,chkRow,outRow,chkcol,timecol,strTitle);

}

//■過去問指示済・実施済等の結果から、実施・次回候補等の一覧を配列差分を使って取得する
// p1:元値。セル内改行で配列に分割。（日本医科大学2018 char(10) 昭和大学2018　char(10) 東邦大学2017　char(10)…）
// p2:削除値。同上　（東邦大学2018　char(10) 昭和大学2018　char(10)…）
// num:0指定（デフォルト）：全て返す　数字指定：戻り値にセットする差分数（差が多そうな場合）
// 戻り値：p1のみの配列値
//＊num指定ありは、候補校表示に使い、2018、2017など年度部分を消されて呼ばれる想定。2018年をつけて戻す
function myDiff(cel1,cel2,num) {
  var ary1 = cel1.split(String.fromCharCode(10));
  var ary2 = cel2.split(String.fromCharCode(10));  
  var _ = Underscore.load();
  var ary3=_.uniq(_.difference(ary1,ary2));
  var cel3='';
  if(num>0){
    //num指定有は候補抽出を想定（年度なしで呼ばれる想定で。2018を付けて返す）
    if(ary3.length >num){
      return  ary3.slice(0,num).join(String.fromCharCode(10)).replace(/大学/g,'大学2018');
    }else{
      return  ary3.join(String.fromCharCode(10)).replace(/大学/g,'大学2018');
    }
  }else{  
    //num指定なしはシンプルに差分
    return ary3.join(String.fromCharCode(10));
  }  

}



//■〇＋未実施の大学を偏差値上位からPICK
// P1:大学名配列
// P2:〇のついた配列
// num:抽出最大数
// shibou:国のみ、国司併願等（国士併願はセンター2018を最後につける）
function myPick(ary1,ary2,num,shibou){

  var wkary1=[],wkary2=[];
  var ctr1=0,ctr2=0,str='',strctr=0;
  for(var i=0; i<ary1[0].length;i=i+2){
    if(ary2[0][i]=='〇' && ary2[0][i+1]==''){
      wkary1[ctr1] = ary1[0][i]+'2018';
      ctr1++;
    }else if((ary2[0][i]=='〇' && ary2[0][i+1]=='2017')||(ary2[0][i]=='〇' && ary2[0][i+1]=='2018')){  
      if(ary2[0][i]=='〇' && ary2[0][i+1]=='2017'){
        wkary2[ctr2] = ary1[0][i]+'2018';
     
      }else{
        wkary2[ctr2] = ary1[0][i]+'2017';
      }
      ctr2++;        
    }  
  }  

  for(var k=0;k < ctr1 && strctr < num;k++){ 
    str = str + wkary1[k] + String.fromCharCode(10);
    strctr++;
  }
  if(strctr < num){
    for(k=0;k < ctr2 && strctr < num;k++){ 
      str = str + wkary2[k] + String.fromCharCode(10);
      strctr++;
    }
  }  
//    
//  if(shibou == "国私併願"){
//    str = str + "センター2018";
//  }
  
  return str;
}

