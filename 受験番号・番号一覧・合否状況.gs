//受験番号シートに送信された受験番号が重複していないか（同じ大学＋番号が重複していないか（同じ人が複数回送ってきた場合と番号間違いを想定）
function double_chk() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh1 = ss.getActiveSheet();
  var lastrow = sh1.getLastRow();
  var ary1 = sh1.getRange(11,sh1.getRange("B1").getColumn(),lastrow-11+1,1).getValues();
  var ary_stu = ary1.concat(ary1).concat(ary1).concat(ary1).concat(ary1);

  sh1.getRange('E:Z').setNumberFormat('@');
  
  var ary2_1_1 = sh1.getRange(11,sh1.getRange("E1").getColumn(),lastrow-11+1,1).getValues();  
  var ary2_1_2 = sh1.getRange(11,sh1.getRange("I1").getColumn(),lastrow-11+1,1).getValues();    
  var ary2_1_3 = sh1.getRange(11,sh1.getRange("M1").getColumn(),lastrow-11+1,1).getValues();      
  var ary2_1_4 = sh1.getRange(11,sh1.getRange("Q1").getColumn(),lastrow-11+1,1).getValues();        
  var ary2_1_5 = sh1.getRange(11,sh1.getRange("U1").getColumn(),lastrow-11+1,1).getValues();          
  var ary_sch = ary2_1_1.concat(ary2_1_2).concat(ary2_1_3).concat(ary2_1_4).concat(ary2_1_5);

  var ary2_2_1 = sh1.getRange(11,sh1.getRange("H1").getColumn(),lastrow-11+1,1).getValues();  
  var ary2_2_2 = sh1.getRange(11,sh1.getRange("L1").getColumn(),lastrow-11+1,1).getValues();    
  var ary2_2_3 = sh1.getRange(11,sh1.getRange("P1").getColumn(),lastrow-11+1,1).getValues();      
  var ary2_2_4 = sh1.getRange(11,sh1.getRange("T1").getColumn(),lastrow-11+1,1).getValues();        
  var ary2_2_5 = sh1.getRange(11,sh1.getRange("X1").getColumn(),lastrow-11+1,1).getValues();          
  var ary_no = ary2_2_1.concat(ary2_2_2).concat(ary2_2_3).concat(ary2_2_4).concat(ary2_2_5);
  
  var ary_wk=[];
  for(var i=0;i<ary_stu.length;i++){
    if(ary_no[i][0] != ""){
      ary_wk.push(ary_sch[i][0] +"-"+ ary_no[i][0]);
    }
  }  
  
  var ary_rtn =
      ary_wk.filter(function (val, idx, arr){
        return arr.indexOf(val) === idx && idx !== arr.lastIndexOf(val);
  });
  
  if(ary_rtn==''){
    Browser.msgBox('重複なし');
  }else{
    Browser.msgBox('次の大学-受験番号が重複\\n'+ary_rtn);
  }  
}

//受験番号シートの番号を、合否状況シートに転記する
function bango_copy() {
  
  var ans = Browser.msgBox("AE列が空欄の行の番号を転記します");
  if(ans != "ok"){return;}
     
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh1 = ss.getActiveSheet();
  var lastrow = sh1.getLastRow();
  sh1.getRange('E:Z').setNumberFormat('@');
  
  var str,j;
  var col_E =sh1.getRange("E1").getColumn();
  var col_X =sh1.getRange("X1").getColumn();  
  var ary_chk = sh1.getRange(11,sh1.getRange("AE1").getColumn(),lastrow-11+1,1).getValues();
  var ary_no = sh1.getRange(11,sh1.getRange("B1").getColumn(),lastrow-11+1,1).getValues();
  var ary_bango =sh1.getRange(11,col_E,lastrow-11+1,sh1.getRange("X1").getColumn()-col_E+1).getValues();
  
  var sh2 = ss.getSheetByName("合否状況");
  var ary_target = sh2.getRange(7,1,1,sh2.getLastColumn()).getValues(); //合否状況シートの大学名＋試験区分（検索用）
  
  for(var i=0;i<lastrow-11+1;i++){
    if(ary_chk[i][0] == '' && ary_no[i][0] != ''){
      j=0;
      //配列5セット分の背景色をクリア
      sh1.getRange(i+11,sh1.getRange("E1").getColumn(),1,col_X-col_E+1).setBackground("white");
      
      //受験番号5セット分検索＆転記(ただし、すでに記入ある場合はSKIPする）
      while(j<20){
        
        if(ary_bango[i][j] != ''){
          if(ary_bango[i][j+1]=='その他' || ary_bango[i][j+2] !=''){
            //その他の試験方式、またはその他に記入があるものはSKIP(帝京、東海の一般は除く（どの日かのメモ）

            str = ary_bango[i][j]+ary_bango[i][j+1];  // 合否状況シートの大学名+試験方式の列位置を検索
            if(ary_bango[i][j] =='帝京大学' && ary_bango[i][j+1] =='一般前期' && ary_target[0].indexOf(str) > -1){
                // 帝京一般前期:番号の3桁目＝1→1日目の位置,2→2日目の位置,3→3列目にシフト 1410114
                if((ary_bango[i][j+3].substr(2,1) == '1' || ary_bango[i][j+3].substr(2,1) == '2' || ary_bango[i][j+3].substr(2,1) == '3' ) &&
                    sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+Number(ary_bango[i][j+3].substr(2,1))).getValue()==''){
                    sh1.getRange(11+i,col_E+j+3,1,1).copyTo(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+Number(ary_bango[i][j+3].substr(2,1))),{contentsOnly:true});
                    sh1.getRange(i+11,col_E+j,1,4).setBackground("silver");     //転記したら背景色をグレイ            
                    }  
            }else if(ary_bango[i][j] =='東海大学'  && ary_bango[i][j+1] =='一般前期' && ary_target[0].indexOf(str) > -1){
                //東海一般前期：頭2桁01=1日目の位置、それ以外=2日目の位置(列位置シフト）
              if(ary_bango[i][j+3].substr(0,2) == '01'){
                if(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+1).getValue()==''){
                    sh1.getRange(11+i,col_E+j+3,1,1).copyTo(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+1),{contentsOnly:true});
                    sh1.getRange(i+11,col_E+j,1,4).setBackground("silver");     //転記したら背景色をグレイ            
                }
              }else{
                if(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+2).getValue()==''){
                    sh1.getRange(11+i,col_E+j+3,1,1).copyTo(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+2),{contentsOnly:true});
                    sh1.getRange(i+11,col_E+j,1,4).setBackground("silver");     //転記したら背景色をグレイ            
                }                
              }  
            }
          }else{
            str = ary_bango[i][j]+ary_bango[i][j+1];  // 合否状況シートの大学名+試験方式の列位置を検索
            if(ary_target[0].indexOf(str) > -1){

              if(ary_bango[i][j] =='帝京大学' && ary_bango[i][j+1] =='一般前期'){
                
                // 帝京一般前期:番号の3桁目＝1→1日目の位置,2→2日目の位置,3→3列目にシフト 1410114
                if((ary_bango[i][j+3].substr(2,1) == '1' || ary_bango[i][j+3].substr(2,1) == '2' || ary_bango[i][j+3].substr(2,1) == '3' ) &&
                    sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+Number(ary_bango[i][j+3].substr(2,1))).getValue()==''){
                      sh1.getRange(11+i,col_E+j+3,1,1).copyTo(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+Number(ary_bango[i][j+3].substr(2,1))),{contentsOnly:true});
                      sh1.getRange(i+11,col_E+j,1,4).setBackground("silver");     
                    }  
              }else if(ary_bango[i][j] =='東海大学'  && ary_bango[i][j+1] =='一般前期'){
                //東海一般前期：頭2桁01=1日目の位置、それ以外=2日目の位置(列位置シフト）
                if(ary_bango[i][j+3].substr(0,2) == '01'){
                  if(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+1).getValue()==''){
                    sh1.getRange(11+i,col_E+j+3,1,1).copyTo(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+1),{contentsOnly:true});
                    sh1.getRange(i+11,col_E+j,1,4).setBackground("silver");     //転記したら背景色をグレイ            
                  }
                }else{
                  if(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+2).getValue()==''){
                    sh1.getRange(11+i,col_E+j+3,1,1).copyTo(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+2),{contentsOnly:true});
                    sh1.getRange(i+11,col_E+j,1,4).setBackground("silver");     //転記したら背景色をグレイ            
                  }                
                }  
              }else{  
                if(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+1).getValue() ==''){ 
                  //該当の大学の列位置に番号転記
                  sh1.getRange(11+i,col_E+j+3,1,1).copyTo(sh2.getRange(ary_no[i][0]+10,ary_target[0].indexOf(str)+1),{contentsOnly:true});
                  sh1.getRange(i+11,col_E+j,1,4).setBackground("silver");                       
                }  
                
              }  
            } 
          }
        
        }else{
          //空欄のセットは背景色をグレイ
          sh1.getRange(i+11,col_E+j,1,4).setBackground("silver");          
        }
        j=j+4; //次の番号セット
      }
      ary_chk[i][0] = "済"; 
    }
    
  } 
  
  sh1.getRange(11,sh1.getRange("AE1").getColumn(),lastrow-11+1,1).setValues(ary_chk);
}

//番号一覧シートに記入された合否を、合否状況シートに転記する
function gouhi_copy() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh1 = ss.getSheetByName("番号一覧");
  var lastrow = sh1.getLastRow();  
  var ary_no = sh1.getRange(4,sh1.getRange("j1").getColumn(),lastrow-4+1,1).getValues();     //受験番号  
  var ary_stuno = sh1.getRange(4,sh1.getRange("g1").getColumn(),lastrow-4+1,1).getValues();
  var ary_col = sh1.getRange(4,sh1.getRange("k1").getColumn(),lastrow-4+1,1).getValues();    //合否状況シートの転記列位置
  var ary_str = sh1.getRange(4,sh1.getRange("n1").getColumn(),lastrow-4+1,1).getValues();    //転記文字
  var ary_done = sh1.getRange(4,sh1.getRange("o1").getColumn(),lastrow-4+1,1).getValues();    //※  
　
  var sh2 = ss.getSheetByName("合否状況");
  var ary_target = sh2.getRange(11,1,sh2.getLastRow(),sh2.getLastColumn()).getValues();
  
  var j;
  //番号一覧の未済の行分、合否列に値があれば転記。×の場合は、O列を済にして以降反映しない。
  for(var i=0;i<ary_done.length;i++){
    if(ary_done[i][0] =='' && ary_str[i][0] != '' && ary_col[i][0] != ''){
      sh2.getRange(ary_stuno[i][0]+10,ary_col[i][0]).setValue(ary_str[i][0]);
      if(ary_str[i][0]=='×'){
        ary_done[i][0]='済';
      }  
    }  
  }
  sh1.getRange(4,sh1.getRange("o1").getColumn(),lastrow-4+1,1).setValues(ary_done);
}


/*************************************************************************/
// メール送信(合否状況用：行ごとにメール件名変更可能　サブルーチンはメール送信gs参照）
/*************************************************************************/
function Manual_SendMail_gouhijoukyou(){
  
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
  
  var shTitle = mySheet.getRange("c5").getValue(); //件名（セルが空白の場合使用）  
  
  //引数(1)起動シート　(2)メアド+開始列+終了列+件名+※+送信日 (3)送信項目✔列 (4)明細開始行 (5)代表件名
  //
  sendmailMain3(mySheet,aryCol,chkRow,outRow,shTitle);

}

