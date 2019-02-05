function myFunction() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("私受験校（練習用）");
  var data1 = sheet.getRange("AQ12:AQ").getValues();//偏差値
  var data2 = sheet.getRange("BN12:CV").getValues();//受験校
  
  for(i=0;i<data1.length;i++){
    
    //60未満→65以上をカット
    if(data1[i][0]<60){
      for(j=0;j<19;j++){
       data2[i][j]="";
      }
      }
     //61未満→66以上をカット
   if(data1[i][0]<61){
      for(j=0;j<14;j++){
       data2[i][j]="";
      }
      }
   　
    //62未満→66以上をカット
    if(data1[i][0]<62){
      for(j=0;j<8;j++){
       data2[i][j]="";
     }
     }
   　　
     //63未満→67以上をカット
   if(data1[i][0]<63){
      for(j=0;j<6;j++){
       data2[i][j]="";
     } 
     }
    
  }
  sheet.getRange("BN12:CV").setValues(data2);
  


}
