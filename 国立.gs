function display(){
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート8");
  var data1=sheet.getRange("T3:T").getValues();//判定の行
  var name=sheet.getRange("B3:B").getValues();//大学名
  var column_A="";
  var column_B="";
  var column_C="";
  var column_D="";
  var column_E="";
  for(i=0;i<data1.length;i++){
  if (data1[i][0]=="A"){
  column_A+=name[i][0]+"/";}
  
  }
  
   sheet.getRange("U2").setValue(column_A);
}



