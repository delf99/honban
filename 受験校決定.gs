function tip1() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("私立受験校");  // シートを取得
  var data1 = sheet.getRange("AQ13:AQ").getValues();
  var data2 = sheet.getRange("BN13:CR").getValues(); 
  for (i=0;i<data1.length;i++){
    //61未満
    if(parseInt(data1[i][0])<61){
      for(j=0;j<19;j++){
       data2[i][j]="";
      }
    }
    
    //62未満
    if(parseInt(data1[i][0])<62){
      for(j=0;j<14;j++){
       data2[i][j]="";
      }
    }  
      
    //64未満
    if(parseInt(data1[i][0])<64){
      for(j=0;j<6;j++){
       data2[i][j]="";
      }
    } 
      
    //65未満
    if(parseInt(data1[i][0])<65){
      for(j=0;j<3;j++){
       data2[i][j]="";
      }
    }  
    //69未満
    if(parseInt(data1[i][0])<69){
      for(j=0;j<1;j++){
       data2[i][j]="";
      }  
    }    
}sheet.getRange("BN13:CR").setValues(data2); 
}
