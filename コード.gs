
function trim2() {
  const ss =SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName("sheet1");
  const sheet2 = ss.getSheetByName("Trim");
  let lastRow = sheet1.getLastRow();
  let data = sheet1.getRange(2,1,lastRow,5).getValues();
  for(let i = 1 ;i<=lastRow;i++){
    if(String(data[i-1][4]).match("・解約希望理由：")){
      data[i-1][4] = Parser.data(data[i-1][4]).from("・解約希望理由：").to("\n").build();
    }else{
      data[i-1][0] = "";
      data[i-1][1] = "";
      data[i-1][2] = "";
      data[i-1][3] = "";
      data[i-1][4] = "";
    }
  }
  sheet2.getRange(2,1,lastRow,5).setValues(data);
  let sortrange = sheet2.getRange(2,1,lastRow,5)
  sortrange.sort({column: 2, ascending: true})
}
