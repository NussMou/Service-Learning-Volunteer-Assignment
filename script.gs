function createSheetsBasedOnFirstChoice() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = spreadsheet.getSheetByName("Sheet1");
  var data = mainSheet.getDataRange().getValues();
  var firstChoiceColumn, mail_col, id_col, name_col;

  // 尋找列索引
  for (var i = 0; i < data[0].length; i++) {
    var column = data[0][i];
    if (column == "志願一") firstChoiceColumn = i;
    else if (column == "Email") mail_col = i;
    else if (column == "學號") id_col = i;
    else if (column == "姓名") name_col = i;
  }

  // 獲取所有現有的工作表名稱，並將它們存儲在一個對象中
  var existingSheets = spreadsheet.getSheets().reduce(function(acc, sheet) {
    acc[sheet.getName()] = true;
    return acc;
  }, {});

  // 從第二行開始迭代（假設第一行是標頭）
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var firstChoice = row[firstChoiceColumn];
    if (firstChoice && !existingSheets[firstChoice]) {
      // 創建新的試算表
      var newSheet = spreadsheet.insertSheet(firstChoice);
      newSheet.appendRow(["姓名", "Email", "學號", "理由/連結/推薦"]);
      existingSheets[firstChoice] = true;
    }
    // 第一志願已經有對應的工作表
    var targetSheet = spreadsheet.getSheetByName(firstChoice);
    var studentInfo = [row[name_col], row[mail_col], row[id_col], row[firstChoiceColumn + 1]];
    targetSheet.appendRow(studentInfo);
    
  }
}
