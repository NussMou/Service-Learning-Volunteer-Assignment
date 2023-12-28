var firstChoiceColumn, secondChoiceColumn, thirdChoiceColumn;
var mail_col, id_col, name_col;
var spreadsheet, mainSheet, bufferSheet, inputSheet;
var mainData;

function createSheetsBasedOnFirstChoice() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  mainSheet = spreadsheet.getSheetByName("Sheet1");
  mainData = mainSheet.getDataRange().getValues();

  // 尋找列索引
  for (var i = 0; i < mainData[0].length; i++) {
    var column = mainData[0][i];
    if (column == "志願一") firstChoiceColumn = i;
    else if (column == "志願二") secondChoiceColumn = i;
    else if (column == "志願三") thirdChoiceColumn = i;
    else if (column == "Email") mail_col = i;
    else if (column == "學號") id_col = i;
    else if (column == "姓名") name_col = i;
  }

  // 獲取所有現有的工作表名稱，並將它們存在一個 object 中
  var existingSheets = spreadsheet.getSheets().reduce(function(acc, sheet) {
    acc[sheet.getName()] = true;
    return acc;
  }, {});

  if(!existingSheets["Buffer"]){
    bufferSheet = spreadsheet.insertSheet("Buffer");
    bufferSheet.appendRow(["margin"]);
    let l = [];
    for (var i = 0; i < mainData[0].length; i++) {
      l.push(mainData[0][i]);
    }
    bufferSheet.appendRow(l);
  }
  else {
    bufferSheet = spreadsheet.getSheetByName("Buffer");
  }
  if(!existingSheets["Input"]){
    inputSheet = spreadsheet.insertSheet("Input");
    inputSheet.appendRow(["margin"]);
  }
  else {
    inputSheet = spreadsheet.getSheetByName("Input");
  }

  // 從第二行開始迭代（假設第一行是標頭）
  for (var i = 1; i < mainData.length; i++) {
    var row = mainData[i];
    var firstChoice = row[firstChoiceColumn];
    if (firstChoice && !existingSheets[firstChoice]) {
      // 創建新的試算表
      var newSheet = spreadsheet.insertSheet(firstChoice);
      newSheet.appendRow(["姓名", "Email", "學號", "志願二", "志願三", "理由/連結/推薦"]);
      existingSheets[firstChoice] = true;
      appendToColumn(inputSheet, 1, firstChoice);
    }
    // 第一志願已經有對應的工作表
    var targetSheet = spreadsheet.getSheetByName(firstChoice);
    var studentInfo = [row[name_col], row[mail_col], row[id_col],row[secondChoiceColumn], row[thirdChoiceColumn], row[firstChoiceColumn + 1]];
    targetSheet.appendRow(studentInfo);
  }
}

function appendToColumn(sheet, rowIndex, value) {
  var lastColumn = sheet.getLastColumn();
  sheet.getRange(rowIndex, lastColumn + 1).setValue(value);
}

// Buffer function 如果不要這個學生，就在自己的col填入學生名字
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var editedValue = e.value;
  var name_col;
  var bufferSheet = e.source.getSheetByName('Buffer'); // 獲取 "Buffer" 工作表
  var mainSheet = e.source.getSheetByName("Sheet1");
  var mainData = mainSheet.getDataRange().getValues();
  // bufferSheet.appendRow(["test",mainData[0].length]);


  for (var i = 0; i < mainData[0].length; i++) {
    var column = mainData[0][i];
    if (column == "姓名") name_col = i;
  }

  // 確保編輯是在 "Input" 工作表上進行的
  if (sheet.getName() !== 'Input') {
    //  if(sheet.getName() !== 'Buffer')
    //   bufferSheet.appendRow(["!Input","test"]);
    if (sheet.getName() == "Sheet1" || sheet.getName() == "Buffer") return; 
    else {
      
    }
    return;
  }
  // else if(!editedValue) return ;
  var tmp_values = range.getValues(); // 這會是一個二維數組
  editedValue = tmp_values[0][0]; // 獲取單個編輯值

  // 檢查主工作表中的每一行
  // bufferSheet.appendRow(["Input","test",editedValue]);
  // 檢查主工作表中的每一行
  for (var i = 0; i < mainData.length; i++) {
    if (mainData[i][name_col] == editedValue) {
      // 找到匹配的學生，將其信息複製到 Buffer 工作表
      bufferSheet.appendRow(mainData[i]);
      return; // 匹配後結束
    }
  }
}


// 這個function用來分割sheet，並把sheet寄給助教
function createNewSpreadsheetWithCopiedSheet() {
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sourceSpreadsheet.getSheetByName("Sheet1"); // 設定要複製的工作表名稱

  // 創建新的試算表
  var newSpreadsheet = SpreadsheetApp.create("new Spreadsheet");

  // 刪除新試算表中的預設工作表（"Sheet1"）
  var sheet = newSpreadsheet.getSheetByName('Sheet1');
  // newSpreadsheet.deleteSheet(sheet);

  // 複製工作表到新的試算表
  sourceSheet.copyTo(newSpreadsheet);

  // 重命名複製的工作表
  var copiedSheet = newSpreadsheet.getSheets()[0];
  copiedSheet.setName("Copied Sheet");

  // 輸出新試算表的 URL
  console.log(newSpreadsheet.getUrl());
}

