var LABEL_NAME = "機器人照會";
var PROGRESS_KEY = "exportProgress";
var SS_ID_KEY = "exportSSId";

function startExport() {
  // 清除舊進度
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROGRESS_KEY);
  props.deleteProperty(SS_ID_KEY);
  
  // 建立試算表
  var ss = SpreadsheetApp.create("機器人照會_信件匯出");
  var sheet = ss.getActiveSheet();
  sheet.appendRow(["日期", "寄件人", "主旨", "內容"]);
  props.setProperty(SS_ID_KEY, ss.getId());
  
  // 建立觸發器
  deleteTrigger();
  ScriptApp.newTrigger("exportBatch")
    .timeBased()
    .everyMinutes(10)
    .create();
  
  Logger.log("✅ 開始匯出，試算表ID: " + ss.getId());
  exportBatch();
}

function exportBatch() {
  var props = PropertiesService.getScriptProperties();
  var start = parseInt(props.getProperty(PROGRESS_KEY) || "0");
  var ssId = props.getProperty(SS_ID_KEY);
  var startTime = new Date().getTime();
  var MAX_RUNTIME = 5 * 60 * 1000;
  
  var label = GmailApp.getUserLabelByName(LABEL_NAME);
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getActiveSheet();
  var offset = start;
  var count = 0;

  while (true) {
    // 檢查是否快超時
    if (new Date().getTime() - startTime > MAX_RUNTIME) {
      props.setProperty(PROGRESS_KEY, String(offset));
      Logger.log("⏱ 暫停於第 " + offset + " 個 thread，下次自動續跑");
      return;
    }

    var threads = label.getThreads(offset, 50);
    if (threads.length === 0) {
      // 全部完成，匯出 Excel
      exportToExcel(ssId);
      deleteTrigger();
      props.deleteProperty(PROGRESS_KEY);
      props.deleteProperty(SS_ID_KEY);
      Logger.log("✅ 全部完成！共寫入 " + (sheet.getLastRow() - 1) + " 封信件");
      return;
    }

    var rows = [];
    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      for (var j = 0; j < messages.length; j++) {
        var msg = messages[j];
        rows.push([
          msg.getDate(),
          msg.getFrom(),
          msg.getSubject(),
          msg.getPlainBody().substring(0, 5000)
        ]);
        count++;
      }
    }

    // 批次寫入
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, 4).setValues(rows);
    offset += threads.length;
  }
}

function exportToExcel(ssId) {
  var url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx";
  var blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getBlob().setName("機器人照會_信件匯出.xlsx");
  
  var file = DriveApp.createFile(blob);
  Logger.log("📁 Excel 下載連結：" + file.getDownloadUrl());
}

function deleteTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "exportBatch") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function checkProgress() {
  var props = PropertiesService.getScriptProperties();
  var progress = props.getProperty(PROGRESS_KEY);
  Logger.log(progress ? "目前進度：第 " + progress + " 個 thread" : "尚未開始或已完成");
}
