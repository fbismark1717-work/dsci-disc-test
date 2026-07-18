/**
 * Apps Script backend for the RCode/Mirro DISC test.
 * Bind this to a Google Sheet: Extensions -> Apps Script, paste this file
 * as Code.gs, then Deploy -> New deployment -> Web app.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Дата", "Имя кандидата", "Роль",
        "D (баллы)", "I (баллы)", "S (баллы)", "C (баллы)",
        "D %", "I %", "S %", "C %",
        "Основной тип", "Смешанный тип", "Смешанный профиль"
      ]);
    }

    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),
      data.name || "",
      data.role || "",
      data.D, data.I, data.S, data.C,
      data.dPct, data.iPct, data.sPct, data.cPct,
      data.primaryName || "",
      data.secondaryName || "",
      data.isBlend ? "Да" : "Нет"
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
