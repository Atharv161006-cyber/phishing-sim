// === Phishing Awareness Simulation — Logging Backend ===
// Deploy this as a Google Apps Script Web App (see DEPLOY_INSTRUCTIONS.txt)
// It logs ONLY event flags + timestamp. It never receives or stores
// the actual password text — the front-end page never sends it.

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Participant ID", "Event"]);
  }

  sheet.appendRow([
    new Date(),
    data.participantId || "unknown",
    data.event || "unknown"
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}
