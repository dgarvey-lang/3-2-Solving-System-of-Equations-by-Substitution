/**
 * OPTIONAL: send student scores straight into a Google Sheet.
 *
 * One-time setup (about 5 minutes):
 *  1. Make a new Google Sheet (name it something like "Radicals Review Scores").
 *  2. In the Sheet: Extensions > Apps Script. Delete what's there and paste this whole file. Save.
 *  3. Click Deploy > New deployment. Click the gear and pick "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone   (students don't need to sign in; they can only ADD rows, not read them)
 *     Click Deploy and allow the permissions it asks for.
 *  4. Copy the "Web app URL" (it ends in /exec).
 *  5. Open alg2b-radicals-quadratics-review.html, find  var SHEET_URL = "";  near the top of the script,
 *     paste the URL between the quotes, and save/commit the file.
 *
 * Every time a student presses "Submit & see my score", a row is added to the "Scores" tab.
 */
var HEADER = ["Received", "Student", "Period", "Points", "Out of", "Percent", "First try right",
  "Attempt", "Started", "Finished", "Problem set", "File OK",
  "Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8", "Q9", "Q10", "Q11", "Q12"];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Scores") || ss.insertSheet("Scores");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADER);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADER.length).setFontWeight("bold");
    }
    var r = JSON.parse(e.postData.contents);
    var check = r.check;
    delete r.check;
    var ok = hash_("pa-alg2b-7Q2x" + JSON.stringify(r)) === check;
    var row = [new Date(), r.name, r.period, r.points, r.outOf, r.percent, r.firstTry,
      r.attempt, r.started, r.finished, r.set, ok ? "yes" : "EDITED"];
    (r.results || []).forEach(function (x) {
      row.push(x.result === "✓" ? 1 : x.result === "½" ? 0.5 : 0);
    });
    sheet.appendRow(row);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

// same checksum the review page uses, so edited submissions are flagged
function hash_(str) {
  var h = 5381;
  for (var i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}
