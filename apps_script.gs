// ===== Spot-the-AI data sink — Google Apps Script =====
// Setup (one time, ~2 min):
//  1. Create a new Google Sheet.
//  2. Extensions > Apps Script. Delete the stub, paste ALL of this.
//  3. Deploy > New deployment > type "Web app".
//       Execute as: Me   |   Who has access: Anyone
//  4. Copy the Web app URL (ends in /exec) and send it to me.
// Data lands in a 'plays' tab, one row per completed game; answers as JSON.
// no-cors POST from the page works because we never read the response.

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('plays') || ss.insertSheet('plays');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['recv_ts','pid','play','client_ts','v','score','lang','ua','answers_json']);
    }
    sh.appendRow([new Date(), data.pid, data.play, data.ts, data.v,
                  data.score, data.lang, data.ua, JSON.stringify(data.answers)]);
  } catch (err) { /* swallow: never block the client */ }
  return ContentService.createTextOutput('ok');
}

function doGet() { return ContentService.createTextOutput('alive'); }
