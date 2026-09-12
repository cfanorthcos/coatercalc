/**
 * Coater Calc — forecast proxy
 *
 * Serves the daypart forecast tab to the Coater Calc app as CSV. Use this if
 * the app's "Test sheet connection" shows the published-CSV and gviz reads
 * being blocked by the browser: an Apps Script web app is served from the same
 * kind of endpoint the app already posts compliance rows to, and always
 * returns the CORS header browsers require.
 *
 * Setup
 *   1. Open the Apps Script project behind your compliance webhook
 *      (Extensions > Apps Script from the compliance sheet).
 *   2. Paste the code below in. If the project already has a doGet, merge the
 *      body of this one into it instead of adding a second doGet.
 *   3. Deploy > Manage deployments > edit the existing deployment >
 *      New version, with "Who has access" set to Anyone. Keeping the same
 *      deployment keeps the /exec URL the app already has.
 *   4. In Coater Calc: Settings > paste that same /exec URL into the
 *      compliance webhook field if it isn't there > Save.
 *
 * The app calls <your /exec URL>?forecast=1 and parses the CSV this returns,
 * so no forecast logic lives here — it just hands over the tab as-is.
 */

// NG Coater Calculator 2.0
var FORECAST_SHEET_ID = '1nelaTqArdP5YoiF6g7TKbM9RacPnHetL9_7GJxnpVAM';

// Name of the tab holding the daypart forecast. Leave '' for the first tab.
var FORECAST_TAB_NAME = '';

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (!params.forecast) {
    return ContentService.createTextOutput('ok');
  }
  var ss = SpreadsheetApp.openById(FORECAST_SHEET_ID);
  var sheet = FORECAST_TAB_NAME ? ss.getSheetByName(FORECAST_TAB_NAME) : ss.getSheets()[0];
  if (!sheet) {
    return ContentService.createTextOutput('tab not found: ' + FORECAST_TAB_NAME)
      .setMimeType(ContentService.MimeType.TEXT);
  }
  var rows = sheet.getDataRange().getDisplayValues();
  var csv = rows.map(function (row) {
    return row.map(csvCell_).join(',');
  }).join('\n');
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}

function csvCell_(value) {
  var text = String(value == null ? '' : value);
  return /[",\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
}
