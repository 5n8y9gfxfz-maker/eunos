// Google Apps Script bound to a Sheet. Deploy > New deployment > Web app,
// Execute as: Me, Who has access: Anyone. Paste the /exec URL into GATE in index.html.
function doPost(e) {
  const email = String(e.parameter.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || email.length > 254) return ContentService.createTextOutput('bad');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const seen = sheet.getRange('A:A').getValues().flat();
  if (!seen.includes(email)) sheet.appendRow([email, new Date()]);
  return ContentService.createTextOutput('ok');
}
