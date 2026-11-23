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

// Sheet menu: Eunos › Send preview to selected rows. Emails Eunos.dmg (from
// your Drive) as an attachment to each selected row not yet sent, and writes
// "sent <date>" in column C so nobody gets it twice.
// ponytail: attachment, not a one-time link; a link can't stop the DMG being forwarded either.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Eunos').addItem('Send preview to selected rows', 'sendPreview').addToUi();
}

function sendPreview() {
  const files = DriveApp.getFilesByName('Eunos.dmg');
  if (!files.hasNext()) return SpreadsheetApp.getUi().alert('Upload Eunos.dmg to your Google Drive first.');
  const dmg = files.next().getBlob();
  const sheet = SpreadsheetApp.getActiveSheet();
  let sent = 0;
  SpreadsheetApp.getActiveRange().getValues().forEach((_, i) => {
    const row = SpreadsheetApp.getActiveRange().getRow() + i;
    const [email, , done] = sheet.getRange(row, 1, 1, 3).getValues()[0];
    if (!email || done) return;
    MailApp.sendEmail({
      to: email,
      subject: 'Your Eunos preview',
      body: 'Thanks for trying Eunos.\n\nThe Mac preview is attached. Open Eunos.dmg and drag Eunos to Applications. It needs macOS 14 or later on Apple silicon.\n\nReply to this email with anything that breaks or feels wrong.',
      attachments: [dmg],
    });
    sheet.getRange(row, 3).setValue('sent ' + new Date().toISOString().slice(0, 10));
    sent++;
  });
  SpreadsheetApp.getUi().alert('Sent to ' + sent + ' address(es).');
}
