// ============================================================
// WEDDING INVITATION - Google Apps Script
// Paste this into: script.google.com → New Project
// Then: Deploy → New deployment → Web app
//   - Execute as: Me
//   - Who has access: Anyone
// Copy the Web App URL and paste into wedding_invitation.html
// ============================================================

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(e.postData.contents);
  
  if (data.type === 'rsvp') {
    const rsvpSheet = sheet.getSheetByName('RSVP') || sheet.insertSheet('RSVP');
    if (rsvpSheet.getLastRow() === 0) {
      rsvpSheet.appendRow(['Timestamp', 'Name', 'Attendance', 'Guests']);
    }
    rsvpSheet.appendRow([
      new Date().toLocaleString('id-ID'),
      data.name,
      data.attend === 'hadir' ? 'HADIR' : 'TIDAK HADIR',
      data.guests
    ]);
  }
  
  if (data.type === 'wish') {
    const wishSheet = sheet.getSheetByName('Wishes') || sheet.insertSheet('Wishes');
    if (wishSheet.getLastRow() === 0) {
      wishSheet.appendRow(['Timestamp', 'Name', 'Message']);
    }
    wishSheet.appendRow([
      new Date().toLocaleString('id-ID'),
      data.name,
      data.message
    ]);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({status: 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};
  
  const rsvpSheet = sheet.getSheetByName('RSVP');
  if (rsvpSheet && rsvpSheet.getLastRow() > 1) {
    const rows = rsvpSheet.getRange(2, 1, rsvpSheet.getLastRow()-1, 4).getValues();
    result.rsvp = rows.map(r => ({ts:r[0], name:r[1], attend:r[2], guests:r[3]}));
  } else { result.rsvp = []; }
  
  const wishSheet = sheet.getSheetByName('Wishes');
  if (wishSheet && wishSheet.getLastRow() > 1) {
    const rows = wishSheet.getRange(2, 1, wishSheet.getLastRow()-1, 3).getValues();
    result.wishes = rows.map(r => ({ts:r[0], name:r[1], message:r[2]}));
  } else { result.wishes = []; }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
