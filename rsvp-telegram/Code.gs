/**
 * Sharaf & Nour wedding RSVP endpoint.
 * Deploy as a Google Apps Script web app (execute as you, access: anyone).
 * Script Properties required: RSVP_SPREADSHEET_ID
 */

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    message: 'Sharaf & Nour RSVP endpoint is ready.'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('RSVP_SPREADSHEET_ID');
    if (!spreadsheetId) throw new Error('Missing RSVP_SPREADSHEET_ID in Script Properties.');

    const request = parseRequest_(e);
    const name = String(request.name || '').trim().replace(/\s+/g, ' ');
    const attendance = String(request.attendance || '').trim().toLowerCase();
    console.log('RSVP received: name present=' + Boolean(name) + ', attendance=' + attendance);
    if (!name || name.length > 100 || !['yes', 'no'].includes(attendance)) {
      return resultPage_(false);
    }

    const reply = attendance === 'yes' ? 'Attending' : 'Unable to attend';
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getSheetByName('Replies') || spreadsheet.insertSheet('Replies');
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Received at', 'Guest name', 'Reply']);
        sheet.setFrozenRows(1);
      }
      sheet.appendRow([new Date(), safeSheetText_(name), reply]);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return resultPage_(true);
  } catch (error) {
    console.error(String(error && error.stack || error));
    return resultPage_(false);
  }
}

function resultPage_(ok) {
  const result = JSON.stringify({ type: 'sn-rsvp-result', ok: Boolean(ok) })
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
  const html = '<!doctype html><meta charset="utf-8"><script>' +
    'window.top.postMessage(' + result + ', "*");' +
    '<\/script><title>RSVP</title>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function parseRequest_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (_) { /* use form fields */ }
  }
  return (e && e.parameter) || {};
}

function safeSheetText_(value) {
  return /^[=+@-]/.test(value) ? "'" + value : value;
}

