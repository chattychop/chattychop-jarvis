// ═══ CHATTY CHOP JARVIS — Google Apps Script Backend ═══
// PASTE THIS ENTIRE CODE into Google Apps Script

const SHEET_ID = '1AmXlYcwPXhwr_IDrMdI4IiQ7Yv3I4aHHQi00rkVYMKA';

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    
    // Add headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Time', 'Country', 'City', 'Device', 'Browser', 'Source', 'Tabs Used', 'Time Spent (s)', 'Session ID']);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#000810').setFontColor('#00f5ff');
    }

    const p = e.parameter;
    
    // Check if this session already exists
    const data = sheet.getDataRange().getValues();
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][8] === p.sessionId) { existingRow = i + 1; break; }
    }

    const row = [
      p.time || new Date().toLocaleString(),
      p.country || 'Unknown',
      p.city || 'Unknown',
      p.device || 'Unknown',
      p.browser || 'Unknown',
      p.source || 'Direct',
      p.tabsUsed || '',
      parseInt(p.timeSpent) || 0,
      p.sessionId || ''
    ];

    if (existingRow > 0) {
      // Update existing row
      sheet.getRange(existingRow, 1, 1, 9).setValues([row]);
    } else {
      // New visitor — add row
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error', msg: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Read all visits for dashboard
function doPost(e) {
  return doGet(e);
}
