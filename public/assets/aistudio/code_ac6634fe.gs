/**
 * RadSafe - Google Sheets Backend (Google Apps Script)
 * Supports:
 *  - GET  : Fetch all incident reports as JSON
 *  - POST : Create new report, Update RCA/CAPA/Status, Delete single report, or Clear all
 */

const SHEET_NAME = 'Incident_Register';

const HEADERS = [
  'id',
  'classification',
  'datetime',
  'modality',
  'category',
  'severity',
  'description',
  'action',
  'factors',
  'reporterName',
  'reporterRole',
  'status',
  'investigator',
  'rca',
  'updatedAt'
];

/**
 * Ensures the target sheet exists and has the exact header row.
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#0d9488')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Helper to return clean JSON responses.
 */
function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles GET requests: returns all reports as a JSON array.
 */
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return jsonResponse({ status: 'ok', reports: [] });
    }

    const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const reports = data
      .filter(row => row[0]) // Ignore empty rows
      .map(row => {
        let parsedFactors = [];
        if (row[8]) {
          try {
            parsedFactors = JSON.parse(row[8]);
          } catch (err) {
            parsedFactors = String(row[8]).split(',').map(s => s.trim()).filter(Boolean);
          }
        }
        return {
          id: String(row[0]),
          classification: String(row[1]),
          datetime: String(row[2]),
          modality: String(row[3]),
          category: String(row[4]),
          severity: String(row[5]),
          description: String(row[6]),
          action: String(row[7]),
          factors: parsedFactors,
          reporterName: String(row[9]),
          reporterRole: String(row[10]),
          status: String(row[11] || 'Submitted'),
          investigator: String(row[12] || ''),
          rca: String(row[13] || '')
        };
      });

    // Return newest reports first
    reports.reverse();
    return jsonResponse({ status: 'ok', reports: reports });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * Handles POST requests:
 *  - action: 'create' | 'update' | 'delete' | 'clear'
 * Uses LockService so simultaneous hospital submissions never overwrite each other.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getOrCreateSheet();
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'create';
    const nowIso = new Date().toISOString();

    // 1. CREATE NEW REPORT
    if (action === 'create') {
      const r = body.report;
      const rowValues = [
        r.id || '',
        r.classification || '',
        r.datetime || '',
        r.modality || '',
        r.category || '',
        r.severity || '',
        r.description || '',
        r.action || '',
        JSON.stringify(r.factors || []),
        r.reporterName || 'Anonymous',
        r.reporterRole || 'Anonymous',
        r.status || 'Submitted',
        r.investigator || '',
        r.rca || '',
        nowIso
      ];
      sheet.appendRow(rowValues);
      return jsonResponse({ status: 'ok', message: 'Report saved to Google Sheet', id: r.id });
    }

    // 2. UPDATE EXISTING REPORT (Status, Investigator, RCA/CAPA)
    if (action === 'update') {
      const r = body.report;
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return jsonResponse({ status: 'error', message: 'Report not found' });

      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      const rowIndex = ids.indexOf(r.id);
      if (rowIndex === -1) return jsonResponse({ status: 'error', message: 'Report ID not found' });

      const targetRow = rowIndex + 2;
      sheet.getRange(targetRow, 12, 1, 4).setValues([[
        r.status || 'Submitted',
        r.investigator || '',
        r.rca || '',
        nowIso
      ]]);
      return jsonResponse({ status: 'ok', message: 'Report updated', id: r.id });
    }

    // 3. DELETE SINGLE REPORT
    if (action === 'delete') {
      const targetId = body.id;
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return jsonResponse({ status: 'error', message: 'No records found' });

      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      const rowIndex = ids.indexOf(targetId);
      if (rowIndex === -1) return jsonResponse({ status: 'error', message: 'Report ID not found' });

      sheet.deleteRow(rowIndex + 2);
      return jsonResponse({ status: 'ok', message: 'Report deleted', id: targetId });
    }

    // 4. CLEAR ALL REPORTS (Keeps Header Row)
    if (action === 'clear') {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      return jsonResponse({ status: 'ok', message: 'All reports cleared' });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}