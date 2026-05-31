/*
 * Student Grade Portal backend.
 *
 * Deploy this file as a Google Apps Script Web App:
 * - Execute as: Me
 * - Who has access: Anyone
 *
 * The frontend sends { prn, password } as text/plain JSON to avoid browser
 * preflight requests. The script reads the configured sheet, finds the PRN
 * column dynamically, and returns every header/value pair for the student row.
 */

const SHEET_ID = '1auZpGAWqwEJhQlvZslZum4AnwhTahpA-fGjvyltvyCA';
const SHEET_NAME = ''; // Leave blank to use the first sheet.
const MASTER_PASSWORD = 'CHANGE_THIS_MASTER_PASSWORD';
const PENDING_VALUE = 'Pending';

function doPost(e) {
  try {
    const request = parseRequest(e);
    const prn = String(request.prn || '').trim();
    const password = String(request.password || '');

    if (!prn || !password) {
      return jsonResponse(false, null, 'Please enter both PRN and password.');
    }

    if (password !== MASTER_PASSWORD) {
      return jsonResponse(false, null, 'Invalid PRN or password.');
    }

    const sheet = getTargetSheet();
    const values = sheet.getDataRange().getDisplayValues();

    if (values.length < 2) {
      return jsonResponse(false, null, 'The sheet does not contain student records.');
    }

    const headers = buildHeaders(values[0]);
    const prnIndex = findPrnColumn(headers);

    if (prnIndex === -1) {
      return jsonResponse(false, null, 'No PRN column was found in the sheet headers.');
    }

    const requestedPrn = normalizeValue(prn);
    const studentRow = values.slice(1).find(function(row) {
      return normalizeValue(row[prnIndex]) === requestedPrn;
    });

    if (!studentRow) {
      return jsonResponse(false, null, 'No student record found for this PRN.');
    }

    const data = {};

    headers.forEach(function(header, index) {
      const value = String(studentRow[index] || '').trim();
      data[header] = value || PENDING_VALUE;
    });

    return jsonResponse(true, data, 'Student record found.');
  } catch (error) {
    return jsonResponse(false, null, 'Server error: ' + error.message);
  }
}

function doGet() {
  return jsonResponse(true, null, 'Student Grade Portal API is running.');
}

function parseRequest(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (error) {
      // Fall back to form parameters below.
    }
  }

  return (e && e.parameter) || {};
}

function getTargetSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = SHEET_NAME
    ? spreadsheet.getSheetByName(SHEET_NAME)
    : spreadsheet.getSheets()[0];

  if (!sheet) {
    throw new Error('Target sheet was not found.');
  }

  return sheet;
}

function buildHeaders(headerRow) {
  const seen = {};

  return headerRow.map(function(rawHeader, index) {
    const baseHeader = String(rawHeader || '').trim() || 'Column ' + (index + 1);
    const key = baseHeader.toLowerCase();
    seen[key] = (seen[key] || 0) + 1;
    return seen[key] === 1 ? baseHeader : baseHeader + ' (' + seen[key] + ')';
  });
}

function findPrnColumn(headers) {
  return headers.findIndex(function(header) {
    const normalized = normalizeHeader(header);
    return normalized === 'prn' || normalized.indexOf('prn') !== -1;
  });
}

function normalizeHeader(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normalizeValue(value) {
  return String(value || '').trim().toLowerCase();
}

function jsonResponse(success, data, message) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: success,
      data: data,
      message: message
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
