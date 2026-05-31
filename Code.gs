/*
 * Student Grade Portal backend.
 *
 * How this works:
 * - Every sheet tab is treated as one subject.
 * - Run sendStudentPasswords() manually from Apps Script to create passwords,
 *   email students, and mark Email_send as Send.
 * - Students log in with Subject + PRN/Application ID + their personal password.
 * - The dashboard never returns internal Password or Email_send columns.
 */

const SHEET_ID = '19PaL6OByZeP8A6eKWYS9Nm7FET5WLEkFGu3A0_jyPOU';
const PASSWORD_COLUMN_NAME = 'Password';
const EMAIL_SEND_COLUMN_NAME = 'Email_send';
const EMAIL_SENT_VALUE = 'Send';
const EMAIL_UNSENT_VALUE = 'unsend';
const PENDING_VALUE = 'Pending';
const PASSWORD_LENGTH = 8;
const PORTAL_URL = 'https://007arjungangwar.github.io/StudentGradePortal/';
const EMAIL_SENDER_NAME = 'Arjun - Student Grade Portal';

function doPost(e) {
  try {
    const request = parseRequest(e);
    const action = String(request.action || 'login').toLowerCase();

    if (action === 'subjects') {
      return jsonResponse(true, {
        subjects: getSubjectNames()
      }, 'Subjects loaded.');
    }

    if (action === 'identifiers') {
      return handleIdentifiers(request);
    }

    if (action === 'login') {
      return handleLogin(request);
    }

    return jsonResponse(false, null, 'Unknown action.');
  } catch (error) {
    return jsonResponse(false, null, 'Server error: ' + error.message);
  }
}

function doGet() {
  return jsonResponse(true, {
    subjects: getSubjectNames()
  }, 'Student Grade Portal API is running.');
}

function handleLogin(request) {
  const subject = String(request.subject || '').trim();
  const identifier = String(
    request.identifier ||
    request.prn ||
    request.applicationId ||
    request.application_id ||
    ''
  ).trim();
  const identifierType = normalizeIdentifierType(
    request.identifierType ||
    request.identifier_type ||
    inferIdentifierType(request)
  );
  const password = String(request.password || '');

  if (!subject) {
    return jsonResponse(false, null, 'Please select a subject.');
  }

  if (!identifier || !password) {
    return jsonResponse(false, null, 'Please enter PRN/Application ID and password.');
  }

  const sheet = getTargetSheet(subject);

  if (!sheet) {
    return jsonResponse(false, null, 'Subject not found.');
  }

  const values = getSheetValues(sheet);

  if (values.length < 2) {
    return jsonResponse(false, null, 'This subject sheet does not contain student records.');
  }

  const headers = buildHeaders(values[0]);
  const prnIndex = findPrnColumn(headers);
  const applicationIndex = findApplicationColumn(headers);
  const passwordIndex = findPasswordColumn(headers);

  if (prnIndex === -1 && applicationIndex === -1) {
    return jsonResponse(false, null, 'PRN/Application ID column not found for this subject.');
  }

  if (passwordIndex === -1) {
    return jsonResponse(false, null, 'Password column not found. Please run sendStudentPasswords() first.');
  }

  const requestedIdentifier = normalizeValue(identifier);
  const studentRow = values.slice(1).find(function(row) {
    return rowMatchesLoginIdentifier(row, prnIndex, applicationIndex, requestedIdentifier, identifierType);
  });

  if (!studentRow) {
    return jsonResponse(false, null, 'PRN/Application ID not found.');
  }

  const storedPassword = String(studentRow[passwordIndex] || '').trim();

  if (!storedPassword) {
    return jsonResponse(false, null, 'Password has not been generated for this record yet.');
  }

  if (password !== storedPassword) {
    return jsonResponse(false, null, 'Password error.');
  }

  return jsonResponse(true, buildStudentData(sheet.getName(), headers, studentRow), 'Student record found.');
}

function handleIdentifiers(request) {
  const subject = String(request.subject || '').trim();

  if (!subject) {
    return jsonResponse(false, null, 'Please select a subject.');
  }

  const sheet = getTargetSheet(subject);

  if (!sheet) {
    return jsonResponse(false, null, 'Subject not found.');
  }

  const values = getSheetValues(sheet);

  if (values.length < 2) {
    return jsonResponse(true, {
      identifiers: []
    }, 'No student records found for this subject.');
  }

  const headers = buildHeaders(values[0]);
  const prnIndex = findPrnColumn(headers);
  const applicationIndex = findApplicationColumn(headers);

  if (prnIndex === -1 && applicationIndex === -1) {
    return jsonResponse(false, null, 'PRN/Application ID column not found for this subject.');
  }

  return jsonResponse(true, {
    identifiers: buildLoginIdentifiers(values.slice(1), prnIndex, applicationIndex)
  }, 'Identifiers loaded.');
}

function sendStudentPasswords() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const results = spreadsheet.getSheets().map(function(sheet) {
    return sendPasswordsForSheet(sheet);
  });

  Logger.log(JSON.stringify(results, null, 2));
  return results;
}

function sendPasswordsForSheet(sheet) {
  const result = {
    subject: sheet.getName(),
    generated: 0,
    sent: 0,
    skippedAlreadySent: 0,
    skippedNoEmail: 0,
    sentRecipients: [],
    errors: []
  };

  try {
    if (sheet.getLastRow() < 2) {
      result.errors.push('No student rows found.');
      return result;
    }

    preparePasswordColumns(sheet);

    const values = getSheetValues(sheet);
    const headers = buildHeaders(values[0]);
    const prnIndex = findPrnColumn(headers);
    const applicationIndex = findApplicationColumn(headers);
    const identifierIndex = findPrimaryIdentifierColumn(headers);
    const emailIndex = findEmailColumn(headers);
    const nameIndex = findNameColumn(headers);
    const passwordIndex = findPasswordColumn(headers);
    const emailSendIndex = findEmailSendColumn(headers);

    if (identifierIndex === -1) {
      result.errors.push('PRN/Application ID column not found.');
      return result;
    }

    if (emailIndex === -1) {
      result.errors.push('Email column not found.');
      return result;
    }

    values.slice(1).forEach(function(row, rowOffset) {
      const sheetRow = rowOffset + 2;
      const prn = prnIndex === -1 ? '' : String(row[prnIndex] || '').trim();
      const applicationId = applicationIndex === -1 ? '' : String(row[applicationIndex] || '').trim();
      const loginIdentifier = prn || applicationId;
      const email = String(row[emailIndex] || '').trim();
      const name = nameIndex === -1 ? '' : String(row[nameIndex] || '').trim();
      let password = String(row[passwordIndex] || '').trim();
      let emailStatus = String(row[emailSendIndex] || '').trim();
      const hadPassword = Boolean(password);

      if (!loginIdentifier) {
        return;
      }

      if (!password) {
        password = generatePassword(PASSWORD_LENGTH);
        sheet.getRange(sheetRow, passwordIndex + 1).setValue(password);
        result.generated++;
      }

      if (!hadPassword || !emailStatus) {
        emailStatus = EMAIL_UNSENT_VALUE;
        sheet.getRange(sheetRow, emailSendIndex + 1).setValue(EMAIL_UNSENT_VALUE);
      }

      if (isEmailAlreadySent(emailStatus)) {
        result.skippedAlreadySent++;
        return;
      }

      if (!email) {
        sheet.getRange(sheetRow, emailSendIndex + 1).setValue(EMAIL_UNSENT_VALUE);
        result.skippedNoEmail++;
        return;
      }

      if (!isValidEmail(email)) {
        sheet.getRange(sheetRow, emailSendIndex + 1).setValue(EMAIL_UNSENT_VALUE);
        result.errors.push('Row ' + sheetRow + ': Invalid email address "' + email + '".');
        return;
      }

      try {
        MailApp.sendEmail({
          to: email,
          subject: 'Student Grade Portal Password - ' + sheet.getName(),
          name: EMAIL_SENDER_NAME,
          body: buildPasswordEmailBody(name, prn, applicationId, password, sheet.getName()),
          htmlBody: buildPasswordEmailHtmlBody(name, prn, applicationId, password, sheet.getName())
        });
        sheet.getRange(sheetRow, emailSendIndex + 1).setValue(EMAIL_SENT_VALUE);
        result.sent++;
        result.sentRecipients.push({
          row: sheetRow,
          prn: prn,
          applicationId: applicationId,
          email: email
        });
      } catch (error) {
        sheet.getRange(sheetRow, emailSendIndex + 1).setValue(EMAIL_UNSENT_VALUE);
        result.errors.push('Row ' + sheetRow + ': ' + error.message);
      }
    });
  } catch (error) {
    result.errors.push(error.message);
  }

  return result;
}

function preparePasswordColumns(sheet) {
  let headers = buildHeaders(getHeaderRow(sheet));
  let identifierIndex = findPrimaryIdentifierColumn(headers);

  if (identifierIndex === -1) {
    throw new Error('PRN/Application ID column not found in ' + sheet.getName() + '.');
  }

  let passwordIndex = findPasswordColumn(headers);

  if (passwordIndex === -1) {
    sheet.insertColumnAfter(identifierIndex + 1);
    sheet.getRange(1, identifierIndex + 2).setValue(PASSWORD_COLUMN_NAME);
  }

  headers = buildHeaders(getHeaderRow(sheet));
  passwordIndex = findPasswordColumn(headers);

  if (passwordIndex === -1) {
    throw new Error('Could not create Password column in ' + sheet.getName() + '.');
  }

  const emailSendIndex = findEmailSendColumn(headers);

  if (emailSendIndex === -1) {
    sheet.insertColumnAfter(passwordIndex + 1);
    sheet.getRange(1, passwordIndex + 2).setValue(EMAIL_SEND_COLUMN_NAME);
  }
}

function getSubjectNames() {
  return SpreadsheetApp
    .openById(SHEET_ID)
    .getSheets()
    .map(function(sheet) {
      return sheet.getName();
    });
}

function getTargetSheet(subject) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(subject);
}

function getHeaderRow(sheet) {
  return sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
}

function getSheetValues(sheet) {
  return sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), 1)).getDisplayValues();
}

function buildStudentData(subject, headers, row) {
  const data = {
    Subject: subject
  };

  headers.forEach(function(header, index) {
    if (isInternalHeader(header)) {
      return;
    }

    const value = String(row[index] || '').trim();
    data[header] = value || PENDING_VALUE;
  });

  return data;
}

function buildPasswordEmailBody(name, prn, applicationId, password, subject) {
  const greetingName = name || 'Student';
  const loginLabel = prn && applicationId ? 'PRN/Application ID' : prn ? 'PRN' : 'Application ID';

  const lines = [
    'Hello ' + greetingName + ',',
    '',
    'Your Student Grade Portal login details are ready.',
    '',
    'Subject: ' + subject
  ];

  if (prn) {
    lines.push('PRN: ' + prn);
  }

  if (applicationId) {
    lines.push('Application ID: ' + applicationId);
  }

  lines.push(
    'Password: ' + password,
    '',
    'I have not uploaded the other entries yet, but I will upload them soon.',
    'Please check again after one or two days. The updated entries will be available here.',
    'If you have any doubts regarding any marks, please contact me.',
    'Thanks for your patience.',
    '',
    'Portal link:',
    PORTAL_URL,
    '',
    'Please keep this password private. Anyone with your ' + loginLabel + ' and password can',
    'view your record for this subject.',
    '',
    'Regards,',
    'Arjun'
  );

  return lines.join('\n');
}

function buildPasswordEmailHtmlBody(name, prn, applicationId, password, subject) {
  const greetingName = name || 'Student';
  const loginLabel = prn && applicationId ? 'PRN/Application ID' : prn ? 'PRN' : 'Application ID';
  const detailRows = [
    buildEmailDetailRow('Subject', subject, '#1f5fbf')
  ];

  if (prn) {
    detailRows.push(buildEmailDetailRow('PRN', prn, '#0b7a75'));
  }

  if (applicationId) {
    detailRows.push(buildEmailDetailRow('Application ID', applicationId, '#7c3aed'));
  }

  detailRows.push(buildEmailDetailRow('Password', password, '#b7791f'));

  return [
    '<div style="margin:0;padding:24px;background:#f3f6fb;font-family:Arial,sans-serif;color:#172033;">',
    '<div style="max-width:620px;margin:0 auto;background:#ffffff;border:1px solid #d7dee8;border-radius:8px;overflow:hidden;">',
    '<div style="background:#12335d;color:#ffffff;padding:20px 24px;">',
    '<div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#c8f7ff;">Student Grade Portal</div>',
    '<h1 style="margin:8px 0 0;font-size:22px;line-height:1.3;">Login details are ready</h1>',
    '</div>',
    '<div style="padding:24px;">',
    '<p style="margin:0 0 14px;font-size:16px;">Hello ' + escapeHtml(greetingName) + ',</p>',
    '<p style="margin:0 0 18px;line-height:1.6;">Your Student Grade Portal login details are ready.</p>',
    '<div style="display:grid;gap:10px;margin:0 0 18px;">',
    detailRows.join(''),
    '</div>',
    '<div style="border-left:4px solid #1f5fbf;background:#eef6ff;padding:14px 16px;margin:0 0 18px;line-height:1.6;">',
    'I have not uploaded the other entries yet, but I will upload them soon.<br>',
    'Please check again after one or two days. The updated entries will be available here.<br>',
    'If you have any doubts regarding any marks, please contact me.<br>',
    'Thanks for your patience.',
    '</div>',
    '<p style="margin:0 0 18px;">',
    '<a href="' + escapeHtml(PORTAL_URL) + '" style="display:inline-block;background:#0b7a75;color:#ffffff;text-decoration:none;font-weight:700;padding:12px 16px;border-radius:8px;">Open Student Portal</a>',
    '</p>',
    '<p style="margin:0;color:#64748b;line-height:1.6;">Please keep this password private. Anyone with your ' + escapeHtml(loginLabel) + ' and password can view your record for this subject.</p>',
    '<p style="margin:22px 0 0;line-height:1.6;">Regards,<br><strong>Arjun</strong></p>',
    '</div>',
    '</div>',
    '</div>'
  ].join('');
}

function buildEmailDetailRow(label, value, color) {
  return [
    '<div style="border:1px solid #d7dee8;border-radius:8px;overflow:hidden;">',
    '<div style="background:' + color + ';color:#ffffff;font-size:12px;font-weight:700;text-transform:uppercase;padding:7px 10px;">' + escapeHtml(label) + '</div>',
    '<div style="background:#ffffff;color:#172033;font-size:17px;font-weight:700;padding:10px;">' + escapeHtml(value) + '</div>',
    '</div>'
  ].join('');
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

function findApplicationColumn(headers) {
  return headers.findIndex(function(header) {
    const normalized = normalizeHeader(header);
    return normalized === 'applicationid'
      || normalized === 'applicationno'
      || normalized === 'applicationnumber'
      || normalized === 'appid'
      || normalized === 'appno'
      || normalized.indexOf('application') !== -1;
  });
}

function findPrimaryIdentifierColumn(headers) {
  const prnIndex = findPrnColumn(headers);

  if (prnIndex !== -1) {
    return prnIndex;
  }

  return findApplicationColumn(headers);
}

function buildLoginIdentifiers(rows, prnIndex, applicationIndex) {
  const seen = {};
  const identifiers = [];

  rows.forEach(function(row) {
    addLoginIdentifier(identifiers, seen, 'PRN', row, prnIndex);
    addLoginIdentifier(identifiers, seen, 'Application ID', row, applicationIndex);
  });

  return identifiers;
}

function addLoginIdentifier(identifiers, seen, type, row, columnIndex) {
  if (columnIndex === -1) {
    return;
  }

  const value = String(row[columnIndex] || '').trim();

  if (!value) {
    return;
  }

  const identifierType = normalizeIdentifierType(type) || normalizeHeader(type);
  const key = identifierType + ':' + normalizeValue(value);

  if (seen[key]) {
    return;
  }

  seen[key] = true;
  identifiers.push({
    key: key,
    type: type,
    value: value,
    label: type + ': ' + value
  });
}

function rowMatchesLoginIdentifier(row, prnIndex, applicationIndex, requestedIdentifier, identifierType) {
  if (identifierType === 'prn') {
    return columnMatchesValue(row, prnIndex, requestedIdentifier);
  }

  if (identifierType === 'application') {
    return columnMatchesValue(row, applicationIndex, requestedIdentifier);
  }

  return columnMatchesValue(row, prnIndex, requestedIdentifier)
    || columnMatchesValue(row, applicationIndex, requestedIdentifier);
}

function columnMatchesValue(row, columnIndex, requestedValue) {
  return columnIndex !== -1 && normalizeValue(row[columnIndex]) === requestedValue;
}

function findEmailColumn(headers) {
  return headers.findIndex(function(header) {
    const normalized = normalizeHeader(header);
    return normalized !== normalizeHeader(EMAIL_SEND_COLUMN_NAME)
      && (normalized === 'email' || normalized === 'emailid' || normalized.indexOf('email') !== -1);
  });
}

function findNameColumn(headers) {
  return headers.findIndex(function(header) {
    const normalized = normalizeHeader(header);
    return normalized === 'name' || normalized === 'studentname' || normalized.indexOf('fullname') !== -1;
  });
}

function findPasswordColumn(headers) {
  return headers.findIndex(function(header) {
    const normalized = normalizeHeader(header);
    return normalized === 'password' || normalized === 'studentpassword' || normalized === 'personalpassword';
  });
}

function findEmailSendColumn(headers) {
  return headers.findIndex(function(header) {
    return normalizeHeader(header) === normalizeHeader(EMAIL_SEND_COLUMN_NAME);
  });
}

function isInternalHeader(header) {
  const normalized = normalizeHeader(header);
  return normalized === normalizeHeader(PASSWORD_COLUMN_NAME)
    || normalized === normalizeHeader(EMAIL_SEND_COLUMN_NAME);
}

function isEmailAlreadySent(value) {
  return normalizeValue(value) === normalizeValue(EMAIL_SENT_VALUE);
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function generatePassword(length) {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
  let password = '';

  for (let i = 0; i < length; i++) {
    password += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  }

  return password;
}

function normalizeHeader(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normalizeValue(value) {
  return String(value || '').trim().toLowerCase();
}

function inferIdentifierType(request) {
  if (request.identifier) {
    return '';
  }

  if (request.prn) {
    return 'prn';
  }

  if (request.applicationId || request.application_id) {
    return 'application';
  }

  return '';
}

function normalizeIdentifierType(value) {
  const normalized = normalizeHeader(value);

  if (normalized === 'prn' || normalized.indexOf('prn') !== -1) {
    return 'prn';
  }

  if (normalized === 'application'
      || normalized === 'applicationid'
      || normalized === 'applicationno'
      || normalized === 'applicationnumber'
      || normalized === 'appid'
      || normalized === 'appno'
      || normalized.indexOf('application') !== -1) {
    return 'application';
  }

  return '';
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
