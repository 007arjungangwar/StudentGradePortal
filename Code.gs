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

const SHEET_ID = '1auZpGAWqwEJhQlvZslZum4AnwhTahpA-fGjvyltvyCA';
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
    return rowMatchesLoginIdentifier(row, prnIndex, applicationIndex, requestedIdentifier);
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
    const emailIndex = findEmailColumn(headers);
    const nameIndex = findNameColumn(headers);
    const passwordIndex = findPasswordColumn(headers);
    const emailSendIndex = findEmailSendColumn(headers);

    if (prnIndex === -1) {
      result.errors.push('PRN column not found.');
      return result;
    }

    if (emailIndex === -1) {
      result.errors.push('Email column not found.');
      return result;
    }

    values.slice(1).forEach(function(row, rowOffset) {
      const sheetRow = rowOffset + 2;
      const prn = String(row[prnIndex] || '').trim();
      const email = String(row[emailIndex] || '').trim();
      const name = nameIndex === -1 ? '' : String(row[nameIndex] || '').trim();
      let password = String(row[passwordIndex] || '').trim();
      let emailStatus = String(row[emailSendIndex] || '').trim();
      const hadPassword = Boolean(password);

      if (!prn) {
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
          body: buildPasswordEmailBody(name, prn, password, sheet.getName()),
          htmlBody: buildPasswordEmailHtmlBody(name, prn, password, sheet.getName())
        });
        sheet.getRange(sheetRow, emailSendIndex + 1).setValue(EMAIL_SENT_VALUE);
        result.sent++;
        result.sentRecipients.push({
          row: sheetRow,
          prn: prn,
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
  let prnIndex = findPrnColumn(headers);

  if (prnIndex === -1) {
    throw new Error('PRN column not found in ' + sheet.getName() + '.');
  }

  let passwordIndex = findPasswordColumn(headers);

  if (passwordIndex === -1) {
    sheet.insertColumnAfter(prnIndex + 1);
    sheet.getRange(1, prnIndex + 2).setValue(PASSWORD_COLUMN_NAME);
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

function buildPasswordEmailBody(name, prn, password, subject) {
  const greetingName = name || 'Student';

  return [
    'Hello ' + greetingName + ',',
    '',
    'Your Student Grade Portal login details are ready.',
    '',
    'Subject: ' + subject,
    'PRN: ' + prn,
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
    'Please keep this password private. Anyone with your PRN and password can',
    'view your record for this subject.',
    '',
    'Regards,',
    'Arjun'
  ].join('\n');
}

function buildPasswordEmailHtmlBody(name, prn, password, subject) {
  return buildPasswordEmailBody(name, prn, password, subject)
    .split('\n')
    .map(function(line) {
      return escapeHtml(line);
    })
    .join('<br>');
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

function rowMatchesLoginIdentifier(row, prnIndex, applicationIndex, requestedIdentifier) {
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

function jsonResponse(success, data, message) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: success,
      data: data,
      message: message
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
