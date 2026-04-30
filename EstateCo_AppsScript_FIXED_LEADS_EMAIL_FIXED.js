// Estate Co. — Quote + Referral handler with full address fields and Drive uploads
// Quote form saves to Sheet1. Referral form saves to Sheet2.
// Supports Google Places full address components when supplied by the website.

var SHEET_NAME = 'Sheet1';
var REFERRAL_SHEET_NAME = 'Sheet2';
var DRIVE_FOLDER_NAME = 'Estate Co. Quotes';

var NOTIFY_EMAIL = 'sales@estateco.com.au,daniel.nyssen@gmail.com';
var BUSINESS_EMAIL = 'sales@estateco.com.au';
var BUSINESS_PHONE = '0401 857 730';
var BUSINESS_NAME = 'Estate Co.';

// Leave blank if this script is created from the Google Sheet:
// Google Sheet → Extensions → Apps Script
// If standalone, paste the long spreadsheet ID here.
var SHEET_ID = '';

function doGet(e) {
  if (e && e.parameter && e.parameter.payload) return doPost(e);
  return jsonResponse({ result: 'success', message: 'Estate Co. script is running' });
}

function doPost(e) {
  try {
    var data = getPayload(e);
    var action = data.action || 'lead';

    if (action === 'file') return handleFileUpload(data);
    if (action === 'clientError') return handleClientError(data);
    if (action === 'referral' || isReferralLeadType(data.leadType || data.actionType)) return handleReferral(data);

    return handleLead(data);
  } catch (err) {
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'Estate Co. Form Error',
        body: 'There was an error processing a website form:\n\n' + err.toString()
      });
    } catch (mailErr) {}

    return jsonResponse({ result: 'error', message: err.toString() });
  }
}

function handleLead(data) {
  if (!data.name || !data.email) throw new Error('Missing required customer details.');

  var folderUrl = '';
  var sheetSaved = false;
  var businessEmailSent = false;
  var customerEmailSent = false;
  var warnings = [];

  try {
    var folder = getQuoteFolder(data);
    folderUrl = folder.getUrl();
  } catch (folderErr) {
    warnings.push('Drive folder warning: ' + folderErr.toString());
  }

  try {
    saveLeadToSheet(data, folderUrl);
    sheetSaved = true;
  } catch (sheetErr) {
    warnings.push('Sheet save warning: ' + sheetErr.toString());
  }

  try {
    sendBusinessEmail(data, folderUrl, warnings);
    businessEmailSent = true;
  } catch (mailErr) {
    warnings.push('Business email warning: ' + mailErr.toString());
  }

  try {
    sendCustomerEmail(data);
    customerEmailSent = true;
  } catch (customerErr) {
    warnings.push('Customer email warning: ' + customerErr.toString());
  }

  if (warnings.length) {
    try {
      MailApp.sendEmail({
        to: 'daniel.nyssen@gmail.com',
        subject: 'Estate Co. lead warning — ' + value(data.reference || data.name),
        body: 'A website quote lead was received, but one or more processing steps had a warning.\n\n' +
          'Reference: ' + value(data.reference) + '\n' +
          'Name: ' + value(data.name) + '\n' +
          'Email: ' + value(data.email) + '\n' +
          'Phone: ' + value(data.phone) + '\n' +
          'Address: ' + value(data.formattedAddress || data.propertyAddress || data.suburb) + '\n\n' +
          'Warnings:\n- ' + warnings.join('\n- ')
      });
    } catch (backupErr) {}
  }

  if (!businessEmailSent && !customerEmailSent && !sheetSaved) {
    throw new Error('Lead received but Sheet and email processing failed: ' + warnings.join(' | '));
  }

  return jsonResponse({
    result: 'success',
    action: 'lead',
    reference: value(data.reference),
    folderUrl: folderUrl,
    sheetSaved: sheetSaved,
    businessEmailSent: businessEmailSent,
    customerEmailSent: customerEmailSent,
    warnings: warnings
  });
}

function handleReferral(data) {
  if (!data.referrerName || !data.referrerEmail) throw new Error('Missing required referrer details.');
  if (!data.clientName || !data.clientPhone) throw new Error('Missing required client details.');

  var folder = getQuoteFolder(data);
  var folderUrl = folder.getUrl();
  var fileLinks = [];
  var files = Array.isArray(data.files) ? data.files : [];

  for (var i = 0; i < files.length; i++) {
    if (files[i] && files[i].data) fileLinks.push(saveSingleBase64File(files[i], folder));
  }

  saveReferralToSheet(data, folderUrl, fileLinks);
  sendReferralBusinessEmail(data, folderUrl, fileLinks);
  sendReferralReferrerEmail(data);

  return jsonResponse({
    result: 'success',
    action: 'referral',
    reference: value(data.reference),
    folderUrl: folderUrl,
    fileLinks: fileLinks
  });
}

function handleFileUpload(data) {
  if (!data.reference) throw new Error('Missing reference for file upload.');
  if (!data.file || !data.file.data) throw new Error('Missing file data.');

  var folder = getQuoteFolder(data);
  var fileLink = saveSingleBase64File(data.file, folder);
  updateSheetWithFileLink(data.reference, fileLink, data.file.name, data.fileIndex, data.fileTotal, data.leadType || data.actionType);

  return jsonResponse({ result: 'success', action: 'file', reference: value(data.reference), fileLink: fileLink });
}

function handleClientError(data) {
  var ss = getSpreadsheet();
  var isReferral = isReferralLeadType(data.leadType || data.actionType);
  var sheet = isReferral ? getReferralSheet(ss) : getQuoteSheet(ss);
  if (isReferral) ensureReferralHeaders(sheet); else ensureQuoteHeaders(sheet);

  var reference = value(data.reference) || 'NO-REFERENCE';
  var row = findRowByReference(sheet, reference);
  var message = value(data.message);
  var extra = value(data.extra);

  if (row) {
    var notesCol = getColumnByHeader(sheet, 'Notes');
    var statusCol = getColumnByHeader(sheet, 'Status');
    var existingNotes = value(sheet.getRange(row, notesCol).getValue());
    sheet.getRange(row, notesCol).setValue(existingNotes + '\n\nBrowser upload warning: ' + message + (extra ? '\n' + extra : ''));
    sheet.getRange(row, statusCol).setValue('Browser warning');
  } else {
    if (isReferral) {
      sheet.appendRow([new Date(), reference, 'Referral', value(data.referrerName), value(data.referrerEmail), value(data.referrerPhone), value(data.referrerCompany || data.company), value(data.clientName || data.name), value(data.clientPhone || data.phone), value(data.propertyAddress || data.suburb), value(data.formattedAddress || data.propertyAddress || data.suburb), value(data.propertyStreet), value(data.propertySuburb || data.suburb), value(data.propertyState), value(data.propertyPostcode), value(data.propertyCountry), value(data.propertyPlaceId), value(data.propertyType), value(data.services || data.jobType), 'Browser upload warning: ' + message + (extra ? '\n' + extra : ''), '', Number(data.fileCount || 0), value(data.fileNames), '', '', 'Browser warning']);
    } else {
      sheet.appendRow([new Date(), reference, value(data.name), value(data.email), value(data.phone), value(data.propertyAddress || data.suburb), value(data.formattedAddress || data.propertyAddress || data.suburb), value(data.propertyStreet), value(data.propertySuburb || data.suburb), value(data.propertyState), value(data.propertyPostcode), value(data.propertyCountry), value(data.propertyPlaceId), value(data.propertyType), value(data.size), value(data.services), 'Browser upload warning: ' + message + (extra ? '\n' + extra : ''), Number(data.fileCount || 0), value(data.fileNames), '', '', 'Browser warning']);
    }
  }

  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'Estate Co. browser upload warning — ' + reference,
      body: 'The website reported a browser-side upload issue.\n\nReference: ' + reference + '\nName: ' + value(data.name || data.clientName) + '\nEmail: ' + value(data.email || data.referrerEmail) + '\nPhone: ' + value(data.phone || data.clientPhone) + '\n\nMessage: ' + message + '\n\nExtra: ' + extra
    });
  } catch (err) {}

  return jsonResponse({ result: 'success', action: 'clientError', reference: reference });
}

function getPayload(e) {
  var raw = '';
  if (e && e.parameter && e.parameter.payload) raw = e.parameter.payload;
  else if (e && e.postData && e.postData.contents) raw = e.postData.contents;
  if (!raw) throw new Error('Missing form payload.');
  return JSON.parse(raw);
}

function saveLeadToSheet(data, folderUrl) {
  var ss = getSpreadsheet();
  var sheet = getQuoteSheet(ss);
  ensureQuoteHeaders(sheet);

  sheet.appendRow([
    new Date(), value(data.reference), value(data.name), value(data.email), value(data.phone),
    value(data.propertyAddress || data.suburb), value(data.formattedAddress || data.propertyAddress || data.suburb),
    value(data.propertyStreet), value(data.propertySuburb || data.suburb), value(data.propertyState), value(data.propertyPostcode), value(data.propertyCountry), value(data.propertyPlaceId),
    value(data.propertyType), value(data.size), value(data.services), value(data.notes),
    Number(data.fileCount || 0), value(data.fileNames), '', folderUrl, 'New'
  ]);
}

function saveReferralToSheet(data, folderUrl, fileLinks) {
  var ss = getSpreadsheet();
  var sheet = getReferralSheet(ss);
  ensureReferralHeaders(sheet);

  sheet.appendRow([
    new Date(), value(data.reference), 'Referral',
    value(data.referrerName), value(data.referrerEmail), value(data.referrerPhone), value(data.referrerCompany || data.company),
    value(data.clientName || data.name), value(data.clientPhone || data.phone),
    value(data.propertyAddress || data.suburb), value(data.formattedAddress || data.propertyAddress || data.suburb),
    value(data.propertyStreet), value(data.propertySuburb || data.suburb), value(data.propertyState), value(data.propertyPostcode), value(data.propertyCountry), value(data.propertyPlaceId),
    value(data.propertyType), value(data.services || data.jobType), value(data.notes),
    value(data.rewardTerms) || '$50 Prezzee gift card for completed jobs $1,000-$4,999; $100 Prezzee gift card for completed jobs $5,000 and above. Private thank-you gift, not communicated to the client.',
    Number(data.fileCount || (fileLinks ? fileLinks.length : 0)), value(data.fileNames), fileLinks && fileLinks.length ? fileLinks.join('\n') : '', folderUrl, 'New referral'
  ]);
}

function updateSheetWithFileLink(reference, fileLink, fileName, fileIndex, fileTotal, leadType) {
  var ss = getSpreadsheet();
  var isReferral = isReferralLeadType(leadType);
  var sheet = isReferral ? getReferralSheet(ss) : getQuoteSheet(ss);
  if (isReferral) ensureReferralHeaders(sheet); else ensureQuoteHeaders(sheet);

  var row = findRowByReference(sheet, reference);
  if (!row) {
    if (isReferral) sheet.appendRow([new Date(), reference, 'Referral', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'File arrived before lead row', '', Number(fileTotal || 1), value(fileName), fileLink, '', 'File uploaded 1/' + value(fileTotal || 1)]);
    else sheet.appendRow([new Date(), reference, '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'File arrived before lead row', Number(fileTotal || 1), value(fileName), fileLink, '', 'File uploaded 1/' + value(fileTotal || 1)]);
    return;
  }

  var fileLinksCol = getColumnByHeader(sheet, 'File Links');
  var statusCol = getColumnByHeader(sheet, 'Status');
  var existing = value(sheet.getRange(row, fileLinksCol).getValue());
  var updated = existing;

  if (existing.indexOf(fileLink) === -1) {
    updated = existing ? existing + '\n' + fileLink : fileLink;
    sheet.getRange(row, fileLinksCol).setValue(updated);
  }

  var uploadedCount = updated ? updated.split('\n').filter(String).length : 0;
  var total = Number(fileTotal || 0);
  var status = total ? ('Files uploaded ' + uploadedCount + '/' + total) : 'Files uploaded';
  if (total && uploadedCount >= total) status = 'Complete — files uploaded';
  sheet.getRange(row, statusCol).setValue(status);
}

function findRowByReference(sheet, reference) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var refs = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = refs.length - 1; i >= 0; i--) {
    if (String(refs[i][0]) === String(reference)) return i + 2;
  }
  return null;
}

function saveSingleBase64File(file, folder) {
  var size = Number(file.size || 0);
  var maxSize = 10 * 1024 * 1024;
  if (size && size > maxSize) throw new Error('File exceeds 10MB limit: ' + value(file.name));

  var bytes = Utilities.base64Decode(file.data);
  var name = cleanFileName(file.name || 'uploaded-file');
  var type = file.type || 'application/octet-stream';
  var blob = Utilities.newBlob(bytes, type, name);
  var saved = folder.createFile(blob);
  saved.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return saved.getUrl();
}

function getQuoteFolder(data) {
  var root = getOrCreateFolder(DRIVE_FOLDER_NAME);
  var yearFolder = getOrCreateFolder(String(new Date().getFullYear()), root);
  var folderName = cleanFileName((data.reference || 'EC') + ' — ' + (data.name || data.clientName || data.referrerName || 'Unknown') + ' — ' + (data.propertySuburb || data.suburb || data.formattedAddress || data.propertyAddress || 'Unknown'));
  return getOrCreateFolder(folderName, yearFolder);
}

function getSpreadsheet() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet not found. Add SHEET_ID or open Apps Script from the Google Sheet.');
  return ss;
}

function getQuoteSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function getReferralSheet(ss) {
  var sheet = ss.getSheetByName(REFERRAL_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(REFERRAL_SHEET_NAME);
  return sheet;
}

function isReferralLeadType(leadType) {
  var t = String(leadType || '').toLowerCase();
  return t === 'referral' || t === 'referral lead' || t === 'referrer';
}

function ensureQuoteHeaders(sheet) {
  ensureHeaderRow(sheet, ['Date','Reference','Name','Email','Phone','Property Address','Formatted Address','Street','Suburb','State','Postcode','Country','Google Place ID','Property Type','Size','Services','Notes','Expected File Count','File Names','File Links','Drive Folder','Status']);
}

function ensureReferralHeaders(sheet) {
  ensureHeaderRow(sheet, ['Date','Reference','Lead Type','Referrer Name','Referrer Email','Referrer Phone','Referrer Company','Client Name','Client Phone','Property Address','Formatted Address','Street','Suburb','State','Postcode','Country','Google Place ID','Property Type','Likely Job Type','Notes','Reward Terms','Expected File Count','File Names','File Links','Drive Folder','Status']);
}

function ensureHeaderRow(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }
  var current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  var needsUpdate = false;
  for (var i = 0; i < headers.length; i++) {
    if (String(current[i] || '') !== String(headers[i])) { needsUpdate = true; break; }
  }
  if (needsUpdate) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function getColumnByHeader(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]) === String(headerName)) return i + 1;
  }
  throw new Error('Missing required sheet header: ' + headerName);
}

function sendBusinessEmail(data, folderUrl, warnings) {
  warnings = warnings || [];
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    replyTo: value(data.email),
    subject: 'New Estate Co. Quote — ' + (data.name || 'New enquiry'),
    body:
      'New quote request received.\n\n' +
      'Reference: ' + value(data.reference) + '\n' +
      'Name: ' + value(data.name) + '\n' +
      'Email: ' + value(data.email) + '\n' +
      'Phone: ' + value(data.phone) + '\n' +
      'Property address: ' + value(data.formattedAddress || data.propertyAddress || data.suburb) + '\n' +
      'Street: ' + value(data.propertyStreet) + '\n' +
      'Suburb: ' + value(data.propertySuburb || data.suburb) + '\n' +
      'State: ' + value(data.propertyState) + '\n' +
      'Postcode: ' + value(data.propertyPostcode) + '\n' +
      'Property Type: ' + value(data.propertyType) + '\n' +
      'Size: ' + value(data.size) + '\n' +
      'Services: ' + value(data.services) + '\n\n' +
      'Notes:\n' + value(data.notes) + '\n\n' +
      'Expected files: ' + value(data.fileCount) + '\n' +
      'File names: ' + value(data.fileNames) + '\n\n' +
      'Drive folder:\n' + value(folderUrl) + '\n\n' +
      (warnings.length ? 'Processing warnings:\n- ' + warnings.join('\n- ') + '\n' : '')
  });
}

function sendCustomerEmail(data) {
  if (!data.email) return;
  var firstName = html(value(data.name).split(' ')[0] || 'there');
  var reference = html(data.reference || 'Pending');
  var address = html(data.formattedAddress || data.propertyAddress || data.suburb);

  MailApp.sendEmail({
    to: data.email,
    replyTo: BUSINESS_EMAIL,
    name: BUSINESS_NAME,
    subject: 'Thank you for your enquiry — Estate Co.',
    htmlBody:
      '<div style="font-family:Arial,sans-serif;background:#f4f7fa;padding:24px;">' +
      '<div style="max-width:620px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 8px 28px rgba(26,60,94,0.08);">' +
      '<div style="background:#1A3C5E;padding:34px;text-align:center;"><h1 style="color:#ffffff;margin:0;font-size:28px;">Estate Co.</h1><p style="color:#d7e6ef;margin:8px 0 0;font-size:14px;">Not just a clean. A fresh start.</p></div>' +
      '<div style="padding:34px;"><h2 style="color:#1A3C5E;margin:0 0 12px;font-size:22px;">Thank you, ' + firstName + '.</h2>' +
      '<p style="color:#4d5f6d;font-size:15px;line-height:1.7;margin:0 0 24px;">We’ve received your enquiry and will review the details you provided. One of our team will be in touch shortly to discuss the property and next steps.</p>' +
      '<div style="background:#eef6fb;border-left:5px solid #2D7FA6;padding:18px 20px;border-radius:8px;margin-bottom:26px;"><p style="margin:0 0 6px;color:#6b7d89;font-size:12px;text-transform:uppercase;letter-spacing:1px;font-weight:bold;">Reference Number</p><p style="margin:0;color:#1A3C5E;font-size:26px;font-weight:bold;">' + reference + '</p></div>' +
      '<h3 style="color:#1A3C5E;font-size:14px;margin:0 0 12px;">Your enquiry summary</h3>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:26px;">' +
      '<tr><td style="padding:10px 0;border-bottom:1px solid #edf0f2;color:#7b8790;">Property address</td><td style="padding:10px 0;border-bottom:1px solid #edf0f2;color:#1A3C5E;font-weight:bold;">' + address + '</td></tr>' +
      '<tr><td style="padding:10px 0;border-bottom:1px solid #edf0f2;color:#7b8790;">Property type</td><td style="padding:10px 0;border-bottom:1px solid #edf0f2;color:#1A3C5E;font-weight:bold;">' + html(data.propertyType) + '</td></tr>' +
      '<tr><td style="padding:10px 0;border-bottom:1px solid #edf0f2;color:#7b8790;">Size</td><td style="padding:10px 0;border-bottom:1px solid #edf0f2;color:#1A3C5E;font-weight:bold;">' + html(data.size) + '</td></tr>' +
      '<tr><td style="padding:10px 0;color:#7b8790;">Services</td><td style="padding:10px 0;color:#1A3C5E;font-weight:bold;">' + html(data.services) + '</td></tr></table>' +
      '<div style="background:#f4f7fa;padding:20px;border-radius:10px;"><p style="margin:0 0 8px;color:#1A3C5E;font-weight:bold;">What happens next?</p><p style="margin:0;color:#4d5f6d;font-size:14px;line-height:1.6;">We’ll review your information and uploaded photos, call you if we need anything clarified, and provide a clear estimate for the work.</p></div>' +
      '<div style="text-align:center;margin-top:28px;"><p style="margin:0 0 8px;color:#7b8790;font-size:13px;">Need to contact us?</p><p style="margin:0;"><a href="tel:0401857730" style="color:#1A3C5E;font-size:22px;font-weight:bold;text-decoration:none;">' + BUSINESS_PHONE + '</a></p><p style="margin:6px 0 0;"><a href="mailto:' + BUSINESS_EMAIL + '" style="color:#2D7FA6;text-decoration:none;">' + BUSINESS_EMAIL + '</a></p></div>' +
      '</div></div><p style="text-align:center;color:#9aa6ad;font-size:11px;margin-top:16px;">© 2026 Estate Co.</p></div>'
  });
}

function sendReferralBusinessEmail(data, folderUrl, fileLinks) {
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    replyTo: value(data.referrerEmail),
    subject: 'New Estate Co. Referral — ' + (data.referrerName || 'New referral'),
    body:
      'New referral received.\n\n' +
      'Reference: ' + value(data.reference) + '\n\n' +
      'REFERRER\n' +
      'Name: ' + value(data.referrerName) + '\n' +
      'Email: ' + value(data.referrerEmail) + '\n' +
      'Phone: ' + value(data.referrerPhone) + '\n' +
      'Company: ' + value(data.referrerCompany || data.company) + '\n\n' +
      'CLIENT / PROPERTY\n' +
      'Client name: ' + value(data.clientName || data.name) + '\n' +
      'Client phone: ' + value(data.clientPhone || data.phone) + '\n' +
      'Property address: ' + value(data.formattedAddress || data.propertyAddress || data.suburb) + '\n' +
      'Street: ' + value(data.propertyStreet) + '\n' +
      'Suburb: ' + value(data.propertySuburb || data.suburb) + '\n' +
      'State: ' + value(data.propertyState) + '\n' +
      'Postcode: ' + value(data.propertyPostcode) + '\n' +
      'Property type: ' + value(data.propertyType) + '\n' +
      'Likely job type: ' + value(data.services || data.jobType) + '\n\n' +
      'Notes:\n' + value(data.notes) + '\n\n' +
      'Reward terms: $50 for completed jobs $1,000-$4,999; $100 for completed jobs $5,000 and above. Private thank-you gift, not communicated to the client.\n\n' +
      'File names: ' + value(data.fileNames) + '\n' +
      'File links:\n' + (fileLinks && fileLinks.length ? fileLinks.join('\n') : '') + '\n\n' +
      'Drive folder:\n' + value(folderUrl) + '\n\n' +
      (warnings.length ? 'Processing warnings:\n- ' + warnings.join('\n- ') + '\n' : '')
  });
}

function sendReferralReferrerEmail(data) {
  if (!data.referrerEmail) return;
  var firstName = html(value(data.referrerName).split(' ')[0] || 'there');
  var reference = html(data.reference || 'Pending');

  MailApp.sendEmail({
    to: data.referrerEmail,
    replyTo: BUSINESS_EMAIL,
    name: BUSINESS_NAME,
    subject: 'Thanks for your referral — Estate Co.',
    htmlBody:
      '<div style="font-family:Arial,sans-serif;background:#f4f7fa;padding:24px;">' +
      '<div style="max-width:620px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 8px 28px rgba(26,60,94,0.08);">' +
      '<div style="background:#1A3C5E;padding:34px;text-align:center;"><h1 style="color:#ffffff;margin:0;font-size:28px;">Estate Co.</h1><p style="color:#d7e6ef;margin:8px 0 0;font-size:14px;">Professional property clearance referrals</p></div>' +
      '<div style="padding:34px;"><h2 style="color:#1A3C5E;margin:0 0 12px;font-size:22px;">Thank you, ' + firstName + '.</h2>' +
      '<p style="color:#4d5f6d;font-size:15px;line-height:1.7;margin:0 0 24px;">We’ve received your referral and will review the details provided. If the job proceeds and is completed, we’ll arrange the applicable private thank-you gift.</p>' +
      '<div style="background:#eef6fb;border-left:5px solid #2D7FA6;padding:18px 20px;border-radius:8px;margin-bottom:26px;"><p style="margin:0 0 6px;color:#6b7d89;font-size:12px;text-transform:uppercase;letter-spacing:1px;font-weight:bold;">Referral Reference</p><p style="margin:0;color:#1A3C5E;font-size:26px;font-weight:bold;">' + reference + '</p></div>' +
      '<div style="background:#f4f7fa;padding:20px;border-radius:10px;margin-bottom:20px;"><p style="margin:0 0 8px;color:#1A3C5E;font-weight:bold;">Referral thank-you</p><p style="margin:0;color:#4d5f6d;font-size:14px;line-height:1.7;">$50 Prezzee gift card for completed jobs valued $1,000–$4,999.<br>$100 Prezzee gift card for completed jobs valued $5,000 and above.</p></div>' +
      '<p style="color:#7b8790;font-size:13px;line-height:1.6;margin:0;">Referral rewards are issued after the job is completed and paid. The reward is a private thank-you to the referrer and is not communicated to the client.</p>' +
      '<div style="text-align:center;margin-top:28px;"><p style="margin:0 0 8px;color:#7b8790;font-size:13px;">Need to contact us?</p><p style="margin:0;"><a href="tel:0401857730" style="color:#1A3C5E;font-size:22px;font-weight:bold;text-decoration:none;">' + BUSINESS_PHONE + '</a></p><p style="margin:6px 0 0;"><a href="mailto:' + BUSINESS_EMAIL + '" style="color:#2D7FA6;text-decoration:none;">' + BUSINESS_EMAIL + '</a></p></div>' +
      '</div></div><p style="text-align:center;color:#9aa6ad;font-size:11px;margin-top:16px;">© 2026 Estate Co.</p></div>'
  });
}

function getOrCreateFolder(name, parent) {
  var folders = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

function cleanFileName(name) {
  return String(name || '').replace(/[\\\/\?\*\[\]\:\|"]/g, '-').replace(/\s+/g, ' ').trim();
}

function value(v) {
  if (v === null || v === undefined) return '';
  return String(v);
}

function html(v) {
  return value(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run once after pasting to approve permissions and test Sheet1 full-address saving.
function testLead() {
  var mock = { parameter: { payload: JSON.stringify({
    action: 'lead', reference: 'EC-TEST-' + new Date().getTime(),
    name: 'Test User', email: NOTIFY_EMAIL, phone: '0400000000',
    propertyAddress: '123 Swan Street, Richmond VIC 3121', formattedAddress: '123 Swan Street, Richmond VIC 3121, Australia',
    propertyStreet: '123 Swan Street', propertySuburb: 'Richmond', propertyState: 'VIC', propertyPostcode: '3121', propertyCountry: 'Australia', propertyPlaceId: 'TEST_PLACE_ID',
    propertyType: 'House', size: '2–3 bedrooms', services: 'Full clearance, Deep cleaning', notes: 'Test lead submission', fileCount: 0, fileNames: ''
  }) } };
  Logger.log(doPost(mock).getContent());
}

// Run once after pasting to approve permissions and test Sheet2 referral saving.
function testReferralLead() {
  var mock = { parameter: { payload: JSON.stringify({
    action: 'referral', leadType: 'Referral', reference: 'EC-REF-TEST-' + new Date().getTime(),
    referrerName: 'Referral Partner', referrerEmail: NOTIFY_EMAIL, referrerPhone: '0411111111', referrerCompany: 'Sample Agency',
    clientName: 'Client Name', clientPhone: '0422222222',
    propertyAddress: '10 Collins Street, Melbourne VIC 3000', formattedAddress: '10 Collins Street, Melbourne VIC 3000, Australia',
    propertyStreet: '10 Collins Street', propertySuburb: 'Melbourne', propertyState: 'VIC', propertyPostcode: '3000', propertyCountry: 'Australia', propertyPlaceId: 'TEST_PLACE_ID',
    propertyType: 'Apartment', services: 'Deceased estate clearance', notes: 'Test referral submission', fileCount: 0, fileNames: ''
  }) } };
  Logger.log(doPost(mock).getContent());
}
