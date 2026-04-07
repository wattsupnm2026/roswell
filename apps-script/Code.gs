// =============================================================================
//  WattsUp NM — Google Apps Script Web App
//  Spreadsheet ID: 1t7nSPtKn1v-rKW3hYcfXbq6REjdZ-14ONctWVUCPT3E
//
//  SETUP:
//  1. Open script.google.com, create a new project linked to the spreadsheet.
//  2. Paste this code.
//  3. Enable the Drive API under Services (Advanced Google Services).
//  4. Deploy as a Web App (Execute as: Me, Access: Anyone).
//  5. Copy the deployment URL into index.html GOOGLE_SCRIPT_URL constant.
// =============================================================================

var SPREADSHEET_ID = '1t7nSPtKn1v-rKW3hYcfXbq6REjdZ-14ONctWVUCPT3E';
var DRIVE_FOLDER_NAME = 'WattsUp Bill Uploads';

// ---------------------------------------------------------------------------
//  doPost — Web App Endpoint
// ---------------------------------------------------------------------------
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var fileBase64 = params.fileBase64 || null;
    var fileName   = params.billFileName || ('upload_' + Date.now());
    var mimeType   = params.fileMimeType || 'application/octet-stream';

    // ── Upload bill file to Drive ──────────────────────────────────────────
    var driveUrl = '';
    var ocrText  = '';
    if (fileBase64) {
      var folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
      var blob   = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, fileName);
      var file   = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveUrl = file.getUrl();

      // ── OCR via Drive API ────────────────────────────────────────────────
      try {
        var resource = { title: fileName + '_ocr', mimeType: 'application/vnd.google-apps.document' };
        var options  = { ocr: true, ocrLanguage: 'en' };
        var ocrFile  = Drive.Files.insert(resource, blob, options);
        var docId    = ocrFile.id;
        var doc      = DocumentApp.openById(docId);
        ocrText      = doc.getBody().getText();
        // Clean up the temporary OCR doc
        DriveApp.getFileById(docId).setTrashed(true);
      } catch (ocrErr) {
        ocrText = 'OCR_ERROR: ' + ocrErr.message;
      }
    }

    // ── Parse OCR text ─────────────────────────────────────────────────────
    var parsed = parseXcelBill(ocrText);

    // ── RHS Flag ───────────────────────────────────────────────────────────
    var rhsFlag = (parsed.rateType && /rhs/i.test(parsed.rateType))
      ? 'YES — INELIGIBLE'
      : 'NO — OK';

    // ── Write row to spreadsheet ───────────────────────────────────────────
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheets()[0];

    // Ensure header row exists
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(getHeaderRow());
    }

    sheet.appendRow([
      new Date(),                                          // 1  Timestamp
      params.fullName        || '',                        // 2  Full Name
      params.email           || '',                        // 3  Email
      params.phone           || '',                        // 4  Phone
      params.eligibleProgram || '',                        // 5  Eligible Program
      params.nameOnBill      || '',                        // 6  Name on Bill
      params.authorizedUser === true ? 'Yes' : (params.authorizedUser === false ? 'No' : ''),  // 7  Authorized User
      fileName,                                            // 8  Bill File Name
      driveUrl,                                            // 9  Bill File URL
      ocrText,                                             // 10 OCR Raw Text
      parsed.accountNumber   || '',                        // 11 Account Number
      parsed.dueDate         || '',                        // 12 Due Date
      parsed.statementNumber || '',                        // 13 Statement Number
      parsed.statementDate   || '',                        // 14 Statement Date
      parsed.amountDue       || '',                        // 15 Amount Due
      parsed.serviceAddress  || '',                        // 16 Service Address (Full)
      parsed.serviceName     || '',                        // 17 Service Name
      parsed.serviceStreet   || '',                        // 18 Service Street
      parsed.serviceCityStateZip || '',                    // 19 Service City State Zip
      parsed.nextReadDate    || '',                        // 20 Next Read Date
      parsed.premisesNumber  || '',                        // 21 Premises Number
      parsed.invoiceNumber   || '',                        // 22 Invoice Number
      parsed.currentReading  || '',                        // 23 Current Reading
      parsed.previousReading || '',                        // 24 Previous Reading
      parsed.usage           || '',                        // 25 Usage (kWh)
      parsed.rateType        || '',                        // 26 Rate Type
      parsed.rateCharge      || '',                        // 27 Rate Charge
      rhsFlag,                                             // 28 RHS Flag
      'New'                                                // 29 Status
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------------------------------------------------------------------------
//  getOrCreateFolder — find or create a Drive folder by name
// ---------------------------------------------------------------------------
function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ---------------------------------------------------------------------------
//  getHeaderRow — returns the header row array
// ---------------------------------------------------------------------------
function getHeaderRow() {
  return [
    'Timestamp', 'Full Name', 'Email', 'Phone', 'Eligible Program',
    'Name on Bill', 'Authorized User', 'Bill File Name', 'Bill File URL',
    'OCR Raw Text', 'Account Number', 'Due Date', 'Statement Number',
    'Statement Date', 'Amount Due', 'Service Address (Full)', 'Service Name',
    'Service Street', 'Service City State Zip', 'Next Read Date',
    'Premises Number', 'Invoice Number', 'Current Reading', 'Previous Reading',
    'Usage (kWh)', 'Rate Type', 'Rate Charge', 'RHS Flag', 'Status'
  ];
}

// ---------------------------------------------------------------------------
//  parseXcelBill — regex extraction from OCR text
// ---------------------------------------------------------------------------
function parseXcelBill(text) {
  if (!text) return {};

  function find(pattern) {
    var m = text.match(pattern);
    return m ? m[1].trim() : '';
  }

  // Account Number  e.g. "54-0010223974-2" or "Account Number: 54-0010223974-2"
  var accountNumber = find(/Account\s*(?:Number|#|No\.?)[\s:]*([0-9\-]+)/i)
    || find(/(\d{2}-\d{10}-\d)/);

  // Due Date
  var dueDate = find(/(?:Payment\s*Due|Amount\s*Due\s*(?:by|on)|Due\s*Date)[\s:]*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i);

  // Statement Number
  var statementNumber = find(/Statement\s*(?:Number|No\.?|#)[\s:]*([0-9]+)/i);

  // Statement Date
  var statementDate = find(/(?:Statement|Bill)\s*Date[\s:]*([A-Za-z]+ \d{1,2},?\s*\d{4}|\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i)
    || find(/(?:Billing Period Ends|Bill Issued)[\s:]*([A-Za-z]+ \d{1,2},?\s*\d{4}|\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/i);

  // Amount Due
  var amountDue = find(/(?:Total\s*Amount\s*Due|Amount\s*Due|Please\s*Pay)[\s:]*\$?([\d,]+\.\d{2})/i)
    || find(/Amount Due[\s:]*\$?([\d,]+\.\d{2})/i);

  // Service address block — look for a multiline address
  var serviceAddress = '';
  var serviceName    = '';
  var serviceStreet  = '';
  var serviceCityStateZip = '';

  var addrBlock = text.match(/Service\s*(?:Address|Location)[\s:\n]+([A-Z ]+\n[A-Z0-9 ]+\n[A-Z ,0-9\-]+)/i);
  if (addrBlock) {
    var lines = addrBlock[1].trim().split('\n');
    serviceName         = (lines[0] || '').trim();
    serviceStreet       = (lines[1] || '').trim();
    serviceCityStateZip = (lines[2] || '').trim();
    serviceAddress      = [serviceName, serviceStreet, serviceCityStateZip].filter(Boolean).join(', ');
  }
  if (!serviceAddress) {
    // Fallback: try to grab name + address from bill header
    var header = text.match(/^([\w ]+)\n([\w .]+)\n([\w ,\-]+\d{5}(?:-\d{4})?)/m);
    if (header) {
      serviceName         = (header[1] || '').trim();
      serviceStreet       = (header[2] || '').trim();
      serviceCityStateZip = (header[3] || '').trim();
      serviceAddress      = [serviceName, serviceStreet, serviceCityStateZip].filter(Boolean).join(', ');
    }
  }

  // Next Read Date
  var nextReadDate = find(/Next\s*(?:Read|Meter Read)\s*(?:Date)?[\s:]*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i);

  // Premises Number
  var premisesNumber = find(/Premises\s*(?:Number|No\.?|#)[\s:]*([0-9]+)/i);

  // Invoice Number
  var invoiceNumber = find(/Invoice\s*(?:Number|No\.?|#)[\s:]*([0-9]+)/i);

  // Meter readings
  var currentReading  = find(/Current\s*(?:Reading|Read)[\s:]*([0-9,]+)/i);
  var previousReading = find(/Previous\s*(?:Reading|Read)[\s:]*([0-9,]+)/i);

  // Usage
  var usage = find(/(?:Usage|kWh Used|kWh)[\s:]*([0-9,]+)\s*kWh/i)
    || find(/([0-9,]+)\s*kWh/i);

  // Rate Type — e.g. "RATE: RHS Res Htg Svc" or "Rate Schedule: RHS"
  var rateType = find(/(?:RATE|Rate\s*(?:Schedule|Type|Code))[\s:]+([A-Za-z0-9 ]+?)(?:\n|$)/i);

  // Rate Charge
  var rateCharge = find(/(?:Rate\s*Charge|Basic\s*Service\s*Charge|Customer\s*Charge)[\s:]*\$?([\d,]+\.\d{2})/i);

  return {
    accountNumber:      accountNumber,
    dueDate:            dueDate,
    statementNumber:    statementNumber,
    statementDate:      statementDate,
    amountDue:          amountDue,
    serviceAddress:     serviceAddress,
    serviceName:        serviceName,
    serviceStreet:      serviceStreet,
    serviceCityStateZip:serviceCityStateZip,
    nextReadDate:       nextReadDate,
    premisesNumber:     premisesNumber,
    invoiceNumber:      invoiceNumber,
    currentReading:     currentReading,
    previousReading:    previousReading,
    usage:              usage,
    rateType:           rateType,
    rateCharge:         rateCharge
  };
}
