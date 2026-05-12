// ============================================================
// Clinical Sample Monitor - Billing Module
// Add this as a new script file in Apps Script: billing_module
// ============================================================

// ── Constants ──────────────────────────────────────────────
const BILLING_FOLDER_ID = '1f3OKo8F-iTjsuDgndYJeRJWZDBnq37YV';


// ============================================================
// MAIN: Calculate Billing
// Finds all eligible samples for a client + billing period
// and writes them to the BILLING tab
// ============================================================
function calculateBilling() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Step 1: Select client ──────────────────────────────
  const clientSheet = ss.getSheetByName('CLIENTS');
  const clientData  = clientSheet.getDataRange().getValues();
  let clientList = '';
  const clientMap = {}; // clientId → { name, currency, billingDay, wgsBillingMode }

  for (let i = 1; i < clientData.length; i++) {
    if (clientData[i][0] && clientData[i][13] !== 'No') {
      clientMap[clientData[i][0]] = {
        name:           clientData[i][1],
        currency:       clientData[i][3] || 'HKD',
        billingDay:     clientData[i][4] || 1,
        wesPrice:       clientData[i][10] || 0,
        wgsBillingMode: clientData[i][11] || '3x_wes',
        wgsPrice:       clientData[i][12] || 0,
      };
      clientList += clientData[i][0] + ' | ' + clientData[i][1] + '\n';
    }
  }

  if (!clientList) {
    ui.alert('⚠️ No active clients found. Please add clients first.');
    return;
  }

  const clientResult = ui.prompt(
    '💰 Calculate Billing — Step 1 of 2',
    'Available clients:\n\n' + clientList +
    '\nEnter ClientID to calculate billing for:',
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResult.getSelectedButton() !== ui.Button.OK) return;
  const clientId = clientResult.getResponseText().trim().toUpperCase();

  if (!clientMap[clientId]) {
    ui.alert('❌ ClientID "' + clientId + '" not found.');
    return;
  }
  const client = clientMap[clientId];

  // ── Step 2: Select billing period ─────────────────────
  const today     = new Date();
  const thisYear  = today.getFullYear();
  const thisMonth = today.getMonth(); // 0-based
  // Default: previous month
  const defStart = new Date(thisYear, thisMonth - 1, 1);
  const defEnd   = new Date(thisYear, thisMonth, 0); // last day of prev month

  const periodResult = ui.prompt(
    '💰 Calculate Billing — Step 2 of 2',
    'Client: ' + client.name + '\n\n' +
    'Enter billing period (or press OK for default):\n' +
    'Line 1: Start date (YYYYMMDD) — default: ' + formatDateYYYYMMDD(defStart) + '\n' +
    'Line 2: End date   (YYYYMMDD) — default: ' + formatDateYYYYMMDD(defEnd),
    ui.ButtonSet.OK_CANCEL
  );
  if (periodResult.getSelectedButton() !== ui.Button.OK) return;

  const periodLines = periodResult.getResponseText().split('\n');
  const startStr = (periodLines[0] || '').trim() || formatDateYYYYMMDD(defStart);
  const endStr   = (periodLines[1] || '').trim() || formatDateYYYYMMDD(defEnd);

  const periodStart = parseYYYYMMDD(startStr);
  const periodEnd   = parseYYYYMMDD(endStr);

  if (!periodStart || !periodEnd) {
    ui.alert('❌ Invalid date format. Please use YYYYMMDD (e.g. 20260401).');
    return;
  }

  // ── Find eligible samples ──────────────────────────────
  const masterSheet = ss.getSheetByName('MASTER_SAMPLES');
  const masterData  = masterSheet.getDataRange().getValues();

  const eligibleSamples = [];

  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    if (row[3] !== clientId) continue; // ClientID col (D)

    const notifDateStr = row[19]; // NotificationEmailDate col (T)
    if (!notifDateStr) continue;

    const notifDate = parseYYYYMMDD(String(notifDateStr).replace(/-/g, ''));
    if (!notifDate) continue;

    // Check if notification date falls within billing period
    if (notifDate >= periodStart && notifDate <= periodEnd) {
      eligibleSamples.push({
        sampleId:    row[0],  // SampleID
        labId:       row[1],  // LabID
        projectId:   row[2],  // ProjectID
        serviceType: row[5],  // ServiceType
        pickupDate:  row[9],  // PickupDatetime
        labInDate:   row[10], // LabInDatetime
        notifDate:   notifDateStr,
        condition:   getConditionForSample(row[3], row[5]), // clientId, serviceType
      });
    }
  }

  if (eligibleSamples.length === 0) {
    ui.alert(
      '⚠️ No eligible samples found for:\n' +
      'Client: ' + client.name + '\n' +
      'Period: ' + startStr + ' – ' + endStr + '\n\n' +
      'Make sure samples have "notification_email_sent" status\nand a Notification Email Date within this period.'
    );
    return;
  }

  // ── Sort by NotificationEmailDate then LabID ───────────
  eligibleSamples.sort((a, b) => {
    const dateCompare = String(a.notifDate).localeCompare(String(b.notifDate));
    if (dateCompare !== 0) return dateCompare;
    return String(a.labId).localeCompare(String(b.labId));
  });

  // ── Expand WGS 3xWES into 3 rows ──────────────────────
  const billingRows = [];
  for (const s of eligibleSamples) {
    const isWGS3x = s.serviceType === 'WGS' && s.condition === '3xWES';

    if (isWGS3x) {
      // 1 WGS = 3 billing rows
      for (let n = 1; n <= 3; n++) {
        billingRows.push({
          ...s,
          pdfRowLabel:   s.labId + '_' + n,
          billingUnits:  1,
          billingMode:   '3x_wes',
          unitPrice:     client.wesPrice,
          currency:      client.currency,
          totalAmount:   client.wesPrice,
        });
      }
    } else {
      // WES or WGS independent = 1 row
      const price = s.serviceType === 'WGS' ? client.wgsPrice : client.wesPrice;
      billingRows.push({
        ...s,
        pdfRowLabel:   s.labId,
        billingUnits:  1,
        billingMode:   s.serviceType === 'WGS' ? 'independent' : 'per_unit',
        unitPrice:     price,
        currency:      client.currency,
        totalAmount:   price,
      });
    }
  }

  // ── Write to BILLING tab ───────────────────────────────
  const billingSheet = ss.getSheetByName('BILLING');
  const now = new Date().toISOString();
  const operator = Session.getActiveUser().getEmail();
  let nextId = getNextID('BILLING', 1);

  // Check for existing records in this period (avoid duplicates)
  const existingBilling = billingSheet.getDataRange().getValues();
  const existingKeys = new Set();
  for (let i = 1; i < existingBilling.length; i++) {
    if (existingBilling[i][3] === clientId) { // ClientID
      existingKeys.add(existingBilling[i][4] + '|' + existingBilling[i][7]); // SampleID|PeriodStart
    }
  }

  let newRows = 0;
  for (const br of billingRows) {
    const key = br.sampleId + '|' + startStr;
    if (existingKeys.has(key)) continue; // skip duplicates

    billingSheet.appendRow([
      nextId++,                           // BillingID
      clientId,                           // ClientID
      client.name,                        // ClientName
      br.sampleId,                        // SampleID
      br.labId,                           // LabID
      br.serviceType,                     // ServiceType
      startStr,                           // BillingPeriodStart
      endStr,                             // BillingPeriodEnd
      formatDateYYYYMMDD(br.notifDate),  // NotificationEmailDate
      formatPickupDate(br.pickupDate),    // PickupDate
      formatPickupDate(br.labInDate),     // LabInDate
      br.billingMode,                     // BillingMode
      br.billingUnits,                    // BillingUnits
      br.unitPrice,                       // UnitPrice
      br.currency,                        // Currency
      br.totalAmount,                     // TotalAmount
      br.pdfRowLabel,                     // PDF_RowLabel
      'No',                               // BillingEmailSent
      '',                                 // BillingEmailDate
      '',                                 // InvoicePDFPath
      '',                                 // Remarks
      now                                 // CreatedDate
    ]);
    newRows++;
  }

  // ── Audit log ──────────────────────────────────────────
  writeAuditLog('CALCULATE_BILLING', 'BILLING', clientId,
    '', '', startStr + ' to ' + endStr,
    newRows + ' billing rows created (' + billingRows.length + ' total, ' +
    eligibleSamples.length + ' samples)'
  );

  ui.alert(
    '✅ Billing Calculated!\n\n' +
    '• Client:          ' + client.name + '\n' +
    '• Period:          ' + startStr + ' – ' + endStr + '\n' +
    '• Samples found:   ' + eligibleSamples.length + '\n' +
    '• Billing rows:    ' + billingRows.length + '\n' +
    '• New rows added:  ' + newRows + '\n\n' +
    'Check the BILLING tab, then click\n"Generate Billing PDF" to create the PDF.'
  );
}


// ============================================================
// MAIN: Generate Billing PDF
// Creates a PDF from BILLING tab data and saves to Drive
// ============================================================
function generateBillingPDF() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Select client ──────────────────────────────────────
  const clientSheet = ss.getSheetByName('CLIENTS');
  const clientData  = clientSheet.getDataRange().getValues();
  let clientList = '';
  const clientMap = {};
  for (let i = 1; i < clientData.length; i++) {
    if (clientData[i][0]) {
      clientMap[clientData[i][0]] = clientData[i][1];
      clientList += clientData[i][0] + ' | ' + clientData[i][1] + '\n';
    }
  }

  const clientResult = ui.prompt(
    '📄 Generate Billing PDF — Step 1 of 2',
    'Available clients:\n\n' + clientList +
    '\nEnter ClientID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResult.getSelectedButton() !== ui.Button.OK) return;
  const clientId   = clientResult.getResponseText().trim().toUpperCase();
  const clientName = clientMap[clientId];
  if (!clientName) {
    ui.alert('❌ ClientID "' + clientId + '" not found.');
    return;
  }

  // ── Select billing period ──────────────────────────────
  const periodResult = ui.prompt(
    '📄 Generate Billing PDF — Step 2 of 2',
    'Enter billing period:\n' +
    'Line 1: Start date (YYYYMMDD)\n' +
    'Line 2: End date   (YYYYMMDD)',
    ui.ButtonSet.OK_CANCEL
  );
  if (periodResult.getSelectedButton() !== ui.Button.OK) return;
  const pLines   = periodResult.getResponseText().split('\n');
  const startStr = (pLines[0] || '').trim();
  const endStr   = (pLines[1] || '').trim();

  // ── Get billing rows from BILLING tab ─────────────────
  const billingSheet = ss.getSheetByName('BILLING');
  const billingData  = billingSheet.getDataRange().getValues();

  const rows = [];
  for (let i = 1; i < billingData.length; i++) {
    const r = billingData[i];
    if (r[1] === clientId &&
        String(r[6]) === startStr &&
        String(r[7]) === endStr) {
      rows.push({
        pdfRowLabel: r[16],  // PDF_RowLabel
        pickupDate:  r[9],   // PickupDate
        labInDate:   r[10],  // LabInDate
        notifDate:   r[8],   // NotificationEmailDate
      });
    }
  }

  if (rows.length === 0) {
    ui.alert(
      '⚠️ No billing records found for:\n' +
      'Client: ' + clientName + '\n' +
      'Period: ' + startStr + ' – ' + endStr + '\n\n' +
      'Please run "Calculate Billing" first.'
    );
    return;
  }

  // ── Create Google Doc ──────────────────────────────────
  const docTitle   = clientId + '_' + startStr + '_' + endStr;
  const doc        = DocumentApp.create(docTitle);
  const body       = doc.getBody();
  const docStyle   = {};
  docStyle[DocumentApp.Attribute.MARGIN_TOP]    = 36;
  docStyle[DocumentApp.Attribute.MARGIN_BOTTOM] = 36;
  docStyle[DocumentApp.Attribute.MARGIN_LEFT]   = 54;
  docStyle[DocumentApp.Attribute.MARGIN_RIGHT]  = 54;
  body.setAttributes(docStyle);

  // ── Add page number footer "Page X of Y" ──────────────
  const footer = doc.addFooter();
  const footerPara = footer.appendParagraph('');
  footerPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  footerPara.appendText('Page ');
  footerPara.appendInlineImage; // placeholder
  // Use page number fields
  const pageNumEl = footerPara.appendText('');
  doc.getFooter().appendParagraph('').setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  // Clear auto-created paragraph and build footer properly
  footer.clear();
  const fp = footer.appendParagraph('');
  fp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  fp.appendText('Page ');
  fp.appendPageNumber();
  fp.appendText(' of ');
  fp.appendPageCount();

  // ── Build table ────────────────────────────────────────
  // Header row
  const tableData = [['#', 'LabID', 'Sample Pickup Date', 'LabIn Date', 'Notification Email Date']];

  for (let i = 0; i < rows.length; i++) {
    tableData.push([
      String(i + 1),
      rows[i].pdfRowLabel,
      String(rows[i].pickupDate || ''),
      String(rows[i].labInDate  || ''),
      String(rows[i].notifDate  || ''),
    ]);
  }

  const table = body.appendTable(tableData);

  // ── Style the table ────────────────────────────────────
  // Column widths (in points): #=30, LabID=160, PickupDate=100, LabInDate=100, NotifDate=120
  const colWidths = [30, 180, 100, 100, 120];
  for (let col = 0; col < colWidths.length; col++) {
    for (let row = 0; row < tableData.length; row++) {
      table.getCell(row, col).setWidth(colWidths[col]);
    }
  }

  // Header row style
  const headerRow = table.getRow(0);
  for (let col = 0; col < 5; col++) {
    const cell = headerRow.getCell(col);
    cell.setBackgroundColor('#000000');
    const para = cell.getParagraphs()[0];
    para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const text = para.editAsText();
    text.setForegroundColor('#FFFFFF');
    text.setBold(true);
    text.setFontSize(9);
  }

  // Data rows style
  for (let row = 1; row < tableData.length; row++) {
    const bgColor = row % 2 === 0 ? '#F8F8F8' : '#FFFFFF';
    for (let col = 0; col < 5; col++) {
      const cell = table.getCell(row, col);
      cell.setBackgroundColor(bgColor);
      const para = cell.getParagraphs()[0];
      // Align # column center, rest left
      para.setAlignment(col === 0
        ? DocumentApp.HorizontalAlignment.CENTER
        : DocumentApp.HorizontalAlignment.LEFT);
      const text = para.editAsText();
      text.setFontSize(8);
      text.setBold(false);
    }
  }

  // ── Save and export as PDF ─────────────────────────────
  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(doc.getId())
    .getAs(MimeType.PDF)
    .setName(docTitle + '.pdf');

  // ── Save to billing Drive folder ───────────────────────
  let billingFolder;
  try {
    billingFolder = DriveApp.getFolderById(BILLING_FOLDER_ID);
  } catch (e) {
    ui.alert('❌ Cannot access billing folder.\nPlease check the folder ID or sharing permissions.');
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return;
  }

  // Check if file already exists — replace if so
  const existingFiles = billingFolder.getFilesByName(docTitle + '.pdf');
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  const savedFile = billingFolder.createFile(pdfBlob);
  const fileUrl   = savedFile.getUrl();

  // ── Delete temp Google Doc ─────────────────────────────
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  // ── Update BILLING tab with PDF path ──────────────────
  const billingRows2 = billingSheet.getDataRange().getValues();
  for (let i = 1; i < billingRows2.length; i++) {
    if (billingRows2[i][1] === clientId &&
        String(billingRows2[i][6]) === startStr &&
        String(billingRows2[i][7]) === endStr) {
      billingSheet.getRange(i + 1, 20).setValue(fileUrl); // InvoicePDFPath col
    }
  }

  // ── Audit log ──────────────────────────────────────────
  writeAuditLog('GENERATE_PDF', 'BILLING', clientId,
    '', '', docTitle + '.pdf', rows.length + ' rows in PDF');

  ui.alert(
    '🎉 Billing PDF Generated!\n\n' +
    '• File: ' + docTitle + '.pdf\n' +
    '• Rows: ' + rows.length + '\n' +
    '• Saved to your billing Drive folder\n\n' +
    'URL:\n' + fileUrl
  );
}


// ============================================================
// HELPER: Get WGS billing condition for a sample
// ============================================================
function getConditionForSample(clientId, serviceType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CLIENT_PREFIXES');
  if (!sheet) return 'Standard';

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clientId && data[i][2] === serviceType) {
      return data[i][3]; // Condition column
    }
  }
  return 'Standard';
}


// ============================================================
// HELPER: Format date as YYYYMMDD string
// ============================================================
function formatDateYYYYMMDD(date) {
  if (!date) return '';
  if (typeof date === 'string') {
    // Already formatted — strip dashes
    return date.replace(/-/g, '').substring(0, 8);
  }
  const d = new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return y + m + day;
}


// ============================================================
// HELPER: Parse YYYYMMDD string to Date object
// ============================================================
function parseYYYYMMDD(str) {
  if (!str) return null;
  const s = String(str).replace(/-/g, '').trim();
  if (s.length !== 8) return null;
  const y = parseInt(s.substring(0, 4));
  const m = parseInt(s.substring(4, 6)) - 1;
  const d = parseInt(s.substring(6, 8));
  const date = new Date(y, m, d);
  if (isNaN(date.getTime())) return null;
  return date;
}


// ============================================================
// HELPER: Format pickup/labIn date to YYYYMMDD
// Handles datetime strings like "2026-03-19 09:00"
// ============================================================
function formatPickupDate(val) {
  if (!val) return '';
  const s = String(val).replace(/-/g, '').replace(/\s.*/g, '').trim();
  return s.substring(0, 8);
}


// ============================================================
// HELPER: Get billing summary for dashboard
// ============================================================
function getBillingSummary(clientId, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BILLING');
  const data  = sheet.getDataRange().getValues();

  let totalUnits = 0;
  let totalAmount = 0;
  let rowCount = 0;

  const periodStart = year + String(month).padStart(2, '0') + '01';

  for (let i = 1; i < data.length; i++) {
    if ((!clientId || data[i][1] === clientId) &&
        String(data[i][6]).startsWith(periodStart.substring(0, 6))) {
      totalUnits  += parseFloat(data[i][12]) || 0;
      totalAmount += parseFloat(data[i][15]) || 0;
      rowCount++;
    }
  }

  return { totalUnits, totalAmount, rowCount };
}
