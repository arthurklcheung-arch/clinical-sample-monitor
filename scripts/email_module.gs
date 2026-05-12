// ============================================================
// Clinical Sample Monitor - Email Notification Module
// Add this as a new script file in Apps Script: email_module
// ============================================================

// ── Drive Folder IDs ───────────────────────────────────────
// Update these after running setupSubmissionFormFolders()
const FORMS_PENDING_FOLDER_ID = 'PENDING_FOLDER_ID_HERE';
const FORMS_SENT_FOLDER_ID    = 'SENT_FOLDER_ID_HERE';


// ============================================================
// SETUP: Update CLIENTS tab with new contact + email columns
// Run ONCE — safely adds new columns without deleting existing
// ============================================================
function upgradeClientsTabForEmail() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CLIENTS');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastCol = headers.length;

  // New columns to add (only if not already present)
  const newCols = [
    // Contact persons
    'Contact_QC',
    'Contact_DataReady',
    'Contact_Billing',
    'Contact_Recollection',
    'Contact_Deletion',
    // Email addresses
    'Email_QC',
    'Email_DataReady',
    'Email_Billing',
    'Email_Recollection',
    'Email_Deletion',
    // CC addresses (optional per email type)
    'CC_QC',
    'CC_DataReady',
    'CC_Billing',
    'CC_Recollection',
    'CC_Deletion',
    // Subject templates ({DATE} = collection date, {CLIENT} = client name)
    'EmailSubject_QC',
    'EmailSubject_DataReady',
    'EmailSubject_Billing',
    'EmailSubject_Recollection',
    'EmailSubject_Deletion',
  ];

  let added = 0;
  for (const colName of newCols) {
    if (!headers.includes(colName)) {
      const nextCol = lastCol + added + 1;
      sheet.getRange(1, nextCol).setValue(colName);
      sheet.getRange(1, nextCol)
        .setBackground('#1a73e8')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      added++;
    }
  }

  sheet.autoResizeColumns(1, sheet.getLastColumn());

  ui.alert(
    '✅ CLIENTS tab upgraded!\n\n' +
    added + ' new columns added:\n' +
    '• Contact_QC / Email_QC / CC_QC\n' +
    '• Contact_DataReady / Email_DataReady\n' +
    '• Contact_Billing / Email_Billing\n' +
    '• Contact_Recollection / Email_Recollection\n' +
    '• Contact_Deletion / Email_Deletion\n' +
    '• EmailSubject_* for each type\n\n' +
    'Please fill in the contact names and\n' +
    'email subject templates for each client.\n\n' +
    'Example subject template:\n' +
    'EmailSubject_QC:\n' +
    'Prenetics x HKCH - WES/WGS/Twist Exon - {DATE} Batch'
  );
}


// ============================================================
// SETUP: Create Google Drive folder structure for forms
// Run ONCE — creates Pending/ and Sent/ folders
// ============================================================
function setupSubmissionFormFolders() {
  const ui = SpreadsheetApp.getUi();

  const rootResult = ui.prompt(
    '📁 Setup Submission Form Folders',
    'Enter the Google Drive folder ID where you want to\n' +
    'create the Submission Forms folders:\n\n' +
    '(Find it in the URL of your Drive folder:\n' +
    'drive.google.com/drive/folders/FOLDER_ID_HERE)\n\n' +
    'Or press OK to create in My Drive root:',
    ui.ButtonSet.OK_CANCEL
  );
  if (rootResult.getSelectedButton() !== ui.Button.OK) return;

  const parentId = rootResult.getResponseText().trim();
  let parent;
  try {
    parent = parentId ? DriveApp.getFolderById(parentId) : DriveApp.getRootFolder();
  } catch (e) {
    ui.alert('❌ Cannot access folder. Please check the folder ID.');
    return;
  }

  // Create main folder
  let mainFolder;
  const existing = parent.getFoldersByName('Submission Forms');
  if (existing.hasNext()) {
    mainFolder = existing.next();
  } else {
    mainFolder = parent.createFolder('Submission Forms');
  }

  // Create Pending subfolder
  let pendingFolder;
  const existingPending = mainFolder.getFoldersByName('Pending');
  if (existingPending.hasNext()) {
    pendingFolder = existingPending.next();
  } else {
    pendingFolder = mainFolder.createFolder('Pending');
  }

  // Create Sent subfolder
  let sentFolder;
  const existingSent = mainFolder.getFoldersByName('Sent');
  if (existingSent.hasNext()) {
    sentFolder = existingSent.next();
  } else {
    sentFolder = mainFolder.createFolder('Sent');
  }

  const pendingId = pendingFolder.getId();
  const sentId    = sentFolder.getId();

  ui.alert(
    '✅ Folders created!\n\n' +
    '📁 Submission Forms/\n' +
    '  ├── 📁 Pending  (drop scanned forms here)\n' +
    '  └── 📁 Sent     (system moves here after email)\n\n' +
    'IMPORTANT: Copy these IDs into the\n' +
    'email_module.gs script at the top:\n\n' +
    'FORMS_PENDING_FOLDER_ID = \n"' + pendingId + '"\n\n' +
    'FORMS_SENT_FOLDER_ID = \n"' + sentId + '"'
  );
}


// ============================================================
// HELPER: Get client email config
// ============================================================
function getClientEmailConfig(clientId, emailType) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName('CLIENTS');
  const data   = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indices
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clientId) {
      return {
        clientName:  data[i][1],
        contact:     data[i][idx['Contact_' + emailType]]     || '',
        email:       data[i][idx['Email_' + emailType]]       || '',
        cc:          data[i][idx['CC_' + emailType]]          || '',
        subject:     data[i][idx['EmailSubject_' + emailType]]|| '',
      };
    }
  }
  return null;
}


// ============================================================
// HELPER: Find submission form in Pending folder
// Looks for file matching: ClientID_YYYYMMDD
// ============================================================
function findSubmissionForm(clientId, dateStr) {
  if (FORMS_PENDING_FOLDER_ID === 'PENDING_FOLDER_ID_HERE') return null;
  try {
    const folder    = DriveApp.getFolderById(FORMS_PENDING_FOLDER_ID);
    const searchName = clientId + '_' + dateStr;
    const files      = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName().replace(/\.[^/.]+$/, ''); // strip extension
      if (name === searchName) return file;
    }
    return null;
  } catch (e) {
    return null;
  }
}


// ============================================================
// HELPER: Move form from Pending → Sent after email sent
// ============================================================
function moveFormToSent(file) {
  if (FORMS_SENT_FOLDER_ID === 'SENT_FOLDER_ID_HERE') return;
  try {
    const sentFolder = DriveApp.getFolderById(FORMS_SENT_FOLDER_ID);
    file.moveTo(sentFolder);
  } catch (e) {
    Logger.log('Could not move form to Sent folder: ' + e);
  }
}


// ============================================================
// HELPER: Record email sent in STATUS_TRACKING + AUDIT_LOG
// ============================================================
function recordEmailSent(sampleIds, emailType, statusCode, operator) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('MASTER_SAMPLES');
  const statusSheet = ss.getSheetByName('STATUS_TRACKING');
  const masterData  = masterSheet.getDataRange().getValues();
  const now         = new Date().toISOString();

  const labelMap = {
    'QC':           'qc_passed_email_sent',
    'DataReady':    'notification_email_sent',
    'Billing':      'billing_email_sent',
    'Recollection': 'sample_received_by_lab', // stays in lab
    'Deletion':     'notification_email_sent', // reuse
  };

  const status = statusCode || labelMap[emailType] || emailType;

  for (const sampleId of sampleIds) {
    // Update MASTER_SAMPLES current status
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][0] === sampleId) {
        const rowNum = i + 1;
        masterSheet.getRange(rowNum, 18).setValue(status);
        masterSheet.getRange(rowNum, 19).setValue(now);
        if (emailType === 'DataReady') {
          masterSheet.getRange(rowNum, 20).setValue(now); // NotificationEmailDate
        }
        break;
      }
    }

    // Append to STATUS_TRACKING
    const trackId = getNextID('STATUS_TRACKING', 1);
    statusSheet.appendRow([
      trackId, sampleId, '', '', '',
      status, 'Email sent: ' + emailType,
      now, operator, emailType + ' email sent'
    ]);
  }
}


// ============================================================
// EMAIL 1: QC Passed Confirmation
// ============================================================
function sendQCConfirmationEmail() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Step 1: Select client ──────────────────────────────
  const clientSheet = ss.getSheetByName('CLIENTS');
  const clientData  = clientSheet.getDataRange().getValues();
  let clientList = '';
  for (let i = 1; i < clientData.length; i++) {
    if (clientData[i][0]) clientList += clientData[i][0] + ' | ' + clientData[i][1] + '\n';
  }

  const clientResult = ui.prompt(
    '📧 QC Confirmation Email — Step 1 of 3',
    'Available clients:\n\n' + clientList +
    '\nEnter ClientID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResult.getSelectedButton() !== ui.Button.OK) return;
  const clientId = clientResult.getResponseText().trim().toUpperCase();

  const config = getClientEmailConfig(clientId, 'QC');
  if (!config) {
    ui.alert('❌ Client "' + clientId + '" not found.');
    return;
  }

  // ── Step 2: Enter collection date (batch date) ─────────
  const dateResult = ui.prompt(
    '📧 QC Confirmation Email — Step 2 of 3',
    'Enter collection/batch date (YYYYMMDD):\n' +
    'This will appear in the subject line.',
    ui.ButtonSet.OK_CANCEL
  );
  if (dateResult.getSelectedButton() !== ui.Button.OK) return;
  const batchDate = dateResult.getResponseText().trim();

  // ── Step 3: Get samples for this batch ────────────────
  const masterSheet = ss.getSheetByName('MASTER_SAMPLES');
  const masterData  = masterSheet.getDataRange().getValues();

  // Find samples for this client with this pickup/lab-in date
  const batchSamples = [];
  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    if (row[3] !== clientId) continue;
    const labInDate = formatDateYYYYMMDD(row[10]); // LabInDatetime
    if (labInDate === batchDate) {
      batchSamples.push({ sampleId: row[0], labId: row[1] });
    }
  }

  if (batchSamples.length === 0) {
    ui.alert(
      '⚠️ No samples found for:\n' +
      'Client: ' + config.clientName + '\n' +
      'Date: ' + batchDate + '\n\n' +
      'Check that LabInDatetime matches this date\nin MASTER_SAMPLES.'
    );
    return;
  }

  // ── Build subject ──────────────────────────────────────
  let subject = config.subject || 'Prenetics x ' + clientId + ' - ' + batchDate + ' Batch';
  subject = subject.replace('{DATE}', batchDate).replace('{CLIENT}', config.clientName);

  // ── Build email body ───────────────────────────────────
  const sampleList = batchSamples.map(s => s.labId).join('\n');
  const contactName = config.contact || 'Sir/Madam';
  const body =
    'Dear ' + contactName + ',\n\n' +
    'The following samples have been received and accepted:\n\n' +
    sampleList + '\n\n' +
    'Regards,\n' +
    'Prenetics Lab Team';

  // ── Check for signed submission form ──────────────────
  const formFile = findSubmissionForm(clientId, batchDate);
  const formFound = formFile !== null;

  // ── Confirm before sending ─────────────────────────────
  const confirm = ui.alert(
    '📧 Confirm QC Email',
    'To:      ' + (config.email || '(no email set — preview only)') + '\n' +
    'CC:      ' + (config.cc || 'none') + '\n' +
    'Subject: ' + subject + '\n\n' +
    'Samples: ' + batchSamples.length + '\n' +
    'Form attached: ' + (formFound ? '✅ ' + formFile.getName() : '⚠️ NOT FOUND') + '\n\n' +
    (formFound ? '' : '⚠️ No submission form found in Pending folder\nfor ' + clientId + '_' + batchDate + '\n\n') +
    'Click OK to send.',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // ── Send email ─────────────────────────────────────────
  const operator = Session.getActiveUser().getEmail();

  if (!config.email) {
    // Preview mode — no email address set yet
    ui.alert(
      '📋 Preview Mode (no email address set)\n\n' +
      'Subject: ' + subject + '\n\n' +
      body + '\n\n' +
      'To send for real: add Email_QC in CLIENTS tab.'
    );
    return;
  }

  try {
    const mailOptions = {
      name: 'Prenetics Lab Team',
      cc:   config.cc || '',
      body: body,
    };

    // Attach signed submission form if found
    if (formFound) {
      mailOptions.attachments = [formFile.getAs(MimeType.PDF)];
    }

    GmailApp.sendEmail(config.email, subject, body, mailOptions);

    // Move form to Sent folder
    if (formFound) moveFormToSent(formFile);

    // Record in STATUS_TRACKING + AUDIT_LOG
    recordEmailSent(
      batchSamples.map(s => s.sampleId),
      'QC', 'qc_passed_email_sent', operator
    );
    writeAuditLog('SEND_EMAIL_QC', 'MASTER_SAMPLES', clientId,
      '', '', subject,
      batchSamples.length + ' samples, form: ' + (formFound ? formFile.getName() : 'none')
    );

    ui.alert(
      '✅ QC Confirmation Email Sent!\n\n' +
      '• To:      ' + config.email + '\n' +
      '• Samples: ' + batchSamples.length + '\n' +
      '• Form:    ' + (formFound ? '✅ Attached & moved to Sent/' : 'Not attached') + '\n' +
      '• Status updated in MASTER_SAMPLES'
    );

  } catch (e) {
    ui.alert('❌ Email failed to send:\n' + e.toString());
  }
}


// ============================================================
// EMAIL TEMPLATE MANAGER
// View and edit email templates per client
// ============================================================
function viewEmailTemplates() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const clientResult = ui.prompt(
    '📋 View Email Templates',
    'Enter ClientID to view templates:',
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResult.getSelectedButton() !== ui.Button.OK) return;
  const clientId = clientResult.getResponseText().trim().toUpperCase();

  const emailTypes = ['QC', 'DataReady', 'Billing', 'Recollection', 'Deletion'];
  const emailLabels = {
    'QC':           '1. QC Passed Confirmation',
    'DataReady':    '2. Data Ready Notification',
    'Billing':      '3. Billing Email',
    'Recollection': '4. Re-collection Request',
    'Deletion':     '5. Data Deletion Notification',
  };

  let summary = 'Email config for ' + clientId + ':\n\n';
  for (const type of emailTypes) {
    const config = getClientEmailConfig(clientId, type);
    if (config) {
      summary += emailLabels[type] + '\n';
      summary += '  Contact: ' + (config.contact || '(not set)') + '\n';
      summary += '  Email:   ' + (config.email   || '(not set)') + '\n';
      summary += '  Subject: ' + (config.subject  || '(not set)') + '\n\n';
    }
  }

  ui.alert(summary);
}


// ============================================================
// UPDATED MENU
// ============================================================
function setupMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧬 Lab System')
    .addItem('➕ Register New Sample',     'registerNewSample')
    .addItem('🔄 Update Sample Status',    'updateSampleStatus')
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 Send Emails')
      .addItem('✅ QC Confirmation',         'sendQCConfirmationEmail')
      .addItem('📤 Data Ready Notification', 'sendDataReadyEmail')
      .addItem('🗑️ Data Deletion Notice',    'sendDataDeletionEmail')
      .addItem('🔁 Re-collection Request',   'sendRecollectionEmail')
      .addItem('💰 Billing Email',           'sendBillingEmail')
    )
    .addSeparator()
    .addItem('💰 Calculate Billing',        'calculateBilling')
    .addItem('📄 Generate Billing PDF',     'generateBillingPDF')
    .addSeparator()
    .addItem('📊 Refresh Dashboard',        'refreshDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Admin Tools')
      .addItem('👥 Add New Client',           'addNewClient')
      .addItem('🔗 Sync Client Prefixes',     'syncExistingClientPrefixes')
      .addItem('✏️ Rename Client',            'renameClient')
      .addItem('✏️ Rename Project',           'renameProject')
      .addItem('🔧 Upgrade CLIENT_PREFIXES',  'upgradeClientPrefixSheet')
      .addItem('📧 Setup Email Columns',      'upgradeClientsTabForEmail')
      .addItem('📁 Setup Form Folders',       'setupSubmissionFormFolders')
      .addItem('📋 View Email Templates',     'viewEmailTemplates')
    )
    .addSeparator()
    .addItem('💾 Backup Now',               'backupDatabase')
    .addToUi();
}

// Placeholder functions for remaining email types
function sendDataReadyEmail() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Data Ready Notification\nPlease provide the email template.');
}
function sendDataDeletionEmail() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Data Deletion Notification\nPlease provide the email template.');
}
function sendRecollectionEmail() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Re-collection Request\nPlease provide the email template.');
}
function sendBillingEmail() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Billing Email\nPlease provide the email template.');
}
