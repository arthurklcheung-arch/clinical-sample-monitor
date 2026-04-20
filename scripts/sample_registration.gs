// ============================================================
// Clinical Sample Monitor - Sample Registration Module
// Add this code to your Apps Script editor (append to existing code)
// ============================================================


// ============================================================
// STEP 1: Add CLIENT_PREFIXES sheet
// Run this ONCE to add the prefix config sheet
// ============================================================
function setupClientPrefixSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('CLIENT_PREFIXES');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('CLIENT_PREFIXES');
  sheet.setTabColor('#00897b');

  const headers = ['ClientID', 'ClientName', 'ServiceType', 'LabID_Prefix', 'Example'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#00897b')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Example data — edit these to match your real clients
  const exampleData = [
    ['CLIENT001', 'Client A', 'WES', 'HKCH-TWIST', 'HKCH-TWIST-NA12878'],
    ['CLIENT001', 'Client A', 'WGS', 'HKCHG',      'HKCHG-NA12878'],
  ];
  sheet.getRange(2, 1, exampleData.length, 5).setValues(exampleData);
  sheet.autoResizeColumns(1, 5);

  // Add dropdown for ServiceType
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['WES', 'WGS'], true)
    .build();
  sheet.getRange(2, 3, 500, 1).setDataValidation(rule);

  // Move sheet to position after CONFIG
  ss.moveActiveSheet(2);
  ss.setActiveSheet(sheet);

  SpreadsheetApp.getUi().alert(
    '✅ CLIENT_PREFIXES tab created!\n\n' +
    'Please:\n' +
    '1. Update the ClientID and ClientName to match your CLIENTS tab\n' +
    '2. Add a row for each Client + Service Type combination\n' +
    '3. Fill in the LabID_Prefix for each row\n\n' +
    'Example:\n' +
    'CLIENT001 | Client A | WES | HKCH-TWIST\n' +
    'CLIENT001 | Client A | WGS | HKCHG'
  );
}


// ============================================================
// HELPER: Get LabID prefix for a client + service type
// Handles multiple prefixes per client/service (Condition column)
// Returns: { prefix, condition } or prompts user to choose
// ============================================================
function getLabIDPrefix(clientId, serviceType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CLIENT_PREFIXES');
  if (!sheet) return { prefix: serviceType, condition: 'Standard' };

  const data = sheet.getDataRange().getValues();
  const matches = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clientId && data[i][2] === serviceType) {
      matches.push({
        condition: data[i][3], // Condition column
        prefix:    data[i][4]  // LabID_Prefix column
      });
    }
  }

  if (matches.length === 0) return { prefix: serviceType, condition: 'Standard' };
  if (matches.length === 1) return matches[0];

  // Multiple prefixes — ask user to choose
  const ui = SpreadsheetApp.getUi();
  let optionList = '';
  matches.forEach((m, idx) => {
    optionList += (idx + 1) + '. ' + m.condition + ' → ' + m.prefix + '\n';
  });

  const result = ui.prompt(
    '⚠️ Multiple Prefixes Found',
    'Client has multiple ' + serviceType + ' prefixes:\n\n' +
    optionList +
    '\nEnter the number of the prefix to use:',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return matches[0];

  const choice = parseInt(result.getResponseText().trim()) - 1;
  if (choice >= 0 && choice < matches.length) return matches[choice];
  return matches[0]; // fallback to first
}


// ============================================================
// HELPER: Generate LabID
// ============================================================
function generateLabID(prefix, sampleId) {
  return prefix + '-' + sampleId;
}


// ============================================================
// HELPER: Get next Tracking ID
// ============================================================
function getNextID(sheetName, idColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const lastID = sheet.getRange(lastRow, idColumn).getValue();
  return (parseInt(lastID) || 0) + 1;
}


// ============================================================
// HELPER: Get client list for dropdown
// ============================================================
function getClientList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CLIENTS');
  const data = sheet.getDataRange().getValues();
  const clients = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][13] !== 'No') { // ClientID exists and IsActive != No
      clients.push(data[i][0] + ' | ' + data[i][1]); // "CLIENT001 | Client A"
    }
  }
  return clients;
}


// ============================================================
// HELPER: Get project list for a client
// ============================================================
function getProjectList(clientId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PROJECTS');
  const data = sheet.getDataRange().getValues();
  const projects = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === clientId && data[i][9] !== 'No') {
      projects.push(data[i][0]); // ProjectID
    }
  }
  return projects;
}


// ============================================================
// HELPER: Write to Audit Log
// ============================================================
function writeAuditLog(action, sheetAffected, recordId, fieldChanged, oldValue, newValue, remarks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AUDIT_LOG');
  const nextId = getNextID('AUDIT_LOG', 1);
  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date().toISOString();

  sheet.appendRow([
    nextId, timestamp, user, action,
    sheetAffected, recordId, fieldChanged,
    oldValue, newValue, remarks
  ]);
}


// ============================================================
// MAIN: Register New Sample
// Called from 🧬 Lab System > Register New Sample
// ============================================================
function registerNewSample() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Step 1: Get Client ─────────────────────────────────
  const clientList = getClientList();
  if (clientList.length === 0) {
    ui.alert('⚠️ No clients found!\nPlease add clients in the CLIENTS tab first.');
    return;
  }

  const clientResult = ui.prompt(
    '📋 Register New Sample — Step 1 of 5',
    'Enter Client ID:\n\nAvailable clients:\n' + clientList.join('\n') +
    '\n\n(Type only the ClientID e.g. CLIENT001)',
    ui.ButtonSet.OK_CANCEL
  );
  if (clientResult.getSelectedButton() !== ui.Button.OK) return;
  const clientId = clientResult.getResponseText().trim().toUpperCase();

  // Validate client
  const clientSheet = ss.getSheetByName('CLIENTS');
  const clientData = clientSheet.getDataRange().getValues();
  let clientName = '';
  for (let i = 1; i < clientData.length; i++) {
    if (clientData[i][0] === clientId) {
      clientName = clientData[i][1];
      break;
    }
  }
  if (!clientName) {
    ui.alert('❌ Client "' + clientId + '" not found.\nPlease check the CLIENTS tab.');
    return;
  }

  // ── Step 2: Get Service Type ───────────────────────────
  const serviceResult = ui.prompt(
    '📋 Register New Sample — Step 2 of 5',
    'Enter Service Type:\n\n• WES\n• WGS\n\n(Type WES or WGS)',
    ui.ButtonSet.OK_CANCEL
  );
  if (serviceResult.getSelectedButton() !== ui.Button.OK) return;
  const serviceType = serviceResult.getResponseText().trim().toUpperCase();
  if (!['WES', 'WGS'].includes(serviceType)) {
    ui.alert('❌ Invalid service type. Please enter WES or WGS.');
    return;
  }

  // ── Step 3: Get Project ────────────────────────────────
  const projectList = getProjectList(clientId);
  const projectPrompt = projectList.length > 0
    ? 'Existing projects for ' + clientName + ':\n' + projectList.join('\n') +
      '\n\nEnter ProjectID (or type a NEW project ID to create one):'
    : 'No existing projects for ' + clientName + '.\nEnter a NEW ProjectID to create one:';

  const projectResult = ui.prompt(
    '📋 Register New Sample — Step 3 of 5',
    projectPrompt,
    ui.ButtonSet.OK_CANCEL
  );
  if (projectResult.getSelectedButton() !== ui.Button.OK) return;
  const projectId = projectResult.getResponseText().trim();
  if (!projectId) {
    ui.alert('❌ ProjectID cannot be empty.');
    return;
  }

  // Auto-create project if new
  const projectSheet = ss.getSheetByName('PROJECTS');
  const projectData = projectSheet.getDataRange().getValues();
  let projectExists = false;
  for (let i = 1; i < projectData.length; i++) {
    if (projectData[i][0] === projectId) { projectExists = true; break; }
  }
  if (!projectExists) {
    const collectionBatch = ui.prompt(
      '📋 New Project',
      'Enter Collection Batch (or leave blank):',
      ui.ButtonSet.OK_CANCEL
    );
    const batch = collectionBatch.getSelectedButton() === ui.Button.OK
      ? collectionBatch.getResponseText().trim() : '';

    projectSheet.appendRow([
      projectId, clientId, clientName, serviceType,
      batch, '', '', '', 0, 'Yes', '', new Date().toISOString()
    ]);
    writeAuditLog('CREATE_PROJECT', 'PROJECTS', projectId, '', '', projectId, 'New project created');
  }

  // ── Step 4: Get Sample Details ─────────────────────────
  const sampleResult = ui.prompt(
    '📋 Register New Sample — Step 4 of 5',
    'Enter sample details (one per line):\n\n' +
    'SampleID (customer name): \n' +
    'Sample Type (Blood/FFPE/DNA/Saliva/Other): \n' +
    'CapID (capture group, leave blank if N/A): \n' +
    'Pickup Datetime (YYYY-MM-DD HH:MM): \n' +
    'Lab-In Datetime (YYYY-MM-DD HH:MM): \n' +
    'SR Condition (e.g. Good/Degraded): \n' +
    'AWB (airway bill, optional): \n\n' +
    'Format: paste each value separated by a new line.',
    ui.ButtonSet.OK_CANCEL
  );
  if (sampleResult.getSelectedButton() !== ui.Button.OK) return;
  const sampleLines = sampleResult.getResponseText().split('\n');
  const sampleId   = (sampleLines[0] || '').trim();
  const sampleType = (sampleLines[1] || '').trim();
  const capId      = (sampleLines[2] || '').trim();
  const pickupDt   = (sampleLines[3] || '').trim();
  const labInDt    = (sampleLines[4] || '').trim();
  const srCond     = (sampleLines[5] || '').trim();
  const awb        = (sampleLines[6] || '').trim();

  if (!sampleId) {
    ui.alert('❌ SampleID cannot be empty.');
    return;
  }

  // Check for duplicate
  const masterSheet = ss.getSheetByName('MASTER_SAMPLES');
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][0] === sampleId) {
      ui.alert('⚠️ SampleID "' + sampleId + '" already exists!\nPlease check MASTER_SAMPLES tab.');
      return;
    }
  }

  // ── Step 5: Generate LabID ─────────────────────────────
  const prefixResult2 = getLabIDPrefix(clientId, serviceType);
  const prefix    = prefixResult2.prefix;
  const condition = prefixResult2.condition;
  const labId     = generateLabID(prefix, sampleId);

  // Confirm before saving
  const confirm = ui.alert(
    '✅ Confirm Registration — Step 5 of 5',
    'Please confirm the following:\n\n' +
    '• SampleID:     ' + sampleId + '\n' +
    '• LabID:        ' + labId + '\n' +
    '• Client:       ' + clientName + '\n' +
    '• Project:      ' + projectId + '\n' +
    '• Service Type: ' + serviceType + '\n' +
    '• Sample Type:  ' + sampleType + '\n' +
    '• CapID:        ' + (capId || 'N/A') + '\n' +
    '• Pickup:       ' + (pickupDt || 'N/A') + '\n' +
    '• Lab-In:       ' + (labInDt || 'N/A') + '\n\n' +
    'Click OK to register, Cancel to abort.',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // ── Save to MASTER_SAMPLES ─────────────────────────────
  const now = new Date().toISOString();
  const operator = Session.getActiveUser().getEmail();

  masterSheet.appendRow([
    sampleId,           // SampleID
    labId,              // LabID
    projectId,          // ProjectID
    clientId,           // ClientID
    clientName,         // ClientName
    serviceType,        // ServiceType
    sampleType,         // SampleType
    capId,              // CapID
    '',                 // PickupBatch
    pickupDt,           // PickupDatetime
    labInDt,            // LabInDatetime
    operator,           // LabInOperator
    srCond,             // SRCondition
    '',                 // ReceivingRemarks
    '',                 // TransitDate
    awb,                // AWB
    '',                 // TWReceiving
    'sample_ordered',   // CurrentStatus
    now,                // StatusUpdatedTime
    '', '', '', '',     // email/deletion dates
    'No',               // IsRecollection
    '',                 // RecollectionOf
    '',                 // SubmissionFormFormat
    now,                // CreatedDate
    operator            // CreatedBy
  ]);

  // ── Save to STATUS_TRACKING ────────────────────────────
  const statusSheet = ss.getSheetByName('STATUS_TRACKING');
  const trackId = getNextID('STATUS_TRACKING', 1);
  statusSheet.appendRow([
    trackId, sampleId, labId, projectId,
    clientName, 'sample_ordered', 'Sample Ordered',
    now, operator, 'Initial registration'
  ]);

  // ── Save to AUDIT_LOG ──────────────────────────────────
  writeAuditLog('REGISTER_SAMPLE', 'MASTER_SAMPLES', sampleId,
    '', '', labId, 'New sample registered');

  ui.alert(
    '🎉 Sample Registered Successfully!\n\n' +
    '• SampleID: ' + sampleId + '\n' +
    '• LabID:    ' + labId + '\n' +
    '• Status:   Sample Ordered\n\n' +
    'Check MASTER_SAMPLES and STATUS_TRACKING tabs.'
  );
}


// ============================================================
// MAIN: Update Sample Status
// Called from 🧬 Lab System > Update Sample Status
// ============================================================
function updateSampleStatus() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get Sample ID
  const sampleResult = ui.prompt(
    '🔄 Update Sample Status — Step 1 of 2',
    'Enter SampleID or LabID to update:',
    ui.ButtonSet.OK_CANCEL
  );
  if (sampleResult.getSelectedButton() !== ui.Button.OK) return;
  const searchId = sampleResult.getResponseText().trim();

  // Find sample
  const masterSheet = ss.getSheetByName('MASTER_SAMPLES');
  const masterData  = masterSheet.getDataRange().getValues();
  let rowIndex = -1, sampleId = '', labId = '', projectId = '', clientName = '', currentStatus = '';

  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][0] === searchId || masterData[i][1] === searchId) {
      rowIndex      = i + 1; // 1-based row
      sampleId      = masterData[i][0];
      labId         = masterData[i][1];
      projectId     = masterData[i][2];
      clientName    = masterData[i][4];
      currentStatus = masterData[i][17];
      break;
    }
  }

  if (rowIndex === -1) {
    ui.alert('❌ Sample "' + searchId + '" not found.\nCheck MASTER_SAMPLES tab.');
    return;
  }

  // Status list
  const statusOptions = [
    '1.  sample_ordered',
    '2.  sample_arrived_registration',
    '3.  sample_received_by_lab',
    '4.  qc_passed_email_sent',
    '5.  processing_started',
    '6.  processing_finished',
    '7.  sequencing_in_progress',
    '8.  sequencing_finished',
    '9.  data_analysis_finished',
    '10. additional_qc_in_progress',
    '11. additional_qc_passed',
    '12. data_ready_for_review',
    '13. data_uploaded_to_sftp',
    '14. notification_email_sent',
    '15. billing_email_sent'
  ];

  const statusResult = ui.prompt(
    '🔄 Update Sample Status — Step 2 of 2',
    'Sample: ' + labId + ' (' + sampleId + ')\n' +
    'Current status: ' + currentStatus + '\n\n' +
    'Select new status (type the status name):\n\n' +
    statusOptions.join('\n') +
    '\n\nOptional — add remarks on next line.',
    ui.ButtonSet.OK_CANCEL
  );
  if (statusResult.getSelectedButton() !== ui.Button.OK) return;

  const lines     = statusResult.getResponseText().split('\n');
  const newStatus = lines[0].trim();
  const remarks   = (lines[1] || '').trim();

  // Validate status
  const validStatuses = [
    'sample_ordered', 'sample_arrived_registration', 'sample_received_by_lab',
    'qc_passed_email_sent', 'processing_started', 'processing_finished',
    'sequencing_in_progress', 'sequencing_finished', 'data_analysis_finished',
    'additional_qc_in_progress', 'additional_qc_passed', 'data_ready_for_review',
    'data_uploaded_to_sftp', 'notification_email_sent', 'billing_email_sent'
  ];
  if (!validStatuses.includes(newStatus)) {
    ui.alert('❌ Invalid status: "' + newStatus + '"\nPlease copy exactly from the list.');
    return;
  }

  const now      = new Date().toISOString();
  const operator = Session.getActiveUser().getEmail();

  // Update MASTER_SAMPLES current status
  masterSheet.getRange(rowIndex, 18).setValue(newStatus);  // CurrentStatus col
  masterSheet.getRange(rowIndex, 19).setValue(now);        // StatusUpdatedTime col

  // Special fields for specific statuses
  if (newStatus === 'notification_email_sent') {
    masterSheet.getRange(rowIndex, 20).setValue(now);      // NotificationEmailDate
  }
  if (newStatus === 'data_uploaded_to_sftp') {
    masterSheet.getRange(rowIndex, 21).setValue(now);      // DataUploadDate
  }

  // Append to STATUS_TRACKING
  const statusSheet = ss.getSheetByName('STATUS_TRACKING');
  const trackId     = getNextID('STATUS_TRACKING', 1);
  const labelMap = {
    'sample_ordered':               'Sample Ordered',
    'sample_arrived_registration':  'Sample Arrived at Registration',
    'sample_received_by_lab':       'Sample Received by Lab',
    'qc_passed_email_sent':         'QC Passed - Email Sent',
    'processing_started':           'Processing Started',
    'processing_finished':          'Processing Finished',
    'sequencing_in_progress':       'Sequencing In Progress',
    'sequencing_finished':          'Sequencing Finished',
    'data_analysis_finished':       'Data Analysis Finished',
    'additional_qc_in_progress':    'Additional QC In Progress',
    'additional_qc_passed':         'Additional QC Passed',
    'data_ready_for_review':        'Data Ready for Review',
    'data_uploaded_to_sftp':        'Data Uploaded to SFTP',
    'notification_email_sent':      'Notification Email Sent',
    'billing_email_sent':           'Billing Email Sent'
  };

  statusSheet.appendRow([
    trackId, sampleId, labId, projectId,
    clientName, newStatus, labelMap[newStatus] || newStatus,
    now, operator, remarks
  ]);

  // Audit log
  writeAuditLog('UPDATE_STATUS', 'MASTER_SAMPLES', sampleId,
    'CurrentStatus', currentStatus, newStatus, remarks);

  ui.alert(
    '✅ Status Updated!\n\n' +
    '• Sample: ' + labId + '\n' +
    '• New Status: ' + (labelMap[newStatus] || newStatus) + '\n' +
    '• Updated by: ' + operator
  );
}
