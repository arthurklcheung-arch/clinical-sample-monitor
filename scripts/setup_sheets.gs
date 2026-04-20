// ============================================================
// Clinical Sample Monitor - Google Sheets Setup Script
// Paste this into Extensions > Apps Script and click Run
// ============================================================

function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getUi().alert('Setting up Clinical Sample Monitor...\nThis may take 30-60 seconds.');

  setupConfigSheet(ss);
  setupClientsSheet(ss);
  setupProjectsSheet(ss);
  setupMasterSamplesSheet(ss);
  setupStatusTrackingSheet(ss);
  setupQCMetricsSheet(ss);
  setupBillingSheet(ss);
  setupDashboardSheet(ss);
  setupAuditLogSheet(ss);
  removeDefaultSheet(ss);
  setupMenu();

  SpreadsheetApp.getUi().alert('✅ Setup Complete!\n\nAll tabs have been created.\nStart by adding your clients in the CLIENTS tab.');
}


// ============================================================
// HELPER FUNCTIONS
// ============================================================
function headerStyle() {
  return SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#FFFFFF').build();
}

function styleHeaders(sheet, numCols, color) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground(color || '#1a73e8');
  headerRange.setTextStyle(headerStyle());
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }
  return sheet;
}

function removeDefaultSheet(ss) {
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);
}

function addDropdownValidation(sheet, row, col, numRows, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col, numRows, 1).setDataValidation(rule);
}


// ============================================================
// SHEET 1: CONFIG
// System settings, dropdowns, service types
// ============================================================
function setupConfigSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'CONFIG');
  sheet.setTabColor('#757575');

  // Section: Status List
  const statusHeaders = [['STATUS_ID', 'STATUS_LABEL', 'DISPLAY_ORDER']];
  const statuses = [
    ['sample_ordered',              'Sample Ordered',                   1],
    ['sample_arrived_registration', 'Sample Arrived at Registration',   2],
    ['sample_received_by_lab',      'Sample Received by Lab',           3],
    ['qc_passed_email_sent',        'QC Passed - Email Sent',           4],
    ['processing_started',          'Processing Started',               5],
    ['processing_finished',         'Processing Finished',              6],
    ['sequencing_in_progress',      'Sequencing In Progress',           7],
    ['sequencing_finished',         'Sequencing Finished',              8],
    ['data_analysis_finished',      'Data Analysis Finished',           9],
    ['additional_qc_in_progress',   'Additional QC In Progress',        10],
    ['additional_qc_passed',        'Additional QC Passed',             11],
    ['data_ready_for_review',       'Data Ready for Review',            12],
    ['data_uploaded_to_sftp',       'Data Uploaded to SFTP',            13],
    ['notification_email_sent',     'Notification Email Sent',          14],
    ['billing_email_sent',          'Billing Email Sent',               15],
  ];

  sheet.getRange(1, 1, 1, 3).setValues(statusHeaders);
  sheet.getRange(2, 1, statuses.length, 3).setValues(statuses);
  styleHeaders(sheet, 3, '#757575');

  // Section: Service Types
  sheet.getRange(1, 5).setValue('SERVICE_TYPE_ID');
  sheet.getRange(1, 6).setValue('SERVICE_TYPE_NAME');
  sheet.getRange(2, 5).setValue('WES');
  sheet.getRange(2, 6).setValue('Whole Exome Sequencing');
  sheet.getRange(3, 5).setValue('WGS');
  sheet.getRange(3, 6).setValue('Whole Genome Sequencing');
  sheet.getRange(1, 5, 1, 2).setBackground('#757575').setFontColor('#FFFFFF').setFontWeight('bold');

  // Section: WGS Billing Mode options
  sheet.getRange(1, 8).setValue('WGS_BILLING_MODE');
  sheet.getRange(1, 9).setValue('DESCRIPTION');
  sheet.getRange(2, 8).setValue('3x_wes');
  sheet.getRange(2, 9).setValue('1 WGS = 3 WES billing units');
  sheet.getRange(3, 8).setValue('independent');
  sheet.getRange(3, 9).setValue('1 WGS = 1 WGS billing unit');
  sheet.getRange(1, 8, 1, 2).setBackground('#757575').setFontColor('#FFFFFF').setFontWeight('bold');

  // Section: Currency options
  sheet.getRange(1, 11).setValue('CURRENCY');
  sheet.getRange(2, 11).setValue('HKD');
  sheet.getRange(3, 11).setValue('USD');
  sheet.getRange(1, 11).setBackground('#757575').setFontColor('#FFFFFF').setFontWeight('bold');

  sheet.autoResizeColumns(1, 11);
}


// ============================================================
// SHEET 2: CLIENTS
// ============================================================
function setupClientsSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'CLIENTS');
  sheet.setTabColor('#0f9d58');

  const headers = [
    'ClientID', 'ClientName', 'ContactEmail', 'BillingCurrency',
    'BillingPeriodDayStart', 'CustomBillingPeriod', 'RequiresAdditionalQC',
    'LabIDPrefix', 'SFTPPath', 'ConfidentialityLevel',
    'WES_UnitPrice', 'WGS_BillingMode', 'WGS_UnitPrice',
    'IsActive', 'Notes', 'CreatedDate'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaders(sheet, headers.length, '#0f9d58');

  // Dropdowns
  addDropdownValidation(sheet, 2, 4, 500, ['HKD', 'USD']);           // BillingCurrency
  addDropdownValidation(sheet, 2, 7, 500, ['Yes', 'No']);            // RequiresAdditionalQC
  addDropdownValidation(sheet, 2, 10, 500, ['general', 'confidential']); // ConfidentialityLevel
  addDropdownValidation(sheet, 2, 12, 500, ['3x_wes', 'independent']); // WGS_BillingMode
  addDropdownValidation(sheet, 2, 14, 500, ['Yes', 'No']);           // IsActive

  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 220);
}


// ============================================================
// SHEET 3: PROJECTS
// ============================================================
function setupProjectsSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'PROJECTS');
  sheet.setTabColor('#0f9d58');

  const headers = [
    'ProjectID', 'ClientID', 'ClientName', 'ServiceType',
    'CollectionBatch', 'SubmissionFormFormat',
    'CustomBillingStart', 'CustomBillingEnd',
    'SampleCount', 'IsActive', 'Notes', 'CreatedDate'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaders(sheet, headers.length, '#0f9d58');

  addDropdownValidation(sheet, 2, 4, 500, ['WES', 'WGS']);   // ServiceType
  addDropdownValidation(sheet, 2, 10, 500, ['Yes', 'No']);    // IsActive

  sheet.autoResizeColumns(1, headers.length);
}


// ============================================================
// SHEET 4: MASTER_SAMPLES
// ============================================================
function setupMasterSamplesSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'MASTER_SAMPLES');
  sheet.setTabColor('#1a73e8');

  const headers = [
    'SampleID',           // Customer sample name
    'LabID',              // Internal: prefix + SampleID
    'ProjectID',
    'ClientID',
    'ClientName',
    'ServiceType',
    'SampleType',
    'CapID',
    'PickupBatch',
    'PickupDatetime',
    'LabInDatetime',
    'LabInOperator',
    'SRCondition',
    'ReceivingRemarks',
    'TransitDate',
    'AWB',
    'TWReceiving',
    'CurrentStatus',
    'StatusUpdatedTime',
    'NotificationEmailDate',
    'DataUploadDate',
    'SFTPDeletionDate',
    'FinalDeletionDate',
    'IsRecollection',
    'RecollectionOf',
    'SubmissionFormFormat',
    'CreatedDate',
    'CreatedBy'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaders(sheet, headers.length, '#1a73e8');

  // Dropdowns
  addDropdownValidation(sheet, 2, 6, 2000, ['WES', 'WGS']);          // ServiceType
  addDropdownValidation(sheet, 2, 7, 2000, ['Blood', 'FFPE', 'DNA', 'Saliva', 'Other']); // SampleType
  addDropdownValidation(sheet, 2, 24, 2000, ['Yes', 'No']);          // IsRecollection

  // Status dropdown (from CONFIG)
  const statusList = [
    'sample_ordered', 'sample_arrived_registration', 'sample_received_by_lab',
    'qc_passed_email_sent', 'processing_started', 'processing_finished',
    'sequencing_in_progress', 'sequencing_finished', 'data_analysis_finished',
    'additional_qc_in_progress', 'additional_qc_passed', 'data_ready_for_review',
    'data_uploaded_to_sftp', 'notification_email_sent', 'billing_email_sent'
  ];
  addDropdownValidation(sheet, 2, 18, 2000, statusList);             // CurrentStatus

  // Freeze and format
  sheet.setFrozenColumns(2);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 160);
}


// ============================================================
// SHEET 5: STATUS_TRACKING
// Full audit trail of every status change
// ============================================================
function setupStatusTrackingSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'STATUS_TRACKING');
  sheet.setTabColor('#f4b400');

  const headers = [
    'TrackingID', 'SampleID', 'LabID', 'ProjectID',
    'ClientName', 'Status', 'StatusLabel',
    'Timestamp', 'Operator', 'Remarks'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaders(sheet, headers.length, '#f4b400');

  sheet.setFrozenColumns(2);
  sheet.autoResizeColumns(1, headers.length);
}


// ============================================================
// SHEET 6: QC_METRICS
// ============================================================
function setupQCMetricsSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'QC_METRICS');
  sheet.setTabColor('#e91e63');

  const headers = [
    'SampleID', 'LabID', 'ProjectID', 'ClientName',
    'RunID', 'RunNumber', 'QC_Conclusion', 'QC_FailReason',
    'MappingRatio', 'TargetDepth', 'BaseQ30', 'GC_Rate',
    'InsertSize', 'DupRatio', 'TitvRatio', 'VariantRatio',
    'HethomRatio', 'SnpindelRatio', 'GenderAnalysed', 'GenderRecord',
    'ContaminationLevel', 'OntargetRatio', 'BasesTarget0x',
    'BasesTargetLt20x', 'PCR_QC', 'SNPQC_Status',
    // SNP markers
    'rs1042713_wes', 'rs1042713_pcr',
    'rs1801133_wes', 'rs1801133_pcr',
    'rs8192678_wes', 'rs8192678_pcr',
    'rs4343_wes',    'rs4343_pcr',
    'rs713598_wes',  'rs713598_pcr',
    'CreatedDate'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaders(sheet, headers.length, '#e91e63');

  addDropdownValidation(sheet, 2, 7, 2000, ['PASS', 'FAIL']);  // QC_Conclusion

  sheet.setFrozenColumns(2);
  sheet.autoResizeColumns(1, headers.length);
}


// ============================================================
// SHEET 7: BILLING
// ============================================================
function setupBillingSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'BILLING');
  sheet.setTabColor('#ff6d00');

  const headers = [
    'BillingID', 'ClientID', 'ClientName',
    'SampleID', 'LabID', 'ServiceType',
    'BillingPeriodStart', 'BillingPeriodEnd',
    'NotificationEmailDate',
    'PickupDate', 'LabInDate',
    'BillingMode', 'BillingUnits', 'UnitPrice', 'Currency',
    'TotalAmount', 'PDF_RowLabel',
    'BillingEmailSent', 'BillingEmailDate',
    'InvoicePDFPath', 'Remarks', 'CreatedDate'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaders(sheet, headers.length, '#ff6d00');

  addDropdownValidation(sheet, 2, 6, 2000, ['WES', 'WGS']);
  addDropdownValidation(sheet, 2, 12, 2000, ['per_unit', '3x_wes', 'independent']);
  addDropdownValidation(sheet, 2, 15, 2000, ['HKD', 'USD']);
  addDropdownValidation(sheet, 2, 18, 2000, ['Yes', 'No']);

  sheet.setFrozenColumns(3);
  sheet.autoResizeColumns(1, headers.length);
}


// ============================================================
// SHEET 8: DASHBOARD
// Auto-calculated summary view
// ============================================================
function setupDashboardSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'DASHBOARD');
  sheet.setTabColor('#9c27b0');

  // Title
  sheet.getRange('A1').setValue('🧬 Clinical Sample Monitor — Dashboard');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');
  sheet.getRange('A2').setValue('Last updated: (click Refresh Dashboard to update)');
  sheet.getRange('A2').setFontColor('#757575').setFontStyle('italic');

  // Section: Sample Summary
  sheet.getRange('A4').setValue('📋 SAMPLE SUMMARY');
  sheet.getRange('A4').setFontWeight('bold').setFontSize(12).setBackground('#e8f0fe');
  sheet.getRange('A4:D4').merge().setBackground('#e8f0fe');

  const summaryLabels = [
    ['Total Samples', ''],
    ['Samples This Month', ''],
    ['Pending (In Progress)', ''],
    ['Completed (Email Sent)', ''],
    ['Failed QC', ''],
  ];
  sheet.getRange('A5:B9').setValues(summaryLabels);
  sheet.getRange('A5:A9').setFontWeight('bold');

  // Section: Status Breakdown
  sheet.getRange('A11').setValue('📊 STATUS BREAKDOWN');
  sheet.getRange('A11').setFontWeight('bold').setFontSize(12).setBackground('#e8f0fe');
  sheet.getRange('A11:D11').merge().setBackground('#e8f0fe');

  const statusLabels = [
    ['Sample Ordered', ''],
    ['Sample Received by Lab', ''],
    ['Processing', ''],
    ['Sequencing In Progress', ''],
    ['Data Analysis Finished', ''],
    ['Notification Email Sent', ''],
    ['Billing Email Sent', ''],
  ];
  sheet.getRange('A12:B18').setValues(statusLabels);
  sheet.getRange('A12:A18').setFontWeight('bold');

  // Section: Revenue Summary
  sheet.getRange('D4').setValue('💰 REVENUE SUMMARY');
  sheet.getRange('D4').setFontWeight('bold').setFontSize(12).setBackground('#fff3e0');
  sheet.getRange('D4:G4').merge().setBackground('#fff3e0');

  const revenueLabels = [
    ['This Month (HKD)', ''],
    ['This Month (USD)', ''],
    ['This Quarter (HKD)', ''],
    ['This Year (HKD)', ''],
    ['Pending Billing', ''],
  ];
  sheet.getRange('D5:E9').setValues(revenueLabels);
  sheet.getRange('D5:D9').setFontWeight('bold');

  // Section: Client Breakdown
  sheet.getRange('D11').setValue('👥 SAMPLES BY CLIENT');
  sheet.getRange('D11').setFontWeight('bold').setFontSize(12).setBackground('#fff3e0');
  sheet.getRange('D11:G11').merge().setBackground('#fff3e0');
  sheet.getRange('D12').setValue('Client');
  sheet.getRange('E12').setValue('Samples');
  sheet.getRange('F12').setValue('Revenue (HKD)');
  sheet.getRange('D12:F12').setFontWeight('bold').setBackground('#fff3e0');

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 140);
}


// ============================================================
// SHEET 9: AUDIT_LOG
// Record of all changes made in the system
// ============================================================
function setupAuditLogSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'AUDIT_LOG');
  sheet.setTabColor('#607d8b');

  const headers = [
    'LogID', 'Timestamp', 'User', 'Action',
    'SheetAffected', 'RecordID', 'FieldChanged',
    'OldValue', 'NewValue', 'Remarks'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaders(sheet, headers.length, '#607d8b');
  sheet.autoResizeColumns(1, headers.length);
}


// ============================================================
// MENU SETUP
// Adds "🧬 Lab System" menu to the top menu bar
// ============================================================
function setupMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧬 Lab System')
    .addItem('➕ Register New Sample', 'registerNewSample')
    .addItem('🔄 Update Sample Status', 'updateSampleStatus')
    .addSeparator()
    .addItem('💰 Calculate Billing', 'calculateBilling')
    .addItem('📄 Generate Billing PDF', 'generateBillingPDF')
    .addItem('📧 Send Billing Email', 'sendBillingEmail')
    .addSeparator()
    .addItem('📊 Refresh Dashboard', 'refreshDashboard')
    .addItem('🔔 Send Notification Email', 'sendNotificationEmail')
    .addSeparator()
    .addItem('💾 Backup Now', 'backupDatabase')
    .addItem('⚙️ Setup (Run Once)', 'setupAllSheets')
    .addToUi();
}

// Placeholder functions (will be built in next phase)
function registerNewSample() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Sample Registration module');
}
function updateSampleStatus() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Status Update module');
}
function calculateBilling() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Billing Calculation module');
}
function generateBillingPDF() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: PDF Generation module');
}
function sendBillingEmail() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Billing Email module');
}
function refreshDashboard() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Dashboard Refresh module');
}
function sendNotificationEmail() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Notification Email module');
}
function backupDatabase() {
  SpreadsheetApp.getUi().alert('🚧 Coming soon: Backup module');
}

// Auto-run menu on open
function onOpen() {
  setupMenu();
}
