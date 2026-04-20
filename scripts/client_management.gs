// ============================================================
// Clinical Sample Monitor - Client Management Module
// Add this as a new script file in Apps Script: client_management
// ============================================================


// ============================================================
// SETUP: Upgrade CLIENT_PREFIXES to add Condition column
// Run this ONCE — it will migrate your existing data safely
// ============================================================
function upgradeClientPrefixSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('CLIENT_PREFIXES');

  // If sheet doesn't exist, create fresh
  if (!sheet) {
    sheet = ss.insertSheet('CLIENT_PREFIXES');
  } else {
    // Back up existing data before modifying
    const existingData = sheet.getDataRange().getValues();
    sheet.clearContents();
    sheet.clearFormats();

    // Re-write with new structure (insert Condition column after ServiceType)
    const newHeaders = ['ClientID', 'ClientName', 'ServiceType', 'Condition', 'LabID_Prefix', 'Example'];
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);

    // Migrate existing rows — set Condition to 'Standard' by default
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0]) {
        sheet.appendRow([
          existingData[i][0],  // ClientID
          existingData[i][1],  // ClientName
          existingData[i][2],  // ServiceType
          'Standard',          // Condition (new — default)
          existingData[i][3],  // LabID_Prefix
          existingData[i][4],  // Example
        ]);
      }
    }
  }

  // Format headers
  const newHeaders = ['ClientID', 'ClientName', 'ServiceType', 'Condition', 'LabID_Prefix', 'Example'];
  sheet.getRange(1, 1, 1, newHeaders.length)
    .setBackground('#00897b')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Dropdowns
  const serviceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['WES', 'WGS'], true).build();
  sheet.getRange(2, 3, 500, 1).setDataValidation(serviceRule);

  sheet.autoResizeColumns(1, 6);
  sheet.setTabColor('#00897b');

  ui.alert(
    '✅ CLIENT_PREFIXES upgraded!\n\n' +
    'A new "Condition" column has been added.\n' +
    'Existing rows have been set to "Standard".\n\n' +
    'You can now add multiple prefixes per client\nby adding rows with different Conditions.\n\n' +
    'Example:\n' +
    'CLIENT001 | Client A | WES | TWIST    | HKCH-TWIST\n' +
    'CLIENT001 | Client A | WES | Research | HKCH-RE\n' +
    'CLIENT001 | Client A | WGS | Standard | HKCHG'
  );
}


// ============================================================
// ADD NEW CLIENT
// Guides user through adding a client + auto-populates prefixes
// ============================================================
function addNewClient() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientSheet = ss.getSheetByName('CLIENTS');

  // ── Step 1: Basic client info ──────────────────────────
  const step1 = ui.prompt(
    '👥 Add New Client — Step 1 of 4',
    'Enter client details (one per line):\n\n' +
    'Line 1: ClientID (e.g. CLIENT002)\n' +
    'Line 2: Client Name\n' +
    'Line 3: Contact Email\n' +
    'Line 4: Billing Currency (HKD or USD)',
    ui.ButtonSet.OK_CANCEL
  );
  if (step1.getSelectedButton() !== ui.Button.OK) return;

  const lines1   = step1.getResponseText().split('\n');
  const clientId = (lines1[0] || '').trim().toUpperCase();
  const clientName = (lines1[1] || '').trim();
  const contactEmail = (lines1[2] || '').trim();
  const currency = (lines1[3] || 'HKD').trim().toUpperCase() || 'HKD';

  if (!clientId || !clientName) {
    ui.alert('❌ ClientID and Client Name are required.');
    return;
  }

  // Check duplicate
  const existingData = clientSheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0] === clientId) {
      ui.alert('❌ ClientID "' + clientId + '" already exists.\nPlease use a different ID.');
      return;
    }
  }

  // ── Step 2: Billing settings ───────────────────────────
  const step2 = ui.prompt(
    '👥 Add New Client — Step 2 of 4',
    'Enter billing settings (one per line):\n\n' +
    'Line 1: Billing period start day (1-28, default: 1)\n' +
    'Line 2: Requires additional QC? (Yes or No)\n' +
    'Line 3: Data confidentiality (general or confidential)\n' +
    'Line 4: Lab ID Prefix base (e.g. HKCH)',
    ui.ButtonSet.OK_CANCEL
  );
  if (step2.getSelectedButton() !== ui.Button.OK) return;

  const lines2         = step2.getResponseText().split('\n');
  const billingDay     = parseInt(lines2[0]) || 1;
  const requiresQC     = (lines2[1] || 'No').trim();
  const confidentiality = (lines2[2] || 'general').trim().toLowerCase();
  const labIdBase      = (lines2[3] || '').trim();

  // ── Step 3: Service types & prefixes ──────────────────
  const step3 = ui.prompt(
    '👥 Add New Client — Step 3 of 4',
    'Enter LabID prefixes for each service type.\n' +
    'Format: ServiceType | Condition | Prefix\n' +
    'One per line. Leave blank to skip a service.\n\n' +
    'Examples:\n' +
    'WES | TWIST | HKCH-TWIST\n' +
    'WES | Research | HKCH-RE\n' +
    'WGS | Standard | HKCHG\n\n' +
    'Enter your prefixes:',
    ui.ButtonSet.OK_CANCEL
  );
  if (step3.getSelectedButton() !== ui.Button.OK) return;

  const prefixLines = step3.getResponseText().split('\n').filter(l => l.trim());
  const prefixes = [];

  for (const line of prefixLines) {
    const parts = line.split('|').map(p => p.trim());
    if (parts.length >= 3) {
      const serviceType = parts[0].toUpperCase();
      const condition   = parts[1];
      const prefix      = parts[2];
      const example     = prefix + '-NA12878';
      prefixes.push([clientId, clientName, serviceType, condition, prefix, example]);
    }
  }

  // ── Step 4: WGS billing mode ───────────────────────────
  const step4 = ui.prompt(
    '👥 Add New Client — Step 4 of 4',
    'WGS Billing Mode:\n\n' +
    '• 3x_wes       = 1 WGS = 3 WES billing units (current)\n' +
    '• independent  = 1 WGS = 1 WGS billing unit (future)\n\n' +
    'Also enter pricing (one per line):\n' +
    'Line 1: WGS billing mode (3x_wes or independent)\n' +
    'Line 2: WES unit price (e.g. 1000)\n' +
    'Line 3: WGS unit price (only if independent mode)\n' +
    'Line 4: SFTP path (optional)',
    ui.ButtonSet.OK_CANCEL
  );
  if (step4.getSelectedButton() !== ui.Button.OK) return;

  const lines4      = step4.getResponseText().split('\n');
  const wgsBilling  = (lines4[0] || '3x_wes').trim();
  const wesPrice    = parseFloat(lines4[1]) || 0;
  const wgsPrice    = parseFloat(lines4[2]) || 0;
  const sftpPath    = (lines4[3] || '').trim();

  // ── Confirm ────────────────────────────────────────────
  let prefixSummary = prefixes.length > 0
    ? prefixes.map(p => '  ' + p[2] + ' | ' + p[3] + ' → ' + p[4]).join('\n')
    : '  (none entered)';

  const confirm = ui.alert(
    '✅ Confirm New Client',
    'Please confirm:\n\n' +
    '• ClientID:       ' + clientId + '\n' +
    '• Name:           ' + clientName + '\n' +
    '• Email:          ' + (contactEmail || 'N/A') + '\n' +
    '• Currency:       ' + currency + '\n' +
    '• Billing day:    ' + billingDay + '\n' +
    '• Additional QC:  ' + requiresQC + '\n' +
    '• Confidentiality:' + confidentiality + '\n' +
    '• WES Price:      ' + wesPrice + ' ' + currency + '\n' +
    '• WGS Mode:       ' + wgsBilling + '\n' +
    '• WGS Price:      ' + (wgsPrice || 'N/A') + '\n\n' +
    'LabID Prefixes:\n' + prefixSummary + '\n\n' +
    'Click OK to save.',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // ── Save to CLIENTS tab ────────────────────────────────
  const now = new Date().toISOString();
  clientSheet.appendRow([
    clientId,           // ClientID
    clientName,         // ClientName
    contactEmail,       // ContactEmail
    currency,           // BillingCurrency
    billingDay,         // BillingPeriodDayStart
    '',                 // CustomBillingPeriod
    requiresQC,         // RequiresAdditionalQC
    labIdBase,          // LabIDPrefix
    sftpPath,           // SFTPPath
    confidentiality,    // ConfidentialityLevel
    wesPrice,           // WES_UnitPrice
    wgsBilling,         // WGS_BillingMode
    wgsPrice || '',     // WGS_UnitPrice
    'Yes',              // IsActive
    '',                 // Notes
    now                 // CreatedDate
  ]);

  // ── Save to CLIENT_PREFIXES tab ────────────────────────
  const prefixSheet = ss.getSheetByName('CLIENT_PREFIXES');
  if (prefixSheet && prefixes.length > 0) {
    for (const prefix of prefixes) {
      prefixSheet.appendRow(prefix);
    }
  }

  // ── Audit Log ──────────────────────────────────────────
  writeAuditLog(
    'ADD_CLIENT', 'CLIENTS', clientId,
    '', '', clientName,
    'New client added with ' + prefixes.length + ' prefix(es)'
  );

  ui.alert(
    '🎉 Client Added Successfully!\n\n' +
    '• ClientID: ' + clientId + '\n' +
    '• Name: ' + clientName + '\n' +
    '• Prefixes added: ' + prefixes.length + '\n\n' +
    'Check CLIENTS and CLIENT_PREFIXES tabs.'
  );
}


// ============================================================
// SYNC EXISTING CLIENT TO CLIENT_PREFIXES
// Use this to sync a client you already added manually
// ============================================================
function syncExistingClientPrefixes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Show existing clients
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

  const idResult = ui.prompt(
    '🔗 Sync Client Prefixes',
    'Clients in CLIENTS tab:\n\n' + clientList +
    '\nEnter ClientID to add prefixes for:',
    ui.ButtonSet.OK_CANCEL
  );
  if (idResult.getSelectedButton() !== ui.Button.OK) return;
  const clientId   = idResult.getResponseText().trim().toUpperCase();
  const clientName = clientMap[clientId];

  if (!clientName) {
    ui.alert('❌ ClientID "' + clientId + '" not found.');
    return;
  }

  // Ask for prefixes
  const prefixResult = ui.prompt(
    '🔗 Add Prefixes for ' + clientName,
    'Enter prefixes (one per line):\n' +
    'Format: ServiceType | Condition | Prefix\n\n' +
    'Examples:\n' +
    'WES | TWIST | HKCH-TWIST\n' +
    'WES | Research | HKCH-RE\n' +
    'WGS | Standard | HKCHG',
    ui.ButtonSet.OK_CANCEL
  );
  if (prefixResult.getSelectedButton() !== ui.Button.OK) return;

  const lines = prefixResult.getResponseText().split('\n').filter(l => l.trim());
  const prefixSheet = ss.getSheetByName('CLIENT_PREFIXES');
  let added = 0;

  for (const line of lines) {
    const parts = line.split('|').map(p => p.trim());
    if (parts.length >= 3) {
      const serviceType = parts[0].toUpperCase();
      const condition   = parts[1];
      const prefix      = parts[2];
      const example     = prefix + '-NA12878';
      prefixSheet.appendRow([clientId, clientName, serviceType, condition, prefix, example]);
      added++;
    }
  }

  writeAuditLog('SYNC_PREFIXES', 'CLIENT_PREFIXES', clientId,
    '', '', clientName, added + ' prefix(es) added');

  ui.alert('✅ Done!\n\n' + added + ' prefix(es) added for ' + clientName + '.\nCheck CLIENT_PREFIXES tab.');
}


// ============================================================
// UPDATED MENU — replaces previous setupMenu()
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
    .addSubMenu(ui.createMenu('⚙️ Admin Tools')
      .addItem('👥 Add New Client', 'addNewClient')
      .addItem('🔗 Sync Client Prefixes', 'syncExistingClientPrefixes')
      .addItem('✏️ Rename Client', 'renameClient')
      .addItem('✏️ Rename Project', 'renameProject')
      .addItem('🔧 Upgrade CLIENT_PREFIXES', 'upgradeClientPrefixSheet')
    )
    .addSeparator()
    .addItem('💾 Backup Now', 'backupDatabase')
    .addToUi();
}
