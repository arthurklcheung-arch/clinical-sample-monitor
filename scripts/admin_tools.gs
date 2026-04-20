// ============================================================
// Clinical Sample Monitor - Admin Tools
// Add this as a new script file in Apps Script
// ============================================================


// ============================================================
// RENAME CLIENT
// Updates client name across ALL sheets
// ============================================================
function renameClient() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Step 1: Show existing clients ─────────────────────
  const clientSheet = ss.getSheetByName('CLIENTS');
  const clientData  = clientSheet.getDataRange().getValues();

  let clientList = '';
  const clientMap = {}; // clientId → clientName
  for (let i = 1; i < clientData.length; i++) {
    if (clientData[i][0]) {
      clientMap[clientData[i][0]] = clientData[i][1];
      clientList += clientData[i][0] + ' | ' + clientData[i][1] + '\n';
    }
  }

  if (!clientList) {
    ui.alert('⚠️ No clients found in CLIENTS tab.');
    return;
  }

  // ── Step 2: Ask which client to rename ────────────────
  const idResult = ui.prompt(
    '✏️ Rename Client — Step 1 of 2',
    'Current clients:\n\n' + clientList +
    '\nEnter the ClientID you want to rename:',
    ui.ButtonSet.OK_CANCEL
  );
  if (idResult.getSelectedButton() !== ui.Button.OK) return;
  const clientId = idResult.getResponseText().trim().toUpperCase();

  if (!clientMap[clientId]) {
    ui.alert('❌ ClientID "' + clientId + '" not found.');
    return;
  }

  const oldName = clientMap[clientId];

  // ── Step 3: Ask for new name ───────────────────────────
  const nameResult = ui.prompt(
    '✏️ Rename Client — Step 2 of 2',
    'ClientID: ' + clientId + '\n' +
    'Current name: ' + oldName + '\n\n' +
    'Enter the NEW client name:',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResult.getSelectedButton() !== ui.Button.OK) return;
  const newName = nameResult.getResponseText().trim();

  if (!newName) {
    ui.alert('❌ New name cannot be empty.');
    return;
  }
  if (newName === oldName) {
    ui.alert('⚠️ New name is the same as the current name. No changes made.');
    return;
  }

  // ── Confirm ────────────────────────────────────────────
  const confirm = ui.alert(
    '⚠️ Confirm Rename',
    'This will update "' + oldName + '" to "' + newName + '" across ALL sheets.\n\n' +
    'Sheets affected:\n' +
    '• CLIENTS\n' +
    '• CLIENT_PREFIXES\n' +
    '• PROJECTS\n' +
    '• MASTER_SAMPLES\n' +
    '• STATUS_TRACKING\n' +
    '• BILLING\n' +
    '• AUDIT_LOG (logged only)\n\n' +
    'Click OK to proceed.',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  let totalUpdated = 0;

  // ── Update CLIENTS ─────────────────────────────────────
  for (let i = 1; i < clientData.length; i++) {
    if (clientData[i][0] === clientId) {
      clientSheet.getRange(i + 1, 2).setValue(newName);
      totalUpdated++;
      break;
    }
  }

  // ── Update CLIENT_PREFIXES ─────────────────────────────
  const prefixSheet = ss.getSheetByName('CLIENT_PREFIXES');
  if (prefixSheet) {
    const prefixData = prefixSheet.getDataRange().getValues();
    for (let i = 1; i < prefixData.length; i++) {
      if (prefixData[i][0] === clientId) {
        prefixSheet.getRange(i + 1, 2).setValue(newName);
        totalUpdated++;
      }
    }
  }

  // ── Update PROJECTS ────────────────────────────────────
  const projectSheet = ss.getSheetByName('PROJECTS');
  if (projectSheet) {
    const projectData = projectSheet.getDataRange().getValues();
    for (let i = 1; i < projectData.length; i++) {
      if (projectData[i][1] === clientId) {
        projectSheet.getRange(i + 1, 3).setValue(newName); // ClientName col
        totalUpdated++;
      }
    }
  }

  // ── Update MASTER_SAMPLES ──────────────────────────────
  const masterSheet = ss.getSheetByName('MASTER_SAMPLES');
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][3] === clientId) { // ClientID col
        masterSheet.getRange(i + 1, 5).setValue(newName); // ClientName col
        totalUpdated++;
      }
    }
  }

  // ── Update STATUS_TRACKING ─────────────────────────────
  const statusSheet = ss.getSheetByName('STATUS_TRACKING');
  if (statusSheet) {
    const statusData = statusSheet.getDataRange().getValues();
    for (let i = 1; i < statusData.length; i++) {
      if (statusData[i][4] === oldName) { // ClientName col
        statusSheet.getRange(i + 1, 5).setValue(newName);
        totalUpdated++;
      }
    }
  }

  // ── Update BILLING ─────────────────────────────────────
  const billingSheet = ss.getSheetByName('BILLING');
  if (billingSheet) {
    const billingData = billingSheet.getDataRange().getValues();
    for (let i = 1; i < billingData.length; i++) {
      if (billingData[i][1] === clientId) { // ClientID col
        billingSheet.getRange(i + 1, 3).setValue(newName); // ClientName col
        totalUpdated++;
      }
    }
  }

  // ── Write to Audit Log ─────────────────────────────────
  writeAuditLog(
    'RENAME_CLIENT', 'CLIENTS', clientId,
    'ClientName', oldName, newName,
    'Client renamed — ' + totalUpdated + ' records updated'
  );

  ui.alert(
    '✅ Client Renamed Successfully!\n\n' +
    '• ClientID:  ' + clientId + '\n' +
    '• Old name:  ' + oldName + '\n' +
    '• New name:  ' + newName + '\n' +
    '• Records updated: ' + totalUpdated
  );
}


// ============================================================
// RENAME PROJECT
// Updates project ID or name across ALL sheets
// ============================================================
function renameProject() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Step 1: Show existing projects ────────────────────
  const projectSheet = ss.getSheetByName('PROJECTS');
  const projectData  = projectSheet.getDataRange().getValues();

  let projectList = '';
  for (let i = 1; i < projectData.length; i++) {
    if (projectData[i][0]) {
      projectList += projectData[i][0] + ' | ' + projectData[i][2] + '\n'; // ProjectID | ClientName
    }
  }

  if (!projectList) {
    ui.alert('⚠️ No projects found in PROJECTS tab.');
    return;
  }

  // ── Step 2: Ask which project ─────────────────────────
  const idResult = ui.prompt(
    '✏️ Rename Project — Step 1 of 2',
    'Current projects (ProjectID | Client):\n\n' + projectList +
    '\nEnter the ProjectID you want to rename:',
    ui.ButtonSet.OK_CANCEL
  );
  if (idResult.getSelectedButton() !== ui.Button.OK) return;
  const oldProjectId = idResult.getResponseText().trim();

  // Validate
  let projectRowIndex = -1;
  for (let i = 1; i < projectData.length; i++) {
    if (projectData[i][0] === oldProjectId) {
      projectRowIndex = i + 1;
      break;
    }
  }
  if (projectRowIndex === -1) {
    ui.alert('❌ Project "' + oldProjectId + '" not found.');
    return;
  }

  // ── Step 3: Ask for new ID ─────────────────────────────
  const newIdResult = ui.prompt(
    '✏️ Rename Project — Step 2 of 2',
    'Current ProjectID: ' + oldProjectId + '\n\n' +
    'Enter the NEW ProjectID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (newIdResult.getSelectedButton() !== ui.Button.OK) return;
  const newProjectId = newIdResult.getResponseText().trim();

  if (!newProjectId) {
    ui.alert('❌ New ProjectID cannot be empty.');
    return;
  }
  if (newProjectId === oldProjectId) {
    ui.alert('⚠️ New ProjectID is the same. No changes made.');
    return;
  }

  // ── Confirm ────────────────────────────────────────────
  const confirm = ui.alert(
    '⚠️ Confirm Rename',
    'Rename project:\n"' + oldProjectId + '" → "' + newProjectId + '"\n\n' +
    'This will update PROJECTS, MASTER_SAMPLES,\nSTATUS_TRACKING and BILLING.\n\n' +
    'Click OK to proceed.',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  let totalUpdated = 0;

  // ── Update PROJECTS ────────────────────────────────────
  projectSheet.getRange(projectRowIndex, 1).setValue(newProjectId);
  totalUpdated++;

  // ── Update MASTER_SAMPLES ──────────────────────────────
  const masterSheet = ss.getSheetByName('MASTER_SAMPLES');
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][2] === oldProjectId) {
        masterSheet.getRange(i + 1, 3).setValue(newProjectId);
        totalUpdated++;
      }
    }
  }

  // ── Update STATUS_TRACKING ─────────────────────────────
  const statusSheet = ss.getSheetByName('STATUS_TRACKING');
  if (statusSheet) {
    const statusData = statusSheet.getDataRange().getValues();
    for (let i = 1; i < statusData.length; i++) {
      if (statusData[i][3] === oldProjectId) {
        statusSheet.getRange(i + 1, 4).setValue(newProjectId);
        totalUpdated++;
      }
    }
  }

  // ── Update BILLING ─────────────────────────────────────
  const billingSheet = ss.getSheetByName('BILLING');
  if (billingSheet) {
    const billingData = billingSheet.getDataRange().getValues();
    for (let i = 1; i < billingData.length; i++) {
      if (billingData[i][4] === oldProjectId) { // check if ProjectID stored
        totalUpdated++;
      }
    }
  }

  // ── Audit Log ──────────────────────────────────────────
  writeAuditLog(
    'RENAME_PROJECT', 'PROJECTS', oldProjectId,
    'ProjectID', oldProjectId, newProjectId,
    totalUpdated + ' records updated'
  );

  ui.alert(
    '✅ Project Renamed Successfully!\n\n' +
    '• Old ProjectID: ' + oldProjectId + '\n' +
    '• New ProjectID: ' + newProjectId + '\n' +
    '• Records updated: ' + totalUpdated
  );
}


// ============================================================
// UPDATE MENU to include Admin Tools
// Replace your existing setupMenu() with this one
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
      .addItem('✏️ Rename Client', 'renameClient')
      .addItem('✏️ Rename Project', 'renameProject')
      .addItem('🗂️ Setup Client Prefixes', 'setupClientPrefixSheet')
    )
    .addSeparator()
    .addItem('💾 Backup Now', 'backupDatabase')
    .addToUi();
}
