// ============================================================
// SQUADRON TRUCKING — Work Order Tracker
// Google Apps Script (paste into Extensions > Apps Script)
// ============================================================
//
// ARCHITECTURE OVERVIEW
// ─────────────────────
// "Work Orders Import"  — raw CSV drop zone. Your FIL's program
//                         writes here weekly via File > Import.
//                         Never manually edited. Script reads from
//                         this tab but never writes to it.
//
// "Work Orders"         — the working tab. Col A (WO Number) is
//                         static. Cols B–T are live VLOOKUP formulas
//                         that pull from the import tab (so Amazon
//                         dispute updates appear automatically).
//                         Cols U+ are your team's internal tracking
//                         — never touched by any script or formula.
//
// "Archive"             — WOs that rolled off the 6-week window.
//                         Moved here automatically when they
//                         disappear from the import tab.
//
// HOW THE SYNC WORKS
// ──────────────────
// An installable onChange trigger watches the spreadsheet.
// When the import tab changes (new CSV loaded), syncWorkOrders()
// runs automatically. It:
//   1. Reads all WO numbers currently in the import tab
//   2. Adds a new row to Work Orders for any WO not yet present
//      (col A = static WO#, cols B–T = VLOOKUP formulas,
//       cols U+ = blank internal tracking fields)
//   3. Auto-sets Status to "✅ Closed - Not Eligible" for any
//      WO where Score Impact = $0 (exempt rows)
//   4. Moves any Work Orders rows whose WO# is no longer in the
//      import tab over to Archive as a permanent static snapshot
//
// The script NEVER modifies cols U–AC on existing rows.
// Your team's notes and dispute tracking are always safe.
// ============================================================


// ─────────────────────────────────────────────────────────────
// CONFIGURATION
// ─────────────────────────────────────────────────────────────
const CONFIG = {
  IMPORT_SHEET:    'Work Orders Import',
  MASTER_SHEET:    'Work Orders',
  ARCHIVE_SHEET:   'Archive',
  DASHBOARD_SHEET: 'Dashboard',

  // Box API — fill in after completing Box setup guide
  BOX_CLIENT_ID:     'w4nxs7myhnmqf8u062j62ewwtfgg84dn',
  BOX_CLIENT_SECRET: 'UpRq55On9VQAp5LVmqh0Aj4MG2OJo2Ej',
  BOX_ACCESS_TOKEN:  '',

  // CPM targets (update if Amazon changes yours)
  CPM_TARGET: 0.168,
  CPM_GOAL:   0.126,   // 75% of target = "Fantastic Plus"

  MAX_ROWS: 500,       // rows to pre-format with dropdowns
};

// ─────────────────────────────────────────────────────────────
// COLUMN MAP — Work Orders tab (1-based)
// ─────────────────────────────────────────────────────────────
// COLUMN MAP — Work Orders tab (1-based)
//
// VISIBLE UP FRONT (cols 1–16):
//   1–7   Key Amazon fields (WO#, Asset, Vendor, Date, Score, AMZ Status/Result)
//   8–16  Your team's internal tracking
//
// REFERENCE / HIDDEN (cols 17–29):
//   Remaining Amazon fields — still used by VLOOKUPs and Box
//   search, just hidden from normal view
// ─────────────────────────────────────────────────────────────
const COL = {
  // ── Visible columns (1–16) ──
  WO_NUMBER:           1,   // A — static anchor
  ASSET_ID:            2,   // B — formula
  VENDOR:              3,   // C — formula
  WO_START_DATE:       4,   // D — formula  (date)
  INVOICE_POST_EXEMPT: 5,   // E — formula  (currency) ← Score Impact
  AMZ_DISPUTE_STATUS:  6,   // F — formula
  AMZ_DISPUTE_DETERM:  7,   // G — formula

  // Internal tracking (cols 8–16) — manual, never overwritten
  STATUS:              8,   // H
  DISPUTE_CATEGORY:    9,   // I
  DISPUTE_DATE_FILED:  10,  // J
  DISPUTE_OUTCOME:     11,  // K
  INVOICE_FOUND_BOX:   12,  // L
  BOX_LINK:            13,  // M
  NOTES:               14,  // N
  LAST_UPDATED:        15,  // O
  FIRST_SEEN:          16,  // P

  // ── Reference columns (17–29) — hidden but active ──
  WO_END_DATE:         17,  // Q — formula
  ASSET_AGE_DAYS:      18,  // R — formula
  FUEL_TYPE:           19,  // S — formula
  INVOICE_NUMBER:      20,  // T — formula
  INVOICE_DATE:        21,  // U — formula  (date)
  PROCESSING_TS:       22,  // V — formula  (datetime)
  REPORT_WEEK:         23,  // W — formula
  AMZ_DISPUTE_OUTCOME: 24,  // X — formula  (currency)
  INVOICE_AMOUNT:      25,  // Y — formula  (currency)
  INVOICE_REVISED:     26,  // Z — formula  (currency)
  EXEMPTION_REASON:    27,  // AA — formula
  T6W_INDICATOR:       28,  // AB — formula
  ALLOWANCE_PERIOD:    29,  // AC — formula
};

// Import tab column positions (0-based) — never changes
const IMP = {
  WO_NUMBER: 0, VENDOR: 1, WO_START_DATE: 2, WO_END_DATE: 3,
  ASSET_ID: 4, ASSET_AGE_DAYS: 5, FUEL_TYPE: 6, INVOICE_NUMBER: 7,
  INVOICE_DATE: 8, PROCESSING_TS: 9, REPORT_WEEK: 10,
  AMZ_DISPUTE_STATUS: 11, AMZ_DISPUTE_DETERM: 12, AMZ_DISPUTE_OUTCOME: 13,
  INVOICE_AMOUNT: 14, INVOICE_REVISED: 15, EXEMPTION_REASON: 16,
  INVOICE_POST_EXEMPT: 17, T6W_INDICATOR: 18, ALLOWANCE_PERIOD: 19,
};

const TOTAL_COLS       = 29;
const AMAZON_COL_COUNT = 20;  // 20 Amazon fields total
const FIRST_INTERNAL   = 8;   // first column owned by your team
const FIRST_HIDDEN     = 17;  // first hidden reference column


// ─────────────────────────────────────────────────────────────
// MENU
// ─────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ WO Tracker')
    .addItem('🔄 Sync Work Orders Now', 'syncWorkOrders')
    .addSeparator()
    .addItem('🔍 Find Invoice in Box (selected row)', 'findInvoiceInBox')
    .addItem('🔍 Find All Unsearched Invoices', 'findAllMissingInvoices')
    .addSeparator()
    .addItem('📊 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('🏗️ Setup: Initialize Sheet Structure', 'initializeSheetStructure')
    .addItem('⚡ Setup: Install Auto-Sync Trigger', 'installSyncTrigger')
    .addSeparator()
    .addItem('🔐 Setup: Authorize Box (one-time OAuth)', 'authorizeBox')
    .addItem('🔑 Setup: Check Box Auth Status', 'checkBoxAuthStatus')
    .addToUi();
}


// ─────────────────────────────────────────────────────────────
// ONE-TIME SETUP
// ─────────────────────────────────────────────────────────────
function initializeSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tabs = [
    CONFIG.IMPORT_SHEET,
    CONFIG.MASTER_SHEET,
    CONFIG.ARCHIVE_SHEET,
    CONFIG.DASHBOARD_SHEET,
    '📋 How-To: Tires',
    '📋 How-To: Tows',
    '📋 How-To: Preventive Maint.',
    '📋 How-To: Overcharges',
    '📋 How-To: Repeat Repairs',
  ];
  tabs.forEach(name => { if (!ss.getSheetByName(name)) ss.insertSheet(name); });

  setupImportSheet_(ss);
  setupMasterSheet_(ss);
  setupArchiveSheet_(ss);
  setupDashboard_(ss);
  setupHowToTires_(ss);
  setupHowToTows_(ss);
  setupHowToPM_(ss);
  setupHowToOvercharges_(ss);
  setupHowToRepeatRepairs_(ss);

  SpreadsheetApp.getUi().alert(
    '✅ Sheet structure initialized!\n\n' +
    'Next steps:\n' +
    '1. Load your first CSV into "Work Orders Import" via File → Import\n' +
    '   (choose "Replace current sheet" when prompted)\n' +
    '2. Run WO Tracker → Sync Work Orders Now\n' +
    '3. Run WO Tracker → Setup: Install Auto-Sync Trigger\n' +
    '   (future imports will sync automatically after this)'
  );
}

function installSyncTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove duplicates first
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onImportChange')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onImportChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Auto-sync trigger installed!\n\n' +
    'Whenever your FIL\'s program loads a new CSV into\n' +
    '"Work Orders Import", the Work Orders tab will\n' +
    'sync automatically within a few seconds.'
  );
}

// Trigger handler — fires on any sheet change, sync is fast & idempotent
function onImportChange(e) {
  syncWorkOrders();
}


// ─────────────────────────────────────────────────────────────
// CORE SYNC
// Safe to run multiple times. Never modifies internal cols.
// ─────────────────────────────────────────────────────────────
function syncWorkOrders() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet  = ss.getSheetByName(CONFIG.IMPORT_SHEET);
  const masterSheet  = ss.getSheetByName(CONFIG.MASTER_SHEET);
  const archiveSheet = ss.getSheetByName(CONFIG.ARCHIVE_SHEET);

  if (!importSheet || !masterSheet || !archiveSheet) return;

  const importLastRow = importSheet.getLastRow();
  if (importLastRow < 2) return;

  // Read import tab (skip row 1 = header)
  const importData = importSheet
    .getRange(2, 1, importLastRow - 1, AMAZON_COL_COUNT)
    .getValues();

  // Build set of WO numbers currently in import
  const importWOSet = new Set();
  importData.forEach(row => {
    const wo = String(row[IMP.WO_NUMBER]).trim();
    if (wo) importWOSet.add(wo);
  });

  // Build map of WO numbers already in Work Orders → sheet row number
  const masterLastRow = masterSheet.getLastRow();
  const existingWOMap = {};
  if (masterLastRow >= 2) {
    masterSheet
      .getRange(2, COL.WO_NUMBER, masterLastRow - 1, 1)
      .getValues()
      .forEach((row, i) => {
        const wo = String(row[0]).trim();
        if (wo) existingWOMap[wo] = i + 2;
      });
  }

  const now = new Date();
  let added = 0, archived = 0;

  // ── STEP 1: Add new WOs ──
  importData.forEach(row => {
    const woNum = String(row[IMP.WO_NUMBER]).trim();
    if (!woNum || existingWOMap[woNum]) return;

    const scoreImpact  = parseFloat(row[IMP.INVOICE_POST_EXEMPT]) || 0;
    const exemptReason = String(row[IMP.EXEMPTION_REASON] || '').trim();
    const initialStatus = scoreImpact === 0 ? '✅ Closed - Not Eligible' : '🔴 Needs Review';
    const autoNote      = (scoreImpact === 0 && exemptReason) ? `Auto-closed: ${exemptReason}` : '';

    masterSheet.appendRow(buildFormulaRow_(woNum, now, initialStatus, autoNote));
    added++;
  });

  // ── STEP 2: Archive WOs that dropped off ──
  const currentLastRow = masterSheet.getLastRow();
  if (currentLastRow >= 2) {
    const allMasterData = masterSheet
      .getRange(2, 1, currentLastRow - 1, TOTAL_COLS)
      .getValues();

    for (let i = allMasterData.length - 1; i >= 0; i--) {
      const woNum = String(allMasterData[i][COL.WO_NUMBER - 1]).trim();
      if (woNum && !importWOSet.has(woNum)) {
        const sheetRow    = i + 2;
        const archiveDate = Utilities.formatDate(now, 'America/Los_Angeles', 'M/d/yyyy');
        const existingNote = String(allMasterData[i][COL.NOTES - 1] || '');
        allMasterData[i][COL.NOTES - 1] = existingNote
          ? `${existingNote} [Archived ${archiveDate}]`
          : `[Archived ${archiveDate}]`;

        archiveSheet.appendRow(materializeRow_(masterSheet, sheetRow, allMasterData[i]));
        masterSheet.deleteRow(sheetRow);
        archived++;
      }
    }
  }

  applyConditionalFormatting_();
  refreshDashboard();

  // Show alert only when run manually (trigger context has no UI)
  try {
    if (added > 0 || archived > 0) {
      SpreadsheetApp.getUi().alert(
        `🔄 Sync Complete\n\n` +
        `• ${added} new work order(s) added\n` +
        `• ${archived} work order(s) archived\n\n` +
        (added > 0 ? 'Review the 🔴 Needs Review rows in Work Orders.' : '')
      );
    }
  } catch(e) { /* silent when running from trigger */ }
}

// Build a full TOTAL_COLS row for a brand-new Work Orders entry.
// Col 1 = static WO number.
// All other Amazon cols = VLOOKUP formulas keyed on WO number.
// Internal cols = seeded with initial values.
//
// VLOOKUP offset = the column position in the import tab (1-based).
// Import tab layout exactly matches the original CSV column order.
function buildFormulaRow_(woNum, now, initialStatus, autoNote) {
  const imp = CONFIG.IMPORT_SHEET;
  // Try text match first, fall back to numeric — handles both import formats
  const vl = (impColOffset) =>
    `=IFERROR(VLOOKUP("${woNum}",'${imp}'!$A:$T,${impColOffset},FALSE),` +
    `IFERROR(VLOOKUP(VALUE("${woNum}"),'${imp}'!$A:$T,${impColOffset},FALSE),""))`;

  const row = new Array(TOTAL_COLS).fill('');

  // ── Col 1: static WO anchor ──
  row[COL.WO_NUMBER - 1]           = woNum;

  // ── Visible Amazon cols (2–7): VLOOKUP offsets match import tab ──
  row[COL.ASSET_ID - 1]            = vl(5);   // import col E
  row[COL.VENDOR - 1]              = vl(2);   // import col B
  row[COL.WO_START_DATE - 1]       = vl(3);   // import col C
  row[COL.INVOICE_POST_EXEMPT - 1] = vl(18);  // import col R
  row[COL.AMZ_DISPUTE_STATUS - 1]  = vl(12);  // import col L
  row[COL.AMZ_DISPUTE_DETERM - 1]  = vl(13);  // import col M

  // ── Hidden reference cols (17–29) ──
  row[COL.WO_END_DATE - 1]         = vl(4);   // import col D
  row[COL.ASSET_AGE_DAYS - 1]      = vl(6);   // import col F
  row[COL.FUEL_TYPE - 1]           = vl(7);   // import col G
  row[COL.INVOICE_NUMBER - 1]      = vl(8);   // import col H
  row[COL.INVOICE_DATE - 1]        = vl(9);   // import col I
  row[COL.PROCESSING_TS - 1]       = vl(10);  // import col J
  row[COL.REPORT_WEEK - 1]         = vl(11);  // import col K
  row[COL.AMZ_DISPUTE_OUTCOME - 1] = vl(14);  // import col N
  row[COL.INVOICE_AMOUNT - 1]      = vl(15);  // import col O
  row[COL.INVOICE_REVISED - 1]     = vl(16);  // import col P
  row[COL.EXEMPTION_REASON - 1]    = vl(17);  // import col Q
  row[COL.T6W_INDICATOR - 1]       = vl(19);  // import col S
  row[COL.ALLOWANCE_PERIOD - 1]    = vl(20);  // import col T

  // ── Internal tracking — seeded once, never overwritten ──
  row[COL.STATUS - 1]              = initialStatus;
  row[COL.DISPUTE_CATEGORY - 1]    = '';
  row[COL.DISPUTE_DATE_FILED - 1]  = '';
  row[COL.DISPUTE_OUTCOME - 1]     = '';
  row[COL.INVOICE_FOUND_BOX - 1]   = 'Not Searched Yet';
  row[COL.BOX_LINK - 1]            = '';
  row[COL.NOTES - 1]               = autoNote;
  row[COL.LAST_UPDATED - 1]        = now;
  row[COL.FIRST_SEEN - 1]          = now;

  return row;
}

function materializeRow_(sheet, sheetRow, rowValues) {
  const staticRow = [...rowValues];
  // Cols 2–20 are formulas — capture their current displayed values
  const displayed = sheet
    .getRange(sheetRow, 2, 1, AMAZON_COL_COUNT - 1)
    .getDisplayValues()[0];
  displayed.forEach((val, i) => { staticRow[i + 1] = val; });
  return staticRow;
}


// ─────────────────────────────────────────────────────────────
// SHEET SETUP FUNCTIONS
// ─────────────────────────────────────────────────────────────
function setupImportSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG.IMPORT_SHEET);

  // IMPORTANT: never clear this tab if it already has CSV data.
  // Only set up the placeholder banner if the sheet is empty.
  if (sheet.getLastRow() > 1) {
    // Data already present — just ensure column width is set, leave everything else alone
    sheet.setColumnWidth(1, 800);
    return;
  }

  sheet.clearContents();
  sheet.clearFormats();
  sheet.setColumnWidth(1, 800);

  sheet.getRange('A1')
    .setValue(
      'WORK ORDERS IMPORT — Raw data from Quicksite. ' +
      'Load new CSVs here via File → Import → Replace current sheet. ' +
      'Do not edit manually. The Work Orders tab reads from here automatically via VLOOKUP.'
    )
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setFontStyle('italic')
    .setFontSize(10)
    .setWrap(true);
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
}

function setupMasterSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET);
  sheet.clearContents();
  sheet.clearFormats();

  // ── Headers: visible (1–16) then hidden reference (17–29) ──
  const headers = [
    // Visible
    'WO #', 'Asset ID', 'Vendor', 'WO Date', 'Score Impact',
    'AMZ Status', 'AMZ Result',
    '⚡ Status', '📂 Category', '📅 Filed', '🏆 Outcome',
    '📄 In Box?', '🔗 Box Link', '📝 Notes', '🕒 Updated', '📆 First Seen',
    // Hidden reference
    'WO End Date', 'Asset Age', 'Fuel', 'Invoice #',
    'Invoice Date', 'Proc. Timestamp', 'Rpt Wk',
    'AMZ Outcome $', 'Inv. Amount', 'Inv. Revised', 'Exemption Reason',
    'T6W', 'Allowance Period',
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Header bar — full navy
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a3a5c').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10)
    .setWrap(true).setVerticalAlignment('middle');

  // Internal tracking zone — green header
  sheet.getRange(1, FIRST_INTERNAL, 1, FIRST_HIDDEN - FIRST_INTERNAL)
    .setBackground('#2d6a4f');

  // Hidden reference zone — dark grey header
  sheet.getRange(1, FIRST_HIDDEN, 1, TOTAL_COLS - FIRST_HIDDEN + 1)
    .setBackground('#444444');

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);  // freeze WO# so it stays visible when scrolling
  sheet.setRowHeight(1, 45);

  // ── Column widths (visible cols) ──
  sheet.setColumnWidth(COL.WO_NUMBER,           130);
  sheet.setColumnWidth(COL.ASSET_ID,             75);
  sheet.setColumnWidth(COL.VENDOR,               80);
  sheet.setColumnWidth(COL.WO_START_DATE,        85);
  sheet.setColumnWidth(COL.INVOICE_POST_EXEMPT,  90);
  sheet.setColumnWidth(COL.AMZ_DISPUTE_STATUS,  110);
  sheet.setColumnWidth(COL.AMZ_DISPUTE_DETERM,  110);
  sheet.setColumnWidth(COL.STATUS,              155);
  sheet.setColumnWidth(COL.DISPUTE_CATEGORY,    135);
  sheet.setColumnWidth(COL.DISPUTE_DATE_FILED,   85);
  sheet.setColumnWidth(COL.DISPUTE_OUTCOME,     120);
  sheet.setColumnWidth(COL.INVOICE_FOUND_BOX,   100);
  sheet.setColumnWidth(COL.BOX_LINK,            180);
  sheet.setColumnWidth(COL.NOTES,               220);
  sheet.setColumnWidth(COL.LAST_UPDATED,        100);
  sheet.setColumnWidth(COL.FIRST_SEEN,           90);

  // ── Hide reference columns (17–29) ──
  sheet.hideColumns(FIRST_HIDDEN, TOTAL_COLS - FIRST_HIDDEN + 1);

  // ── Date & currency formats on data rows ──
  const n = CONFIG.MAX_ROWS;
  sheet.getRange(2, COL.WO_START_DATE,       n, 1).setNumberFormat('M/d/yyyy');
  sheet.getRange(2, COL.INVOICE_DATE,        n, 1).setNumberFormat('M/d/yyyy');
  sheet.getRange(2, COL.WO_END_DATE,         n, 1).setNumberFormat('M/d/yyyy');
  sheet.getRange(2, COL.PROCESSING_TS,       n, 1).setNumberFormat('M/d/yyyy h:mm am/pm');
  sheet.getRange(2, COL.INVOICE_POST_EXEMPT, n, 1).setNumberFormat('"$"#,##0.00');
  sheet.getRange(2, COL.INVOICE_AMOUNT,      n, 1).setNumberFormat('"$"#,##0.00');
  sheet.getRange(2, COL.INVOICE_REVISED,     n, 1).setNumberFormat('"$"#,##0.00');
  sheet.getRange(2, COL.AMZ_DISPUTE_OUTCOME, n, 1).setNumberFormat('"$"#,##0.00');

  // ── Warning-only protection on all Amazon formula cells ──
  // Covers both visible (cols 2–7) and hidden (cols 17–29)
  sheet.getRange(2, 2, n, 6)
    .protect().setDescription('Amazon data — auto-populated by formula.').setWarningOnly(true);
  sheet.getRange(2, FIRST_HIDDEN, n, TOTAL_COLS - FIRST_HIDDEN + 1)
    .protect().setDescription('Amazon reference data — auto-populated by formula.').setWarningOnly(true);

  addMasterDropdowns_(sheet);
}

function addMasterDropdowns_(sheet) {
  const n = CONFIG.MAX_ROWS;

  sheet.getRange(2, COL.STATUS, n).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([
      '🔴 Needs Review','✅ Closed - Not Eligible',
      '📤 Closed - Disputed','⏳ Awaiting Outcome',
      '⚠️ Invoice Not Found in Box',
    ], true).setAllowInvalid(false).build()
  );

  sheet.getRange(2, COL.DISPUTE_CATEGORY, n).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([
      '','Tire','Tow','Preventive Maintenance','Overcharge','Repeat Repair',
    ], true).setAllowInvalid(false).build()
  );

  sheet.getRange(2, COL.DISPUTE_OUTCOME, n).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([
      '','🏆 Won - Full','🥈 Won - Partial','❌ Denied','⏳ Pending',
    ], true).setAllowInvalid(false).build()
  );

  sheet.getRange(2, COL.INVOICE_FOUND_BOX, n).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([
      '','Yes','No - Not Found in Box','Not Searched Yet',
    ], true).setAllowInvalid(false).build()
  );
}

function applyConditionalFormatting_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET);
  const last  = Math.max(sheet.getLastRow(), 2);
  const range = sheet.getRange(2, 1, last - 1, FIRST_HIDDEN - 1);  // visible cols only
  const scoreRange = sheet.getRange(2, COL.INVOICE_POST_EXEMPT, last - 1, 1);
  const S = colLetter_(COL.STATUS);

  sheet.clearConditionalFormatRules();
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${S}2="🔴 Needs Review"`)
      .setBackground('#fce8e6').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${S}2="⏳ Awaiting Outcome"`)
      .setBackground('#fff9c4').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${S}2="✅ Closed - Not Eligible"`)
      .setBackground('#f1f3f4').setFontColor('#9aa0a6').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${S}2="📤 Closed - Disputed"`)
      .setBackground('#e8f0fe').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${S}2="⚠️ Invoice Not Found in Box"`)
      .setBackground('#fef3e2').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(1000)
      .setBold(true).setFontColor('#c5221f').setRanges([scoreRange]).build(),
  ]);
}

function setupArchiveSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG.ARCHIVE_SHEET);
  sheet.clearContents();
  sheet.clearFormats();

  sheet.getRange(1,1)
    .setValue(
      'ARCHIVE — WOs that rolled off the 6-week window. ' +
      'Values are static snapshots (formulas replaced with values at archive time). ' +
      'Do not edit manually. Filter by Asset ID to find prior repairs for repeat-repair disputes.'
    )
    .setBackground('#f1f3f4').setFontColor('#777777')
    .setFontStyle('italic').setFontSize(10).setWrap(true);
  sheet.setRowHeight(1, 36);

  // Copy header row from master
  const master = ss.getSheetByName(CONFIG.MASTER_SHEET);
  if (master) {
    sheet.getRange(2, 1, 1, TOTAL_COLS)
      .setValues(master.getRange(1, 1, 1, TOTAL_COLS).getValues());
    sheet.getRange(2, 1, 1, TOTAL_COLS)
      .setBackground('#666666').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(2);
  }
}


// ─────────────────────────────────────────────────────────────
// BOX OAUTH — Persistent token management
//
// HOW IT WORKS:
// 1. First time: run "Setup: Authorize Box" from the menu.
//    You'll be shown a URL — open it, approve access, then
//    paste the code back into the prompt. This stores a
//    refresh token in PropertiesService (secure, per-user).
// 2. Every Box API call uses getBoxAccessToken_() which
//    automatically exchanges the refresh token for a fresh
//    access token. Refresh tokens renew on every use so
//    as long as the sheet is used at least once every 60
//    days, it never needs to be re-authorized.
// 3. Fallback: if BOX_ACCESS_TOKEN is set in CONFIG (dev
//    token), that is used instead — so dev token still works
//    for quick testing.
// ─────────────────────────────────────────────────────────────

function getBoxAccessToken_() {
  // If a manual dev token is set in CONFIG, use it (for testing)
  if (CONFIG.BOX_ACCESS_TOKEN) return CONFIG.BOX_ACCESS_TOKEN;

  const props = PropertiesService.getUserProperties();
  const refreshToken = props.getProperty('BOX_REFRESH_TOKEN');

  if (!refreshToken) return null;  // not authorized yet

  // Exchange refresh token for new access token
  try {
    const resp = UrlFetchApp.fetch('https://api.box.com/oauth2/token', {
      method: 'POST',
      payload: {
        grant_type:    'refresh_token',
        refresh_token: refreshToken,
        client_id:     CONFIG.BOX_CLIENT_ID,
        client_secret: CONFIG.BOX_CLIENT_SECRET,
      },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      // Refresh token expired (>60 days unused) — need re-auth
      props.deleteProperty('BOX_REFRESH_TOKEN');
      return null;
    }

    const json = JSON.parse(resp.getContentText());
    // Store the new refresh token (it rotates on every use)
    props.setProperty('BOX_REFRESH_TOKEN', json.refresh_token);
    return json.access_token;

  } catch(e) {
    return null;
  }
}

// Step 1 of OAuth: generate the authorization URL and show it to the user
function authorizeBox() {
  const ui = SpreadsheetApp.getUi();

  if (!CONFIG.BOX_CLIENT_ID || !CONFIG.BOX_CLIENT_SECRET) {
    ui.alert('❌ BOX_CLIENT_ID and BOX_CLIENT_SECRET must be set in CONFIG before authorizing.');
    return;
  }

  const authUrl = 'https://account.box.com/api/oauth2/authorize' +
    `?response_type=code` +
    `&client_id=${CONFIG.BOX_CLIENT_ID}` +
    `&redirect_uri=https://script.google.com/oauthcallback` +
    `&state=squadron_wo_tracker`;

  const result = ui.prompt(
    '🔐 Authorize Box — Step 1 of 2',
    'Open this URL in your browser, approve access, then copy the "code" parameter ' +
    'from the redirect URL and paste it below.\n\n' + authUrl,
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const code = result.getResponseText().trim();
  if (!code) { ui.alert('No code entered. Authorization cancelled.'); return; }

  exchangeBoxCode_(code);
}

// Step 2 of OAuth: exchange the auth code for access + refresh tokens
function exchangeBoxCode_(code) {
  const ui = SpreadsheetApp.getUi();
  try {
    const resp = UrlFetchApp.fetch('https://api.box.com/oauth2/token', {
      method: 'POST',
      payload: {
        grant_type:   'authorization_code',
        code:          code,
        client_id:     CONFIG.BOX_CLIENT_ID,
        client_secret: CONFIG.BOX_CLIENT_SECRET,
        redirect_uri:  'https://script.google.com/oauthcallback',
      },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) {
      ui.alert('❌ Authorization failed. The code may have expired (they last ~30 seconds).\n\nRun "Authorize Box" again and paste the code immediately.');
      return;
    }

    const json = JSON.parse(resp.getContentText());
    PropertiesService.getUserProperties().setProperty('BOX_REFRESH_TOKEN', json.refresh_token);

    ui.alert('✅ Box authorized successfully!\n\nYour credentials are stored securely. You will not need to do this again unless the sheet goes unused for more than 60 days.');
  } catch(e) {
    ui.alert('❌ Error during authorization: ' + e.message);
  }
}

// Check current auth status
function checkBoxAuthStatus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
  const hasRefresh = !!props.getProperty('BOX_REFRESH_TOKEN');
  const hasDevToken = !!CONFIG.BOX_ACCESS_TOKEN;

  if (hasDevToken) {
    ui.alert('🔑 Using manual Developer Token from CONFIG.\n\nThis expires every 60 minutes. Run Authorize Box to set up permanent OAuth.');
  } else if (hasRefresh) {
    const token = getBoxAccessToken_();
    if (token) {
      ui.alert('✅ Box OAuth is active. Token refreshes automatically — no action needed.');
    } else {
      ui.alert('⚠️ Refresh token expired (sheet unused >60 days).\n\nRun Authorize Box to re-authorize.');
    }
  } else {
    ui.alert('❌ Box is not authorized. Run Setup: Authorize Box from the WO Tracker menu.');
  }
}


// ─────────────────────────────────────────────────────────────
// BOX API — Find Invoice (3-pass search)
//
// Pass 1: asset ID + date  → exact match, link written
// Pass 2: asset ID + formatted dollar amount ($x,xxx.xx)
//         → used when Pass 1 returns nothing
// Pass 3: asset ID alone   → last resort, no link written
//         (too low confidence to be useful)
// ─────────────────────────────────────────────────────────────
function findInvoiceInBox() {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET);
  const ui     = SpreadsheetApp.getUi();
  const active = sheet.getActiveRange().getRow();

  if (active < 2) { ui.alert('Click on a data row first.'); return; }

  const row = sheet.getRange(active, 1, 1, TOTAL_COLS).getValues()[0];
  const assetId   = String(row[COL.ASSET_ID - 1]).trim();
  const startDate = row[COL.WO_START_DATE - 1];
  const amount    = row[COL.INVOICE_AMOUNT - 1];

  if (!assetId) { ui.alert('No Asset ID on this row.'); return; }

  const token = getBoxAccessToken_();
  if (!token) {
    ui.alert('❌ Box is not authorized. Run WO Tracker Setup: Authorize Box first.');
    return;
  }

  searchBoxForInvoice_(sheet, active, assetId, startDate, amount, token);
  sheet.getRange(active, COL.LAST_UPDATED).setValue(new Date());
}

function findAllMissingInvoices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET);
  const ui    = SpreadsheetApp.getUi();
  if (sheet.getLastRow() < 2) { ui.alert('No data rows found.'); return; }

  const token = getBoxAccessToken_();
  if (!token) {
    ui.alert('❌ Box is not authorized. Run WO Tracker Setup: Authorize Box first.');
    return;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, TOTAL_COLS).getValues();
  let searched = 0;
  data.forEach((row, i) => {
    const box = String(row[COL.INVOICE_FOUND_BOX - 1]).trim();
    if (box === 'Not Searched Yet' || box === '') {
      const r       = i + 2;
      const assetId = String(row[COL.ASSET_ID - 1]).trim();
      if (assetId) {
        Utilities.sleep(300);
        searchBoxForInvoice_(sheet, r, assetId,
          row[COL.WO_START_DATE - 1], row[COL.INVOICE_AMOUNT - 1], token);
        sheet.getRange(r, COL.LAST_UPDATED).setValue(new Date());
        searched++;
      }
    }
  });
  ui.alert(`✅ Done! Searched Box for ${searched} invoice(s).`);
}

function searchBoxForInvoice_(sheet, sheetRow, assetId, startDateRaw, amount, token) {
  try {
    // ── Build all date format variants ──
    const dateObj = new Date(startDateRaw);
    const mRaw  = parseInt(Utilities.formatDate(dateObj, 'America/Los_Angeles', 'M'));
    const dRaw  = parseInt(Utilities.formatDate(dateObj, 'America/Los_Angeles', 'd'));
    const yyyy  = Utilities.formatDate(dateObj, 'America/Los_Angeles', 'yyyy');
    const yy    = Utilities.formatDate(dateObj, 'America/Los_Angeles', 'yy');
    const mmDD  = Utilities.formatDate(dateObj, 'America/Los_Angeles', 'MM');
    const ddDD  = Utilities.formatDate(dateObj, 'America/Los_Angeles', 'dd');

    // All dot variants (4-digit and 2-digit year)
    const dotVariants = [
      `${mmDD}.${ddDD}.${yyyy}`,  // 03.19.2026
      `${mRaw}.${ddDD}.${yyyy}`,  // 3.19.2026
      `${mmDD}.${dRaw}.${yyyy}`,  // 03.19.2026
      `${mRaw}.${dRaw}.${yyyy}`,  // 3.19.2026
      `${mmDD}.${ddDD}.${yy}`,    // 03.19.26  ← confirmed Kooner format
      `${mRaw}.${ddDD}.${yy}`,    // 3.19.26
      `${mmDD}.${dRaw}.${yy}`,    // 03.19.26
      `${mRaw}.${dRaw}.${yy}`,    // 3.19.26
    ];
    const slashVariants = [
      Utilities.formatDate(dateObj, 'America/Los_Angeles', 'MM/dd/yyyy'),
      Utilities.formatDate(dateObj, 'America/Los_Angeles', 'yyyy-MM-dd'),
    ];

    // Dollar amount formatted exactly as it appears in PDF text: $3,606.39
    const amtNum     = parseFloat(amount) || 0;
    const amtDollar  = '$' + amtNum.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
    const amtPlain   = amtNum.toFixed(2);           // 3606.39
    const amtRound   = Math.round(amtNum).toString(); // 3606

    const existingNote = sheet.getRange(sheetRow, COL.NOTES).getValue();

    // ── PASS 1: Asset ID + date (2-digit year — most common in filenames) ──
    const pass1Query = `"${assetId}" "${dotVariants[4]}"`;  // mm.dd.yy
    const pass1Result = boxSearch_(pass1Query, token);

    if (pass1Result === null) { handleBoxError_(sheet, sheetRow, 'API error on pass 1'); return; }

    if (pass1Result.length > 0) {
      const match = scoreDateMatch_(pass1Result, assetId, dotVariants, slashVariants, amtPlain, amtRound);
      if (match.hasAsset && match.hasDate) {
        writeBoxMatch_(sheet, sheetRow, match.best, pass1Result.length, existingNote);
        return;
      }
    }

    // ── PASS 2: Asset ID + 4-digit year date ──
    const pass2Query = `"${assetId}" "${dotVariants[0]}"`;  // mm.dd.yyyy
    const pass2Result = boxSearch_(pass2Query, token);

    if (pass2Result !== null && pass2Result.length > 0) {
      const match = scoreDateMatch_(pass2Result, assetId, dotVariants, slashVariants, amtPlain, amtRound);
      if (match.hasAsset && match.hasDate) {
        writeBoxMatch_(sheet, sheetRow, match.best, pass2Result.length, existingNote);
        return;
      }
    }

    // ── PASS 3: Asset ID + formatted dollar amount + date proximity check ──
    // Box full-text search finds the amount INSIDE the PDF, so we must verify
    // the file's created_at date is within ±10 days of the WO date to avoid
    // false matches where the amount appears as a line item in a different invoice.
    const pass3Query = `"${assetId}" "${amtDollar}"`;
    const pass3Result = boxSearch_(pass3Query, token);

    if (pass3Result !== null && pass3Result.length > 0) {
      const match = scoreAmountMatch_(pass3Result, assetId, amtPlain, amtRound, amtDollar, dateObj);
      if (match.hasAsset && match.hasAmount && match.withinDateWindow) {
        const boxUrl = `https://app.box.com/file/${match.best.id}`;
        sheet.getRange(sheetRow, COL.INVOICE_FOUND_BOX).setValue('Yes');
        sheet.getRange(sheetRow, COL.BOX_LINK).setValue(boxUrl);
        sheet.getRange(sheetRow, COL.NOTES).setValue(
          appendNote_(existingNote,
            `[✅ Box match (via amount ${amtDollar}): ${match.best.name} | ${pass3Result.length} result(s)]`));
        return;
      }
    }

    // ── All passes failed — check if asset exists at all ──
    const assetOnlyResult = boxSearch_(`"${assetId}"`, token);
    if (assetOnlyResult !== null && assetOnlyResult.length > 0) {
      // Asset exists in Box but no file matches this date or amount
      sheet.getRange(sheetRow, COL.INVOICE_FOUND_BOX).setValue('No - Not Found in Box');
      sheet.getRange(sheetRow, COL.BOX_LINK).setValue('');
      sheet.getRange(sheetRow, COL.NOTES).setValue(
        appendNote_(existingNote,
          `[⚠️ NOT IN BOX — asset ${assetId} exists in Box but no file matches ` +
          `date ${dotVariants[4]} or amount ${amtDollar}. Invoice may be missing from Box.]`));
    } else {
      // Asset not in Box at all
      sheet.getRange(sheetRow, COL.INVOICE_FOUND_BOX).setValue('No - Not Found in Box');
      sheet.getRange(sheetRow, COL.BOX_LINK).setValue('');
      sheet.getRange(sheetRow, COL.NOTES).setValue(
        appendNote_(existingNote,
          `[❌ NOT IN BOX — no files found for asset ${assetId} at all]`));
    }

  } catch(e) {
    sheet.getRange(sheetRow, COL.NOTES).setValue(
      appendNote_(sheet.getRange(sheetRow, COL.NOTES).getValue(),
                  `[Box error: ${e.message}]`));
  }
}

// Make a Box search API call, return entries array or null on error
function boxSearch_(query, token) {
  try {
    const url = `https://api.box.com/2.0/search?query=${encodeURIComponent(query)}` +
                `&type=file&file_extensions=pdf&limit=10&content_types=file_content,name` +
                `&fields=id,name,created_at,modified_at`;  // request date fields for proximity check
    const resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() === 401) return null;  // token issue
    return JSON.parse(resp.getContentText()).entries || [];
  } catch(e) {
    return null;
  }
}

// Score entries by asset ID + date match, return best result with flags
function scoreDateMatch_(entries, assetId, dotVariants, slashVariants, amtPlain, amtRound) {
  const scored = entries.map(e => {
    const name = (e.name || '').toLowerCase();
    const assetMatch = name.includes(assetId.toLowerCase());
    const dotDateMatch = dotVariants.some(d => name.includes(d.toLowerCase()));
    const slashDateMatch = slashVariants.some(d => name.includes(d.toLowerCase()));
    const amtMatch = name.includes(amtPlain) || name.includes(amtRound);
    let s = 0;
    if (assetMatch)      s += 3;
    if (dotDateMatch)    s += 2;
    else if (slashDateMatch) s += 1;
    if (amtMatch)        s += 1;
    return { ...e, score: s, assetMatch, hasDate: dotDateMatch || slashDateMatch };
  }).sort((a, b) => b.score - a.score);
  const best = scored[0];
  return { best, hasAsset: best.assetMatch, hasDate: best.hasDate };
}

// Score entries by asset ID + amount match + date proximity
// dateObj = the WO start date — file must be within ±10 days
function scoreAmountMatch_(entries, assetId, amtPlain, amtRound, amtDollar, dateObj) {
  const DATE_WINDOW_MS = 10 * 24 * 60 * 60 * 1000;  // 10 days in milliseconds
  const woTime = dateObj.getTime();

  const scored = entries.map(e => {
    const name = (e.name || '').toLowerCase();
    const assetMatch = name.includes(assetId.toLowerCase());
    const amtMatch   = name.includes(amtPlain) || name.includes(amtRound) ||
                       name.includes(amtDollar.toLowerCase());

    // Check file creation date is within ±10 days of WO date
    // Box returns created_at as ISO string e.g. "2026-03-31T12:00:00-07:00"
    let withinDateWindow = false;
    if (e.created_at) {
      const fileTime = new Date(e.created_at).getTime();
      withinDateWindow = Math.abs(fileTime - woTime) <= DATE_WINDOW_MS;
    }

    let s = 0;
    if (assetMatch)       s += 3;
    if (amtMatch)         s += 2;
    if (withinDateWindow) s += 2;
    return { ...e, score: s, assetMatch, hasAmount: amtMatch, withinDateWindow };
  }).sort((a, b) => b.score - a.score);

  const best = scored[0];
  return {
    best,
    hasAsset: best.assetMatch,
    hasAmount: best.hasAmount,
    withinDateWindow: best.withinDateWindow,
  };
}

// Write a confirmed match to the sheet
function writeBoxMatch_(sheet, sheetRow, entry, resultCount, existingNote) {
  const boxUrl = `https://app.box.com/file/${entry.id}`;
  sheet.getRange(sheetRow, COL.INVOICE_FOUND_BOX).setValue('Yes');
  sheet.getRange(sheetRow, COL.BOX_LINK).setValue(boxUrl);
  sheet.getRange(sheetRow, COL.NOTES).setValue(
    appendNote_(existingNote,
      `[✅ Box match: ${entry.name} | ${resultCount} result(s)]`));
}

function handleBoxError_(sheet, sheetRow, msg) {
  sheet.getRange(sheetRow, COL.INVOICE_FOUND_BOX).setValue('Not Searched Yet');
  sheet.getRange(sheetRow, COL.NOTES).setValue(
    appendNote_(sheet.getRange(sheetRow, COL.NOTES).getValue(),
                `[Box error: ${msg}]`));
}


// ─────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────
function setupDashboard_(ss) {
  const sheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);
  sheet.clearContents(); sheet.clearFormats();
  sheet.setColumnWidth(1, 280); sheet.setColumnWidth(2, 160);

  sheet.getRange('A1').setValue('SQUADRON TRUCKING — WO Tracker Dashboard')
    .setFontSize(16).setFontWeight('bold').setFontColor('#1a3a5c');
  sheet.setRowHeight(1, 40);
  sheet.getRange('A2').setValue('Last refreshed:').setFontColor('#777777');
  sheet.getRange('B2').setValue(new Date()).setNumberFormat('M/d/yyyy h:mm am/pm');
  sheet.getRange('B4').setValue(CONFIG.CPM_TARGET).setNumberFormat('$0.000');
  sheet.getRange('B5').setValue(CONFIG.CPM_GOAL).setNumberFormat('$0.000');

  const labels = [
    [4,'CPM TARGET','#1a3a5c',true], [5,'CPM GOAL (Fantastic Plus)','#1a3a5c',true],
    [7,'TOTAL WOs IN WINDOW','#1a3a5c',true],
    [8,'🔴 Needs Review','#c5221f',false], [9,'⏳ Awaiting Outcome','#b45309',false],
    [10,'✅ Closed - Not Eligible','#2d6a4f',false], [11,'📤 Closed - Disputed','#1967d2',false],
    [13,'INVOICES FOUND IN BOX','#1a3a5c',true],
    [14,'INVOICES NOT FOUND IN BOX','#c5221f',false], [15,'NOT YET SEARCHED','#b45309',false],
    [17,'DISPUTES FILED (total)','#1a3a5c',true],
    [18,'Won - Full','#2d6a4f',false], [19,'Won - Partial','#2d6a4f',false],
    [20,'Denied','#c5221f',false], [21,'Pending','#b45309',false],
  ];
  labels.forEach(([r,label,color,bold]) =>
    sheet.getRange(r,1).setValue(label).setFontColor(color).setFontWeight(bold?'bold':'normal'));

  sheet.getRange('A23').setValue('Click WO Tracker → Refresh Dashboard to update counts.')
    .setFontColor('#999999').setFontStyle('italic').setFontSize(9);
}

function refreshDashboard() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.MASTER_SHEET);
  const dash   = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);
  if (!master || !dash) return;

  const data = master.getLastRow() > 1
    ? master.getRange(2, 1, master.getLastRow() - 1, TOTAL_COLS).getValues() : [];

  let total=0, needsReview=0, awaiting=0, notEligible=0, disputed=0;
  let foundBox=0, notFoundBox=0, notSearched=0;
  let wonFull=0, wonPartial=0, denied=0, pending=0;

  data.forEach(row => {
    if (!String(row[COL.WO_NUMBER-1]).trim()) return;
    total++;
    const st  = String(row[COL.STATUS-1]);
    const box = String(row[COL.INVOICE_FOUND_BOX-1]);
    const out = String(row[COL.DISPUTE_OUTCOME-1]);
    if (st.includes('Needs Review')) needsReview++;
    else if (st.includes('Awaiting')) awaiting++;
    else if (st.includes('Not Eligible')) notEligible++;
    else if (st.includes('Disputed')) disputed++;
    if (box==='Yes') foundBox++;
    else if (box.includes('Not Found')) notFoundBox++;
    else notSearched++;
    if (out.includes('Won - Full')) wonFull++;
    else if (out.includes('Won - Partial')) wonPartial++;
    else if (out.includes('Denied')) denied++;
    else if (out.includes('Pending')) pending++;
  });

  dash.getRange('B2').setValue(new Date()).setNumberFormat('M/d/yyyy h:mm am/pm');
  [[7,total],[8,needsReview],[9,awaiting],[10,notEligible],[11,disputed],
   [13,foundBox],[14,notFoundBox],[15,notSearched],
   [17,wonFull+wonPartial+denied+pending],[18,wonFull],[19,wonPartial],[20,denied],[21,pending]]
  .forEach(([r,v]) => dash.getRange(r,2).setValue(v));
  dash.getRange(8,2).setFontColor(needsReview>0?'#c5221f':'#000');
  dash.getRange(14,2).setFontColor(notFoundBox>0?'#c5221f':'#000');
  dash.getRange(15,2).setFontColor(notSearched>0?'#b45309':'#000');
}


// ─────────────────────────────────────────────────────────────
// HOW-TO TABS
// ─────────────────────────────────────────────────────────────
function writeHowToTab_(sheet, content, accentColor) {
  content.forEach((item, i) => {
    const row  = i + 1;
    const cell = sheet.getRange(row, 1);
    if (item.type === 'spacer') { sheet.setRowHeight(row, 8); return; }
    cell.setValue(item.text).setWrap(true);
    switch (item.type) {
      case 'title':
        cell.setFontSize(13).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1a3a5c');
        sheet.setRowHeight(row, 36); break;
      case 'section':
        cell.setFontSize(10).setFontWeight('bold').setFontColor('#1a3a5c').setBackground(accentColor+'44');
        sheet.setRowHeight(row, 24); break;
      case 'body':
        cell.setFontSize(10).setFontColor('#333333'); sheet.setRowHeight(row, 48); break;
      case 'check':
        cell.setFontSize(10).setFontColor('#333333'); sheet.setRowHeight(row, 22); break;
      case 'step':
        cell.setFontSize(10).setFontColor('#1a3a5c'); sheet.setRowHeight(row, 22); break;
      case 'tip':
        cell.setFontSize(10).setFontColor('#2d6a4f').setFontStyle('italic');
        sheet.setRowHeight(row, 22); break;
    }
  });
}

function setupHowToTires_(ss) {
  const sheet = ss.getSheetByName('📋 How-To: Tires');
  sheet.clearContents(); sheet.clearFormats(); sheet.setColumnWidth(1, 720);
  writeHowToTab_(sheet, [
    {text:'TIRE DISPUTES — How-To Guide',type:'title'},
    {text:'',type:'spacer'},
    {text:'WHAT QUALIFIES FOR A TIRE DISPUTE?',type:'section'},
    {text:'Tires may be disputed when charged at above-market rates, when the replaced tire had significant remaining tread life, when the same tire position was replaced recently (possible defective product or warranty), or when the brand/spec charged does not match what was actually installed.',type:'body'},
    {text:'',type:'spacer'},
    {text:'ELIGIBILITY CHECKLIST — answer YES to any of these to consider disputing:',type:'section'},
    {text:'□  Was the price per tire above current market rate for that size/brand?',type:'check'},
    {text:'□  Was the tire replaced before 50% tread wear? (Check prior PM records)',type:'check'},
    {text:'□  Has this same tire position been replaced on this truck within the last 6 months?',type:'check'},
    {text:'□  Was a premium brand charged but a lesser brand installed? (Verify on invoice PDF)',type:'check'},
    {text:'□  Were disposal fees, mounting, or balancing charged at unusually high rates?',type:'check'},
    {text:'',type:'spacer'},
    {text:'EVIDENCE YOU NEED BEFORE FILING:',type:'section'},
    {text:'1. Invoice PDF from Box — confirm actual brand, size, and price per tire',type:'step'},
    {text:'2. Prior PM report showing tread depth at last inspection',type:'step'},
    {text:'3. Market price comps for same tire size (Goodyear fleet portal, TireHub, or supplier quote)',type:'step'},
    {text:'4. If repeat position: prior WO for same axle — pull from Archive tab by filtering Asset ID',type:'step'},
    {text:'',type:'spacer'},
    {text:'HOW TO FILE IN AMAZON QUICKSITE:',type:'section'},
    {text:'1. Quicksite → Maintenance → Work Orders → find WO by WO Number',type:'step'},
    {text:'2. Click WO → click "Dispute" button (top right)',type:'step'},
    {text:'3. Select reason: "Tire - Overcharge" or "Tire - Premature Replacement"',type:'step'},
    {text:'4. Upload evidence (invoice PDF + price comps)',type:'step'},
    {text:'5. Write a clear 2–3 sentence explanation in the comments box',type:'step'},
    {text:'6. Submit — note confirmation number in the Notes column of this sheet',type:'step'},
    {text:'7. Update: Status → "📤 Closed - Disputed" | Category → "Tire" | Dispute Filed date',type:'step'},
    {text:'',type:'spacer'},
    {text:'TIPS:',type:'section'},
    {text:'• Goodyear invoices tend to be the most winnable — pricing errors on their direct submissions are common',type:'tip'},
    {text:'• Always check Score Impact $ (col R) vs Invoice Amount (col O) — if different, a partial dispute may already be applied',type:'tip'},
    {text:'• CNG trucks use different tire specs — compare only against CNG-rated tire pricing',type:'tip'},
  ], '#f4b942');
}

function setupHowToTows_(ss) {
  const sheet = ss.getSheetByName('📋 How-To: Tows');
  sheet.clearContents(); sheet.clearFormats(); sheet.setColumnWidth(1, 720);
  writeHowToTab_(sheet, [
    {text:'TOW DISPUTES — How-To Guide',type:'title'},
    {text:'',type:'spacer'},
    {text:'WHAT QUALIFIES FOR A TOW DISPUTE?',type:'section'},
    {text:'Tows can be disputed when the breakdown was caused by a non-maintenance issue (driver error, accident, fuel problem), when the truck was towed past a closer approved shop, when the rate per mile is above market, or when the tow was to a non-Amazon-approved vendor.',type:'body'},
    {text:'',type:'spacer'},
    {text:'ELIGIBILITY CHECKLIST:',type:'section'},
    {text:'□  Was the breakdown caused by something other than a maintenance failure? (accident, out-of-fuel, driver error)',type:'check'},
    {text:'□  Was the truck towed more than ~25 miles when a closer approved shop was available?',type:'check'},
    {text:'□  Is the per-mile tow rate above market? (typical heavy truck rate: $6–$8/mile)',type:'check'},
    {text:'□  Was the tow vendor not on the Amazon-approved vendor list?',type:'check'},
    {text:'□  Was the truck towed to a non-approved shop when an approved shop was in range?',type:'check'},
    {text:'',type:'spacer'},
    {text:'EVIDENCE YOU NEED BEFORE FILING:',type:'section'},
    {text:'1. Tow invoice from Box — pickup location, drop-off, mileage billed, rate per mile',type:'step'},
    {text:'2. If driver error/accident: incident report or dispatch notes from that day',type:'step'},
    {text:'3. Google Maps screenshot: breakdown location → nearest approved shop vs. actual tow destination',type:'step'},
    {text:'4. Amazon approved vendor list showing what was available in that area',type:'step'},
    {text:'',type:'spacer'},
    {text:'HOW TO FILE IN AMAZON QUICKSITE:',type:'section'},
    {text:'1. Quicksite → Work Orders → find the WO',type:'step'},
    {text:'2. Click "Dispute" → select "Tow - Not Maintenance Related" or "Tow - Excessive Distance/Rate"',type:'step'},
    {text:'3. Upload tow invoice + supporting evidence',type:'step'},
    {text:'4. Comments: state breakdown cause, actual tow distance, nearest available approved shop, and a reasonable charge',type:'step'},
    {text:'5. Submit and log confirmation number in Notes column',type:'step'},
    {text:'6. Update: Status → "📤 Closed - Disputed" | Category → "Tow" | Dispute Filed date',type:'step'},
    {text:'',type:'spacer'},
    {text:'TIPS:',type:'section'},
    {text:'• Tows are typically your highest single-invoice amounts — even a partial win is worth the filing time',type:'tip'},
    {text:'• Driver messages or dispatch logs from the day of breakdown are the strongest possible evidence',type:'tip'},
    {text:'• If towed to an approved shop, focus your dispute on the rate per mile rather than the destination',type:'tip'},
  ], '#e57373');
}

function setupHowToPM_(ss) {
  const sheet = ss.getSheetByName('📋 How-To: Preventive Maint.');
  sheet.clearContents(); sheet.clearFormats(); sheet.setColumnWidth(1, 720);
  writeHowToTab_(sheet, [
    {text:'SCHEDULED PREVENTIVE MAINTENANCE DISPUTES — How-To Guide',type:'title'},
    {text:'',type:'spacer'},
    {text:'WHAT QUALIFIES?',type:'section'},
    {text:'PM disputes apply when a service was billed but not yet due per manufacturer interval or your own PM schedule, when a PM was already completed recently by another vendor, or when add-on services were bundled in that fall outside the standard PM scope.',type:'body'},
    {text:'',type:'spacer'},
    {text:'ELIGIBILITY CHECKLIST:',type:'section'},
    {text:'□  Was this PM performed before the recommended interval since the last PM on this truck?',type:'check'},
    {text:'□  Do you have records showing this PM was already completed by your own vendor?',type:'check'},
    {text:'□  Was the PM interval shorter than the OEM-recommended interval for this engine/mileage?',type:'check'},
    {text:'□  Were add-on services included that are outside the standard PM scope?',type:'check'},
    {text:'',type:'spacer'},
    {text:'EVIDENCE YOU NEED BEFORE FILING:',type:'section'},
    {text:'1. Previous PM invoice for this truck — date and mileage of last service',type:'step'},
    {text:'2. Current mileage at time of disputed PM (from dispatch or telematics)',type:'step'},
    {text:'3. OEM-recommended service intervals for this specific tractor model and engine',type:'step'},
    {text:'4. If duplicate: your vendor\'s invoice showing same service was already performed',type:'step'},
    {text:'',type:'spacer'},
    {text:'HOW TO FILE IN AMAZON QUICKSITE:',type:'section'},
    {text:'1. Quicksite → Work Orders → find the WO',type:'step'},
    {text:'2. Click "Dispute" → select "Scheduled PM - Not Due" or "Scheduled PM - Duplicate"',type:'step'},
    {text:'3. Upload prior PM invoice and current mileage documentation',type:'step'},
    {text:'4. Comments: state last PM date/mileage, interval that should have been used vs. actual interval used',type:'step'},
    {text:'5. Submit and log in this sheet',type:'step'},
    {text:'',type:'spacer'},
    {text:'TIPS:',type:'section'},
    {text:'• A per-truck PM log is the single most powerful tool for winning PM disputes — if you don\'t have one, start now',type:'tip'},
    {text:'• Kenworth and TA vendors tend to PM on their own preferred (shorter) intervals — always cross-check',type:'tip'},
    {text:'• Newer trucks (under 1 year old) often have longer OEM intervals — use that to your advantage',type:'tip'},
  ], '#81c784');
}

function setupHowToOvercharges_(ss) {
  const sheet = ss.getSheetByName('📋 How-To: Overcharges');
  sheet.clearContents(); sheet.clearFormats(); sheet.setColumnWidth(1, 720);
  writeHowToTab_(sheet, [
    {text:'OVERCHARGE DISPUTES — How-To Guide',type:'title'},
    {text:'',type:'spacer'},
    {text:'WHAT QUALIFIES?',type:'section'},
    {text:'Overcharges occur when labor rates exceed Amazon\'s negotiated or market rate for that vendor, when parts are marked up excessively above MSRP, when more labor hours are billed than the repair reasonably requires, or when line items are for work not actually performed.',type:'body'},
    {text:'',type:'spacer'},
    {text:'ELIGIBILITY CHECKLIST:',type:'section'},
    {text:'□  Does the labor rate exceed Amazon\'s published standard rate for this vendor?',type:'check'},
    {text:'□  Are parts billed above MSRP or significantly above market price?',type:'check'},
    {text:'□  Does the labor time seem excessive for this repair type? (e.g., >4 hours for a brake adjustment)',type:'check'},
    {text:'□  Are there line items for work you cannot verify was performed?',type:'check'},
    {text:'□  Does the final invoice significantly exceed any original estimate?',type:'check'},
    {text:'',type:'spacer'},
    {text:'EVIDENCE YOU NEED BEFORE FILING:',type:'section'},
    {text:'1. Full invoice from Box with all line items clearly visible',type:'step'},
    {text:'2. Amazon\'s published labor rate schedule for that vendor (Quicksite vendor portal)',type:'step'},
    {text:'3. MSRP or market price for any parts you believe were overpriced (dealer site, Parts.com, etc.)',type:'step'},
    {text:'4. Industry standard time guide for the repair type (Mitchell1, Alldata, or OEM service manual)',type:'step'},
    {text:'',type:'spacer'},
    {text:'HOW TO FILE IN AMAZON QUICKSITE:',type:'section'},
    {text:'1. Quicksite → Work Orders → find the WO',type:'step'},
    {text:'2. Click "Dispute" → select "Overcharge - Labor Rate", "Overcharge - Parts", or "Overcharge - Excessive Hours"',type:'step'},
    {text:'3. List each overcharged line item: what was billed vs. what should have been billed',type:'step'},
    {text:'4. Upload invoice + your rate/price documentation',type:'step'},
    {text:'5. Submit and log confirmation number in Notes column',type:'step'},
    {text:'',type:'spacer'},
    {text:'TIPS:',type:'section'},
    {text:'• Line-item specificity wins overcharge disputes — "price seems high" rarely succeeds on its own',type:'tip'},
    {text:'• Study rows in your sheet where col N already has a value — those partially granted disputes are your templates',type:'tip'},
    {text:'• TA (TravelCenters) invoices are the most common overcharge target given their volume in your fleet',type:'tip'},
  ], '#64b5f6');
}

function setupHowToRepeatRepairs_(ss) {
  const sheet = ss.getSheetByName('📋 How-To: Repeat Repairs');
  sheet.clearContents(); sheet.clearFormats(); sheet.setColumnWidth(1, 720);
  writeHowToTab_(sheet, [
    {text:'REPEAT REPAIR DISPUTES — How-To Guide',type:'title'},
    {text:'',type:'spacer'},
    {text:'WHAT QUALIFIES?',type:'section'},
    {text:'A repeat repair dispute applies when the same system or component on the same truck was repaired within 30–90 days and failed again — meaning the vendor\'s prior work was not durable and should be covered under a workmanship warranty.',type:'body'},
    {text:'',type:'spacer'},
    {text:'ELIGIBILITY CHECKLIST:',type:'section'},
    {text:'□  Has this truck had the same system/component repaired within the last 90 days?',type:'check'},
    {text:'□  Was the prior repair done by the same vendor? (stronger case, but not required)',type:'check'},
    {text:'□  Is the current repair failure directly related to the prior repair?',type:'check'},
    {text:'□  Can you document both repairs with invoices showing same Asset ID and same system?',type:'check'},
    {text:'',type:'spacer'},
    {text:'HOW TO IDENTIFY REPEAT REPAIRS IN THIS SHEET:',type:'section'},
    {text:'1. Filter the Work Orders tab by Asset ID (col E) to see all WOs for a specific truck',type:'step'},
    {text:'2. Check the Archive tab too — prior repairs that rolled off may still be within 90 days',type:'step'},
    {text:'3. Look for same vendor + same asset + similar repair description within 90 days',type:'step'},
    {text:'4. Sort by Asset ID then WO Start Date to see patterns across your fleet quickly',type:'step'},
    {text:'',type:'spacer'},
    {text:'EVIDENCE YOU NEED BEFORE FILING:',type:'section'},
    {text:'1. Prior WO invoice showing the original repair (Box link from the Archive row)',type:'step'},
    {text:'2. Current WO invoice showing the repeat failure',type:'step'},
    {text:'3. Both invoices must show same Asset ID and same system/component',type:'step'},
    {text:'4. Calculate and note the days elapsed between the two repairs',type:'step'},
    {text:'',type:'spacer'},
    {text:'HOW TO FILE IN AMAZON QUICKSITE:',type:'section'},
    {text:'1. Quicksite → Work Orders → find the CURRENT (new) WO',type:'step'},
    {text:'2. Click "Dispute" → select "Repeat Repair"',type:'step'},
    {text:'3. Reference the prior WO number explicitly in your comments',type:'step'},
    {text:'4. Upload BOTH invoices',type:'step'},
    {text:'5. Comments template: "This is a repeat repair of WO [#] performed on [date] for the same system on Asset [ID]. The repair failed within [X days]. We request this charge be waived under workmanship warranty."',type:'step'},
    {text:'6. Submit and log in this sheet',type:'step'},
    {text:'',type:'spacer'},
    {text:'TIPS:',type:'section'},
    {text:'• Highest win rate when same vendor did both repairs — they have the most incentive to honor their own warranty',type:'tip'},
    {text:'• Amazon takes repeat repairs seriously because they signal a vendor quality issue — your dispute helps them too',type:'tip'},
    {text:'• Your Archive tab becomes more valuable every week — it\'s your repeat repair detection database over time',type:'tip'},
  ], '#ce93d8');
}


// ─────────────────────────────────────────────────────────────
// UTILITIES
// ─────────────────────────────────────────────────────────────
function colLetter_(n) {
  let r = '';
  while (n > 0) { n--; r = String.fromCharCode(65 + (n % 26)) + r; n = Math.floor(n/26); }
  return r;
}

function appendNote_(existing, addition) {
  const e = String(existing || '').trim();
  return e ? `${e} ${addition}` : addition;
}