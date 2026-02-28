/***** CONFIG *****/
const PURCHASES_FOLDER_ID = '1lNJdDrFNR5bvdcCAvqxqM560H2GOvapR';

// Default mode when using "Start Report Wizard"
const TEST_MODE_DEFAULT = true;
const TEST_FOLDER_NAME = 'Generated Reports - TEST';
const LIVE_FOLDER_NAME = 'Generated Reports';

// Report templates copied into Generated Reports folder
const REPORT_TEMPLATES = {
  'DTSC-ALL': 'https://docs.google.com/spreadsheets/d/1OV_6SjNNC1quXXkysNLdRxPmwYAXvBJUz33EQMG3jtg/edit?gid=308431959#gid=308431959',
  'DTSC': 'https://docs.google.com/spreadsheets/d/1QM63U20jXQDHbWvnndL2YY3vD2Ye_EwvXcM1Gqu8Cps/edit?gid=308431959#gid=308431959',
  'NETCOST': 'https://docs.google.com/spreadsheets/d/1ppJeXJCmmFnKYY4YMToX1iWnEJ2T0QDu6u3mbIUaxjo/edit?gid=308431959#gid=308431959',
};

// Source tabs used to refresh invoice PO tabs
const PO_SHEET_SOURCES = {
  'EWASTE-PO': {
    spreadsheetId: '1-UsOuVDDIcBxkiFPzHjGt8XQ9MwPrLaRP7PCzM0h-OU',
    sheetName: '(DTSC)EWASTE-PO',
  },
  'SB20-PO': {
    spreadsheetId: '1azaSdv9dCYGY3wiirSmUJSiuGTncKx5WbNnPy9JUO40',
    sheetName: '(NETCOST)SB20-PO',
  },
};

/***** MENU *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Reports')
    .addItem('Open Wizard Sidebar', 'openWizardSidebar')
    .addSeparator()
    .addItem('Start Report Wizard', 'startReportWizard')
    .addItem('Start Report Wizard (TEST)', 'startReportWizard_TEST')
    .addItem('Start Report Wizard (LIVE)', 'startReportWizard_LIVE')
    .addToUi();
}

function openWizardSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Report Wizard')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function startReportWizard() {
  startReportWizardCore_({ testMode: TEST_MODE_DEFAULT, runId: null });
}

function startReportWizard_TEST() {
  startReportWizardCore_({ testMode: true, runId: null });
}

function startReportWizard_LIVE() {
  startReportWizardCore_({ testMode: false, runId: null });
}

/***** SIDEBAR RUN CONTROL *****/
function beginWizardRun(testMode) {
  const runId = `run_${Date.now()}_${Math.floor(Math.random() * 1000000)}`;
  const state = {
    runId,
    status: 'running',
    progress: 0,
    logs: [`${timestamp_()} Run created (${testMode ? 'TEST' : 'LIVE'})`],
    error: '',
    mode: testMode ? 'TEST' : 'LIVE',
  };
  saveRunState_(runId, state);
  return runId;
}

function runWizardWithProgress(runId, testMode) {
  try {
    setStatus_('Starting wizard...', 2, runId);
    startReportWizardCore_({ testMode: !!testMode, runId });
    setStatus_('Completed successfully.', 100, runId);
    finishRun_(runId);
  } catch (error) {
    failRun_(runId, error);
    throw error;
  }
}

function getWizardRunStatus(runId) {
  return loadRunState_(runId) || {
    runId,
    status: 'not_found',
    progress: 0,
    logs: [`${timestamp_()} No run state found.`],
    error: 'Run not found',
    mode: '-',
  };
}

function saveRunState_(runId, state) {
  PropertiesService.getUserProperties().setProperty(`RW_${runId}`, JSON.stringify(state));
}

function loadRunState_(runId) {
  const raw = PropertiesService.getUserProperties().getProperty(`RW_${runId}`);
  return raw ? JSON.parse(raw) : null;
}

function finishRun_(runId) {
  const state = loadRunState_(runId);
  if (!state) return;
  state.status = 'done';
  state.progress = 100;
  state.logs.push(`${timestamp_()} Finished.`);
  saveRunState_(runId, state);
}

function failRun_(runId, error) {
  const state = loadRunState_(runId) || {
    runId,
    status: 'error',
    progress: 0,
    logs: [],
    error: '',
    mode: '-',
  };
  state.status = 'error';
  state.error = error && error.message ? error.message : String(error);
  state.logs.push(`${timestamp_()} ERROR: ${state.error}`);
  saveRunState_(runId, state);
}

function setStatus_(message, progress, runId) {
  Logger.log(message);

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Report Wizard', 3);
  } catch (_) {}

  if (!runId) return;

  const state = loadRunState_(runId);
  if (!state) return;

  state.logs.push(`${timestamp_()} ${message}`);
  if (typeof progress === 'number') {
    state.progress = Math.max(0, Math.min(100, progress));
  }
  state.logs = state.logs.slice(-400);
  saveRunState_(runId, state);
}

function timestamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
}

/***** MAIN FLOW *****/
function startReportWizardCore_(opts) {
  const ui = SpreadsheetApp.getUi();
  const testMode = !!(opts && opts.testMode);
  const runId = opts && opts.runId ? opts.runId : null;

  try {
    setStatus_('Connecting to Purchases folder...', 5, runId);
    const purchasesRoot = DriveApp.getFolderById(PURCHASES_FOLDER_ID);

    setStatus_('Prompting for customer number...', 8, runId);
    const custResp = ui.prompt(
      'Customer Number',
      'Enter customer number (example: 103 or #199-1):',
      ui.ButtonSet.OK_CANCEL
    );
    if (custResp.getSelectedButton() !== ui.Button.OK) {
      setStatus_('Cancelled by user at customer prompt.', 100, runId);
      return;
    }

    const customerNumber = (custResp.getResponseText() || '').trim();
    if (!customerNumber) {
      ui.alert('Customer number is required.');
      setStatus_('Cancelled: customer number missing.', 100, runId);
      return;
    }

    setStatus_('Finding customer folder...', 12, runId);
    const customerFolder = findCustomerPoFolder_(purchasesRoot, customerNumber);
    if (!customerFolder) {
      ui.alert(
        `No customer folder found for ${customerNumber}.\nExpected names like #${customerNumber}-POs or ${customerNumber}-POs`
      );
      setStatus_('Customer folder not found.', 100, runId);
      return;
    }

    const confirmFolder = ui.alert(
      'Confirm Customer Folder',
      `Mode: ${testMode ? 'TEST' : 'LIVE'}\n\nUse this customer folder?\n\n${customerFolder.getName()}\n${customerFolder.getUrl()}`,
      ui.ButtonSet.YES_NO
    );
    if (confirmFolder !== ui.Button.YES) {
      setStatus_('Cancelled by user at customer folder confirmation.', 100, runId);
      return;
    }

    setStatus_('Prompting SB-20 selection...', 18, runId);
    const sb20Response = ui.alert(
      'SB-20 Items',
      'Does this customer have SB-20 items?',
      ui.ButtonSet.YES_NO
    );
    const hasSb20 = sb20Response === ui.Button.YES;

    const generatedFolderName = testMode ? TEST_FOLDER_NAME : LIVE_FOLDER_NAME;
    setStatus_(`Preparing "${generatedFolderName}" folder...`, 24, runId);
    const generatedReportsFolder = getOrCreateSubfolder_(customerFolder, generatedFolderName);

    const templatesToUse = hasSb20
      ? REPORT_TEMPLATES
      : {
          'DTSC-ALL': REPORT_TEMPLATES['DTSC-ALL'],
          'DTSC': REPORT_TEMPLATES['DTSC'],
        };

    setStatus_('Copying template report files...', 32, runId);
    const reportFilesByKey = copyAllTemplates_(templatesToUse, generatedReportsFolder, customerNumber);

    setStatus_('Prompting for invoice number(s)...', 38, runId);
    const invResp = ui.prompt(
      'Invoice Number(s)',
      'Enter invoice number(s). Use comma or new line for multiple.\nExample: #199-1-01-08-26',
      ui.ButtonSet.OK_CANCEL
    );
    if (invResp.getSelectedButton() !== ui.Button.OK) {
      setStatus_('Cancelled by user at invoice prompt.', 100, runId);
      return;
    }

    const invoiceInputs = (invResp.getResponseText() || '')
      .split(/[\n,]+/)
      .map(s => s.trim())
      .filter(Boolean);

    if (!invoiceInputs.length) {
      ui.alert('No invoice number entered.');
      setStatus_('Cancelled: no invoice numbers entered.', 100, runId);
      return;
    }

    setStatus_(`Searching ${invoiceInputs.length} invoice folder(s)...`, 45, runId);
    const found = [];
    const notFound = [];
    const processed = [];

    invoiceInputs.forEach((invoiceNo) => {
      const match = findInvoiceSheetInPurchases_(purchasesRoot, invoiceNo, customerNumber);
      if (match) found.push({ invoiceNo, folder: match.folder, file: match.file });
      else notFound.push(invoiceNo);
    });

    if (found.length) {
      setStatus_(`Matched ${found.length} invoice(s).`, 52, runId);
      showInvoiceLinksDialog_(found, testMode);

      found.forEach(item => {
        setStatus_(`Processing invoice ${item.invoiceNo}...`, null, runId);
        const invoiceFileToProcess = testMode
          ? createInvoiceTestCopy_(item.file, generatedReportsFolder, customerNumber)
          : item.file;

        const refreshed = upsertPoSheetsInInvoice_(invoiceFileToProcess.getId(), item.invoiceNo, hasSb20, runId);
        if (!refreshed) {
          processed.push(`Skipped (missing PO or user cancelled): ${item.invoiceNo}`);
          return;
        }

        const imported = importInvoiceTabsToGeneratedReports_(
          invoiceFileToProcess.getId(),
          item.invoiceNo,
          reportFilesByKey,
          hasSb20
        );

        processed.push(`${item.invoiceNo} -> ${imported.join(', ') || 'No import tabs found'}`);
      });
    }

    const copiedReportKeys = Object.keys(reportFilesByKey);
    let msg =
      `${testMode ? 'TEST' : 'LIVE'} complete.\n\n` +
      `Generated folder:\n${generatedReportsFolder.getUrl()}\n\n` +
      `Copied ${copiedReportKeys.length} report file(s):\n` +
      copiedReportKeys.map(k => `• ${reportFilesByKey[k].getName()}`).join('\n');

    if (processed.length) {
      msg += `\n\nProcessed invoices:\n${processed.map(x => `• ${x}`).join('\n')}`;
    }

    if (notFound.length) {
      msg += `\n\nInvoice folder/sheet not found for:\n${notFound.map(x => `• ${x}`).join('\n')}`;
    }

    setStatus_('Finalizing summary...', 98, runId);
    ui.alert(msg);
  } catch (error) {
    setStatus_(`Error: ${error.message}`, 100, runId);
    ui.alert(`Error: ${error.message}`);
    throw error;
  }
}

/***** CUSTOMER FOLDER LOOKUP *****/
function findCustomerPoFolder_(purchasesRoot, customerNumber) {
  const normalized = customerNumber.replace(/^#/, '').trim();

  const exactNames = [`#${normalized}-POs`, `${normalized}-POs`];
  let folders = purchasesRoot.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (exactNames.includes(f.getName())) return f;
  }

  folders = purchasesRoot.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    const n = f.getName().toLowerCase();
    if (n.includes(normalized.toLowerCase()) && n.includes('po')) return f;
  }

  return null;
}

/***** INVOICE LOOKUP (SEARCH FROM PURCHASES ROOT) *****/
function findInvoiceSheetInPurchases_(purchasesRoot, invoiceNumber, customerNumber) {
  const matches = [];
  collectFoldersByExactNameRecursive_(purchasesRoot, invoiceNumber, matches);

  if (!matches.length) return null;

  const preferred = pickBestInvoiceFolderMatch_(matches, customerNumber) || matches[0];
  const sheetFile = findInvoiceSheetInFolder_(preferred, invoiceNumber);
  if (!sheetFile) return null;

  return { folder: preferred, file: sheetFile };
}

function collectFoldersByExactNameRecursive_(folder, exactName, out) {
  if (folder.getName() === exactName) out.push(folder);

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    collectFoldersByExactNameRecursive_(subfolders.next(), exactName, out);
  }
}

function pickBestInvoiceFolderMatch_(folders, customerNumber) {
  if (!customerNumber) return null;
  const normCustomer = normalizeName_(customerNumber).replace(/^#/, '');

  for (let i = 0; i < folders.length; i++) {
    if (folderPathContainsCustomer_(folders[i], normCustomer)) return folders[i];
  }
  return null;
}

function folderPathContainsCustomer_(folder, normCustomer) {
  let current = folder;
  for (let i = 0; i < 30; i++) {
    const n = normalizeName_(current.getName()).replace(/^#/, '');
    if (n.includes(normCustomer)) return true;

    const parents = current.getParents();
    if (!parents.hasNext()) break;
    current = parents.next();
  }
  return false;
}

function findInvoiceSheetInFolder_(invoiceFolder, invoiceNumber) {
  const exact = invoiceFolder.getFilesByName(invoiceNumber);
  while (exact.hasNext()) {
    const f = exact.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) return f;
  }

  const anySheets = invoiceFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  return anySheets.hasNext() ? anySheets.next() : null;
}

function normalizeName_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^a-z0-9#]/g, '');
}

/***** GENERATED REPORT FILES *****/
function copyAllTemplates_(templates, targetFolder, customerNumber) {
  const result = {};
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');

  Object.keys(templates).forEach((reportKey) => {
    const url = templates[reportKey];
    const fileId = extractGoogleFileId_(url);
    if (!fileId) return;

    const sourceFile = DriveApp.getFileById(fileId);
    const newName = `${reportKey} - ${customerNumber} - ${stamp}`;
    const copy = sourceFile.makeCopy(newName, targetFolder);
    result[reportKey] = copy;
  });

  return result;
}

function extractGoogleFileId_(url) {
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

function getOrCreateSubfolder_(parentFolder, subfolderName) {
  const existing = parentFolder.getFoldersByName(subfolderName);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(subfolderName);
}

function createInvoiceTestCopy_(invoiceFile, targetFolder, customerNumber) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const name = `[TEST] ${invoiceFile.getName()} - ${customerNumber} - ${stamp}`;
  return invoiceFile.makeCopy(name, targetFolder);
}

/***** INVOICE TAB REFRESH *****/
function upsertPoSheetsInInvoice_(invoiceSpreadsheetId, invoiceNo, hasSb20, runId) {
  const ui = SpreadsheetApp.getUi();

  setStatus_(`Checking for required "PO" tab in ${invoiceNo}...`, null, runId);
  const canContinue = ensurePoSheetExistsOrCancel_(invoiceSpreadsheetId, invoiceNo);
  if (!canContinue) return false;

  const targetSs = SpreadsheetApp.openById(invoiceSpreadsheetId);
  const targetNames = hasSb20 ? ['EWASTE-PO', 'SB20-PO'] : ['EWASTE-PO'];

  const existing = targetNames.filter(name => targetSs.getSheetByName(name));

  if (existing.length) {
    const confirm = ui.alert(
      'Update PO Tabs',
      `Invoice ${invoiceNo} already has: ${existing.join(', ')}.\n\nDelete and import updated tab(s)?`,
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return false;
  }

  setStatus_(`Importing ${targetNames.join(' + ')} into ${invoiceNo}...`, null, runId);
  targetNames.forEach(name => {
    if (name === 'SB20-PO' && !hasSb20) return;
    const oldSheet = targetSs.getSheetByName(name);
    if (oldSheet) targetSs.deleteSheet(oldSheet);
    importUpdatedPoSheet_(targetSs, name);
  });

  SpreadsheetApp.flush();
  Utilities.sleep(1200);
  setStatus_(`Freezing formulas to values for ${invoiceNo}...`, null, runId);
  hardPasteValuesInPoSheets_(targetSs, targetNames);

  return true;
}

function ensurePoSheetExistsOrCancel_(invoiceSpreadsheetId, invoiceNo) {
  const ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.openById(invoiceSpreadsheetId);

  while (!ss.getSheetByName('PO')) {
    const decision = ui.alert(
      'Missing Required Sheet: PO',
      `Invoice ${invoiceNo} does not have a sheet named "PO".\n\n` +
      `Please add/rename a tab to exactly "PO", then click YES to re-check.\n` +
      `Click NO to skip this invoice.`,
      ui.ButtonSet.YES_NO
    );
    if (decision !== ui.Button.YES) return false;

    Utilities.sleep(500);
    ss = SpreadsheetApp.openById(invoiceSpreadsheetId);
  }

  return true;
}

function importUpdatedPoSheet_(targetSs, tabName) {
  const src = PO_SHEET_SOURCES[tabName];
  if (!src || !src.spreadsheetId) throw new Error(`Missing source config for ${tabName}`);

  const srcSs = SpreadsheetApp.openById(src.spreadsheetId);
  const srcSheet = srcSs.getSheetByName(src.sheetName);
  if (!srcSheet) throw new Error(`Source tab "${src.sheetName}" not found for ${tabName}`);

  const copied = srcSheet.copyTo(targetSs);
  copied.setName(tabName);
  normalizePoSheetNames_(targetSs);
}

function normalizePoSheetNames_(ss) {
  const renameMap = {
    'Copy of (NETCOST)SB20-PO': 'SB20-PO',
    'Copy of (DTSC)EWASTE-PO': 'EWASTE-PO',
    '(NETCOST)SB20-PO': 'SB20-PO',
    '(DTSC)EWASTE-PO': 'EWASTE-PO',
  };

  Object.keys(renameMap).forEach(oldName => {
    const newName = renameMap[oldName];
    const sh = ss.getSheetByName(oldName);
    if (!sh) return;

    const existingTarget = ss.getSheetByName(newName);
    if (existingTarget) {
      ss.deleteSheet(sh);
      return;
    }
    sh.setName(newName);
  });
}

function hardPasteValuesInPoSheets_(ss, sheetNames) {
  sheetNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastRow = Math.max(sh.getLastRow(), 1);
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const rng = sh.getRange(1, 1, lastRow, lastCol);

    // Ctrl+A + Ctrl+C + Ctrl+Shift+V equivalent
    rng.copyTo(rng, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  });
}

/***** INVOICE -> GENERATED REPORT IMPORTS *****/
function importInvoiceTabsToGeneratedReports_(invoiceSpreadsheetId, invoiceNo, reportFilesByKey, hasSb20) {
  const importedInto = [];
  const invoiceSs = SpreadsheetApp.openById(invoiceSpreadsheetId);

  const ewasteSheet = invoiceSs.getSheetByName('EWASTE-PO');
  if (ewasteSheet) {
    ['DTSC-ALL', 'DTSC'].forEach(reportKey => {
      if (!reportFilesByKey[reportKey]) return;
      const reportSs = SpreadsheetApp.openById(reportFilesByKey[reportKey].getId());
      const tabName = buildTargetTabName_('EWASTE-PO', invoiceNo);
      copySheetIntoReport_(ewasteSheet, reportSs, tabName);
      importedInto.push(`${reportKey}:${tabName}`);
    });
  }

  if (hasSb20) {
    const sb20Sheet = invoiceSs.getSheetByName('SB20-PO');
    if (sb20Sheet && reportFilesByKey['NETCOST']) {
      const netcostSs = SpreadsheetApp.openById(reportFilesByKey['NETCOST'].getId());
      const tabName = buildTargetTabName_('SB20-PO', invoiceNo);
      copySheetIntoReport_(sb20Sheet, netcostSs, tabName);
      importedInto.push(`NETCOST:${tabName}`);
    }
  }

  return importedInto;
}

function copySheetIntoReport_(sourceSheet, targetSpreadsheet, targetTabName) {
  const existing = targetSpreadsheet.getSheetByName(targetTabName);
  if (existing) targetSpreadsheet.deleteSheet(existing);

  const copied = sourceSheet.copyTo(targetSpreadsheet);
  copied.setName(targetTabName);
}

function buildTargetTabName_(baseName, invoiceNo) {
  const safeSuffix = sanitizeTabSuffix_(invoiceNo);
  const raw = `${baseName}-${safeSuffix}`;
  return raw.length <= 100 ? raw : raw.substring(0, 100);
}

function sanitizeTabSuffix_(text) {
  return String(text || '')
    .replace(/[\\\/\?\*\[\]:]/g, '-')
    .trim();
}

/***** UI HELPERS *****/
function showInvoiceLinksDialog_(foundList, testMode) {
  const rows = foundList.map(item => {
    const safeInvoice = escapeHtml_(item.invoiceNo);
    const safeName = escapeHtml_(item.file.getName());
    const url = item.file.getUrl();
    return `<li><b>${safeInvoice}</b>: <a href="${url}" target="_blank">${safeName}</a></li>`;
  }).join('');

  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;padding:12px;">
      <h3>Matched Invoice Sheets (${testMode ? 'TEST' : 'LIVE'})</h3>
      <ul>${rows}</ul>
      <p>These invoice files will be processed.</p>
    </div>`
  ).setWidth(540).setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(html, 'Matched Invoice Sheets');
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}