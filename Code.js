/***** CONFIG *****/
const PURCHASES_FOLDER_ID = '1lNJdDrFNR5bvdcCAvqxqM560H2GOvapR';

// Default mode when using "Start Report Wizard"
const TEST_MODE_DEFAULT = true;
const TEST_FOLDER_NAME = 'Generated Reports - TEST';
const LIVE_FOLDER_NAME = 'Generated Reports';
const ENABLE_PURCHASES_WIDE_FALLBACK = false;

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
    .createMenu('DTSC Report Wizard')
    .addItem('Open Progress Window', 'openWizardSidebar')
    .addToUi();
}

function openWizardSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Report Wizard Progress')
    .setWidth(560)
    .setHeight(760);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Report Wizard Progress');
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

function beginWizardRunWithParams(params) {
  const safe = params || {};
  const testMode = !!safe.testMode;
  const runId = beginWizardRun(testMode);
  return runId;
}

function runWizardWithParams(runId, params) {
  const safe = params || {};
  try {
    setStatus_('Starting window-run workflow...', 2, runId);
    startReportWizardWithInputs_(safe, runId);
    setStatus_('Completed successfully.', 100, runId);
    finishRun_(runId);
  } catch (error) {
    failRun_(runId, error);
    throw error;
  }
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
    const processedInvoiceNos = [];

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

        processedInvoiceNos.push(item.invoiceNo);
        processed.push(`${item.invoiceNo} -> ${imported.join(', ') || 'No import tabs found'}`);
      });
    }

    const renameResult = renameReportsByInvoicePeriod_(
      reportFilesByKey,
      customerNumber,
      processedInvoiceNos
    );
    if (renameResult.warning) {
      setStatus_(`Report naming warning: ${renameResult.warning}`, null, runId);
    } else if (renameResult.renamed) {
      setStatus_(`Renamed reports to ${renameResult.periodLabel}.`, null, runId);
    }

    if (hasSb20 && reportFilesByKey['NETCOST'] && processedInvoiceNos.length) {
      setStatus_('Updating NETCOST monthly totals...', 94, runId);
      const invoiceMonth = resolveSingleInvoiceMonth_(processedInvoiceNos);
      writeNetcostMonthlyTotals_(reportFilesByKey['NETCOST'], invoiceMonth);
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

function startReportWizardWithInputs_(params, runId) {
  const customerNumber = String(params.customerNumber || '').trim();
  const hasSb20 = !!params.hasSb20;
  const testMode = !!params.testMode;
  const invoiceInputs = String(params.invoiceNumbers || '')
    .split(/[\n,]+/)
    .map(s => s.trim())
    .filter(Boolean);

  if (!customerNumber) {
    throw new Error('Customer Number is required.');
  }
  if (!invoiceInputs.length) {
    throw new Error('At least one Invoice Number is required.');
  }

  setStatus_('Connecting to Purchases folder...', 5, runId);
  const purchasesRoot = DriveApp.getFolderById(PURCHASES_FOLDER_ID);

  setStatus_('Finding customer folder...', 12, runId);
  const customerFolder = findCustomerPoFolder_(purchasesRoot, customerNumber);
  if (!customerFolder) {
    throw new Error(`Customer folder not found for ${customerNumber}.`);
  }

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

  setStatus_(`Searching ${invoiceInputs.length} invoice folder(s)...`, 45, runId);
  const found = [];
  const notFound = [];
  const processed = [];
  const processedInvoiceNos = [];

  invoiceInputs.forEach((invoiceNo) => {
    const match = findInvoiceSheetForCustomer_(
      customerFolder,
      purchasesRoot,
      invoiceNo,
      customerNumber,
      runId
    );
    if (match) found.push({ invoiceNo, folder: match.folder, file: match.file });
    else notFound.push(invoiceNo);
  });

  if (found.length) {
    setStatus_(`Matched ${found.length} invoice(s).`, 52, runId);

    found.forEach(item => {
      setStatus_(`Processing invoice ${item.invoiceNo}...`, null, runId);
      const invoiceFileToProcess = testMode
        ? createInvoiceTestCopy_(item.file, generatedReportsFolder, customerNumber)
        : item.file;

      setStatus_(
        testMode
          ? `TEST mode: updating copy "${invoiceFileToProcess.getName()}" (original is unchanged).`
          : `LIVE mode: updating original invoice file "${invoiceFileToProcess.getName()}".`,
        null,
        runId
      );

      const refreshed = upsertPoSheetsInInvoice_(
        invoiceFileToProcess.getId(),
        item.invoiceNo,
        hasSb20,
        runId,
        { interactive: false }
      );

      if (!refreshed) {
        processed.push(`Skipped (missing PO): ${item.invoiceNo}`);
        return;
      }

      const imported = importInvoiceTabsToGeneratedReports_(
        invoiceFileToProcess.getId(),
        item.invoiceNo,
        reportFilesByKey,
        hasSb20
      );

      processedInvoiceNos.push(item.invoiceNo);
      processed.push(`${item.invoiceNo} -> ${imported.join(', ') || 'No import tabs found'}`);
    });
  }

  const renameResult = renameReportsByInvoicePeriod_(
    reportFilesByKey,
    customerNumber,
    processedInvoiceNos
  );
  if (renameResult.warning) {
    setStatus_(`Report naming warning: ${renameResult.warning}`, null, runId);
  } else if (renameResult.renamed) {
    setStatus_(`Renamed reports to ${renameResult.periodLabel}.`, null, runId);
  }

  if (hasSb20 && reportFilesByKey['NETCOST'] && processedInvoiceNos.length) {
    setStatus_('Updating NETCOST monthly totals...', 94, runId);
    const invoiceMonth = resolveSingleInvoiceMonth_(processedInvoiceNos);
    writeNetcostMonthlyTotals_(reportFilesByKey['NETCOST'], invoiceMonth);
  }

  const copiedReportKeys = Object.keys(reportFilesByKey);
  let summary =
    `${testMode ? 'TEST' : 'LIVE'} complete.\n\n` +
    `Generated folder:\n${generatedReportsFolder.getUrl()}\n\n` +
    `Copied ${copiedReportKeys.length} report file(s):\n` +
    copiedReportKeys.map(k => `• ${reportFilesByKey[k].getName()}`).join('\n');

  if (processed.length) {
    summary += `\n\nProcessed invoices:\n${processed.map(x => `• ${x}`).join('\n')}`;
  }
  if (notFound.length) {
    summary += `\n\nInvoice folder/sheet not found for:\n${notFound.map(x => `• ${x}`).join('\n')}`;
  }

  setStatus_('Finalizing summary...', 98, runId);
  setStatus_(summary, 99, runId);
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

function findInvoiceSheetForCustomer_(customerFolder, purchasesRoot, invoiceNumber, customerNumber, runId) {
  setStatus_(`Looking for invoice ${invoiceNumber} in customer folder first...`, null, runId);
  const inCustomer = findInvoiceSheetInCustomerFolderRecursive_(customerFolder, invoiceNumber);
  if (inCustomer) return inCustomer;

  if (!ENABLE_PURCHASES_WIDE_FALLBACK) {
    return null;
  }

  setStatus_(`Not found in customer folder; searching all Purchases (slower)...`, null, runId);
  return findInvoiceSheetInPurchases_(purchasesRoot, invoiceNumber, customerNumber);
}

function findInvoiceSheetInCustomerFolderRecursive_(customerFolder, invoiceNumber) {
  const matches = [];
  collectFoldersByExactNameRecursive_(customerFolder, invoiceNumber, matches);
  if (!matches.length) return null;

  const invoiceFolder = matches[0];
  const sheetFile = findInvoiceSheetInFolder_(invoiceFolder, invoiceNumber);
  if (!sheetFile) return null;

  return { folder: invoiceFolder, file: sheetFile };
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
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd-yyyy');
  const customerTag = buildCustomerTag_(customerNumber);

  Object.keys(templates).forEach((reportKey) => {
    const url = templates[reportKey];
    const fileId = extractGoogleFileId_(url);
    if (!fileId) return;

    const sourceFile = DriveApp.getFileById(fileId);
    const newName = customerTag
      ? `${customerTag}-${reportKey}-${stamp}`
      : `${reportKey}-${stamp}`;
    const copy = sourceFile.makeCopy(newName, targetFolder);
    result[reportKey] = copy;
  });

  return result;
}

function buildCustomerTag_(customerNumber) {
  const rawCustomer = String(customerNumber || '').trim();
  if (!rawCustomer) return '';
  return rawCustomer.startsWith('#') ? rawCustomer : `#${rawCustomer}`;
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
function upsertPoSheetsInInvoice_(invoiceSpreadsheetId, invoiceNo, hasSb20, runId, options) {
  const ui = SpreadsheetApp.getUi();
  const interactive = !options || options.interactive !== false;

  setStatus_(`Checking for required "PO" tab in ${invoiceNo}...`, null, runId);
  const canContinue = ensurePoSheetExistsOrCancel_(invoiceSpreadsheetId, invoiceNo, interactive);
  if (!canContinue) return false;

  const targetSs = SpreadsheetApp.openById(invoiceSpreadsheetId);
  const targetNames = hasSb20 ? ['EWASTE-PO', 'SB20-PO'] : ['EWASTE-PO'];

  const existing = targetNames.filter(name => targetSs.getSheetByName(name));

  if (existing.length && interactive) {
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

function ensurePoSheetExistsOrCancel_(invoiceSpreadsheetId, invoiceNo, interactive) {
  const ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.openById(invoiceSpreadsheetId);

  if (interactive === false) {
    return !!ss.getSheetByName('PO');
  }

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
  const importedTabsByReport = {
    'DTSC-ALL': [],
    'DTSC': [],
    'NETCOST': [],
  };
  const invoiceSs = SpreadsheetApp.openById(invoiceSpreadsheetId);

  const poSheet = invoiceSs.getSheetByName('PO');
  if (poSheet && reportFilesByKey['DTSC-ALL']) {
    const dtscAllSs = SpreadsheetApp.openById(reportFilesByKey['DTSC-ALL'].getId());
    const tabName = buildTargetTabName_('PO', invoiceNo);
    copySheetIntoReport_(poSheet, dtscAllSs, tabName);
    importedTabsByReport['DTSC-ALL'].push(tabName);
    importedInto.push(`DTSC-ALL:${tabName}`);
  }

  const ewasteSheet = invoiceSs.getSheetByName('EWASTE-PO');
  if (ewasteSheet && reportFilesByKey['DTSC']) {
    const dtscSs = SpreadsheetApp.openById(reportFilesByKey['DTSC'].getId());
    const tabName = buildTargetTabName_('EWASTE-PO', invoiceNo);
    copySheetIntoReport_(ewasteSheet, dtscSs, tabName);
    importedTabsByReport['DTSC'].push(tabName);
    importedInto.push(`DTSC:${tabName}`);
  }

  if (hasSb20) {
    const sb20Sheet = invoiceSs.getSheetByName('SB20-PO');
    if (sb20Sheet && reportFilesByKey['NETCOST']) {
      const netcostSs = SpreadsheetApp.openById(reportFilesByKey['NETCOST'].getId());
      const tabName = buildTargetTabName_('SB20-PO', invoiceNo);
      copySheetIntoReport_(sb20Sheet, netcostSs, tabName);
      importedTabsByReport['NETCOST'].push(tabName);
      importedInto.push(`NETCOST:${tabName}`);
    }
  }

  Object.keys(importedTabsByReport).forEach((reportKey) => {
    if (!reportFilesByKey[reportKey]) return;
    const reportSs = SpreadsheetApp.openById(reportFilesByKey[reportKey].getId());
    updateImportedDataSheet_(reportSs, importedTabsByReport[reportKey]);
  });

  return importedInto;
}

function updateImportedDataSheet_(reportSpreadsheet, importedTabNames) {
  let importedDataSheet = reportSpreadsheet.getSheetByName('IMPORTED-DATA');
  if (!importedDataSheet) {
    importedDataSheet = reportSpreadsheet.insertSheet('IMPORTED-DATA');
  }

  const existing = importedDataSheet
    .getRange('A2:A31')
    .getValues()
    .map(row => String(row[0] || '').trim())
    .filter(Boolean);

  const incoming = (importedTabNames || [])
    .map(name => String(name || '').trim())
    .filter(Boolean);

  const names = existing.concat(incoming).slice(0, 30);

  importedDataSheet.getRange('A2:A31').clearContent();
  if (!names.length) return;

  const values = names.map(name => [name]);
  importedDataSheet.getRange(2, 1, values.length, 1).setValues(values);
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

function parseInvoiceMonthYear_(invoiceNo) {
  const raw = String(invoiceNo || '').trim();
  if (!raw) return null;

  const parts = raw.split('-').map(p => p.trim()).filter(Boolean);
  if (parts.length < 3) return null;

  const monthToken = parts[parts.length - 3];
  const yearToken = parts[parts.length - 1];

  if (!/^\d{1,2}$/.test(monthToken)) return null;
  if (!/^\d{2,4}$/.test(yearToken)) return null;

  const month = Number(monthToken);
  if (month < 1 || month > 12) return null;

  const yearNum = Number(yearToken);
  if (Number.isNaN(yearNum)) return null;
  const year = yearToken.length === 2 ? 2000 + yearNum : yearNum;

  if (year < 2000 || year > 2099) return null;

  return { month, year };
}

function resolveSingleInvoicePeriodForNaming_(invoiceNumbers) {
  const periods = new Set();
  const parseErrors = [];

  (invoiceNumbers || []).forEach((invoiceNo) => {
    const parsed = parseInvoiceMonthYear_(invoiceNo);
    if (!parsed) {
      parseErrors.push(invoiceNo);
      return;
    }
    periods.add(`${parsed.year}-${String(parsed.month).padStart(2, '0')}`);
  });

  if (!invoiceNumbers || !invoiceNumbers.length) {
    return {
      warning: 'No successfully processed invoices, so report names were left unchanged.',
      period: null,
    };
  }

  if (parseErrors.length) {
    return {
      warning:
        `Could not parse invoice month/year for ${parseErrors.length} invoice(s), ` +
        'so report names were left unchanged.',
      period: null,
    };
  }

  if (periods.size !== 1) {
    return {
      warning: 'Processed invoices include multiple months, so report names were left unchanged.',
      period: null,
    };
  }

  const periodKey = Array.from(periods)[0];
  const [yearText, monthText] = periodKey.split('-');
  return {
    warning: '',
    period: {
      year: Number(yearText),
      month: Number(monthText),
    },
  };
}

function renameReportsByInvoicePeriod_(reportFilesByKey, customerNumber, processedInvoiceNos) {
  const resolved = resolveSingleInvoicePeriodForNaming_(processedInvoiceNos);
  if (!resolved.period) {
    return {
      renamed: false,
      warning: resolved.warning,
      periodLabel: '',
    };
  }

  const periodLabel = `${monthNumberToName_(resolved.period.month)}-${resolved.period.year}`;
  const customerTag = buildCustomerTag_(customerNumber);

  Object.keys(reportFilesByKey || {}).forEach((reportKey) => {
    const reportFile = reportFilesByKey[reportKey];
    if (!reportFile) return;

    const newName = customerTag
      ? `${customerTag}-${reportKey}-${periodLabel}`
      : `${reportKey}-${periodLabel}`;
    reportFile.setName(newName);
  });

  return {
    renamed: true,
    warning: '',
    periodLabel,
  };
}

function resolveSingleInvoiceMonth_(invoiceNumbers) {
  const months = new Set();

  (invoiceNumbers || []).forEach((invoiceNo) => {
    const parsed = parseInvoiceMonthYear_(invoiceNo);
    const month = parsed ? parsed.month : null;
    if (!month) {
      throw new Error(
        `Could not parse month from invoice "${invoiceNo}". Expected format like #199-1-01-08-26.`
      );
    }
    months.add(month);
  });

  if (!months.size) {
    throw new Error('No invoice month found for NETCOST monthly update.');
  }

  if (months.size > 1) {
    const monthNames = Array.from(months)
      .sort((a, b) => a - b)
      .map(monthNumberToName_)
      .join(', ');
    throw new Error(
      `Multiple invoice months found (${monthNames}). Run one month at a time for monthly reports.`
    );
  }

  return Array.from(months)[0];
}

function monthNumberToName_(monthNumber) {
  const names = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ];
  return names[monthNumber - 1] || '';
}

function monthLabelToNumber_(value) {
  const raw = String(value || '').trim();
  if (!raw) return null;

  if (/^\d{1,2}$/.test(raw)) {
    const numeric = Number(raw);
    return numeric >= 1 && numeric <= 12 ? numeric : null;
  }

  const normalized = raw.toLowerCase();
  const map = {
    jan: 1,
    january: 1,
    feb: 2,
    february: 2,
    mar: 3,
    march: 3,
    apr: 4,
    april: 4,
    may: 5,
    jun: 6,
    june: 6,
    jul: 7,
    july: 7,
    aug: 8,
    august: 8,
    sep: 9,
    sept: 9,
    september: 9,
    oct: 10,
    october: 10,
    nov: 11,
    november: 11,
    dec: 12,
    december: 12,
  };

  if (map[normalized]) return map[normalized];

  const cleaned = normalized.replace(/[^a-z]/g, '');
  if (map[cleaned]) return map[cleaned];

  return null;
}

function writeNetcostMonthlyTotals_(netcostFile, invoiceMonth) {
  const reportSs = SpreadsheetApp.openById(netcostFile.getId());
  const netcostSheet = reportSs.getSheetByName('NETCOST');
  if (!netcostSheet) {
    throw new Error('NETCOST report is missing required "NETCOST" sheet.');
  }

  const totals = netcostSheet.getRange('B20:R20').getValues();
  const monthCells = netcostSheet.getRange('A25:A36').getDisplayValues();

  let targetRow = null;
  for (let i = 0; i < monthCells.length; i++) {
    const labelMonth = monthLabelToNumber_(monthCells[i][0]);
    if (labelMonth === invoiceMonth) {
      targetRow = 25 + i;
      break;
    }
  }

  if (!targetRow) {
    throw new Error(
      `Could not find month row for ${monthNumberToName_(invoiceMonth)} in NETCOST!A25:A36.`
    );
  }

  netcostSheet.getRange(targetRow, 2, 1, 17).setValues(totals);
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