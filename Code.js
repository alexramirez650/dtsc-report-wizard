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
    .addItem('Start Report Wizard', 'openWizardSidebar')
    .addToUi();
}

function openWizardSidebar() {
  openSidebarWithMode_('wizard');
}

function openCompileSidebar() {
  openSidebarWithMode_('compile');
}

function openSidebarWithMode_(mode) {
  const tpl = HtmlService.createTemplateFromFile('sidebar');
  tpl.initialMode = mode || 'wizard';
  const title = mode === 'compile' ? 'Compile Month Totals' : 'Report Wizard Progress';
  const html = tpl.evaluate()
    .setTitle(title)
    .setWidth(620)
    .setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, title);
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

function beginCompileRunWithParams(params) {
  const safe = params || {};
  const runId = `run_${Date.now()}_${Math.floor(Math.random() * 1000000)}`;
  const state = {
    runId,
    status: 'running',
    progress: 0,
    logs: [`${timestamp_()} Compile run created (${safe.testMode ? 'TEST' : 'LIVE'})`],
    error: '',
    mode: safe.testMode ? 'COMPILE-TEST' : 'COMPILE-LIVE',
  };
  saveRunState_(runId, state);
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

function runCompileMonthTotalsWithParams(runId, params) {
  const safe = params || {};
  try {
    setStatus_('Starting compile workflow...', 2, runId);
    compileMonthTotalsWithInputs_(safe, runId);
    setStatus_('Completed successfully.', 100, runId);
    finishRun_(runId);
  } catch (error) {
    failRun_(runId, error);
    throw error;
  }
}

function listCompileTargetReports(params) {
  const safe = params || {};
  const customerNumber = String(safe.customerNumber || '').trim();
  const reportType = String(safe.reportType || '').trim().toUpperCase();
  const testMode = !!safe.testMode;
  const year = Number(safe.year || new Date().getFullYear());

  if (!customerNumber || !reportType || Number.isNaN(year)) return [];
  if (!['DTSC-ALL', 'DTSC', 'NETCOST'].includes(reportType)) return [];

  const purchasesRoot = DriveApp.getFolderById(PURCHASES_FOLDER_ID);
  const customerFolder = findCustomerPoFolder_(purchasesRoot, customerNumber);
  if (!customerFolder) return [];

  const customerTag = buildCustomerTag_(customerNumber);
  const candidates = collectReportFilesForCompile_(customerFolder, customerTag, testMode, reportType, year);
  return candidates.map((item) => ({
    id: item.file.getId(),
    name: item.file.getName(),
    month: item.month,
    year: item.year,
  }));
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

    const folderRename = renameGeneratedReportsFolderByInvoicePeriod_(
      generatedReportsFolder,
      testMode,
      customerNumber,
      processedInvoiceNos
    );
    if (folderRename.warning) {
      setStatus_(`Folder naming warning: ${folderRename.warning}`, null, runId);
    } else if (folderRename.renamed) {
      setStatus_(`Renamed folder to ${folderRename.folderName}.`, null, runId);
    }

    const backfillResult = backfillPreviousMonthsTotals_(
      customerFolder,
      generatedReportsFolder,
      reportFilesByKey,
      customerNumber,
      testMode,
      processedInvoiceNos
    );
    if (backfillResult.warning) {
      setStatus_(`Backfill warning: ${backfillResult.warning}`, null, runId);
    } else if (backfillResult.updatedReports.length) {
      setStatus_(`Backfilled previous months for: ${backfillResult.updatedReports.join(', ')}`, 91, runId);
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
  const invoiceLinkInputs = String(params.invoiceLinks || '')
    .split(/[\n,]+/)
    .map(s => s.trim())
    .filter(Boolean);

  if (!customerNumber) {
    throw new Error('Customer Number is required.');
  }
  if (!invoiceInputs.length && !invoiceLinkInputs.length) {
    throw new Error('At least one Invoice Number or Invoice Link is required.');
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

  setStatus_(`Resolving ${invoiceInputs.length + invoiceLinkInputs.length} invoice input(s)...`, 45, runId);
  const found = [];
  const notFound = [];
  const processed = [];
  const processedInvoiceNos = [];

  invoiceLinkInputs.forEach((linkText) => {
    const fileId = extractGoogleFileId_(linkText) || parsePossibleDriveFileId_(linkText);
    if (!fileId) {
      notFound.push(linkText);
      return;
    }

    try {
      const file = DriveApp.getFileById(fileId);
      if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
        notFound.push(linkText);
        return;
      }

      const invoiceNo = deriveInvoiceNumberForFile_(file, linkText);
      found.push({ invoiceNo, folder: null, file });
    } catch (_) {
      notFound.push(linkText);
    }
  });

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

  const folderRename = renameGeneratedReportsFolderByInvoicePeriod_(
    generatedReportsFolder,
    testMode,
    customerNumber,
    processedInvoiceNos
  );
  if (folderRename.warning) {
    setStatus_(`Folder naming warning: ${folderRename.warning}`, null, runId);
  } else if (folderRename.renamed) {
    setStatus_(`Renamed folder to ${folderRename.folderName}.`, null, runId);
  }

  const backfillResult = backfillPreviousMonthsTotals_(
    customerFolder,
    generatedReportsFolder,
    reportFilesByKey,
    customerNumber,
    testMode,
    processedInvoiceNos
  );
  if (backfillResult.warning) {
    setStatus_(`Backfill warning: ${backfillResult.warning}`, null, runId);
  } else if (backfillResult.updatedReports.length) {
    setStatus_(`Backfilled previous months for: ${backfillResult.updatedReports.join(', ')}`, 91, runId);
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

function parsePossibleDriveFileId_(text) {
  const raw = String(text || '').trim();
  if (!raw) return '';
  return /^[a-zA-Z0-9_-]{20,}$/.test(raw) ? raw : '';
}

function deriveInvoiceNumberForFile_(file, fallbackText) {
  const name = String(file && file.getName ? file.getName() : '').trim();
  const fromName = extractInvoiceNumberFromText_(name);
  if (fromName) return fromName;

  const fromFallback = extractInvoiceNumberFromText_(fallbackText);
  if (fromFallback) return fromFallback;

  return name || String(fallbackText || '').trim() || file.getId();
}

function extractInvoiceNumberFromText_(text) {
  const raw = String(text || '').trim();
  if (!raw) return '';

  const m = raw.match(/#?\d+-\d+-\d{1,2}-\d{1,2}-(?:\d{2}|20\d{2})/);
  return m ? m[0] : '';
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
  collectInvoiceFoldersByMatchRecursive_(purchasesRoot, invoiceNumber, matches);

  if (!matches.length) return null;

  const preferred = pickBestInvoiceFolderMatch_(matches, customerNumber, invoiceNumber) || matches[0];
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
  collectInvoiceFoldersByMatchRecursive_(customerFolder, invoiceNumber, matches);
  if (!matches.length) return null;

  const invoiceFolder = pickBestInvoiceFolderMatch_(matches, '', invoiceNumber) || matches[0];
  const sheetFile = findInvoiceSheetInFolder_(invoiceFolder, invoiceNumber);
  if (!sheetFile) return null;

  return { folder: invoiceFolder, file: sheetFile };
}

function collectInvoiceFoldersByMatchRecursive_(folder, invoiceNumber, out) {
  const targetKeys = buildInvoiceMatchKeys_(invoiceNumber);
  collectInvoiceFoldersByMatchRecursiveInner_(folder, invoiceNumber, targetKeys, out);
}

function collectInvoiceFoldersByMatchRecursiveInner_(folder, invoiceNumber, targetKeys, out) {
  if (isInvoiceFolderNameMatch_(folder.getName(), invoiceNumber, targetKeys)) {
    out.push(folder);
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    collectInvoiceFoldersByMatchRecursiveInner_(
      subfolders.next(),
      invoiceNumber,
      targetKeys,
      out
    );
  }
}

function buildInvoiceMatchKeys_(text) {
  const raw = String(text || '').trim();
  if (!raw) return [];

  const variants = new Set([raw]);
  const tokens = raw.split('-').map(t => t.trim()).filter(Boolean);
  const tokenShapes = [];
  if (tokens.length) {
    tokenShapes.push(tokens);
    tokenShapes.push(tokens.map(normalizeInvoiceToken_));
  }

  const shapeSeen = new Set();
  tokenShapes.forEach((shape) => {
    const sig = shape.join('|');
    if (shapeSeen[sig]) return;
    shapeSeen[sig] = true;

    addInvoiceYearVariants_(shape, variants);

    if (shape.length >= 3) {
      const dateOnly = shape.slice(shape.length - 3);
      addInvoiceYearVariants_(dateOnly, variants);
    }
  });

  const keys = new Set();
  variants.forEach((value) => {
    const norm = normalizeName_(value);
    if (!norm) return;
    keys.add(norm);
    keys.add(norm.replace(/^#/, ''));
  });

  return Array.from(keys);
}

function normalizeInvoiceToken_(token) {
  const value = String(token || '').trim();
  if (!value) return value;

  if (/^\d+$/.test(value)) {
    return String(Number(value));
  }

  if (/^#\d+$/.test(value)) {
    return `#${Number(value.slice(1))}`;
  }

  return value;
}

function addInvoiceYearVariants_(tokens, outSet) {
  const base = (tokens || []).join('-');
  if (!base) return;
  outSet.add(base);

  if (!tokens || !tokens.length) return;
  const year = String(tokens[tokens.length - 1] || '').trim();
  if (/^\d{2}$/.test(year)) {
    const alt = tokens.slice();
    alt[alt.length - 1] = `20${year}`;
    outSet.add(alt.join('-'));
  } else if (/^20\d{2}$/.test(year)) {
    const alt = tokens.slice();
    alt[alt.length - 1] = year.slice(2);
    outSet.add(alt.join('-'));
  }
}

function isInvoiceFolderNameMatch_(folderName, invoiceNumber, targetKeys) {
  if (String(folderName || '') === String(invoiceNumber || '')) return true;

  const folderNorm = normalizeName_(folderName);
  if (!folderNorm || !targetKeys || !targetKeys.length) return false;
  const folderNoHash = folderNorm.replace(/^#/, '');
  const folderKeys = [folderNorm, folderNoHash];

  for (let i = 0; i < folderKeys.length; i++) {
    const folderKey = folderKeys[i];
    for (let j = 0; j < targetKeys.length; j++) {
      const targetKey = targetKeys[j];
      if (!targetKey) continue;
      if (folderKey === targetKey) return true;
      if (folderKey.startsWith(targetKey)) return true;
      if (folderKey.includes(targetKey)) return true;
    }
  }

  return false;
}

function scoreInvoiceFolderNameMatch_(folderName, invoiceNumber) {
  const rawFolder = String(folderName || '');
  const rawInvoice = String(invoiceNumber || '');
  if (rawFolder === rawInvoice) return 100;

  const targetKeys = buildInvoiceMatchKeys_(invoiceNumber);
  const folderNorm = normalizeName_(folderName);
  const folderNoHash = folderNorm ? folderNorm.replace(/^#/, '') : '';
  const folderKeys = [folderNorm, folderNoHash].filter(Boolean);

  if (!folderKeys.length || !targetKeys.length) return 0;

  let score = 0;
  for (let i = 0; i < folderKeys.length; i++) {
    const folderKey = folderKeys[i];
    for (let j = 0; j < targetKeys.length; j++) {
      const targetKey = targetKeys[j];
      if (!targetKey) continue;

      if (folderKey === targetKey) return 90;
      if (folderKey.startsWith(targetKey)) {
        score = Math.max(score, 70);
        continue;
      }
      if (folderKey.includes(targetKey)) {
        score = Math.max(score, 50);
      }
    }
  }

  if (score > 0) return score;
  return 0;
}

function pickBestInvoiceFolderMatch_(folders, customerNumber, invoiceNumber) {
  const normCustomer = normalizeName_(customerNumber).replace(/^#/, '');
  let best = null;
  let bestScore = -1;

  for (let i = 0; i < folders.length; i++) {
    const folder = folders[i];
    const nameScore = scoreInvoiceFolderNameMatch_(folder.getName(), invoiceNumber);
    const customerBonus = normCustomer && folderPathContainsCustomer_(folder, normCustomer) ? 10 : 0;
    const score = nameScore + customerBonus;

    if (score > bestScore) {
      bestScore = score;
      best = folder;
    }
  }

  return best;
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
  if (!parts.length) return null;

  const parseYearToken = (token) => {
    if (!/^\d{2,4}$/.test(token)) return null;
    const yearNum = Number(token);
    if (Number.isNaN(yearNum)) return null;
    const year = token.length === 2 ? 2000 + yearNum : yearNum;
    return year >= 2000 && year <= 2099 ? year : null;
  };

  const parseMonthToken = (token) => {
    if (!/^\d{1,2}$/.test(token)) return null;
    const month = Number(token);
    return month >= 1 && month <= 12 ? month : null;
  };

  const parseDayToken = (token) => {
    if (!/^\d{1,2}$/.test(token)) return null;
    const day = Number(token);
    return day >= 1 && day <= 31 ? day : null;
  };

  const yearToken = parts[parts.length - 1];
  const year = parseYearToken(yearToken);
  if (year === null) return null;

  const prevToken = parts.length >= 2 ? parts[parts.length - 2] : '';
  const prevPrevToken = parts.length >= 3 ? parts[parts.length - 3] : '';

  const dayBeforeYear = parseDayToken(prevToken);
  if (dayBeforeYear !== null) {
    const monthBeforeDay = parseMonthToken(prevPrevToken);
    if (monthBeforeDay !== null) {
      return { month: monthBeforeDay, year };
    }
  }

  const monthBeforeYear = parseMonthToken(prevToken);
  if (monthBeforeYear !== null) {
    return { month: monthBeforeYear, year };
  }

  return null;
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

function renameGeneratedReportsFolderByInvoicePeriod_(folder, testMode, customerNumber, processedInvoiceNos) {
  const resolved = resolveSingleInvoicePeriodForNaming_(processedInvoiceNos);
  if (!resolved.period) {
    return {
      renamed: false,
      warning: resolved.warning,
      folderName: '',
    };
  }

  const periodLabel = `${monthNumberToName_(resolved.period.month)}-${resolved.period.year}`;
  const baseName = testMode ? TEST_FOLDER_NAME : LIVE_FOLDER_NAME;
  const customerTag = buildCustomerTag_(customerNumber);
  const folderName = customerTag
    ? `${customerTag} - ${baseName} - ${periodLabel}`
    : `${baseName} - ${periodLabel}`;

  if (folder.getName() !== folderName) {
    folder.setName(folderName);
  }

  return {
    renamed: true,
    warning: '',
    folderName,
  };
}

function backfillPreviousMonthsTotals_(
  customerFolder,
  currentGeneratedFolder,
  reportFilesByKey,
  customerNumber,
  testMode,
  processedInvoiceNos
) {
  const resolved = resolveSingleInvoicePeriodForNaming_(processedInvoiceNos);
  if (!resolved.period) {
    return {
      warning: resolved.warning,
      updatedReports: [],
    };
  }

  const currentMonth = resolved.period.month;
  const currentYear = resolved.period.year;
  if (currentMonth <= 1) {
    return {
      warning: '',
      updatedReports: [],
    };
  }

  const customerTag = buildCustomerTag_(customerNumber);
  const updatedReports = [];

  Object.keys(reportFilesByKey || {}).forEach((reportKey) => {
    const currentFile = reportFilesByKey[reportKey];
    if (!currentFile) return;

    const previousFile = findMostRecentPreviousReportFile_(
      customerFolder,
      currentGeneratedFolder,
      reportKey,
      customerTag,
      testMode,
      currentYear,
      currentMonth,
      currentFile.getId()
    );
    if (!previousFile) return;

    if (copyPreviousMonthBlockByReportKey_(previousFile, currentFile, reportKey, currentMonth)) {
      updatedReports.push(reportKey);
    }
  });

  return {
    warning: '',
    updatedReports,
  };
}

function findMostRecentPreviousReportFile_(
  customerFolder,
  currentGeneratedFolder,
  reportKey,
  customerTag,
  testMode,
  currentYear,
  currentMonth,
  excludeFileId
) {
  const targetPeriod = currentYear * 100 + currentMonth;
  const candidateFolders = [];

  if (currentGeneratedFolder) {
    candidateFolders.push(currentGeneratedFolder);
  }

  const allGeneratedFolders = collectGeneratedReportFoldersRecursive_(
    customerFolder,
    customerTag,
    testMode,
    { maxFolders: 5000, maxDepth: 30 }
  );
  allGeneratedFolders.forEach((folder) => {
    if (currentGeneratedFolder && folder.getId() === currentGeneratedFolder.getId()) return;
    candidateFolders.push(folder);
  });

  let best = null;

  candidateFolders.forEach((folder) => {
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const file = files.next();
      if (file.getId() === excludeFileId) continue;

      const parsed = parseReportFilePeriodAndKey_(file.getName());
      if (!parsed) continue;
      if (parsed.reportKey !== reportKey) continue;
      if (parsed.year !== currentYear) continue;

      const periodValue = parsed.year * 100 + parsed.month;
      if (periodValue >= targetPeriod) continue;

      if (!best || periodValue > best.periodValue ||
          (periodValue === best.periodValue && file.getLastUpdated().getTime() > best.lastUpdated)) {
        best = {
          file,
          periodValue,
          lastUpdated: file.getLastUpdated().getTime(),
        };
      }
    }
  });

  return best ? best.file : null;
}

function isGeneratedReportsFolderCandidate_(folderName, customerTag, testMode) {
  const name = String(folderName || '').trim();
  const modeBase = testMode ? TEST_FOLDER_NAME : LIVE_FOLDER_NAME;
  const taggedBase = customerTag ? `${customerTag} - ${modeBase}` : modeBase;

  if (name === modeBase || name.startsWith(`${modeBase} - `)) return true;
  if (name === taggedBase || name.startsWith(`${taggedBase} - `)) return true;

  return false;
}

function collectGeneratedReportFoldersRecursive_(rootFolder, customerTag, testMode, options) {
  const opts = options || {};
  const maxFolders = Number(opts.maxFolders || 5000);
  const maxDepth = Number(opts.maxDepth || 30);

  const result = [];
  const seen = {};
  const stack = [{ folder: rootFolder, depth: 0 }];
  let scanned = 0;

  while (stack.length) {
    const node = stack.pop();
    if (!node || !node.folder) continue;

    const folder = node.folder;
    const depth = node.depth;
    const folderId = folder.getId();
    if (seen[folderId]) continue;
    seen[folderId] = true;

    if (depth > 0 && isGeneratedReportsFolderCandidate_(folder.getName(), customerTag, testMode)) {
      result.push(folder);
    }

    if (depth >= maxDepth || scanned >= maxFolders) continue;

    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      stack.push({ folder: subfolders.next(), depth: depth + 1 });
      scanned += 1;
      if (scanned >= maxFolders) break;
    }
  }

  return result;
}

function parseReportFilePeriodAndKey_(fileName) {
  const text = String(fileName || '').trim();
  const monthRegex =
    '(January|February|March|April|May|June|July|August|September|October|November|December)';
  const keyRegex = '(DTSC-ALL|DTSC|NETCOST)';
  const re = new RegExp(`${keyRegex}-${monthRegex}-(\\d{4})$`, 'i');
  const match = text.match(re);
  if (!match) return null;

  const reportKey = match[1].toUpperCase();
  const monthName = match[2];
  const year = Number(match[3]);
  const month = monthLabelToNumber_(monthName);

  if (!month || Number.isNaN(year)) return null;

  return {
    reportKey,
    month,
    year,
  };
}

function copyPreviousMonthBlockByReportKey_(sourceFile, targetFile, reportKey, currentMonth) {
  const map = {
    'NETCOST': { sheetName: 'NETCOST', startCol: 2, numCols: 17 },
    'DTSC-ALL': { sheetName: 'DTSC ALL ITEMS', startCol: 2, numCols: 13 },
    'DTSC': { sheetName: 'DTSC', startCol: 2, numCols: 11 },
  };

  const cfg = map[reportKey];
  if (!cfg) return false;

  const rowsToCopy = currentMonth - 1;
  if (rowsToCopy <= 0) return false;

  const sourceSs = SpreadsheetApp.openById(sourceFile.getId());
  const targetSs = SpreadsheetApp.openById(targetFile.getId());
  const sourceSheet = sourceSs.getSheetByName(cfg.sheetName);
  const targetSheet = targetSs.getSheetByName(cfg.sheetName);
  if (!sourceSheet || !targetSheet) return false;

  const sourceValues = sourceSheet.getRange(25, cfg.startCol, rowsToCopy, cfg.numCols).getValues();
  targetSheet.getRange(25, cfg.startCol, rowsToCopy, cfg.numCols).setValues(sourceValues);
  return true;
}

function compileMonthTotalsWithInputs_(params, runId) {
  const safe = params || {};
  const customerNumber = String(safe.customerNumber || '').trim();
  const reportType = String(safe.reportType || '').trim().toUpperCase();
  const targetReportId = String(safe.targetReportId || '').trim();
  const year = Number(safe.year || new Date().getFullYear());
  const testMode = !!safe.testMode;
  const months = normalizeCompileMonths_(safe.months);

  if (!customerNumber) throw new Error('Customer Number is required.');
  if (!['DTSC-ALL', 'DTSC', 'NETCOST'].includes(reportType)) {
    throw new Error('Report type must be DTSC-ALL, DTSC, or NETCOST.');
  }
  if (!targetReportId) throw new Error('Target report is required.');
  if (Number.isNaN(year) || year < 2000 || year > 2099) throw new Error('Year is invalid.');
  if (!months.length) throw new Error('Select at least one month to compile.');

  setStatus_('Connecting to Purchases folder...', 6, runId);
  const purchasesRoot = DriveApp.getFolderById(PURCHASES_FOLDER_ID);
  const customerFolder = findCustomerPoFolder_(purchasesRoot, customerNumber);
  if (!customerFolder) throw new Error(`Customer folder not found for ${customerNumber}.`);

  const customerTag = buildCustomerTag_(customerNumber);
  const targetFile = DriveApp.getFileById(targetReportId);
  const updated = [];
  const missing = [];

  months.forEach((month, idx) => {
    const progress = 15 + Math.floor(((idx + 1) / months.length) * 75);
    const monthName = monthNumberToName_(month);
    setStatus_(`Compiling ${reportType} totals for ${monthName} ${year}...`, progress, runId);

    const sourceFile = findReportFileByExactPeriod_(
      customerFolder,
      customerTag,
      testMode,
      reportType,
      year,
      month,
      targetReportId
    );

    if (!sourceFile) {
      missing.push(`${monthName}-${year}`);
      return;
    }

    copyMonthlyTotalsFromSourceToTarget_(sourceFile, targetFile, reportType, month);
    updated.push(`${monthName}-${year}`);
  });

  let summary =
    `Compile complete for ${reportType}.\n` +
    `Target:\n${targetFile.getName()}\n${targetFile.getUrl()}`;

  if (updated.length) {
    summary += `\n\nUpdated month rows:\n${updated.map(m => `• ${m}`).join('\n')}`;
  }
  if (missing.length) {
    summary += `\n\nMissing source reports:\n${missing.map(m => `• ${m}`).join('\n')}`;
  }

  setStatus_('Finalizing compile summary...', 98, runId);
  setStatus_(summary, 99, runId);
}

function normalizeCompileMonths_(monthsInput) {
  const raw = Array.isArray(monthsInput) ? monthsInput : [];
  const set = new Set();

  raw.forEach((value) => {
    const month = Number(value);
    if (!Number.isNaN(month) && month >= 1 && month <= 12) {
      set.add(month);
    }
  });

  return Array.from(set).sort((a, b) => a - b);
}

function collectReportFilesForCompile_(customerFolder, customerTag, testMode, reportType, year) {
  const matches = [];
  const seen = {};

  const folders = collectGeneratedReportFoldersRecursive_(
    customerFolder,
    customerTag,
    testMode,
    { maxFolders: 5000, maxDepth: 30 }
  );

  folders.forEach((folder) => {
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const file = files.next();
      if (seen[file.getId()]) continue;
      seen[file.getId()] = true;

      const parsed = parseReportFilePeriodAndKey_(file.getName());
      if (!parsed) continue;
      if (parsed.reportKey !== reportType || parsed.year !== year) continue;

      matches.push({
        file,
        month: parsed.month,
        year: parsed.year,
        updatedAt: file.getLastUpdated().getTime(),
      });
    }
  });

  matches.sort((a, b) => {
    if (a.month !== b.month) return a.month - b.month;
    return b.updatedAt - a.updatedAt;
  });

  return matches;
}

function findReportFileByExactPeriod_(
  customerFolder,
  customerTag,
  testMode,
  reportType,
  year,
  month,
  preferredFileId
) {
  const all = collectReportFilesForCompile_(customerFolder, customerTag, testMode, reportType, year)
    .filter(item => item.month === month);
  if (!all.length) return null;

  const preferred = all.find(item => item.file.getId() === preferredFileId);
  if (preferred) return preferred.file;

  all.sort((a, b) => b.updatedAt - a.updatedAt);
  return all[0].file;
}

function getMonthlyReportConfig_(reportType) {
  const map = {
    'NETCOST': { sheetName: 'NETCOST', startCol: 2, numCols: 17 },
    'DTSC-ALL': { sheetName: 'DTSC ALL ITEMS', startCol: 2, numCols: 13 },
    'DTSC': { sheetName: 'DTSC', startCol: 2, numCols: 11 },
  };
  return map[reportType] || null;
}

function copyMonthlyTotalsFromSourceToTarget_(sourceFile, targetFile, reportType, targetMonth) {
  const cfg = getMonthlyReportConfig_(reportType);
  if (!cfg) throw new Error(`Unsupported report type: ${reportType}`);

  const sourceSs = SpreadsheetApp.openById(sourceFile.getId());
  const targetSs = SpreadsheetApp.openById(targetFile.getId());
  const sourceSheet = sourceSs.getSheetByName(cfg.sheetName);
  const targetSheet = targetSs.getSheetByName(cfg.sheetName);
  if (!sourceSheet || !targetSheet) {
    throw new Error(`Required sheet "${cfg.sheetName}" not found in source/target report.`);
  }

  const totals = sourceSheet.getRange(20, cfg.startCol, 1, cfg.numCols).getValues();
  const monthCells = targetSheet.getRange('A25:A36').getDisplayValues();

  let targetRow = null;
  for (let i = 0; i < monthCells.length; i++) {
    const labelMonth = monthLabelToNumber_(monthCells[i][0]);
    if (labelMonth === targetMonth) {
      targetRow = 25 + i;
      break;
    }
  }

  if (!targetRow) {
    throw new Error(
      `Could not find month row for ${monthNumberToName_(targetMonth)} in ${cfg.sheetName}!A25:A36.`
    );
  }

  targetSheet.getRange(targetRow, cfg.startCol, 1, cfg.numCols).setValues(totals);
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

function writeDtscAllMonthlyTotals_(dtscAllFile, invoiceMonth) {
  const reportSs = SpreadsheetApp.openById(dtscAllFile.getId());
  const dtscAllSheet = reportSs.getSheetByName('DTSC ALL ITEMS');
  if (!dtscAllSheet) {
    throw new Error('DTSC-ALL report is missing required "DTSC ALL ITEMS" sheet.');
  }

  const totals = dtscAllSheet.getRange('B20:N20').getValues();
  const monthCells = dtscAllSheet.getRange('A25:A36').getDisplayValues();

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
      `Could not find month row for ${monthNumberToName_(invoiceMonth)} in DTSC ALL ITEMS!A25:A36.`
    );
  }

  dtscAllSheet.getRange(targetRow, 2, 1, 13).setValues(totals);
}

function writeDtscMonthlyTotals_(dtscFile, invoiceMonth) {
  const reportSs = SpreadsheetApp.openById(dtscFile.getId());
  const dtscSheet = reportSs.getSheetByName('DTSC');
  if (!dtscSheet) {
    throw new Error('DTSC report is missing required "DTSC" sheet.');
  }

  const totals = dtscSheet.getRange('B20:L20').getValues();
  const monthCells = dtscSheet.getRange('A25:A36').getDisplayValues();

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
      `Could not find month row for ${monthNumberToName_(invoiceMonth)} in DTSC!A25:A36.`
    );
  }

  dtscSheet.getRange(targetRow, 2, 1, 11).setValues(totals);
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