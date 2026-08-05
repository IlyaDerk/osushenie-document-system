/**
 * Ручная административная архивация листа «История изменений».
 * Автоматические триггеры и правило 90 дней намеренно не добавляются.
 */
const CHANGE_HISTORY_ARCHIVE_FOLDER_PROPERTY_ = 'CHANGE_HISTORY_ARCHIVE_FOLDER_ID';

/** Запрашивает дату отсечения и архивирует изменения строго раньше неё. */
function archiveChangeHistoryByDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Архивация истории изменений',
    'Введите дату отсечения в формате ДД.ММ.ГГГГ. Будут архивированы строки строго раньше этой даты.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return { cancelled: true };
  }
  const cutoffText = String(response.getResponseText() || '').trim();
  const cutoffDate = parseArchiveCutoffDate_(cutoffText);
  const startedAt = new Date();
  const userEmail = getActiveUserEmail_();
  let operationId = '';

  try {
    const result = withDocumentLock_(function () {
      operationId = generateOperationId_(startedAt);
      return archiveChangeHistoryUnderLock_(operationId, startedAt, userEmail, cutoffText, cutoffDate);
    });
    ui.alert('Архивация истории изменений', result.report.text, ui.ButtonSet.OK);
    return result.report;
  } catch (error) {
    const finishedAt = new Date();
    const counters = error.archiveState && error.archiveState.counters
      ? error.archiveState.counters
      : emptyArchiveCounters_();
    const report = buildArchiveReport_(
      SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
      cutoffText,
      counters,
      error.archiveState && error.archiveState.warnings ? error.archiveState.warnings : [],
      String(error.message || error),
      startedAt,
      finishedAt
    );
    try {
      withDocumentLock_(function () {
        if (!operationId) operationId = generateOperationId_(startedAt);
        writeArchiveOperationHistory_(operationId, startedAt, finishedAt, userEmail, report.text, counters, 1,
          SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR);
      });
    } catch (historyError) {
      // Best effort: исходная причина остаётся главной.
    }
    showCriticalOperationError_('Ошибка архивации истории изменений', report.text.split('\n\nКРИТИЧЕСКАЯ ОШИБКА')[0], String(error.message || error));
    throw error;
  }
}

/** Настраивает папку Drive для архивных Spreadsheet-файлов. */
function setupChangeHistoryArchiveFolder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Настройка архива истории изменений',
    'Введите ID папки Google Drive для архивных файлов.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const folderId = String(response.getResponseText() || '').trim();
  if (!folderId) throw new Error('ID папки архива не заполнен.');
  const folder = DriveApp.getFolderById(folderId);
  folder.getName();
  PropertiesService.getScriptProperties().setProperty(
    CHANGE_HISTORY_ARCHIVE_FOLDER_PROPERTY_,
    folderId
  );
  ui.alert('Настройка архива истории изменений', 'Папка архива сохранена: ' + folder.getName(), ui.ButtonSet.OK);
}

function archiveChangeHistoryUnderLock_(operationId, startedAt, userEmail, cutoffText, cutoffDate) {
  assertSystemSheetsStructure_(['CHANGE_HISTORY', 'OPERATION_HISTORY']);
  const folderId = PropertiesService.getScriptProperties().getProperty(CHANGE_HISTORY_ARCHIVE_FOLDER_PROPERTY_);
  if (!folderId) throw new Error('Не настроена папка архива. Запустите setupChangeHistoryArchiveFolder.');
  const folder = DriveApp.getFolderById(folderId);
  folder.getName();

  const context = getSystemSheetContext_('CHANGE_HISTORY');
  const rows = readArchiveSourceRows_(context);
  const state = { counters: emptyArchiveCounters_(), warnings: [] };
  state.counters.checkedRows = rows.length;
  assertArchiveSourceIds_(context, rows, state);

  const selected = rows.filter(function (item) {
    const changedAt = item.values[archiveColumnIndex_(context, H.CHANGE_DATETIME)];
    return changedAt instanceof Date && !isNaN(changedAt.getTime()) && changedAt.getTime() < cutoffDate.getTime();
  });
  state.counters.selectedRows = selected.length;
  if (selected.length === 0) {
    const finishedAt = new Date();
    const report = buildArchiveReport_(SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES, cutoffText, state.counters, state.warnings, '', startedAt, finishedAt);
    writeArchiveOperationHistory_(operationId, startedAt, finishedAt, userEmail, report.text, state.counters, 0, SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES);
    return { report: report };
  }

  const groups = groupArchiveRowsByMonth_(context, selected);
  state.counters.archiveFiles = Object.keys(groups).length;
  Object.keys(groups).forEach(function (monthKey) {
    processArchiveMonth_(folder, context, monthKey, groups[monthKey], state);
  });

  deleteArchiveSourceRows_(context.sheet, selected.map(function (item) { return item.sheetRow; }));
  state.counters.deletedRows = selected.length;
  const finishedAt = new Date();
  const status = state.warnings.length
    ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS
    : SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS;
  const report = buildArchiveReport_(status, cutoffText, state.counters, state.warnings, '', startedAt, finishedAt);
  writeArchiveOperationHistory_(operationId, startedAt, finishedAt, userEmail, report.text, state.counters, 0, status);
  return { report: report };
}

function processArchiveMonth_(folder, sourceContext, monthKey, items, state) {
  const spreadsheet = getOrCreateArchiveSpreadsheet_(folder, monthKey);
  const archiveSheet = getOrCreateArchiveSheet_(spreadsheet, sourceContext);
  const archiveContext = buildArchiveContext_(archiveSheet, sourceContext);
  const existing = readArchiveExistingById_(archiveContext);
  const rowsToWrite = [];
  items.forEach(function (item) {
    const id = archiveChangeId_(sourceContext, item.values);
    if (existing[id]) {
      if (!archiveRowsEqual_(existing[id], item.values)) {
        archiveFail_(state, 'В архиве ' + monthKey + ' уже есть ID изменения «' + id + '» с отличающимися данными.');
      }
      state.counters.alreadyArchived += 1;
      return;
    }
    rowsToWrite.push(item.values);
  });
  if (rowsToWrite.length > 0) {
    const startRow = Math.max(archiveSheet.getLastRow() + 1, sourceContext.config.dataStartRow);
    archiveSheet.getRange(startRow, 1, rowsToWrite.length, sourceContext.headers.length).setValues(rowsToWrite);
    state.counters.writtenRows += rowsToWrite.length;
  }
  const after = readArchiveExistingById_(archiveContext);
  items.forEach(function (item) {
    const id = archiveChangeId_(sourceContext, item.values);
    if (!after[id]) archiveFail_(state, 'После записи архив ' + monthKey + ' не содержит ID изменения «' + id + '».');
    if (!archiveRowsEqual_(after[id], item.values)) archiveFail_(state, 'После записи архив ' + monthKey + ' содержит отличающиеся данные для ID изменения «' + id + '».');
  });
}

function readArchiveSourceRows_(context) {
  const lastRow = context.sheet.getLastRow();
  if (lastRow < context.config.dataStartRow) return [];
  return context.sheet.getRange(context.config.dataStartRow, 1, lastRow - context.config.dataStartRow + 1, context.headers.length)
    .getValues()
    .map(function (row, offset) { return { sheetRow: context.config.dataStartRow + offset, values: row }; });
}

function assertArchiveSourceIds_(context, rows, state) {
  const seen = {};
  rows.forEach(function (item) {
    const id = archiveChangeId_(context, item.values);
    if (!id) archiveFail_(state, 'В строке ' + item.sheetRow + ' пустой ID изменения. Архивация остановлена.');
    if (seen[id]) archiveFail_(state, 'ID изменения «' + id + '» повторяется в строках ' + seen[id] + ', ' + item.sheetRow + '. Архивация остановлена.');
    seen[id] = item.sheetRow;
  });
}

function groupArchiveRowsByMonth_(context, rows) {
  const groups = {};
  rows.forEach(function (item) {
    const date = item.values[archiveColumnIndex_(context, H.CHANGE_DATETIME)];
    const monthKey = Utilities.formatDate(date, getSystemSpreadsheet_().getSpreadsheetTimeZone(), 'yyyy-MM');
    if (!groups[monthKey]) groups[monthKey] = [];
    groups[monthKey].push(item);
  });
  return groups;
}

function getOrCreateArchiveSpreadsheet_(folder, monthKey) {
  const name = 'Осушение — Архив истории изменений ' + monthKey;
  const files = folder.getFilesByName(name);
  if (files.hasNext()) return SpreadsheetApp.openById(files.next().getId());
  const spreadsheet = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(spreadsheet.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (error) {}
  return spreadsheet;
}

function getOrCreateArchiveSheet_(spreadsheet, sourceContext) {
  let sheet = spreadsheet.getSheetByName(sourceContext.config.name);
  if (!sheet) sheet = spreadsheet.insertSheet(sourceContext.config.name);
  ensureArchiveHeaders_(sheet, sourceContext);
  return sheet;
}

function ensureArchiveHeaders_(sheet, sourceContext) {
  const headerRow = sourceContext.config.headerRow;
  const current = sheet.getRange(headerRow, 1, 1, sourceContext.headers.length).getValues()[0];
  const hasAny = current.some(function (value) { return String(value || '').trim() !== ''; });
  if (!hasAny) {
    sheet.getRange(headerRow, 1, 1, sourceContext.headers.length).setValues([sourceContext.headers]);
    return;
  }
  sourceContext.headers.forEach(function (header, index) {
    if (sysNormalizeHeader_(current[index]) !== sysNormalizeHeader_(header)) {
      throw new Error('Архивный лист «' + sourceContext.config.name + '» имеет некорректный заголовок «' + header + '».');
    }
  });
}

function buildArchiveContext_(sheet, sourceContext) {
  return { sheet: sheet, config: sourceContext.config, headers: sourceContext.headers, headerMap: sourceContext.headerMap };
}

function readArchiveExistingById_(context) {
  const result = {};
  readArchiveSourceRows_(context).forEach(function (item) {
    const id = archiveChangeId_(context, item.values);
    if (id) result[id] = item.values;
  });
  return result;
}

function deleteArchiveSourceRows_(sheet, rowNumbers) {
  rowNumbers.sort(function (a, b) { return b - a; });
  let index = 0;
  while (index < rowNumbers.length) {
    const endRow = rowNumbers[index];
    let startRow = endRow;
    index += 1;
    while (index < rowNumbers.length && rowNumbers[index] === startRow - 1) {
      startRow = rowNumbers[index];
      index += 1;
    }
    sheet.deleteRows(startRow, endRow - startRow + 1);
  }
}

function parseArchiveCutoffDate_(text) {
  const match = String(text || '').trim().match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!match) throw new Error('Дата отсечения должна быть в формате ДД.ММ.ГГГГ.');
  const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  if (date.getFullYear() !== Number(match[3]) || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[1])) {
    throw new Error('Дата отсечения некорректна: ' + text);
  }
  return date;
}

function archiveColumnIndex_(context, header) { return context.headerMap[sysNormalizeHeader_(header)] - 1; }
function archiveChangeId_(context, row) { return String(row[archiveColumnIndex_(context, H.CHANGE_ID)] == null ? '' : row[archiveColumnIndex_(context, H.CHANGE_ID)]).trim(); }
function archiveRowsEqual_(left, right) {
  if (!left || !right || left.length !== right.length) return false;
  return left.every(function (value, index) {
    const other = right[index];
    if (value instanceof Date || other instanceof Date) {
      return value instanceof Date && other instanceof Date && value.getTime() === other.getTime();
    }
    return String(value == null ? '' : value) === String(other == null ? '' : other);
  });
}
function archiveFail_(state, message) {
  const error = new Error(message);
  error.archiveState = state;
  throw error;
}
function emptyArchiveCounters_() {
  return { checkedRows: 0, selectedRows: 0, alreadyArchived: 0, writtenRows: 0, deletedRows: 0, archiveFiles: 0 };
}
function buildArchiveReport_(status, cutoffText, counters, warnings, critical, startedAt, finishedAt) {
  const lines = [
    'Статус: ' + status,
    'Дата отсечения: ' + cutoffText,
    'Строк истории проверено: ' + counters.checkedRows,
    'Строк выбрано для архивации: ' + counters.selectedRows,
    'Строк уже находилось в архиве: ' + counters.alreadyArchived,
    'Строк записано в архив: ' + counters.writtenRows,
    'Строк удалено из рабочего файла: ' + counters.deletedRows,
    'Архивных файлов: ' + counters.archiveFiles,
    'Время выполнения, сек.: ' + Math.max(0, ((finishedAt.getTime() - startedAt.getTime()) / 1000).toFixed(3))
  ];
  if (warnings.length) {
    lines.push('Предупреждения:');
    warnings.forEach(function (warning) { lines.push('- ' + warning); });
  }
  if (critical) lines.push('', 'КРИТИЧЕСКАЯ ОШИБКА', critical);
  return { status: status, counters: counters, text: lines.join('\n') };
}
function writeArchiveOperationHistory_(operationId, startedAt, finishedAt, userEmail, comment, counters, errorsCount, status) {
  const context = getSystemSheetContext_('OPERATION_HISTORY');
  const values = {};
  values[H.OPERATION_ID] = operationId;
  values[H.OPERATION_STARTED_AT] = startedAt;
  values[H.OPERATION_FINISHED_AT] = finishedAt;
  values[H.OPERATION_STARTED_BY] = userEmail;
  values[H.OPERATION_SOURCE] = SYSTEM_CONFIG.VALUES.CHANGE_HISTORY_ARCHIVE_SOURCE;
  values[H.OPERATION_TYPE] = SYSTEM_CONFIG.VALUES.CHANGE_HISTORY_ARCHIVE_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = status;
  values[H.DOCUMENTS_LOADED] = 0;
  values[H.DOCUMENTS_CHANGED] = 0;
  values[H.FACT_ROWS_UPDATED] = counters.deletedRows;
  values[H.FIELDS_CHANGED] = 0;
  values[H.DUPLICATE_IDS_FOUND] = 0;
  values[H.ERRORS_COUNT] = errorsCount;
  values[H.EXECUTION_SECONDS] = Math.max(0, (finishedAt.getTime() - startedAt.getTime()) / 1000);
  values[H.ERROR_TEXT] = comment;
  const row = new Array(context.headers.length).fill('');
  Object.keys(values).forEach(function (header) { row[archiveColumnIndex_(context, header)] = values[header]; });
  const startRow = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(startRow, 1, 1, context.headers.length).setValues([row]);
}
