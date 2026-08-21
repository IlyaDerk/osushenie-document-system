/** Поля объекта, переносимые из листа «Объекты» в активные документы. */
const OBJECT_SYNC_FIELDS_ = Object.freeze([
  H.OBJECT_NAME,
  H.CONTRACT_NUMBER,
  H.OBJECT_STATUS,
  H.RESPONSIBLE_FOREMAN,
  H.RESPONSIBLE_FOREMAN_ID
]);

/** Guards v2 document-owned fields from accidental object synchronization. */
function assertObjectSyncFieldContract_() {
  const documentOwned = [
    H.DOCUMENT_TYPE,
    H.DOCUMENT_TYPE_ID,
    H.DOCUMENT_NUMBER,
    H.DOCUMENT_DATE,
    H.CUSTOMER_SIGNING_RESPONSIBLE,
    // These dates are copied only by document creation. Afterwards they belong
    // to the document and may be changed only by the dedicated card save flow.
    H.WORK_START_DATE,
    H.WORK_END_PLAN,
    H.WORK_END_FACT
  ];
  const conflicts = OBJECT_SYNC_FIELDS_.filter(function (header) {
    return documentOwned.indexOf(header) !== -1;
  });
  if (conflicts.length > 0) {
    throw new Error(
      'Ошибка контракта синхронизации: документные поля не могут ' +
      'копироваться из листа «' + SYSTEM_CONFIG.SHEETS.OBJECTS.name + '»: ' +
      conflicts.join(', ')
    );
  }
}

/** Ручной запуск кнопки «Синхронизировать данные объектов». */
function syncObjectDataToDocuments() {
  const startedAt = new Date();
  const userEmail = getActiveUserEmail_();
  let operationId = '';
  try {
    const result = withDocumentLock_(function () {
      operationId = generateOperationId_(startedAt);
      return syncObjectDataUnderLock_(operationId, startedAt, userEmail);
    });
    showObjectSyncReport_(result.report, false);
    return result.report;
  } catch (error) {
    const state = error.objectSyncState || emptyObjectSyncState_();
    const finishedAt = new Date();
    const reason = String(error.message || error);
    const report = buildObjectSyncReport_(
      SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
      state.counters,
      state.skips,
      state.warnings,
      reason,
      startedAt,
      finishedAt
    );
    try {
      withDocumentLock_(function () {
        if (!operationId) {
          operationId = generateOperationId_(startedAt);
        }
        writeObjectSyncOperationHistory_({
          operationId: operationId,
          startedAt: startedAt,
          finishedAt: finishedAt,
          userEmail: userEmail,
          status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
          counters: state.counters,
          comment: report.text,
          errorsCount: 1
        });
      });
    } catch (historyError) {
      // Best effort; исходная причина остаётся главной.
    }
    showObjectSyncReport_(report, true);
    throw error;
  }
}

/** Вся проверка, повторное чтение и запись выполняются под общей блокировкой. */
function syncObjectDataUnderLock_(operationId, startedAt, userEmail) {
  assertObjectSyncFieldContract_();
  assertSystemSheetsStructure_([
    'OBJECTS',
    'DOCUMENTS',
    'CARD_DICTIONARY',
    'OPERATION_HISTORY',
    'CHANGE_HISTORY'
  ]);
  assertObjectSyncActiveDictionaryValue_();

  const objectResult = readObjectSyncSource_();
  const documentResult = readObjectSyncDocuments_();
  const now = new Date();
  const plan = buildObjectSyncPlan_(
    objectResult.validById,
    documentResult,
    now,
    userEmail,
    operationId
  );
  const counters = {
    checkedObjects: objectResult.checkedObjects,
    validObjects: Object.keys(objectResult.validById).length,
    skippedObjects: objectResult.skips.length,
    activeDocumentRowsChecked: documentResult.activeRowsChecked,
    documentsChanged: plan.changedDocumentIds.length,
    factRowsUpdated: plan.changedRows.length,
    fieldsChanged: plan.changes.length,
    duplicateIdsFound: documentResult.duplicateDocumentIds.length
  };
  const warnings = objectResult.warnings.concat(documentResult.warnings);
  const state = { counters: counters, skips: objectResult.skips, warnings: warnings };

  if (plan.changes.length > 9999) {
    throwObjectSyncError_(
      new Error('Исчерпан предел 9999 изменений внутри операции. Запись данных отменена.'),
      state
    );
  }

  if (plan.changedRows.length > 0) {
    writeObjectSyncFacts_(documentResult.context, plan.changedRows);
    try {
      writeObjectSyncChangeHistory_(plan.changes);
    } catch (error) {
      throwObjectSyncError_(
        new Error(
          'Данные документов обновлены, но история изменений записана не полностью: ' +
          error.message
        ),
        state
      );
    }
  }

  const status = plan.changedRows.length === 0
    ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES
    : warnings.length > 0 || objectResult.skips.length > 0
      ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS
      : SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS;
  const finishedAt = new Date();
  const report = buildObjectSyncReport_(
    status,
    counters,
    objectResult.skips,
    warnings,
    '',
    startedAt,
    finishedAt
  );
  try {
    writeObjectSyncOperationHistory_({
      operationId: operationId,
      startedAt: startedAt,
      finishedAt: finishedAt,
      userEmail: userEmail,
      status: status,
      counters: counters,
      comment: report.text,
      errorsCount: 0
    });
  } catch (error) {
    throwObjectSyncError_(
      new Error(
        (plan.changedRows.length > 0 ? 'Данные документов обновлены, но ' : '') +
        'история операций не записана: ' + error.message
      ),
      state
    );
  }
  return { report: report };
}

function assertObjectSyncActiveDictionaryValue_() {
  const context = getSystemSheetContext_('CARD_DICTIONARY');
  const rows = objectSyncReadRows_(context);
  const index = objectSyncColumnIndex_(context, H.RECORD_STATUS);
  const found = rows.some(function (row) {
    return objectSyncNormalized_(row[index]) ===
      objectSyncNormalized_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS);
  });
  if (!found) {
    throw new Error(
      'На листе «' + context.config.name + '» отсутствует системное значение «' +
      SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS + '» в поле «' +
      H.RECORD_STATUS + '».'
    );
  }
}

function readObjectSyncSource_() {
  const context = getSystemSheetContext_('OBJECTS');
  const rows = objectSyncReadRows_(context);
  const indexes = {};
  [H.OBJECT_ID].concat(OBJECT_SYNC_FIELDS_).forEach(function (header) {
    indexes[header] = objectSyncColumnIndex_(context, header);
  });
  const candidates = [];
  const idRows = {};
  rows.forEach(function (row, offset) {
    if (objectSyncRowEmpty_(row)) return;
    const sheetRow = context.config.dataStartRow + offset;
    const id = objectSyncText_(row[indexes[H.OBJECT_ID]]);
    if (!id) return;
    candidates.push({ id: id, row: row, sheetRow: sheetRow });
    if (!idRows[id]) idRows[id] = [];
    idRows[id].push(sheetRow);
  });
  const validById = {};
  const skips = [];
  candidates.forEach(function (candidate) {
    const reasons = [];
    if (idRows[candidate.id].length > 1) {
      reasons.push('ID объекта повторяется в строках ' + idRows[candidate.id].join(', '));
    }
    [H.WORK_START_DATE, H.WORK_END_PLAN, H.WORK_END_FACT].forEach(function (header) {
      const value = candidate.row[indexes[header]];
      if (!objectSyncEmpty_(value) && !objectSyncValidDate_(value)) {
        reasons.push('поле «' + header + '» не является корректной датой');
      }
    });
    const start = candidate.row[indexes[H.WORK_START_DATE]];
    const plan = candidate.row[indexes[H.WORK_END_PLAN]];
    if (objectSyncValidDate_(start) && objectSyncValidDate_(plan) &&
        plan.getTime() < start.getTime()) {
      reasons.push('плановое окончание раньше даты начала работ');
    }
    if (reasons.length > 0) {
      skips.push({ sheetRow: candidate.sheetRow, id: candidate.id, reasons: reasons });
      return;
    }
    const values = {};
    OBJECT_SYNC_FIELDS_.forEach(function (header) {
      values[header] = candidate.row[indexes[header]];
    });
    validById[candidate.id] = values;
  });
  return {
    checkedObjects: candidates.length,
    validById: validById,
    skips: skips,
    warnings: []
  };
}

function readObjectSyncDocuments_() {
  const context = getSystemSheetContext_('DOCUMENTS');
  const rows = objectSyncReadRows_(context);
  const indexes = {};
  [H.DOCUMENT_ID, H.OBJECT_ID, H.RECORD_STATUS, H.UPDATED_AT, H.UPDATED_BY_EMAIL]
    .concat(OBJECT_SYNC_FIELDS_).forEach(function (header) {
      indexes[header] = objectSyncColumnIndex_(context, header);
    });
  const activeRows = [];
  let activeRowsChecked = 0;
  const warnings = [];
  const documentIdRows = {};
  rows.forEach(function (row, offset) {
    const sheetRow = context.config.dataStartRow + offset;
    if (objectSyncNormalized_(row[indexes[H.RECORD_STATUS]]) !==
        objectSyncNormalized_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS)) return;
    activeRowsChecked += 1;
    const documentId = objectSyncText_(row[indexes[H.DOCUMENT_ID]]);
    const objectId = objectSyncText_(row[indexes[H.OBJECT_ID]]);
    if (!documentId || !objectId) {
      warnings.push(
        'Активная строка документа ' + sheetRow + ' пропущена: отсутствует ' +
        (!documentId ? H.DOCUMENT_ID : H.OBJECT_ID) + '.'
      );
      return;
    }
    activeRows.push({
      row: row.slice(),
      sheetRow: sheetRow,
      documentId: documentId,
      objectId: objectId
    });
    if (!documentIdRows[documentId]) documentIdRows[documentId] = [];
    documentIdRows[documentId].push(sheetRow);
  });
  const duplicateDocumentIds = Object.keys(documentIdRows).filter(function (id) {
    return documentIdRows[id].length > 1;
  });
  duplicateDocumentIds.forEach(function (id) {
    warnings.push(
      'ID документа «' + id + '» повторяется в активных строках ' +
      documentIdRows[id].join(', ') + '; обновлены все подходящие строки.'
    );
  });
  return {
    context: context,
    indexes: indexes,
    activeRows: activeRows,
    activeRowsChecked: activeRowsChecked,
    duplicateDocumentIds: duplicateDocumentIds,
    warnings: warnings
  };
}

function buildObjectSyncPlan_(objects, documents, now, userEmail, operationId) {
  const changedRows = [];
  const changes = [];
  const changedIds = {};
  documents.activeRows.forEach(function (document) {
    const source = objects[document.objectId];
    if (!source) return;
    const rowChanges = [];
    OBJECT_SYNC_FIELDS_.forEach(function (header) {
      const index = documents.indexes[header];
      if (!objectSyncEqual_(document.row[index], source[header], objectSyncDateHeader_(header))) {
        rowChanges.push({ header: header, oldValue: document.row[index], newValue: source[header] });
        document.row[index] = source[header];
      }
    });
    if (rowChanges.length === 0) return;
    const businessChanges = rowChanges.slice();
    const updatedAtIndex = documents.indexes[H.UPDATED_AT];
    rowChanges.push({ header: H.UPDATED_AT, oldValue: document.row[updatedAtIndex], newValue: now });
    document.row[updatedAtIndex] = now;
    const updatedByIndex = documents.indexes[H.UPDATED_BY_EMAIL];
    if (!objectSyncEqual_(document.row[updatedByIndex], userEmail, false)) {
      rowChanges.push({
        header: H.UPDATED_BY_EMAIL,
        oldValue: document.row[updatedByIndex],
        newValue: userEmail
      });
      document.row[updatedByIndex] = userEmail;
    }
    businessChanges.forEach(function (change) {
      changes.push({
        changeId: generateChangeId_(operationId, changes.length + 1),
        operationId: operationId,
        changedAt: now,
        userEmail: userEmail,
        documentId: document.documentId,
        objectId: document.objectId,
        factRow: document.sheetRow,
        fieldName: change.header,
        oldValue: objectSyncHistoryValue_(change.oldValue, objectSyncDateHeader_(change.header)),
        newValue: objectSyncHistoryValue_(change.newValue, objectSyncDateHeader_(change.header))
      });
    });
    document.syncWriteHeaders = rowChanges.map(function (change) {
      return change.header;
    });
    changedRows.push(document);
    changedIds[document.documentId] = true;
  });
  return {
    changedRows: changedRows,
    changes: changes,
    changedDocumentIds: Object.keys(changedIds)
  };
}

/**
 * Пакетно пишет только разрешённые изменившиеся поля, группируя соседние
 * физические строки внутри каждого столбца. Остальные столбцы не затрагиваются.
 */
function writeObjectSyncFacts_(context, changedRows) {
  const writableHeaders = OBJECT_SYNC_FIELDS_.concat([
    H.UPDATED_AT,
    H.UPDATED_BY_EMAIL
  ]);
  writableHeaders.forEach(function (header) {
    const columnIndex = objectSyncColumnIndex_(context, header);
    const items = changedRows.filter(function (item) {
      return item.syncWriteHeaders.indexOf(header) !== -1;
    }).sort(function (left, right) {
      return left.sheetRow - right.sheetRow;
    });
    const groups = [];
    items.forEach(function (item) {
      const group = groups[groups.length - 1];
      const value = [item.row[columnIndex]];
      if (group && item.sheetRow === group.startRow + group.values.length) {
        group.values.push(value);
      } else {
        groups.push({ startRow: item.sheetRow, values: [value] });
      }
    });
    groups.forEach(function (group) {
      context.sheet.getRange(
        group.startRow,
        columnIndex + 1,
        group.values.length,
        1
      ).setValues(group.values);
    });
  });
}

function writeObjectSyncChangeHistory_(changes) {
  if (changes.length === 0) return;
  const context = getSystemSheetContext_('CHANGE_HISTORY');
  const rows = changes.map(function (change) {
    const values = {};
    values[H.CHANGE_ID] = change.changeId;
    values[H.OPERATION_ID] = change.operationId;
    values[H.CHANGE_DATETIME] = change.changedAt;
    values[H.CHANGED_BY_EMAIL] = change.userEmail;
    values[H.ACTION_TYPE] = SYSTEM_CONFIG.VALUES.CHANGE_ACTION_SYNC;
    values[H.DOCUMENT_ID] = change.documentId;
    values[H.OBJECT_ID] = change.objectId;
    values[H.FACT_ROW_NUMBER] = change.factRow;
    values[H.FIELD_NAME] = change.fieldName;
    values[H.OLD_VALUE] = change.oldValue;
    values[H.NEW_VALUE] = change.newValue;
    values[H.CHANGE_SOURCE] = SYSTEM_CONFIG.VALUES.OBJECT_SHEET_CHANGE_SOURCE;
    return objectSyncRowFromValues_(context, values);
  });
  const startRow = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(startRow, 1, rows.length, context.headers.length).setValues(rows);
}

function writeObjectSyncOperationHistory_(operation) {
  const context = getSystemSheetContext_('OPERATION_HISTORY');
  const values = {};
  values[H.OPERATION_ID] = operation.operationId;
  values[H.OPERATION_STARTED_AT] = operation.startedAt;
  values[H.OPERATION_FINISHED_AT] = operation.finishedAt;
  values[H.OPERATION_STARTED_BY] = operation.userEmail;
  values[H.OPERATION_SOURCE] = SYSTEM_CONFIG.VALUES.OBJECT_DATA_SYNC_SOURCE;
  values[H.OPERATION_TYPE] = SYSTEM_CONFIG.VALUES.OBJECT_DATA_SYNC_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = operation.status;
  values[H.DOCUMENTS_LOADED] = 0;
  values[H.DOCUMENTS_CHANGED] = operation.counters.documentsChanged;
  values[H.FACT_ROWS_UPDATED] = operation.counters.factRowsUpdated;
  values[H.FIELDS_CHANGED] = operation.counters.fieldsChanged;
  values[H.DUPLICATE_IDS_FOUND] = operation.counters.duplicateIdsFound;
  values[H.ERRORS_COUNT] = operation.errorsCount;
  values[H.EXECUTION_SECONDS] = Math.max(
    0,
    (operation.finishedAt.getTime() - operation.startedAt.getTime()) / 1000
  );
  values[H.ERROR_TEXT] = operation.comment;
  const row = objectSyncRowFromValues_(context, values);
  const startRow = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(startRow, 1, 1, context.headers.length).setValues([row]);
}

function buildObjectSyncReport_(status, counters, skips, warnings, critical, startedAt, finishedAt) {
  const lines = [
    'Статус: ' + status,
    'Объектов проверено: ' + counters.checkedObjects,
    'Корректных объектов: ' + counters.validObjects,
    'Пропущено объектов: ' + counters.skippedObjects,
    'Активных строк документов проверено: ' + counters.activeDocumentRowsChecked,
    'Документов с изменениями: ' + counters.documentsChanged,
    'Строк обновлено: ' + counters.factRowsUpdated,
    'Полей изменено: ' + counters.fieldsChanged,
    'Дублирующихся ID найдено: ' + counters.duplicateIdsFound,
    'Время выполнения, сек.: ' + Math.max(
      0,
      ((finishedAt.getTime() - startedAt.getTime()) / 1000).toFixed(3)
    )
  ];
  if (skips.length) {
    lines.push('Причины пропусков:');
    skips.forEach(function (item) {
      lines.push('- строка ' + item.sheetRow + ', ID ' + item.id + ': ' + item.reasons.join('; '));
    });
  }
  if (warnings.length) {
    lines.push('Предупреждения:');
    warnings.forEach(function (warning) { lines.push('- ' + warning); });
  }
  if (critical) {
    lines.push('', 'КРИТИЧЕСКАЯ ОШИБКА', critical);
  }
  return { status: status, counters: counters, text: lines.join('\n') };
}

function showObjectSyncReport_(report, isError) {
  try {
    if (isError) {
      const marker = 'КРИТИЧЕСКАЯ ОШИБКА\n';
      const position = report.text.lastIndexOf(marker);
      showCriticalOperationError_(
        'Ошибка синхронизации объектов',
        position < 0 ? report.text : report.text.slice(0, position).trim(),
        position < 0 ? report.text : report.text.slice(position + marker.length)
      );
      return;
    }
    SpreadsheetApp.getUi().alert(
      'Синхронизация данных объектов',
      report.text,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (uiError) {
    // Возвращаемый отчёт и история остаются доступны без UI.
  }
}

function emptyObjectSyncState_() {
  return {
    counters: {
      checkedObjects: 0,
      validObjects: 0,
      skippedObjects: 0,
      activeDocumentRowsChecked: 0,
      documentsChanged: 0,
      factRowsUpdated: 0,
      fieldsChanged: 0,
      duplicateIdsFound: 0
    },
    skips: [],
    warnings: []
  };
}

function throwObjectSyncError_(error, state) {
  error.objectSyncState = state;
  throw error;
}

function objectSyncReadRows_(context) {
  const lastRow = context.sheet.getLastRow();
  if (lastRow < context.config.dataStartRow) return [];
  return context.sheet.getRange(
    context.config.dataStartRow,
    1,
    lastRow - context.config.dataStartRow + 1,
    context.headers.length
  ).getValues();
}
function objectSyncColumnIndex_(context, header) {
  return context.headerMap[sysNormalizeHeader_(header)] - 1;
}
function objectSyncRowFromValues_(context, values) {
  const row = new Array(context.headers.length).fill('');
  Object.keys(values).forEach(function (header) {
    row[objectSyncColumnIndex_(context, header)] = values[header];
  });
  return row;
}
function objectSyncText_(value) {
  return String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim();
}
function objectSyncNormalized_(value) {
  return objectSyncText_(value).replace(/\s+/g, ' ').toLocaleLowerCase();
}
function objectSyncEmpty_(value) {
  return value == null || (typeof value === 'string' && objectSyncText_(value) === '');
}
function objectSyncRowEmpty_(row) { return row.every(objectSyncEmpty_); }
function objectSyncValidDate_(value) {
  return value instanceof Date && !isNaN(value.getTime());
}
function objectSyncDateHeader_(header) {
  return header === H.WORK_START_DATE || header === H.WORK_END_PLAN ||
    header === H.WORK_END_FACT || header === H.UPDATED_AT;
}
function objectSyncEqual_(left, right, isDate) {
  if (objectSyncEmpty_(left) && objectSyncEmpty_(right)) return true;
  if (isDate) {
    return objectSyncValidDate_(left) && objectSyncValidDate_(right) &&
      left.getTime() === right.getTime();
  }
  return left === right;
}
function objectSyncHistoryValue_(value, isDate) {
  if (objectSyncEmpty_(value)) return '';
  if (isDate && objectSyncValidDate_(value)) {
    const timezone = getSystemSpreadsheet_().getSpreadsheetTimeZone();
    const pattern = value.getHours() || value.getMinutes() || value.getSeconds()
      ? 'dd.MM.yyyy HH:mm:ss'
      : 'dd.MM.yyyy';
    return Utilities.formatDate(value, timezone, pattern);
  }
  return String(value);
}
