/**
 * Массово создаёт недостающие документы для всех корректных объектов.
 * Функция предназначена для назначения кнопке Google Sheets.
 */
function createMissingDocumentsForAllObjects() {
  const startedAt = new Date();
  let operationId = '';
  const userEmail = getActiveUserEmail_();
  let result;

  try {
    result = withDocumentLock_(function () {
      operationId = generateOperationId_(startedAt);
      return createObjectDocumentsUnderLock_(
        operationId,
        startedAt,
        userEmail
      );
    });

    showCreateDocumentsReport_(result.report, false);
    return result.report;
  } catch (error) {
    const failure = buildCreationFailure_(
      error,
      operationId,
      startedAt,
      userEmail
    );

    try {
      withDocumentLock_(function () {
        if (!operationId) {
          operationId = generateOperationId_(startedAt);
          failure.operationRow.operationId = operationId;
        }
        writeCreationOperationHistory_(failure.operationRow);
      });
    } catch (historyError) {
      // Best effort: исходная ошибка всегда имеет приоритет.
    }

    showCreateDocumentsReport_(failure.report, true);
    throw error;
  }
}


/** Выполняет проверку, повторное чтение и все записи под общей блокировкой. */
function createObjectDocumentsUnderLock_(operationId, startedAt, userEmail) {
  assertSystemSheetsStructure_([
    'OBJECTS',
    'DOCUMENT_TYPES',
    'DOCUMENTS',
    'CARD_DICTIONARY',
    'OPERATION_HISTORY',
    'CHANGE_HISTORY'
  ]);

  assertCreationDictionaryValues_();

  const rules = readAutomaticDocumentRules_();
  const objectResult = readAndValidateCreationObjects_();
  const factResult = readExistingCreationFacts_();
  const plan = buildMissingDocumentsPlan_(
    objectResult.validObjects,
    rules,
    factResult
  );
  const counters = {
    checkedObjects: objectResult.checkedObjects,
    validObjects: objectResult.validObjects.length,
    skippedObjects: objectResult.skippedObjects.length,
    creationRules: rules.length,
    existingDocuments: plan.existingDocuments,
    createdDocuments: 0,
    duplicateObjectTypeKeys: factResult.duplicateObjectTypeKeys.length,
    duplicateDocumentIds: factResult.duplicateDocumentIds.length,
    incompleteFactKeys: factResult.incompleteFactRows.length,
    changeRows: 0
  };
  const warnings = factResult.warnings.slice();
  const documentsContext = getSystemSheetContext_('DOCUMENTS');
  const firstFactRow = Math.max(
    documentsContext.sheet.getLastRow() + 1,
    documentsContext.config.dataStartRow
  );
  const now = new Date();
  const prepared = prepareDocumentRows_(
    plan.missing,
    factResult,
    documentsContext,
    firstFactRow,
    now,
    userEmail,
    operationId
  );

  if (prepared.documentRows.length > 0) {
    documentsContext.sheet
      .getRange(
        firstFactRow,
        1,
        prepared.documentRows.length,
        documentsContext.headers.length
      )
      .setValues(prepared.documentRows);
    counters.createdDocuments = prepared.documentRows.length;

    try {
      writeCreationChangeHistory_(prepared.changeRows);
      counters.changeRows = prepared.changeRows.length;
    } catch (error) {
      const message =
        'Документы созданы, но запись истории изменений ' +
        'завершилась ошибкой';
      const wrapped = new Error(message + ': ' + error.message);
      wrapped.creationState = {
        counters: counters,
        warnings: warnings,
        skippedObjects: objectResult.skippedObjects,
        criticalMessage: message
      };
      throw wrapped;
    }
  }

  const status = warnings.length > 0 || objectResult.skippedObjects.length > 0
    ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS
    : counters.createdDocuments > 0
      ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS
      : SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES;
  const finishedAt = new Date();
  const report = buildCreationReport_(
    status,
    counters,
    objectResult.skippedObjects,
    warnings,
    '',
    startedAt,
    finishedAt
  );

  try {
    writeCreationOperationHistory_({
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
    const wrapped = new Error(
      'Документы созданы, но запись истории операций завершилась ошибкой: ' +
      error.message
    );
    wrapped.creationState = {
      counters: counters,
      warnings: warnings,
      skippedObjects: objectResult.skippedObjects,
      criticalMessage: wrapped.message
    };
    throw wrapped;
  }

  return { report: report };
}


/** Проверяет обязательные значения справочника карточки. */
function assertCreationDictionaryValues_() {
  const context = getSystemSheetContext_('CARD_DICTIONARY');
  const values = readCreationSheetValues_(context);
  const statusIndex = creationColumnIndex_(context, H.DOCUMENT_STATUS);
  const recordIndex = creationColumnIndex_(context, H.RECORD_STATUS);
  const hasInitialStatus = values.some(function (row) {
    return creationNormalizedValue_(row[statusIndex]) ===
      creationNormalizedValue_(
        SYSTEM_CONFIG.VALUES.INITIAL_DOCUMENT_STATUS
      );
  });
  const hasActiveStatus = values.some(function (row) {
    return creationNormalizedValue_(row[recordIndex]) ===
      creationNormalizedValue_(
        SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS
      );
  });

  if (!hasInitialStatus || !hasActiveStatus) {
    const missing = [];
    if (!hasInitialStatus) {
      missing.push(
        '«' + SYSTEM_CONFIG.VALUES.INITIAL_DOCUMENT_STATUS +
        '» в поле «' + H.DOCUMENT_STATUS + '»'
      );
    }
    if (!hasActiveStatus) {
      missing.push(
        '«' + SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS +
        '» в поле «' + H.RECORD_STATUS + '»'
      );
    }
    throw new Error(
      'На листе «' + context.config.name +
      '» отсутствуют обязательные системные значения: ' +
      missing.join(', ') + '.'
    );
  }
}


/** Читает и строго проверяет автоматические правила. */
function readAutomaticDocumentRules_() {
  const context = getSystemSheetContext_('DOCUMENT_TYPES');
  const rows = readCreationSheetValues_(context);
  const idIndex = creationColumnIndex_(context, H.DOCUMENT_TYPE_ID);
  const nameIndex = creationColumnIndex_(context, H.DOCUMENT_TYPE);
  const flagIndex = creationColumnIndex_(
    context,
    H.CREATE_ON_OBJECT_CREATION
  );
  const rules = [];
  const ids = {};
  const names = {};
  const errors = [];

  rows.forEach(function (row, offset) {
    if (
      creationNormalizedValue_(row[flagIndex]) !==
      creationNormalizedValue_(SYSTEM_CONFIG.VALUES.AUTO_CREATE_DOCUMENTS)
    ) {
      return;
    }

    const sheetRow = context.config.dataStartRow + offset;
    const id = String(row[idIndex] == null ? '' : row[idIndex]).trim();
    const name = String(row[nameIndex] == null ? '' : row[nameIndex]).trim();
    const normalizedName = creationNormalizedValue_(name);

    if (!id) {
      errors.push(
        'лист «' + context.config.name + '», строка ' + sheetRow +
        ': не заполнено поле «' + H.DOCUMENT_TYPE_ID + '»'
      );
    }
    if (!name) {
      errors.push(
        'лист «' + context.config.name + '», строка ' + sheetRow +
        ': не заполнено поле «' + H.DOCUMENT_TYPE + '»'
      );
    }
    if (!id || !name) {
      return;
    }
    if (ids[id]) {
      errors.push(
        'ID типа документа «' + id + '» повторяется в строках ' +
        ids[id] + ' и ' + sheetRow
      );
    } else {
      ids[id] = sheetRow;
    }
    if (names[normalizedName] && names[normalizedName].id !== id) {
      errors.push(
        'тип документа «' + name + '» связан с разными ID в строках ' +
        names[normalizedName].row + ' и ' + sheetRow
      );
    } else if (!names[normalizedName]) {
      names[normalizedName] = { id: id, row: sheetRow };
    }
    rules.push({ id: id, name: name, sheetRow: sheetRow });
  });

  if (rules.length === 0 && errors.length === 0) {
    errors.push('не найдено ни одного корректного автоматического правила');
  }
  if (errors.length > 0) {
    throw new Error('Ошибки справочника документов:\n- ' + errors.join('\n- '));
  }
  return rules;
}


/** Читает объекты, пропуская некорректные и все строки с повторным ID. */
function readAndValidateCreationObjects_() {
  const context = getSystemSheetContext_('OBJECTS');
  const rows = readCreationSheetValues_(context);
  const requiredHeaders = [
    H.OBJECT_ID,
    H.OBJECT_NAME,
    H.CONTRACT_NUMBER,
    H.RESPONSIBLE_FOREMAN_ID,
    H.RESPONSIBLE_FOREMAN,
    H.OBJECT_STATUS,
    H.WORK_START_DATE,
    H.WORK_END_PLAN
  ];
  const indexes = {};
  context.config.requiredHeaders.forEach(function (header) {
    indexes[header] = creationColumnIndex_(context, header);
  });
  const candidates = [];
  const idRows = {};

  rows.forEach(function (row, offset) {
    if (creationRowIsEmpty_(row)) {
      return;
    }
    const sheetRow = context.config.dataStartRow + offset;
    const id = String(row[indexes[H.OBJECT_ID]] == null
      ? '' : row[indexes[H.OBJECT_ID]])
      .replace(/\u00A0/g, ' ')
      .trim();
    candidates.push({ row: row, sheetRow: sheetRow, id: id });
    if (id) {
      if (!idRows[id]) {
        idRows[id] = [];
      }
      idRows[id].push(sheetRow);
    }
  });

  const duplicateIds = {};
  Object.keys(idRows).forEach(function (id) {
    if (idRows[id].length > 1) {
      duplicateIds[id] = idRows[id];
    }
  });

  const validObjects = [];
  const skippedObjects = [];
  candidates.forEach(function (candidate) {
    const reasons = [];
    requiredHeaders.forEach(function (header) {
      if (header === H.OBJECT_ID) {
        return;
      }
      if (creationValueIsEmpty_(candidate.row[indexes[header]])) {
        reasons.push('не заполнено поле «' + header + '»');
      }
    });
    if (!candidate.id) {
      reasons.push('не заполнено поле «' + H.OBJECT_ID + '»');
    }
    if (candidate.id && duplicateIds[candidate.id]) {
      reasons.push(
        'ID объекта повторяется в строках ' +
        duplicateIds[candidate.id].join(', ')
      );
    }
    const objectStatus = candidate.row[indexes[H.OBJECT_STATUS]];
    if (!creationValueIsEmpty_(objectStatus) && !SYSTEM_CONFIG.VALUES.OBJECT_STATUSES.some(function (status) {
      return creationNormalizedValue_(status) === creationNormalizedValue_(objectStatus);
    })) {
      reasons.push('поле «' + H.OBJECT_STATUS + '» содержит недопустимое значение «' + objectStatus + '»');
    }
    const start = candidate.row[indexes[H.WORK_START_DATE]];
    const plan = candidate.row[indexes[H.WORK_END_PLAN]];
    const fact = candidate.row[indexes[H.WORK_END_FACT]];
    if (!creationValueIsEmpty_(start) && !creationIsValidDate_(start)) {
      reasons.push('поле «' + H.WORK_START_DATE + '» не является датой');
    }
    if (!creationValueIsEmpty_(plan) && !creationIsValidDate_(plan)) {
      reasons.push('поле «' + H.WORK_END_PLAN + '» не является датой');
    }
    if (!creationValueIsEmpty_(fact) && !creationIsValidDate_(fact)) {
      reasons.push('поле «' + H.WORK_END_FACT + '» не является датой');
    }
    if (
      creationIsValidDate_(start) && creationIsValidDate_(plan) &&
      plan.getTime() < start.getTime()
    ) {
      reasons.push('плановое окончание раньше даты начала работ');
    }
    if (reasons.length > 0) {
      skippedObjects.push({
        sheetRow: candidate.sheetRow,
        id: candidate.id || '(ID не указан)',
        reasons: reasons
      });
      return;
    }
    validObjects.push({
      id: candidate.id,
      name: candidate.row[indexes[H.OBJECT_NAME]],
      contractNumber: candidate.row[indexes[H.CONTRACT_NUMBER]],
      responsibleForemanId:
        candidate.row[indexes[H.RESPONSIBLE_FOREMAN_ID]],
      responsibleForeman:
        candidate.row[indexes[H.RESPONSIBLE_FOREMAN]],
      objectStatus: candidate.row[indexes[H.OBJECT_STATUS]],
      workStartDate: start,
      workEndPlan: plan,
      workEndFact: fact,
      sheetRow: candidate.sheetRow
    });
  });

  return {
    checkedObjects: candidates.length,
    validObjects: validObjects,
    skippedObjects: skippedObjects,
    warnings: []
  };
}


/** Анализирует ключи и ID уже существующих строк фактов. */
function readExistingCreationFacts_() {
  const context = getSystemSheetContext_('DOCUMENTS');
  const rows = readCreationSheetValues_(context);
  const objectIndex = creationColumnIndex_(context, H.OBJECT_ID);
  const typeIndex = creationColumnIndex_(context, H.DOCUMENT_TYPE_ID);
  const documentIndex = creationColumnIndex_(context, H.DOCUMENT_ID);
  const keys = {};
  const documentIds = {};
  const incompleteFactRows = [];

  rows.forEach(function (row, offset) {
    if (creationRowIsEmpty_(row)) {
      return;
    }
    const sheetRow = context.config.dataStartRow + offset;
    const objectId = String(row[objectIndex] == null ? '' : row[objectIndex]).trim();
    const typeId = String(row[typeIndex] == null ? '' : row[typeIndex]).trim();
    const documentId = String(
      row[documentIndex] == null ? '' : row[documentIndex]
    ).trim();
    if (!objectId || !typeId) {
      incompleteFactRows.push(sheetRow);
    } else {
      const key = creationCompositeKey_(objectId, typeId);
      if (!keys[key]) {
        keys[key] = [];
      }
      keys[key].push(sheetRow);
    }
    if (documentId) {
      if (!documentIds[documentId]) {
        documentIds[documentId] = [];
      }
      documentIds[documentId].push(sheetRow);
    }
  });

  const duplicateObjectTypeKeys = Object.keys(keys).filter(function (key) {
    return keys[key].length > 1;
  });
  const duplicateDocumentIds = Object.keys(documentIds).filter(function (id) {
    return documentIds[id].length > 1;
  });
  const warnings = [];
  duplicateObjectTypeKeys.forEach(function (key) {
    warnings.push(
      'Дубль сочетания объект/тип в строках ' + keys[key].join(', ') + '.'
    );
  });
  duplicateDocumentIds.forEach(function (id) {
    warnings.push(
      'ID документа «' + id + '» повторяется в строках ' +
      documentIds[id].join(', ') + '.'
    );
  });
  incompleteFactRows.forEach(function (row) {
    warnings.push('Неполный ключ в строке фактов ' + row + '.');
  });
  return {
    rows: rows,
    keys: keys,
    documentIds: documentIds,
    duplicateObjectTypeKeys: duplicateObjectTypeKeys,
    duplicateDocumentIds: duplicateDocumentIds,
    incompleteFactRows: incompleteFactRows,
    warnings: warnings
  };
}


/** Строит список отсутствующих сочетаний, сохраняя порядок объектов и правил. */
function buildMissingDocumentsPlan_(objects, rules, facts) {
  const missing = [];
  let existingDocuments = 0;
  objects.forEach(function (object) {
    rules.forEach(function (rule) {
      const key = creationCompositeKey_(object.id, rule.id);
      if (facts.keys[key]) {
        existingDocuments += 1;
      } else {
        missing.push({ object: object, rule: rule });
      }
    });
  });
  return { missing: missing, existingDocuments: existingDocuments };
}


/** Готовит пакет фактов и связанную историю изменений. */
function prepareDocumentRows_(
  missing,
  facts,
  context,
  firstFactRow,
  now,
  userEmail,
  operationId
) {
  const documentRows = [];
  const changeRows = [];
  const nextNumbers = {};
  const usedIds = {};
  Object.keys(facts.documentIds).forEach(function (id) {
    usedIds[id] = true;
  });

  missing.forEach(function (item, offset) {
    if (!nextNumbers[item.object.id]) {
      nextNumbers[item.object.id] =
        findNextCreationDocumentNumber_(item.object.id, facts.documentIds);
    }
    let documentId;
    do {
      if (nextNumbers[item.object.id] > 9999) {
        throw new Error(
          'Для объекта «' + item.object.id +
          '» исчерпан диапазон четырёхзначных номеров документов.'
        );
      }
      documentId = SYSTEM_CONFIG.ID_PREFIXES.DOCUMENT + item.object.id + '-' +
        creationPadFour_(nextNumbers[item.object.id]);
      nextNumbers[item.object.id] += 1;
    } while (usedIds[documentId]);
    usedIds[documentId] = true;

    const values = {};
    values[H.DOCUMENT_ID] = documentId;
    values[H.OBJECT_ID] = item.object.id;
    values[H.OBJECT_NAME] = item.object.name;
    values[H.DOCUMENT_TYPE] = item.rule.name;
    values[H.CONTRACT_NUMBER] = item.object.contractNumber;
    values[H.DOCUMENT_NUMBER] = '';
    values[H.DOCUMENT_STATUS] =
      SYSTEM_CONFIG.VALUES.INITIAL_DOCUMENT_STATUS;
    values[H.OBJECT_STATUS] = item.object.objectStatus;
    values[H.WORK_START_DATE] = item.object.workStartDate;
    values[H.WORK_END_PLAN] = item.object.workEndPlan;
    values[H.WORK_END_FACT] = item.object.workEndFact;
    values[H.CREATED_AT] = now;
    values[H.UPDATED_AT] = now;
    values[H.DOCUMENT_TYPE_ID] = item.rule.id;
    values[H.UPDATED_BY_EMAIL] = userEmail;
    values[H.RESPONSIBLE_FOREMAN] = item.object.responsibleForeman;
    values[H.RESPONSIBLE_FOREMAN_ID] = item.object.responsibleForemanId;
    values[H.CUSTOMER_SIGNING_RESPONSIBLE] = '';
    values[H.DOCUMENT_STATUS_CHANGED_AT] = now;
    values[H.CREATION_SOURCE] =
      SYSTEM_CONFIG.VALUES.OBJECT_CREATION_SOURCE;
    values[H.RECORD_STATUS] =
      SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS;

    const row = new Array(context.headers.length).fill('');
    Object.keys(values).forEach(function (header) {
      row[creationColumnIndex_(context, header)] = values[header];
    });
    documentRows.push(row);
    appendPreparedCreationChangeRow_(
      changeRows,
      values,
      documentId,
      item.object.id,
      firstFactRow + offset,
      operationId,
      userEmail,
      now
    );
  });
  return { documentRows: documentRows, changeRows: changeRows };
}


/** Создаёт одну логическую запись истории на созданный документ. */
function appendPreparedCreationChangeRow_(
  target,
  values,
  documentId,
  objectId,
  factRow,
  operationId,
  userEmail,
  changedAt
) {
  target.push({
    changeId: generateChangeId_(operationId, target.length + 1),
    operationId: operationId,
    changedAt: changedAt,
    userEmail: userEmail,
    documentId: documentId,
    objectId: objectId,
    factRow: factRow,
    fieldName: 'Создание документа',
    newValue: buildCreationInitialSnapshot_(values)
  });
}


/** Формирует компактный многострочный снимок непустых начальных полей. */
function buildCreationInitialSnapshot_(values) {
  const dateHeaders = {};
  dateHeaders[H.WORK_START_DATE] = true;
  dateHeaders[H.WORK_END_PLAN] = true;
  dateHeaders[H.WORK_END_FACT] = true;
  const snapshotHeaders = [
    H.DOCUMENT_ID,
    H.OBJECT_ID,
    H.OBJECT_NAME,
    H.DOCUMENT_TYPE,
    H.DOCUMENT_TYPE_ID,
    H.CONTRACT_NUMBER,
    H.DOCUMENT_NUMBER,
    H.DOCUMENT_STATUS,
    H.OBJECT_STATUS,
    H.WORK_START_DATE,
    H.WORK_END_PLAN,
    H.WORK_END_FACT,
    H.RESPONSIBLE_FOREMAN,
    H.RESPONSIBLE_FOREMAN_ID,
    H.CREATION_SOURCE,
    H.RECORD_STATUS
  ];
  return snapshotHeaders.reduce(function (lines, header) {
    const value = values[header];
    if (!creationValueIsEmpty_(value)) {
      lines.push(
        header + ': ' + formatCreationHistoryValue_(
          value,
          dateHeaders[header] ? 'date' : 'text'
        )
      );
    }
    return lines;
  }, []).join('\n');
}

/** Пакетно записывает историю изменений. */
function writeCreationChangeHistory_(changes) {
  if (changes.length === 0) {
    return;
  }
  const context = getSystemSheetContext_('CHANGE_HISTORY');
  const rows = changes.map(function (change) {
    const values = {};
    values[H.CHANGE_ID] = change.changeId;
    values[H.OPERATION_ID] = change.operationId;
    values[H.CHANGE_DATETIME] = change.changedAt;
    values[H.CHANGED_BY_EMAIL] = change.userEmail;
    values[H.ACTION_TYPE] = SYSTEM_CONFIG.VALUES.CHANGE_ACTION_CREATE;
    values[H.DOCUMENT_ID] = change.documentId;
    values[H.OBJECT_ID] = change.objectId;
    values[H.FACT_ROW_NUMBER] = change.factRow;
    values[H.FIELD_NAME] = change.fieldName;
    values[H.OLD_VALUE] = '';
    values[H.NEW_VALUE] = change.newValue;
    values[H.CHANGE_SOURCE] =
      SYSTEM_CONFIG.VALUES.OBJECT_CREATION_SOURCE;
    return creationRowFromValues_(context, values);
  });
  const startRow = Math.max(
    context.sheet.getLastRow() + 1,
    context.config.dataStartRow
  );
  context.sheet.getRange(
    startRow,
    1,
    rows.length,
    context.headers.length
  ).setValues(rows);
}


/** Записывает одну строку истории операции. */
function writeCreationOperationHistory_(operation) {
  const context = getSystemSheetContext_('OPERATION_HISTORY');
  const values = {};
  values[H.OPERATION_ID] = operation.operationId;
  values[H.OPERATION_STARTED_AT] = operation.startedAt;
  values[H.OPERATION_FINISHED_AT] = operation.finishedAt;
  values[H.OPERATION_STARTED_BY] = operation.userEmail;
  values[H.OPERATION_SOURCE] =
    SYSTEM_CONFIG.VALUES.OBJECT_CREATION_SOURCE;
  values[H.OPERATION_TYPE] =
    SYSTEM_CONFIG.VALUES.OBJECT_DOCUMENT_CREATION_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = operation.status;
  values[H.DOCUMENTS_LOADED] = 0;
  values[H.DOCUMENTS_CHANGED] = operation.counters.createdDocuments;
  values[H.FACT_ROWS_UPDATED] = operation.counters.createdDocuments;
  values[H.FIELDS_CHANGED] = operation.counters.changeRows;
  values[H.DUPLICATE_IDS_FOUND] =
    operation.counters.duplicateDocumentIds;
  values[H.ERRORS_COUNT] = operation.errorsCount;
  values[H.EXECUTION_SECONDS] = Math.max(
    0,
    (operation.finishedAt.getTime() - operation.startedAt.getTime()) / 1000
  );
  values[H.ERROR_TEXT] = operation.comment;
  const row = creationRowFromValues_(context, values);
  const startRow = Math.max(
    context.sheet.getLastRow() + 1,
    context.config.dataStartRow
  );
  context.sheet.getRange(
    startRow,
    1,
    1,
    context.headers.length
  ).setValues([row]);
}


/** Формирует исход критической ошибки, сохраняя известные счётчики. */
function buildCreationFailure_(error, operationId, startedAt, userEmail) {
  const state = error.creationState || {};
  const counters = state.counters || {
    checkedObjects: 0,
    validObjects: 0,
    skippedObjects: 0,
    creationRules: 0,
    existingDocuments: 0,
    createdDocuments: 0,
    duplicateObjectTypeKeys: 0,
    duplicateDocumentIds: 0,
    incompleteFactKeys: 0,
    changeRows: 0
  };
  const finishedAt = new Date();
  const message = state.criticalMessage || String(error.message || error);
  const report = buildCreationReport_(
    SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
    counters,
    state.skippedObjects || [],
    state.warnings || [],
    message,
    startedAt,
    finishedAt
  );
  return {
    report: report,
    operationRow: {
      operationId: operationId,
      startedAt: startedAt,
      finishedAt: finishedAt,
      userEmail: userEmail,
      status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
      counters: counters,
      comment: report.text,
      errorsCount: 1
    }
  };
}


/** Собирает единый текст для истории и пользовательского отчёта. */
function buildCreationReport_(
  status,
  counters,
  skippedObjects,
  warnings,
  criticalError,
  startedAt,
  finishedAt
) {
  const lines = [
    'Статус: ' + status,
    'Проверено объектов: ' + counters.checkedObjects,
    'Корректных объектов: ' + counters.validObjects,
    'Пропущено объектов: ' + counters.skippedObjects,
    'Правил создания: ' + counters.creationRules,
    'Документов уже существовало: ' + counters.existingDocuments,
    'Документов создано: ' + counters.createdDocuments,
    'Дублей объект/тип: ' + counters.duplicateObjectTypeKeys,
    'Дублирующихся ID документа: ' + counters.duplicateDocumentIds,
    'Строк с неполным ключом: ' + counters.incompleteFactKeys,
    'Время выполнения, сек.: ' + Math.max(
      0,
      ((finishedAt.getTime() - startedAt.getTime()) / 1000).toFixed(3)
    )
  ];
  if (skippedObjects.length > 0) {
    lines.push('Причины пропуска объектов:');
    skippedObjects.forEach(function (item) {
      lines.push(
        '- строка ' + item.sheetRow + ', ID ' + item.id + ': ' +
        item.reasons.join('; ')
      );
    });
  }
  if (warnings.length > 0) {
    lines.push('Предупреждения:');
    warnings.forEach(function (warning) {
      lines.push('- ' + warning);
    });
  }
  if (criticalError) {
    lines.push('');
    lines.push('КРИТИЧЕСКАЯ ОШИБКА');
    lines.push(criticalError);
  }
  return { status: status, counters: counters, text: lines.join('\n') };
}


/** Показывает результат кнопочного запуска. */
function showCreateDocumentsReport_(report, isError) {
  try {
    if (isError) {
      const marker = 'КРИТИЧЕСКАЯ ОШИБКА\n';
      const position = String(report.text).lastIndexOf(marker);
      showCriticalOperationError_(
        'Ошибка создания документов',
        position < 0 ? report.text : report.text.slice(0, position).trim(),
        creationCriticalMessage_(report.text)
      );
      return;
    }
    SpreadsheetApp.getUi().alert(
      isError ? 'Ошибка создания документов' : 'Создание документов',
      report.text,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (uiError) {
    // Отчёт всё равно возвращается вызывающему коду и пишется в историю.
  }
}


function creationCriticalMessage_(text) {
  const marker = 'КРИТИЧЕСКАЯ ОШИБКА\n';
  const position = String(text).lastIndexOf(marker);
  return position < 0 ? String(text) : String(text).slice(position + marker.length);
}


function readCreationSheetValues_(context) {
  const lastRow = context.sheet.getLastRow();
  if (lastRow < context.config.dataStartRow) {
    return [];
  }
  return context.sheet.getRange(
    context.config.dataStartRow,
    1,
    lastRow - context.config.dataStartRow + 1,
    context.headers.length
  ).getValues();
}


function creationColumnIndex_(context, header) {
  return context.headerMap[sysNormalizeHeader_(header)] - 1;
}


function creationRowFromValues_(context, values) {
  const row = new Array(context.headers.length).fill('');
  Object.keys(values).forEach(function (header) {
    row[creationColumnIndex_(context, header)] = values[header];
  });
  return row;
}


function creationNormalizedValue_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00A0/g, ' ')
    .trim()
    .replace(/\s+/g, ' ')
    .toLocaleLowerCase();
}


function creationValueIsEmpty_(value) {
  if (value === null || typeof value === 'undefined') {
    return true;
  }

  if (typeof value === 'string') {
    return value.replace(/\u00A0/g, ' ').trim() === '';
  }

  return false;
}


function creationRowIsEmpty_(row) {
  return row.every(creationValueIsEmpty_);
}


function creationIsValidDate_(value) {
  return value instanceof Date && !isNaN(value.getTime());
}


function creationCompositeKey_(objectId, typeId) {
  return String(objectId) + '\u0000' + String(typeId);
}


function findNextCreationDocumentNumber_(objectId, documentIds) {
  const prefix = SYSTEM_CONFIG.ID_PREFIXES.DOCUMENT + objectId + '-';
  let maximum = 0;
  Object.keys(documentIds).forEach(function (id) {
    if (id.indexOf(prefix) !== 0) {
      return;
    }
    const suffix = id.slice(prefix.length);
    if (/^\d{4,}$/.test(suffix)) {
      maximum = Math.max(maximum, Number(suffix));
    }
  });
  return maximum + 1;
}


function creationPadFour_(number) {
  return String(number).padStart(4, '0');
}


function formatCreationHistoryValue_(value, kind) {
  if (kind === 'date' || kind === 'datetime') {
    const timezone = getSystemSpreadsheet_().getSpreadsheetTimeZone();
    return Utilities.formatDate(
      value,
      timezone,
      kind === 'date' ? 'dd.MM.yyyy' : 'dd.MM.yyyy HH:mm:ss'
    );
  }
  return String(value);
}
