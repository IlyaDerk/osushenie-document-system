/** Web App создания документов. Редактирование фактов здесь намеренно отсутствует. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('DocumentWebApp')
    .setTitle('Документы — Осушение')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function webAppGetBootstrap(sessionToken) {
  try {
    webAppRequireSession_(sessionToken);
    assertSystemSheetsStructure_([
      'OBJECTS', 'DOCUMENT_TYPES', 'CARD_DICTIONARY'
    ]);
    return {
      ok: true,
      objects: webAppReadObjects_().map(function (object) {
        return { id: object.id, name: object.name };
      }),
      documentTypes: webAppReadDocumentTypes_().map(function (type) {
        return { id: type.id, name: type.name, repeatability: type.repeatability };
      }),
      documentStatuses: webAppReadStatuses_()
    };
  } catch (error) {
    return webAppFailureResponse_(error);
  }
}

function webAppGetExistingDocuments(sessionToken, objectId, documentTypeId) {
  try {
    webAppRequireSession_(sessionToken);
    assertSystemSheetsStructure_(['OBJECTS', 'DOCUMENT_TYPES', 'DOCUMENTS']);
    const object = webAppFindUnique_(webAppReadObjects_(), objectId, 'объекта');
    webAppValidateObject_(object);
    const type = webAppFindUnique_(
      webAppReadDocumentTypes_(), documentTypeId, 'типа документа'
    );
    const documents = webAppReadMatchingDocuments_(object.id, type.id);
    return {
      ok: true,
      object: { id: object.id, name: object.name },
      documentType: { id: type.id, name: type.name, repeatability: type.repeatability },
      documents: documents,
      canCreate: type.repeatability === 'Много' || documents.length === 0,
      nextDocumentName: type.repeatability === 'Много'
        ? type.name + ' №' + (documents.length + 1) : type.name
    };
  } catch (error) {
    return webAppFailureResponse_(error);
  }
}

function webAppCreateDocument(sessionToken, payload) {
  let user;
  try {
    user = webAppRequireSession_(sessionToken);
  } catch (error) {
    return webAppFailureResponse_(error);
  }
  const request = payload || {};
  try {
    return withDocumentLock_(function () {
      const startedAt = new Date();
      const operationId = generateOperationId_(startedAt);
      try {
        return webAppCreateDocumentUnderLock_(user, request, startedAt, operationId);
      } catch (error) {
        try {
          webAppWriteOperation_({
            operationId: operationId, startedAt: startedAt, finishedAt: new Date(),
            actor: user.actor, status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
            created: error.webAppFactWritten ? 1 : 0, errors: 1,
            comment: webAppSafeError_(error)
          });
        } catch (historyError) {
          // Исходная ошибка и признак записи факта важнее best-effort журнала.
        }
        throw error;
      }
    });
  } catch (error) {
    return webAppFailureResponse_(error, !!error.webAppFactWritten);
  }
}

function webAppCreateDocumentUnderLock_(user, payload, startedAt, operationId) {
  assertSystemSheetsStructure_([
    'OBJECTS', 'DOCUMENT_TYPES', 'DOCUMENTS', 'CARD_DICTIONARY',
    'CHANGE_HISTORY', 'OPERATION_HISTORY'
  ]);
  const object = webAppFindUnique_(webAppReadObjects_(), payload.objectId, 'объекта');
  webAppValidateObject_(object);
  const type = webAppFindUnique_(
    webAppReadDocumentTypes_(), payload.documentTypeId, 'типа документа'
  );
  const status = String(payload.documentStatus == null ? '' : payload.documentStatus).trim();
  if (webAppReadStatuses_().indexOf(status) < 0) {
    throw new Error('Выбран несуществующий статус документа. Обновите данные.');
  }
  const existing = webAppReadMatchingDocuments_(object.id, type.id);
  if (type.repeatability === 'Один' && existing.length > 0) {
    webAppWriteOperation_({
      operationId: operationId, startedAt: startedAt, finishedAt: new Date(),
      actor: user.actor, status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES,
      created: 0, errors: 0,
      comment: 'Документ не создан. Повторяемость «Один»; документ уже существует. ' +
        webAppOperationSubject_(user, object, type, '', status)
    });
    return {
      ok: false, businessBlocked: true,
      message: 'Документ не создан.\n\nДокумент этого типа уже существует.\n' +
        'Повторяемость — «Один».',
      documents: existing
    };
  }
  const documentsContext = getSystemSheetContext_('DOCUMENTS');
  const facts = readExistingCreationFacts_();
  const documentId = webAppNextDocumentId_(object.id, facts.documentIds);
  const actualName = type.repeatability === 'Много'
    ? type.name + ' №' + (existing.length + 1) : type.name;
  const now = new Date();
  const values = webAppDocumentValues_(object, type, actualName, status, documentId, user.actor, now);
  const row = creationRowFromValues_(documentsContext, values);
  const factRow = Math.max(
    documentsContext.sheet.getLastRow() + 1,
    documentsContext.config.dataStartRow
  );
  documentsContext.sheet.getRange(
    factRow, 1, 1, documentsContext.headers.length
  ).setValues([row]);
  let factWritten = true;
  try {
    webAppWriteChange_(operationId, user.actor, factRow, values, now);
    webAppWriteOperation_({
      operationId: operationId, startedAt: startedAt, finishedAt: new Date(),
      actor: user.actor, status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS,
      created: 1, errors: 0,
      comment: webAppOperationSubject_(user, object, type, documentId, status)
    });
  } catch (error) {
    error.webAppFactWritten = factWritten;
    throw error;
  }
  return {
    ok: true,
    document: {
      id: documentId, objectId: object.id, objectName: object.name,
      name: actualName, status: status
    }
  };
}

function webAppReadObjects_() {
  const context = getSystemSheetContext_('OBJECTS');
  const rows = readCreationSheetValues_(context);
  const headers = [H.OBJECT_ID, H.OBJECT_NAME, H.CONTRACT_NUMBER, H.OBJECT_STATUS,
    H.WORK_START_DATE, H.WORK_END_PLAN, H.WORK_END_FACT,
    H.RESPONSIBLE_FOREMAN, H.RESPONSIBLE_FOREMAN_ID];
  const indexes = {};
  headers.forEach(function (header) { indexes[header] = creationColumnIndex_(context, header); });
  return rows.reduce(function (items, row) {
    const id = String(row[indexes[H.OBJECT_ID]] == null ? '' : row[indexes[H.OBJECT_ID]]).trim();
    const name = String(row[indexes[H.OBJECT_NAME]] == null ? '' : row[indexes[H.OBJECT_NAME]]).trim();
    if (!id || !name || webAppNormalize_(id) === webAppNormalize_(SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL) ||
        webAppNormalize_(name) === webAppNormalize_(SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL)) return items;
    items.push({ id: id, name: name, contractNumber: row[indexes[H.CONTRACT_NUMBER]],
      objectStatus: row[indexes[H.OBJECT_STATUS]], workStartDate: row[indexes[H.WORK_START_DATE]],
      workEndPlan: row[indexes[H.WORK_END_PLAN]], workEndFact: row[indexes[H.WORK_END_FACT]],
      responsibleForeman: row[indexes[H.RESPONSIBLE_FOREMAN]],
      responsibleForemanId: row[indexes[H.RESPONSIBLE_FOREMAN_ID]] });
    return items;
  }, []);
}

function webAppReadDocumentTypes_() {
  const context = getSystemSheetContext_('DOCUMENT_TYPES');
  const rows = readCreationSheetValues_(context);
  const idIndex = creationColumnIndex_(context, H.DOCUMENT_TYPE_ID);
  const nameIndex = creationColumnIndex_(context, H.DOCUMENT_TYPE);
  const repeatIndex = creationColumnIndex_(context, H.REPEATABILITY);
  return rows.reduce(function (items, row) {
    const id = String(row[idIndex] == null ? '' : row[idIndex]).trim();
    const name = String(row[nameIndex] == null ? '' : row[nameIndex]).trim();
    if (!id || !name) return items;
    const normalized = webAppNormalize_(row[repeatIndex]);
    if (normalized !== webAppNormalize_('Один') && normalized !== webAppNormalize_('Много')) {
      throw new Error('Для этого типа документа некорректно настроена повторяемость. ' +
        'Обратитесь к администратору.');
    }
    items.push({ id: id, name: name, repeatability: normalized === webAppNormalize_('Один') ? 'Один' : 'Много' });
    return items;
  }, []);
}

function webAppReadStatuses_() {
  const context = getSystemSheetContext_('CARD_DICTIONARY');
  const index = creationColumnIndex_(context, H.DOCUMENT_STATUS);
  const seen = {};
  return readCreationSheetValues_(context).reduce(function (values, row) {
    const value = String(row[index] == null ? '' : row[index]).trim();
    if (value && !seen[value]) { seen[value] = true; values.push(value); }
    return values;
  }, []);
}

function webAppFindUnique_(items, requestedId, entityName) {
  const id = String(requestedId == null ? '' : requestedId).trim();
  const matches = items.filter(function (item) { return item.id === id; });
  if (!id || matches.length === 0) throw new Error('Не найдены актуальные данны ' + entityName + '.');
  if (matches.length > 1) throw new Error('В справочнике найден дублирующийся ID ' + entityName + '.');
  return matches[0];
}

function webAppValidateObject_(object) {
  const required = [
    ['name', H.OBJECT_NAME], ['contractNumber', H.CONTRACT_NUMBER],
    ['objectStatus', H.OBJECT_STATUS], ['workStartDate', H.WORK_START_DATE],
    ['workEndPlan', H.WORK_END_PLAN],
    ['responsibleForeman', H.RESPONSIBLE_FOREMAN],
    ['responsibleForemanId', H.RESPONSIBLE_FOREMAN_ID]
  ];
  const missing = required.filter(function (item) {
    return creationValueIsEmpty_(object[item[0]]);
  }).map(function (item) { return item[1]; });
  if (missing.length) {
    throw new Error('Для объекта не заполнены обязательные данны: ' + missing.join(', ') + '.');
  }
  [H.WORK_START_DATE, H.WORK_END_PLAN, H.WORK_END_FACT].forEach(function (header) {
    const field = header === H.WORK_START_DATE ? 'workStartDate'
      : header === H.WORK_END_PLAN ? 'workEndPlan' : 'workEndFact';
    if (!creationValueIsEmpty_(object[field]) && !creationIsValidDate_(object[field])) {
      throw new Error('В данных объекта поле «' + header + '» не является датой.');
    }
  });
  if (object.workEndPlan.getTime() < object.workStartDate.getTime()) {
    throw new Error('В данных объекта плановое окончание раньше даты начала работ.');
  }
}

function webAppReadMatchingDocuments_(objectId, typeId) {
  const context = getSystemSheetContext_('DOCUMENTS');
  const indexes = {};
  [H.OBJECT_ID, H.DOCUMENT_TYPE_ID, H.DOCUMENT_ID, H.DOCUMENT_TYPE,
    H.DOCUMENT_STATUS, H.RECORD_STATUS, H.CREATED_AT].forEach(function (header) {
    indexes[header] = creationColumnIndex_(context, header);
  });
  return readCreationSheetValues_(context).reduce(function (documents, row) {
    if (String(row[indexes[H.OBJECT_ID]]).trim() !== objectId ||
        String(row[indexes[H.DOCUMENT_TYPE_ID]]).trim() !== typeId) return documents;
    documents.push({ id: String(row[indexes[H.DOCUMENT_ID]]), name: String(row[indexes[H.DOCUMENT_TYPE]]),
      status: String(row[indexes[H.DOCUMENT_STATUS]]), recordStatus: String(row[indexes[H.RECORD_STATUS]]),
      createdAt: row[indexes[H.CREATED_AT]] });
    return documents;
  }, []);
}

function webAppNextDocumentId_(objectId, documentIds) {
  let number = findNextCreationDocumentNumber_(objectId, documentIds);
  let id;
  do {
    if (number > 9999) throw new Error('Невозможно безопасно создать ID документа: диапазон исчерпан.');
    id = SYSTEM_CONFIG.ID_PREFIXES.DOCUMENT + objectId + '-' + creationPadFour_(number++);
  } while (documentIds[id]);
  return id;
}

function webAppDocumentValues_(object, type, name, status, id, actor, now) {
  const values = {};
  values[H.DOCUMENT_ID] = id; values[H.OBJECT_ID] = object.id; values[H.DOCUMENT_TYPE] = name;
  values[H.CONTRACT_NUMBER] = object.contractNumber; values[H.DOCUMENT_STATUS] = status;
  values[H.OBJECT_STATUS] = object.objectStatus; values[H.WORK_START_DATE] = object.workStartDate;
  values[H.WORK_END_PLAN] = object.workEndPlan; values[H.WORK_END_FACT] = object.workEndFact;
  values[H.CREATED_AT] = now; values[H.UPDATED_AT] = now; values[H.DOCUMENT_TYPE_ID] = type.id;
  values[H.UPDATED_BY_EMAIL] = actor; values[H.RESPONSIBLE_FOREMAN] = object.responsibleForeman;
  values[H.RESPONSIBLE_FOREMAN_ID] = object.responsibleForemanId;
  values[H.DOCUMENT_STATUS_CHANGED_AT] = now;
  values[H.CREATION_SOURCE] = SYSTEM_CONFIG.VALUES.WEB_APP_SOURCE;
  values[H.RECORD_STATUS] = SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS;
  return values;
}

function webAppWriteChange_(operationId, actor, factRow, values, now) {
  const context = getSystemSheetContext_('CHANGE_HISTORY');
  const item = {};
  item[H.CHANGE_ID] = generateChangeId_(operationId, 1); item[H.OPERATION_ID] = operationId;
  item[H.CHANGE_DATETIME] = now; item[H.CHANGED_BY_EMAIL] = actor;
  item[H.ACTION_TYPE] = SYSTEM_CONFIG.VALUES.CHANGE_ACTION_CREATE;
  item[H.DOCUMENT_ID] = values[H.DOCUMENT_ID]; item[H.OBJECT_ID] = values[H.OBJECT_ID];
  item[H.FACT_ROW_NUMBER] = factRow; item[H.FIELD_NAME] = 'Создание документа';
  item[H.OLD_VALUE] = ''; item[H.NEW_VALUE] = buildCreationInitialSnapshot_(values);
  item[H.CHANGE_SOURCE] = SYSTEM_CONFIG.VALUES.WEB_APP_SOURCE;
  const row = creationRowFromValues_(context, item);
  context.sheet.getRange(Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow),
    1, 1, context.headers.length).setValues([row]);
}

function webAppWriteOperation_(operation) {
  const context = getSystemSheetContext_('OPERATION_HISTORY');
  const values = {};
  values[H.OPERATION_ID] = operation.operationId; values[H.OPERATION_STARTED_AT] = operation.startedAt;
  values[H.OPERATION_FINISHED_AT] = operation.finishedAt; values[H.OPERATION_STARTED_BY] = operation.actor;
  values[H.OPERATION_SOURCE] = SYSTEM_CONFIG.VALUES.WEB_APP_SOURCE;
  values[H.OPERATION_TYPE] = SYSTEM_CONFIG.VALUES.WEB_APP_CREATE_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = operation.status; values[H.DOCUMENTS_LOADED] = 0;
  values[H.DOCUMENTS_CHANGED] = operation.created; values[H.FACT_ROWS_UPDATED] = operation.created;
  values[H.FIELDS_CHANGED] = operation.created; values[H.DUPLICATE_IDS_FOUND] = 0;
  values[H.ERRORS_COUNT] = operation.errors; values[H.EXECUTION_SECONDS] = Math.max(0,
    (operation.finishedAt.getTime() - operation.startedAt.getTime()) / 1000);
  values[H.ERROR_TEXT] = operation.comment;
  const row = creationRowFromValues_(context, values);
  context.sheet.getRange(Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow),
    1, 1, context.headers.length).setValues([row]);
}

function webAppOperationSubject_(user, object, type, id, status) {
  return 'Пользователь: ' + (user.fullName || user.login) + '; Объект: ' + object.name +
    '; Тип: ' + type.name + '; Созданный ID: ' + (id || '—') + '; Статус: ' + status;
}

function webAppFailureResponse_(error, factWritten) {
  const expired = String(error && error.message).indexOf('Сессия истекла') === 0;
  return { ok: false, sessionExpired: expired, uncertain: !!factWritten,
    message: factWritten ? 'Документ мог быть создан, но завершение операции прошло с ошибкой.\n' +
      'Обновите данные перед повторной попыткой.' : webAppSafeError_(error) };
}
