/** Web App для создания одного документа. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('DocumentWebApp')
    .setTitle('Документы — Осушение')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function webAppGetInitialData(sessionToken) {
  const user = webAppRequireSession_(sessionToken);
  assertSystemSheetsStructure_([
    'OBJECTS', 'DOCUMENT_TYPES', 'CARD_DICTIONARY', 'DOCUMENTS'
  ]);
  const objectResult = readAndValidateCreationObjects_();
  const types = webAppReadDocumentTypes_().filter(function (type) {
    return webAppValidRepeatability_(type.repeatability);
  });
  return {
    user: user,
    objects: objectResult.validObjects
      .filter(function (object) {
        return object.id !== SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL;
      })
      .map(function (object) { return { id: object.id, name: object.name }; }),
    documentTypes: types.map(function (type) {
      return { id: type.id, name: type.name, repeatability: type.repeatability };
    }),
    statuses: webAppReadStatuses_()
  };
}

function webAppGetExistingDocuments(sessionToken, objectId, documentTypeId) {
  webAppRequireSession_(sessionToken);
  assertSystemSheetsStructure_(['OBJECTS', 'DOCUMENT_TYPES', 'DOCUMENTS']);
  const object = webAppFindObject_(objectId);
  const type = webAppFindDocumentType_(documentTypeId);
  const existing = webAppReadMatchingDocuments_(object.id, type.id, type.name);
  const blocked = type.repeatability === SYSTEM_CONFIG.VALUES.DOCUMENT_REPEATABILITY_ONE && existing.length > 0;
  const nextNumber = type.repeatability === SYSTEM_CONFIG.VALUES.DOCUMENT_REPEATABILITY_MANY
    ? webAppNextDocumentNumber_(existing, type.name) : '';
  return {
    objectName: object.name,
    documentType: { id: type.id, name: type.name, repeatability: type.repeatability },
    documents: existing,
    allowed: !blocked,
    nextDocumentName: type.repeatability === SYSTEM_CONFIG.VALUES.DOCUMENT_REPEATABILITY_MANY
      ? webAppDisplayDocumentName_(type.name, nextNumber) : '',
    warning: blocked
      ? 'Документ этого типа уже существует.\n' +
        'Повторяемость типа документа — «Один», поэтому создать второй документ нельзя.' : ''
  };
}

function webAppCreateDocument(sessionToken, payload) {
  webAppRequireSession_(sessionToken);
  const startedAt = new Date();
  return withDocumentLock_(function () {
    const user = webAppRequireSession_(sessionToken);
    let operationId = '';
    let factWritten = false;
    let changeWritten = false;
    try {
      assertSystemSheetsStructure_([
        'OBJECTS', 'DOCUMENT_TYPES', 'CARD_DICTIONARY', 'DOCUMENTS',
        'CHANGE_HISTORY', 'OPERATION_HISTORY'
      ]);
      operationId = generateOperationId_(startedAt);
      const clean = webAppValidatePayload_(payload);
      const object = webAppFindObject_(clean.objectId);
      const type = webAppFindDocumentType_(clean.documentTypeId);
      const statuses = webAppReadStatuses_();
      if (statuses.indexOf(clean.documentStatus) < 0) {
        throw new Error('Выбран несуществующий статус документа.');
      }
      // Окончательный номер всегда рассчитывается после повторного
      // чтения фактов внутри общей блокировки.
      const existing = webAppReadMatchingDocuments_(object.id, type.id, type.name);
      if (type.repeatability === SYSTEM_CONFIG.VALUES.DOCUMENT_REPEATABILITY_ONE && existing.length > 0) {
        webAppWriteOperation_({
          operationId: operationId, startedAt: startedAt, finishedAt: new Date(),
          actor: user.actor, status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES,
          changed: 0, errors: 0,
          comment: 'Документ этого типа уже существует. Повторяемость — «Один».'
        });
        return {
          ok: false, blocked: true,
          message: 'Документ не создан.\n\nДокумент этого типа уже существует.\nПовторяемость — «Один».'
        };
      }

      const fact = readExistingCreationFacts_();
      const documentsContext = getSystemSheetContext_('DOCUMENTS');
      const factRow = Math.max(
        documentsContext.sheet.getLastRow() + 1,
        documentsContext.config.dataStartRow
      );
      const now = new Date();
      const documentNumber = type.repeatability === SYSTEM_CONFIG.VALUES.DOCUMENT_REPEATABILITY_MANY
        ? webAppNextDocumentNumber_(existing, type.name) : '';
      if (documentNumber !== '' && webAppDocumentNumberOccupied_(
        existing, type.name, documentNumber
      )) {
        throw new Error('Рассчитанный номер документа уже занят. Обновите данные.');
      }
      const prepared = webAppPrepareDocument_(
        object, type, type.name, documentNumber, clean.documentStatus,
        fact, documentsContext,
        factRow, now, user.actor, operationId
      );
      documentsContext.sheet.getRange(
        factRow, 1, 1, documentsContext.headers.length
      ).setValues([prepared.row]);
      factWritten = true;
      webAppWriteChange_(prepared.change);
      changeWritten = true;
      webAppWriteOperation_({
        operationId: operationId, startedAt: startedAt, finishedAt: new Date(),
        actor: user.actor, status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS,
        changed: 1, errors: 0, comment: 'Документ создан: ' + prepared.documentId
      });
      return {
        ok: true, objectName: object.name,
        documentType: type.name, documentNumber: documentNumber,
        documentName: webAppDisplayDocumentName_(type.name, documentNumber),
        documentStatus: clean.documentStatus, documentId: prepared.documentId
      };
    } catch (error) {
      if (factWritten) {
        if (!changeWritten && operationId) {
          try {
            webAppWriteOperation_({
              operationId: operationId, startedAt: startedAt, finishedAt: new Date(),
              actor: user.actor, status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
              changed: 1, errors: 1,
              comment: 'Факт записан, но история изменений не записана.'
            });
          } catch (ignoredHistory) { /* Автоповтор записи факта запрещён. */ }
        }
        throw new Error(
          'Документ мог быть создан, но завершение операции прошло с ошибкой. ' +
          'Обновите данные перед повторной попыткой.'
        );
      }
      if (operationId) {
        try {
          webAppWriteOperation_({
            operationId: operationId, startedAt: startedAt, finishedAt: new Date(),
            actor: user.actor, status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
            changed: 0, errors: 1, comment: String(error.message || error)
          });
        } catch (ignored) { /* Исходная ошибка имеет приоритет. */ }
      }
      throw new Error(String(error.message || error));
    }
  });
}

function webAppReadDocumentTypes_() {
  const context = getSystemSheetContext_('DOCUMENT_TYPES');
  const rows = webAppReadRows_(context);
  const indexes = webAppIndexes_(context, [
    H.DOCUMENT_TYPE_ID, H.DOCUMENT_TYPE, H.REPEATABILITY
  ]);
  const result = [];
  const ids = {};
  rows.forEach(function (row, offset) {
    if (webAppRowEmpty_(row)) return;
    const id = String(row[indexes[H.DOCUMENT_TYPE_ID]] == null
      ? '' : row[indexes[H.DOCUMENT_TYPE_ID]]).trim();
    const name = String(row[indexes[H.DOCUMENT_TYPE]] == null
      ? '' : row[indexes[H.DOCUMENT_TYPE]]).trim();
    const repeatability = String(row[indexes[H.REPEATABILITY]] == null
      ? '' : row[indexes[H.REPEATABILITY]]).trim();
    if (!id || !name) return;
    if (ids[id]) throw new Error('В справочнике документов повторяется ID типа документа «' + id + '».');
    ids[id] = context.config.dataStartRow + offset;
    result.push({ id: id, name: name, repeatability: repeatability });
  });
  return result;
}

function webAppFindDocumentType_(id) {
  const cleanId = String(id == null ? '' : id).trim();
  const matches = webAppReadDocumentTypes_().filter(function (type) {
    return type.id === cleanId;
  });
  if (matches.length !== 1) throw new Error('Тип документа не найден.');
  if (!webAppValidRepeatability_(matches[0].repeatability)) {
    throw new Error('Для этого типа документа некорректно настроена повторяемость. Обратитесь к администратору.');
  }
  return matches[0];
}

function webAppValidRepeatability_(value) {
  return value === SYSTEM_CONFIG.VALUES.DOCUMENT_REPEATABILITY_ONE ||
    value === SYSTEM_CONFIG.VALUES.DOCUMENT_REPEATABILITY_MANY;
}

function webAppFindObject_(id) {
  const cleanId = String(id == null ? '' : id).trim();
  if (!cleanId || cleanId === SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL) {
    throw new Error('Выберите конкретный объект.');
  }
  const result = readAndValidateCreationObjects_();
  const matches = result.validObjects.filter(function (object) { return object.id === cleanId; });
  if (matches.length !== 1) {
    const duplicate = result.skippedObjects.some(function (object) {
      return object.id === cleanId && object.reasons.some(function (reason) {
        return reason.indexOf('ID объекта повторяется') >= 0;
      });
    });
    throw new Error(duplicate ? 'В справочнике повторяется ID выбранного объекта.' : 'Объект не найден или некорректно настроен.');
  }
  return matches[0];
}

function webAppReadStatuses_() {
  const context = getSystemSheetContext_('CARD_DICTIONARY');
  const index = webAppIndexes_(context, [H.DOCUMENT_STATUS])[H.DOCUMENT_STATUS];
  const seen = {};
  const statuses = webAppReadRows_(context).reduce(function (result, row) {
    const status = String(row[index] == null ? '' : row[index]).trim();
    if (status && !seen[status]) { seen[status] = true; result.push(status); }
    return result;
  }, []);
  if (statuses.length === 0) {
    throw new Error('В справочнике для КО не настроены статусы документа.');
  }
  return statuses;
}

function webAppReadMatchingDocuments_(objectId, typeId, canonicalDocumentType) {
  const context = getSystemSheetContext_('DOCUMENTS');
  const fields = [H.OBJECT_ID, H.DOCUMENT_TYPE_ID, H.DOCUMENT_TYPE,
    H.DOCUMENT_NUMBER, H.DOCUMENT_ID, H.DOCUMENT_STATUS, H.RECORD_STATUS,
    H.CREATED_AT];
  const indexes = webAppIndexes_(context, fields);
  return webAppReadRows_(context).reduce(function (items, row) {
    if (String(row[indexes[H.OBJECT_ID]]).trim() !== objectId ||
        String(row[indexes[H.DOCUMENT_TYPE_ID]]).trim() !== typeId) return items;
    const storedType = String(row[indexes[H.DOCUMENT_TYPE]] == null
      ? '' : row[indexes[H.DOCUMENT_TYPE]]).trim();
    const canonicalType = canonicalDocumentType || storedType;
    const explicitNumber = row[indexes[H.DOCUMENT_NUMBER]];
    const effectiveNumber = webAppEffectiveDocumentNumber_(
      explicitNumber, storedType, canonicalType
    );
    items.push({
      documentType: canonicalType,
      storedDocumentType: storedType,
      documentNumber: webAppClientValue_(explicitNumber),
      displayName: webAppDisplayDocumentName_(canonicalType, effectiveNumber),
      documentId: row[indexes[H.DOCUMENT_ID]],
      documentStatus: row[indexes[H.DOCUMENT_STATUS]], recordStatus: row[indexes[H.RECORD_STATUS]],
      createdAt: webAppClientValue_(row[indexes[H.CREATED_AT]])
    });
    return items;
  }, []);
}

function webAppPrepareDocument_(object, type, canonicalDocumentType,
  documentNumber, status, facts, context, factRow, now, actor, operationId) {
  const number = findNextCreationDocumentNumber_(object.id, facts.documentIds);
  if (number > 9999) throw new Error('Невозможно безопасно сгенерировать ID документа.');
  const documentId = SYSTEM_CONFIG.ID_PREFIXES.DOCUMENT + object.id + '-' + creationPadFour_(number);
  if (facts.documentIds[documentId]) throw new Error('Невозможно безопасно сгенерировать ID документа.');
  const values = {};
  values[H.DOCUMENT_ID] = documentId; values[H.OBJECT_ID] = object.id;
  values[H.OBJECT_NAME] = object.name;
  values[H.DOCUMENT_TYPE] = canonicalDocumentType;
  values[H.DOCUMENT_NUMBER] = documentNumber;
  values[H.CONTRACT_NUMBER] = object.contractNumber; values[H.DOCUMENT_STATUS] = status;
  values[H.OBJECT_STATUS] = object.objectStatus; values[H.WORK_START_DATE] = object.workStartDate;
  values[H.WORK_END_PLAN] = object.workEndPlan; values[H.WORK_END_FACT] = object.workEndFact;
  values[H.CREATED_AT] = now; values[H.UPDATED_AT] = now; values[H.DOCUMENT_TYPE_ID] = type.id;
  values[H.UPDATED_BY_EMAIL] = actor; values[H.RESPONSIBLE_FOREMAN] = object.responsibleForeman;
  values[H.RESPONSIBLE_FOREMAN_ID] = object.responsibleForemanId;
  values[H.DOCUMENT_STATUS_CHANGED_AT] = now;
  values[H.CREATION_SOURCE] = SYSTEM_CONFIG.VALUES.WEB_APP_SOURCE;
  values[H.RECORD_STATUS] = SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS;
  return {
    documentId: documentId, row: creationRowFromValues_(context, values),
    change: { changeId: generateChangeId_(operationId, 1), operationId: operationId,
      changedAt: now, actor: actor, documentId: documentId, objectId: object.id,
      factRow: factRow, snapshot: buildCreationInitialSnapshot_(values) }
  };
}

/** Возвращает только положительный целый номер. */
function webAppPositiveDocumentNumber_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!/^\d+$/.test(text)) return null;
  const number = Number(text);
  return Number.isSafeInteger(number) && number > 0 ? number : null;
}

/** Отдельное поле имеет приоритет над legacy-суффиксом. */
function webAppEffectiveDocumentNumber_(explicitNumber, storedType, canonicalType) {
  if (String(explicitNumber == null ? '' : explicitNumber).trim() !== '') {
    return webAppPositiveDocumentNumber_(explicitNumber) ||
      webAppPositiveDocumentNumber_(
        documentArchitectureExtractLegacyNumber_(
          explicitNumber, canonicalType, ''
        )
      );
  }
  return webAppPositiveDocumentNumber_(
    documentArchitectureExtractLegacyNumber_(storedType, canonicalType, '')
  );
}

function webAppNextDocumentNumber_(documents, canonicalType) {
  const maximum = documents.reduce(function (current, document) {
    const number = webAppEffectiveDocumentNumber_(
      document.documentNumber, document.storedDocumentType || document.documentType,
      canonicalType
    );
    return number == null ? current : Math.max(current, number);
  }, 0);
  if (maximum >= Number.MAX_SAFE_INTEGER) {
    throw new Error('Исчерпан диапазон номеров документа.');
  }
  return maximum + 1;
}

function webAppDocumentNumberOccupied_(documents, canonicalType, number) {
  return documents.some(function (document) {
    return webAppEffectiveDocumentNumber_(
      document.documentNumber, document.storedDocumentType || document.documentType,
      canonicalType
    ) === number;
  });
}

function webAppDisplayDocumentName_(canonicalType, documentNumber) {
  return documentNumber === '' || documentNumber == null
    ? canonicalType : canonicalType + ' №' + documentNumber;
}

function webAppWriteChange_(change) {
  const context = getSystemSheetContext_('CHANGE_HISTORY');
  const values = {};
  values[H.CHANGE_ID] = change.changeId; values[H.OPERATION_ID] = change.operationId;
  values[H.CHANGE_DATETIME] = change.changedAt; values[H.CHANGED_BY_EMAIL] = change.actor;
  values[H.ACTION_TYPE] = SYSTEM_CONFIG.VALUES.CHANGE_ACTION_CREATE;
  values[H.DOCUMENT_ID] = change.documentId; values[H.OBJECT_ID] = change.objectId;
  values[H.FACT_ROW_NUMBER] = change.factRow; values[H.FIELD_NAME] = 'Создание документа';
  values[H.OLD_VALUE] = ''; values[H.NEW_VALUE] = change.snapshot;
  values[H.CHANGE_SOURCE] = SYSTEM_CONFIG.VALUES.WEB_APP_SOURCE;
  webAppWriteOneRow_(context, creationRowFromValues_(context, values));
}

function webAppWriteOperation_(operation) {
  const context = getSystemSheetContext_('OPERATION_HISTORY');
  const values = {};
  values[H.OPERATION_ID] = operation.operationId; values[H.OPERATION_STARTED_AT] = operation.startedAt;
  values[H.OPERATION_FINISHED_AT] = operation.finishedAt; values[H.OPERATION_STARTED_BY] = operation.actor;
  values[H.OPERATION_SOURCE] = SYSTEM_CONFIG.VALUES.WEB_APP_SOURCE;
  values[H.OPERATION_TYPE] = SYSTEM_CONFIG.VALUES.WEB_APP_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = operation.status; values[H.DOCUMENTS_LOADED] = 0;
  values[H.DOCUMENTS_CHANGED] = operation.changed; values[H.FACT_ROWS_UPDATED] = operation.changed;
  values[H.FIELDS_CHANGED] = operation.changed; values[H.DUPLICATE_IDS_FOUND] = 0;
  values[H.ERRORS_COUNT] = operation.errors;
  values[H.EXECUTION_SECONDS] = Math.max(0, (operation.finishedAt.getTime() - operation.startedAt.getTime()) / 1000);
  values[H.ERROR_TEXT] = operation.comment;
  webAppWriteOneRow_(context, creationRowFromValues_(context, values));
}

function webAppWriteOneRow_(context, row) {
  const target = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(target, 1, 1, context.headers.length).setValues([row]);
}
function webAppValidatePayload_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Некорректный запрос.');
  return { objectId: String(payload.objectId == null ? '' : payload.objectId).trim(),
    documentTypeId: String(payload.documentTypeId == null ? '' : payload.documentTypeId).trim(),
    documentStatus: String(payload.documentStatus == null ? '' : payload.documentStatus).trim() };
}
function webAppIndexes_(context, headers) {
  const result = {}; headers.forEach(function (header) { result[header] = creationColumnIndex_(context, header); }); return result;
}
function webAppReadRows_(context) { return readCreationSheetValues_(context); }
function webAppRowEmpty_(row) { return row.every(function (value) { return value === '' || value == null; }); }
function webAppClientValue_(value) { return value instanceof Date ? value.toISOString() : String(value == null ? '' : value); }
