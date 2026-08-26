/** Read-only loader for «Карточка операциониста». */
function operatorCardNormalizeText_(value) {
  return String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim().replace(/\s+/g, ' ');
}

function operatorCardFold_(value) {
  return operatorCardNormalizeText_(value).toLocaleLowerCase('ru');
}

/** Canonical client-contact label. Missing parts are never guessed. */
function operatorCardClientDisplay_(contactName, organizationName) {
  const contact = operatorCardNormalizeText_(contactName);
  const organization = operatorCardNormalizeText_(organizationName);
  return contact && organization ? contact + ' — ' + organization : '';
}

function operatorCardIsActiveClient_(client) {
  return operatorCardFold_(client.status) === operatorCardFold_('Активный');
}

function operatorCardClientsForObject_(clients, objectName) {
  const key = operatorCardFold_(objectName);
  return (clients || []).filter(function (client) {
    return operatorCardIsActiveClient_(client) && operatorCardFold_(client.objectName) === key;
  });
}

function operatorCardReadDictionary_(sheetKey, fields) {
  const context = getSystemSheetContext_(sheetKey);
  const lastRow = context.sheet.getLastRow();
  if (lastRow < context.config.dataStartRow) return [];
  const fieldColumns = fields.map(function (field) {
    return context.headerMap[sysNormalizeHeader_(field.header)];
  });
  const firstColumn = Math.min.apply(null, fieldColumns);
  const lastColumn = Math.max.apply(null, fieldColumns);
  const values = context.sheet.getRange(
    context.config.dataStartRow, firstColumn,
    lastRow - context.config.dataStartRow + 1,
    lastColumn - firstColumn + 1
  ).getValues();
  return values.map(function (row) {
    const item = {};
    fields.forEach(function (field, index) {
      item[field.key] = operatorCardNormalizeText_(row[fieldColumns[index] - firstColumn]);
    });
    return item;
  }).filter(function (item) { return item.id && item.name; });
}

function operatorCardReadUniqueColumn_(sheetKey, header) {
  const context = getSystemSheetContext_(sheetKey);
  const lastRow = context.sheet.getLastRow();
  if (lastRow < context.config.dataStartRow) return [];
  const column = context.headerMap[sysNormalizeHeader_(header)];
  const values = context.sheet.getRange(
    context.config.dataStartRow,
    column,
    lastRow - context.config.dataStartRow + 1,
    1
  ).getValues();
  const seen = {};
  const result = [];
  values.forEach(function (row) {
    const value = operatorCardNormalizeText_(row[0]);
    const key = operatorCardFold_(value);
    if (!value || seen[key]) return;
    seen[key] = true;
    result.push(value);
  });
  return result;
}

function operatorCardWithAll_(values) {
  return [SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL].concat(values);
}

function operatorCardBuildObjects_(rows) {
  const allName = SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL;
  if (rows.some(function (item) { return operatorCardFold_(item.name) === operatorCardFold_(allName); })) {
    throw new Error('В листе «Объекты» обнаружен реальный объект с системным названием «Все». Переименуйте объект или удалите конфликтующую строку.');
  }
  rows.sort(function (a, b) { return operatorCardNaturalCompare_(a.name, b.name) || operatorCardNaturalCompare_(a.id, b.id); });
  return [{ id: '', name: allName, isAllObjects: true }].concat(rows.map(function (item) {
    return { id: item.id, name: item.name, isAllObjects: false };
  }));
}

function operatorCardGetFilterData_() {
  const objects = operatorCardBuildObjects_(operatorCardReadDictionary_('OBJECTS', [
    { key: 'id', header: H.OBJECT_ID }, { key: 'name', header: H.OBJECT_NAME }
  ]));
  const employees = operatorCardReadDictionary_('EMPLOYEES', [
    { key: 'id', header: H.EMPLOYEE_ID }, { key: 'name', header: H.EMPLOYEE_NAME },
    { key: 'position', header: H.EMPLOYEE_POSITION }
  ]).sort(function (a, b) { return operatorCardNaturalCompare_(a.name, b.name) || operatorCardNaturalCompare_(a.id, b.id); });
  const clients = operatorCardReadDictionary_('CLIENTS', [
    { key: 'id', header: H.CLIENT_ID },
    { key: 'organizationName', header: H.CLIENT_NAME },
    { key: 'objectName', header: H.OBJECT_NAME },
    { key: 'name', header: H.EMPLOYEE_NAME },
    { key: 'status', header: H.CLIENT_STATUS }
  ]).map(function (client) {
    client.contactName = client.name;
    client.name = operatorCardClientDisplay_(client.contactName, client.organizationName);
    return client;
  }).filter(function (client) { return client.name; }).sort(function (a, b) {
    return operatorCardNaturalCompare_(a.name, b.name) || operatorCardNaturalCompare_(a.id, b.id);
  });
  const clientIds = {};
  clients.forEach(function (client) {
    if (clientIds[client.id]) {
      throw new Error('В листе «Справочник клиентов» повторяется ID клиента «' + client.id + '».');
    }
    clientIds[client.id] = true;
  });
  const holders = employees.map(function (x) { return { id: x.id, name: x.name, type: 'employee' }; }).concat(clients.map(function (x) { return { id: x.id, name: x.name, type: 'client' }; }));
  const documentTypes = operatorCardReadDictionary_('DOCUMENT_TYPES', [
    { key: 'id', header: H.DOCUMENT_TYPE_ID }, { key: 'name', header: H.DOCUMENT_TYPE }
  ]).sort(function (a, b) { return operatorCardNaturalCompare_(a.name, b.name) || operatorCardNaturalCompare_(a.id, b.id); });
  const documentStatuses = operatorCardWithAll_(
    operatorCardReadUniqueColumn_('CARD_DICTIONARY', H.DOCUMENT_STATUS)
  );
  return {
    objects: objects,
    foremen: employees.filter(function (employee) {
      return SYSTEM_CONFIG.VALUES.FOREMAN_POSITIONS.some(function (position) {
        return operatorCardFold_(employee.position) === operatorCardFold_(position);
      });
    }),
    employees: employees,
    clients: clients,
    holders: holders,
    documentTypes: documentTypes,
    documentStatuses: documentStatuses
  };
}

/** Detects a card row whose document identity still points to its physical fact row. */
function operatorCardHasLoadedRow_() {
  const card = getSystemSheetContext_('OPERATOR_CARD');
  const documents = getSystemSheetContext_('DOCUMENTS');
  const cardIdColumn = card.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)];
  const cardFactRowColumn = card.headerMap[sysNormalizeHeader_(H.FACT_ROW_NUMBER)];
  const lastCardRow = card.sheet.getLastRow();
  if (lastCardRow < card.config.dataStartRow) return false;

  const firstColumn = Math.min(cardIdColumn, cardFactRowColumn);
  const values = card.sheet.getRange(
    card.config.dataStartRow, firstColumn,
    lastCardRow - card.config.dataStartRow + 1,
    Math.max(cardIdColumn, cardFactRowColumn) - firstColumn + 1
  ).getValues();
  const idOffset = cardIdColumn - firstColumn;
  const rowOffset = cardFactRowColumn - firstColumn;
  const documentIdColumn = documents.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)];
  const lastDocumentRow = documents.sheet.getLastRow();

  return values.some(function (row) {
    const documentId = operatorCardNormalizeText_(row[idOffset]);
    const factRow = Number(row[rowOffset]);
    if (!documentId || !Number.isInteger(factRow) ||
        factRow < documents.config.dataStartRow || factRow > lastDocumentRow) return false;
    const factDocumentId = documents.sheet.getRange(factRow, documentIdColumn).getValue();
    return operatorCardNormalizeText_(factDocumentId) === documentId;
  });
}

function operatorCardNormalizeObjectFilter_(filter) {
  const value = filter || {};
  const id = operatorCardNormalizeText_(value.objectId);
  const name = operatorCardNormalizeText_(value.objectName);
  if (!filter || value.allObjects === true || (value.allObjects !== false && !id && !name) || operatorCardFold_(name) === operatorCardFold_(SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL)) {
    return { allObjects: true, objectId: '', objectName: SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL };
  }
  if (value.allObjects === false && !id) throw new Error('Для конкретного объекта необходимо передать ID объекта.');
  if (!id) throw new Error('Не удалось определить ID выбранного объекта.');
  return { allObjects: false, objectId: id, objectName: name };
}

function operatorCardNormalizeFilters_(filters) {
  const source = filters || {};
  const object = operatorCardNormalizeObjectFilter_(source.object);
  return {
    object: object,
    foremanId: operatorCardNormalizeText_(source.foremanId),
    foremanName: operatorCardNormalizeText_(source.foremanName),
    documentStatus: operatorCardFold_(source.documentStatus) === operatorCardFold_(SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL)
      ? '' : operatorCardNormalizeText_(source.documentStatus),
    documentTypeId: operatorCardNormalizeText_(source.documentTypeId),
    documentTypeName: operatorCardNormalizeText_(source.documentTypeName),
    holderId: operatorCardNormalizeText_(source.holderId),
    holderName: operatorCardNormalizeText_(source.holderName),
    dateFrom: operatorCardNormalizeText_(source.dateFrom),
    dateTo: operatorCardNormalizeText_(source.dateTo)
  };
}

function operatorCardValidateSelection_(filters, data) {
  function requireId(id, items, label) {
    if (id && !items.some(function (item) { return item.id === id; })) throw new Error('Выбран неизвестный ' + label + ' (ID «' + id + '»). Обновите справочники.');
  }
  if (!filters.object.allObjects) requireId(filters.object.objectId, data.objects, 'объект');
  requireId(filters.foremanId, data.foremen, 'ответственный прораб');
  requireId(filters.documentTypeId, data.documentTypes, 'тип документа');
  requireId(filters.holderId, data.holders, 'держатель документа');
  if (filters.documentStatus && !data.documentStatuses.some(function (status) {
    return operatorCardFold_(status) === operatorCardFold_(filters.documentStatus);
  })) throw new Error('Выбран неизвестный статус документа «' + filters.documentStatus + '». Обновите справочники.');
}

function operatorCardMergeUniqueValues_(groups) {
  const seen = {};
  const result = [];
  groups.forEach(function (values) {
    values.forEach(function (value) {
      const normalized = operatorCardNormalizeText_(value);
      const key = operatorCardFold_(normalized);
      if (!normalized || seen[key]) return;
      seen[key] = true;
      result.push(normalized);
    });
  });
  return result;
}

function operatorCardDisplayLabel_(name, id) {
  return id ? operatorCardNormalizeText_(name) + ' [' + operatorCardNormalizeText_(id) + ']' : operatorCardNormalizeText_(name);
}

function operatorCardGetValidationData_(filterData) {
  const data = filterData || operatorCardGetFilterData_();
  const employeeLabels = data.employees.map(function (x) { return operatorCardDisplayLabel_(x.name, x.id); });
  const holderLabels = data.holders.map(function (x) { return operatorCardDisplayLabel_(x.name, x.id); });
  return [
    { header: H.DOCUMENT_STATUS, values: operatorCardReadUniqueColumn_('CARD_DICTIONARY', H.DOCUMENT_STATUS) },
    { header: H.ORIGINAL_EDO, values: operatorCardReadUniqueColumn_('CARD_DICTIONARY', H.ORIGINAL_EDO) },
    { header: H.DOCUMENT_HOLDER, values: holderLabels },
    { header: H.DOCUMENT_LOCATION, values: operatorCardReadUniqueColumn_('CARD_DICTIONARY', H.DOCUMENT_LOCATION) },
    { header: H.TRANSFERRED_BY, values: employeeLabels },
    { header: H.PAID, values: operatorCardReadUniqueColumn_('CARD_DICTIONARY', H.PAID) },
    { header: H.GU_FLAG, values: operatorCardReadUniqueColumn_('CARD_DICTIONARY', H.GU_FLAG) },
    { header: H.RECORD_STATUS, values: operatorCardReadUniqueColumn_('CARD_DICTIONARY', H.RECORD_STATUS) }
  ];
}

/** Applies and clears object-specific validations with one batch write. */
function operatorCardApplyRowValidations_(cardContext, cardRows, filterData) {
  if (!cardContext.sheet.getRange) return;
  const cardStartColumn = cardContext.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)];
  const objectOffset = cardContext.headerMap[sysNormalizeHeader_(H.OBJECT_NAME)] - cardStartColumn;
  const firstColumn = cardContext.headerMap[sysNormalizeHeader_(H.TRANSFERRED_BY)];
  const signerColumn = cardContext.headerMap[sysNormalizeHeader_(H.CUSTOMER_SIGNING_RESPONSIBLE)];
  const width = signerColumn - firstColumn + 1;
  const rowCount = cardContext.sheet.getMaxRows
    ? cardContext.sheet.getMaxRows() - cardContext.config.dataStartRow + 1 : cardRows.length;
  const employeeLabels = filterData.employees.map(function (employee) {
    return operatorCardDisplayLabel_(employee.name, employee.id);
  });
  const rules = Array.from({ length: rowCount }, function (_, rowIndex) {
    const output = Array(width).fill(null);
    if (rowIndex < cardRows.length) {
      const row = cardRows[rowIndex];
      const clients = operatorCardClientsForObject_(filterData.clients, row[objectOffset]);
      const clientLabels = clients.map(function (client) { return client.name; });
      const transferred = operatorCardMergeUniqueValues_([employeeLabels, clientLabels]);
      output[0] = transferred.length ? SpreadsheetApp.newDataValidation()
        .requireValueInList(transferred, true).setAllowInvalid(true).build() : null;
      output[signerColumn - firstColumn] = clientLabels.length
        ? SpreadsheetApp.newDataValidation().requireValueInList(clientLabels, true)
          .setAllowInvalid(true).build() : null;
    }
    return output;
  });
  cardContext.sheet.getRange(
    cardContext.config.dataStartRow, firstColumn, rowCount, width
  ).setDataValidations(rules);
}

function operatorCardApplyValidations_(cardContext, validationData) {
  const rowCount = cardContext.sheet.getMaxRows() - cardContext.config.dataStartRow + 1;
  if (rowCount < 1) return;
  validationData.forEach(function (validation) {
    if (!validation.values.length) return;
    const column = cardContext.headerMap[sysNormalizeHeader_(validation.header)];
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(validation.values, true)
      .setAllowInvalid(true)
      .build();
    cardContext.sheet.getRange(
      cardContext.config.dataStartRow,
      column,
      rowCount,
      1
    ).setDataValidation(rule);
  });
}

function operatorCardParseDateRange_(fromText, toText, timezone) {
  const pattern = /^\d{4}-\d{2}-\d{2}$/;
  function parse(text) {
    if (!text) return null;
    if (!pattern.test(text)) throw new Error('Некорректная календарная дата: «' + text + '».');
    const parts = text.split('-').map(Number);
    const canonical = parts[2] + '.' + parts[1] + '.' + parts[0];
    const result = Utilities.parseDate(canonical, timezone, 'dd.MM.yyyy');
    if (!result || typeof result.getTime !== 'function' || isNaN(result.getTime()) || Utilities.formatDate(result, timezone, 'yyyy-MM-dd') !== text) {
      throw new Error('Некорректная календарная дата: «' + text + '».');
    }
    return result;
  }
  const from = parse(fromText);
  const to = parse(toText);
  if (from && to && from.getTime() > to.getTime()) throw new Error('Дата от не может быть позже даты до.');
  let toExclusive = null;
  if (toText) {
    const parts = toText.split('-').map(Number);
    const nextCalendarDay = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + 1, 12));
    const nextText = Utilities.formatDate(nextCalendarDay, 'UTC', 'yyyy-MM-dd');
    toExclusive = parse(nextText);
  }
  return { active: !!(from || to), from: from, toExclusive: toExclusive };
}

function operatorCardNaturalCompare_(left, right) {
  return String(left == null ? '' : left).localeCompare(String(right == null ? '' : right), 'ru', { numeric: true, sensitivity: 'base' });
}

function operatorCardValidateHeaders_(documentsHeaders, cardHeaders, documentsName, cardName, documentsStartColumn, cardStartColumn) {
  const expectedCount = SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.requiredHeaders.length;
  const documentsStartIndex = documentsStartColumn == null
    ? documentsHeaders.map(operatorCardNormalizeText_).indexOf(H.DOCUMENT_ID)
    : documentsStartColumn - 1;
  const cardStartIndex = cardStartColumn == null
    ? cardHeaders.map(operatorCardNormalizeText_).indexOf(H.DOCUMENT_ID)
    : cardStartColumn - 1;
  if (documentsStartIndex < 0) throw new Error('Лист «' + documentsName + '»: не найден заголовок «' + H.DOCUMENT_ID + '».');
  if (cardStartIndex < 0) throw new Error('Лист «' + cardName + '»: не найден заголовок «' + H.DOCUMENT_ID + '».');
  for (let index = 0; index < expectedCount; index++) {
    const configured = SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.requiredHeaders[index];
    const actual = operatorCardNormalizeText_(cardHeaders[cardStartIndex + index]);
    if (index < expectedCount - 2) {
      const expected = operatorCardNormalizeText_(documentsHeaders[documentsStartIndex + index]);
      if (expected !== configured) throw new Error('Лист «' + documentsName + '», позиция ' + (index + 1) + ' рабочего блока: ожидается заголовок «' + configured + '», фактически «' + expected + '».');
    }
    if (actual !== operatorCardNormalizeText_(configured)) throw new Error('Лист «' + cardName + '», позиция ' + (index + 1) + ' рабочего блока: ожидается заголовок «' + configured + '», фактически «' + actual + '».');
  }
  for (let index = cardStartIndex + expectedCount; index < cardHeaders.length; index++) {
    const actual = operatorCardNormalizeText_(cardHeaders[index]);
    if (actual) throw new Error('Лист «' + cardName + '», физическая колонка ' + (index + 1) + ': ожидается пустой заголовок после утверждённых ' + expectedCount + ' полей, фактически «' + actual + '».');
  }
  return true;
}

function operatorCardPrepareRows_(rows, indexes, filters, dateRange, dataStartRow, sourceStartIndex, dictionaries) {
  sourceStartIndex = sourceStartIndex == null ? 0 : sourceStartIndex;
  const warnings = [];
  const active = [];
  rows.forEach(function (row, offset) {
    if (row.every(function (value) { return value === '' || value == null; })) return;
    if (operatorCardFold_(row[indexes.recordStatus]) !== operatorCardFold_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS)) return;
    active.push({ row: row, sheetRow: dataStartRow + offset });
  });
  const groups = {};
  active.forEach(function (item) {
    const id = operatorCardNormalizeText_(item.row[indexes.documentId]);
    if (id) (groups[id] = groups[id] || []).push(item.sheetRow);
  });
  const duplicateGroups = {};
  Object.keys(groups).forEach(function (id) { if (groups[id].length > 1) duplicateGroups[id] = groups[id]; });
  let invalidDates = 0;
  const selected = active.filter(function (item) {
    const row = item.row;
    if (!filters.object.allObjects && operatorCardNormalizeText_(row[indexes.objectId]) !== filters.object.objectId) return false;
    if (filters.foremanId && operatorCardNormalizeText_(row[indexes.foremanId]) !== filters.foremanId) return false;
    if (filters.documentStatus && operatorCardFold_(row[indexes.documentStatus]) !== operatorCardFold_(filters.documentStatus)) return false;
    if (filters.documentTypeId && operatorCardNormalizeText_(row[indexes.documentTypeId]) !== filters.documentTypeId) return false;
    if (filters.holderId && operatorCardNormalizeText_(row[indexes.holderId]) !== filters.holderId) return false;
    if (dateRange.active) {
      const date = row[indexes.createdAt];
      if (date === '' || date == null) return false;
      if (!(date instanceof Date) || isNaN(date.getTime())) { invalidDates++; return false; }
      if (dateRange.from && date < dateRange.from) return false;
      if (dateRange.toExclusive && date >= dateRange.toExclusive) return false;
    }
    return true;
  });
  selected.sort(function (a, b) {
    return operatorCardNaturalCompare_(a.row[indexes.objectId], b.row[indexes.objectId]) ||
      operatorCardNaturalCompare_(a.row[indexes.documentTypeId], b.row[indexes.documentTypeId]) ||
      operatorCardNaturalCompare_(a.row[indexes.documentId], b.row[indexes.documentId]) || a.sheetRow - b.sheetRow;
  });
  const selectedDuplicateIds = {};
  selected.forEach(function (item) {
    const id = operatorCardNormalizeText_(item.row[indexes.documentId]);
    if (duplicateGroups[id]) selectedDuplicateIds[id] = true;
  });
  Object.keys(selectedDuplicateIds).sort(operatorCardNaturalCompare_).forEach(function (id) {
    warnings.push('ID документа «' + id + '» повторяется в активных строках ' + duplicateGroups[id].join(' и ') + '.');
  });
  if (invalidDates) warnings.push('Строк с некорректной датой создания пропущено: ' + invalidDates + '.');
  return { activeCount: active.length, rows: selected, cardRows: selected.map(function (item) {
    const sharedFieldCount = SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.requiredHeaders.length - 2;
    const cardRow = item.row.slice(sourceStartIndex, sourceStartIndex + sharedFieldCount);
    cardRow[indexes.documentHolder - sourceStartIndex] = operatorCardDisplayLabel_(item.row[indexes.documentHolder], item.row[indexes.holderId]);
    const transferredId = operatorCardNormalizeText_(item.row[indexes.transferredById]);
    cardRow[indexes.transferredBy - sourceStartIndex] = transferredId.indexOf(SYSTEM_CONFIG.ID_PREFIXES.CLIENT) === 0
      ? operatorCardNormalizeText_(item.row[indexes.transferredBy])
      : operatorCardDisplayLabel_(item.row[indexes.transferredBy], transferredId);
    return cardRow.concat([item.row[indexes.recordStatus], item.sheetRow]);
  }), warnings: warnings, duplicateIdsCount: Object.keys(selectedDuplicateIds).length, invalidDatesCount: invalidDates };
}

function operatorCardReplace_(cardContext, newRows) {
  const sheet = cardContext.sheet;
  const start = cardContext.config.dataStartRow;
  const cardStartColumn = cardContext.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)];
  const available = Math.max(sheet.getLastRow() - start + 1, 0);
  let oldCount = 0;
  if (available) {
    const ids = sheet.getRange(start, cardStartColumn, available, 1).getValues();
    ids.forEach(function (row, index) { if (operatorCardNormalizeText_(row[0])) oldCount = index + 1; });
  }
  const writeCount = Math.max(oldCount, newRows.length);
  if (!writeCount) return;
  const output = Array.from({ length: writeCount }, function (_, index) {
    return index < newRows.length ? newRows[index].slice() :
      Array(cardContext.config.requiredHeaders.length).fill('');
  });
  const cardWidth = cardContext.config.requiredHeaders.length;
  sheet.getRange(start, cardStartColumn, writeCount, cardWidth).setValues(output);
  if (sheet.hideColumns) {
    sheet.hideColumns(cardStartColumn, 2);
    sheet.hideColumns(cardStartColumn + cardWidth - 1);
  }
}

/** Restores manually edited read-only card cells from their physical fact rows. */
function operatorCardHandleReadOnlyEdit_(event) {
  if (!event || !event.range ||
      event.range.getSheet().getName() !== SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.name) return;
  const card = getSystemSheetContext_('OPERATOR_CARD');
  const edit = event.range;
  const firstRow = Math.max(edit.getRow(), card.config.dataStartRow);
  const lastRow = edit.getRow() + edit.getNumRows() - 1;
  if (lastRow < firstRow) return;
  const firstColumn = edit.getColumn();
  const lastColumn = firstColumn + edit.getNumColumns() - 1;
  const readOnlyMappings = SYSTEM_CONFIG.CARD_FIELD_MAP.filter(function (mapping) {
    if (mapping.editable !== false) return false;
    const column = card.headerMap[sysNormalizeHeader_(mapping.cardHeader)];
    return column >= firstColumn && column <= lastColumn;
  });
  if (!readOnlyMappings.length) return;

  const cardStartColumn = card.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)];
  const cardWidth = card.config.requiredHeaders.length;
  const cardRows = card.sheet.getRange(firstRow, cardStartColumn, lastRow - firstRow + 1, cardWidth).getValues();
  const rowColumn = card.headerMap[sysNormalizeHeader_(H.FACT_ROW_NUMBER)] - cardStartColumn;
  const documents = getSystemSheetContext_('DOCUMENTS');
  const factRows = cardRows.map(function (row) { return Number(row[rowColumn]); });
  const validFactRows = factRows.filter(function (row) {
    return Number.isInteger(row) && row >= documents.config.dataStartRow && row <= documents.sheet.getLastRow();
  });
  const minFactRow = validFactRows.length ? Math.min.apply(null, validFactRows) : 0;
  const maxFactRow = validFactRows.length ? Math.max.apply(null, validFactRows) : -1;
  const factValues = validFactRows.length ? documents.sheet.getRange(
    minFactRow, 1, maxFactRow - minFactRow + 1, documents.headers.length
  ).getValues() : [];
  const identityHeaders = [H.DOCUMENT_ID, H.OBJECT_ID, H.DOCUMENT_TYPE_ID];
  const automaticNote = 'Поле заполняется автоматически и недоступно для ручного изменения.';
  const unsafeNote = 'Не удалось безопасно восстановить поле: неверная физическая строка или identity документа. Обновите карточку.';

  cardRows.forEach(function (cardRow, offset) {
    const sheetRow = firstRow + offset;
    const factRow = factRows[offset];
    const affectedHeaders = {};
    readOnlyMappings.forEach(function (mapping) { affectedHeaders[mapping.cardHeader] = true; });
    const fact = Number.isInteger(factRow) && factRow >= minFactRow && factRow <= maxFactRow
      ? factValues[factRow - minFactRow] : null;
    const identityMatches = !!fact && identityHeaders.every(function (header) {
      if (affectedHeaders[header]) return true;
      const cardIndex = card.headerMap[sysNormalizeHeader_(header)] - cardStartColumn;
      const factIndex = documents.headerMap[sysNormalizeHeader_(header)] - 1;
      return operatorCardNormalizeText_(cardRow[cardIndex]) === operatorCardNormalizeText_(fact[factIndex]);
    });
    const hasIdentityAnchor = identityHeaders.some(function (header) { return !affectedHeaders[header]; });
    readOnlyMappings.forEach(function (mapping) {
      const column = card.headerMap[sysNormalizeHeader_(mapping.cardHeader)];
      const cell = card.sheet.getRange(sheetRow, column);
      if (!identityMatches || !hasIdentityAnchor || !mapping.factHeader) {
        cell.setNote(unsafeNote);
        return;
      }
      const factIndex = documents.headerMap[sysNormalizeHeader_(mapping.factHeader)] - 1;
      cell.setValue(fact[factIndex]).setNote(automaticNote);
    });
  });
}

function operatorCardComment_(filters, result, criticalError) {
  const parts = [
    'Объект: ' + (filters.object.allObjects ? 'Все' : (filters.object.objectName || filters.object.objectId)),
    'Прораб: ' + (filters.foremanName || filters.foremanId || 'Все'), 'Статус: ' + (filters.documentStatus || 'Все'),
    'Тип: ' + (filters.documentTypeName || filters.documentTypeId || 'Все'), 'Держатель: ' + (filters.holderName || filters.holderId || 'Все'),
    'Даты: ' + (filters.dateFrom || 'без начала') + ' — ' + (filters.dateTo || 'без окончания'),
    'Активных строк: ' + (result.activeCount || 0), 'Загружено: ' + (result.loadedCount || 0),
    'Дублирующихся ID: ' + (result.duplicateIdsCount || 0), 'Некорректных дат: ' + (result.invalidDatesCount || 0)
  ];
  if (result.warnings && result.warnings.length) parts.push('Предупреждения: ' + result.warnings.join(' | '));
  if (criticalError) parts.push('Ошибка: ' + criticalError);
  return parts.join('; ');
}

function operatorCardWriteOperation_(operation) {
  const context = getSystemSheetContext_('OPERATION_HISTORY');
  const values = {};
  values[H.OPERATION_ID] = operation.id; values[H.OPERATION_STARTED_AT] = operation.startedAt;
  values[H.OPERATION_FINISHED_AT] = operation.finishedAt; values[H.OPERATION_STARTED_BY] = operation.email;
  values[H.OPERATION_SOURCE] = SYSTEM_CONFIG.VALUES.OPERATOR_CARD_SOURCE;
  values[H.OPERATION_TYPE] = SYSTEM_CONFIG.VALUES.OPERATOR_CARD_LOAD_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = operation.status; values[H.DOCUMENTS_LOADED] = operation.loadedCount;
  values[H.DOCUMENTS_CHANGED] = 0; values[H.FACT_ROWS_UPDATED] = 0; values[H.FIELDS_CHANGED] = 0;
  values[H.DUPLICATE_IDS_FOUND] = operation.duplicateIdsCount; values[H.ERRORS_COUNT] = operation.errorsCount;
  values[H.EXECUTION_SECONDS] = (operation.finishedAt.getTime() - operation.startedAt.getTime()) / 1000;
  values[H.ERROR_TEXT] = operation.comment;
  const row = context.headers.map(function (header) { return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : ''; });
  const start = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(start, 1, 1, context.headers.length).setValues([row]);
}

function operatorCardApply_(rawFilters, suppressJournal) {
  const startedAt = new Date();
  return withDocumentLock_(function () {
    let filters = {
      object: { allObjects: true, objectId: '', objectName: SYSTEM_CONFIG.VALUES.ALL_OBJECTS_LABEL },
      foremanId: '', foremanName: '', documentStatus: '', documentTypeId: '',
      documentTypeName: '', holderId: '', holderName: '', dateFrom: '', dateTo: ''
    };
    let operationId = '';
    let cardWritten = false;
    let result = { activeCount: 0, loadedCount: 0, duplicateIdsCount: 0, invalidDatesCount: 0, warnings: [] };
    try {
      filters = operatorCardNormalizeFilters_(rawFilters);
      if (!suppressJournal) operationId = generateOperationId_(startedAt);
      const dictionaries = operatorCardGetFilterData_();
      operatorCardValidateSelection_(filters, dictionaries);
      const ss = getSystemSpreadsheet_();
      const range = operatorCardParseDateRange_(filters.dateFrom, filters.dateTo, ss.getSpreadsheetTimeZone());
      const documents = getSystemSheetContext_('DOCUMENTS');
      const card = getSystemSheetContext_('OPERATOR_CARD');
      const documentsStartColumn = documents.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)];
      const cardStartColumn = card.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)];
      operatorCardValidateHeaders_(documents.headers, card.headers, documents.config.name, card.config.name, documentsStartColumn, cardStartColumn);
      const lastRow = documents.sheet.getLastRow();
      const rows = lastRow < documents.config.dataStartRow ? [] : documents.sheet.getRange(documents.config.dataStartRow, 1, lastRow - documents.config.dataStartRow + 1, documents.headers.length).getValues();
      function index(header) { return documents.headerMap[sysNormalizeHeader_(header)] - 1; }
      result = operatorCardPrepareRows_(rows, {
        documentId: index(H.DOCUMENT_ID), objectId: index(H.OBJECT_ID), documentTypeId: index(H.DOCUMENT_TYPE_ID),
        documentStatus: index(H.DOCUMENT_STATUS), holderId: index(H.HOLDER_EMPLOYEE_ID), foremanId: index(H.RESPONSIBLE_FOREMAN_ID),
        createdAt: index(H.CREATED_AT), recordStatus: index(H.RECORD_STATUS), transferredById: index(H.TRANSFERRED_BY_EMPLOYEE_ID),
        documentHolder: index(H.DOCUMENT_HOLDER), transferredBy: index(H.TRANSFERRED_BY)
      }, filters, range, documents.config.dataStartRow, documentsStartColumn - 1, dictionaries);
      result.loadedCount = result.cardRows.length;
      operatorCardReplace_(card, result.cardRows);
      operatorCardApplyValidations_(card, operatorCardGetValidationData_(dictionaries));
      operatorCardApplyRowValidations_(card, result.cardRows, dictionaries);
      cardWritten = true;
      const status = result.loadedCount === 0 ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES : (result.warnings.length ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS : SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS);
      const finishedAt = new Date();
      try {
        if (!suppressJournal) operatorCardWriteOperation_({ id: operationId, startedAt: startedAt, finishedAt: finishedAt, email: getActiveUserEmail_(), status: status, loadedCount: result.loadedCount, duplicateIdsCount: result.duplicateIdsCount, errorsCount: 0, comment: operatorCardComment_(filters, result, '') });
      } catch (historyError) {
        const historyMessage = historyError && historyError.message ? historyError.message : String(historyError);
        const warning = 'Карточка загружена, но запись в «История операций» завершилась ошибкой. Причина: ' + historyMessage;
        result.warnings.push(warning);
        if (typeof console !== 'undefined' && console.error) console.error(warning);
        return { success: true, status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS, operationId: operationId, loadedCount: result.loadedCount, duplicateIdsCount: result.duplicateIdsCount, warnings: result.warnings, message: result.warnings[result.warnings.length - 1] };
      }
      return { success: true, status: status, operationId: operationId, loadedCount: result.loadedCount, duplicateIdsCount: result.duplicateIdsCount, warnings: result.warnings, message: result.loadedCount ? ('Загружено документов: ' + result.loadedCount + (result.warnings.length ? '. Найдены предупреждения.' : '.')) : 'Документы не найдены. Предыдущая выдача карточки очищена.' };
    } catch (error) {
      const message = error && error.message ? error.message : String(error);
      if (!operationId && !suppressJournal) { try { operationId = generateOperationId_(startedAt); } catch (ignored) {} }
      try {
        if (!suppressJournal) operatorCardWriteOperation_({ id: operationId, startedAt: startedAt, finishedAt: new Date(), email: getActiveUserEmail_(), status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR, loadedCount: cardWritten ? result.loadedCount : 0, duplicateIdsCount: result.duplicateIdsCount || 0, errorsCount: 1, comment: operatorCardComment_(filters, result, message) });
      } catch (ignoredHistoryError) {}
      throw new Error(message + (cardWritten ? '' : ' Карточка не была изменена. На листе осталась предыдущая выдача.'));
    }
  });
}
