/**
 * Data-only migration to document architecture v2.
 *
 * Run only after the owner has manually installed the v2 headers. The function
 * deliberately never inserts, deletes, moves, or rewrites columns or headers.
 * It resolves every field through SystemCore and is safe to run repeatedly.
 */
function migrateDocumentsToArchitectureV2() {
  return withDocumentLock_(function () {
    assertSystemSheetsStructure_([
      'DOCUMENT_TYPES', 'DOCUMENTS', 'OPERATION_HISTORY', 'CHANGE_HISTORY'
    ]);

    const startedAt = new Date();
    const operationId = generateOperationId_(startedAt);
    const email = getActiveUserEmail_();
    const typeContext = getSystemSheetContext_('DOCUMENT_TYPES');
    const documentContext = getSystemSheetContext_('DOCUMENTS');
    const typeRows = migrationReadRows_(typeContext);
    const documentRows = migrationReadRows_(documentContext);
    const typeIdColumn = migrationColumn_(typeContext, H.DOCUMENT_TYPE_ID);
    const typeNameColumn = migrationColumn_(typeContext, H.DOCUMENT_TYPE);
    const types = {};

    typeRows.forEach(function (row, offset) {
      const id = migrationText_(row[typeIdColumn]);
      const name = migrationText_(row[typeNameColumn]);
      if (!id || !name) return;
      if (types[id] && types[id].name !== name) {
        throw new Error('ID типа документа «' + id + '» связан с разными названиями в справочнике.');
      }
      types[id] = { name: name, row: typeContext.config.dataStartRow + offset };
    });

    const indexes = {};
    [H.DOCUMENT_ID, H.OBJECT_ID, H.DOCUMENT_TYPE_ID, H.DOCUMENT_TYPE,
      H.DOCUMENT_NUMBER, H.UPDATED_AT, H.UPDATED_BY_EMAIL].forEach(function (header) {
      indexes[header] = migrationColumn_(documentContext, header);
    });
    const changes = [];
    const changedRows = [];
    const now = new Date();

    documentRows.forEach(function (row, offset) {
      const typeId = migrationText_(row[indexes[H.DOCUMENT_TYPE_ID]]);
      if (!typeId) return;
      const canonical = types[typeId];
      if (!canonical) {
        throw new Error('В строке ' + (documentContext.config.dataStartRow + offset) +
          ' неизвестный «' + H.DOCUMENT_TYPE_ID + '»: «' + typeId + '».');
      }
      const oldType = migrationText_(row[indexes[H.DOCUMENT_TYPE]]);
      const oldNumber = row[indexes[H.DOCUMENT_NUMBER]];
      const number = documentArchitectureExtractLegacyNumber_(
        oldType, canonical.name, oldNumber
      );
      if (migrationEmpty_(oldNumber) && !number &&
          documentArchitectureHasLegacyMarker_(oldType, canonical.name)) {
        throw new Error('В строке ' +
          (documentContext.config.dataStartRow + offset) +
          ' legacy-номер в поле «' + H.DOCUMENT_TYPE +
          '» не является положительным целым числом. Миграция отменена.');
      }
      const rowChanges = [];
      if (oldType !== canonical.name) {
        rowChanges.push({ header: H.DOCUMENT_TYPE, oldValue: oldType, newValue: canonical.name });
        row[indexes[H.DOCUMENT_TYPE]] = canonical.name;
      }
      if (migrationEmpty_(oldNumber) && number) {
        rowChanges.push({ header: H.DOCUMENT_NUMBER, oldValue: '', newValue: number });
        row[indexes[H.DOCUMENT_NUMBER]] = number;
      }
      if (!rowChanges.length) return;
      row[indexes[H.UPDATED_AT]] = now;
      row[indexes[H.UPDATED_BY_EMAIL]] = email;
      const sheetRow = documentContext.config.dataStartRow + offset;
      rowChanges.forEach(function (change) {
        changes.push({
          changeId: generateChangeId_(operationId, changes.length + 1),
          operationId: operationId, changedAt: now, userEmail: email,
          documentId: migrationText_(row[indexes[H.DOCUMENT_ID]]),
          objectId: migrationText_(row[indexes[H.OBJECT_ID]]), factRow: sheetRow,
          fieldName: change.header, oldValue: change.oldValue, newValue: change.newValue
        });
      });
      changedRows.push({ sheetRow: sheetRow, row: row });
    });

    migrationWriteColumns_(documentContext, changedRows, [
      H.DOCUMENT_TYPE, H.DOCUMENT_NUMBER, H.UPDATED_AT, H.UPDATED_BY_EMAIL
    ]);
    migrationWriteHistory_(changes);
    migrationWriteOperation_(operationId, startedAt, new Date(), email,
      changedRows.length, changes.length);
    return { operationId: operationId, rowsChanged: changedRows.length,
      fieldsChanged: changes.length };
  });
}

function migrationReadRows_(context) {
  const count = context.sheet.getLastRow() - context.config.dataStartRow + 1;
  return count > 0 ? context.sheet.getRange(context.config.dataStartRow, 1,
    count, context.headers.length).getValues() : [];
}
function migrationColumn_(context, header) {
  return context.headerMap[sysNormalizeHeader_(header)] - 1;
}
function migrationText_(value) {
  return String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim();
}
function migrationEmpty_(value) { return value == null || migrationText_(value) === ''; }

function migrationWriteColumns_(context, items, headers) {
  headers.forEach(function (header) {
    const index = migrationColumn_(context, header);
    const sorted = items.slice().sort(function (left, right) {
      return left.sheetRow - right.sheetRow;
    });
    const groups = [];
    sorted.forEach(function (item) {
      const group = groups[groups.length - 1];
      const value = [item.row[index]];
      if (group && item.sheetRow === group.startRow + group.values.length) {
        group.values.push(value);
      } else {
        groups.push({ startRow: item.sheetRow, values: [value] });
      }
    });
    groups.forEach(function (group) {
      context.sheet.getRange(
        group.startRow, index + 1, group.values.length, 1
      ).setValues(group.values);
    });
  });
}

function migrationWriteHistory_(changes) {
  if (!changes.length) return;
  const context = getSystemSheetContext_('CHANGE_HISTORY');
  const rows = changes.map(function (change) {
    const values = {};
    values[H.CHANGE_ID] = change.changeId; values[H.OPERATION_ID] = change.operationId;
    values[H.CHANGE_DATETIME] = change.changedAt; values[H.CHANGED_BY_EMAIL] = change.userEmail;
    values[H.ACTION_TYPE] = SYSTEM_CONFIG.VALUES.CHANGE_ACTION_EDIT;
    values[H.DOCUMENT_ID] = change.documentId; values[H.OBJECT_ID] = change.objectId;
    values[H.FACT_ROW_NUMBER] = change.factRow; values[H.FIELD_NAME] = change.fieldName;
    values[H.OLD_VALUE] = change.oldValue; values[H.NEW_VALUE] = change.newValue;
    values[H.CHANGE_SOURCE] = SYSTEM_CONFIG.VALUES.DOCUMENT_V2_MIGRATION_SOURCE;
    return context.headers.map(function (header) {
      return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
    });
  });
  const start = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(start, 1, rows.length, context.headers.length).setValues(rows);
}

function migrationWriteOperation_(id, startedAt, finishedAt, email, rows, fields) {
  const context = getSystemSheetContext_('OPERATION_HISTORY');
  const values = {};
  values[H.OPERATION_ID] = id; values[H.OPERATION_STARTED_AT] = startedAt;
  values[H.OPERATION_FINISHED_AT] = finishedAt; values[H.OPERATION_STARTED_BY] = email;
  values[H.OPERATION_SOURCE] = SYSTEM_CONFIG.VALUES.DOCUMENT_V2_MIGRATION_SOURCE;
  values[H.OPERATION_TYPE] = SYSTEM_CONFIG.VALUES.DOCUMENT_V2_MIGRATION_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = rows ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS :
    SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES;
  values[H.DOCUMENTS_CHANGED] = rows; values[H.FACT_ROWS_UPDATED] = rows;
  values[H.FIELDS_CHANGED] = fields; values[H.ERRORS_COUNT] = 0;
  values[H.EXECUTION_SECONDS] = (finishedAt.getTime() - startedAt.getTime()) / 1000;
  values[H.ERROR_TEXT] = 'Структура листов не изменялась.';
  const row = context.headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
  });
  const start = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(start, 1, 1, context.headers.length).setValues([row]);
}
