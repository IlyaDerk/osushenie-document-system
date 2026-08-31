/** Contracts, pure validation helpers, and data-only workflow migrations. */

function documentWorkflowIsBlank_(value) {
  return value == null || (typeof value === 'string' && value.trim() === '');
}

function isValidTransferredAt_(value) {
  return documentWorkflowIsBlank_(value) ||
    (value instanceof Date && !isNaN(value.getTime()));
}

function isValidImplementationDays_(value) {
  return documentWorkflowIsBlank_(value) ||
    (typeof value === 'number' && isFinite(value) &&
      Math.floor(value) === value && value >= 0);
}

/** Returns a value only; it never mutates the supplied document or rule. */
function getEffectiveImplementationDays_(manualValue, matchingRule) {
  if (!documentWorkflowIsBlank_(manualValue)) {
    if (!isValidImplementationDays_(manualValue)) {
      throw new Error('«' + H.IMPLEMENTATION_DAYS + '» must be a non-negative integer.');
    }
    return manualValue;
  }
  const ruleValue = matchingRule && Object.prototype.hasOwnProperty.call(
    matchingRule, 'implementationDays'
  ) ? matchingRule.implementationDays : matchingRule;
  if (!documentWorkflowIsBlank_(ruleValue)) {
    if (!isValidImplementationDays_(ruleValue)) {
      throw new Error('The active rule contains an invalid «' + H.IMPLEMENTATION_DAYS + '» value.');
    }
    return ruleValue;
  }
  return 7;
}

function getWorkflowControlDate_(transferredAt, effectiveDays) {
  if (documentWorkflowIsBlank_(transferredAt)) return null;
  if (!isValidTransferredAt_(transferredAt)) {
    throw new Error('«' + H.TRANSFERRED_AT + '» must be a Date.');
  }
  if (!isValidImplementationDays_(effectiveDays) ||
      documentWorkflowIsBlank_(effectiveDays)) {
    throw new Error('Effective days must be a non-negative integer.');
  }
  const result = new Date(transferredAt.getTime());
  result.setDate(result.getDate() + effectiveDays);
  return result;
}

function isDocumentFlowEmployee_(employee) {
  return !!employee && employee.active !== false &&
    sysNormalizeHeader_(employee.participatesInDocumentFlow) ===
      sysNormalizeHeader_(SYSTEM_CONFIG.VALUES.DOCUMENT_FLOW_YES);
}

function documentWorkflowNormalizedConditions_(value) {
  const fold = function (condition) {
    return sysNormalizeHeader_(condition).toLocaleLowerCase('ru');
  };
  return [
    fold(value.documentType),
    fold(value.documentStatus),
    fold(value.documentLocation)
  ];
}

/** Header-driven reader. Invalid active rules fail closed instead of being guessed. */
function readActiveDocumentWorkflowRules_() {
  const context = getSystemSheetContext_('WORKFLOW_RULES');
  const count = context.sheet.getLastRow() - context.config.dataStartRow + 1;
  if (count <= 0) return [];
  const rows = context.sheet.getRange(
    context.config.dataStartRow, 1, count, context.headers.length
  ).getValues();
  const column = function (header) {
    return context.headerMap[sysNormalizeHeader_(header)] - 1;
  };
  const activeRules = rows.map(function (row, offset) {
    return {
      row: row,
      sheetRow: context.config.dataStartRow + offset
    };
  }).filter(function (item) {
    return sysNormalizeHeader_(item.row[column(H.ACTIVE)]) ===
      sysNormalizeHeader_(SYSTEM_CONFIG.VALUES.ACTIVE_RULE);
  }).map(function (item) {
    const row = item.row;
    const days = row[column(H.IMPLEMENTATION_DAYS)];
    if (!isValidImplementationDays_(days) || documentWorkflowIsBlank_(days)) {
      throw new Error(
        'Invalid «' + H.IMPLEMENTATION_DAYS + '» in active rule at row ' +
        item.sheetRow + '.'
      );
    }
    return {
      id: String(row[column(H.RULE_ID)] || '').trim(),
      documentType: String(row[column(H.DOCUMENT_TYPE)] || '').trim(),
      documentStatus: String(row[column(H.DOCUMENT_STATUS)] || '').trim(),
      documentLocation: String(row[column(H.DOCUMENT_LOCATION)] || '').trim(),
      implementationDays: days,
      action: String(row[column(H.ACTION)] || '').trim(),
      sheetRow: item.sheetRow
    };
  });

  const rulesByConditions = {};
  activeRules.forEach(function (rule) {
    const key = JSON.stringify(documentWorkflowNormalizedConditions_(rule));
    if (!rulesByConditions[key]) rulesByConditions[key] = [];
    rulesByConditions[key].push(rule);
  });
  const conflicts = Object.keys(rulesByConditions).filter(function (key) {
    return rulesByConditions[key].length > 1;
  });
  if (conflicts.length) {
    const details = conflicts.map(function (key) {
      return rulesByConditions[key].map(function (rule) {
        return 'row ' + rule.sheetRow + (rule.id ? ' (ID «' + rule.id + '»)' : '');
      }).join(', ');
    });
    throw new Error(
      'Ambiguous active workflow rules with identical conditions: ' +
      details.join('; ') + '.'
    );
  }
  return activeRules;
}

function findActiveDocumentWorkflowRule_(rules, document) {
  const expected = documentWorkflowNormalizedConditions_(document);
  return (rules || []).filter(function (rule) {
    const actual = documentWorkflowNormalizedConditions_(rule);
    return actual.every(function (value, index) {
      return value === expected[index];
    });
  })[0] || null;
}

function migrateDocumentStatuses() {
  return withDocumentLock_(function () {
    return runDocumentWorkflowMigration_({ statuses: true });
  });
}

function migrateKs2Ks3DocumentType(retainedTypeId, legacyTypeName) {
  if (!String(retainedTypeId || '').trim()) {
    throw new Error('The retained KS-2 document type ID must be supplied explicitly.');
  }
  return withDocumentLock_(function () {
    return runDocumentWorkflowMigration_({
      retainedTypeId: String(retainedTypeId).trim(),
      legacyTypeName: String(legacyTypeName || 'КС-2').trim()
    });
  });
}

function migrateDocumentWorkflowData(retainedTypeId, legacyTypeName) {
  if (!String(retainedTypeId || '').trim()) {
    throw new Error('The retained KS-2 document type ID must be supplied explicitly.');
  }
  return withDocumentLock_(function () {
    return runDocumentWorkflowMigration_({
      statuses: true,
      retainedTypeId: String(retainedTypeId).trim(),
      legacyTypeName: String(legacyTypeName || 'КС-2').trim()
    });
  });
}

function buildDocumentWorkflowMigrationPlan_(rows, indexes, options, dataStartRow) {
  const output = { rows: [], changes: [], statusChanges: 0, typeChanges: 0,
    numberChanges: 0, checkedRows: 0 };
  const statusMap = {
    'Ожидает заполнения': 'На подготовке',
    'Передан заказчику': 'На согласовании у заказчика'
  };
  const escape = function (text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  };
  const legacyPattern = options.retainedTypeId
    ? new RegExp('^' + escape(options.legacyTypeName) + ' №([1-9][0-9]*)$') : null;
  rows.forEach(function (sourceRow, offset) {
    if (sourceRow.every(documentWorkflowIsBlank_)) return;
    output.checkedRows += 1;
    const row = sourceRow.slice();
    const sheetRow = dataStartRow + offset;
    const add = function (header, newValue, counter) {
      const oldValue = row[indexes[header]];
      if (oldValue === newValue) return;
      row[indexes[header]] = newValue;
      output.changes.push({ sheetRow: sheetRow, header: header,
        oldValue: oldValue, newValue: newValue, row: row });
      output[counter] += 1;
    };
    if (options.statuses) {
      const oldStatus = String(row[indexes[H.DOCUMENT_STATUS]] || '').trim();
      if (statusMap[oldStatus]) add(H.DOCUMENT_STATUS, statusMap[oldStatus], 'statusChanges');
    }
    if (options.retainedTypeId &&
        String(row[indexes[H.DOCUMENT_TYPE_ID]] || '').trim() === options.retainedTypeId) {
      const oldType = String(row[indexes[H.DOCUMENT_TYPE]] || '').trim();
      const oldNumber = String(row[indexes[H.DOCUMENT_NUMBER]] || '').trim();
      const generated = legacyPattern.exec(oldNumber);
      add(H.DOCUMENT_TYPE, 'КС2/КС3', 'typeChanges');
      if (generated) add(H.DOCUMENT_NUMBER, 'КС2/КС3 №' + generated[1], 'numberChanges');
    }
    if (output.changes.some(function (change) { return change.sheetRow === sheetRow; })) {
      output.rows.push({ sheetRow: sheetRow, row: row });
    }
  });
  return output;
}

function runDocumentWorkflowMigration_(options) {
  assertSystemSheetsStructure_(['DOCUMENTS', 'OPERATION_HISTORY', 'CHANGE_HISTORY']);
  const startedAt = new Date();
  const operationId = generateOperationId_(startedAt);
  const email = getActiveUserEmail_();
  const context = getSystemSheetContext_('DOCUMENTS');
  const rows = migrationReadRows_(context);
  const indexes = {};
  [H.DOCUMENT_ID, H.OBJECT_ID, H.DOCUMENT_STATUS, H.DOCUMENT_TYPE_ID,
    H.DOCUMENT_TYPE, H.DOCUMENT_NUMBER, H.UPDATED_AT, H.UPDATED_BY_EMAIL]
    .forEach(function (header) { indexes[header] = migrationColumn_(context, header); });
  const plan = buildDocumentWorkflowMigrationPlan_(
    rows, indexes, options, context.config.dataStartRow
  );
  const now = new Date();
  plan.rows.forEach(function (item) {
    item.row[indexes[H.UPDATED_AT]] = now;
    item.row[indexes[H.UPDATED_BY_EMAIL]] = email;
  });
  migrationWriteColumns_(context, plan.rows, [H.DOCUMENT_STATUS, H.DOCUMENT_TYPE,
    H.DOCUMENT_NUMBER, H.UPDATED_AT, H.UPDATED_BY_EMAIL]);
  writeDocumentWorkflowHistory_(context, plan, operationId, now, email, indexes);
  writeDocumentWorkflowOperation_(plan, operationId, startedAt, new Date(), email);
  return { operationId: operationId, checkedRows: plan.checkedRows,
    changedRows: plan.rows.length, changedFields: plan.changes.length,
    statusChanges: plan.statusChanges, typeChanges: plan.typeChanges,
    numberChanges: plan.numberChanges };
}

function writeDocumentWorkflowHistory_(documents, plan, operationId, now, email, indexes) {
  if (!plan.changes.length) return;
  const context = getSystemSheetContext_('CHANGE_HISTORY');
  const rows = plan.changes.map(function (change, index) {
    const values = {};
    const documentRow = change.row;
    values[H.CHANGE_ID] = generateChangeId_(operationId, index + 1);
    values[H.OPERATION_ID] = operationId;
    values[H.CHANGE_DATETIME] = now;
    values[H.CHANGED_BY_EMAIL] = email;
    values[H.ACTION_TYPE] = SYSTEM_CONFIG.VALUES.CHANGE_ACTION_EDIT;
    values[H.DOCUMENT_ID] = documentRow[indexes[H.DOCUMENT_ID]];
    values[H.OBJECT_ID] = documentRow[indexes[H.OBJECT_ID]];
    values[H.FACT_ROW_NUMBER] = change.sheetRow;
    values[H.FIELD_NAME] = change.header;
    values[H.OLD_VALUE] = change.oldValue;
    values[H.NEW_VALUE] = change.newValue;
    values[H.CHANGE_SOURCE] = SYSTEM_CONFIG.VALUES.DOCUMENT_WORKFLOW_MIGRATION_SOURCE;
    return context.headers.map(function (header) {
      return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
    });
  });
  const start = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(start, 1, rows.length, context.headers.length).setValues(rows);
}

function writeDocumentWorkflowOperation_(plan, id, startedAt, finishedAt, email) {
  const context = getSystemSheetContext_('OPERATION_HISTORY');
  const values = {};
  values[H.OPERATION_ID] = id;
  values[H.OPERATION_STARTED_AT] = startedAt;
  values[H.OPERATION_FINISHED_AT] = finishedAt;
  values[H.OPERATION_STARTED_BY] = email;
  values[H.OPERATION_SOURCE] = SYSTEM_CONFIG.VALUES.DOCUMENT_WORKFLOW_MIGRATION_SOURCE;
  values[H.OPERATION_TYPE] = SYSTEM_CONFIG.VALUES.DOCUMENT_WORKFLOW_MIGRATION_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = plan.rows.length
    ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS
    : SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES;
  values[H.DOCUMENTS_CHANGED] = plan.rows.length;
  values[H.FACT_ROWS_UPDATED] = plan.rows.length;
  values[H.FIELDS_CHANGED] = plan.changes.length;
  values[H.ERRORS_COUNT] = 0;
  values[H.EXECUTION_SECONDS] = (finishedAt.getTime() - startedAt.getTime()) / 1000;
  values[H.ERROR_TEXT] = 'Headers and columns were not changed by the migration.';
  const row = context.headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
  });
  const start = Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow);
  context.sheet.getRange(start, 1, 1, context.headers.length).setValues([row]);
}
