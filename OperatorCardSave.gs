/** Safe partial write-back workflow for «Карточка операциониста». */
const OPERATOR_CARD_SAVE_WRITABLE_ = [
  H.DOCUMENT_NUMBER, H.DOCUMENT_DATE, H.DOCUMENT_STATUS, H.ORIGINAL_EDO, H.COMMENT,
  H.DOCUMENT_HOLDER, H.DOCUMENT_LOCATION, H.TRANSFERRED_BY, H.PAID,
  H.CUSTOMER_SIGNING_RESPONSIBLE, H.DOCUMENT_AMOUNT, H.GU_FLAG, H.GU_TERMS,
  H.RECORD_STATUS
];

function operatorCardSaveEqual_(left, right) {
  if (left instanceof Date || right instanceof Date) {
    return left instanceof Date && right instanceof Date && left.getTime() === right.getTime();
  }
  return left === right;
}

function operatorCardSaveEmpty_(value) {
  return value === '' || value == null;
}

function operatorCardSaveLabelMap_(items) {
  const map = {};
  items.forEach(function (item) {
    map[operatorCardDisplayLabel_(item.name, item.id)] = item;
  });
  return map;
}

function operatorCardSaveAssertDictionary_(value, allowed, header) {
  if (operatorCardSaveEmpty_(value)) return;
  if (!allowed.some(function (item) {
    return operatorCardFold_(item) === operatorCardFold_(value);
  })) {
    throw new Error('Поле «' + header + '»: неизвестное значение «' + value + '». Выберите значение из актуального справочника.');
  }
}

function operatorCardSaveVersion_(value) {
  return value instanceof Date && !isNaN(value.getTime())
    ? value.getTime()
    : String(value == null ? '' : value);
}

function operatorCardSaveError_(card, factRow, message) {
  return {
    cardRow: card.sheetRow,
    factRow: Number(factRow) || null,
    documentId: operatorCardNormalizeText_(card.values[H.DOCUMENT_ID]),
    message: message
  };
}

function operatorCardSaveProposeRow_(card, fact, dictionaries) {
  const holderMap = operatorCardSaveLabelMap_(dictionaries.holders);
  const employeeMap = operatorCardSaveLabelMap_(dictionaries.employees);
  const dictionaryHeaders = [
    H.DOCUMENT_STATUS, H.ORIGINAL_EDO, H.DOCUMENT_LOCATION,
    H.PAID, H.GU_FLAG, H.RECORD_STATUS
  ];
  const proposed = {};

  OPERATOR_CARD_SAVE_WRITABLE_.forEach(function (header) {
    let value = card.values[header];
    if (header === H.DOCUMENT_DATE && !operatorCardSaveEmpty_(value) &&
        (!(value instanceof Date) || isNaN(value.getTime()))) {
      throw new Error('Поле «' + header + '» должно быть пустым или корректной датой.');
    }
    if (header === H.DOCUMENT_AMOUNT && !operatorCardSaveEmpty_(value) &&
        (typeof value !== 'number' || !isFinite(value))) {
      throw new Error('Поле «' + header + '» должно быть пустым или числом.');
    }
    if (dictionaryHeaders.indexOf(header) !== -1) {
      operatorCardSaveAssertDictionary_(value, dictionaries.card[header] || [], header);
    }
    if (header === H.DOCUMENT_HOLDER) {
      if (operatorCardSaveEmpty_(value)) {
        value = '';
        proposed[H.HOLDER_EMPLOYEE_ID] = '';
      } else if (holderMap[value]) {
        proposed[H.HOLDER_EMPLOYEE_ID] = holderMap[value].id;
        value = holderMap[value].name;
      } else if (!fact.values[H.HOLDER_EMPLOYEE_ID] && value === fact.values[H.DOCUMENT_HOLDER]) {
        proposed[H.HOLDER_EMPLOYEE_ID] = '';
      } else {
        throw new Error('Держатель не найден в актуальных справочниках сотрудников и клиентов.');
      }
    }
    if (header === H.TRANSFERRED_BY) {
      if (operatorCardSaveEmpty_(value)) {
        value = '';
        proposed[H.TRANSFERRED_BY_EMPLOYEE_ID] = '';
      } else if (employeeMap[value]) {
        proposed[H.TRANSFERRED_BY_EMPLOYEE_ID] = employeeMap[value].id;
        value = employeeMap[value].name;
      } else if (!fact.values[H.TRANSFERRED_BY_EMPLOYEE_ID] && value === fact.values[H.TRANSFERRED_BY]) {
        proposed[H.TRANSFERRED_BY_EMPLOYEE_ID] = '';
      } else {
        throw new Error('«Кто передал» должен быть выбран из актуального справочника сотрудников.');
      }
    }
    proposed[header] = value;
  });
  return proposed;
}

/** Builds independent row plans; local failures are returned, never thrown. */
function operatorCardBuildSavePlan_(cardItems, factItems, dictionaries, now, email, operationId) {
  const factByRow = {};
  const activeGroups = {};
  factItems.forEach(function (item) {
    factByRow[item.sheetRow] = item;
    if (operatorCardFold_(item.values[H.RECORD_STATUS]) ===
        operatorCardFold_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS)) {
      const documentId = operatorCardNormalizeText_(item.values[H.DOCUMENT_ID]);
      if (documentId) (activeGroups[documentId] = activeGroups[documentId] || []).push(item.sheetRow);
    }
  });

  const rowPlans = [];
  const rowErrors = [];
  const cardRowsByDocument = {};
  cardItems.forEach(function (card) {
    const factRow = Number(card.values[H.FACT_ROW_NUMBER]);
    try {
      if (!Number.isInteger(factRow)) {
        throw new Error('Отсутствует корректный скрытый «' + H.FACT_ROW_NUMBER + '».');
      }
      const fact = factByRow[factRow];
      if (!fact) throw new Error('Физическая строка факта не существует. Нажмите «Применить» заново.');

      [H.DOCUMENT_ID, H.OBJECT_ID, H.DOCUMENT_TYPE_ID].forEach(function (header) {
        if (!operatorCardSaveEqual_(card.values[header], fact.values[header])) {
          throw new Error('Строка факта больше не соответствует полю «' + header + '».');
        }
      });

      const documentId = operatorCardNormalizeText_(fact.values[H.DOCUMENT_ID]);
      (cardRowsByDocument[documentId] = cardRowsByDocument[documentId] || []).push(factRow);
      const preliminaryChanged = OPERATOR_CARD_SAVE_WRITABLE_.some(function (header) {
        let factValue = fact.values[header];
        if (header === H.DOCUMENT_HOLDER) {
          factValue = operatorCardDisplayLabel_(factValue, fact.values[H.HOLDER_EMPLOYEE_ID]);
        } else if (header === H.TRANSFERRED_BY) {
          factValue = operatorCardDisplayLabel_(factValue, fact.values[H.TRANSFERRED_BY_EMPLOYEE_ID]);
        }
        return !operatorCardSaveEqual_(card.values[header], factValue);
      });
      const changedReadOnlyHeaders = SYSTEM_CONFIG.CARD_FIELD_MAP.filter(function (mapping) {
        return !mapping.editable && !mapping.technical &&
          mapping.cardHeader !== H.UPDATED_AT && mapping.cardHeader !== H.UPDATED_BY_EMAIL &&
          !operatorCardSaveEqual_(card.values[mapping.cardHeader], fact.values[mapping.factHeader]);
      }).map(function (mapping) {
        return mapping.cardHeader;
      });
      if (!preliminaryChanged && !changedReadOnlyHeaders.length) return;
      if (changedReadOnlyHeaders.length) {
        throw new Error('Изменено поле только для чтения «' + changedReadOnlyHeaders.join('», «') + '».');
      }

      const proposed = operatorCardSaveProposeRow_(card, fact, dictionaries);
      const changes = [];
      Object.keys(proposed).forEach(function (header) {
        if (!operatorCardSaveEqual_(proposed[header], fact.values[header])) {
          changes.push({ header: header, oldValue: fact.values[header], newValue: proposed[header] });
        }
      });

      // An untouched row never prevents another row from being saved.
      if (!changes.length) return;

      if (operatorCardFold_(fact.values[H.RECORD_STATUS]) !==
          operatorCardFold_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS)) {
        throw new Error('Строка факта уже не активна. Нажмите «Применить» заново.');
      }
      if (operatorCardSaveVersion_(card.values[H.UPDATED_AT]) !==
          operatorCardSaveVersion_(fact.values[H.UPDATED_AT])) {
        throw new Error('Данные изменились после загрузки. Нажмите «Применить» или исправьте строку после обновления.');
      }
      const statusChange = changes.some(function (change) { return change.header === H.RECORD_STATUS; });
      const businessChange = changes.some(function (change) { return change.header !== H.RECORD_STATUS; });
      if (statusChange) {
        const status = proposed[H.RECORD_STATUS];
        const allowed = [
          SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS,
          SYSTEM_CONFIG.VALUES.ARCHIVED_RECORD_STATUS,
          SYSTEM_CONFIG.VALUES.DELETED_RECORD_STATUS
        ];
        if (!allowed.some(function (value) { return operatorCardFold_(value) === operatorCardFold_(status); })) {
          throw new Error('Недопустимый переход «Статус записи»: Активная → ' + status + '.');
        }
      }
      if (statusChange && businessChange) {
        throw new Error('Нельзя одновременно менять «Статус записи» и другие поля.');
      }

      rowPlans.push({
        cardRow: card.sheetRow,
        sheetRow: factRow,
        documentId: documentId,
        objectId: fact.values[H.OBJECT_ID],
        changes: changes
      });
    } catch (error) {
      rowErrors.push(operatorCardSaveError_(card, factRow, error.message || String(error)));
    }
  });

  const groupErrors = [];
  const touchedIds = {};
  rowPlans.forEach(function (plan) { touchedIds[plan.documentId] = true; });
  Object.keys(touchedIds).forEach(function (documentId) {
    const activeRows = activeGroups[documentId] || [];
    if (activeRows.length < 2) return;
    const plans = rowPlans.filter(function (plan) { return plan.documentId === documentId; });
    const present = (cardRowsByDocument[documentId] || []).slice().sort(function (a, b) { return a - b; });
    const expected = activeRows.slice().sort(function (a, b) { return a - b; });
    let message = '';
    if (present.join(',') !== expected.join(',')) {
      message = 'Дубль представлен в карточке не полностью.';
    } else if (plans.some(function (plan) {
      return plan.changes.some(function (change) { return change.header !== H.RECORD_STATUS; });
    })) {
      message = 'Для дубля разрешено менять только «Статус записи».';
    } else {
      const changedByRow = {};
      plans.forEach(function (plan) {
        const change = plan.changes.filter(function (item) { return item.header === H.RECORD_STATUS; })[0];
        if (change) changedByRow[plan.sheetRow] = change.newValue;
      });
      const remaining = activeRows.filter(function (row) {
        return !changedByRow[row] || operatorCardFold_(changedByRow[row]) ===
          operatorCardFold_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS);
      }).length;
      if (remaining > 1) message = 'Дубль не устранён: после save останется ' + remaining + ' активных строк.';
    }
    if (message) {
      groupErrors.push({ documentId: documentId, factRows: expected, message: message });
      for (let index = rowPlans.length - 1; index >= 0; index--) {
        if (rowPlans[index].documentId === documentId) rowPlans.splice(index, 1);
      }
    }
  });

  const changes = [];
  let sequence = 0;
  rowPlans.forEach(function (plan) {
    plan.changes.forEach(function (change) {
      changes.push({
        changeId: generateChangeId_(operationId, ++sequence),
        operationId: operationId,
        changedAt: now,
        userEmail: email,
        documentId: plan.documentId,
        objectId: plan.objectId,
        factRow: plan.sheetRow,
        fieldName: change.header,
        oldValue: change.oldValue,
        newValue: change.newValue
      });
    });
  });
  return {
    rows: rowPlans,
    changes: changes,
    rowErrors: rowErrors,
    groupErrors: groupErrors,
    duplicateIds: Object.keys(touchedIds).filter(function (id) {
      return (activeGroups[id] || []).length > 1;
    })
  };
}

function operatorCardSaveRead_(context) {
  const count = Math.max(context.sheet.getLastRow() - context.config.dataStartRow + 1, 0);
  if (!count) return [];
  const rows = context.sheet.getRange(
    context.config.dataStartRow, 1, count, context.headers.length
  ).getValues();
  const result = [];
  rows.forEach(function (row, offset) {
    if (!row.some(function (value) { return !operatorCardSaveEmpty_(value); })) return;
    const values = {};
    context.headers.forEach(function (header, index) { values[header] = row[index]; });
    result.push({ sheetRow: context.config.dataStartRow + offset, values: values });
  });
  return result;
}

function operatorCardSaveDictionaries_() {
  const filterData = operatorCardGetFilterData_();
  const card = {};
  [H.DOCUMENT_STATUS, H.ORIGINAL_EDO, H.DOCUMENT_LOCATION, H.PAID, H.GU_FLAG, H.RECORD_STATUS]
    .forEach(function (header) {
      card[header] = operatorCardReadUniqueColumn_('CARD_DICTIONARY', header);
    });
  [SYSTEM_CONFIG.VALUES.ARCHIVED_RECORD_STATUS, SYSTEM_CONFIG.VALUES.DELETED_RECORD_STATUS]
    .forEach(function (required) {
      if (!card[H.RECORD_STATUS].some(function (value) {
        return operatorCardFold_(value) === operatorCardFold_(required);
      })) {
        throw new Error('В «Справочник для КО» → «Статус записи» отсутствует обязательное значение «' + required + '». Добавьте его вручную.');
      }
    });
  return { holders: filterData.holders, employees: filterData.employees, card: card };
}

function operatorCardSaveWriteColumnGroups_(sheet, column, items) {
  const groups = [];
  items.sort(function (left, right) { return left.row - right.row; }).forEach(function (item) {
    const group = groups[groups.length - 1];
    if (group && item.row === group.startRow + group.values.length) {
      group.values.push([item.value]);
    } else {
      groups.push({ startRow: item.row, values: [[item.value]] });
    }
  });
  groups.forEach(function (group) {
    sheet.getRange(group.startRow, column, group.values.length, 1).setValues(group.values);
  });
}

function operatorCardSaveWriteFacts_(context, plan, now, email) {
  plan.rows.forEach(function (rowPlan) {
    rowPlan.changes.push({ header: H.UPDATED_AT, newValue: now });
    rowPlan.changes.push({ header: H.UPDATED_BY_EMAIL, newValue: email });
    if (rowPlan.changes.some(function (change) { return change.header === H.DOCUMENT_STATUS; })) {
      rowPlan.changes.push({ header: H.DOCUMENT_STATUS_CHANGED_AT, newValue: now });
    }
  });
  const headers = OPERATOR_CARD_SAVE_WRITABLE_.concat([
    H.HOLDER_EMPLOYEE_ID, H.TRANSFERRED_BY_EMPLOYEE_ID,
    H.UPDATED_AT, H.UPDATED_BY_EMAIL, H.DOCUMENT_STATUS_CHANGED_AT
  ]);
  headers.forEach(function (header) {
    const items = [];
    plan.rows.forEach(function (rowPlan) {
      const change = rowPlan.changes.filter(function (item) { return item.header === header; })[0];
      if (change) items.push({ row: rowPlan.sheetRow, value: change.newValue });
    });
    if (items.length) {
      operatorCardSaveWriteColumnGroups_(
        context.sheet, context.headerMap[sysNormalizeHeader_(header)], items
      );
    }
  });
}

function operatorCardSaveSyncSuccessfulCardRows_(cardContext, plan, now, email) {
  const systemValues = {};
  systemValues[H.UPDATED_AT] = now;
  systemValues[H.UPDATED_BY_EMAIL] = email;
  [H.UPDATED_AT, H.UPDATED_BY_EMAIL].forEach(function (header) {
    const items = plan.rows.map(function (rowPlan) {
      return { row: rowPlan.cardRow, value: systemValues[header] };
    });
    if (items.length) {
      operatorCardSaveWriteColumnGroups_(
        cardContext.sheet, cardContext.headerMap[sysNormalizeHeader_(header)], items
      );
    }
  });
}

function operatorCardSaveWriteChanges_(context, changes) {
  if (!changes.length) return;
  const rows = changes.map(function (change) {
    const values = {};
    values[H.CHANGE_ID] = change.changeId;
    values[H.OPERATION_ID] = change.operationId;
    values[H.CHANGE_DATETIME] = change.changedAt;
    values[H.CHANGED_BY_EMAIL] = change.userEmail;
    values[H.ACTION_TYPE] = SYSTEM_CONFIG.VALUES.CHANGE_ACTION_EDIT;
    values[H.DOCUMENT_ID] = change.documentId;
    values[H.OBJECT_ID] = change.objectId;
    values[H.FACT_ROW_NUMBER] = change.factRow;
    values[H.FIELD_NAME] = change.fieldName;
    values[H.OLD_VALUE] = change.oldValue;
    values[H.NEW_VALUE] = change.newValue;
    values[H.CHANGE_SOURCE] = SYSTEM_CONFIG.VALUES.OPERATOR_CARD_SOURCE;
    return context.headers.map(function (header) {
      return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
    });
  });
  context.sheet.getRange(
    Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow),
    1, rows.length, context.headers.length
  ).setValues(rows);
}

function operatorCardSaveWriteOperation_(context, operation) {
  const values = {};
  values[H.OPERATION_ID] = operation.id;
  values[H.OPERATION_STARTED_AT] = operation.started;
  values[H.OPERATION_FINISHED_AT] = operation.finished;
  values[H.OPERATION_STARTED_BY] = operation.email;
  values[H.OPERATION_SOURCE] = SYSTEM_CONFIG.VALUES.OPERATOR_CARD_SOURCE;
  values[H.OPERATION_TYPE] = SYSTEM_CONFIG.VALUES.OPERATOR_CARD_SAVE_OPERATION_TYPE;
  values[H.OPERATION_STATUS] = operation.status;
  values[H.DOCUMENTS_LOADED] = 0;
  values[H.DOCUMENTS_CHANGED] = operation.documents;
  values[H.FACT_ROWS_UPDATED] = operation.rows;
  values[H.FIELDS_CHANGED] = operation.fields;
  values[H.DUPLICATE_IDS_FOUND] = operation.duplicates;
  values[H.ERRORS_COUNT] = operation.errors;
  values[H.EXECUTION_SECONDS] = (operation.finished.getTime() - operation.started.getTime()) / 1000;
  values[H.ERROR_TEXT] = operation.comment || '';
  const row = context.headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '';
  });
  context.sheet.getRange(
    Math.max(context.sheet.getLastRow() + 1, context.config.dataStartRow),
    1, 1, context.headers.length
  ).setValues([row]);
}

function operatorCardSaveProblemComment_(rowErrors, groupErrors) {
  return rowErrors.map(function (error) {
    return (error.documentId || 'строка карточки ' + error.cardRow) +
      (error.factRow ? ', факт ' + error.factRow : '') + ': ' + error.message;
  }).concat(groupErrors.map(function (error) {
    return error.documentId + ', факты ' + error.factRows.join(', ') + ': ' + error.message;
  })).join(' | ');
}

function operatorCardSaveStatus_(savedRows, localErrors, postWriteWarning) {
  if (savedRows > 0) {
    return localErrors > 0 || postWriteWarning
      ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS
      : SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS;
  }
  return localErrors > 0
    ? SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR
    : SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES;
}

function operatorCardSave_(filters) {
  const started = new Date();
  let result;
  result = withDocumentLock_(function () {
    let operationId = '';
    let operationContext = null;
    let state = { documents: 0, rows: 0, fields: 0, duplicates: 0, errors: 0 };
    let factWriteStarted = false;
    let factWriteCompleted = false;
    const email = getActiveUserEmail_();
    try {
      // All required sources and destinations are validated before the first fact write.
      const documents = getSystemSheetContext_('DOCUMENTS');
      const card = getSystemSheetContext_('OPERATOR_CARD');
      getSystemSheetContext_('CARD_DICTIONARY');
      getSystemSheetContext_('EMPLOYEES');
      getSystemSheetContext_('CLIENTS');
      const changeContext = getSystemSheetContext_('CHANGE_HISTORY');
      operationContext = getSystemSheetContext_('OPERATION_HISTORY');
      operatorCardValidateHeaders_(
        documents.headers, card.headers, documents.config.name, card.config.name,
        documents.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)],
        card.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)]
      );
      const dictionaries = operatorCardSaveDictionaries_();
      operationId = generateOperationId_(started);
      const now = new Date();
      const plan = operatorCardBuildSavePlan_(
        operatorCardSaveRead_(card), operatorCardSaveRead_(documents),
        dictionaries, now, email, operationId
      );
      const changedIds = {};
      plan.rows.forEach(function (rowPlan) { changedIds[rowPlan.documentId] = true; });
      state = {
        documents: Object.keys(changedIds).length,
        rows: plan.rows.length,
        fields: plan.changes.length,
        duplicates: plan.duplicateIds.length,
        errors: plan.rowErrors.length + plan.groupErrors.length
      };

      factWriteStarted = true;
      operatorCardSaveWriteFacts_(documents, plan, now, email);
      factWriteCompleted = true;
      let postWriteWarning = '';
      try {
        operatorCardSaveWriteChanges_(changeContext, plan.changes);
      } catch (historyError) {
        postWriteWarning = 'Данные документов сохранены, но «История изменений» не записана: ' +
          (historyError.message || String(historyError));
      }

      const hasSaved = state.rows > 0;
      const hasRejected = state.errors > 0;
      let status = operatorCardSaveStatus_(state.rows, state.errors, postWriteWarning);
      const problemComment = operatorCardSaveProblemComment_(plan.rowErrors, plan.groupErrors);
      const comment = [problemComment, postWriteWarning].filter(function (value) { return value; }).join(' | ');
      try {
        operatorCardSaveWriteOperation_(operationContext, {
          id: operationId, started: started, finished: new Date(), email: email,
          status: status, documents: state.documents, rows: state.rows,
          fields: state.fields, duplicates: state.duplicates,
          errors: state.errors + (postWriteWarning ? 1 : 0), comment: comment
        });
      } catch (operationError) {
        postWriteWarning = (postWriteWarning ? postWriteWarning + ' ' : '') +
          'Данные документов сохранены, но «История операций» не записана: ' +
          (operationError.message || String(operationError));
        if (hasSaved) status = SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS;
      }

      // Partial save keeps rejected rows intact and advances versions only for saved card rows.
      if (hasSaved && hasRejected) operatorCardSaveSyncSuccessfulCardRows_(card, plan, now, email);
      return {
        success: hasSaved || !hasRejected,
        status: status,
        operationId: operationId,
        documentsChanged: state.documents,
        factRowsUpdated: state.rows,
        fieldsChanged: state.fields,
        duplicateIdsCount: state.duplicates,
        rejectedCount: state.errors,
        rowErrors: plan.rowErrors,
        groupErrors: plan.groupErrors,
        warnings: postWriteWarning ? [postWriteWarning] : [],
        partial: hasSaved && hasRejected,
        needsFullRefresh: !hasRejected
      };
    } catch (error) {
      const message = error.message || String(error);
      const uncertainWriteMessage = 'Во время записи данных произошла ошибка. ' +
        'Часть изменений могла быть записана. Не продолжайте редактирование до повторной загрузки карточки.';
      // An ID may only be reserved while the document lock is held.
      if (!operationId) {
        try { operationId = generateOperationId_(started); } catch (ignoredIdError) {}
      }
      if (!operationContext) {
        try { operationContext = getSystemSheetContext_('OPERATION_HISTORY'); } catch (ignoredContextError) {}
      }
      if (operationContext && operationId) {
        try {
          operatorCardSaveWriteOperation_(operationContext, {
            id: operationId, started: started, finished: new Date(), email: email,
            status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,
            documents: state.documents, rows: state.rows, fields: state.fields,
            duplicates: state.duplicates, errors: Math.max(1, state.errors),
            comment: factWriteStarted && !factWriteCompleted
              ? uncertainWriteMessage + ' Причина: ' + message
              : (factWriteCompleted
                ? 'Данные документов сохранены, но save завершился ошибкой: ' + message
                : message)
          });
        } catch (ignoredOperationError) {}
      }
      if (factWriteStarted && !factWriteCompleted) {
        throw new Error(uncertainWriteMessage + ' Причина: ' + message);
      }
      if (factWriteCompleted) {
        return {
          success: true,
          status: SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS,
          operationId: operationId,
          documentsChanged: state.documents,
          factRowsUpdated: state.rows,
          fieldsChanged: state.fields,
          duplicateIdsCount: state.duplicates,
          rejectedCount: state.errors,
          rowErrors: [], groupErrors: [], partial: true, needsFullRefresh: false,
          warnings: ['Данные документов сохранены, но save завершился ошибкой: ' + message]
        };
      }
      throw error;
    }
  });

  if (result.needsFullRefresh) {
    try {
      operatorCardApply_(filters, true);
      result.refreshed = true;
    } catch (refreshError) {
      result.refreshed = false;
      result.warnings.push('Данные сохранены, но карточку не удалось обновить. Нажмите «Применить» вручную: ' +
        (refreshError.message || String(refreshError)));
      if (result.factRowsUpdated) result.status = SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS;
    }
  } else {
    result.refreshed = false;
  }
  result.message = 'Сохранено документов: ' + result.documentsChanged + '. ' +
    (result.rejectedCount ? 'Не сохранено: ' + result.rejectedCount + '. ' : '') +
    'Строк: ' + result.factRowsUpdated + ', полей: ' + result.fieldsChanged + '.';
  return result;
}
