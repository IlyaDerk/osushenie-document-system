/** Pure workflow-notification selection and report-model helpers. */
const WORKFLOW_NOTIFICATION_ = Object.freeze({
  TIMEZONE: 'Europe/Moscow',
  ACTIVE_RECORD: SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS,
  TERMINAL_STATUS: 'Подписан с обеих сторон',
  OFFICE_LOCATION: 'В офисе',
  YES: 'Да',
  NO: 'Нет',
  ACTION_NOT_CONFIGURED: 'не настроено'
});

function workflowNotificationFold_(value) {
  return documentWorkflowNormalizedConditions_({
    documentType: value,
    documentStatus: '',
    documentLocation: ''
  })[0];
}

function workflowNotificationDateOrdinal_(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) {
    throw new Error('A valid Date is required for Moscow calendar conversion.');
  }
  const key = Utilities.formatDate(
    value, WORKFLOW_NOTIFICATION_.TIMEZONE, 'yyyy-MM-dd'
  );
  const parts = key.split('-').map(Number);
  return {
    key: key,
    day: Date.UTC(parts[0], parts[1] - 1, parts[2]) / 86400000
  };
}

function workflowNotificationDateKeyFromDay_(day) {
  return new Date(day * 86400000).toISOString().slice(0, 10);
}

function workflowNotificationWarning_(code, message, details) {
  return { code: code, message: message, details: details || {} };
}

function workflowNotificationSkip_(reason, document, warnings) {
  return {
    status: 'SKIPPED',
    reason: reason,
    documentId: String(document.documentId || '').trim(),
    warnings: warnings || []
  };
}

function workflowNotificationEvent_(daysRemaining) {
  if (daysRemaining === 7) return 'D_MINUS_7';
  if (daysRemaining === 3) return 'D_MINUS_3';
  if (daysRemaining === 0) return 'DUE_TODAY';
  if (daysRemaining < 0) return 'OVERDUE_DAILY';
  return null;
}

function evaluateWorkflowNotification_(document, activeRules, businessAt) {
  const warnings = [];
  const fold = workflowNotificationFold_;
  if (fold(document.recordStatus) !== fold(WORKFLOW_NOTIFICATION_.ACTIVE_RECORD)) {
    return workflowNotificationSkip_('SKIPPED_INACTIVE_RECORD', document);
  }
  if (fold(document.documentStatus) === fold(WORKFLOW_NOTIFICATION_.TERMINAL_STATUS)) {
    return workflowNotificationSkip_('SKIPPED_TERMINAL_STATUS', document);
  }
  if (fold(document.documentLocation) !== fold(WORKFLOW_NOTIFICATION_.OFFICE_LOCATION)) {
    return workflowNotificationSkip_('SKIPPED_DOCUMENT_NOT_IN_OFFICE', document);
  }
  if (documentWorkflowIsBlank_(document.transferredAt)) {
    return workflowNotificationSkip_('SKIPPED_TRANSFER_DATE_BLANK', document);
  }
  if (!isValidTransferredAt_(document.transferredAt)) {
    warnings.push(workflowNotificationWarning_(
      'INVALID_TRANSFER_DATE', 'Некорректное значение «' + H.TRANSFERRED_AT + '».'
    ));
    return workflowNotificationSkip_(
      'SKIPPED_INVALID_TRANSFER_DATE', document, warnings
    );
  }

  const transferDate = workflowNotificationDateOrdinal_(document.transferredAt);
  if (!documentWorkflowIsBlank_(document.statusChangedAt)) {
    if (!isValidTransferredAt_(document.statusChangedAt)) {
      warnings.push(workflowNotificationWarning_(
        'INVALID_STATUS_CHANGE_DATE',
        'Некорректное значение «' + H.DOCUMENT_STATUS_CHANGED_AT + '».'
      ));
      return workflowNotificationSkip_(
        'SKIPPED_INVALID_STATUS_CHANGE_DATE', document, warnings
      );
    }
    const statusDate = workflowNotificationDateOrdinal_(document.statusChangedAt);
    if (transferDate.day < statusDate.day) {
      warnings.push(workflowNotificationWarning_(
        'WORKFLOW_CYCLE_NOT_STARTED',
        'Требуется указать "' + H.TRANSFERRED_AT + '" для текущего статуса'
      ));
      return workflowNotificationSkip_(
        'SKIPPED_WORKFLOW_CYCLE_NOT_STARTED', document, warnings
      );
    }
  }

  const rule = findActiveDocumentWorkflowRule_(activeRules, document);
  if (rule) {
    const permission = fold(rule.notify);
    if (permission === fold(WORKFLOW_NOTIFICATION_.NO)) {
      return workflowNotificationSkip_(
        'SKIPPED_RULE_NOTIFICATIONS_DISABLED', document
      );
    }
    if (permission !== fold(WORKFLOW_NOTIFICATION_.YES)) {
      warnings.push(workflowNotificationWarning_(
        'INVALID_RULE_NOTIFICATION_PERMISSION',
        'В active rule поле «' + H.NOTIFY + '» должно быть «Да» или «Нет».',
        {
          ruleId: rule.id,
          sheetRow: rule.sheetRow,
          invalidValue: rule.notify,
          workflowConditions: documentWorkflowNormalizedConditions_(rule).slice()
        }
      ));
      return workflowNotificationSkip_(
        'SKIPPED_INVALID_RULE_NOTIFICATION_PERMISSION', document, warnings
      );
    }
  } else {
    warnings.push(workflowNotificationWarning_(
      'WORKFLOW_RULE_MISSING',
      'Для workflow-state не найдено active rule.',
      { workflowConditions: documentWorkflowNormalizedConditions_(document).slice() }
    ));
  }

  let action = rule ? String(rule.action || '').trim() : '';
  if (!action) {
    action = WORKFLOW_NOTIFICATION_.ACTION_NOT_CONFIGURED;
    if (rule) {
      warnings.push(workflowNotificationWarning_(
        'WORKFLOW_ACTION_MISSING', 'В active rule не настроено «' + H.ACTION + '».',
        { ruleId: rule.id, sheetRow: rule.sheetRow }
      ));
    }
  }

  const effectiveDays = getEffectiveImplementationDays_(
    document.implementationDays, rule
  );
  const businessDate = workflowNotificationDateOrdinal_(businessAt);
  const controlDay = transferDate.day + effectiveDays;
  const daysRemaining = controlDay - businessDate.day;
  const event = workflowNotificationEvent_(daysRemaining);
  if (!event) {
    return workflowNotificationSkip_('SKIPPED_NO_EVENT_TODAY', document, warnings);
  }
  return {
    status: 'CANDIDATE',
    documentId: String(document.documentId || '').trim(),
    objectId: String(document.objectId || '').trim(),
    objectName: String(document.objectName || '').trim(),
    documentType: String(document.documentType || '').trim(),
    documentNumber: String(document.documentNumber || '').trim(),
    documentStatus: String(document.documentStatus || '').trim(),
    location: String(document.documentLocation || '').trim(),
    transferredAt: document.transferredAt,
    statusChangedAt: document.statusChangedAt,
    effectiveDays: effectiveDays,
    controlDateKey: workflowNotificationDateKeyFromDay_(controlDay),
    daysRemaining: daysRemaining,
    event: event,
    action: action,
    warnings: warnings
  };
}


function workflowNotificationCompareText_(left, right) {
  return String(left || '').localeCompare(String(right || ''), 'ru', {
    numeric: true,
    sensitivity: 'base'
  });
}

function buildWorkflowNotificationModel_(results) {
  const candidates = (results || []).filter(function (result) {
    return result && result.status === 'CANDIDATE';
  }).slice().sort(function (left, right) {
    return workflowNotificationCompareText_(left.objectName, right.objectName) ||
      workflowNotificationCompareText_(left.objectId, right.objectId) ||
      workflowNotificationCompareText_(left.documentType, right.documentType) ||
      workflowNotificationCompareText_(left.documentNumber, right.documentNumber) ||
      workflowNotificationCompareText_(left.documentId, right.documentId);
  });
  const objects = [];
  const byKey = {};
  candidates.forEach(function (candidate) {
    const key = candidate.objectId
      ? 'id:' + workflowNotificationFold_(candidate.objectId)
      : 'name:' + workflowNotificationFold_(candidate.objectName);
    if (!byKey[key]) {
      byKey[key] = {
        objectId: candidate.objectId,
        objectName: candidate.objectName,
        documents: []
      };
      objects.push(byKey[key]);
    }
    byKey[key].documents.push(candidate);
  });
  return { candidateCount: candidates.length, objects: objects };
}
