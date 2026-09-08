/** Append-only repository and pure state model for workflow notifications. */
const WORKFLOW_NOTIFICATION_HISTORY_ = Object.freeze({
  CHANNEL_TELEGRAM: 'TELEGRAM',
  TYPES: Object.freeze({
    BUSINESS_RESERVATION: 'BUSINESS_RESERVATION',
    DELIVERY_PREPARED: 'DELIVERY_PREPARED',
    DELIVERY_RESULT: 'DELIVERY_RESULT',
    CONFIG_WARNING: 'CONFIG_WARNING',
    TEST_DELIVERY: 'TEST_DELIVERY'
  }),
  RESULTS: Object.freeze({
    RESERVED: 'RESERVED', PREPARED: 'PREPARED', SENT: 'SENT',
    FAILED_TECHNICAL: 'FAILED_TECHNICAL',
    FAILED_PERMANENT: 'FAILED_PERMANENT',
    FAILED_CONFIGURATION: 'FAILED_CONFIGURATION',
    UNKNOWN_DELIVERY_OUTCOME: 'UNKNOWN_DELIVERY_OUTCOME',
    SKIPPED_STALE_BEFORE_RETRY: 'SKIPPED_STALE_BEFORE_RETRY'
  })
});

function notificationHistoryString_(value) {
  return String(value == null ? '' : value).trim();
}

function notificationBusinessDateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return workflowNotificationDateOrdinal_(value).key;
  }
  const text = notificationHistoryString_(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    throw new Error('Notification business date must have yyyy-MM-dd format.');
  }
  const parts = text.split('-').map(Number);
  if (new Date(Date.UTC(parts[0], parts[1] - 1, parts[2])).toISOString().slice(0, 10) !== text) {
    throw new Error('Notification business date is invalid.');
  }
  return text;
}

function buildNotificationBusinessKey_(documentId, businessDate, event, channel, target) {
  const values = [
    notificationHistoryString_(documentId),
    notificationBusinessDateKey_(businessDate),
    notificationHistoryString_(event).toUpperCase(),
    notificationHistoryString_(channel).toUpperCase(),
    notificationHistoryString_(target)
  ];
  if (values.some(function (value) { return !value; })) {
    throw new Error('All business notification key values are required.');
  }
  return values.join('|');
}

function generateNotificationId_(prefixKey) {
  const prefix = SYSTEM_CONFIG.ID_PREFIXES[prefixKey];
  if (!prefix) throw new Error('Unknown notification ID prefix: ' + prefixKey);
  return prefix + Utilities.getUuid();
}

function notificationHistoryBaseRecord_(input, type, result) {
  const source = input || {};
  return {
    recordId: notificationHistoryString_(source.recordId) || generateNotificationId_('NOTIFICATION_HISTORY'),
    runId: notificationHistoryString_(source.runId), recordType: type,
    deliveryId: notificationHistoryString_(source.deliveryId),
    documentId: notificationHistoryString_(source.documentId),
    objectId: notificationHistoryString_(source.objectId),
    timestamp: source.timestamp instanceof Date ? source.timestamp : new Date(),
    notificationDate: source.notificationDate ? notificationBusinessDateKey_(source.notificationDate) : '',
    event: notificationHistoryString_(source.event).toUpperCase(),
    controlDate: source.controlDate || '', documentStatus: source.documentStatus || '',
    documentLocation: source.documentLocation || '',
    channel: notificationHistoryString_(source.channel).toUpperCase(),
    target: notificationHistoryString_(source.target), result: result,
    telegramMessageId: notificationHistoryString_(source.telegramMessageId),
    errorCode: notificationHistoryString_(source.errorCode), errorText: source.errorText || '',
    attemptNumber: source.attemptNumber === '' ? '' : source.attemptNumber,
    configWarning: source.configWarning || '', messageHash: source.messageHash || '',
    messageText: source.messageText || ''
  };
}

function buildBusinessReservationRecord_(input) {
  const record = notificationHistoryBaseRecord_(input,
    WORKFLOW_NOTIFICATION_HISTORY_.TYPES.BUSINESS_RESERVATION,
    WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.RESERVED);
  record.businessKey = buildNotificationBusinessKey_(record.documentId,
    record.notificationDate, record.event, record.channel, record.target);
  return record;
}

function buildDeliveryPreparedRecord_(input) {
  return notificationHistoryBaseRecord_(input,
    WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_PREPARED,
    WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.PREPARED);
}

function buildDeliveryResultRecord_(input) {
  return notificationHistoryBaseRecord_(input,
    WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_RESULT,
    notificationHistoryString_((input || {}).result).toUpperCase());
}

function buildConfigWarningRecord_(input) {
  return notificationHistoryBaseRecord_(input,
    WORKFLOW_NOTIFICATION_HISTORY_.TYPES.CONFIG_WARNING,
    WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.FAILED_CONFIGURATION);
}

function buildTestDeliveryRecord_(input) {
  return notificationHistoryBaseRecord_(input,
    WORKFLOW_NOTIFICATION_HISTORY_.TYPES.TEST_DELIVERY,
    notificationHistoryString_((input || {}).result).toUpperCase() ||
      WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.PREPARED);
}

const NOTIFICATION_HISTORY_FIELDS_ = Object.freeze([
  [H.NOTIFICATION_RECORD_ID, 'recordId'], [H.NOTIFICATION_RUN_ID, 'runId'],
  [H.NOTIFICATION_RECORD_TYPE, 'recordType'], [H.DELIVERY_ID, 'deliveryId'],
  [H.DOCUMENT_ID, 'documentId'], [H.OBJECT_ID, 'objectId'],
  [H.NOTIFICATION_DATETIME, 'timestamp'], [H.NOTIFICATION_DATE, 'notificationDate'],
  [H.NOTIFICATION_EVENT, 'event'], [H.CONTROL_DATE, 'controlDate'],
  [H.DOCUMENT_STATUS, 'documentStatus'], [H.DOCUMENT_LOCATION, 'documentLocation'],
  [H.CHANNEL, 'channel'], [H.RECIPIENT, 'target'], [H.RESULT, 'result'],
  [H.TELEGRAM_MESSAGE_ID, 'telegramMessageId'], [H.ERROR_CODE, 'errorCode'],
  [H.NOTIFICATION_ERROR_TEXT, 'errorText'], [H.ATTEMPT_NUMBER, 'attemptNumber'],
  [H.CONFIG_WARNING, 'configWarning'], [H.MESSAGE_HASH, 'messageHash'],
  [H.MESSAGE_TEXT, 'messageText']
]);

function notificationHistoryRecordToRow_(record, context) {
  const row = context.headers.map(function () { return ''; });
  NOTIFICATION_HISTORY_FIELDS_.forEach(function (mapping) {
    const column = context.headerMap[sysNormalizeHeader_(mapping[0])];
    if (!column) throw new Error('Missing notification history header: ' + mapping[0]);
    row[column - 1] = record[mapping[1]] == null ? '' : record[mapping[1]];
  });
  return row;
}

function readWorkflowNotificationHistory_() {
  const context = getSystemSheetContext_('NOTIFICATION_HISTORY');
  const count = context.sheet.getLastRow() - context.config.dataStartRow + 1;
  if (count <= 0) return { context: context, records: [] };
  const rows = context.sheet.getRange(context.config.dataStartRow, 1, count,
    context.headers.length).getValues();
  const records = rows.map(function (row) {
    const record = {};
    NOTIFICATION_HISTORY_FIELDS_.forEach(function (mapping) {
      record[mapping[1]] = row[context.headerMap[sysNormalizeHeader_(mapping[0])] - 1];
    });
    return record;
  }).filter(function (record) { return notificationHistoryString_(record.recordId); });
  return { context: context, records: records };
}

function foldWorkflowNotificationHistory_(records) {
  const state = { business: {}, deliveries: {} };
  (records || []).forEach(function (record) {
    const type = notificationHistoryString_(record.recordType).toUpperCase();
    if (type === WORKFLOW_NOTIFICATION_HISTORY_.TYPES.BUSINESS_RESERVATION) {
      const key = buildNotificationBusinessKey_(record.documentId, record.notificationDate,
        record.event, record.channel, record.target);
      if (!state.business[key]) state.business[key] = {
        reserved: true, reservationRecord: record,
        deliveryId: notificationHistoryString_(record.deliveryId)
      };
      return;
    }
    if (type !== WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_PREPARED &&
        type !== WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_RESULT) return;
    const id = notificationHistoryString_(record.deliveryId);
    if (!id) return;
    if (!state.deliveries[id]) state.deliveries[id] = {
      deliveryId: id, preparedAttempts: [], results: [], finalSent: false,
      failedTechnical: false, failedPermanent: false, unknownOutcome: false,
      latestAttempt: 0, retryAlreadyAttempted: false
    };
    const delivery = state.deliveries[id];
    const attempt = Number(record.attemptNumber) || 0;
    delivery.latestAttempt = Math.max(delivery.latestAttempt, attempt);
    delivery.retryAlreadyAttempted = delivery.latestAttempt > 1;
    if (type === WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_PREPARED) {
      delivery.preparedAttempts.push(record);
    } else {
      delivery.results.push(record);
      delivery.finalSent = delivery.finalSent || record.result === WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.SENT;
      delivery.failedTechnical = delivery.failedTechnical || record.result === WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.FAILED_TECHNICAL;
      delivery.failedPermanent = delivery.failedPermanent || record.result === WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.FAILED_PERMANENT;
      delivery.unknownOutcome = delivery.unknownOutcome || record.result === WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.UNKNOWN_DELIVERY_OUTCOME;
    }
  });
  return state;
}

function validateAtomicNotificationPlan_(plan) {
  if (!plan || !notificationHistoryString_(plan.runId)) throw new Error('Notification run ID is required.');
  const reservations = plan.businessReservations || [], deliveries = plan.deliveries || [];
  const records = reservations.concat(deliveries), ids = {}, deliveryIds = {},
    deliveryReferences = {}, keys = {};
  deliveries.forEach(function (record) {
    if (record.recordType !== WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_PREPARED) throw new Error('Only prepared deliveries are allowed.');
    if (!record.deliveryId || deliveryIds[record.deliveryId]) throw new Error('Delivery IDs must be filled and unique.');
    deliveryIds[record.deliveryId] = true;
    deliveryReferences[record.deliveryId] = 0;
    if (!Number.isInteger(Number(record.attemptNumber)) || Number(record.attemptNumber) < 1) throw new Error('Attempt number must be a positive integer.');
    if (record.channel !== WORKFLOW_NOTIFICATION_HISTORY_.CHANNEL_TELEGRAM || !record.target) {
      throw new Error('Prepared delivery channel and target are required.');
    }
    if (!record.messageHash || !notificationHistoryString_(record.messageText)) {
      throw new Error('Prepared delivery message hash and text are required.');
    }
  });
  reservations.forEach(function (record) {
    if (record.recordType !== WORKFLOW_NOTIFICATION_HISTORY_.TYPES.BUSINESS_RESERVATION) throw new Error('TEST records cannot be mixed into a production reservation plan.');
    if (!record.documentId || !deliveryIds[record.deliveryId]) throw new Error('Every reservation must reference an existing delivery.');
    if (record.channel !== WORKFLOW_NOTIFICATION_HISTORY_.CHANNEL_TELEGRAM || !record.target ||
        ['D_MINUS_7', 'D_MINUS_3', 'DUE_TODAY', 'OVERDUE_DAILY'].indexOf(record.event) < 0) {
      throw new Error('Reservation event, channel, and target must be canonical.');
    }
    deliveryReferences[record.deliveryId]++;
    const key = record.businessKey || buildNotificationBusinessKey_(record.documentId,
      record.notificationDate, record.event, record.channel, record.target);
    if (keys[key]) throw new Error('Duplicate business key in notification plan: ' + key);
    keys[key] = true;
  });
  records.forEach(function (record) {
    if (record.runId !== plan.runId) throw new Error('Every record must use the plan run ID.');
    if (!record.recordId || ids[record.recordId]) throw new Error('History record IDs must be filled and unique.');
    ids[record.recordId] = true;
  });
  Object.keys(deliveryReferences).forEach(function (deliveryId) {
    if (!deliveryReferences[deliveryId]) {
      throw new Error('A prepared delivery cannot be empty: ' + deliveryId);
    }
  });
  return plan;
}

/**
 * Atomically rechecks dedup, invokes a pure builder for accepted candidates,
 * validates the final plan and appends reservations + prepared deliveries once.
 */
function reserveWorkflowNotificationPlan_(input, deliveryBuilder) {
  if (!input || !notificationHistoryString_(input.runId)) throw new Error('Notification run ID is required.');
  if (typeof deliveryBuilder !== 'function') throw new TypeError('A pure delivery builder is required.');
  const candidates = input.businessCandidates || [], preliminary = {};
  candidates.forEach(function (candidate) {
    const key = buildNotificationBusinessKey_(candidate.documentId, candidate.notificationDate,
      candidate.event, candidate.channel, candidate.target);
    if (preliminary[key]) throw new Error('Duplicate business key in input: ' + key);
    preliminary[key] = true;
  });
  return withDocumentLock_(function () {
    const history = readWorkflowNotificationHistory_();
    const folded = foldWorkflowNotificationHistory_(history.records);
    const accepted = [], alreadyReserved = [];
    candidates.forEach(function (candidate) {
      const key = buildNotificationBusinessKey_(candidate.documentId, candidate.notificationDate,
        candidate.event, candidate.channel, candidate.target);
      (folded.business[key] ? alreadyReserved : accepted).push(candidate);
    });
    const built = deliveryBuilder(accepted.slice()) || {};
    const plan = {
      runId: notificationHistoryString_(input.runId),
      businessReservations: (built.businessReservations || []).map(buildBusinessReservationRecord_),
      deliveries: (built.deliveries || []).map(buildDeliveryPreparedRecord_)
    };
    validateAtomicNotificationPlan_(plan);
    if (plan.businessReservations.length !== accepted.length) {
      throw new Error('Delivery builder must map every accepted business event exactly once.');
    }
    const acceptedKeys = {};
    accepted.forEach(function (item) { acceptedKeys[buildNotificationBusinessKey_(item.documentId,
      item.notificationDate, item.event, item.channel, item.target)] = true; });
    plan.businessReservations.forEach(function (record) {
      if (!acceptedKeys[record.businessKey]) throw new Error('Delivery builder returned an unaccepted business event.');
    });
    const records = plan.businessReservations.concat(plan.deliveries);
    if (records.length) {
      const rows = records.map(function (record) {
        return notificationHistoryRecordToRow_(record, history.context);
      });
      const start = Math.max(history.context.sheet.getLastRow() + 1,
        history.context.config.dataStartRow);
      history.context.sheet.getRange(start, 1, rows.length,
        history.context.headers.length).setValues(rows);
    }
    return { runId: plan.runId, accepted: accepted, alreadyReserved: alreadyReserved,
      businessReservations: plan.businessReservations, deliveries: plan.deliveries };
  });
}

/** Appends non-reservation lifecycle records in one contiguous batch. */
function appendWorkflowNotificationHistoryRecords_(records) {
  const supplied = records || [];
  if (!supplied.length) return [];
  supplied.forEach(function (record) {
    if ([WORKFLOW_NOTIFICATION_HISTORY_.TYPES.BUSINESS_RESERVATION,
      WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_PREPARED].indexOf(record.recordType) >= 0) {
      throw new Error('Pre-send records must use reserveWorkflowNotificationPlan_.');
    }
  });
  return withDocumentLock_(function () {
    const history = readWorkflowNotificationHistory_(), ids = {};
    history.records.forEach(function (record) { ids[notificationHistoryString_(record.recordId)] = true; });
    supplied.forEach(function (record) {
      if (!record.recordId || ids[record.recordId]) throw new Error('History record IDs must be unique.');
      ids[record.recordId] = true;
    });
    const rows = supplied.map(function (record) {
      return notificationHistoryRecordToRow_(record, history.context);
    });
    const start = Math.max(history.context.sheet.getLastRow() + 1,
      history.context.config.dataStartRow);
    history.context.sheet.getRange(start, 1, rows.length,
      history.context.headers.length).setValues(rows);
    return supplied.slice();
  });
}
