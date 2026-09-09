/** Read-only notification orchestration. Telegram formatting remains pure. */
const WORKFLOW_NOTIFICATION_DRY_RUN_TARGET_ = 'DRY_RUN_GROUP_TARGET';
const WORKFLOW_NOTIFICATION_RETRY_AFTER_MINUTES_ = 60;
const WORKFLOW_NOTIFICATION_RESULT_PERSIST_ATTEMPTS_ = 3;

function readWorkflowNotificationDocuments_() {
  const context = getSystemSheetContext_('DOCUMENTS');
  const count = context.sheet.getLastRow() - context.config.dataStartRow + 1;
  if (count <= 0) return [];
  const rows = context.sheet.getRange(context.config.dataStartRow, 1, count,
    context.headers.length).getValues();
  const column = function (header) {
    return context.headerMap[sysNormalizeHeader_(header)] - 1;
  };
  return rows.map(function (row, index) {
    return {
      sheetRow: context.config.dataStartRow + index,
      documentId: row[column(H.DOCUMENT_ID)], objectId: row[column(H.OBJECT_ID)],
      objectName: row[column(H.OBJECT_NAME)], documentType: row[column(H.DOCUMENT_TYPE)],
      documentNumber: row[column(H.DOCUMENT_NUMBER)],
      documentStatus: row[column(H.DOCUMENT_STATUS)],
      documentLocation: row[column(H.DOCUMENT_LOCATION)],
      transferredAt: row[column(H.TRANSFERRED_AT)],
      implementationDays: row[column(H.IMPLEMENTATION_DAYS)],
      statusChangedAt: row[column(H.DOCUMENT_STATUS_CHANGED_AT)],
      recordStatus: row[column(H.RECORD_STATUS)]
    };
  });
}

function aggregateWorkflowNotificationWarnings_(results, runId) {
  const byKey = {};
  (results || []).forEach(function (result) {
    (result.warnings || []).forEach(function (warning) {
      const details = warning.details || {};
      const workflow = (details.workflowConditions || []).map(workflowNotificationFold_);
      const key = [notificationHistoryString_(runId), warning.code || '',
        details.ruleId || '', details.sheetRow || '', JSON.stringify(workflow)].join('|');
      if (!byKey[key]) byKey[key] = {
        code: warning.code, message: warning.message, details: details, count: 0,
        affectedDocumentIds: [], affectedObjectIds: [],
        documentIdsSeen: {}, objectIdsSeen: {}
      };
      const aggregated = byKey[key];
      aggregated.count++;
      const documentId = notificationHistoryString_(result.documentId);
      const objectId = notificationHistoryString_(result.objectId);
      if (documentId && !aggregated.documentIdsSeen[documentId]) {
        aggregated.documentIdsSeen[documentId] = true;
        aggregated.affectedDocumentIds.push(documentId);
      }
      if (objectId && !aggregated.objectIdsSeen[objectId]) {
        aggregated.objectIdsSeen[objectId] = true;
        aggregated.affectedObjectIds.push(objectId);
      }
    });
  });
  return Object.keys(byKey).sort().map(function (key) {
    const warning = byKey[key];
    warning.affectedDocumentIds.sort(workflowNotificationCompareText_);
    warning.affectedObjectIds.sort(workflowNotificationCompareText_);
    delete warning.documentIdsSeen;
    delete warning.objectIdsSeen;
    return warning;
  });
}

function buildWorkflowNotificationDryRun_(options) {
  const input = options || {}, documents = input.documents || [], rules = input.rules || [];
  const businessDate = notificationBusinessDateKey_(input.businessDate);
  const businessAt = input.businessAt instanceof Date ? input.businessAt :
    new Date(businessDate + 'T09:00:00+03:00');
  const target = notificationHistoryString_(input.target) || WORKFLOW_NOTIFICATION_DRY_RUN_TARGET_;
  const results = documents.map(function (document) {
    return evaluateWorkflowNotification_(document, rules, businessAt);
  });
  const state = foldWorkflowNotificationHistory_(input.historyRecords || []);
  const candidates = results.filter(function (result) { return result.status === 'CANDIDATE'; });
  const newEvents = [], alreadyReserved = [];
  candidates.forEach(function (candidate) {
    const key = buildNotificationBusinessKey_(candidate.documentId, businessDate,
      candidate.event, WORKFLOW_NOTIFICATION_HISTORY_.CHANNEL_TELEGRAM, target);
    const entry = { key: key, candidate: candidate };
    (state.business[key] ? alreadyReserved : newEvents).push(entry);
  });
  const sendResults = newEvents.map(function (entry) { return entry.candidate; });
  const warnings = aggregateWorkflowNotificationWarnings_(results, input.runId || 'DRY_RUN');
  const warningCounts = {}, skippedReasons = {};
  warnings.forEach(function (warning) {
    warningCounts[warning.code] = (warningCounts[warning.code] || 0) + warning.count;
  });
  results.forEach(function (result) {
    if (result.status === 'SKIPPED') skippedReasons[result.reason] =
      (skippedReasons[result.reason] || 0) + 1;
  });
  const eventCounts = { D_MINUS_7: 0, D_MINUS_3: 0, DUE_TODAY: 0, OVERDUE_DAILY: 0 };
  sendResults.forEach(function (candidate) { eventCounts[candidate.event]++; });
  const model = buildWorkflowNotificationModel_(sendResults);
  const deliveryPlan = input.telegramPreview === false ? [] : buildTelegramDeliveryPlan_(model, {
    businessDate: businessDate, target: target,
    operatorCardUrl: input.operatorCardUrl || '', testMode: false
  });
  return {
    mode: 'DRY_RUN', businessDate: businessDate,
    generatedAt: input.generatedAt instanceof Date ? input.generatedAt : new Date(),
    counts: {
      rowsRead: documents.length,
      active: documents.filter(function (document) {
        return workflowNotificationFold_(document.recordStatus) ===
          workflowNotificationFold_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS);
      }).length,
      candidatesBeforeDedup: candidates.length, wouldSend: sendResults.length,
      alreadyReserved: alreadyReserved.length,
      skipped: results.filter(function (result) { return result.status === 'SKIPPED'; }).length
    },
    eventCounts: eventCounts, warningCounts: warningCounts, warnings: warnings,
    skippedReasons: skippedReasons,
    dedup: { newEvents: newEvents, alreadyReserved: alreadyReserved },
    model: model,
    telegramPreview: {
      target: target, partCount: deliveryPlan.length,
      parts: deliveryPlan.map(function (delivery) { return {
        partNumber: delivery.partNumber, length: delivery.messageText.length,
        hash: delivery.messageHash, text: delivery.messageText
      }; })
    }
  };
}

function workflowOperatorCardUrl_() {
  const spreadsheet = getSystemSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.name);
  if (!sheet) throw new Error('Не найден лист «' + SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.name + '».');
  return spreadsheet.getUrl() + '#gid=' + sheet.getSheetId();
}

function dryRunWorkflowNotifications() {
  const config = requireWorkflowNotificationConfig_(readWorkflowNotificationConfig_(), false);
  assertSystemSheetsStructure_(['DOCUMENTS', 'WORKFLOW_RULES', 'NOTIFICATION_HISTORY']);
  const now = new Date();
  const history = readWorkflowNotificationHistory_();
  return buildWorkflowNotificationDryRun_({
    documents: readWorkflowNotificationDocuments_(),
    rules: readActiveDocumentWorkflowRules_(), historyRecords: history.records,
    businessDate: workflowNotificationDateOrdinal_(now).key,
    businessAt: now, generatedAt: now, target: config.chatId,
    operatorCardUrl: workflowOperatorCardUrl_()
  });
}

function workflowTestHistoryRecord_(delivery, runId, businessDate, result, now) {
  return buildTestDeliveryRecord_({
    runId: runId, deliveryId: delivery.deliveryId, timestamp: now,
    notificationDate: businessDate, channel: delivery.channel, target: delivery.target,
    result: result.result || result, telegramMessageId: result.telegramMessageId || '',
    errorCode: result.errorCode || '', errorText: result.errorText || '', attemptNumber: 1,
    messageHash: delivery.messageHash, messageText: delivery.messageText
  });
}

/** Manual-only real Telegram send. It never consults or writes production dedup records. */
function testSendWorkflowNotifications() {
  const config = requireWorkflowNotificationConfig_(readWorkflowNotificationConfig_(), true);
  assertSystemSheetsStructure_(['DOCUMENTS', 'WORKFLOW_RULES', 'NOTIFICATION_HISTORY']);
  const now = new Date(), businessDate = workflowNotificationDateOrdinal_(now).key;
  const rules = readActiveDocumentWorkflowRules_();
  const evaluated = readWorkflowNotificationDocuments_().map(function (document) {
    return evaluateWorkflowNotification_(document, rules, now);
  });
  const model = buildWorkflowNotificationModel_(evaluated);
  const runId = generateNotificationId_('NOTIFICATION_RUN');
  const deliveries = buildTelegramDeliveryPlan_(model, { businessDate: businessDate,
    target: config.chatId, operatorCardUrl: workflowOperatorCardUrl_(), testMode: true });
  const prepared = deliveries.map(function (delivery) {
    return workflowTestHistoryRecord_(delivery, runId, businessDate,
      WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.PREPARED, now);
  });
  appendWorkflowNotificationHistoryRecords_(prepared);
  const outcomes = deliveries.map(function (delivery) {
    return sendTelegramMessage_(config.botToken, config.chatId, delivery.messageText);
  });
  const outcomeRecords = deliveries.map(function (delivery, index) {
    return workflowTestHistoryRecord_(delivery, runId, businessDate, outcomes[index], new Date());
  });
  let historyPersisted = true, historyError = '';
  try { appendWorkflowNotificationHistoryRecords_(outcomeRecords); }
  catch (error) {
    historyPersisted = false;
    historyError = redactWorkflowNotificationSecret_(error && error.message ? error.message : error,
      config.botToken);
  }
  return {
    mode: 'TEST_SEND', runId: runId, target: config.chatId, partCount: deliveries.length,
    historyPersisted: historyPersisted, historyError: historyError,
    outcomes: outcomes.map(function (outcome, index) { return {
      deliveryId: deliveries[index].deliveryId, partNumber: index + 1,
      sent: outcome.sent, result: outcome.result,
      telegramMessageId: outcome.telegramMessageId, httpStatus: outcome.httpStatus,
      errorCode: outcome.errorCode, errorText: outcome.errorText,
      retryEligible: outcome.retryEligible, messageHash: deliveries[index].messageHash
    }; })
  };
}

function workflowProductionCandidate_(candidate, businessDate, target) {
  const copy = Object.assign({}, candidate);
  copy.notificationDate = businessDate;
  copy.channel = WORKFLOW_NOTIFICATION_HISTORY_.CHANNEL_TELEGRAM;
  copy.target = notificationHistoryString_(target);
  return copy;
}

function workflowReservationWarning_(candidate) {
  return (candidate.warnings || []).filter(function (warning) {
    return warning.code === 'WORKFLOW_RULE_MISSING' || warning.code === 'WORKFLOW_ACTION_MISSING';
  }).map(function (warning) { return warning.code; }).sort().join(', ');
}

function buildWorkflowProductionPlan_(accepted, input) {
  const deliveries = buildTelegramDeliveryPlan_(buildWorkflowNotificationModel_(accepted), {
    businessDate: input.businessDate, target: input.target,
    operatorCardUrl: input.operatorCardUrl, testMode: false
  });
  const referenceDelivery = {};
  deliveries.forEach(function (delivery) {
    delivery.references.forEach(function (candidate) {
      const key = buildNotificationBusinessKey_(candidate.documentId,
        candidate.notificationDate, candidate.event, candidate.channel, candidate.target);
      if (referenceDelivery[key]) throw new Error('Business candidate appears in two deliveries.');
      referenceDelivery[key] = delivery.deliveryId;
    });
  });
  const reservations = accepted.map(function (candidate) {
    const key = buildNotificationBusinessKey_(candidate.documentId,
      candidate.notificationDate, candidate.event, candidate.channel, candidate.target);
    if (!referenceDelivery[key]) throw new Error('Business candidate has no physical delivery.');
    return {
      runId: input.runId, deliveryId: referenceDelivery[key], documentId: candidate.documentId,
      objectId: candidate.objectId, timestamp: input.now,
      notificationDate: candidate.notificationDate, event: candidate.event,
      controlDate: candidate.controlDateKey, documentStatus: candidate.documentStatus,
      documentLocation: candidate.location, channel: candidate.channel, target: candidate.target,
      configWarning: workflowReservationWarning_(candidate)
    };
  });
  return {
    businessReservations: reservations,
    deliveries: deliveries.map(function (delivery) { return {
      runId: input.runId, deliveryId: delivery.deliveryId, timestamp: input.now,
      notificationDate: input.businessDate, channel: delivery.channel, target: delivery.target,
      attemptNumber: 1, messageHash: delivery.messageHash, messageText: delivery.messageText
    }; })
  };
}

function workflowDeliveryResult_(delivery, outcome, recordId, attempt, now, botToken) {
  return buildDeliveryResultRecord_({
    recordId: recordId, runId: delivery.runId, deliveryId: delivery.deliveryId,
    timestamp: now, notificationDate: delivery.notificationDate,
    channel: delivery.channel, target: delivery.target, attemptNumber: attempt,
    result: outcome.result, telegramMessageId: outcome.telegramMessageId || '',
    errorCode: outcome.errorCode || '',
    errorText: redactWorkflowNotificationSecret_(outcome.errorText || '', botToken),
    messageHash: delivery.messageHash
  });
}

function persistWorkflowResultLocally_(record, botToken) {
  let errorText = '';
  for (let attempt = 1; attempt <= WORKFLOW_NOTIFICATION_RESULT_PERSIST_ATTEMPTS_; attempt++) {
    try {
      persistWorkflowNotificationResultIdempotently_(record);
      return { persisted: true, attempts: attempt, error: '' };
    } catch (error) {
      errorText = redactWorkflowNotificationSecret_(
        error && error.message ? error.message : error, botToken);
    }
  }
  return { persisted: false, attempts: WORKFLOW_NOTIFICATION_RESULT_PERSIST_ATTEMPTS_,
    error: errorText };
}

function workflowSendPreparedDeliveries_(deliveries, config, attempt) {
  const outcomes = [], unresolved = [];
  (deliveries || []).forEach(function (delivery) {
    const resultRecordId = generateNotificationId_('NOTIFICATION_HISTORY');
    const outcome = sendTelegramMessage_(config.botToken, config.chatId, delivery.messageText);
    const record = workflowDeliveryResult_(delivery, outcome, resultRecordId,
      attempt, new Date(), config.botToken);
    const persistence = persistWorkflowResultLocally_(record, config.botToken);
    const item = {
      deliveryId: delivery.deliveryId, result: outcome.result,
      telegramMessageId: outcome.telegramMessageId || '', historyPersisted: persistence.persisted,
      historyAttempts: persistence.attempts,
      historyError: persistence.persisted ? '' : persistence.error
    };
    outcomes.push(item);
    if (!persistence.persisted) unresolved.push(item);
  });
  return { outcomes: outcomes, unresolved: unresolved };
}

function workflowConfigWarningRecords_(evaluated, runId, now, businessDate) {
  return aggregateWorkflowNotificationWarnings_(evaluated, runId).filter(function (warning) {
    return warning.code === 'INVALID_RULE_NOTIFICATION_PERMISSION';
  }).map(function (warning) {
    return buildConfigWarningRecord_({ runId: runId, timestamp: now,
      notificationDate: businessDate,
      configWarning: warning.code + ': ' + warning.message + ' (documents=' + warning.count + ')' });
  });
}

/** Scheduled production entry point. Disabled is an immediate property-only no-op. */
function runWorkflowNotificationsDaily() {
  const config = readWorkflowNotificationConfig_();
  if (config.enabled !== true) return { mode: 'PRODUCTION', status: 'DISABLED' };
  requireWorkflowNotificationConfig_(config, true);
  assertSystemSheetsStructure_(['DOCUMENTS', 'WORKFLOW_RULES', 'NOTIFICATION_HISTORY']);
  const now = new Date(), businessDate = workflowNotificationDateOrdinal_(now).key;
  const rules = readActiveDocumentWorkflowRules_();
  const evaluated = readWorkflowNotificationDocuments_().map(function (document) {
    return evaluateWorkflowNotification_(document, rules, now);
  });
  const candidates = evaluated.filter(function (result) {
    return result.status === 'CANDIDATE';
  }).map(function (candidate) {
    return workflowProductionCandidate_(candidate, businessDate, config.chatId);
  });
  const runId = generateNotificationId_('NOTIFICATION_RUN');
  const buildInput = { runId: runId, businessDate: businessDate, target: config.chatId,
    operatorCardUrl: workflowOperatorCardUrl_(), now: now };
  const reserved = reserveWorkflowNotificationPlan_({ runId: runId,
    businessCandidates: candidates }, function (accepted) {
    return buildWorkflowProductionPlan_(accepted, buildInput);
  });
  const warningRecords = workflowConfigWarningRecords_(evaluated, runId, now, businessDate);
  let warningsPersisted = true, warningError = '';
  if (warningRecords.length) {
    try { appendWorkflowNotificationHistoryRecords_(warningRecords); }
    catch (error) {
      warningsPersisted = false;
      warningError = redactWorkflowNotificationSecret_(error && error.message ? error.message : error,
        config.botToken);
    }
  }
  const sending = workflowSendPreparedDeliveries_(reserved.deliveries, config, 1);
  return { mode: 'PRODUCTION', status: 'COMPLETED', runId: runId,
    businessDate: businessDate, candidates: candidates.length,
    reserved: reserved.businessReservations.length, alreadyReserved: reserved.alreadyReserved.length,
    deliveries: reserved.deliveries.length, outcomes: sending.outcomes,
    historyUnresolved: sending.unresolved, warningsPersisted: warningsPersisted,
    warningError: warningError };
}

function workflowRetryCandidates_(records, now) {
  const byDelivery = {};
  (records || []).forEach(function (record) {
    const type = notificationHistoryString_(record.recordType).toUpperCase();
    if (type !== WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_PREPARED &&
        type !== WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_RESULT) return;
    const id = notificationHistoryString_(record.deliveryId);
    if (!id) return;
    if (!byDelivery[id]) byDelivery[id] = { deliveryId: id, prepared1: null,
      technical1: null, attempt2: false, terminal: false };
    const item = byDelivery[id], attempt = Number(record.attemptNumber);
    if (attempt === 2) item.attempt2 = true;
    if (type === WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_PREPARED && attempt === 1) {
      item.prepared1 = item.prepared1 || record;
    }
    if (type === WORKFLOW_NOTIFICATION_HISTORY_.TYPES.DELIVERY_RESULT) {
      if (attempt === 1 && record.result === WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.FAILED_TECHNICAL) {
        item.technical1 = item.technical1 || record;
      }
      if ([WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.SENT,
        WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.UNKNOWN_DELIVERY_OUTCOME,
        WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.FAILED_PERMANENT,
        WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.FAILED_CONFIGURATION,
        WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.SKIPPED_STALE_BEFORE_RETRY].indexOf(record.result) >= 0) {
        item.terminal = true;
      }
    }
  });
  return Object.keys(byDelivery).map(function (id) { return byDelivery[id]; }).filter(function (item) {
    if (!item.prepared1 || !item.technical1 || item.attempt2 || item.terminal) return false;
    const at = item.technical1.timestamp;
    if (!(at instanceof Date) || isNaN(at.getTime())) return false;
    item.failureAt = at;
    return now.getTime() - at.getTime() >= WORKFLOW_NOTIFICATION_RETRY_AFTER_MINUTES_ * 60000;
  }).sort(function (left, right) {
    return left.failureAt.getTime() - right.failureAt.getTime() ||
      workflowNotificationCompareText_(left.deliveryId, right.deliveryId);
  });
}

function workflowRetryStaleReasons_(reservations, documents, rules) {
  if (!reservations.length) return ['RESERVATIONS_MISSING'];
  const byId = {};
  documents.forEach(function (document) {
    const id = notificationHistoryString_(document.documentId);
    if (!byId[id]) byId[id] = [];
    byId[id].push(document);
  });
  const reasons = [];
  reservations.forEach(function (reservation) {
    const matches = byId[notificationHistoryString_(reservation.documentId)] || [];
    if (!matches.length) { reasons.push(reservation.documentId + ':MISSING_DOCUMENT'); return; }
    if (matches.length > 1) { reasons.push(reservation.documentId + ':DUPLICATE_DOCUMENT_ID'); return; }
    const document = matches[0], prefix = reservation.documentId + ':';
    if (workflowNotificationFold_(document.recordStatus) !== workflowNotificationFold_(WORKFLOW_NOTIFICATION_.ACTIVE_RECORD)) reasons.push(prefix + 'INACTIVE');
    if (workflowNotificationFold_(document.documentStatus) === workflowNotificationFold_(WORKFLOW_NOTIFICATION_.TERMINAL_STATUS)) reasons.push(prefix + 'TERMINAL');
    if (workflowNotificationFold_(document.documentLocation) !== workflowNotificationFold_(WORKFLOW_NOTIFICATION_.OFFICE_LOCATION)) reasons.push(prefix + 'LOCATION');
    if (workflowNotificationFold_(document.documentStatus) !== workflowNotificationFold_(reservation.documentStatus)) reasons.push(prefix + 'STATUS_CHANGED');
    if (documentWorkflowIsBlank_(document.transferredAt)) reasons.push(prefix + 'TRANSFER_BLANK');
    else if (!isValidTransferredAt_(document.transferredAt)) reasons.push(prefix + 'TRANSFER_INVALID');
    if (!documentWorkflowIsBlank_(document.statusChangedAt) && !isValidTransferredAt_(document.statusChangedAt)) reasons.push(prefix + 'STATUS_DATE_INVALID');
    if (isValidTransferredAt_(document.transferredAt) && !documentWorkflowIsBlank_(document.transferredAt) &&
        isValidTransferredAt_(document.statusChangedAt) && !documentWorkflowIsBlank_(document.statusChangedAt) &&
        workflowNotificationDateOrdinal_(document.transferredAt).day < workflowNotificationDateOrdinal_(document.statusChangedAt).day) reasons.push(prefix + 'CYCLE_NOT_STARTED');
    const normalized = documentWorkflowNormalizedConditions_(document), matching = (rules || []).filter(function (rule) {
      const ruleConditions = documentWorkflowNormalizedConditions_(rule);
      return ruleConditions.every(function (value, index) { return value === normalized[index]; });
    });
    if (matching.length > 1) reasons.push(prefix + 'RULE_AMBIGUOUS');
    if (matching.length === 1) {
      const permission = workflowNotificationFold_(matching[0].notify);
      if (permission === workflowNotificationFold_(WORKFLOW_NOTIFICATION_.NO)) reasons.push(prefix + 'RULE_NOTIFY_NO');
      else if (permission !== workflowNotificationFold_(WORKFLOW_NOTIFICATION_.YES)) reasons.push(prefix + 'RULE_NOTIFY_INVALID');
    }
  });
  return reasons;
}

/** Shared worker: only clear technical attempt 1 failures can reserve attempt 2. */
function retryWorkflowNotificationDeliveries() {
  const config = readWorkflowNotificationConfig_();
  if (config.enabled !== true) return { mode: 'RETRY', status: 'DISABLED' };
  requireWorkflowNotificationConfig_(config, true);
  assertSystemSheetsStructure_(['DOCUMENTS', 'WORKFLOW_RULES', 'NOTIFICATION_HISTORY']);
  const now = new Date(), history = readWorkflowNotificationHistory_();
  const due = workflowRetryCandidates_(history.records, now);
  const documents = due.length ? readWorkflowNotificationDocuments_() : [];
  const rules = due.length ? readActiveDocumentWorkflowRules_({ allowAmbiguous: true }) : [];
  const summary = { mode: 'RETRY', status: 'COMPLETED', scanned: Object.keys(
    foldWorkflowNotificationHistory_(history.records).deliveries).length,
    due: due.length, retried: 0, sent: 0, staleSkipped: 0, technicalFailed: 0,
    permanentFailed: 0, configurationFailed: 0, unknown: 0,
    historyUnresolved: 0, notEligible: 0, outcomes: [] };
  due.forEach(function (item) {
    const reservations = history.records.filter(function (record) {
      return record.recordType === WORKFLOW_NOTIFICATION_HISTORY_.TYPES.BUSINESS_RESERVATION &&
        notificationHistoryString_(record.deliveryId) === item.deliveryId;
    });
    const stale = workflowRetryStaleReasons_(reservations, documents, rules);
    if (stale.length) {
      const staleRecord = buildDeliveryResultRecord_({ runId: item.prepared1.runId,
        deliveryId: item.deliveryId, timestamp: new Date(),
        notificationDate: item.prepared1.notificationDate,
        channel: item.prepared1.channel, target: item.prepared1.target, attemptNumber: 2,
        result: WORKFLOW_NOTIFICATION_HISTORY_.RESULTS.SKIPPED_STALE_BEFORE_RETRY,
        errorCode: 'STALE_BEFORE_RETRY', errorText: stale.join('; '),
        messageHash: item.prepared1.messageHash });
      const skipped = reserveWorkflowNotificationRetryAttempt_(staleRecord);
      if (skipped.reserved) summary.staleSkipped++; else summary.notEligible++;
      return;
    }
    const prepared2 = buildDeliveryPreparedRecord_({ runId: item.prepared1.runId,
      deliveryId: item.deliveryId, timestamp: new Date(),
      notificationDate: item.prepared1.notificationDate,
      channel: item.prepared1.channel, target: item.prepared1.target, attemptNumber: 2,
      messageHash: item.prepared1.messageHash, messageText: item.prepared1.messageText });
    const reservation = reserveWorkflowNotificationRetryAttempt_(prepared2);
    if (!reservation.reserved) { summary.notEligible++; return; }
    summary.retried++;
    const sending = workflowSendPreparedDeliveries_([prepared2], config, 2);
    const outcome = sending.outcomes[0];
    summary.outcomes.push(outcome);
    if (!outcome.historyPersisted) summary.historyUnresolved++;
    if (outcome.result === 'SENT') summary.sent++;
    else if (outcome.result === 'FAILED_TECHNICAL') summary.technicalFailed++;
    else if (outcome.result === 'FAILED_PERMANENT') summary.permanentFailed++;
    else if (outcome.result === 'FAILED_CONFIGURATION') summary.configurationFailed++;
    else summary.unknown++;
  });
  summary.notEligible += Math.max(0, summary.scanned - summary.due);
  return summary;
}

const WORKFLOW_NOTIFICATION_TRIGGER_HANDLERS_ = Object.freeze([
  'runWorkflowNotificationsDaily', 'retryWorkflowNotificationDeliveries'
]);

function removeWorkflowNotificationTriggers() {
  const removed = [];
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    const handler = trigger.getHandlerFunction();
    if (WORKFLOW_NOTIFICATION_TRIGGER_HANDLERS_.indexOf(handler) >= 0) {
      ScriptApp.deleteTrigger(trigger);
      removed.push(handler);
    }
  });
  return { removed: removed.length, handlers: removed };
}

function setupWorkflowNotificationTriggers() {
  const removed = removeWorkflowNotificationTriggers();
  ScriptApp.newTrigger('runWorkflowNotificationsDaily').timeBased().atHour(8)
    .nearMinute(0).everyDays(1).inTimezone(WORKFLOW_NOTIFICATION_.TIMEZONE).create();
  ScriptApp.newTrigger('retryWorkflowNotificationDeliveries').timeBased()
    .everyMinutes(15).create();
  return { status: 'CONFIGURED', removed: removed.removed, created: 2,
    handlers: WORKFLOW_NOTIFICATION_TRIGGER_HANDLERS_.slice() };
}
