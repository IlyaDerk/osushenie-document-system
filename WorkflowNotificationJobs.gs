/** Read-only Stage 2 notification preview. No delivery or configuration side effects. */
const WORKFLOW_NOTIFICATION_DRY_RUN_TARGET_ = 'DRY_RUN_GROUP_TARGET';

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
        code: warning.code, message: warning.message, details: details, count: 0
      };
      byKey[key].count++;
    });
  });
  return Object.keys(byKey).sort().map(function (key) { return byKey[key]; });
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
    model: buildWorkflowNotificationModel_(sendResults)
  };
}

function dryRunWorkflowNotifications() {
  assertSystemSheetsStructure_(['DOCUMENTS', 'WORKFLOW_RULES', 'NOTIFICATION_HISTORY']);
  const now = new Date();
  const history = readWorkflowNotificationHistory_();
  return buildWorkflowNotificationDryRun_({
    documents: readWorkflowNotificationDocuments_(),
    rules: readActiveDocumentWorkflowRules_(), historyRecords: history.records,
    businessDate: workflowNotificationDateOrdinal_(now).key,
    businessAt: now, generatedAt: now, target: WORKFLOW_NOTIFICATION_DRY_RUN_TARGET_
  });
}
