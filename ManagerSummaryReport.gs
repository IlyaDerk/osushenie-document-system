/** Generated, read-only report «Сводная руководителя». */
const MANAGER_SUMMARY_ = Object.freeze({
  SHEET_NAME: 'Сводная руководителя',
  TITLE: 'Сводная руководителя',
  SIGNED: 'Подписан с обеих сторон',
  FIX: 'Требует исправления',
  ORPHANS: '⚠ Документы без найденного объекта',
  ACTIONS: Object.freeze([
    'Просрочено', 'Требует исправления', 'Срок сегодня',
    'Срок до 14 дней', 'Без плановой даты', 'Нет документов', 'Не завершено'
  ])
});

/** Validates, reads, plans and publishes one snapshot under one lock. */
function rebuildManagerSummary() {
  return withDocumentLock_(function () {
    assertSystemSheetsStructure_(['OBJECTS', 'DOCUMENTS', 'CARD_DICTIONARY']);
    const spreadsheet = getSystemSpreadsheet_();
    const timezone = spreadsheet.getSpreadsheetTimeZone();
    const snapshotAt = new Date();
    const objectRows = objectCardReadRows_('OBJECTS', [
      H.OBJECT_ID, H.OBJECT_NAME, H.RESPONSIBLE_FOREMAN, H.OBJECT_STATUS
    ]);
    const documentRows = objectCardReadRows_('DOCUMENTS', [
      H.DOCUMENT_ID, H.OBJECT_ID, H.DOCUMENT_STATUS, H.RECORD_STATUS,
      H.WORK_END_PLAN
    ]);
    const dictionaryRows = objectCardReadRows_(
      'CARD_DICTIONARY', [H.DOCUMENT_STATUS]
    );
    const cardModel = objectCardBuildModel_(
      objectRows, documentRows, dictionaryRows, snapshotAt, timezone
    );
    const model = managerSummaryBuildModel_(cardModel, objectRows);
    managerSummaryValidate_(model);
    return managerSummaryPublish_(spreadsheet, model);
  });
}

function managerSummaryDay_(date, timezone) {
  return Date.parse(Utilities.formatDate(date, timezone, 'yyyy-MM-dd') + 'T00:00:00Z') / 86400000;
}

function managerSummaryDeadline_(document, today, timezone) {
  const value = document.values[H.WORK_END_PLAN];
  if (value === '' || value == null) return { bucket: 'missing', date: null };
  if (Object.prototype.toString.call(value) !== '[object Date]' ||
      isNaN(value.getTime())) {
    throw new Error('Лист «Документы объектов», строка ' + document.sourceRow +
      ': некорректное значение «' + H.WORK_END_PLAN + '».');
  }
  const day = managerSummaryDay_(value, timezone);
  return { date: value, day: day,
    bucket: day < today ? 'overdue' : (day <= today + 14 ? 'soon' : 'later') };
}

function managerSummaryMetrics_(documents, statuses, today, timezone) {
  const signedKey = objectCardKey_(MANAGER_SUMMARY_.SIGNED);
  const fixKey = objectCardKey_(MANAGER_SUMMARY_.FIX);
  const result = { total: documents.length, signed: 0, unsigned: 0, fix: 0,
    overdue: 0, soon: 0, missing: 0, today: 0, nearest: null,
    deadlineAssignments: 0 };
  documents.forEach(function (document) {
    const status = objectCardKey_(document.values[H.DOCUMENT_STATUS]);
    // Date is part of the source contract even when the signed document is
    // deliberately excluded from every deadline calculation.
    const deadline = managerSummaryDeadline_(document, today, timezone);
    if (status === signedKey) { result.signed++; return; }
    result.unsigned++;
    if (status === fixKey) result.fix++;
    if (deadline.date && (!result.nearest || deadline.date < result.nearest)) {
      result.nearest = deadline.date;
    }
    if (deadline.bucket === 'overdue') result.overdue++;
    if (deadline.bucket === 'soon') {
      result.soon++;
      if (deadline.day === today) result.today++;
    }
    if (deadline.bucket === 'missing') result.missing++;
    if (deadline.bucket !== 'later') result.deadlineAssignments++;
  });
  result.readiness = result.total ? result.signed / result.total : 0;
  result.action = result.overdue ? MANAGER_SUMMARY_.ACTIONS[0] :
    result.fix ? MANAGER_SUMMARY_.ACTIONS[1] :
    result.today ? MANAGER_SUMMARY_.ACTIONS[2] :
    result.soon ? MANAGER_SUMMARY_.ACTIONS[3] :
    result.missing ? MANAGER_SUMMARY_.ACTIONS[4] :
    !result.total ? MANAGER_SUMMARY_.ACTIONS[5] : MANAGER_SUMMARY_.ACTIONS[6];
  return result;
}

/** Pure management projection over the canonical Object Card model. */
function managerSummaryBuildModel_(cardModel, objectRows) {
  const metadata = {};
  objectRows.forEach(function (row) {
    metadata[objectCardKey_(row.values[H.OBJECT_ID])] = {
      foreman: objectCardText_(row.values[H.RESPONSIBLE_FOREMAN]),
      status: objectCardText_(row.values[H.OBJECT_STATUS])
    };
  });
  const today = managerSummaryDay_(cardModel.snapshotAt, cardModel.timezone);
  const objects = cardModel.objects.map(function (object) {
    const metrics = managerSummaryMetrics_(
      object.documents, cardModel.statuses, today, cardModel.timezone
    );
    const meta = metadata[objectCardKey_(object.id)];
    return { id: object.id, name: object.name, foreman: meta.foreman,
      status: meta.status, metrics: metrics, problem: metrics.unsigned > 0 || !metrics.total };
  });
  const orphan = managerSummaryMetrics_(
    cardModel.orphans, cardModel.statuses, today, cardModel.timezone
  );
  const global = managerSummaryMetrics_(
    cardModel.objects.reduce(function (all, object) {
      return all.concat(object.documents);
    }, []).concat(cardModel.orphans), cardModel.statuses, today, cardModel.timezone
  );
  const statusCounts = cardModel.globalStats.counts.slice();
  const priority = {};
  MANAGER_SUMMARY_.ACTIONS.forEach(function (action, index) { priority[action] = index; });
  const problems = objectCardStableSort_(objects.filter(function (object) {
    return object.problem;
  }), function (left, right) {
    const action = priority[left.metrics.action] - priority[right.metrics.action];
    if (action) return action;
    const ld = left.metrics.nearest, rd = right.metrics.nearest;
    if (ld && rd && ld.getTime() !== rd.getTime()) return ld - rd;
    if (ld && !rd) return -1;
    if (!ld && rd) return 1;
    return objectCardNaturalCompare_(left.id, right.id);
  });
  return { snapshotAt: cardModel.snapshotAt, timezone: cardModel.timezone,
    statuses: cardModel.statuses.slice(), statusCounts: statusCounts,
    objects: objects, problems: problems, orphan: orphan,
    hasOrphans: cardModel.orphans.length > 0, global: global,
    objectsWithoutDocuments: objects.filter(function (o) { return !o.metrics.total; }).length,
    cardGlobalTotal: cardModel.globalStats.total,
    cardAssignedTotal: objects.reduce(function (n, o) { return n + o.metrics.total; }, 0) + orphan.total };
}

function managerSummaryValidate_(model) {
  const fail = function (message) { throw new Error('Сводная руководителя: ' + message); };
  const statusTotal = model.statusCounts.reduce(function (a, b) { return a + b; }, 0);
  if (statusTotal !== model.global.total) fail('нарушен инвариант статусов.');
  if (model.cardGlobalTotal !== model.global.total || model.cardAssignedTotal !== model.global.total) {
    fail('расчёты не совпадают с моделью Карточки объектов.');
  }
  if (model.global.signed + model.global.unsigned !== model.global.total) fail('нарушен signed/unsigned.');
  if (model.problems.length + model.objects.filter(function (o) { return !o.problem; }).length !== model.objects.length) {
    fail('не все объекты классифицированы.');
  }
  const ids = {};
  model.problems.forEach(function (o) {
    const key = objectCardKey_(o.id); if (ids[key]) fail('объект повторён.'); ids[key] = true;
  });
  [model.global, model.orphan].concat(model.objects.map(function (o) { return o.metrics; }))
    .forEach(function (m) {
      if (m.overdue + m.soon + m.missing !== m.deadlineAssignments) fail('deadline buckets пересекаются.');
    });
}

function managerSummaryMetricRow_(metrics) {
  return [metrics.total, metrics.signed, metrics.unsigned, metrics.fix,
    metrics.overdue, metrics.soon, metrics.missing, metrics.nearest || '',
    metrics.readiness, metrics.action];
}

function managerSummaryBuildOutput_(model) {
  const width = Math.max(13, model.statuses.length + 1);
  const blank = function () { return Array(width).fill(''); };
  const rows = [], sectionRows = [], headerRows = [], metricValueRows = [];
  function add(values) { rows.push(values.concat(Array(width - values.length).fill(''))); }
  function addTracked(values, target) { add(values); target.push(rows.length); }
  add([MANAGER_SUMMARY_.TITLE]);
  add(['Состояние на: ' + Utilities.formatDate(model.snapshotAt, model.timezone, 'dd.MM.yyyy HH:mm')]);
  add([]); addTracked(['Общее состояние'], sectionRows);
  addTracked(['Всего объектов', 'Проблемных объектов', 'Объектов без документов',
    'Всего активных документов', 'Подписано с обеих сторон', 'Не подписано',
    'Требует исправления', 'Готовность документов, %', 'Документов без найденного объекта'], headerRows);
  addTracked([model.objects.length, model.problems.length, model.objectsWithoutDocuments,
    model.global.total, model.global.signed, model.global.unsigned, model.global.fix,
    model.global.readiness, model.orphan.total], metricValueRows);
  add([]); addTracked(['Контроль сроков'], sectionRows);
  addTracked(['Просрочено документов', 'Срок в ближайшие 14 дней', 'Без плановой даты'], headerRows);
  addTracked([model.global.overdue, model.global.soon, model.global.missing], metricValueRows);
  add([]); addTracked(['Распределение документов по статусам'], sectionRows);
  addTracked(['Всего документов'].concat(model.statuses), headerRows);
  addTracked([model.global.total].concat(model.statusCounts), metricValueRows);
  add([]); addTracked(['Проблемные объекты'], sectionRows);
  addTracked(['Название объекта', 'Ответственный прораб', 'Статус объекта', 'Всего документов',
    'Подписано с обеих сторон', 'Не подписано', 'Требует исправления', 'Просрочено',
    'До 14 дней', 'Без плановой даты', 'Ближайший срок', 'Готовность, %', 'Контроль'], headerRows);
  const problemStart = rows.length + 1;
  model.problems.forEach(function (o) {
    const m = o.metrics; add([o.name, o.foreman, o.status, m.total, m.signed,
      m.unsigned, m.fix, m.overdue, m.soon, m.missing, m.nearest || '', m.readiness, m.action]);
  });
  if (model.hasOrphans) {
    const m = model.orphan; add([MANAGER_SUMMARY_.ORPHANS, '', '', m.total, m.signed,
      m.unsigned, m.fix, m.overdue, m.soon, m.missing, m.nearest || '', m.readiness, m.action]);
  }
  return { rows: rows, width: width, sectionRows: sectionRows,
    headerRows: headerRows, metricValueRows: metricValueRows,
    problemStart: problemStart, problemCount: model.problems.length + (model.hasOrphans ? 1 : 0) };
}

/** The only destructive phase; source sheets are never written. */
function managerSummaryPublish_(spreadsheet, model) {
  const output = managerSummaryBuildOutput_(model);
  let sheet = spreadsheet.getSheetByName(MANAGER_SUMMARY_.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(MANAGER_SUMMARY_.SHEET_NAME);
  sheet.clear();
  if (sheet.getMaxRows() < output.rows.length) sheet.insertRowsAfter(sheet.getMaxRows(), output.rows.length - sheet.getMaxRows());
  if (sheet.getMaxColumns() < output.width) sheet.insertColumnsAfter(sheet.getMaxColumns(), output.width - sheet.getMaxColumns());
  sheet.getRange(1, 1, output.rows.length, output.width).setValues(output.rows);
  sheet.setFrozenRows(2);
  sheet.getRange(1, 1, output.rows.length, output.width).setVerticalAlignment('middle').setWrap(true);
  sheet.getRange(1, 1, 1, output.width).setBackground('#1f4e78').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14);
  output.sectionRows.forEach(function (row) { sheet.getRange(row, 1, 1, output.width).setBackground('#eeeeee').setFontWeight('bold'); });
  output.headerRows.forEach(function (row) { sheet.getRange(row, 1, 1, output.width).setBackground('#5b9bd5').setFontColor('#ffffff').setFontWeight('bold'); });
  output.metricValueRows.forEach(function (row) { sheet.getRange(row, 1, 1, output.width).setBackground('#d9eaf7'); });
  if (output.problemCount) sheet.getRange(output.problemStart, 1, output.problemCount, 13).setBackground('#eaf2f8');
  sheet.getRange(output.metricValueRows[0], 8).setNumberFormat('0.00%');
  if (output.problemCount) {
    sheet.getRange(output.problemStart, 11, output.problemCount, 1).setNumberFormat('dd.MM.yyyy');
    sheet.getRange(output.problemStart, 12, output.problemCount, 1).setNumberFormat('0.00%');
  }
  for (let column = 1; column <= output.width; column++) sheet.setColumnWidth(column, column <= 3 ? 190 : 125);
  return { objectsRead: model.objects.length, activeDocumentsRead: model.global.total,
    problemObjects: model.problems.length, orphanDocuments: model.orphan.total,
    outputRows: output.rows.length };
}
