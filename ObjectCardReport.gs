/** Generated, read-only management report «Карточка объектов». */
const OBJECT_CARD_REPORT_ = Object.freeze({
  SHEET_NAME: 'Карточка объектов',
  TITLE: 'Карточка объектов',
  ALL_OBJECTS: 'Все объекты',
  ORPHANS: '⚠ Документы без найденного объекта',
  NO_STATUS: 'Без статуса',
  SUMMARY_HEADERS: Object.freeze(['Название объекта', 'Всего документов']),
  DETAIL_HEADERS: Object.freeze([
    H.DOCUMENT_TYPE,
    H.CONTRACT_NUMBER,
    H.DOCUMENT_NUMBER,
    H.DOCUMENT_DATE,
    H.DOCUMENT_STATUS,
    H.CUSTOMER_SIGNING_RESPONSIBLE
  ])
});

/**
 * Rebuilds the complete report from a validated in-memory snapshot.
 * Source sheets are only read. The existing report is touched after every
 * source contract and accounting invariant has passed.
 */
function rebuildObjectCard() {
  return withDocumentLock_(function () {
    assertSystemSheetsStructure_(['OBJECTS', 'DOCUMENTS', 'CARD_DICTIONARY']);
    const spreadsheet = getSystemSpreadsheet_();
    const timezone = spreadsheet.getSpreadsheetTimeZone();
    const snapshotAt = new Date();
    const model = objectCardBuildModel_(
      objectCardReadRows_('OBJECTS', [H.OBJECT_ID, H.OBJECT_NAME]),
      objectCardReadRows_('DOCUMENTS', [
        H.DOCUMENT_ID, H.OBJECT_ID, H.DOCUMENT_TYPE, H.CONTRACT_NUMBER,
        H.DOCUMENT_NUMBER, H.DOCUMENT_DATE, H.DOCUMENT_STATUS,
        H.CUSTOMER_SIGNING_RESPONSIBLE, H.RECORD_STATUS
      ]),
      objectCardReadRows_('CARD_DICTIONARY', [H.DOCUMENT_STATUS]),
      snapshotAt,
      timezone
    );
    return objectCardPublish_(spreadsheet, model);
  });
}

/** Reads one source in a single batch and maps values by exact headers. */
function objectCardReadRows_(sheetKey, requiredHeaders) {
  const context = getSystemSheetContext_(sheetKey);
  const indexes = {};
  requiredHeaders.forEach(function (header) {
    const column = context.headerMap[sysNormalizeHeader_(header)];
    if (!column) {
      throw new Error('На листе «' + context.config.name +
        '» отсутствует обязательный заголовок «' + header + '».');
    }
    indexes[header] = column - 1;
  });
  const count = Math.max(
    context.sheet.getLastRow() - context.config.dataStartRow + 1, 0
  );
  if (!count) return [];
  const rows = context.sheet.getRange(
    context.config.dataStartRow, 1, count, context.headers.length
  ).getValues();
  return rows.map(function (row, offset) {
    const values = {};
    requiredHeaders.forEach(function (header) {
      values[header] = row[indexes[header]];
    });
    return {
      sourceRow: context.config.dataStartRow + offset,
      values: values
    };
  });
}

function objectCardText_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00a0/g, ' ').trim().replace(/\s+/g, ' ');
}

function objectCardKey_(value) {
  return objectCardText_(value).toLocaleLowerCase('ru');
}

function objectCardNaturalCompare_(left, right) {
  return objectCardText_(left).localeCompare(
    objectCardText_(right), 'ru', { numeric: true, sensitivity: 'base' }
  );
}

function objectCardStableSort_(items, comparator) {
  return items.map(function (item, index) {
    return { item: item, index: index };
  }).sort(function (left, right) {
    return comparator(left.item, right.item) || left.index - right.index;
  }).map(function (entry) { return entry.item; });
}

function objectCardBuildStatuses_(dictionaryRows, activeDocuments) {
  const dictionary = [];
  const known = {};
  const noStatusKey = objectCardKey_(OBJECT_CARD_REPORT_.NO_STATUS);
  let needsNoStatus = false;
  dictionaryRows.forEach(function (row) {
    const value = objectCardText_(row.values[H.DOCUMENT_STATUS]);
    const key = objectCardKey_(value);
    if (!value || known[key]) return;
    known[key] = value;
    if (key === noStatusKey) {
      needsNoStatus = true;
      return;
    }
    dictionary.push(value);
  });
  const unknownByKey = {};
  activeDocuments.forEach(function (document) {
    const value = objectCardText_(document.values[H.DOCUMENT_STATUS]);
    if (!value) {
      needsNoStatus = true;
      return;
    }
    const key = objectCardKey_(value);
    if (key === noStatusKey) {
      needsNoStatus = true;
      return;
    }
    if (!known[key] && !unknownByKey[key]) unknownByKey[key] = value;
  });
  const unknown = objectCardStableSort_(
    Object.keys(unknownByKey).map(function (key) { return unknownByKey[key]; }),
    objectCardNaturalCompare_
  );
  return dictionary.concat(unknown).concat(
    needsNoStatus ? [OBJECT_CARD_REPORT_.NO_STATUS] : []
  );
}

function objectCardStats_(documents, statuses) {
  const counts = statuses.map(function () { return 0; });
  const statusIndex = {};
  statuses.forEach(function (status, index) {
    statusIndex[objectCardKey_(status)] = index;
  });
  documents.forEach(function (document) {
    const actual = objectCardText_(document.values[H.DOCUMENT_STATUS]);
    const key = objectCardKey_(actual || OBJECT_CARD_REPORT_.NO_STATUS);
    if (!Object.prototype.hasOwnProperty.call(statusIndex, key)) {
      throw new Error('Активный документ имеет неучтённый статус «' + actual + '».');
    }
    counts[statusIndex[key]]++;
  });
  const total = documents.length;
  if (counts.reduce(function (sum, count) { return sum + count; }, 0) !== total) {
    throw new Error('Нарушен инвариант учёта статусов документов.');
  }
  return { total: total, counts: counts };
}

function objectCardDocumentCompare_(left, right) {
  return objectCardNaturalCompare_(
    left.values[H.DOCUMENT_TYPE], right.values[H.DOCUMENT_TYPE]
  ) || objectCardNaturalCompare_(
    left.values[H.DOCUMENT_NUMBER], right.values[H.DOCUMENT_NUMBER]
  );
}

/** Pure snapshot planner, deliberately separated from destructive output. */
function objectCardBuildModel_(objectRows, documentRows, dictionaryRows,
  snapshotAt, timezone) {
  const objectsById = {};
  const objects = [];
  objectRows.forEach(function (row) {
    const id = objectCardText_(row.values[H.OBJECT_ID]);
    const key = objectCardKey_(id);
    if (!id) throw new Error('Лист «Объекты», строка ' + row.sourceRow +
      ': не указан «' + H.OBJECT_ID + '».');
    if (objectsById[key]) throw new Error('Лист «Объекты»: повторяется «' +
      H.OBJECT_ID + '» ' + id + '.');
    const object = {
      id: id,
      name: objectCardText_(row.values[H.OBJECT_NAME]),
      documents: []
    };
    objectsById[key] = object;
    objects.push(object);
  });
  const activeDocuments = documentRows.filter(function (row) {
    return objectCardKey_(row.values[H.RECORD_STATUS]) === objectCardKey_(
      SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS
    );
  });
  const orphans = [];
  activeDocuments.forEach(function (document) {
    const object = objectsById[objectCardKey_(document.values[H.OBJECT_ID])];
    (object ? object.documents : orphans).push(document);
  });
  const statuses = objectCardBuildStatuses_(dictionaryRows, activeDocuments);
  const sortedObjects = objectCardStableSort_(objects, function (left, right) {
    return objectCardNaturalCompare_(left.id, right.id);
  });
  sortedObjects.forEach(function (object) {
    object.documents = objectCardStableSort_(
      object.documents, objectCardDocumentCompare_
    );
    object.stats = objectCardStats_(object.documents, statuses);
  });
  const sortedOrphans = objectCardStableSort_(orphans, objectCardDocumentCompare_);
  const globalStats = objectCardStats_(activeDocuments, statuses);
  const orphanStats = objectCardStats_(sortedOrphans, statuses);
  const assigned = sortedObjects.reduce(function (sum, object) {
    return sum + object.stats.total;
  }, 0) + orphanStats.total;
  if (assigned !== globalStats.total) {
    throw new Error('Нарушен инвариант распределения активных документов.');
  }
  return {
    snapshotAt: snapshotAt,
    timezone: timezone,
    statuses: statuses,
    objects: sortedObjects,
    orphans: sortedOrphans,
    orphanStats: orphanStats,
    globalStats: globalStats
  };
}

function objectCardSummaryRow_(label, stats, width) {
  return [label, stats.total].concat(stats.counts).concat(
    Array(Math.max(width - stats.counts.length - 2, 0)).fill('')
  );
}

function objectCardDetailRow_(document, width) {
  const values = OBJECT_CARD_REPORT_.DETAIL_HEADERS.map(function (header) {
    return document.values[header];
  });
  return values.concat(Array(Math.max(width - values.length, 0)).fill(''));
}

function objectCardBuildOutput_(model) {
  const width = Math.max(6, 2 + model.statuses.length);
  const blank = Array(width).fill('');
  const rows = [];
  const groups = [];
  const summaryRows = [];
  const detailHeaderRows = [];
  rows.push([OBJECT_CARD_REPORT_.TITLE].concat(blank.slice(1)));
  rows.push(['Состояние на: ' + Utilities.formatDate(
    model.snapshotAt, model.timezone, 'dd.MM.yyyy HH:mm'
  )].concat(blank.slice(1)));
  rows.push(blank.slice());
  rows.push(OBJECT_CARD_REPORT_.SUMMARY_HEADERS.concat(model.statuses)
    .concat(blank.slice(2 + model.statuses.length)));
  rows.push(objectCardSummaryRow_(
    OBJECT_CARD_REPORT_.ALL_OBJECTS, model.globalStats, width
  ));
  summaryRows.push(5);
  rows.push(blank.slice());

  function appendSection(label, documents, stats) {
    rows.push(objectCardSummaryRow_(label, stats, width));
    const summaryRow = rows.length;
    summaryRows.push(summaryRow);
    if (!documents.length) return;
    rows.push(OBJECT_CARD_REPORT_.DETAIL_HEADERS.concat(blank.slice(6)));
    const firstDetailRow = rows.length;
    detailHeaderRows.push(firstDetailRow);
    documents.forEach(function (document) {
      rows.push(objectCardDetailRow_(document, width));
    });
    groups.push({ startRow: firstDetailRow, rowCount: documents.length + 1,
      summaryRow: summaryRow });
  }
  model.objects.forEach(function (object) {
    appendSection(object.name, object.documents, object.stats);
  });
  if (model.orphans.length) {
    appendSection(OBJECT_CARD_REPORT_.ORPHANS, model.orphans, model.orphanStats);
  }
  return { rows: rows, width: width, groups: groups,
    summaryRows: summaryRows, detailHeaderRows: detailHeaderRows };
}

function objectCardRemoveGroups_(sheet) {
  // Google Sheets supports at most eight outline levels. Applying the maximum
  // negative shift below the first row removes every old level in one API
  // operation and remains safe when row-group controls are positioned BEFORE.
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 1) return;
  sheet.getRange(2, 1, maxRows - 1, 1).shiftRowGroupDepth(-8);
}

/** The only destructive phase. It owns only the generated output sheet. */
function objectCardPublish_(spreadsheet, model) {
  const output = objectCardBuildOutput_(model);
  let sheet = spreadsheet.getSheetByName(OBJECT_CARD_REPORT_.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(OBJECT_CARD_REPORT_.SHEET_NAME);
  objectCardRemoveGroups_(sheet);
  sheet.clear();
  if (sheet.getMaxRows() < output.rows.length) {
    sheet.insertRowsAfter(sheet.getMaxRows(), output.rows.length - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < output.width) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(), output.width - sheet.getMaxColumns()
    );
  }
  sheet.getRange(1, 1, output.rows.length, output.width).setValues(output.rows);
  sheet.setFrozenRows(4);
  sheet.getRange(1, 1, output.rows.length, output.width)
    .setVerticalAlignment('middle').setWrap(true);
  sheet.getRange(1, 1, 1, output.width)
    .setBackground('#1f4e78').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(14);
  sheet.getRange(4, 1, 1, output.width)
    .setBackground('#5b9bd5').setFontColor('#ffffff').setFontWeight('bold');
  output.summaryRows.forEach(function (row) {
    sheet.getRange(row, 1, 1, output.width)
      .setBackground(row === 5 ? '#d9eaf7' : '#eaf2f8').setFontWeight('bold');
  });
  output.detailHeaderRows.forEach(function (row) {
    sheet.getRange(row, 1, 1, output.width)
      .setBackground('#eeeeee').setFontWeight('bold');
  });
  output.groups.forEach(function (group) {
    sheet.getRange(group.startRow + 1, 4, group.rowCount - 1, 1)
      .setNumberFormat('dd.MM.yyyy');
  });
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 130);
  for (let column = 3; column <= output.width; column++) {
    sheet.setColumnWidth(column, column === 4 ? 110 : 145);
  }
  output.groups.forEach(function (group) {
    sheet.getRange(group.startRow, 1, group.rowCount, 1).shiftRowGroupDepth(1);
    const rowGroup = sheet.getRowGroup(group.startRow, 1);
    if (!rowGroup) throw new Error('Не удалось создать группу строк отчёта.');
    rowGroup.collapse();
  });
  return {
    timestamp: Utilities.formatDate(
      model.snapshotAt, model.timezone, 'dd.MM.yyyy HH:mm'
    ),
    objectsRead: model.objects.length,
    activeDocumentsRead: model.globalStats.total,
    objectsRendered: model.objects.length,
    objectsWithoutDocuments: model.objects.filter(function (object) {
      return object.stats.total === 0;
    }).length,
    orphanDocuments: model.orphanStats.total,
    statusColumns: model.statuses.slice(),
    outputRows: output.rows.length,
    groupsCreated: output.groups.length
  };
}
