'use strict';
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');
const cp = require('child_process');
const files = ['SystemCore.gs', 'DocumentArchitectureCore.gs', 'CreateObjectDocuments.gs', 'SyncObjectData.gs', 'ArchiveChangeHistory.gs', 'OperatorCard.gs', 'OperatorCardSave.gs', 'ObjectSheetControls.gs', 'Code.gs', 'WebAppAuth.gs', 'DocumentWebAppServer.gs', 'DocumentArchitectureV2Migration.gs'];
const source = files.map(f => fs.readFileSync(f, 'utf8')).join('\n');
let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('✓ ' + name); }
  catch (error) { console.error('✗ ' + name + '\n  ' + error.stack); process.exitCode = 1; }
}
function baseContext(extra = {}) {
  const sandbox = {
    console,
    Utilities: {
      formatDate(date, zone, pattern) {
        if (pattern === 'yyyyMMdd') return '20260804';
        if (pattern === 'yyyy-MM') return date.toISOString().slice(0, 7);
        if (pattern === 'yyyy-MM-dd') {
          const parts = new Intl.DateTimeFormat('en-CA', { timeZone: zone, day: '2-digit', month: '2-digit', year: 'numeric' }).formatToParts(date);
          const map = Object.fromEntries(parts.map(part => [part.type, part.value]));
          return `${map.year}-${map.month}-${map.day}`;
        }
        if (pattern === 'dd.MM.yyyy') {
          const parts = new Intl.DateTimeFormat('en-GB', { timeZone: zone, day: '2-digit', month: '2-digit', year: 'numeric' }).formatToParts(date);
          const map = Object.fromEntries(parts.map(part => [part.type, part.value]));
          return `${map.day}.${map.month}.${map.year}`;
        }
        return date.toISOString();
      },
      parseDate(text, zone, pattern) {
        if (pattern !== 'dd.MM.yyyy') throw new Error('unexpected pattern');
        const [day, month, year] = text.split('.').map(Number);
        const utc = Date.UTC(year, month - 1, day, 0, 0, 0);
        const probe = new Date(utc);
        const local = new Intl.DateTimeFormat('en-US', {
          timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit',
          hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23'
        }).formatToParts(probe);
        const map = Object.fromEntries(local.map(part => [part.type, part.value]));
        const asUtc = Date.UTC(Number(map.year), Number(map.month) - 1, Number(map.day), Number(map.hour), Number(map.minute), Number(map.second));
        return new Date(utc - (asUtc - utc));
      }
    },
    PropertiesService: {
      getDocumentProperties() {
        return extra.properties || { getProperty() { return null; }, setProperty() {} };
      },
      getScriptProperties() {
        return extra.scriptProperties || { getProperty() { return null; }, setProperty() {}, deleteProperty() {} };
      }
    },
    SpreadsheetApp: extra.SpreadsheetApp || { flush() {} },
    LockService: extra.LockService || {},
    Session: { getActiveUser() { return { getEmail() { return 'tester@example.com'; } }; } },
    HtmlService: extra.HtmlService || {},
    CacheService: extra.CacheService || { getScriptCache() { return { get() { return null; }, put() {}, remove() {} }; } }
  };
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox);
  return sandbox;
}
function idContext(operationIds, changeOperationIds, propertyStart) {
  let property = propertyStart == null ? null : String(propertyStart);
  const ctx = baseContext({
    properties: {
      getProperty() { return property; },
      setProperty(key, value) { property = value; }
    }
  });
  ctx.getSystemSpreadsheet_ = () => ({ getSpreadsheetTimeZone: () => 'Europe/Moscow' });
  const sets = { OPERATION_HISTORY: operationIds, CHANGE_HISTORY: changeOperationIds };
  ctx.getSystemSheetContext_ = key => ({
    config: { dataStartRow: 3 },
    sheet: {
      getLastRow: () => sets[key].length + 2,
      getRange: () => ({ getValues: () => sets[key].map(value => [value]) })
    }
  });
  ctx.getSystemColumn_ = () => 1;
  return { ctx, property: () => property };
}

test('1. operation ID has OP-YYYYMMDD-NNNN format', () => {
  const { ctx } = idContext([], [], null);
  assert.match(ctx.generateOperationId_(new Date()), /^OP-20260804-0001$/);
});
test('2. two same-day operations are sequential', () => {
  const state = idContext([], [], null);
  assert.equal(state.ctx.generateOperationId_(new Date()), 'OP-20260804-0001');
  assert.equal(state.ctx.generateOperationId_(new Date()), 'OP-20260804-0002');
});
test('3. legacy UUIDs do not affect the next sequence', () => {
  const { ctx } = idContext(['OP-old-uuid'], ['CHG-old-uuid'], null);
  assert.equal(ctx.generateOperationId_(new Date()), 'OP-20260804-0001');
});
test('4. maximum operation sequence is read from both histories', () => {
  const { ctx } = idContext(['OP-20260804-0003'], ['OP-20260804-0007'], null);
  assert.equal(ctx.generateOperationId_(new Date()), 'OP-20260804-0008');
});
test('5. change IDs are linked to operation ID', () => {
  const { ctx } = idContext([], [], null);
  assert.equal(ctx.generateChangeId_('OP-20260804-0012', 3), 'CHG-20260804-0012-0003');
});
test('6. serialized concurrent reservations cannot duplicate', () => {
  const state = idContext([], [], null);
  const ids = [state.ctx.generateOperationId_(new Date()), state.ctx.generateOperationId_(new Date())];
  assert.equal(new Set(ids).size, 2);
  assert.equal(state.property(), '2');
});

function syncFixture(sourceValues, docs) {
  const ctx = baseContext();
  ctx.getSystemSpreadsheet_ = () => ({ getSpreadsheetTimeZone: () => 'UTC' });
  const headers = [
    'ID документа', 'ID объекта', 'Название объекта', 'Статус записи', 'Номер договора', 'Статус объекта',
    'Дата начала работ', 'Дата окончания (по плану)', 'Дата окончания (по факту)',
    'Ответственный прораб', 'ID ответственного прораба', 'Дата обновления', 'Кто обновил (email)',
    'Дата изменения статуса документа'
  ];
  const indexes = Object.fromEntries(headers.map((h, i) => [h, i]));
  const activeRows = docs.map((values, i) => ({
    row: headers.map(h => values[h] === undefined ? '' : values[h]),
    sheetRow: i + 4,
    documentId: values['ID документа'],
    objectId: values['ID объекта']
  }));
  const result = ctx.buildObjectSyncPlan_(
    { OBJ1: sourceValues },
    { indexes, activeRows },
    new Date('2026-08-04T10:00:00Z'),
    'tester@example.com',
    'OP-20260804-0001'
  );
  return { ctx, result, headers };
}
const baseline = {
  'Название объекта': 'Объект Альфа', 'Номер договора': 'A-1', 'Статус объекта': 'В работе',
  'Дата начала работ': '',
  'Дата окончания (по плану)': '',
  'Дата окончания (по факту)': '', 'Ответственный прораб': 'Иванов',
  'ID ответственного прораба': 'ST-1'
};
function document(extra = {}) {
  return Object.assign({
    'ID документа': 'DOC-OBJ1-0001', 'ID объекта': 'OBJ1', 'Статус записи': 'Активная',
    'Дата обновления': new Date('2026-08-03T10:00:00Z'), 'Кто обновил (email)': 'old@example.com',
    'Дата изменения статуса документа': new Date('2026-07-01T00:00:00Z')
  }, baseline, extra);
}
test('7. synchronization plan without differences has no changes', () => {
  const fixture = syncFixture(baseline, [document()]);
  assert.equal(fixture.result.changedRows.length, 0);
});
test('8. one changed object field updates every active document', () => {
  const source = Object.assign({}, baseline, { 'Статус объекта': 'Завершён' });
  const fixture = syncFixture(source, [document(), document({ 'ID документа': 'DOC-OBJ1-0002' })]);
  assert.equal(fixture.result.changedRows.length, 2);
});
test('9. empty source field clears document value', () => {
  const source = Object.assign({}, baseline, { 'Ответственный прораб': '' });
  const fixture = syncFixture(source, [document()]);
  assert.equal(fixture.result.changedRows[0].row[8], '');
});
test('10. non-active rows are excluded before planning', () => {
  const text = fs.readFileSync('SyncObjectData.gs', 'utf8');
  assert.match(text, /ACTIVE_RECORD_STATUS[\s\S]*return;/);
  assert.doesNotMatch(text, /ARCHIVED_RECORD_STATUS[\s\S]*setValues/);
});
test('11. document status change timestamp is never part of sync fields', () => {
  const source = Object.assign({}, baseline, { 'Статус объекта': 'Завершён' });
  const fixture = syncFixture(source, [document()]);
  assert.ok(!fixture.result.changes.some(c => c.fieldName === 'Дата изменения статуса документа'));
});
test('12. invalid nonempty date skips source object', () => {
  const ctx = baseContext();
  const text = fs.readFileSync('SyncObjectData.gs', 'utf8');
  assert.match(text, /!objectSyncEmpty_\(value\) && !objectSyncValidDate_\(value\)/);
  assert.equal(ctx.objectSyncValidDate_('04.08.2026'), false);
});
test('13. duplicate object IDs are rejected as a group', () => {
  const text = fs.readFileSync('SyncObjectData.gs', 'utf8');
  assert.match(text, /idRows\[candidate\.id\]\.length > 1/);
  assert.match(text, /ID объекта повторяется в строках/);
});
test('14. history records old and new values per changed field', () => {
  const source = Object.assign({}, baseline, { 'Номер договора': 'A-2' });
  const fixture = syncFixture(source, [document()]);
  const change = fixture.result.changes.find(c => c.fieldName === 'Номер договора');
  assert.equal(change.oldValue, 'A-1'); assert.equal(change.newValue, 'A-2');
});
test('15. critical errors use escaped HtmlService modal dialog', () => {
  let shown = false; let htmlText = '';
  const output = { setWidth() { return this; }, setHeight() { return this; } };
  const ctx = baseContext({
    HtmlService: { createHtmlOutput(html) { htmlText = html; return output; } },
    SpreadsheetApp: { getUi() { return { showModalDialog() { shown = true; } }; } }
  });
  ctx.showCriticalOperationError_('Ошибка', '<bad>', '&reason');
  assert.ok(shown); assert.ok(htmlText.includes('&lt;bad&gt;')); assert.ok(htmlText.includes('&amp;reason'));
  assert.ok(htmlText.includes('КРИТИЧЕСКАЯ ОШИБКА'));
});
test('16. ordinary dictionary values are not validated by creation', () => {
  const text = fs.readFileSync('CreateObjectDocuments.gs', 'utf8');
  assert.ok(!text.includes('DOCUMENT_LOCATION') || !/assertCreationDictionaryValues_[\s\S]{0,1500}DOCUMENT_LOCATION/.test(text));
});
test('17. missing initial document status stops creation', () => {
  const text = fs.readFileSync('CreateObjectDocuments.gs', 'utf8');
  assert.match(text, /!hasInitialStatus/); assert.match(text, /INITIAL_DOCUMENT_STATUS/);
});
test('18. missing active status stops creation and synchronization', () => {
  assert.match(fs.readFileSync('CreateObjectDocuments.gs', 'utf8'), /!hasActiveStatus/);
  assert.match(fs.readFileSync('SyncObjectData.gs', 'utf8'), /if \(!found\)/);
});
test('19. fact writes precede history and history failure states facts remain', () => {
  for (const f of ['CreateObjectDocuments.gs', 'SyncObjectData.gs']) {
    const text = fs.readFileSync(f, 'utf8');
    const fact = text.indexOf(f === 'CreateObjectDocuments.gs' ? '.setValues(prepared.documentRows)' : 'writeObjectSyncFacts_');
    const history = text.indexOf(f === 'CreateObjectDocuments.gs' ? 'writeCreationChangeHistory_(prepared.changeRows)' : 'writeObjectSyncChangeHistory_(plan.changes)');
    assert.ok(fact >= 0 && history > fact);
  }
});
test('20. operator endpoints delegate to the dedicated module', () => {
  const text = fs.readFileSync('Code.gs', 'utf8');
  assert.match(text, /return operatorCardGetFilterData_\(\)/);
  assert.match(text, /return operatorCardApply_\(filters\)/);
  assert.doesNotMatch(text, /OPERATOR_CONFIG|loadOperatorCard_/);
});


test('21. fact writer preserves unrelated values and formulas', () => {
  const fixture = syncFixture(
    Object.assign({}, baseline, { 'Номер договора': 'A-2' }),
    [document()]
  );
  const headers = fixture.headers.concat(['Комментарий', 'Сумма документа']);
  const commentIndex = headers.indexOf('Комментарий');
  const formulaIndex = headers.indexOf('Сумма документа');
  const contractIndex = headers.indexOf('Номер договора');
  fixture.result.changedRows[0].row.push('Не менять', '=SUM(1;2)');
  const storedRow = document();
  const stored = headers.map(header => storedRow[header] === undefined ? '' : storedRow[header]);
  stored[commentIndex] = 'Не менять';
  stored[formulaIndex] = '=SUM(1;2)';
  const headerMap = Object.fromEntries(headers.map((header, index) => [header, index + 1]));
  const context = {
    headers,
    headerMap,
    sheet: {
      getRange(startRow, startColumn, rowCount, columnCount) {
        assert.equal(columnCount, 1, 'writer must only request one permitted column');
        return {
          setValues(values) {
            values.forEach((value, offset) => { stored[startColumn - 1] = value[0]; });
          }
        };
      }
    }
  };
  fixture.ctx.writeObjectSyncFacts_(context, fixture.result.changedRows);
  assert.equal(stored[contractIndex], 'A-2');
  assert.equal(stored[commentIndex], 'Не менять');
  assert.equal(stored[formulaIndex], '=SUM(1;2)');
});

test('22. dictionary guard never weakens foreign protection and stays idempotent', () => {
  const ctx = baseContext({ SpreadsheetApp: { ProtectionType: { RANGE: 'RANGE' } } });
  const description = 'Системное значение документооборота';
  let created = 0;
  function protection(desc, row, column) {
    return {
      warningCalls: 0,
      getDescription: () => desc,
      getRange: () => ({
        getRow: () => row,
        getColumn: () => column,
        getNumRows: () => 1,
        getNumColumns: () => 1
      }),
      setDescription(value) { desc = value; return this; },
      setWarningOnly() { this.warningCalls++; return this; }
    };
  }
  const foreign = protection('Чужая строгая защита', 3, 1);
  const own = protection(description, 3, 1);
  const protectionsByCell = { '3:1': [foreign, own], '3:2': [protection('Чужая', 3, 2)] };
  function cell(row, column) {
    return {
      getRow: () => row,
      getColumn: () => column,
      setNote() { return this; },
      setBackground() { return this; },
      setFontColor() { return this; },
      setFontWeight() { return this; },
      getProtections: () => protectionsByCell[`${row}:${column}`],
      protect() {
        created++;
        const result = protection('', row, column);
        protectionsByCell[`${row}:${column}`].push(result);
        return result;
      }
    };
  }
  const sheet = {
    getLastRow: () => 3,
    getRange(row, column, rowCount, columnCount) {
      if (rowCount === undefined) return cell(row, column);
      const value = column === 1 ? 'Ожидает заполнения' : 'Активная';
      return { getValues: () => [[value]] };
    }
  };
  ctx.assertSystemSheetsStructure_ = () => {};
  ctx.getSystemSheetContext_ = () => ({
    sheet,
    config: { dataStartRow: 3 },
    headers: ['Статус документа', 'Статус записи'],
    headerMap: { 'Статус документа': 1, 'Статус записи': 2 }
  });
  ctx.getSystemColumn_ = (key, header) => header === 'Статус документа' ? 1 : 2;
  ctx.setupSystemDictionaryGuards();
  ctx.setupSystemDictionaryGuards();
  assert.equal(foreign.warningCalls, 0, 'foreign strict protection was weakened');
  assert.equal(own.warningCalls, 2, 'own protection should be reused');
  assert.equal(created, 1, 'second run must reuse the function-owned protection');
});


function creationPrepareFixture(count) {
  const ctx = baseContext();
  ctx.getSystemSpreadsheet_ = () => ({ getSpreadsheetTimeZone: () => 'UTC' });
  const headers = [
    'ID документа','ID объекта','Название объекта','Тип документа','ID типа документа','Номер договора','Номер документа','Статус документа',
    'Статус объекта','Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)',
    'Дата создания','Дата обновления','Кто обновил (email)','Ответственный прораб','ID ответственного прораба','Кто ответственный за подписание (заказчик)',
    'Дата изменения статуса документа','Источник создания','Статус записи'
  ];
  const context = { headers, headerMap: Object.fromEntries(headers.map((h, i) => [h, i + 1])) };
  const missing = Array.from({ length: count }, (_, i) => ({
    object: {
      id: 'OBJ' + String(i + 1).padStart(2, '0'), name: 'Объект ' + (i + 1), contractNumber: 'DOG-' + (i + 1),
      objectStatus: 'Действующий', workStartDate: new Date('2026-08-01T00:00:00Z'),
      workEndPlan: new Date('2026-08-31T00:00:00Z'), workEndFact: i === 0 ? new Date('2026-08-20T00:00:00Z') : '',
      responsibleForeman: 'Иванов', responsibleForemanId: 'ST-1'
    },
    rule: { id: 'DT-1', name: 'Акт', repeatability: 'Много' }
  }));
  return ctx.prepareDocumentRows_(missing, { documentIds: {} }, context, 4, new Date('2026-08-04T10:00:00Z'), 'tester@example.com', 'OP-20260804-0001');
}

test('23. one created document produces one history row', () => {
  const prepared = creationPrepareFixture(1);
  assert.equal(prepared.documentRows.length, 1);
  assert.equal(prepared.changeRows.length, 1);
  assert.equal(prepared.changeRows[0].fieldName, 'Создание документа');
});
test('24. 18 created documents produce 18 history rows, not field-per-value rows', () => {
  const prepared = creationPrepareFixture(18);
  assert.equal(prepared.documentRows.length, 18);
  assert.equal(prepared.changeRows.length, 18);
  assert.notEqual(prepared.changeRows.length, 324);
});
test('25. initial creation snapshot contains main nonempty fields', () => {
  const snapshot = creationPrepareFixture(1).changeRows[0].newValue;
  ['ID документа: DOC-OBJ01-0001','ID объекта: OBJ01','Название объекта: Объект 1','Тип документа: Акт','ID типа документа: DT-1',
   'Номер договора: DOG-1','Номер документа: 1','Статус документа: Ожидает заполнения','Статус объекта: Действующий',
   'Дата начала работ:','Дата окончания (по плану):','Дата окончания (по факту):',
   'Ответственный прораб: Иванов','ID ответственного прораба: ST-1','Источник создания: Создание документов по объекту',
   'Статус записи: Активная'].forEach(part => assert.ok(snapshot.includes(part), part));
});
test('26. sync does not create separate history row for update timestamp', () => {
  const fixture = syncFixture(Object.assign({}, baseline, { 'Номер договора': 'A-2' }), [document()]);
  assert.ok(!fixture.result.changes.some(c => c.fieldName === 'Дата обновления'));
});
test('27. sync does not create separate history row for update email', () => {
  const fixture = syncFixture(Object.assign({}, baseline, { 'Номер договора': 'A-2' }), [document()]);
  assert.ok(!fixture.result.changes.some(c => c.fieldName === 'Кто обновил (email)'));
});
test('28. sync fields counter counts only changed business fields', () => {
  const fixture = syncFixture(Object.assign({}, baseline, { 'Номер договора': 'A-2' }), [document()]);
  assert.equal(fixture.result.changedRows.length, 1);
  assert.equal(fixture.result.changes.length, 1);
});
test('29. archive cutoff selects only rows older than cutoff', () => {
  const ctx = baseContext();
  ctx.getSystemSpreadsheet_ = () => ({ getSpreadsheetTimeZone: () => 'UTC' });
  const cutoff = ctx.parseArchiveCutoffDate_('01.08.2026');
  assert.ok(new Date(2026, 6, 31).getTime() < cutoff.getTime());
  assert.ok(!(new Date(2026, 7, 1).getTime() < cutoff.getTime()));
});
test('30. archive cutoff uses spreadsheet timezone and keeps cutoff-day rows active', () => {
  const ctx = baseContext();
  ctx.getSystemSpreadsheet_ = () => ({ getSpreadsheetTimeZone: () => 'Asia/Tokyo' });
  const cutoff = ctx.parseArchiveCutoffDate_('01.08.2026');
  assert.equal(ctx.Utilities.formatDate(cutoff, 'Asia/Tokyo', 'dd.MM.yyyy'), '01.08.2026');
  assert.ok(new Date('2026-07-31T14:59:59Z').getTime() < cutoff.getTime());
  assert.ok(!(new Date('2026-07-31T15:00:00Z').getTime() < cutoff.getTime()));
});
test('31. archive rows older than cutoff are grouped by month', () => {
  const ctx = baseContext();
  ctx.getSystemSpreadsheet_ = () => ({ getSpreadsheetTimeZone: () => 'UTC' });
  const context = { headerMap: { 'Дата и время изменения': 2 } };
  const groups = ctx.groupArchiveRowsByMonth_(context, [
    { values: ['CHG-1', new Date('2026-07-31T10:00:00Z')] },
    { values: ['CHG-2', new Date('2026-08-01T10:00:00Z')] }
  ]);
  assert.deepEqual(Object.keys(groups).sort(), ['2026-07', '2026-08']);
});
test('32. archive writer skips duplicate writes for identical existing ID', () => {
  const ctx = baseContext();
  const state = { counters: ctx.emptyArchiveCounters_(), warnings: [] };
  const row = ['CHG-1', new Date('2026-07-31T10:00:00Z'), 'x'];
  const sourceContext = archiveTestContext(ctx, [row]);
  const archiveSheet = archiveMockSheet(sourceContext, [row]);
  ctx.getOrCreateArchiveSpreadsheet_ = () => ({ marker: true });
  ctx.getOrCreateArchiveSheet_ = () => archiveSheet;
  ctx.processArchiveMonth_({}, sourceContext, '2026-07', [{ values: row }], state);
  assert.equal(state.counters.alreadyArchived, 1);
  assert.equal(state.counters.writtenRows, 0);
});
test('33. matching ID with different data is a critical archive error', () => {
  const ctx = baseContext();
  const state = { counters: ctx.emptyArchiveCounters_(), warnings: [] };
  const row = ['CHG-1', new Date('2026-07-31T10:00:00Z'), 'source'];
  const sourceContext = archiveTestContext(ctx, [row]);
  const archiveSheet = archiveMockSheet(sourceContext, [['CHG-1', new Date('2026-07-31T10:00:00Z'), 'archive']]);
  ctx.getOrCreateArchiveSpreadsheet_ = () => ({});
  ctx.getOrCreateArchiveSheet_ = () => archiveSheet;
  assert.throws(() => ctx.processArchiveMonth_({}, sourceContext, '2026-07', [{ values: row }], state), /отличающимися данными/);
});
test('34. archive write error keeps source rows undeleted', () => {
  const ctx = baseContext();
  const text = fs.readFileSync('ArchiveChangeHistory.gs', 'utf8');
  assert.ok(text.indexOf('processArchiveMonth_') < text.indexOf('deleteArchiveSourceRows_'));
});
test('35. source rows are deleted only after archive ID verification', () => {
  const text = fs.readFileSync('ArchiveChangeHistory.gs', 'utf8');
  const callIndex = text.indexOf('processArchiveMonth_(folder, context');
  const deleteIndex = text.indexOf('deleteArchiveSourceRows_(context.sheet');
  const verifyIndex = text.indexOf('const after = readArchiveExistingById_', text.indexOf('function processArchiveMonth_'));
  assert.ok(callIndex >= 0 && deleteIndex > callIndex && verifyIndex >= 0);
});
test('36. empty source change ID stops archiving', () => {
  const ctx = baseContext();
  ctx.getSystemSheetContext_ = () => archiveTestContext(ctx, []);
  const state = { counters: ctx.emptyArchiveCounters_(), warnings: [] };
  assert.throws(() => ctx.assertArchiveSourceIds_(archiveTestContext(ctx, []), [{ sheetRow: 3, values: ['', new Date(), 'x'] }], state), /пустой ID изменения/);
});
test('37. duplicate source change ID stops archiving', () => {
  const ctx = baseContext();
  ctx.getSystemSheetContext_ = () => archiveTestContext(ctx, []);
  const state = { counters: ctx.emptyArchiveCounters_(), warnings: [] };
  assert.throws(() => ctx.assertArchiveSourceIds_(archiveTestContext(ctx, []), [
    { sheetRow: 3, values: ['CHG-1', new Date(), 'x'] },
    { sheetRow: 4, values: ['CHG-1', new Date(), 'y'] }
  ], state), /повторяется/);
});

function archiveHeaderErrorFixture() {
  const ctx = baseContext();
  const state = { counters: Object.assign(ctx.emptyArchiveCounters_(), {
    checkedRows: 48,
    selectedRows: 1,
    archiveFiles: 1
  }), warnings: [] };
  const row = ['CHG-1', new Date('2026-07-31T10:00:00Z'), 'x'];
  const sourceContext = archiveTestContext(ctx, [row]);
  const badSheet = archiveMockSheet(sourceContext, []);
  badSheet.getRange = function (rowNumber, column, rowCount, columnCount) {
    return {
      getValues() {
        if (rowNumber === 2) return [['Повреждённый заголовок', 'Дата и время изменения', 'Новое значение']];
        return [];
      },
      setValues() { throw new Error('Не должна выполняться запись при повреждённом заголовке'); }
    };
  };
  ctx.getOrCreateArchiveSpreadsheet_ = () => ({});
  ctx.getOrCreateArchiveSheet_ = function (spreadsheet, context) {
    ctx.ensureArchiveHeaders_(badSheet, context);
    return badSheet;
  };
  let caught;
  try { ctx.processArchiveMonth_({}, sourceContext, '2026-07', [{ values: row }], state); }
  catch (error) { caught = error; }
  assert.ok(caught, 'expected corrupted header error');
  return { error: caught, state };
}

test('38. corrupted archive header error preserves checkedRows', () => {
  const fixture = archiveHeaderErrorFixture();
  assert.equal(fixture.error.archiveState.counters.checkedRows, 48);
});
test('39. corrupted archive header error preserves selectedRows', () => {
  const fixture = archiveHeaderErrorFixture();
  assert.equal(fixture.error.archiveState.counters.selectedRows, 1);
});
test('40. corrupted archive header error preserves archiveFiles', () => {
  const fixture = archiveHeaderErrorFixture();
  assert.equal(fixture.error.archiveState.counters.archiveFiles, 1);
});
test('41. corrupted archive header leaves writtenRows and deletedRows zero', () => {
  const fixture = archiveHeaderErrorFixture();
  assert.equal(fixture.error.archiveState.counters.writtenRows, 0);
  assert.equal(fixture.error.archiveState.counters.deletedRows, 0);
});
test('42. corrupted archive header keeps original error message', () => {
  const fixture = archiveHeaderErrorFixture();
  assert.match(fixture.error.message, /некорректный заголовок/);
});
test('43. ordinary Drive or Spreadsheet error inside processArchiveMonth receives archiveState', () => {
  const ctx = baseContext();
  const state = { counters: Object.assign(ctx.emptyArchiveCounters_(), { checkedRows: 48, selectedRows: 1, archiveFiles: 1 }), warnings: [] };
  const sourceContext = archiveTestContext(ctx, [['CHG-1', new Date('2026-07-31T10:00:00Z'), 'x']]);
  ctx.getOrCreateArchiveSpreadsheet_ = () => { throw new Error('DriveApp недоступен'); };
  assert.throws(() => ctx.processArchiveMonth_({}, sourceContext, '2026-07', [{ values: sourceContext.sheet._rows[0] }], state), error => {
    assert.equal(error.message, 'DriveApp недоступен');
    assert.equal(error.archiveState, state);
    return true;
  });
});
test('44. runChangeHistoryArchive calls archiveChangeHistoryByDate', () => {
  const ctx = baseContext();
  let called = false;
  ctx.archiveChangeHistoryByDate = () => { called = true; return 'ok'; };
  assert.equal(ctx.runChangeHistoryArchive(), 'ok');
  assert.equal(called, true);
});
test('45. processArchiveMonth preserves existing archiveState object', () => {
  const ctx = baseContext();
  const state = { counters: ctx.emptyArchiveCounters_(), warnings: [] };
  const sourceContext = archiveTestContext(ctx, [['CHG-1', new Date('2026-07-31T10:00:00Z'), 'x']]);
  const existing = new Error('already wrapped');
  existing.archiveState = state;
  ctx.getOrCreateArchiveSpreadsheet_ = () => { throw existing; };
  assert.throws(() => ctx.processArchiveMonth_({}, sourceContext, '2026-07', [{ values: sourceContext.sheet._rows[0] }], state), error => error.archiveState === state && error.message === 'already wrapped');
});

function archiveEntryPointErrorFixture(errorFactory) {
  let criticalSummary = '';
  const ui = {
    Button: { OK: 'OK' },
    ButtonSet: { OK_CANCEL: 'OK_CANCEL', OK: 'OK' },
    prompt() {
      return { getSelectedButton: () => 'OK', getResponseText: () => '01.08.2026' };
    },
    alert() {}
  };
  const ctx = baseContext({ SpreadsheetApp: { getUi: () => ui, flush() {} } });
  ctx.parseArchiveCutoffDate_ = () => new Date('2026-08-01T00:00:00Z');
  ctx.getActiveUserEmail_ = () => 'tester@example.com';
  ctx.generateOperationId_ = () => 'OP-20260804-0001';
  ctx.withDocumentLock_ = callback => callback();
  ctx.archiveChangeHistoryUnderLock_ = function (operationId, startedAt, userEmail, cutoffText, cutoffDate, state) {
    state.counters.checkedRows = 48;
    state.counters.selectedRows = 1;
    state.counters.archiveFiles = 1;
    throw errorFactory(ctx, state);
  };
  ctx.writeArchiveOperationHistory_ = () => {};
  ctx.showCriticalOperationError_ = (title, summary) => { criticalSummary = summary; };
  let caught;
  try { ctx.archiveChangeHistoryByDate(); } catch (error) { caught = error; }
  assert.ok(caught, 'expected archive entry point error');
  return { caught, criticalSummary };
}

test('46. archive entry point report falls back to externally held state', () => {
  const fixture = archiveEntryPointErrorFixture(() => new Error('Повреждён заголовок архива'));
  assert.match(fixture.criticalSummary, /Строк истории проверено: 48/);
  assert.match(fixture.criticalSummary, /Строк выбрано для архивации: 1/);
  assert.match(fixture.criticalSummary, /Строк записано в архив: 0/);
  assert.match(fixture.criticalSummary, /Строк удалено из рабочего файла: 0/);
  assert.match(fixture.criticalSummary, /Архивных файлов: 1/);
  assert.equal(fixture.caught.message, 'Повреждён заголовок архива');
});

test('47. archive entry point still prefers error archiveState', () => {
  const fixture = archiveEntryPointErrorFixture((ctx) => {
    const error = new Error('Ошибка с вложенным состоянием');
    error.archiveState = {
      counters: Object.assign(ctx.emptyArchiveCounters_(), { checkedRows: 7, selectedRows: 2, archiveFiles: 2 }),
      warnings: ['Состояние из ошибки']
    };
    return error;
  });
  assert.match(fixture.criticalSummary, /Строк истории проверено: 7/);
  assert.match(fixture.criticalSummary, /Строк выбрано для архивации: 2/);
  assert.match(fixture.criticalSummary, /Архивных файлов: 2/);
  assert.match(fixture.criticalSummary, /Состояние из ошибки/);
  assert.equal(fixture.caught.message, 'Ошибка с вложенным состоянием');
});

test('48. operator sidebar contains all filters and no alert', () => {
  const text = fs.readFileSync('OperatorSidebar.html', 'utf8');
  for (const id of ['object', 'foreman', 'status', 'type', 'holder', 'from', 'to']) assert.match(text, new RegExp('id=\"' + id + '\"'));
  assert.doesNotMatch(text, /alert\s*\(/);
});

function archiveTestContext(ctx, rows) {
  const headers = ['ID изменения', 'Дата и время изменения', 'Новое значение'];
  return {
    sheet: archiveMockSheet({ headers, config: { dataStartRow: 3, headerRow: 2 } }, rows),
    config: { dataStartRow: 3, headerRow: 2, name: 'История изменений' },
    headers,
    headerMap: Object.fromEntries(headers.map((h, i) => [h, i + 1]))
  };
}
function archiveMockSheet(context, initialRows) {
  const rows = initialRows.map(r => r.slice());
  return {
    getLastRow: () => rows.length + 2,
    getRange(row, column, rowCount, columnCount) {
      return {
        getValues() {
          if (row === 2) return [context.headers.slice(0, columnCount || context.headers.length)];
          return rows.slice(row - 3, row - 3 + rowCount).map(r => r.slice(column - 1, column - 1 + columnCount));
        },
        setValues(values) {
          values.forEach((value, offset) => { rows[row - 3 + offset] = value.slice(); });
        }
      };
    },
    deleteRows(start, count) { rows.splice(start - 3, count); },
    _rows: rows
  };
}


test('49. missing, empty and explicit all-object filters normalize to Все', () => {
  const ctx = baseContext();
  for (const value of [undefined, {}, { allObjects: true, objectId: 'ignored' }, { objectName: ' Все ' }]) {
    const result = ctx.operatorCardNormalizeObjectFilter_(value);
    assert.equal(result.allObjects, true); assert.equal(result.objectId, ''); assert.equal(result.objectName, 'Все');
  }
});
test('50. concrete object requires and preserves ID', () => {
  const ctx = baseContext();
  assert.throws(() => ctx.operatorCardNormalizeObjectFilter_({ allObjects: false }), /ID объекта/);
  const result = ctx.operatorCardNormalizeObjectFilter_({ allObjects: false, objectId: ' 3 ', objectName: 'Дом' });
  assert.equal(result.objectId, '3'); assert.equal(result.allObjects, false);
});
test('51. programmatic Все is unique, first, and real objects are naturally sorted', () => {
  const ctx = baseContext();
  const result = ctx.operatorCardBuildObjects_([{ id: '10', name: 'Дом 10' }, { id: '2', name: 'Дом 2' }]);
  assert.deepEqual(Array.from(result, x => x.name), ['Все', 'Дом 2', 'Дом 10']);
  assert.equal(result.filter(x => x.isAllObjects).length, 1);
});
test('52. real object named Все is rejected', () => {
  const ctx = baseContext();
  assert.throws(() => ctx.operatorCardBuildObjects_([{ id: '1', name: ' ВСЕ ' }]), /системным названием/);
});
test('53. unknown IDs are rejected by server validation', () => {
  const ctx = baseContext();
  const filters = ctx.operatorCardNormalizeFilters_({ object: { allObjects: false, objectId: '404', objectName: 'X' } });
  assert.throws(() => ctx.operatorCardValidateSelection_(filters, { objects: [], foremen: [], employees: [], documentTypes: [] }), /неизвестный объект/);
});
test('54. header contract accepts extra fact columns but rejects reordering', () => {
  const ctx = baseContext();
  const headers24 = ['ID документа','ID объекта','Название объекта','Тип документа','Номер договора','Номер документа','Дата документа','Статус документа','Оригинал / ЭДО','Комментарий','У кого документ','Где документ','Кто передал','Кто ответственный за подписание (заказчик)','Оплачен','Сумма документа','ГУ (Да/Нет)','Условия ГУ','Статус объекта','Ответственный прораб','Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)','Дата создания','Дата обновления','ID типа документа','Кто обновил (email)']; const headers=headers24.concat(['Статус записи','Номер строки в таблице фактов']);
  assert.equal(ctx.operatorCardValidateHeaders_(headers.concat(['extra']), headers, 'Документы объектов', 'Карточка'), true);
  const swapped = headers.slice(); [swapped[0], swapped[1]] = [swapped[1], swapped[0]];
  assert.throws(() => ctx.operatorCardValidateHeaders_(swapped, headers, 'Документы объектов', 'Карточка'), /позиция 2/);
  assert.throws(() => ctx.operatorCardValidateHeaders_(headers24.concat(['Техническое поле']), headers.concat(['Лишний заголовок']), 'Документы объектов', 'Карточка'), /позиция 27|колонка 30/);
  assert.equal(ctx.operatorCardValidateHeaders_(headers24.concat(['Техническое поле']), headers, 'Документы объектов', 'Карточка'), true);
});
test('55. active rows, combined ID filters and normalized status are applied', () => {
  const ctx = baseContext();
  const indexes = { documentId:0, objectId:1, documentTypeId:2, documentStatus:3, holderId:4, foremanId:5, createdAt:6, recordStatus:7,transferredById:30,documentHolder:9,transferredBy:11 };
  const rows = [
    ['DOC-2','2','T1',' ГОТОВ ','E1','F1','', ' Активная '],
    ['DOC-3','2','T1','готов','E1','F1','', 'Архивная'],
    ['DOC-4','2','T2','готов','E1','F1','', 'Удалённая']
  ]; while(rows[0].length<31) rows.forEach(r=>r.push(''));
  const filters = ctx.operatorCardNormalizeFilters_({ object:{allObjects:false,objectId:'2'}, documentTypeId:'T1', documentStatus:'готов', holderId:'E1', foremanId:'F1' });
  const result = ctx.operatorCardPrepareRows_(rows,indexes,filters,{active:false},4);
  assert.equal(result.activeCount,1); assert.equal(result.rows.length,1); assert.equal(result.cardRows[0].length,29);
});
test('56. natural sorting uses object, type, document and physical row', () => {
  const ctx = baseContext(); const i={documentId:0,objectId:1,documentTypeId:2,documentStatus:3,holderId:4,foremanId:5,createdAt:6,recordStatus:7,transferredById:30,documentHolder:9,transferredBy:11};
  const rows=[['DOC-10','2','10','','','','','Активная'],['DOC-2','2','2','','','','','Активная'],['DOC-1','10','1','','','','','Активная']]; rows.forEach(r=>{while(r.length<31)r.push('')});
  const result=ctx.operatorCardPrepareRows_(rows,i,ctx.operatorCardNormalizeFilters_({}),{active:false},4);
  assert.deepEqual(Array.from(result.rows,x=>x.row[0]),['DOC-2','DOC-10','DOC-1']);
});
test('57. active duplicate group is retained and reports physical rows once per ID', () => {
  const ctx=baseContext(); const i={documentId:0,objectId:1,documentTypeId:2,documentStatus:3,holderId:4,foremanId:5,createdAt:6,recordStatus:7,transferredById:30,documentHolder:9,transferredBy:11};
  const rows=[['D1','A','','','','','','Активная'],['D1','B','','','','','','Активная'],['D1','A','','','','','','Архивная']]; rows.forEach(r=>{while(r.length<31)r.push('')});
  const result=ctx.operatorCardPrepareRows_(rows,i,ctx.operatorCardNormalizeFilters_({object:{allObjects:false,objectId:'A'}}),{active:false},18);
  assert.equal(result.rows.length,1); assert.equal(result.duplicateIdsCount,1); assert.match(result.warnings[0],/18 и 19/); assert.doesNotMatch(result.warnings[0],/20/);
});
test('58. card replacement performs one 26-column setValues and clears tail', () => {
  const ctx=baseContext(); let calls=0, written;
  const sheet={getLastRow:()=>8,getRange(row,col,count,width){return {getValues:()=>[['old1'],['old2'],['']],setValues(values){calls++;written={row,col,count,width,values}}}}};
  ctx.operatorCardReplace_({sheet,config:{dataStartRow:6,requiredHeaders:Array(29)},headerMap:{'ID документа':1}},[[1,2].concat(Array(27).fill(''))]);
  assert.equal(calls,1); assert.equal(written.width,29); assert.equal(written.values.length,2); assert.ok(written.values[1].every(v=>v===''));
});
test('59. zero result clears old card in one batch', () => {
  const ctx=baseContext(); let values; const sheet={getLastRow:()=>6,getRange(){return {getValues:()=>[['old']],setValues(v){values=v}}}};
  ctx.operatorCardReplace_({sheet,config:{dataStartRow:6,requiredHeaders:Array(29)},headerMap:{'ID документа':1}},[]); assert.equal(values.length,1); assert.equal(values[0].length,29); assert.ok(values[0].every(v=>v===''));
});
test('60. operator implementation never writes facts, dictionaries, or change history', () => {
  const text=fs.readFileSync('OperatorCard.gs','utf8');
  assert.doesNotMatch(text,/appendRow|\.clear\s*\(|deleteRows|insertRows/);
  assert.doesNotMatch(text,/getSystemSheetContext_\('CHANGE_HISTORY'\)/);
  assert.match(text,/getRange\(start, cardStartColumn, writeCount, cardWidth\)\.setValues/);
});


test('61. date ranges support empty, from-only, to-only and reject reversed dates', () => {
  const ctx=baseContext();
  assert.equal(ctx.operatorCardParseDateRange_('','','Europe/Berlin').active,false);
  assert.ok(ctx.operatorCardParseDateRange_('2026-03-29','','Europe/Berlin').from);
  assert.ok(ctx.operatorCardParseDateRange_('','2026-10-25','Europe/Berlin').toExclusive);
  assert.throws(()=>ctx.operatorCardParseDateRange_('2026-08-11','2026-08-10','UTC'),/Дата от/);
});
test('62. inclusive date boundaries use spreadsheet timezone and next calendar day', () => {
  const ctx=baseContext(); const range=ctx.operatorCardParseDateRange_('2026-03-29','2026-03-29','Europe/Berlin');
  assert.equal(range.from.toISOString(),'2026-03-28T23:00:00.000Z');
  assert.equal(range.toExclusive.toISOString(),'2026-03-29T22:00:00.000Z');
  assert.equal(range.toExclusive-range.from,23*60*60*1000);
  assert.doesNotMatch(fs.readFileSync('OperatorCard.gs','utf8'),/24\s*\*\s*60\s*\*\s*60|86400000/);
});
test('63. active date filter excludes blank and warns about invalid nonblank dates', () => {
  const ctx=baseContext(); const i={documentId:0,objectId:1,documentTypeId:2,documentStatus:3,holderId:4,foremanId:5,createdAt:6,recordStatus:7,transferredById:30,documentHolder:9,transferredBy:11};
  const valid=vm.runInContext("new Date('2026-08-10T12:00:00Z')",ctx);
  const rows=[['A','','','','','',valid,'Активная'],['B','','','','','','','Активная'],['C','','','','','','bad','Активная']];rows.forEach(r=>{while(r.length<31)r.push('')});
  const range=ctx.operatorCardParseDateRange_('2026-08-10','2026-08-10','UTC');
  const result=ctx.operatorCardPrepareRows_(rows,i,ctx.operatorCardNormalizeFilters_({}),range,4);
  assert.equal(result.rows.length,1);assert.equal(result.invalidDatesCount,1);assert.match(result.warnings[0],/некорректной датой/);
});


const operatorWorkflowCardHeaders = ['ID документа','ID объекта','Название объекта','Тип документа','Номер договора','Номер документа','Дата документа','Статус документа','Оригинал / ЭДО','Комментарий','У кого документ','Где документ','Кто передал','Кто ответственный за подписание (заказчик)','Оплачен','Сумма документа','ГУ (Да/Нет)','Условия ГУ','Статус объекта','Ответственный прораб','Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)','Дата создания','Дата обновления','ID типа документа','Кто обновил (email)'].concat(['Статус записи','Номер строки в таблице фактов']);
const operatorWorkflowDocumentHeaders = operatorWorkflowCardHeaders.slice(0,27).concat(['ID сотрудника — у кого документ','ID сотрудника — кто передал','ID ответственного прораба','Дата изменения статуса документа','Отчётный период','Источник создания','Статус записи']);
const operatorWorkflowOperationHeaders = ['ID операции','Дата и время начала','Дата и время завершения','Кто запустил (email)','Источник операции','Тип операции','Статус операции','Документов загружено в карточку','Документов с изменениями','Строк факта обновлено','Полей изменено','Дублирующихся ID найдено','Ошибок','Время выполнения, сек.','Текст ошибки / комментарий'];

function operatorWorkflowFixture(options = {}) {
  const ctx = baseContext();
  const calls = { cardWrites: 0, operationRows: [], sheetKeys: [] };
  const documentMap = Object.fromEntries(operatorWorkflowDocumentHeaders.map((header, index) => [header, index + 1]));
  const documents = {
    config: { name: 'Документы объектов', headerRow: 3, dataStartRow: 4 }, headers: operatorWorkflowDocumentHeaders.slice(), headerMap: documentMap,
    sheet: { getLastRow: () => 3, getRange() { throw new Error('fact data range must not be read for empty fixture'); } }
  };
  const card = { config: { name: 'Карточка операциониста', headerRow: 5, dataStartRow: 6 }, headers: operatorWorkflowCardHeaders.slice(), headerMap: Object.fromEntries(operatorWorkflowCardHeaders.map((h,i)=>[h,i+1])), sheet: {} };
  const operation = {
    config: { name: 'История операций', headerRow: 2, dataStartRow: 3 }, headers: operatorWorkflowOperationHeaders.slice(),
    sheet: { getLastRow: () => 2, getRange() { return { setValues(rows) { calls.operationRows.push(rows[0].slice()); } }; } }
  };
  ctx.withDocumentLock_ = callback => callback();
  ctx.generateOperationId_ = () => 'OP-20260810-0001';
  ctx.getActiveUserEmail_ = () => 'tester@example.com';
  ctx.getSystemSpreadsheet_ = () => ({ getSpreadsheetTimeZone: () => 'UTC' });
  ctx.operatorCardGetFilterData_ = () => ({ objects: [{ id:'', name:'Все', isAllObjects:true }], foremen: [], employees: [], documentTypes: [] });
  ctx.operatorCardValidateSelection_ = options.validationError ? () => { throw new Error(options.validationError); } : () => {};
  ctx.operatorCardPrepareRows_ = () => ({ activeCount: options.activeCount == null ? 2 : options.activeCount, rows: [], cardRows: Array.from({length: options.loadedCount == null ? 2 : options.loadedCount}, (_, index) => [index + 1].concat(Array(26).fill(''))), warnings: (options.warnings || []).slice(), duplicateIdsCount: options.duplicateIdsCount || 0, invalidDatesCount: 0 });
  ctx.operatorCardGetValidationData_ = () => [];
  ctx.operatorCardApplyValidations_ = () => {};
  ctx.operatorCardReplace_ = () => { calls.cardWrites++; };
  ctx.getSystemSheetContext_ = key => {
    calls.sheetKeys.push(key);
    if (key === 'DOCUMENTS') return documents;
    if (key === 'OPERATOR_CARD') return card;
    if (key === 'OPERATION_HISTORY') {
      if (options.historyError) throw new Error(options.historyError);
      return operation;
    }
    throw new Error('Unexpected sheet key: ' + key);
  };
  return { ctx, calls };
}

function operatorWorkflowAssertCounters(row, expectedLoaded) {
  assert.equal(row[7], expectedLoaded, 'real loaded count');
  assert.equal(row[8], 0, 'documents changed');
  assert.equal(row[9], 0, 'fact rows updated');
  assert.equal(row[10], 0, 'fields changed');
}

test('64. full workflow success writes card and exactly one successful operation', () => {
  const fixture = operatorWorkflowFixture({ loadedCount: 2 });
  const response = fixture.ctx.operatorCardApply_({ object: {} });
  assert.equal(response.status, 'Успешно'); assert.equal(response.loadedCount, 2);
  assert.equal(fixture.calls.cardWrites, 1); assert.equal(fixture.calls.operationRows.length, 1);
  assert.equal(fixture.calls.operationRows[0][6], 'Успешно');
  operatorWorkflowAssertCounters(fixture.calls.operationRows[0], 2);
  assert.match(fixture.calls.operationRows[0][14], /Объект: Все/);
  assert.ok(!fixture.calls.sheetKeys.includes('CHANGE_HISTORY'));
});

test('65. full workflow duplicate warning writes exactly one warning operation', () => {
  const warning = 'ID документа «D-1» повторяется в активных строках 4 и 5.';
  const fixture = operatorWorkflowFixture({ loadedCount: 2, warnings: [warning], duplicateIdsCount: 1 });
  const response = fixture.ctx.operatorCardApply_({});
  assert.equal(response.status, 'Успешно с предупреждениями'); assert.equal(response.duplicateIdsCount, 1);
  assert.equal(fixture.calls.operationRows.length, 1); assert.equal(fixture.calls.operationRows[0][6], 'Успешно с предупреждениями');
  assert.equal(fixture.calls.operationRows[0][11], 1); assert.match(fixture.calls.operationRows[0][14], /D-1/);
  operatorWorkflowAssertCounters(fixture.calls.operationRows[0], 2);
});

test('66. full workflow zero result writes one no-changes operation with real zero count', () => {
  const fixture = operatorWorkflowFixture({ loadedCount: 0, activeCount: 3 });
  const response = fixture.ctx.operatorCardApply_({});
  assert.equal(response.status, 'Без изменений'); assert.equal(response.loadedCount, 0);
  assert.equal(fixture.calls.cardWrites, 1); assert.equal(fixture.calls.operationRows.length, 1);
  assert.equal(fixture.calls.operationRows[0][6], 'Без изменений'); operatorWorkflowAssertCounters(fixture.calls.operationRows[0], 0);
});

test('67. critical pre-write error preserves card and creates one best-effort error operation', () => {
  const fixture = operatorWorkflowFixture({ validationError: 'Справочник повреждён' });
  assert.throws(() => fixture.ctx.operatorCardApply_({}), /Справочник повреждён.*Карточка не была изменена/);
  assert.equal(fixture.calls.cardWrites, 0); assert.equal(fixture.calls.operationRows.length, 1);
  assert.equal(fixture.calls.operationRows[0][6], 'Ошибка'); assert.equal(fixture.calls.operationRows[0][12], 1);
  operatorWorkflowAssertCounters(fixture.calls.operationRows[0], 0);
});

test('68. filter normalization error uses common error path and logs Все without changing card', () => {
  const fixture = operatorWorkflowFixture();
  assert.throws(() => fixture.ctx.operatorCardApply_({ object: { allObjects: false } }), /ID объекта.*Карточка не была изменена/);
  assert.equal(fixture.calls.cardWrites, 0); assert.equal(fixture.calls.operationRows.length, 1);
  assert.equal(fixture.calls.operationRows[0][6], 'Ошибка'); assert.match(fixture.calls.operationRows[0][14], /Объект: Все/);
});

test('69. history failure after successful card write keeps card and exposes original cause', () => {
  const fixture = operatorWorkflowFixture({ loadedCount: 1, historyError: 'Лист журнала защищён' });
  const originalError = console.error; const errors = []; console.error = message => errors.push(message);
  let response;
  try { response = fixture.ctx.operatorCardApply_({}); } finally { console.error = originalError; }
  assert.equal(fixture.calls.cardWrites, 1); assert.equal(response.success, true); assert.equal(response.status, 'Успешно с предупреждениями');
  assert.match(response.warnings.join(' '), /Карточка загружена.*Лист журнала защищён/);
  assert.match(errors.join(' '), /Лист журнала защищён/);
  assert.equal(fixture.calls.operationRows.length, 0);
});


test('70. shifted facts and independently shifted card use their ID-document start columns', () => {
  const ctx = baseContext();
  const factHeaders = [''].concat(operatorWorkflowDocumentHeaders);
  const cardHeaders = ['', ''].concat(operatorWorkflowCardHeaders).concat(['', '  ']);
  assert.equal(ctx.operatorCardValidateHeaders_(factHeaders, cardHeaders, 'Документы объектов', 'Карточка', 2, 3), true);

  const workValues = operatorWorkflowDocumentHeaders.map(header => {
    if (header === 'ID документа') return 'DOC-7';
    if (header === 'ID объекта') return 'OBJ-2';
    if (header === 'ID типа документа') return 'TYPE-3';
    if (header === 'Статус записи') return 'Активная';
    return header;
  });
  const factRow = ['служебное значение слева'].concat(workValues);
  const physicalIndexes = Object.fromEntries(operatorWorkflowDocumentHeaders.map((header, index) => [header, index + 1]));
  const prepared = ctx.operatorCardPrepareRows_([factRow], {
    documentId: physicalIndexes['ID документа'], objectId: physicalIndexes['ID объекта'],
    documentTypeId: physicalIndexes['ID типа документа'], documentStatus: physicalIndexes['Статус документа'],
    holderId: physicalIndexes['ID сотрудника — у кого документ'], foremanId: physicalIndexes['ID ответственного прораба'],
    createdAt: physicalIndexes['Дата создания'], recordStatus: physicalIndexes['Статус записи'], transferredById: physicalIndexes['ID сотрудника — кто передал'],
    documentHolder: physicalIndexes['У кого документ'], transferredBy: physicalIndexes['Кто передал']
  }, ctx.operatorCardNormalizeFilters_({}), {active:false}, 4, 1);
  assert.equal(prepared.cardRows.length, 1);
  assert.equal(prepared.cardRows[0].length, 29);
  assert.equal(prepared.cardRows[0][0], 'DOC-7');
  assert.ok(!prepared.cardRows[0].includes('служебное значение слева'));

  const rangeCalls = [];
  const sheet = {
    getLastRow: () => 6,
    getRange(row, column, rowCount, columnCount) {
      rangeCalls.push({row, column, rowCount, columnCount});
      return { getValues: () => [['OLD-DOC']], setValues(values) { this.values = values; } };
    }
  };
  ctx.operatorCardReplace_({sheet, config:{dataStartRow:6,requiredHeaders:Array(29)}, headerMap:{'ID документа':3}}, prepared.cardRows);
  assert.equal(rangeCalls.length, 2);
  assert.deepEqual(rangeCalls.map(call => call.column), [3, 3]);
  assert.equal(rangeCalls[0].columnCount, 1);
  assert.equal(rangeCalls[1].columnCount, 29);
  assert.ok(rangeCalls.every(call => call.column >= 3), 'columns left of the card block must not be touched');
});


function operatorDictionaryContext(headers, dataStartRow, rows) {
  const headerMap = Object.fromEntries(headers.map((header, index) => [header, index + 1]).filter(entry => entry[0]));
  return {
    headers, headerMap, config: {dataStartRow},
    sheet: {
      getLastRow: () => dataStartRow + rows.length - 1,
      getRange(row, column, rowCount, columnCount) {
        return {getValues: () => rows.slice(0, rowCount).map(values => values.slice(column - 1, column - 1 + columnCount))};
      }
    }
  };
}

test('71. filter data reads shifted real objects, test foremen, and unique card statuses', () => {
  const ctx = baseContext();
  const contexts = {
    OBJECTS: operatorDictionaryContext(['','ID объекта','Название объекта'], 3, [
      ['', '10', 'Дом 10'], ['', '2', 'Дом 2']
    ]),
    EMPLOYEES: operatorDictionaryContext(['','ID Сотрудника','ФИО сотрудника','Должность'], 5, [
      ['', 'E1', 'Иванов', '  гЛаВнЫй\u00a0 ИнЖеНеР '], ['', 'E2', 'Петров', 'Бухгалтер']
    ]),
    CLIENTS: operatorDictionaryContext(['','ID клиента','Наименование клиента'], 5, []),
    DOCUMENT_TYPES: operatorDictionaryContext(['','ID типа документа','Тип документа'], 5, [
      ['', 'T1', 'Акт']
    ]),
    CARD_DICTIONARY: operatorDictionaryContext(['','Статус документа'], 3, [
      ['', ' Подписан '], ['', ''], ['', 'ПОДПИСАН'], ['', 'На согласовании']
    ])
  };
  ctx.getSystemSheetContext_ = key => contexts[key];
  const result = ctx.operatorCardGetFilterData_();
  assert.deepEqual(Array.from(result.objects, item => item.name), ['Все','Дом 2','Дом 10']);
  assert.equal(result.objects[0].isAllObjects, true);
  assert.equal(result.objects.filter(item => item.isAllObjects).length, 1);
  assert.deepEqual(Array.from(result.objects.slice(1), item => ({id:item.id,name:item.name,isAllObjects:item.isAllObjects})), [
    {id:'2',name:'Дом 2',isAllObjects:false},{id:'10',name:'Дом 10',isAllObjects:false}
  ]);
  assert.deepEqual(Array.from(result.foremen, item => item.id), ['E1']);
  assert.deepEqual(Array.from(result.documentStatuses), ['Все','Подписан','На согласовании']);
});

test('72. document status server validation accepts dictionary values only', () => {
  const ctx = baseContext();
  const data = {objects:[{id:'2'}],foremen:[],documentTypes:[],employees:[],holders:[],documentStatuses:['Все','Подписан']};
  const valid = ctx.operatorCardNormalizeFilters_({documentStatus:'  пОдПиСаН '});
  assert.doesNotThrow(() => ctx.operatorCardValidateSelection_(valid, data));
  const all = ctx.operatorCardNormalizeFilters_({documentStatus:'Все'});
  assert.equal(all.documentStatus, '');
  assert.throws(() => ctx.operatorCardValidateSelection_(ctx.operatorCardNormalizeFilters_({documentStatus:'Удалён'}), data), /неизвестный статус/);
});

test('73. validation dictionaries use configured sources and preserve employee-before-client order', () => {
  const ctx = baseContext();
  const values = {
    'EMPLOYEES|ФИО сотрудника':['Сотрудник 1 [ST-1]','Сотрудник 2 [ST-2]'],
    'CLIENTS|Наименование клиента':['Клиент 1'],
    'CARD_DICTIONARY|Статус документа':['Подписан'],
    'CARD_DICTIONARY|Оригинал / ЭДО':['Оригинал','ЭДО'],
    'CARD_DICTIONARY|Где документ':['Мытищи'],
    'CARD_DICTIONARY|Оплачен':['Оплачен'],
    'CARD_DICTIONARY|ГУ (Да/Нет)':['Да','Нет'],
    'CARD_DICTIONARY|Статус записи':['Активная','Архивная','Удалённая']
  };
  ctx.operatorCardReadUniqueColumn_ = (key, header) => values[key+'|'+header].slice();
  const result = ctx.operatorCardGetValidationData_({employees:[{id:'ST-1',name:'Сотрудник 1'},{id:'ST-2',name:'Сотрудник 2'}],holders:[{id:'ST-1',name:'Сотрудник 1'},{id:'ST-2',name:'Сотрудник 2'},{id:'CL-1',name:'Клиент 1'}]});
  const byHeader = Object.fromEntries(Array.from(result, item => [item.header, Array.from(item.values)]));
  assert.deepEqual(byHeader['Статус документа'], ['Подписан']);
  assert.deepEqual(byHeader['Оригинал / ЭДО'], ['Оригинал','ЭДО']);
  assert.deepEqual(byHeader['У кого документ'], ['Сотрудник 1 [ST-1]','Сотрудник 2 [ST-2]','Клиент 1 [CL-1]']);
  assert.deepEqual(byHeader['Где документ'], ['Мытищи']);
  assert.deepEqual(byHeader['Кто передал'], ['Сотрудник 1 [ST-1]','Сотрудник 2 [ST-2]']);
  assert.deepEqual(byHeader['Оплачен'], ['Оплачен']);
  assert.deepEqual(byHeader['ГУ (Да/Нет)'], ['Да','Нет']);
});

test('74. card validations use headerMap ranges and survive reload, zero result, and repeat load', () => {
  const ctx = baseContext();
  const validationColumns = {'Статус документа':8,'Оригинал / ЭДО':9,'У кого документ':11,'Где документ':12,'Кто передал':13,'Оплачен':14,'ГУ (Да/Нет)':16};
  const validations = {};
  const writes = [];
  ctx.SpreadsheetApp.newDataValidation = () => {
    const state = {};
    return {requireValueInList(values, showDropdown){state.values=values.slice();state.showDropdown=showDropdown;return this},setAllowInvalid(value){state.allowInvalid=value;return this},build(){return state}};
  };
  const sheet = {
    getMaxRows: () => 20,
    getLastRow: () => 6,
    getRange(row,column,rowCount,columnCount) {
      return {
        getValues: () => [['OLD']],
        setValues(values) { writes.push({row,column,rowCount,columnCount,values}); },
        setDataValidation(rule) { validations[column] = {row,rowCount,columnCount,rule}; }
      };
    }
  };
  const context = {sheet,config:{dataStartRow:6,requiredHeaders:Array(29)},headerMap:Object.assign({'ID документа':3},validationColumns)};
  const data = Object.keys(validationColumns).map(header => ({header,values:[header+' value']}));
  ctx.operatorCardApplyValidations_(context, data);
  assert.deepEqual(Object.keys(validations).map(Number).sort((a,b)=>a-b), [8,9,11,12,13,14,16]);
  assert.ok(Object.values(validations).every(item => item.row===6 && item.rowCount===15 && item.columnCount===1));
  assert.ok(Object.values(validations).every(item => item.rule.allowInvalid===true));
  const snapshot = JSON.stringify(validations);
  const row = ['DOC'].concat(Array(26).fill(''));
  ctx.operatorCardReplace_(context,[row]);
  ctx.operatorCardReplace_(context,[]);
  ctx.operatorCardReplace_(context,[row]);
  assert.equal(JSON.stringify(validations), snapshot, 'setValues must preserve validations');
  assert.equal(writes.length,3);
  assert.ok(writes.every(write => write.column===3 && write.columnCount===29));
});

test('75. sidebar groups creation dates and reset is local-only for all filters', () => {
  const text = fs.readFileSync('OperatorSidebar.html','utf8');
  assert.match(text, /<div class="group-label">Дата создания<\/div>/);
  assert.match(text, /for="from">От<\/label><input id="from" type="date"/);
  assert.match(text, /for="to">До<\/label><input id="to" type="date"/);
  assert.match(text, />Очистить фильтры<\/button>/);
  for (const id of ['object','foreman','status','type','holder']) {
    assert.match(text, new RegExp('id="' + id + '"[^>]*placeholder="Все"'));
  }
  assert.match(text, /function resetFilters\(\)\{\['object','foreman','status','type','holder','from','to'\]\.forEach\(id=>el\(id\)\.value=''\)/);
  const resetBody = text.match(/function resetFilters\(\)\{([\s\S]*?)\}\nel\('reset'\)/)[1];
  assert.doesNotMatch(resetBody,/google\.script\.run|applyOperatorFilters/);
  assert.match(text,/fillValues\('statuses',data\.documentStatuses\|\|\[\]\)/);
  assert.match(text,/items\.filter\(item=>!item\.isAllObjects\)/);
  assert.match(text,/values\.filter\(value=>value!=='Все'\)/);
  assert.doesNotMatch(text,/\.value='Все'/);
});

test('76. empty sidebar filter contract means all objects and no additional restrictions', () => {
  const ctx = baseContext();
  const filters = ctx.operatorCardNormalizeFilters_({object:{},foremanId:'',documentStatus:'',documentTypeId:'',holderId:'',dateFrom:'',dateTo:''});
  assert.equal(filters.object.allObjects,true);
  assert.equal(filters.foremanId,''); assert.equal(filters.documentStatus,'');
  assert.equal(filters.documentTypeId,''); assert.equal(filters.holderId,'');
  const indexes={documentId:0,objectId:1,documentTypeId:2,documentStatus:3,holderId:4,foremanId:5,createdAt:6,recordStatus:7,transferredById:30,documentHolder:9,transferredBy:11};
  const rows=[['D1','OBJ1','T1','Подписан','E1','F1','','Активная'],['D2','OBJ2','T2','Новый','E2','F2','','Активная']];
  rows.forEach(row=>{while(row.length<31)row.push('')});
  const result=ctx.operatorCardPrepareRows_(rows,indexes,filters,{active:false},4,0);
  assert.equal(result.rows.length,2);
});

test('77. save card contract has unchanged first 24 fields plus status and physical row', () => {
  const ctx=baseContext(); const headers=Array.from(ctx.SYSTEM_CONFIG ? ctx.SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.requiredHeaders : []);
  // top-level const is lexical in vm; inspect source for the runtime contract instead.
  const text=fs.readFileSync('SystemCore.gs','utf8');
  assert.match(text,/H\.UPDATED_BY_EMAIL,\s*H\.RECORD_STATUS,\s*H\.FACT_ROW_NUMBER/);
  assert.match(text,/cardHeader: H\.CONTRACT_NUMBER,[\s\S]{0,80}editable: false/);
});
test('78. save endpoint is thin and critical workflow uses document lock', () => {
  assert.match(fs.readFileSync('Code.gs','utf8'),/saveOperatorCardChanges\(filters\)[\s\S]*return operatorCardSave_\(filters\)/);
  assert.match(fs.readFileSync('OperatorCardSave.gs','utf8'),/withDocumentLock_\(function/);
});
test('79. loader copies 25 fact fields then appends status and physical row', () => {
  const text=fs.readFileSync('OperatorCard.gs','utf8');
  assert.match(text,/slice\(sourceStartIndex, sourceStartIndex \+ sharedFieldCount\)/);
  assert.match(text,/concat\(\[item\.row\[indexes\.recordStatus\], item\.sheetRow\]\)/);
  assert.match(text,/hideColumns\(cardStartColumn, 2\)/);
});
test('80. holder dictionary is employee then client and labels contain IDs', () => {
  const text=fs.readFileSync('OperatorCard.gs','utf8');
  assert.match(text,/employees\.map[\s\S]*\.concat\(clients\.map/);
  assert.equal(baseContext().operatorCardDisplayLabel_('ООО Ромашка','CL-0001'),'ООО Ромашка [CL-0001]');
});
test('81. save module resolves holder from combined map and transfer only from employees', () => {
  const text=fs.readFileSync('OperatorCardSave.gs','utf8');
  assert.match(text,/holderMap\[value\][\s\S]*HOLDER_EMPLOYEE_ID/);
  assert.match(text,/employeeMap\[value\][\s\S]*TRANSFERRED_BY_EMPLOYEE_ID/);
  assert.match(text,/value = ''[\s\S]*HOLDER_EMPLOYEE_ID\] = ''/);
});
test('82. save validates optimistic version and all row identity keys before writes', () => {
  const text=fs.readFileSync('OperatorCardSave.gs','utf8');
  for(const token of ['H.DOCUMENT_ID','H.OBJECT_ID','H.DOCUMENT_TYPE_ID','H.UPDATED_AT','H.FACT_ROW_NUMBER']) assert.ok(text.includes(token));
  assert.ok(text.indexOf('operatorCardBuildSavePlan_') < text.indexOf('operatorCardSaveWriteFacts_'));
});
test('83. duplicate resolution requires complete group, status-only and at most one active', () => {
  const text=fs.readFileSync('OperatorCardSave.gs','utf8');
  assert.match(text,/представлен в карточке не полностью/); assert.match(text,/разрешено менять только/); assert.match(text,/remaining > 1/);
});
test('84. fact writer whitelist excludes contract and includes derived technical fields', () => {
  const text=fs.readFileSync('OperatorCardSave.gs','utf8');
  const list=text.match(/const OPERATOR_CARD_SAVE_WRITABLE_ = \[([\s\S]*?)\];/)[1];
  assert.ok(!list.includes('CONTRACT_NUMBER')); assert.ok(list.includes('RECORD_STATUS'));
  assert.match(text,/H\.UPDATED_AT, H\.UPDATED_BY_EMAIL, H\.DOCUMENT_STATUS_CHANGED_AT/);
});
test('85. save histories use centralized operation source and edit action', () => {
  const core=fs.readFileSync('SystemCore.gs','utf8'),save=fs.readFileSync('OperatorCardSave.gs','utf8');
  assert.match(core,/OPERATOR_CARD_SAVE_OPERATION_TYPE/); assert.match(core,/CHANGE_ACTION_EDIT/);
  assert.match(save,/SYSTEM_CONFIG\.VALUES\.CHANGE_ACTION_EDIT/); assert.match(save,/SYSTEM_CONFIG\.VALUES\.OPERATOR_CARD_SAVE_OPERATION_TYPE/);
});
test('86. sidebar preserves lastAppliedFilters and reset remains local-only', () => {
  const text=fs.readFileSync('OperatorSidebar.html','utf8');
  assert.match(text,/let lastAppliedFilters=null/); assert.match(text,/lastAppliedFilters=filters/); assert.match(text,/saveOperatorCardChanges\(lastAppliedFilters\)/);
  const reset=text.match(/function resetFilters\(\)\{([\s\S]*?)\}\nel\('reset'\)/)[1]; assert.ok(!reset.includes('lastAppliedFilters'));
});

function partialSaveFixture(ctx, options = {}) {
  const updated = options.updated || 'v1';
  const documentId = options.documentId || 'DOC-1';
  const factRow = options.factRow || 10;
  const base = {
    'ID документа':documentId,'ID объекта':'OBJ-1','Название объекта':'Объект Альфа','ID типа документа':'TYPE-1',
    'Тип документа':'Акт','Номер договора':'CN-1','Дата документа':'',
    'Статус документа':'Новый','Оригинал / ЭДО':'Оригинал','Комментарий':'old',
    'У кого документ':'','Где документ':'Офис','Кто передал':'','Оплачен':'Нет',
    'Сумма документа':'','ГУ (Да/Нет)':'Нет','Условия ГУ':'','Статус объекта':'Работа',
    'Дата начала работ':'','Дата окончания (по плану)':'','Дата окончания (по факту)':'',
    'Дата создания':'created','Дата обновления':updated,'Кто обновил (email)':'old@example.com',
    'Ответственный прораб':'Иванов','Статус записи':'Активная',
    'ID сотрудника — у кого документ':'','ID сотрудника — кто передал':''
  };
  const fact = Object.assign({},base,options.fact||{});
  const card = Object.assign({},base,options.card||{}, {'Номер строки в таблице фактов':factRow});
  return {
    card:{sheetRow:options.cardRow||6,values:card},
    fact:{sheetRow:factRow,values:fact}
  };
}
function partialDictionaries() {
  return {
    holders:[{id:'ST-1',name:'Иванов',type:'employee'},{id:'CL-1',name:'ООО Ромашка',type:'client'}],
    employees:[{id:'ST-1',name:'Иванов'}],
    card:{
      'Статус документа':['Новый','Готов'], 'Оригинал / ЭДО':['Оригинал','ЭДО'],
      'Где документ':['Офис','Архив'], 'Оплачен':['Нет','Да'], 'ГУ (Да/Нет)':['Нет','Да'],
      'Статус записи':['Активная','Архивная','Удалённая']
    }
  };
}

test('87. one invalid row does not reject an unrelated valid row', () => {
  const ctx=baseContext();
  const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const invalid=partialSaveFixture(ctx,{documentId:'DOC-2',factRow:99,card:{'Комментарий':'bad'}});
  const plan=ctx.operatorCardBuildSavePlan_([valid.card,invalid.card],[valid.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.deepEqual(Array.from(plan.rows,p=>p.documentId),['DOC-1']); assert.equal(plan.rowErrors.length,1);
});
test('88. stale and invalid dictionary rows are local while valid changes survive', () => {
  const ctx=baseContext();
  const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const stale=partialSaveFixture(ctx,{documentId:'DOC-2',factRow:11,updated:'old',fact:{'Дата обновления':'new'},card:{'Комментарий':'edit'}});
  const dictionary=partialSaveFixture(ctx,{documentId:'DOC-3',factRow:12,card:{'Статус документа':'UNKNOWN'}});
  const plan=ctx.operatorCardBuildSavePlan_([valid.card,stale.card,dictionary.card],[valid.fact,stale.fact,dictionary.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.equal(plan.rows.length,1); assert.equal(plan.rowErrors.length,2);
  assert.match(plan.rowErrors.map(x=>x.message).join(' '),/изменились после загрузки/);
});
test('89. an untouched stale row produces no error', () => {
  const ctx=baseContext(); const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const stale=partialSaveFixture(ctx,{documentId:'DOC-2',factRow:11,updated:'old',fact:{'Дата обновления':'new'}});
  const plan=ctx.operatorCardBuildSavePlan_([valid.card,stale.card],[valid.fact,stale.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.equal(plan.rows.length,1); assert.equal(plan.rowErrors.length,0);
});
test('90. partial and unresolved duplicate groups do not reject unrelated valid rows', () => {
  const ctx=baseContext(); const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const duplicate=partialSaveFixture(ctx,{documentId:'DUP',factRow:20,card:{'Статус записи':'Удалённая'}});
  const otherFact=partialSaveFixture(ctx,{documentId:'DUP',factRow:21}).fact;
  const plan=ctx.operatorCardBuildSavePlan_([valid.card,duplicate.card],[valid.fact,duplicate.fact,otherFact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.deepEqual(Array.from(plan.rows,p=>p.documentId),['DOC-1']); assert.equal(plan.groupErrors.length,1);
});
test('91. duplicate outside current card has no effect', () => {
  const ctx=baseContext(); const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const d1=partialSaveFixture(ctx,{documentId:'D100',factRow:20}).fact,d2=partialSaveFixture(ctx,{documentId:'D100',factRow:21}).fact;
  const plan=ctx.operatorCardBuildSavePlan_([valid.card],[valid.fact,d1,d2],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.equal(plan.rows.length,1); assert.equal(plan.groupErrors.length,0); assert.equal(plan.duplicateIds.length,0);
});
test('92. resolved duplicate group and unrelated valid row are both planned', () => {
  const ctx=baseContext(); const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const first=partialSaveFixture(ctx,{documentId:'DUP',factRow:20});
  const second=partialSaveFixture(ctx,{documentId:'DUP',factRow:21,card:{'Статус записи':'Удалённая'}});
  const plan=ctx.operatorCardBuildSavePlan_([valid.card,first.card,second.card],[valid.fact,first.fact,second.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.equal(plan.rows.length,2); assert.equal(plan.groupErrors.length,0); assert.equal(plan.duplicateIds.length,1);
});
test('93. rejected rows never create change-history plans', () => {
  const ctx=baseContext(); const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const stale=partialSaveFixture(ctx,{documentId:'DOC-2',factRow:11,updated:'old',fact:{'Дата обновления':'new'},card:{'Комментарий':'edit'}});
  const plan=ctx.operatorCardBuildSavePlan_([valid.card,stale.card],[valid.fact,stale.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.ok(plan.changes.every(change=>change.documentId==='DOC-1'));
});
test('94. physical-row reader preserves offsets around blank rows', () => {
  const ctx=baseContext(); const rows=[['D10'],[''],['D12']];
  const context={headers:['ID документа'],config:{dataStartRow:10},sheet:{getLastRow:()=>12,getRange:()=>({getValues:()=>rows})}};
  const read=ctx.operatorCardSaveRead_(context); assert.deepEqual(Array.from(read,x=>x.sheetRow),[10,12]);
});
test('95. fact writes batch adjacent physical rows by header', () => {
  const ctx=baseContext(); const calls=[];
  const sheet={getRange(row,col,count,width){return{setValues(values){calls.push({row,col,count,width,values})}}}};
  ctx.operatorCardSaveWriteColumnGroups_(sheet,5,[{row:10,value:'a'},{row:11,value:'b'},{row:13,value:'c'}]);
  assert.equal(calls.length,2); assert.equal(calls[0].count,2); assert.equal(calls[1].row,13);
});
test('96. operation ID reservation occurs inside the document lock', () => {
  const text=fs.readFileSync('OperatorCardSave.gs','utf8');
  assert.ok(text.indexOf('withDocumentLock_(function') < text.indexOf('generateOperationId_(started)'));
});
test('97. all destination structures are resolved before fact writes', () => {
  const text=fs.readFileSync('OperatorCardSave.gs','utf8'); const write=text.indexOf('operatorCardSaveWriteFacts_(documents');
  for(const key of ['DOCUMENTS','OPERATOR_CARD','CARD_DICTIONARY','EMPLOYEES','CLIENTS','CHANGE_HISTORY','OPERATION_HISTORY']) assert.ok(text.indexOf("getSystemSheetContext_('"+key+"')")<write);
});
test('98. partial success preserves rejected rows and skips full refresh', () => {
  const save=fs.readFileSync('OperatorCardSave.gs','utf8'),side=fs.readFileSync('OperatorSidebar.html','utf8');
  assert.match(save,/hasSaved && hasRejected[\s\S]*operatorCardSaveSyncSuccessfulCardRows_/);
  assert.match(save,/needsFullRefresh: !hasRejected/); assert.match(save,/if \(result\.needsFullRefresh\)/);
  assert.match(side,/saveProblems\(response\)/); assert.match(side,/изменения не сохранены/);
});
test('99. unresolved complete duplicate remains local and unrelated row survives', () => {
  const ctx=baseContext(); const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const first=partialSaveFixture(ctx,{documentId:'DUP',factRow:20,card:{'Статус записи':'Активная'}});
  const second=partialSaveFixture(ctx,{documentId:'DUP',factRow:21,card:{'Статус записи':'Активная','Комментарий':'touch'}});
  // Touching a duplicate business field rejects only the duplicate group.
  const plan=ctx.operatorCardBuildSavePlan_([valid.card,first.card,second.card],[valid.fact,first.fact,second.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.deepEqual(Array.from(plan.rows,p=>p.documentId),['DOC-1']); assert.equal(plan.groupErrors.length,1);
});
test('100. partial-save status matrix follows the business contract', () => {
  const ctx=baseContext();
  assert.equal(ctx.operatorCardSaveStatus_(2,0,''),'Успешно');
  assert.equal(ctx.operatorCardSaveStatus_(2,1,''),'Успешно с предупреждениями');
  assert.equal(ctx.operatorCardSaveStatus_(0,0,''),'Без изменений');
  assert.equal(ctx.operatorCardSaveStatus_(0,2,''),'Ошибка');
});
test('101. malformed change-history is global and facts remain untouched', () => {
  const ctx=baseContext(); let factWrites=0,insideLock=false,idInsideLock=false;
  const minimal={headers:['ID документа'],headerMap:{'ID документа':1},config:{name:'X',dataStartRow:2},sheet:{getLastRow:()=>1,getRange:()=>({getValues:()=>[],setValues(){}})}};
  const documents={headers:['ID документа'],headerMap:{'ID документа':1},config:{name:'Документы',dataStartRow:2},sheet:{getLastRow:()=>1,getRange:()=>({getValues:()=>[],setValues(){factWrites++}})}};
  ctx.withDocumentLock_=callback=>{insideLock=true;try{return callback()}finally{insideLock=false}};
  ctx.generateOperationId_=()=>{idInsideLock=insideLock;return 'OP-20260804-0001'};
  ctx.getActiveUserEmail_=()=> 'x@x';
  ctx.getSystemSheetContext_=key=>{if(key==='CHANGE_HISTORY')throw new Error('malformed change history');return key==='DOCUMENTS'?documents:minimal};
  assert.throws(()=>ctx.operatorCardSave_({}),/malformed change history/);
  assert.equal(factWrites,0); assert.equal(idInsideLock,true);
});

test('102. an edit to only a read-only field produces a row error naming the field', () => {
  const ctx=baseContext();
  const readOnly=partialSaveFixture(ctx,{documentId:'DOC-RO',factRow:30,card:{'Номер договора':'CHANGED'}});
  const plan=ctx.operatorCardBuildSavePlan_([readOnly.card],[readOnly.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.equal(plan.rows.length,0); assert.equal(plan.rowErrors.length,1);
  assert.match(plan.rowErrors[0].message,/Номер договора/);
});
test('103. a read-only-only error does not reject an unrelated valid row', () => {
  const ctx=baseContext();
  const valid=partialSaveFixture(ctx,{documentId:'DOC-1',factRow:10,card:{'Комментарий':'new'}});
  const readOnly=partialSaveFixture(ctx,{documentId:'DOC-RO',factRow:30,card:{'Номер договора':'CHANGED'}});
  const plan=ctx.operatorCardBuildSavePlan_([valid.card,readOnly.card],[valid.fact,readOnly.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.deepEqual(Array.from(plan.rows,item=>item.documentId),['DOC-1']);
  assert.equal(plan.rowErrors.length,1); assert.match(plan.rowErrors[0].message,/Номер договора/);
});
test('104. a fact-write failure is fatal and never claims guaranteed save success', () => {
  const ctx=baseContext(); let operationRow=null;
  const operationHeaders=['ID операции','Статус операции','Документов с изменениями','Строк факта обновлено','Полей изменено','Ошибок','Текст ошибки / комментарий'];
  const generic={headers:[],headerMap:{'ID документа':1},config:{name:'X',dataStartRow:2},sheet:{getLastRow:()=>1,getRange:()=>({getValues:()=>[],setValues(){}})}};
  const operation={headers:operationHeaders,headerMap:{},config:{name:'История операций',dataStartRow:3},sheet:{getLastRow:()=>2,getRange:()=>({setValues(rows){operationRow=rows[0]}})}};
  ctx.withDocumentLock_=callback=>callback(); ctx.getActiveUserEmail_=()=> 'x@x';
  ctx.getSystemSheetContext_=key=>key==='OPERATION_HISTORY'?operation:generic;
  ctx.operatorCardValidateHeaders_=()=>true; ctx.operatorCardSaveDictionaries_=()=>({});
  ctx.operatorCardSaveRead_=()=>[]; ctx.generateOperationId_=()=> 'OP-20260804-0001';
  ctx.operatorCardBuildSavePlan_=()=>({rows:[{documentId:'DOC-1'}],changes:[{}],rowErrors:[],groupErrors:[],duplicateIds:[]});
  ctx.operatorCardSaveWriteFacts_=()=>{throw new Error('write boom')};
  let thrown; try{ctx.operatorCardSave_({})}catch(error){thrown=error}
  assert.ok(thrown); assert.match(thrown.message,/Часть изменений могла быть записана/);
  assert.doesNotMatch(thrown.message,/Данные документов сохранены/);
  assert.equal(operationRow[operationHeaders.indexOf('Статус операции')],'Ошибка');
  assert.match(operationRow[operationHeaders.indexOf('Текст ошибки / комментарий')],/неопредел|Часть изменений могла быть записана/);
});

test('105. document and object IDs remain hidden technical card fields', () => {
  const ctx = baseContext();
  const map = Array.from(vm.runInContext('SYSTEM_CONFIG.CARD_FIELD_MAP',ctx));
  assert.equal(map.find(x => x.cardHeader === 'ID документа').technical, true);
  assert.equal(map.find(x => x.cardHeader === 'ID объекта').technical, true);
  assert.match(fs.readFileSync('OperatorCard.gs','utf8'), /hideColumns\(cardStartColumn, 2\)/);
});
test('106. object name is read directly from the fact row', () => {
  const ctx=baseContext(); const i={documentId:0,objectId:1,documentTypeId:22,documentStatus:6,holderId:25,foremanId:27,createdAt:20,recordStatus:31,transferredById:26,documentHolder:9,transferredBy:11};
  const row=['D1','OBJ-1','Башня']; while(row.length<32) row.push(''); row[31]='Активная';
  const result=ctx.operatorCardPrepareRows_([row],i,ctx.operatorCardNormalizeFilters_({}),{active:false},4,0,{});
  assert.equal(result.cardRows[0][2],'Башня');
  assert.doesNotMatch(fs.readFileSync('OperatorCard.gs','utf8'),/objectNames|Не удалось определить название объекта/);
});
test('107. approved foreman positions normalize case whitespace and NBSP', () => {
  const ctx=baseContext();
  assert.equal(ctx.objectControlsNormalize_('  ГЛАВНЫЙ\u00a0  ИНЖЕНЕР '),'главный инженер');
  const allowed=Array.from(vm.runInContext('SYSTEM_CONFIG.VALUES.FOREMAN_POSITIONS',ctx));
  assert.ok(allowed.includes('Главный инженер')); assert.ok(!allowed.includes('Бухгалтер'));
});
test('108. duplicate normalized foreman names with different ST-IDs fail safely', () => {
  const ctx=baseContext();
  const rows=[['ST-1',' Иванов\u00a0Иван ','Главный инженер'],['ST-2','иванов  иван','Начальник Участка'],['ST-3','Петров','Бухгалтер']];
  ctx.getSystemSheetContext_=()=>({config:{dataStartRow:5},sheet:{getLastRow:()=>7,getRange:()=>({getValues:()=>rows})}});
  ctx.getSystemColumn_=(key,h)=>({'ID Сотрудника':1,'ФИО сотрудника':2,'Должность':3})[h];
  assert.throws(()=>ctx.objectControlsForemen_(),/одинаковые ФИО.*ST-1, ST-2.*Невозможно однозначно определить ID/);
});
test('109. object statuses are centralized and preserve approved spelling', () => {
  const statuses=Array.from(vm.runInContext('SYSTEM_CONFIG.VALUES.OBJECT_STATUSES',baseContext()));
  assert.deepEqual(statuses,['Действующий','Завершён','Отменён']);
  for(const documentStatus of ['На подготовке','Передан заказчику','Требует исправления','Подписан с обеих сторон']) assert.ok(!statuses.includes(documentStatus));
});
test('110. object synchronization still carries foreman name, ST-ID and status', () => {
  const text=fs.readFileSync('SyncObjectData.gs','utf8');
  assert.match(text,/H\.OBJECT_STATUS/); assert.match(text,/H\.RESPONSIBLE_FOREMAN/); assert.match(text,/H\.RESPONSIBLE_FOREMAN_ID/);
});

test('111. objects start at physical row 3 and modules use configured start rows', () => {
  const ctx=baseContext();
  assert.equal(vm.runInContext('SYSTEM_CONFIG.SHEETS.OBJECTS.dataStartRow',ctx),3);
  for(const file of ['CreateObjectDocuments.gs','SyncObjectData.gs','OperatorCard.gs','ObjectSheetControls.gs']) {
    const text=fs.readFileSync(file,'utf8');
    assert.doesNotMatch(text,/OBJECTS[^\n]{0,120}dataStartRow\s*[:=]\s*4/);
  }
});
test('112. foreman dropdown uses clean names only', () => {
  const ctx=baseContext({SpreadsheetApp:{flush(){},newDataValidation(){const rule={requireValueInList(values){rule.values=values;return rule},setAllowInvalid(){return rule},build(){return rule}};return rule}}});
  const rules=[]; const objects={config:{dataStartRow:3},sheet:{getMaxRows:()=>10,getRange:()=>({setDataValidation(rule){rules.push(rule)}}),hideColumns(){}}};
  ctx.assertSystemSheetsStructure_=()=>{}; ctx.getSystemSheetContext_=key=>key==='OBJECTS'?objects:{};
  ctx.objectControlsForemen_=()=>[{id:'ST-16',name:'Иванов Иван Иванович'}]; ctx.objectControlsEnsureStatuses_=()=>['Действующий','Завершён','Отменён'];
  ctx.getSystemColumn_=(key,h)=>h==='Ответственный прораб'?6:h==='Статус объекта'?7:5;
  ctx.setupObjectSheetControls();
  assert.deepEqual(Array.from(rules[0].values),['Иванов Иван Иванович']);
  assert.ok(!rules[0].values[0].includes('[ST-'));
});
test('113. onEdit preserves clean FIO, writes ST-ID, and clearing FIO clears ID', () => {
  const ctx=baseContext(); let fio=''; let id='OLD'; let note='';
  const idCell={setValue(v){id=v;return this},clearContent(){id='';return this}};
  const sheet={getName:()=>'Объекты',getSheetId:()=>10,getRange:()=>idCell};
  const range={getNumRows:()=>1,getNumColumns:()=>1,getSheet:()=>sheet,getRow:()=>3,getColumn:()=>6,setValue(v){fio=v;return this},setNote(v){note=v;return this},clearContent(){fio='';return this}};
  ctx.getSystemSheetContext_=()=>({config:{dataStartRow:3},sheet});
  ctx.getSystemColumn_=(key,h)=>h==='Ответственный прораб'?6:5;
  ctx.objectControlsForemen_=()=>[{id:'ST-16',name:'Иванов Иван Иванович'}];
  ctx.onEdit({range,value:'Иванов Иван Иванович'});
  assert.equal(fio,'Иванов Иван Иванович'); assert.equal(id,'ST-16'); assert.match(note,/ST-ID/);
  ctx.onEdit({range,value:''}); assert.equal(id,'');
  id='OLD'; fio='bad'; ctx.onEdit({range,value:'Неизвестный'}); assert.equal(id,''); assert.equal(fio,''); assert.match(note,/выпадающего списка/);
});

test('114. final object-name contracts match physical sheets', () => {
  const ctx=baseContext();
  assert.equal(vm.runInContext('H.OBJECT_NAME',ctx),'Название объекта');
  const documents=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.DOCUMENTS.requiredHeaders',ctx));
  const card=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.requiredHeaders',ctx));
  assert.equal(documents.length,34); assert.equal(documents[2],'Название объекта');
  assert.equal(card.length,29); assert.equal(card[2],'Название объекта');
  assert.deepEqual(card.slice(0,27),documents.slice(0,27));
});
test('115. creation writes object name and sync histories name changes', () => {
  const created=creationPrepareFixture(1); const headers=['ID документа','ID объекта','Название объекта','Тип документа','Номер договора','Статус документа','Статус объекта','Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)','Дата создания','Дата обновления','Кто обновил (email)','Ответственный прораб','ID ответственного прораба','Дата изменения статуса документа','Источник создания','Статус записи'];
  assert.equal(created.documentRows[0][headers.indexOf('Название объекта')],'Объект 1');
  const fixture=syncFixture(Object.assign({},baseline,{'Название объекта':'Новое имя'}),[document()]);
  const change=fixture.result.changes.find(item=>item.fieldName==='Название объекта');
  assert.ok(change); assert.equal(change.oldValue,'Объект Альфа'); assert.equal(change.newValue,'Новое имя');
});
test('116. object name is a normal read-only fact field', () => {
  const ctx=baseContext(); const mapping=Array.from(vm.runInContext('SYSTEM_CONFIG.CARD_FIELD_MAP',ctx)).find(item=>item.cardHeader==='Название объекта');
  assert.equal(mapping.factHeader,'Название объекта'); assert.equal(mapping.editable,false); assert.equal(mapping.derived,undefined);
  assert.doesNotMatch(fs.readFileSync('OperatorCardSave.gs','utf8'),/mapping\.derived/);
  const changed=partialSaveFixture(ctx,{card:{'Название объекта':'Подмена'}});
  const rejected=ctx.operatorCardBuildSavePlan_([changed.card],[changed.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0001');
  assert.equal(rejected.rows.length,0); assert.match(rejected.rowErrors[0].message,/Название объекта/);
  const editable=partialSaveFixture(ctx,{card:{'Комментарий':'new'}});
  const accepted=ctx.operatorCardBuildSavePlan_([editable.card],[editable.fact],partialDictionaries(),new Date(),'x@x','OP-20260804-0002');
  assert.equal(accepted.rows.length,1);
});

test('117. creation accepts every object status and rejects an empty status', () => {
  const ctx=baseContext();
  const headers=['ID объекта','Название объекта','Адрес','Номер договора','ID ответственного прораба','Ответственный прораб','Статус объекта','Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)'];
  const start=vm.runInContext("new Date('2026-08-01T00:00:00Z')",ctx);
  const end=vm.runInContext("new Date('2026-08-31T00:00:00Z')",ctx);
  const statuses=['Действующий','Завершён','Отменён','','На подготовке'];
  const rows=statuses.map((status,index)=>['OBJ-'+index,'Объект '+index,'','DOG-'+index,'ST-1','Иванов',status,start,end,'']);
  const context={headers,headerMap:Object.fromEntries(headers.map((h,i)=>[h,i+1])),config:{dataStartRow:3,requiredHeaders:headers},sheet:{getLastRow:()=>7,getRange:()=>({getValues:()=>rows})}};
  ctx.getSystemSheetContext_=()=>context;
  const result=ctx.readAndValidateCreationObjects_();
  assert.deepEqual(Array.from(result.validObjects,item=>item.objectStatus),['Действующий','Завершён','Отменён']);
  assert.equal(result.skippedObjects.length,2); assert.match(result.skippedObjects[0].reasons.join(' '),/не заполнено поле «Статус объекта»/);
  assert.match(result.skippedObjects[1].reasons.join(' '),/недопустимое значение «На подготовке»/);
});

function readOnlyGuardFixture(editColumn, editWidth, mutate) {
  const ctx=baseContext();
  const cardHeaders=operatorWorkflowCardHeaders.slice();
  const factHeaders=operatorWorkflowDocumentHeaders.slice();
  const cardValues={'ID документа':'DOC-1','ID объекта':'OBJ-1','Название объекта':'Объект А','Тип документа':'Акт','Номер договора':'DOG-1','Статус документа':'Новый','Оригинал / ЭДО':'Оригинал','Комментарий':'old','Где документ':'Офис','Оплачен':'Нет','ГУ (Да/Нет)':'Нет','Статус объекта':'Действующий','Ответственный прораб':'Иванов','Дата создания':'created','Дата обновления':'v1','ID типа документа':'TYPE-1','Кто обновил (email)':'old@example.com','Статус записи':'Активная','Номер строки в таблице фактов':4};
  const card=cardHeaders.map(header=>Object.prototype.hasOwnProperty.call(cardValues,header)?cardValues[header]:'');
  const factValues=Object.assign({},cardValues,{'ID ответственного прораба':'ST-1'});
  const fact=factHeaders.map(header=>Object.prototype.hasOwnProperty.call(factValues,header)?factValues[header]:'');
  if (mutate) mutate(card);
  const notes={}; let cardReads=0,factReads=0;
  const cardSheet={getName:()=> 'Карточка операциониста',getSheetId:()=>20,getRange(row,column,rowCount,columnCount){
    if (rowCount!==undefined) { cardReads++; return {getValues:()=>[card.slice(column-1,column-1+columnCount)]}; }
    return {setValue(value){card[column-1]=value;return this},setNote(value){notes[column]=value;return this}};
  }};
  const factSheet={getLastRow:()=>4,getRange(row,column,rowCount,columnCount){factReads++;return {getValues:()=>[fact.slice()]}}};
  const cardContext={config:{dataStartRow:6,requiredHeaders:cardHeaders},headers:cardHeaders,headerMap:Object.fromEntries(cardHeaders.map((h,i)=>[h,i+1])),sheet:cardSheet};
  const factContext={config:{dataStartRow:4},headers:factHeaders,headerMap:Object.fromEntries(factHeaders.map((h,i)=>[h,i+1])),sheet:factSheet};
  ctx.getSystemSheetContext_=key=>key==='OPERATOR_CARD'?cardContext:factContext;
  const range={getSheet:()=>cardSheet,getRow:()=>6,getNumRows:()=>1,getColumn:()=>editColumn,getNumColumns:()=>editWidth};
  return {ctx,card,fact,notes,range,reads:()=>({cardReads,factReads})};
}
test('118. project has exactly one global onEdit router', () => {
  const count=files.reduce((sum,file)=>sum+(fs.readFileSync(file,'utf8').match(/function\s+onEdit\s*\(/g)||[]).length,0);
  assert.equal(count,1); const text=fs.readFileSync('ObjectSheetControls.gs','utf8');
  assert.match(text,/objectControlsHandleEdit_\(event\)/); assert.match(text,/operatorCardHandleReadOnlyEdit_\(event\)/);
});
test('119. card guard restores read-only object fields from physical fact row', () => {
  for(const [header,bad] of [['Название объекта','ТЕСТ'],['Номер договора','BAD'],['Ответственный прораб','BAD NAME']]) {
    const column=operatorWorkflowCardHeaders.indexOf(header)+1;
    const fixture=readOnlyGuardFixture(column,1,card=>{card[column-1]=bad}); const expected=fixture.fact[column-1];
    fixture.ctx.operatorCardHandleReadOnlyEdit_({range:fixture.range});
    assert.equal(fixture.card[column-1],expected,header); assert.match(fixture.notes[column],/недоступно для ручного изменения/);
    assert.deepEqual(fixture.reads(),{cardReads:1,factReads:1});
  }
});
test('120. card guard ignores editable comment and document status', () => {
  for(const header of ['Комментарий','Статус документа']) {
    const column=operatorWorkflowCardHeaders.indexOf(header)+1;
    const fixture=readOnlyGuardFixture(column,1,card=>{card[column-1]='USER VALUE'});
    fixture.ctx.operatorCardHandleReadOnlyEdit_({range:fixture.range});
    assert.equal(fixture.card[column-1],'USER VALUE'); assert.deepEqual(fixture.reads(),{cardReads:0,factReads:0});
  }
});
test('121. mixed paste restores only read-only cells and preserves editable cells', () => {
  const first=3,width=8;
  const fixture=readOnlyGuardFixture(first,width,card=>{
    for(let i=first-1;i<first-1+width;i++) card[i]='PASTE-'+i;
    card[0]='DOC-1'; card[1]='OBJ-1'; card[operatorWorkflowCardHeaders.indexOf('ID типа документа')]='TYPE-1';
  });
  fixture.ctx.operatorCardHandleReadOnlyEdit_({range:fixture.range});
  for(const header of ['Название объекта','Тип документа','Номер договора']) {
    const index=operatorWorkflowCardHeaders.indexOf(header); assert.equal(fixture.card[index],fixture.fact[index],header);
  }
  for(const header of ['Дата документа','Статус документа','Оригинал / ЭДО','Комментарий']) {
    const index=operatorWorkflowCardHeaders.indexOf(header); assert.equal(fixture.card[index],'PASTE-'+index,header);
  }
});
test('122. identity mismatch never restores from another physical fact row', () => {
  const column=operatorWorkflowCardHeaders.indexOf('Название объекта')+1;
  const fixture=readOnlyGuardFixture(column,1,card=>{card[1]='OTHER-OBJECT';card[column-1]='ТЕСТ'});
  fixture.ctx.operatorCardHandleReadOnlyEdit_({range:fixture.range});
  assert.equal(fixture.card[column-1],'ТЕСТ'); assert.match(fixture.notes[column],/identity/);
});

function webAuthFixture(users) {
  const values = {};
  const properties = { getProperty(key) { return values[key] || null; }, setProperty(key, value) { values[key] = value; }, deleteProperty(key) { delete values[key]; } };
  const state = { users: users };
  const ctx = baseContext({ scriptProperties: properties });
  ctx.Utilities.getUuid = () => 'uuid-' + Object.keys(values).length;
  ctx.webAppReadUsers_ = () => state.users;
  return { ctx, values, state };
}
test('123. web login accepts enabled user and never returns password', () => {
  const fixture = webAuthFixture([{ login:'irina',password:'Secret',fullName:'Ирина',contact:'irina@work',access:'Да' }]);
  const result = fixture.ctx.webAppLogin(' irina ', 'Secret');
  assert.equal(result.user.actor, 'irina@work'); assert.equal(result.user.password, undefined);
  assert.ok(result.token); assert.ok(!JSON.stringify(result).includes('Secret'));
});
test('124. web login rejects wrong password, unknown login and denied access', () => {
  const enabled = webAuthFixture([{ login:'irina',password:'Secret',fullName:'Ирина',contact:'',access:'Да' }]).ctx;
  assert.throws(() => enabled.webAppLogin('irina','secret'), /Неверный/);
  assert.throws(() => enabled.webAppLogin('unknown','Secret'), /Неверный/);
  const denied = webAuthFixture([{ login:'irina',password:'Secret',fullName:'Ирина',contact:'',access:'Нет' }]).ctx;
  assert.throws(() => denied.webAppLogin('irina','Secret'), /запрещён/);
});
test('125. duplicate web login is blocked and contact falls back to login', () => {
  const users = [{ login:'irina',password:'a',fullName:'',contact:'',access:'Да' },{ login:'irina',password:'a',fullName:'',contact:'',access:'Да' }];
  assert.throws(() => webAuthFixture(users).ctx.webAppLogin('irina','a'), /повторяющийся логин/);
  assert.equal(webAuthFixture([]).ctx.webAppPublicUser_({login:'user',fullName:'',contact:''}).actor, 'user');
});
test('126. web sessions persist without expiry and logout removes them', () => {
  const fixture=webAuthFixture([{login:'u',password:'p',fullName:'U',contact:'',access:'Да'}]);
  const result=fixture.ctx.webAppLogin('u','p'); assert.equal(fixture.ctx.webAppRequireSession_(result.token).login,'u');
  assert.deepEqual(JSON.parse(Object.values(fixture.values)[0]),{login:'u'});
  assert.ok(!JSON.stringify(fixture.values).includes('"password"'));
  assert.equal(fixture.ctx.webAppRequireSession_(result.token).login,'u');
  fixture.ctx.webAppLogout(result.token); assert.throws(()=>fixture.ctx.webAppRequireSession_(result.token),/Сессия недействительна/);
});
test('127. web configuration is centralized and excludes special object', () => {
  const ctx=baseContext(); const config=vm.runInContext('SYSTEM_CONFIG',ctx); const headers=vm.runInContext('SYSTEM_HEADERS',ctx);
  assert.deepEqual([config.SHEETS.WEB_USERS.name,config.SHEETS.WEB_USERS.headerRow,config.SHEETS.WEB_USERS.dataStartRow],['Справочник Web-пользователей',4,5]);
  assert.deepEqual([headers.WEB_LOGIN,headers.WEB_PASSWORD,headers.WEB_FULL_NAME,headers.WEB_CONTACT,headers.WEB_ACCESS],['Логин','Пароль','ФИО','Контакт','Доступ']);
  const text=fs.readFileSync('DocumentWebAppServer.gs','utf8'); assert.match(text,/object\.id !== SYSTEM_CONFIG\.VALUES\.ALL_OBJECTS_LABEL/);
});
test('128. repeatability rules preserve base type ID and use max numbering', () => {
  const ctx=baseContext(); assert.equal(ctx.webAppValidRepeatability_('Один'),true); assert.equal(ctx.webAppValidRepeatability_('Много'),true); assert.equal(ctx.webAppValidRepeatability_(''),false);
  const text=fs.readFileSync('DocumentWebAppServer.gs','utf8'); assert.doesNotMatch(text,/existing\.length \+ 1/); assert.match(text,/values\[H\.DOCUMENT_TYPE_ID\] = type\.id/);
  assert.match(text,/webAppNextDocumentNumber_/);
  assert.doesNotMatch(text,/ACTIVE_RECORD_STATUS[\s\S]{0,200}webAppReadMatchingDocuments_/);
});
test('129. web document preparation fills system/object fields and leaves business fields empty', () => {
  const ctx=baseContext();
  ctx.getSystemSpreadsheet_=()=>({getSpreadsheetTimeZone:()=> 'UTC'});
  const headers=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.DOCUMENTS.requiredHeaders',ctx));
  const context={headers,headerMap:Object.fromEntries(headers.map((h,i)=>[h,i+1]))};
  const now=vm.runInContext("new Date('2026-08-17T10:00:00Z')",ctx);
  const facts={documentIds:{'DOC-15-0001':[4],'DOC-15-0003':[5]}};
  ctx.generateChangeId_=()=> 'CHG-1';
  const result=ctx.webAppPrepareDocument_({id:'15',name:'Объект',contractNumber:'D-1',objectStatus:'Действующий',workStartDate:now,workEndPlan:now,workEndFact:'',responsibleForeman:'Иванов',responsibleForemanId:'ST-1'},{id:'TYPE-1'},'Акт',3,'Подписан',facts,context,6,now,'actor','OP-1');
  const value=h=>result.row[headers.indexOf(h)];
  assert.equal(result.documentId,'DOC-15-0004'); assert.equal(value('ID типа документа'),'TYPE-1'); assert.equal(value('Тип документа'),'Акт');
  assert.equal(value('Номер документа'),3); assert.equal(result.row.length,34);
  assert.equal(value('Статус записи'),'Активная'); assert.equal(value('Источник создания'),'Web-приложение'); assert.equal(value('Кто обновил (email)'),'actor');
  for(const h of ['Дата создания','Дата обновления','Дата изменения статуса документа']) assert.equal(Object.prototype.toString.call(value(h)),'[object Date]',h);
  for(const h of ['Дата документа','Оригинал / ЭДО','Комментарий','У кого документ','Где документ','Кто передал','Оплачен','Сумма документа','ГУ (Да/Нет)','Условия ГУ','ID сотрудника — у кого документ','ID сотрудника — кто передал','Отчётный период']) assert.equal(value(h),'',h);
  assert.equal(result.change.factRow,6); assert.equal(result.change.operationId,'OP-1');
});
test('130. create endpoint revalidates everything under the shared lock', () => {
  const text=fs.readFileSync('DocumentWebAppServer.gs','utf8');
  const start=text.indexOf('function webAppCreateDocument'); const lock=text.indexOf('withDocumentLock_',start); const object=text.indexOf('webAppFindObject_',lock); const type=text.indexOf('webAppFindDocumentType_',lock); const status=text.indexOf('webAppReadStatuses_',lock); const facts=text.indexOf('webAppReadMatchingDocuments_',lock);
  assert.ok(lock>start && object>lock && type>object && status>type && facts>status); assert.match(text,/OPERATION_STATUS_NO_CHANGES/); assert.match(text,/factWritten/);
});
test('131. web writes are batched and operation/change contracts are explicit', () => {
  const text=fs.readFileSync('DocumentWebAppServer.gs','utf8'); assert.doesNotMatch(text,/appendRow\s*\(/); assert.match(text,/\.setValues\(\[row\]\)/);
  assert.match(text,/WEB_APP_OPERATION_TYPE/); assert.match(text,/CHANGE_ACTION_CREATE/); assert.match(text,/FIELDS_CHANGED\] = operation\.changed/);
  assert.match(text,/generateOperationId_/); assert.match(text,/generateChangeId_\(operationId, 1\)/);
});
test('132. web HTML uses safe DOM output, sessionStorage and exact payload', () => {
  const html=fs.readFileSync('DocumentWebApp.html','utf8'); assert.doesNotMatch(html,/\.innerHTML\s*=/); assert.match(html,/textContent/); assert.match(html,/sessionStorage/);
  assert.match(html,/\{ objectId: el\('object'\)\.value, documentTypeId: el\('documentType'\)\.value, documentStatus: el\('status'\)\.value \}/);
  assert.doesNotMatch(fs.readFileSync('WebAppAuth.gs','utf8'),/return\s+\{[^}]*password/);
});
test('133. protected request refreshes current contact and falls back to login', () => {
  const fixture=webAuthFixture([{login:'u',password:'p',fullName:'Old',contact:'old@work',access:'Да'}]);
  const token=fixture.ctx.webAppLogin('u','p').token;
  fixture.state.users=[{login:'u',password:'changed',fullName:'New',contact:'new@work',access:'Да'}];
  assert.equal(fixture.ctx.webAppRequireSession_(token).actor,'new@work');
  fixture.state.users=[{login:'u',password:'changed',fullName:'New',contact:'',access:'Да'}];
  assert.equal(fixture.ctx.webAppRequireSession_(token).actor,'u');
});
test('134. revoked access deletes the server session permanently', () => {
  const fixture=webAuthFixture([{login:'u',password:'p',fullName:'U',contact:'',access:'Да'}]);
  const token=fixture.ctx.webAppLogin('u','p').token;
  fixture.state.users=[{login:'u',password:'p',fullName:'U',contact:'',access:'Нет'}];
  assert.throws(()=>fixture.ctx.webAppRequireSession_(token),/Доступ.*прекращён/); assert.equal(Object.keys(fixture.values).length,0);
  fixture.state.users=[{login:'u',password:'p',fullName:'U',contact:'',access:'Да'}];
  assert.throws(()=>fixture.ctx.webAppRequireSession_(token),/Сессия недействительна/);
});
test('135. missing user and duplicate login both revoke and delete sessions', () => {
  for(const replacement of [[],[{login:'u',password:'p',fullName:'1',contact:'',access:'Да'},{login:'u',password:'p',fullName:'2',contact:'',access:'Да'}]]) {
    const fixture=webAuthFixture([{login:'u',password:'p',fullName:'U',contact:'',access:'Да'}]); const token=fixture.ctx.webAppLogin('u','p').token;
    fixture.state.users=replacement; assert.throws(()=>fixture.ctx.webAppRequireSession_(token),/Доступ.*прекращён/); assert.equal(Object.keys(fixture.values).length,0);
  }
});
test('136. session storage uses Script Properties without expiry or CacheService', () => {
  const text=fs.readFileSync('WebAppAuth.gs','utf8'); assert.match(text,/PropertiesService\.getScriptProperties\(\)/);
  assert.match(text,/\.setProperty\(/); assert.match(text,/\.deleteProperty\(/); assert.doesNotMatch(text,/CacheService|SESSION_SECONDS|setProperty\([^)]*,[^)]*,/);
});
test('137. document creation rechecks access inside lock and revoked access writes nothing', () => {
  const ctx=baseContext(); let checks=0, locked=false, touched=false;
  ctx.webAppRequireSession_=()=>{ checks++; if(checks===1)return {actor:'old'}; assert.equal(locked,true); throw new Error('Доступ к приложению прекращён. Войдите снова.'); };
  ctx.withDocumentLock_=callback=>{locked=true;return callback();};
  ctx.assertSystemSheetsStructure_=()=>{touched=true;}; ctx.getSystemSheetContext_=()=>{touched=true;};
  assert.throws(()=>ctx.webAppCreateDocument('token',{}),/Доступ.*прекращён/); assert.equal(checks,2); assert.equal(touched,false);
  const text=fs.readFileSync('DocumentWebAppServer.gs','utf8'); const start=text.indexOf('function webAppCreateDocument'); const lock=text.indexOf('withDocumentLock_',start); const second=text.indexOf('webAppRequireSession_',lock); assert.ok(second>lock);
});
test('138. client removes revoked tokens, returns to login and never retries create', () => {
  const html=fs.readFileSync('DocumentWebApp.html','utf8'); assert.match(html,/function handleSessionError[\s\S]*sessionStorage\.removeItem\('documentWebToken'\)[\s\S]*showLogin/);
  assert.match(html,/Доступ к приложению прекращён/); assert.match(html,/Сессия недействительна/);
  const createHandler=html.slice(html.indexOf("el('documentForm').addEventListener"),html.indexOf("el('logout').addEventListener"));
  assert.equal((createHandler.match(/webAppCreateDocument/g)||[]).length,1); assert.doesNotMatch(createHandler,/setTimeout|retry|повтор/i);
});
test('139. system spreadsheet prefers the active spreadsheet', () => {
  const active={name:'active'}; let propertiesRead=0,opened=0;
  const ctx=baseContext({
    SpreadsheetApp:{getActiveSpreadsheet:()=>active,openById(){opened++;}},
    scriptProperties:{getProperty(){propertiesRead++;return 'saved-id';},setProperty(){},deleteProperty(){}}
  });
  assert.equal(ctx.getSystemSpreadsheet_(),active); assert.equal(propertiesRead,0); assert.equal(opened,0);
});
test('140. system spreadsheet opens saved ID when active spreadsheet is unavailable', () => {
  const opened={name:'opened'}; let requested='';
  const ctx=baseContext({
    SpreadsheetApp:{getActiveSpreadsheet:()=>null,openById(id){requested=id;return opened;}},
    scriptProperties:{getProperty(key){assert.equal(key,'SYSTEM_SPREADSHEET_ID');return ' sheet-id ';},setProperty(){},deleteProperty(){}}
  });
  assert.equal(ctx.getSystemSpreadsheet_(),opened); assert.equal(requested,'sheet-id');
});
test('141. missing spreadsheet connection reports the setup instruction', () => {
  const ctx=baseContext({
    SpreadsheetApp:{getActiveSpreadsheet:()=>null,openById(){throw new Error('must not open');}},
    scriptProperties:{getProperty(){return '';},setProperty(){},deleteProperty(){}}
  });
  assert.throws(()=>ctx.getSystemSpreadsheet_(),/setupSystemSpreadsheetConnection\(\)/);
});
test('142. spreadsheet setup saves active ID and returns confirmation', () => {
  let savedKey='',savedValue=''; const active={getId:()=> 'main-sheet-id',getName:()=> 'Осушение'};
  const ctx=baseContext({
    SpreadsheetApp:{getActiveSpreadsheet:()=>active},
    scriptProperties:{getProperty(){return null;},setProperty(key,value){savedKey=key;savedValue=value;},deleteProperty(){}}
  });
  const result=ctx.setupSystemSpreadsheetConnection(); assert.equal(savedKey,'SYSTEM_SPREADSHEET_ID'); assert.equal(savedValue,'main-sheet-id');
  assert.equal(result.ok,true); assert.equal(result.spreadsheetId,'main-sheet-id'); assert.equal(result.spreadsheetName,'Осушение'); assert.match(result.message,/успешно/);
});
test('143. spreadsheet setup rejects missing active spreadsheet without writes', () => {
  let writes=0; const ctx=baseContext({
    SpreadsheetApp:{getActiveSpreadsheet:()=>null},
    scriptProperties:{getProperty(){return null;},setProperty(){writes++;},deleteProperty(){}}
  });
  assert.throws(()=>ctx.setupSystemSpreadsheetConnection(),/активную Google-таблицу/); assert.equal(writes,0);
});
test('144. renamed server loads and doGet still opens DocumentWebApp HTML', () => {
  assert.ok(files.includes('DocumentWebAppServer.gs')); assert.ok(!files.includes('DocumentWebApp.gs'));
  let requested=''; const output={setTitle(){return this;},addMetaTag(){return this;}};
  const ctx=baseContext({HtmlService:{createHtmlOutputFromFile(name){requested=name;return output;}}});
  assert.equal(ctx.doGet(),output); assert.equal(requested,'DocumentWebApp');
});


test('145. v2 contracts separate canonical document type from document number', () => {
  const ctx=baseContext();
  assert.equal(vm.runInContext('H.DOCUMENT_NUMBER',ctx),'Номер документа');
  assert.equal(vm.runInContext('H.CUSTOMER_SIGNING_RESPONSIBLE',ctx),'Кто ответственный за подписание (заказчик)');
  const documents=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.DOCUMENTS.requiredHeaders',ctx));
  assert.deepEqual(documents.slice(3,7),['Тип документа','Номер договора','Номер документа','Дата документа']);
  assert.deepEqual(documents.slice(18,21),['Статус объекта','Ответственный прораб','Дата начала работ']);
  const map=Array.from(vm.runInContext('SYSTEM_CONFIG.CARD_FIELD_MAP',ctx));
  assert.equal(map.find(item=>item.cardHeader==='Тип документа').editable,false);
  assert.equal(map.find(item=>item.cardHeader==='Номер документа').editable,true);
  assert.equal(map.find(item=>item.cardHeader==='Кто ответственный за подписание (заказчик)').editable,true);
});
test('146. v2 migration extracts only a suffix number and preserves an existing number', () => {
  const ctx=baseContext();
  assert.equal(ctx.documentArchitectureExtractLegacyNumber_('Дополнительное соглашение №3','Дополнительное соглашение',''),'3');
  assert.equal(ctx.documentArchitectureExtractLegacyNumber_('КС-2 — акт о приёмке выполненных работ №2','КС-2 — акт о приёмке выполненных работ',''),'2');
  assert.equal(ctx.documentArchitectureExtractLegacyNumber_('Чужой тип №7','Дополнительное соглашение',''),'');
  assert.equal(ctx.documentArchitectureExtractLegacyNumber_('Дополнительное соглашение №3','Дополнительное соглашение','ручной-5'),'');
});
test('147. v2 migration is data-only and uses centralized contracts', () => {
  const text=fs.readFileSync('DocumentArchitectureV2Migration.gs','utf8');
  assert.doesNotMatch(text,/insertColumn|deleteColumn|moveColumn|setFrozen|setName/);
  assert.match(text,/assertSystemSheetsStructure_/);
  assert.match(text,/H\.DOCUMENT_NUMBER/);
  assert.match(text,/DOCUMENT_V2_MIGRATION_SOURCE/);
  assert.doesNotMatch(text,/\.setValue\s*\(/);
  assert.match(text,/\.setValues\(group\.values\)/);
});

test('148. Web App numbering uses max plus one and never fills gaps', () => {
  const ctx=baseContext();
  const documents=[1,2,4].map(number=>({documentType:'КС-2',storedDocumentType:'КС-2',documentNumber:number}));
  assert.equal(ctx.webAppNextDocumentNumber_(documents,'КС-2'),5);
  assert.equal(ctx.webAppNextDocumentNumber_([],'КС-2'),1);
  assert.equal(ctx.webAppDocumentNumberOccupied_(documents,'КС-2',4),true);
  assert.equal(ctx.webAppDocumentNumberOccupied_(documents,'КС-2',5),false);
  assert.equal(ctx.webAppNextDocumentNumber_([
    {documentType:'КС-2',storedDocumentType:'КС-2',documentNumber:'КС-2 №7'}
  ],'КС-2'),8);
});

test('149. Web App numbering supports legacy suffix with explicit-number priority', () => {
  const ctx=baseContext();
  assert.equal(ctx.webAppNextDocumentNumber_([
    {documentType:'КС-2',storedDocumentType:'КС-2 №3',documentNumber:''}
  ],'КС-2'),4);
  assert.equal(ctx.webAppNextDocumentNumber_([
    {documentType:'КС-2',storedDocumentType:'КС-2 №99',documentNumber:'4'}
  ],'КС-2'),5);
  assert.equal(ctx.webAppNextDocumentNumber_([
    {documentType:'КС-2',storedDocumentType:'КС-2 2026 акт',documentNumber:''},
    {documentType:'КС-2',storedDocumentType:'КС-2',documentNumber:'0'},
    {documentType:'КС-2',storedDocumentType:'КС-2',documentNumber:'2.5'}
  ],'КС-2'),1);
});

test('150. repeatability One stores an empty number and Many stores it separately', () => {
  const ctx=baseContext();
  const headers=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.DOCUMENTS.requiredHeaders',ctx));
  const context={headers,headerMap:Object.fromEntries(headers.map((header,index)=>[header,index+1]))};
  const object={id:'7',name:'Объект',contractNumber:'',objectStatus:'Действующий',workStartDate:'',workEndPlan:'',workEndFact:'',responsibleForeman:'',responsibleForemanId:''};
  const facts={documentIds:{}}; const now=new Date(); ctx.generateChangeId_=()=> 'CHG-1';
  const one=ctx.webAppPrepareDocument_(object,{id:'ONE'},'Договор','', 'Новый',facts,context,4,now,'actor','OP-1');
  const many=ctx.webAppPrepareDocument_(object,{id:'MANY'},'КС-2',3, 'Новый',facts,context,5,now,'actor','OP-2');
  const at=(row,header)=>row[headers.indexOf(header)];
  assert.equal(at(one.row,'Номер документа'),'');
  assert.equal(at(many.row,'Тип документа'),'КС-2');
  assert.equal(at(many.row,'Номер документа'),3);
  assert.doesNotMatch(at(many.row,'Тип документа'),/№/);
});

test('151. existing Web App documents expose number and display name separately', () => {
  const ctx=baseContext();
  const headers=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.DOCUMENTS.requiredHeaders',ctx));
  const makeRow=values=>headers.map(header=>Object.prototype.hasOwnProperty.call(values,header)?values[header]:'');
  const rows=[
    makeRow({'ID объекта':'OBJ','ID типа документа':'TYPE','Тип документа':'КС-2','Номер документа':4,'ID документа':'D1'}),
    makeRow({'ID объекта':'OBJ','ID типа документа':'TYPE','Тип документа':'КС-2 №5','Номер документа':'','ID документа':'D2'})
  ];
  ctx.getSystemSheetContext_=()=>({headers,headerMap:Object.fromEntries(headers.map((header,index)=>[header,index+1])),config:{dataStartRow:4},sheet:{getLastRow:()=>5,getRange:()=>({getValues:()=>rows})}});
  const result=ctx.webAppReadMatchingDocuments_('OBJ','TYPE','КС-2');
  assert.equal(result[0].documentNumber,'4'); assert.equal(result[0].displayName,'КС-2 №4');
  assert.equal(result[1].documentNumber,''); assert.equal(result[1].displayName,'КС-2 №5');
});

test('152. preview and locked creation share the numbering algorithm', () => {
  const text=fs.readFileSync('DocumentWebAppServer.gs','utf8');
  const preview=text.slice(text.indexOf('function webAppGetExistingDocuments'),text.indexOf('function webAppCreateDocument'));
  const create=text.slice(text.indexOf('function webAppCreateDocument'),text.indexOf('function webAppReadDocumentTypes_'));
  assert.match(preview,/webAppNextDocumentNumber_/);
  assert.match(create,/withDocumentLock_[\s\S]*webAppReadMatchingDocuments_[\s\S]*webAppNextDocumentNumber_[\s\S]*webAppDocumentNumberOccupied_/);
  assert.doesNotMatch(text,/values\[H\.DOCUMENT_TYPE\]\s*=\s*[^;]*\+[^;]*№/);
});

test('153. creation snapshot keeps canonical type and separate document number', () => {
  const ctx=baseContext();
  const snapshot=ctx.buildCreationInitialSnapshot_({
    'Тип документа':'КС-2','Номер документа':6,'ID типа документа':'TYPE'
  });
  assert.match(snapshot,/Тип документа: КС-2/);
  assert.match(snapshot,/Номер документа: 6/);
  assert.doesNotMatch(snapshot,/КС-2 №6/);
});

test('154. standalone Web App loads and numbers without migration module', () => {
  const standaloneFiles=['SystemCore.gs','DocumentArchitectureCore.gs',
    'CreateObjectDocuments.gs','WebAppAuth.gs','DocumentWebAppServer.gs'];
  assert.ok(!standaloneFiles.includes('DocumentArchitectureV2Migration.gs'));
  const standaloneSource=standaloneFiles.map(file=>fs.readFileSync(file,'utf8')).join('\n');
  const sandbox={console,Number,Date,Math}; vm.createContext(sandbox);
  vm.runInContext(standaloneSource,sandbox);
  assert.equal(sandbox.documentArchitectureExtractLegacyNumber_('КС-2 №4','КС-2',''),'4');
  assert.equal(sandbox.webAppNextDocumentNumber_([
    {documentType:'КС-2',storedDocumentType:'КС-2 №4',documentNumber:''}
  ],'КС-2'),5);
});

test('155. shared legacy parser accepts only a positive trailing integer', () => {
  const ctx=baseContext();
  assert.equal(ctx.documentArchitectureExtractLegacyNumber_('КС-2 №003','КС-2',''),'3');
  for(const value of ['КС-2 №0','КС-2 №-1','КС-2 №2.5','КС-2 №3 копия','КС-2 2026']) {
    assert.equal(ctx.documentArchitectureExtractLegacyNumber_(value,'КС-2',''),'',value);
  }
  assert.equal(ctx.documentArchitectureExtractLegacyNumber_('КС-2 №99','КС-2','7'),'');
});

test('156. creation and sync preserve v2 document-owned fields', () => {
  const create=fs.readFileSync('CreateObjectDocuments.gs','utf8');
  assert.match(create,/findNextCreationTypeDocumentNumber_/);
  assert.match(create,/values\[H\.CUSTOMER_SIGNING_RESPONSIBLE\] = ''/);
  const syncFields=Array.from(vm.runInContext('OBJECT_SYNC_FIELDS_',baseContext()));
  assert.ok(!syncFields.includes('Тип документа'));
  assert.ok(!syncFields.includes('Номер документа'));
  assert.ok(!syncFields.includes('Кто ответственный за подписание (заказчик)'));
  assert.equal(baseContext().assertObjectSyncFieldContract_(),undefined);
  assert.match(fs.readFileSync('SyncObjectData.gs','utf8'),/assertObjectSyncFieldContract_\(\)/);
  assert.equal(baseContext().findNextCreationTypeDocumentNumber_([
    {documentType:'Акт',documentNumber:'Акт №4'}
  ],'Акт'),5);
});

test('157. migration partially skips problems and returns the exact report contract', () => {
  const ctx=baseContext();
  const documentHeaders=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.DOCUMENTS.requiredHeaders',ctx));
  const typeHeaders=['ID типа документа','Тип документа'];
  const historyHeaders=['ID изменения','ID операции','Дата и время изменения','Кто изменил (email)','Тип действия','ID документа','ID объекта','Номер строки в таблице фактов','Название поля','Старое значение','Новое значение','Источник изменения'];
  const operationHeaders=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.OPERATION_HISTORY.requiredHeaders',ctx));
  const row=values=>documentHeaders.map(header=>Object.prototype.hasOwnProperty.call(values,header)?values[header]:'');
  const documentRows=[
    row({'ID документа':'D1','ID объекта':'O1','ID типа документа':'T1','Тип документа':'Дополнительное соглашение №3'}),
    row({'ID документа':'D2','ID объекта':'O1','Тип документа':'Дополнительное соглашение №4'}),
    row({'ID документа':'D3','ID объекта':'O1','ID типа документа':'UNKNOWN','Тип документа':'X'}),
    row({'ID документа':'D4','ID объекта':'O1','ID типа документа':'DUP','Тип документа':'Дубль'}),
    row({'ID документа':'D5','ID объекта':'O1','ID типа документа':'T1','Тип документа':'Дополнительное соглашение №X'}),
    row({'ID документа':'D6','ID объекта':'O1','ID типа документа':'T1','Тип документа':'Дополнительное соглашение','Номер документа':'already'})
  ];
  const writes=[];
  function context(headers,dataStartRow,rows) {
    return {headers,headerMap:Object.fromEntries(headers.map((header,index)=>[header,index+1])),config:{dataStartRow},sheet:{
      getLastRow:()=>dataStartRow+rows.length-1,
      getRange(start,column,count,width){return {getValues:()=>rows,setValues(values){writes.push({start,column,count,width,values})}}}
    }};
  }
  const contexts={
    DOCUMENT_TYPES:context(typeHeaders,5,[['T1','Дополнительное соглашение'],['DUP','Дубль'],['DUP','Дубль']]),
    DOCUMENTS:context(documentHeaders,4,documentRows),
    CHANGE_HISTORY:context(historyHeaders,3,[]),
    OPERATION_HISTORY:context(operationHeaders,3,[])
  };
  ctx.withDocumentLock_=callback=>callback(); ctx.assertSystemSheetsStructure_=()=>{};
  ctx.generateOperationId_=()=> 'OP-1'; ctx.generateChangeId_=(id,n)=>'CHG-'+n;
  ctx.getActiveUserEmail_=()=> 'actor'; ctx.getSystemSheetContext_=key=>contexts[key];
  const report=ctx.migrateDocumentNumberStructure();
  assert.deepEqual(Object.keys(report),['checkedRows','changedRows','unchangedRows','skippedRows','warningsCount','problems','operationId']);
  assert.deepEqual({...report,problems:undefined},{checkedRows:6,changedRows:1,unchangedRows:1,skippedRows:4,warningsCount:4,problems:undefined,operationId:'OP-1'});
  assert.equal(report.problems.length,4); assert.ok(report.problems.every(problem=>Number.isInteger(problem.sheetRow)&&problem.reason));
  const numberWrite=writes.find(write=>write.column===documentHeaders.indexOf('Номер документа')+1);
  assert.equal(numberWrite.values[0][0],'Дополнительное соглашение №3');
});

test('158. mass creation applies One blank and Many max-plus-one numbering', () => {
  const ctx=baseContext(); ctx.getSystemSpreadsheet_=()=>({getSpreadsheetTimeZone:()=> 'UTC'});
  const headers=Array.from(vm.runInContext('SYSTEM_CONFIG.SHEETS.DOCUMENTS.requiredHeaders',ctx));
  const context={headers,headerMap:Object.fromEntries(headers.map((header,index)=>[header,index+1]))};
  const object={id:'OBJ',name:'Объект',contractNumber:'DOG',objectStatus:'Действующий',workStartDate:new Date(),workEndPlan:new Date(),workEndFact:'',responsibleForeman:'Иванов',responsibleForemanId:'ST-1'};
  const facts={documentIds:{},documentNumbersByKey:{
    'OBJ\u0000MANY':[
      {documentType:'Акт',documentNumber:2},
      {documentType:'Акт',documentNumber:'Акт №4'},
      {documentType:'Акт №3',documentNumber:''}
    ]
  }};
  ctx.generateChangeId_=(id,index)=>'CHG-'+index;
  const prepared=ctx.prepareDocumentRows_([
    {object,rule:{id:'ONE',name:'Договор',repeatability:'Один'}},
    {object,rule:{id:'MANY',name:'Акт',repeatability:'Много'}}
  ],facts,context,4,new Date(),'actor','OP-1');
  const numberIndex=headers.indexOf('Номер документа');
  const typeIndex=headers.indexOf('Тип документа');
  assert.equal(prepared.documentRows[0][numberIndex],'');
  assert.equal(prepared.documentRows[1][numberIndex],5);
  assert.equal(prepared.documentRows[0][typeIndex],'Договор');
  assert.equal(prepared.documentRows[1][typeIndex],'Акт');
});

test('159. operator card accepts and histories the three document-owned dates', () => {
  const ctx=baseContext();
  ctx.generateChangeId_=(operationId,index)=>operationId+'-'+index;
  const dates=vm.runInContext('[new Date(2026,0,2),new Date(2026,1,3),new Date(2026,2,4)]',ctx);
  const item=partialSaveFixture(ctx,{card:{
    'Дата начала работ':dates[0],
    'Дата окончания (по плану)':dates[1],
    'Дата окончания (по факту)':dates[2]
  }});
  const plan=ctx.operatorCardBuildSavePlan_([item.card],[item.fact],partialDictionaries(),dates[0],'actor','OP-1');
  assert.equal(plan.rowErrors.length,0);
  assert.deepEqual(Array.from(plan.changes,change=>change.fieldName),[
    'Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)'
  ]);
});

test('160. operator card rejects non-Date values for each work date', () => {
  for(const header of ['Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)']) {
    const ctx=baseContext();
    const item=partialSaveFixture(ctx,{card:{[header]:'2026-08-21'}});
    const plan=ctx.operatorCardBuildSavePlan_([item.card],[item.fact],partialDictionaries(),new Date(),'actor','OP-1');
    assert.equal(plan.rows.length,0,header);
    assert.match(plan.rowErrors[0].message,/должно быть пустым или корректной датой/,header);
  }
});

test('161. object synchronization cannot overwrite document work dates', () => {
  const ctx=baseContext();
  const fields=Array.from(vm.runInContext('OBJECT_SYNC_FIELDS_',ctx));
  for(const header of ['Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)']) {
    assert.ok(!fields.includes(header),header);
  }
  assert.equal(ctx.assertObjectSyncFieldContract_(),undefined);
  const sync=fs.readFileSync('SyncObjectData.gs','utf8');
  assert.match(sync,/documentOwned[\s\S]*H\.WORK_START_DATE[\s\S]*H\.WORK_END_PLAN[\s\S]*H\.WORK_END_FACT/);
});

test('162. loaded-card state validates the mapped document identity and physical row', () => {
  const ctx=baseContext();
  const contexts={
    OPERATOR_CARD:{config:{dataStartRow:6},headerMap:{'ID документа':2,'Номер строки в таблице фактов':5},sheet:{
      getLastRow:()=>6,getRange:()=>({getValues:()=>[['DOC-1','','',9]]})
    }},
    DOCUMENTS:{config:{dataStartRow:4},headerMap:{'ID документа':3},sheet:{
      getLastRow:()=>10,getRange:(row,column)=>({getValue:()=>row===9&&column===3?'DOC-1':''})
    }}
  };
  ctx.getSystemSheetContext_=key=>contexts[key];
  assert.equal(ctx.operatorCardHasLoadedRow_(),true);
});

test('163. loaded-card state rejects a corrupted identity row', () => {
  const ctx=baseContext();
  const contexts={
    OPERATOR_CARD:{config:{dataStartRow:6},headerMap:{'ID документа':1,'Номер строки в таблице фактов':2},sheet:{
      getLastRow:()=>6,getRange:()=>({getValues:()=>[['DOC-1',9]]})
    }},
    DOCUMENTS:{config:{dataStartRow:4},headerMap:{'ID документа':1},sheet:{
      getLastRow:()=>10,getRange:()=>({getValue:()=> 'OTHER'})
    }}
  };
  ctx.getSystemSheetContext_=key=>contexts[key];
  assert.equal(ctx.operatorCardHasLoadedRow_(),false);
});

test('164. reopening and resetting sidebar never reloads or clears a loaded card', () => {
  const sidebar=fs.readFileSync('OperatorSidebar.html','utf8');
  const onload=sidebar.slice(sidebar.indexOf('window.onload='));
  assert.match(onload,/getOperatorSidebarData\(\)/);
  assert.match(onload,/if\(data\.hasLoadedCard\)lastAppliedFilters=emptyFilters\(\)/);
  assert.doesNotMatch(onload,/applyOperatorFilters/);
  const reset=sidebar.match(/function resetFilters\(\)[^\n]+/)[0];
  assert.doesNotMatch(reset,/lastAppliedFilters|null|save/);
});

if (!process.exitCode) console.log(`\n${passed} tests passed.`);
