'use strict';
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');
const cp = require('child_process');
const files = ['SystemCore.gs', 'CreateObjectDocuments.gs', 'SyncObjectData.gs', 'ArchiveChangeHistory.gs', 'OperatorCard.gs', 'Code.gs'];
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
      }
    },
    SpreadsheetApp: extra.SpreadsheetApp || { flush() {} },
    LockService: extra.LockService || {},
    Session: { getActiveUser() { return { getEmail() { return 'tester@example.com'; } }; } },
    HtmlService: extra.HtmlService || {}
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
    'ID документа', 'ID объекта', 'Статус записи', 'Номер договора', 'Статус объекта',
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
  'Номер договора': 'A-1', 'Статус объекта': 'В работе',
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
    'ID документа','ID объекта','Тип документа','ID типа документа','Номер договора','Статус документа',
    'Статус объекта','Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)',
    'Дата создания','Дата обновления','Кто обновил (email)','Ответственный прораб','ID ответственного прораба',
    'Дата изменения статуса документа','Источник создания','Статус записи'
  ];
  const context = { headers, headerMap: Object.fromEntries(headers.map((h, i) => [h, i + 1])) };
  const missing = Array.from({ length: count }, (_, i) => ({
    object: {
      id: 'OBJ' + String(i + 1).padStart(2, '0'), contractNumber: 'DOG-' + (i + 1),
      objectStatus: 'В работе', workStartDate: new Date('2026-08-01T00:00:00Z'),
      workEndPlan: new Date('2026-08-31T00:00:00Z'), workEndFact: i === 0 ? new Date('2026-08-20T00:00:00Z') : '',
      responsibleForeman: 'Иванов', responsibleForemanId: 'ST-1'
    },
    rule: { id: 'DT-1', name: 'Акт' }
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
  ['ID документа: DOC-OBJ01-0001','ID объекта: OBJ01','Тип документа: Акт','ID типа документа: DT-1',
   'Номер договора: DOG-1','Статус документа: Ожидает заполнения','Статус объекта: В работе',
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
  const headers = ['ID документа','ID объекта','Тип документа','Номер договора','Дата документа','Статус документа','Оригинал / ЭДО','Комментарий','У кого документ','Где документ','Кто передал','Оплачен','Сумма документа','ГУ (Да/Нет)','Условия ГУ','Статус объекта','Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)','Дата создания','Дата обновления','ID типа документа','Кто обновил (email)','Ответственный прораб'];
  assert.equal(ctx.operatorCardValidateHeaders_(headers.concat(['extra']), headers, 'Документы объектов', 'Карточка'), true);
  const swapped = headers.slice(); [swapped[0], swapped[1]] = [swapped[1], swapped[0]];
  assert.throws(() => ctx.operatorCardValidateHeaders_(swapped, headers, 'Документы объектов', 'Карточка'), /позиция 2/);
  assert.throws(() => ctx.operatorCardValidateHeaders_(headers.concat(['Техническое поле']), headers.concat(['Лишний заголовок']), 'Документы объектов', 'Карточка'), /колонка 25.*Лишний заголовок/);
  assert.equal(ctx.operatorCardValidateHeaders_(headers.concat(['Техническое поле']), headers.concat(['', '  ']), 'Документы объектов', 'Карточка'), true);
});
test('55. active rows, combined ID filters and normalized status are applied', () => {
  const ctx = baseContext();
  const indexes = { documentId:0, objectId:1, documentTypeId:2, documentStatus:3, holderId:4, foremanId:5, createdAt:6, recordStatus:7 };
  const rows = [
    ['DOC-2','2','T1',' ГОТОВ ','E1','F1','', ' Активная '],
    ['DOC-3','2','T1','готов','E1','F1','', 'Архивная'],
    ['DOC-4','2','T2','готов','E1','F1','', 'Удалённая']
  ]; while(rows[0].length<31) rows.forEach(r=>r.push(''));
  const filters = ctx.operatorCardNormalizeFilters_({ object:{allObjects:false,objectId:'2'}, documentTypeId:'T1', documentStatus:'готов', holderId:'E1', foremanId:'F1' });
  const result = ctx.operatorCardPrepareRows_(rows,indexes,filters,{active:false},4);
  assert.equal(result.activeCount,1); assert.equal(result.rows.length,1); assert.equal(result.cardRows[0].length,24);
});
test('56. natural sorting uses object, type, document and physical row', () => {
  const ctx = baseContext(); const i={documentId:0,objectId:1,documentTypeId:2,documentStatus:3,holderId:4,foremanId:5,createdAt:6,recordStatus:7};
  const rows=[['DOC-10','2','10','','','','','Активная'],['DOC-2','2','2','','','','','Активная'],['DOC-1','10','1','','','','','Активная']]; rows.forEach(r=>{while(r.length<31)r.push('')});
  const result=ctx.operatorCardPrepareRows_(rows,i,ctx.operatorCardNormalizeFilters_({}),{active:false},4);
  assert.deepEqual(Array.from(result.rows,x=>x.row[0]),['DOC-2','DOC-10','DOC-1']);
});
test('57. active duplicate group is retained and reports physical rows once per ID', () => {
  const ctx=baseContext(); const i={documentId:0,objectId:1,documentTypeId:2,documentStatus:3,holderId:4,foremanId:5,createdAt:6,recordStatus:7};
  const rows=[['D1','A','','','','','','Активная'],['D1','B','','','','','','Активная'],['D1','A','','','','','','Архивная']]; rows.forEach(r=>{while(r.length<31)r.push('')});
  const result=ctx.operatorCardPrepareRows_(rows,i,ctx.operatorCardNormalizeFilters_({object:{allObjects:false,objectId:'A'}}),{active:false},18);
  assert.equal(result.rows.length,1); assert.equal(result.duplicateIdsCount,1); assert.match(result.warnings[0],/18 и 19/); assert.doesNotMatch(result.warnings[0],/20/);
});
test('58. card replacement performs one 24-column setValues and clears tail', () => {
  const ctx=baseContext(); let calls=0, written;
  const sheet={getLastRow:()=>8,getRange(row,col,count,width){return {getValues:()=>[['old1'],['old2'],['']],setValues(values){calls++;written={row,col,count,width,values}}}}};
  ctx.operatorCardReplace_({sheet,config:{dataStartRow:6},headerMap:{'ID документа':1}},[[1,2].concat(Array(22).fill(''))]);
  assert.equal(calls,1); assert.equal(written.width,24); assert.equal(written.values.length,2); assert.ok(written.values[1].every(v=>v===''));
});
test('59. zero result clears old card in one batch', () => {
  const ctx=baseContext(); let values; const sheet={getLastRow:()=>6,getRange(){return {getValues:()=>[['old']],setValues(v){values=v}}}};
  ctx.operatorCardReplace_({sheet,config:{dataStartRow:6},headerMap:{'ID документа':1}},[]); assert.equal(values.length,1); assert.equal(values[0].length,24); assert.ok(values[0].every(v=>v===''));
});
test('60. operator implementation never writes facts, dictionaries, or change history', () => {
  const text=fs.readFileSync('OperatorCard.gs','utf8');
  assert.doesNotMatch(text,/appendRow|\.clear\s*\(|deleteRows|insertRows/);
  assert.doesNotMatch(text,/getSystemSheetContext_\('CHANGE_HISTORY'\)/);
  assert.match(text,/getRange\(start, cardStartColumn, writeCount, 24\)\.setValues/);
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
  const ctx=baseContext(); const i={documentId:0,objectId:1,documentTypeId:2,documentStatus:3,holderId:4,foremanId:5,createdAt:6,recordStatus:7};
  const valid=vm.runInContext("new Date('2026-08-10T12:00:00Z')",ctx);
  const rows=[['A','','','','','',valid,'Активная'],['B','','','','','','','Активная'],['C','','','','','','bad','Активная']];rows.forEach(r=>{while(r.length<31)r.push('')});
  const range=ctx.operatorCardParseDateRange_('2026-08-10','2026-08-10','UTC');
  const result=ctx.operatorCardPrepareRows_(rows,i,ctx.operatorCardNormalizeFilters_({}),range,4);
  assert.equal(result.rows.length,1);assert.equal(result.invalidDatesCount,1);assert.match(result.warnings[0],/некорректной датой/);
});


const operatorWorkflowCardHeaders = ['ID документа','ID объекта','Тип документа','Номер договора','Дата документа','Статус документа','Оригинал / ЭДО','Комментарий','У кого документ','Где документ','Кто передал','Оплачен','Сумма документа','ГУ (Да/Нет)','Условия ГУ','Статус объекта','Дата начала работ','Дата окончания (по плану)','Дата окончания (по факту)','Дата создания','Дата обновления','ID типа документа','Кто обновил (email)','Ответственный прораб'];
const operatorWorkflowDocumentHeaders = operatorWorkflowCardHeaders.concat(['ID сотрудника — у кого документ','ID сотрудника — кто передал','ID ответственного прораба','Дата изменения статуса документа','Отчётный период','Источник создания','Статус записи']);
const operatorWorkflowOperationHeaders = ['ID операции','Дата и время начала','Дата и время завершения','Кто запустил (email)','Источник операции','Тип операции','Статус операции','Документов загружено в карточку','Документов с изменениями','Строк факта обновлено','Полей изменено','Дублирующихся ID найдено','Ошибок','Время выполнения, сек.','Текст ошибки / комментарий'];

function operatorWorkflowFixture(options = {}) {
  const ctx = baseContext();
  const calls = { cardWrites: 0, operationRows: [], sheetKeys: [] };
  const documentMap = Object.fromEntries(operatorWorkflowDocumentHeaders.map((header, index) => [header, index + 1]));
  const documents = {
    config: { name: 'Документы объектов', headerRow: 3, dataStartRow: 4 }, headers: operatorWorkflowDocumentHeaders.slice(), headerMap: documentMap,
    sheet: { getLastRow: () => 3, getRange() { throw new Error('fact data range must not be read for empty fixture'); } }
  };
  const card = { config: { name: 'Карточка операциониста', headerRow: 5, dataStartRow: 6 }, headers: operatorWorkflowCardHeaders.slice(), headerMap: {'ID документа': 1}, sheet: {} };
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
  ctx.operatorCardPrepareRows_ = () => ({ activeCount: options.activeCount == null ? 2 : options.activeCount, rows: [], cardRows: Array.from({length: options.loadedCount == null ? 2 : options.loadedCount}, (_, index) => [index + 1].concat(Array(23).fill(''))), warnings: (options.warnings || []).slice(), duplicateIdsCount: options.duplicateIdsCount || 0, invalidDatesCount: 0 });
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
    createdAt: physicalIndexes['Дата создания'], recordStatus: physicalIndexes['Статус записи']
  }, ctx.operatorCardNormalizeFilters_({}), {active:false}, 4, 1);
  assert.equal(prepared.cardRows.length, 1);
  assert.equal(prepared.cardRows[0].length, 24);
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
  ctx.operatorCardReplace_({sheet, config:{dataStartRow:6}, headerMap:{'ID документа':3}}, prepared.cardRows);
  assert.equal(rangeCalls.length, 2);
  assert.deepEqual(rangeCalls.map(call => call.column), [3, 3]);
  assert.equal(rangeCalls[0].columnCount, 1);
  assert.equal(rangeCalls[1].columnCount, 24);
  assert.ok(rangeCalls.every(call => call.column >= 3), 'columns left of the card block must not be touched');
});

if (!process.exitCode) console.log(`\n${passed} tests passed.`);
