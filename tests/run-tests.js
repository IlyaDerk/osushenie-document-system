'use strict';
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');
const cp = require('child_process');
const files = ['SystemCore.gs', 'CreateObjectDocuments.gs', 'SyncObjectData.gs'];
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
        return date.toISOString();
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
test('20. protected UI files and onOpen are unchanged from merged baseline', () => {
  for (const f of ['Code.gs', 'OperatorSidebar.html']) {
    cp.execFileSync('git', ['diff', '--quiet', '36eaaf7', '--', f]);
  }
  cp.execFileSync('git', ['diff', '--quiet', '36eaaf7', '--', 'Code.gs']);
});

if (!process.exitCode) console.log(`\n${passed}/20 tests passed.`);
