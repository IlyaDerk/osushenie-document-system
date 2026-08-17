/** Controlled inputs for the source-of-truth sheet «Объекты». */
function objectControlsNormalize_(value) {
  return String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim().replace(/\s+/g, ' ').toLocaleLowerCase('ru');
}

function objectControlsForemen_() {
  const context = getSystemSheetContext_('EMPLOYEES');
  const lastRow = context.sheet.getLastRow();
  if (lastRow < context.config.dataStartRow) return [];
  const idColumn = getSystemColumn_('EMPLOYEES', H.EMPLOYEE_ID);
  const nameColumn = getSystemColumn_('EMPLOYEES', H.EMPLOYEE_NAME);
  const positionColumn = getSystemColumn_('EMPLOYEES', H.EMPLOYEE_POSITION);
  const firstColumn = Math.min(idColumn, nameColumn, positionColumn);
  const lastColumn = Math.max(idColumn, nameColumn, positionColumn);
  const allowed = {};
  SYSTEM_CONFIG.VALUES.FOREMAN_POSITIONS.forEach(function (position) {
    allowed[objectControlsNormalize_(position)] = true;
  });
  return context.sheet.getRange(
    context.config.dataStartRow, firstColumn,
    lastRow - context.config.dataStartRow + 1, lastColumn - firstColumn + 1
  ).getValues().map(function (row) {
    return {
      id: String(row[idColumn - firstColumn] == null ? '' : row[idColumn - firstColumn]).trim(),
      name: String(row[nameColumn - firstColumn] == null ? '' : row[nameColumn - firstColumn]).trim(),
      position: row[positionColumn - firstColumn]
    };
  }).filter(function (employee) {
    return employee.id && employee.name && allowed[objectControlsNormalize_(employee.position)];
  }).map(function (employee) {
    employee.label = employee.name + ' [' + employee.id + ']';
    return employee;
  });
}

function objectControlsEnsureStatuses_(context) {
  const column = getSystemColumn_('CARD_DICTIONARY', H.OBJECT_STATUS);
  const lastRow = Math.max(context.sheet.getLastRow(), context.config.dataStartRow - 1);
  const count = Math.max(lastRow - context.config.dataStartRow + 1, 0);
  const values = count ? context.sheet.getRange(context.config.dataStartRow, column, count, 1).getValues() : [];
  const existing = {};
  values.forEach(function (row) { existing[objectControlsNormalize_(row[0])] = true; });
  const missing = SYSTEM_CONFIG.VALUES.OBJECT_STATUSES.filter(function (status) {
    return !existing[objectControlsNormalize_(status)];
  });
  if (missing.length) {
    context.sheet.getRange(lastRow + 1, column, missing.length, 1).setValues(missing.map(function (status) { return [status]; }));
  }
  return SYSTEM_CONFIG.VALUES.OBJECT_STATUSES.slice();
}

/** Idempotently installs header-driven validations without changing object data. */
function setupObjectSheetControls() {
  assertSystemSheetsStructure_(['OBJECTS', 'EMPLOYEES', 'CARD_DICTIONARY']);
  const objects = getSystemSheetContext_('OBJECTS');
  const dictionary = getSystemSheetContext_('CARD_DICTIONARY');
  const foremen = objectControlsForemen_();
  const statuses = objectControlsEnsureStatuses_(dictionary);
  const rowCount = objects.sheet.getMaxRows() - objects.config.dataStartRow + 1;
  if (rowCount < 1) return { foremen: foremen.length, statuses: statuses.length };
  const foremanRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(foremen.map(function (employee) { return employee.label; }), true)
    .setAllowInvalid(false).build();
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statuses, true).setAllowInvalid(false).build();
  objects.sheet.getRange(objects.config.dataStartRow, getSystemColumn_('OBJECTS', H.RESPONSIBLE_FOREMAN), rowCount, 1).setDataValidation(foremanRule);
  objects.sheet.getRange(objects.config.dataStartRow, getSystemColumn_('OBJECTS', H.OBJECT_STATUS), rowCount, 1).setDataValidation(statusRule);
  objects.sheet.hideColumns(getSystemColumn_('OBJECTS', H.RESPONSIBLE_FOREMAN_ID));
  return { foremen: foremen.length, statuses: statuses.length };
}

/** Simple edit trigger: stores clean FIO and its unambiguous ST-ID after dropdown selection. */
function onEdit(event) {
  if (!event || !event.range || event.range.getNumRows() !== 1 || event.range.getNumColumns() !== 1) return;
  const objects = getSystemSheetContext_('OBJECTS');
  if (event.range.getSheet().getSheetId() !== objects.sheet.getSheetId() ||
      event.range.getRow() < objects.config.dataStartRow ||
      event.range.getColumn() !== getSystemColumn_('OBJECTS', H.RESPONSIBLE_FOREMAN)) return;
  const idCell = objects.sheet.getRange(event.range.getRow(), getSystemColumn_('OBJECTS', H.RESPONSIBLE_FOREMAN_ID));
  const value = String(event.value == null ? '' : event.value).trim();
  if (!value) { idCell.clearContent(); return; }
  const employee = objectControlsForemen_().filter(function (item) { return item.label === value; })[0];
  if (!employee) {
    idCell.clearContent();
    event.range.clearContent().setNote('Выберите сотрудника из выпадающего списка.');
    return;
  }
  event.range.setValue(employee.name).setNote('ST-ID определён автоматически: ' + employee.id);
  idCell.setValue(employee.id);
}
