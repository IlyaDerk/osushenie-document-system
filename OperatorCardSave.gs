/** Safe write-back workflow for «Карточка операциониста». */
const OPERATOR_CARD_SAVE_WRITABLE_ = [
  H.DOCUMENT_DATE, H.DOCUMENT_STATUS, H.ORIGINAL_EDO, H.COMMENT,
  H.DOCUMENT_HOLDER, H.DOCUMENT_LOCATION, H.TRANSFERRED_BY, H.PAID,
  H.DOCUMENT_AMOUNT, H.GU_FLAG, H.GU_TERMS, H.RECORD_STATUS
];

function operatorCardSaveEqual_(a, b) {
  if (a instanceof Date || b instanceof Date) return a instanceof Date && b instanceof Date && a.getTime() === b.getTime();
  return a === b;
}
function operatorCardSaveEmpty_(v) { return v === '' || v == null; }
function operatorCardSaveLabelMap_(items) {
  const map = {};
  items.forEach(function (x) { map[operatorCardDisplayLabel_(x.name, x.id)] = x; });
  return map;
}
function operatorCardSaveAssertDictionary_(value, allowed, header) {
  if (operatorCardSaveEmpty_(value)) return;
  if (!allowed.some(function (x) { return operatorCardFold_(x) === operatorCardFold_(value); }))
    throw new Error('Поле «' + header + '»: неизвестное значение «' + value + '». Выберите значение из актуального справочника.');
}
function operatorCardSaveVersion_(v) { return v instanceof Date && !isNaN(v.getTime()) ? v.getTime() : String(v == null ? '' : v); }

/** Pure pre-write planner; it performs every identity/type/dictionary/duplicate check. */
function operatorCardBuildSavePlan_(cardItems, factItems, dictionaries, now, email, operationId) {
  const factByRow = {}, activeGroups = {};
  factItems.forEach(function (item) {
    factByRow[item.sheetRow] = item;
    if (operatorCardFold_(item.values[H.RECORD_STATUS]) === operatorCardFold_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS)) {
      const id = operatorCardNormalizeText_(item.values[H.DOCUMENT_ID]);
      (activeGroups[id] = activeGroups[id] || []).push(item.sheetRow);
    }
  });
  const duplicateIds = Object.keys(activeGroups).filter(function (id) { return id && activeGroups[id].length > 1; });
  const holderMap = operatorCardSaveLabelMap_(dictionaries.holders);
  const employeeMap = operatorCardSaveLabelMap_(dictionaries.employees);
  const dictionaryHeaders = [H.DOCUMENT_STATUS,H.ORIGINAL_EDO,H.DOCUMENT_LOCATION,H.PAID,H.GU_FLAG,H.RECORD_STATUS];
  const cardRowsByDocument = {}, rowPlans = [], changes = [];
  cardItems.forEach(function (card) {
    const rowNumber = Number(card.values[H.FACT_ROW_NUMBER]);
    if (!Number.isInteger(rowNumber)) throw new Error('Строка карточки «' + card.values[H.DOCUMENT_ID] + '»: отсутствует корректный скрытый «' + H.FACT_ROW_NUMBER + '».');
    const fact = factByRow[rowNumber];
    if (!fact) throw new Error('Физическая строка факта ' + rowNumber + ' не существует. Нажмите «Применить» заново.');
    [H.DOCUMENT_ID,H.OBJECT_ID,H.DOCUMENT_TYPE_ID].forEach(function (header) {
      if (!operatorCardSaveEqual_(card.values[header], fact.values[header])) throw new Error('Строка факта ' + rowNumber + ': не совпадает «' + header + '». Сохранение заблокировано.');
    });
    if (operatorCardFold_(fact.values[H.RECORD_STATUS]) !== operatorCardFold_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS)) throw new Error('Строка факта ' + rowNumber + ' уже не активна. Нажмите «Применить» заново.');
    if (operatorCardSaveVersion_(card.values[H.UPDATED_AT]) !== operatorCardSaveVersion_(fact.values[H.UPDATED_AT])) throw new Error('Документ «' + fact.values[H.DOCUMENT_ID] + '», строка факта ' + rowNumber + ': «' + H.UPDATED_AT + '» изменилась. Нажмите «Применить» заново.');
    SYSTEM_CONFIG.CARD_FIELD_MAP.filter(function (m) { return !m.editable && !m.technical; }).forEach(function (m) {
      if (!operatorCardSaveEqual_(card.values[m.cardHeader], fact.values[m.factHeader])) throw new Error('Поле только для чтения «' + m.cardHeader + '» было изменено. Сохранение заблокировано.');
    });
    const id = operatorCardNormalizeText_(fact.values[H.DOCUMENT_ID]);
    (cardRowsByDocument[id] = cardRowsByDocument[id] || []).push(rowNumber);
    const proposed = {}, rowChanges = [];
    OPERATOR_CARD_SAVE_WRITABLE_.forEach(function (header) {
      let value = card.values[header];
      if (header === H.DOCUMENT_DATE && !operatorCardSaveEmpty_(value) && (!(value instanceof Date) || isNaN(value.getTime()))) throw new Error('Поле «' + header + '» должно быть пустым или корректной датой.');
      if (header === H.DOCUMENT_AMOUNT && !operatorCardSaveEmpty_(value) && (typeof value !== 'number' || !isFinite(value))) throw new Error('Поле «' + header + '» должно быть пустым или числом.');
      if (dictionaryHeaders.indexOf(header) !== -1) operatorCardSaveAssertDictionary_(value, dictionaries.card[header] || [], header);
      if (header === H.DOCUMENT_HOLDER) {
        if (operatorCardSaveEmpty_(value)) { value=''; proposed[H.HOLDER_EMPLOYEE_ID]=''; }
        else if (holderMap[value]) { proposed[H.HOLDER_EMPLOYEE_ID]=holderMap[value].id; value=holderMap[value].name; }
        else if (!fact.values[H.HOLDER_EMPLOYEE_ID] && value === fact.values[H.DOCUMENT_HOLDER]) { proposed[H.HOLDER_EMPLOYEE_ID]=''; }
        else throw new Error('Поле «' + header + '»: держатель не найден в справочниках сотрудников/клиентов.');
      }
      if (header === H.TRANSFERRED_BY) {
        if (operatorCardSaveEmpty_(value)) { value=''; proposed[H.TRANSFERRED_BY_EMPLOYEE_ID]=''; }
        else if (employeeMap[value]) { proposed[H.TRANSFERRED_BY_EMPLOYEE_ID]=employeeMap[value].id; value=employeeMap[value].name; }
        else if (!fact.values[H.TRANSFERRED_BY_EMPLOYEE_ID] && value === fact.values[H.TRANSFERRED_BY]) { proposed[H.TRANSFERRED_BY_EMPLOYEE_ID]=''; }
        else throw new Error('Поле «' + header + '»: «Кто передал» должен быть выбран из справочника сотрудников.');
      }
      proposed[header] = value;
    });
    Object.keys(proposed).forEach(function (header) {
      if (!operatorCardSaveEqual_(proposed[header], fact.values[header])) rowChanges.push({ header:header, oldValue:fact.values[header], newValue:proposed[header] });
    });
    const statusChange = rowChanges.some(function(c){return c.header===H.RECORD_STATUS;});
    const businessChange = rowChanges.some(function(c){return c.header!==H.RECORD_STATUS;});
    if (statusChange && businessChange) throw new Error('Строка факта ' + rowNumber + ': нельзя одновременно менять «Статус записи» и другие поля.');
    rowPlans.push({ sheetRow:rowNumber, documentId:id, objectId:fact.values[H.OBJECT_ID], changes:rowChanges, fact:fact });
  });
  duplicateIds.forEach(function (id) {
    const expected = activeGroups[id].slice().sort().join(','), present = (cardRowsByDocument[id] || []).slice().sort().join(',');
    if (expected !== present) throw new Error('Дубль «' + id + '» представлен в карточке не полностью. Активные строки факта: ' + activeGroups[id].join(', ') + '.');
    const plans = rowPlans.filter(function(p){return p.documentId===id;});
    if (plans.some(function(p){return p.changes.some(function(c){return c.header!==H.RECORD_STATUS;});})) throw new Error('Для дубля «' + id + '» разрешено менять только «Статус записи».');
    const remaining = plans.filter(function(p){ const c=p.changes.find(function(x){return x.header===H.RECORD_STATUS;}); return !c || operatorCardFold_(c.newValue)===operatorCardFold_(SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS); }).length;
    if (remaining > 1) throw new Error('Дубль «' + id + '» не устранён: после сохранения останется ' + remaining + ' активных строк.');
  });
  let sequence=0; rowPlans.forEach(function(plan){ plan.changes.forEach(function(c){ changes.push({changeId:generateChangeId_(operationId,++sequence),operationId:operationId,changedAt:now,userEmail:email,documentId:plan.documentId,objectId:plan.objectId,factRow:plan.sheetRow,fieldName:c.header,oldValue:c.oldValue,newValue:c.newValue}); }); });
  return { rows:rowPlans.filter(function(p){return p.changes.length;}), changes:changes, duplicateIds:duplicateIds };
}

function operatorCardSaveRead_(context) {
  const count = Math.max(context.sheet.getLastRow()-context.config.dataStartRow+1,0); if (!count) return [];
  const rows=context.sheet.getRange(context.config.dataStartRow,1,count,context.headers.length).getValues();
  return rows.filter(function(r){return r.some(function(v){return !operatorCardSaveEmpty_(v);});}).map(function(row,i){const values={};context.headers.forEach(function(h,j){values[h]=row[j];});return {sheetRow:context.config.dataStartRow+i,values:values};});
}
function operatorCardSaveDictionaries_() {
  const filter=operatorCardGetFilterData_(), card={};
  [H.DOCUMENT_STATUS,H.ORIGINAL_EDO,H.DOCUMENT_LOCATION,H.PAID,H.GU_FLAG,H.RECORD_STATUS].forEach(function(h){card[h]=operatorCardReadUniqueColumn_('CARD_DICTIONARY',h);});
  [SYSTEM_CONFIG.VALUES.ARCHIVED_RECORD_STATUS,SYSTEM_CONFIG.VALUES.DELETED_RECORD_STATUS].forEach(function(v){if(!card[H.RECORD_STATUS].some(function(x){return operatorCardFold_(x)===operatorCardFold_(v);}))throw new Error('В «Справочник для КО» → «Статус записи» отсутствует обязательное значение «'+v+'». Добавьте его вручную.');});
  return {holders:filter.holders,employees:filter.employees,card:card};
}
function operatorCardSaveWriteFacts_(context, plan, now, email) {
  plan.rows.forEach(function(p){p.changes.push({header:H.UPDATED_AT,newValue:now},{header:H.UPDATED_BY_EMAIL,newValue:email});if(p.changes.some(function(c){return c.header===H.DOCUMENT_STATUS;}))p.changes.push({header:H.DOCUMENT_STATUS_CHANGED_AT,newValue:now});});
  const headers=OPERATOR_CARD_SAVE_WRITABLE_.concat([H.HOLDER_EMPLOYEE_ID,H.TRANSFERRED_BY_EMPLOYEE_ID,H.UPDATED_AT,H.UPDATED_BY_EMAIL,H.DOCUMENT_STATUS_CHANGED_AT]);
  headers.forEach(function(header){const col=context.headerMap[sysNormalizeHeader_(header)];const items=plan.rows.map(function(p){const c=p.changes.find(function(x){return x.header===header;});return c?{row:p.sheetRow,value:c.newValue}:null;}).filter(Boolean).sort(function(a,b){return a.row-b.row;});items.forEach(function(x){context.sheet.getRange(x.row,col,1,1).setValues([[x.value]]);});});
}
function operatorCardSaveWriteChanges_(changes){if(!changes.length)return;const c=getSystemSheetContext_('CHANGE_HISTORY');const rows=changes.map(function(x){const v={};v[H.CHANGE_ID]=x.changeId;v[H.OPERATION_ID]=x.operationId;v[H.CHANGE_DATETIME]=x.changedAt;v[H.CHANGED_BY_EMAIL]=x.userEmail;v[H.ACTION_TYPE]=SYSTEM_CONFIG.VALUES.CHANGE_ACTION_EDIT;v[H.DOCUMENT_ID]=x.documentId;v[H.OBJECT_ID]=x.objectId;v[H.FACT_ROW_NUMBER]=x.factRow;v[H.FIELD_NAME]=x.fieldName;v[H.OLD_VALUE]=x.oldValue;v[H.NEW_VALUE]=x.newValue;v[H.CHANGE_SOURCE]=SYSTEM_CONFIG.VALUES.OPERATOR_CARD_SOURCE;return c.headers.map(function(h){return Object.prototype.hasOwnProperty.call(v,h)?v[h]:'';});});c.sheet.getRange(Math.max(c.sheet.getLastRow()+1,c.config.dataStartRow),1,rows.length,c.headers.length).setValues(rows);}
function operatorCardSaveWriteOperation_(o){const c=getSystemSheetContext_('OPERATION_HISTORY'),v={};v[H.OPERATION_ID]=o.id;v[H.OPERATION_STARTED_AT]=o.started;v[H.OPERATION_FINISHED_AT]=o.finished;v[H.OPERATION_STARTED_BY]=o.email;v[H.OPERATION_SOURCE]=SYSTEM_CONFIG.VALUES.OPERATOR_CARD_SOURCE;v[H.OPERATION_TYPE]=SYSTEM_CONFIG.VALUES.OPERATOR_CARD_SAVE_OPERATION_TYPE;v[H.OPERATION_STATUS]=o.status;v[H.DOCUMENTS_LOADED]=0;v[H.DOCUMENTS_CHANGED]=o.documents;v[H.FACT_ROWS_UPDATED]=o.rows;v[H.FIELDS_CHANGED]=o.fields;v[H.DUPLICATE_IDS_FOUND]=o.duplicates;v[H.ERRORS_COUNT]=o.error?1:0;v[H.EXECUTION_SECONDS]=(o.finished-o.started)/1000;v[H.ERROR_TEXT]=o.comment||'';c.sheet.getRange(Math.max(c.sheet.getLastRow()+1,c.config.dataStartRow),1,1,c.headers.length).setValues([c.headers.map(function(h){return Object.prototype.hasOwnProperty.call(v,h)?v[h]:'';})]);}

function operatorCardSave_(filters) {
  const started=new Date(), id=generateOperationId_(started), email=getActiveUserEmail_(); let result;
  try { result=withDocumentLock_(function(){const docs=getSystemSheetContext_('DOCUMENTS'),card=getSystemSheetContext_('OPERATOR_CARD');operatorCardValidateHeaders_(docs.headers,card.headers,docs.config.name,card.config.name,docs.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)],card.headerMap[sysNormalizeHeader_(H.DOCUMENT_ID)]);const plan=operatorCardBuildSavePlan_(operatorCardSaveRead_(card),operatorCardSaveRead_(docs),operatorCardSaveDictionaries_(),new Date(),email,id);operatorCardSaveWriteFacts_(docs,plan,new Date(),email);operatorCardSaveWriteChanges_(plan.changes);const changed={};plan.rows.forEach(function(p){changed[p.documentId]=true;});const status=plan.rows.length?(plan.duplicateIds.length?SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS:SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS):SYSTEM_CONFIG.VALUES.OPERATION_STATUS_NO_CHANGES;operatorCardSaveWriteOperation_({id:id,started:started,finished:new Date(),email:email,status:status,documents:Object.keys(changed).length,rows:plan.rows.length,fields:plan.changes.length,duplicates:plan.duplicateIds.length,error:false,comment:plan.duplicateIds.length?'Дубли устранены через физические строки: '+plan.duplicateIds.join(', '):''});return {success:true,status:status,operationId:id,documentsChanged:Object.keys(changed).length,factRowsUpdated:plan.rows.length,fieldsChanged:plan.changes.length,duplicateIdsCount:plan.duplicateIds.length,warnings:plan.duplicateIds.length?['Дубли устранены: '+plan.duplicateIds.join(', ')]:[]};}); }
  catch(e){try{operatorCardSaveWriteOperation_({id:id,started:started,finished:new Date(),email:email,status:SYSTEM_CONFIG.VALUES.OPERATION_STATUS_ERROR,documents:0,rows:0,fields:0,duplicates:0,error:true,comment:e.message||String(e)});}catch(ignored){}throw e;}
  try { operatorCardApply_(filters,true); result.refreshed=true; } catch(e){result.refreshed=false;result.warnings.push('Данные сохранены, но карточку не удалось обновить. Нажмите «Применить» вручную: '+(e.message||e));result.status=SYSTEM_CONFIG.VALUES.OPERATION_STATUS_SUCCESS_WITH_WARNINGS;}
  result.message='Изменено документов: '+result.documentsChanged+', строк: '+result.factRowsUpdated+', полей: '+result.fieldsChanged+'.';return result;
}
