/** Public entry points for the operator card. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Документы')
    .addItem('Открыть фильтры карточки', 'showOperatorSidebar')
    .addSeparator()
    .addItem('Проверить структуру системы', 'validateSystemStructure')
    .addItem('Проверить email оператора', 'testActiveUserEmail')
    .addToUi();
}

function showOperatorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('OperatorSidebar')
    .setTitle('Фильтры карточки');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getOperatorFilterData() {
  return operatorCardGetFilterData_();
}

function getOperatorCardState() {
  return operatorCardGetLoadedState_();
}

function applyOperatorFilters(filters) {
  return operatorCardApply_(filters);
}

function saveOperatorCardChanges(filters) {
  return operatorCardSave_(filters);
}
