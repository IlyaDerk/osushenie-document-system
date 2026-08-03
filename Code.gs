// ======================================================
// НАСТРОЙКИ
// ======================================================

const OPERATOR_CONFIG = {

  // Справочники
  OBJECTS_SHEET: 'Объекты',
  EMPLOYEES_SHEET: 'Справочник сотрудников',
  DOCUMENT_TYPES_SHEET: 'Документы Справочник',

  // Рабочий лист
  CARD_SHEET: 'Карточка операциониста',

  // Должность прораба
  FOREMAN_POSITION: 'Производитель работ (прораб)'
};


// ======================================================
// МЕНЮ ПРИ ОТКРЫТИИ GOOGLE SHEETS
// ======================================================

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Документы')
    .addItem(
      'Открыть фильтры карточки',
      'showOperatorSidebar'
    )
    .addSeparator()
    .addItem(
      'Проверить структуру системы',
      'validateSystemStructure'
    )
    .addItem(
      'Проверить email оператора',
      'testActiveUserEmail'
    )
    .addToUi();
}


// ======================================================
// ОТКРЫТИЕ SIDEBAR
// ======================================================

function showOperatorSidebar() {

  const html =
    HtmlService
      .createHtmlOutputFromFile(
        'OperatorSidebar'
      )
      .setTitle(
        'Фильтры карточки'
      );

  SpreadsheetApp
    .getUi()
    .showSidebar(html);

}


// ======================================================
// ВСЕ ДАННЫЕ ДЛЯ SIDEBAR
// ======================================================

function getOperatorFilterData() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  // ------------------------------
  // Получаем сотрудников
  // ------------------------------

  const employees =
    getEmployeesForFilter_(ss);


  // ------------------------------
  // Отдельно выделяем прорабов
  // ------------------------------

  const foremen =
    employees.filter(employee => {

      return normalize_(
        employee.position
      ) === normalize_(
        OPERATOR_CONFIG.FOREMAN_POSITION
      );

    });


  // ------------------------------
  // Возвращаем всё одним объектом
  // ------------------------------

  return {

    objects:
      getObjectsForFilter_(ss),

    foremen:
      foremen,

    employees:
      employees,

    documentTypes:
      getDocumentTypesForFilter_(ss)

  };

}


// ======================================================
// ОБЪЕКТЫ
// ======================================================

function getObjectsForFilter_(ss) {

  const sheet =
    ss.getSheetByName(
      OPERATOR_CONFIG.OBJECTS_SHEET
    );


  if (!sheet) {

    throw new Error(
      'Не найден лист "' +
      OPERATOR_CONFIG.OBJECTS_SHEET +
      '"'
    );

  }


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (!values.length) {

    return [];

  }


  // ------------------------------------
  // Ищем строку заголовков
  // ------------------------------------

  const headerRowIndex =
    values.findIndex(row => {

      return (
        findHeaderIndex_(
          row,
          ['ID Объекта']
        ) !== -1
        &&
        findHeaderIndex_(
          row,
          ['Название']
        ) !== -1
      );

    });


  if (headerRowIndex === -1) {

    throw new Error(
      'На листе "Объекты" не найдены ' +
      'заголовки "ID Объекта" и "Название".'
    );

  }


  const headers =
    values[headerRowIndex];


  // ------------------------------------
  // Индексы столбцов
  // ------------------------------------

  const idIndex =
    findHeaderIndex_(
      headers,
      ['ID Объекта']
    );


  const nameIndex =
    findHeaderIndex_(
      headers,
      ['Название']
    );


  const responsibleIdIndex =
    findHeaderIndex_(
      headers,
      [
        'ID Ответственного',
        'ID ответственного'
      ]
    );


  const responsibleNameIndex =
    findHeaderIndex_(
      headers,
      [
        'Имя ответственного',
        'ФИО ответственного'
      ]
    );


  const statusIndex =
    findHeaderIndex_(
      headers,
      ['Статус объекта']
    );


  // ------------------------------------
  // Формируем массив объектов
  // ------------------------------------

  const objects =
    values
      .slice(
        headerRowIndex + 1
      )

      .filter(row => {

        return (
          String(
            row[idIndex] || ''
          ).trim() !== ''
          &&
          String(
            row[nameIndex] || ''
          ).trim() !== ''
        );

      })

      .map(row => {

        return {

          id:
            String(
              row[idIndex] || ''
            ).trim(),

          name:
            String(
              row[nameIndex] || ''
            ).trim(),

          responsibleId:
            responsibleIdIndex >= 0
              ? String(
                  row[
                    responsibleIdIndex
                  ] || ''
                ).trim()
              : '',

          responsibleName:
            responsibleNameIndex >= 0
              ? String(
                  row[
                    responsibleNameIndex
                  ] || ''
                ).trim()
              : '',

          status:
            statusIndex >= 0
              ? String(
                  row[
                    statusIndex
                  ] || ''
                ).trim()
              : ''

        };

      });


  // ------------------------------------
  // Сортируем по названию
  // ------------------------------------

  objects.sort(
    (a, b) =>
      a.name.localeCompare(
        b.name,
        'ru'
      )
  );


  return objects;

}


// ======================================================
// СОТРУДНИКИ
// ======================================================

function getEmployeesForFilter_(ss) {

  const sheet =
    ss.getSheetByName(
      OPERATOR_CONFIG.EMPLOYEES_SHEET
    );


  if (!sheet) {

    throw new Error(
      'Не найден лист "' +
      OPERATOR_CONFIG.EMPLOYEES_SHEET +
      '"'
    );

  }


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (!values.length) {

    return [];

  }


  // ------------------------------------
  // Ищем строку заголовков
  // ------------------------------------

  const headerRowIndex =
    values.findIndex(row => {

      return (
        findHeaderIndex_(
          row,
          ['ID Сотрудника']
        ) !== -1
        &&
        findHeaderIndex_(
          row,
          [
            'ФИО сотрудника',
            'ФИО'
          ]
        ) !== -1
        &&
        findHeaderIndex_(
          row,
          ['Должность']
        ) !== -1
      );

    });


  if (headerRowIndex === -1) {

    throw new Error(
      'На листе "Справочник сотрудников" ' +
      'не найдены заголовки: ' +
      '"ID Сотрудника", ' +
      '"ФИО сотрудника", ' +
      '"Должность".'
    );

  }


  const headers =
    values[headerRowIndex];


  // ------------------------------------
  // Индексы столбцов
  // ------------------------------------

  const idIndex =
    findHeaderIndex_(
      headers,
      ['ID Сотрудника']
    );


  const nameIndex =
    findHeaderIndex_(
      headers,
      [
        'ФИО сотрудника',
        'ФИО'
      ]
    );


  const positionIndex =
    findHeaderIndex_(
      headers,
      ['Должность']
    );


  const emailIndex =
    findHeaderIndex_(
      headers,
      [
        'Электронная почта',
        'Email',
        'E-mail'
      ]
    );


  const hireDateIndex =
    findHeaderIndex_(
      headers,
      [
        'Дата найма',
        'Дата Найма'
      ]
    );


  const dismissalDateIndex =
    findHeaderIndex_(
      headers,
      [
        'Дата увольнения',
        'Дата Увольнения'
      ]
    );


  // ------------------------------------
  // Формируем сотрудников
  // ------------------------------------

  const employees =
    values
      .slice(
        headerRowIndex + 1
      )

      .filter(row => {

        return (
          String(
            row[idIndex] || ''
          ).trim() !== ''
          &&
          String(
            row[nameIndex] || ''
          ).trim() !== ''
        );

      })

      .map(row => {

        return {

          id:
            String(
              row[idIndex] || ''
            ).trim(),

          name:
            String(
              row[nameIndex] || ''
            ).trim(),

          position:
            String(
              row[positionIndex] || ''
            ).trim(),

          email:
            emailIndex >= 0
              ? String(
                  row[emailIndex] || ''
                ).trim()
              : '',

          hireDate:
            hireDateIndex >= 0
              ? String(
                  row[hireDateIndex] || ''
                ).trim()
              : '',

          dismissalDate:
            dismissalDateIndex >= 0
              ? String(
                  row[
                    dismissalDateIndex
                  ] || ''
                ).trim()
              : ''

        };

      });


  // ------------------------------------
  // Сортируем по ФИО
  // ------------------------------------

  employees.sort(
    (a, b) =>
      a.name.localeCompare(
        b.name,
        'ru'
      )
  );


  return employees;

}


// ======================================================
// ТИПЫ ДОКУМЕНТОВ
// ======================================================

function getDocumentTypesForFilter_(ss) {

  const sheet =
    ss.getSheetByName(
      OPERATOR_CONFIG.DOCUMENT_TYPES_SHEET
    );


  if (!sheet) {

    throw new Error(
      'Не найден лист "' +
      OPERATOR_CONFIG.DOCUMENT_TYPES_SHEET +
      '"'
    );

  }


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (!values.length) {

    return [];

  }


  // ------------------------------------
  // Ищем строку с заголовком
  // "Тип документа"
  // ------------------------------------

  const headerRowIndex =
    values.findIndex(row => {

      return (
        findHeaderIndex_(
          row,
          ['Тип документа']
        ) !== -1
      );

    });


  if (headerRowIndex === -1) {

    throw new Error(
      'На листе "Документы Справочник" ' +
      'не найден столбец "Тип документа".'
    );

  }


  const headers =
    values[headerRowIndex];


  // ------------------------------------
  // Индексы
  // ------------------------------------

  const nameIndex =
    findHeaderIndex_(
      headers,
      ['Тип документа']
    );


  const idIndex =
    findHeaderIndex_(
      headers,
      [
        'ID типа',
        'ID типа документа'
      ]
    );


  // ------------------------------------
  // Формируем справочник
  // ------------------------------------

  const documentTypes =
    values
      .slice(
        headerRowIndex + 1
      )

      .filter(row => {

        return (
          String(
            row[nameIndex] || ''
          ).trim() !== ''
        );

      })

      .map(row => {

        return {

          id:
            idIndex >= 0
              ? String(
                  row[idIndex] || ''
                ).trim()
              : '',

          name:
            String(
              row[nameIndex] || ''
            ).trim()

        };

      });


  documentTypes.sort(
    (a, b) =>
      a.name.localeCompare(
        b.name,
        'ru'
      )
  );


  return documentTypes;

}


// ======================================================
// ПОЛУЧЕНИЕ ФИЛЬТРОВ ИЗ SIDEBAR
// ======================================================

function applyOperatorFilters(filters) {

  // ------------------------------------
  // Защита
  // ------------------------------------

  filters =
    filters || {};


  // ------------------------------------
  // Сейчас выводим в журнал.
  //
  // На следующем этапе здесь будет:
  //
  // 1. Чтение ФКТ_Документы
  // 2. Фильтрация
  // 3. Обогащение справочниками
  // 4. Очистка карточки
  // 5. setValues()
  //
  // ------------------------------------

  console.log(
    JSON.stringify(
      filters,
      null,
      2
    )
  );


  return loadOperatorCard_(filters);

}


// ======================================================
// ЗАГРУЗКА КАРТОЧКИ
// ПОКА ЗАГЛУШКА
// ======================================================

function loadOperatorCard_(filters) {

  // Пока фактическую таблицу документов
  // мы ещё не подключили.
  //
  // Эта функция нужна, чтобы Sidebar
  // уже полностью работал и передавал
  // выбранные условия в Code.gs.


  const selected = [];


  if (filters.objectName) {

    selected.push(
      'Объект: ' +
      filters.objectName
    );

  }


  if (filters.responsibleName) {

    selected.push(
      'Прораб: ' +
      filters.responsibleName
    );

  }


  if (filters.status) {

    selected.push(
      'Статус: ' +
      filters.status
    );

  }


  if (filters.documentType) {

    selected.push(
      'Тип документа: ' +
      filters.documentType
    );

  }


  if (filters.holderName) {

    selected.push(
      'У кого: ' +
      filters.holderName
    );

  }


  return {

    success: true,

    message:
      selected.length
        ? 'Фильтры приняты'
        : 'Выбраны все документы',

    filters:
      filters

  };

}


// ======================================================
// ПОИСК НОМЕРА СТОЛБЦА ПО НАЗВАНИЮ
// ======================================================

function findHeaderIndex_(
  headers,
  possibleNames
) {

  const normalizedHeaders =
    headers.map(
      header =>
        normalize_(header)
    );


  for (
    let i = 0;
    i < possibleNames.length;
    i++
  ) {

    const target =
      normalize_(
        possibleNames[i]
      );


    const index =
      normalizedHeaders.indexOf(
        target
      );


    if (index !== -1) {

      return index;

    }

  }


  return -1;

}


// ======================================================
// НОРМАЛИЗАЦИЯ ТЕКСТА
// ======================================================

function normalize_(value) {

  return String(
    value || ''
  )
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

}
