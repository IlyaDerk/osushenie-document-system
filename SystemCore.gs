/**
 * ============================================================
 * ТЕХНИЧЕСКОЕ ЯДРО СИСТЕМЫ ДОКУМЕНТООБОРОТА «ОСУШЕНИЕ»
 * ============================================================
 *
 * Файл: SystemCore.gs
 *
 * Назначение:
 * 1. Хранить единый справочник названий заголовков.
 * 2. Хранить конфигурацию листов и строк.
 * 3. Проверять наличие листов и обязательных заголовков.
 * 4. Проверять повторяющиеся заголовки.
 * 5. Находить столбцы по названию, а не по букве.
 * 6. Проверять соответствие карточки таблице документов.
 * 7. Получать email пользователя.
 * 8. Блокировать одновременную запись.
 *
 * Заголовки сравниваются строго с учётом регистра.
 * Лишние пробелы и неразрывные пробелы игнорируются.
 */


/**
 * ============================================================
 * 1. ЕДИНЫЙ СПРАВОЧНИК ЗАГОЛОВКОВ
 * ============================================================
 *
 * Один смысловой столбец получает одно каноническое название.
 * Во всех конфигурациях ниже используются значения из этого объекта.
 */
const SYSTEM_HEADERS = Object.freeze({
  // Общие идентификаторы
  CHANGE_ID: 'ID изменения',
  OPERATION_ID: 'ID операции',
  DOCUMENT_ID: 'ID документа',
  OBJECT_ID: 'ID объекта',
  DOCUMENT_TYPE_ID: 'ID типа документа',

  // Основные поля документа
  DOCUMENT_TYPE: 'Тип документа',
  CONTRACT_NUMBER: 'Номер договора',
  DOCUMENT_DATE: 'Дата документа',
  DOCUMENT_STATUS: 'Статус документа',
  ORIGINAL_EDO: 'Оригинал / ЭДО',
  COMMENT: 'Комментарий',

  // Движение документа
  DOCUMENT_HOLDER: 'У кого документ',
  DOCUMENT_LOCATION: 'Где документ',
  TRANSFERRED_BY: 'Кто передал',

  // Финансовые поля
  PAID: 'Оплачен',
  DOCUMENT_AMOUNT: 'Сумма документа',
  GU_FLAG: 'ГУ (Да/Нет)',
  GU_TERMS: 'Условия ГУ',

  // Данные объекта
  OBJECT_NAME: 'Название',
  OBJECT_ADDRESS: 'Адрес',
  OBJECT_STATUS: 'Статус объекта',
  WORK_START_DATE: 'Дата начала работ',
  WORK_END_PLAN: 'Дата окончания (по плану)',
  WORK_END_FACT: 'Дата окончания (по факту)',
  RESPONSIBLE_FOREMAN: 'Ответственный прораб',
  RESPONSIBLE_FOREMAN_ID: 'ID ответственного прораба',

  // Системные поля документа
  CREATED_AT: 'Дата создания',
  UPDATED_AT: 'Дата обновления',
  UPDATED_BY_EMAIL: 'Кто обновил (email)',
  HOLDER_EMPLOYEE_ID: 'ID сотрудника — у кого документ',
  TRANSFERRED_BY_EMPLOYEE_ID: 'ID сотрудника — кто передал',
  DOCUMENT_STATUS_CHANGED_AT: 'Дата изменения статуса документа',
  REPORTING_PERIOD: 'Отчётный период',
  CREATION_SOURCE: 'Источник создания',
  RECORD_STATUS: 'Статус записи',

  // История изменений
  CHANGE_DATETIME: 'Дата и время изменения',
  CHANGED_BY_EMAIL: 'Кто изменил (email)',
  ACTION_TYPE: 'Тип действия',
  FACT_ROW_NUMBER: 'Номер строки в таблице фактов',
  FIELD_NAME: 'Название поля',
  OLD_VALUE: 'Старое значение',
  NEW_VALUE: 'Новое значение',
  CHANGE_SOURCE: 'Источник изменения',

  // История операций
  OPERATION_STARTED_AT: 'Дата и время начала',
  OPERATION_FINISHED_AT: 'Дата и время завершения',
  OPERATION_STARTED_BY: 'Кто запустил (email)',
  OPERATION_SOURCE: 'Источник операции',
  OPERATION_TYPE: 'Тип операции',
  OPERATION_STATUS: 'Статус операции',
  DOCUMENTS_LOADED: 'Документов загружено в карточку',
  DOCUMENTS_CHANGED: 'Документов с изменениями',
  FACT_ROWS_UPDATED: 'Строк факта обновлено',
  FIELDS_CHANGED: 'Полей изменено',
  DUPLICATE_IDS_FOUND: 'Дублирующихся ID найдено',
  ERRORS_COUNT: 'Ошибок',
  EXECUTION_SECONDS: 'Время выполнения, сек.',
  ERROR_TEXT: 'Текст ошибки / комментарий',

  // Справочник сотрудников
  EMPLOYEE_ID: 'ID Сотрудника',
  EMPLOYEE_NAME: 'ФИО сотрудника',
  EMPLOYEE_POSITION: 'Должность',
  WORK_PHONE: 'Телефон рабочий',
  TELEGRAM_NICK: 'ник в телеграмм',
  EMAIL: 'Электронная почта',
  EMPLOYEE_GROUP: 'Группа',
  HIRE_DATE: 'Дата найма',
  DISMISSAL_DATE: 'Дата увольнения',

  // Справочник клиентов
  CLIENT_ID: 'ID клиента',
  CLIENT_NAME: 'Наименование клиента',
  PHONE: 'Телефон',
  CLIENT_STATUS: 'Статус клиента',

  // Справочник типов документов
  MANDATORY: 'Обязательность',
  CREATE_ON_OBJECT_CREATION: 'Создавать при создании объекта',
  EXECUTION_DEADLINE: 'Срок на исполнение от даты создания',
  REPEATABILITY: 'Повторяемость'
});


/**
 * Короткое имя для использования в конфигурации.
 */
const H = SYSTEM_HEADERS;


/**
 * ============================================================
 * 2. КОНФИГУРАЦИЯ СИСТЕМЫ
 * ============================================================
 */
const SYSTEM_CONFIG = {
  VERSION: '0.4.0',

  /**
   * Максимальное время ожидания блокировки.
   */
  LOCK_WAIT_MS: 30000,

  /**
   * Системные значения.
   */
  VALUES: {
    ACTIVE_RECORD_STATUS: 'Активная',
    ARCHIVED_RECORD_STATUS: 'Архивная',
    DELETED_RECORD_STATUS: 'Удалённая',

    INITIAL_DOCUMENT_STATUS: 'Ожидает заполнения',

    AUTO_CREATE_DOCUMENTS: 'Автоматически',

    OPERATION_STATUS_SUCCESS: 'Успешно',
    OPERATION_STATUS_SUCCESS_WITH_WARNINGS:
      'Успешно с предупреждениями',
    OPERATION_STATUS_NO_CHANGES: 'Без изменений',
    OPERATION_STATUS_ERROR: 'Ошибка',

    OBJECT_DOCUMENT_CREATION_OPERATION_TYPE:
      'Создание недостающих документов по объектам',

    OBJECT_DATA_SYNC_OPERATION_TYPE: 'Синхронизация данных объектов',

    CHANGE_ACTION_CREATE: 'Создание',
    CHANGE_ACTION_SYNC: 'Синхронизация',

    OPERATOR_CARD_SOURCE: 'Карточка операциониста',
    OBJECT_CREATION_SOURCE: 'Создание документов по объекту',
    OBJECT_DATA_SYNC_SOURCE: 'Синхронизация данных объектов',
    OBJECT_SHEET_CHANGE_SOURCE: 'Лист объектов',
    AUTOMATION_SOURCE: 'Автоматизация'
  },

  /**
   * Префиксы будущих технических ID.
   */
  ID_PREFIXES: {
    DOCUMENT: 'DOC-',
    EMPLOYEE: 'ST-',
    CLIENT: 'CL-',
    OPERATION: 'OP-',
    CHANGE: 'CHG-'
  },

  /**
   * Конфигурация всех рабочих листов.
   */
  SHEETS: {
    CHANGE_HISTORY: {
      name: 'История изменений',
      headerRow: 2,
      dataStartRow: 3,
      requiredHeaders: [
        H.CHANGE_ID,
        H.OPERATION_ID,
        H.CHANGE_DATETIME,
        H.CHANGED_BY_EMAIL,
        H.ACTION_TYPE,
        H.DOCUMENT_ID,
        H.OBJECT_ID,
        H.FACT_ROW_NUMBER,
        H.FIELD_NAME,
        H.OLD_VALUE,
        H.NEW_VALUE,
        H.CHANGE_SOURCE
      ]
    },

    OPERATION_HISTORY: {
      name: 'История операций',
      headerRow: 2,
      dataStartRow: 3,
      requiredHeaders: [
        H.OPERATION_ID,
        H.OPERATION_STARTED_AT,
        H.OPERATION_FINISHED_AT,
        H.OPERATION_STARTED_BY,
        H.OPERATION_SOURCE,
        H.OPERATION_TYPE,
        H.OPERATION_STATUS,
        H.DOCUMENTS_LOADED,
        H.DOCUMENTS_CHANGED,
        H.FACT_ROWS_UPDATED,
        H.FIELDS_CHANGED,
        H.DUPLICATE_IDS_FOUND,
        H.ERRORS_COUNT,
        H.EXECUTION_SECONDS,
        H.ERROR_TEXT
      ]
    },

    OPERATOR_CARD: {
      name: 'Карточка операциониста',
      headerRow: 5,
      dataStartRow: 6,
      requiredHeaders: [
        H.DOCUMENT_ID,
        H.OBJECT_ID,
        H.DOCUMENT_TYPE,
        H.CONTRACT_NUMBER,
        H.DOCUMENT_DATE,
        H.DOCUMENT_STATUS,
        H.ORIGINAL_EDO,
        H.COMMENT,
        H.DOCUMENT_HOLDER,
        H.DOCUMENT_LOCATION,
        H.TRANSFERRED_BY,
        H.PAID,
        H.DOCUMENT_AMOUNT,
        H.GU_FLAG,
        H.GU_TERMS,
        H.OBJECT_STATUS,
        H.WORK_START_DATE,
        H.WORK_END_PLAN,
        H.WORK_END_FACT,
        H.CREATED_AT,
        H.UPDATED_AT,
        H.DOCUMENT_TYPE_ID,
        H.UPDATED_BY_EMAIL,
        H.RESPONSIBLE_FOREMAN
      ]
    },

    DOCUMENTS: {
      name: 'Документы объектов',
      headerRow: 3,
      dataStartRow: 4,
      requiredHeaders: [
        H.DOCUMENT_ID,
        H.OBJECT_ID,
        H.DOCUMENT_TYPE,
        H.CONTRACT_NUMBER,
        H.DOCUMENT_DATE,
        H.DOCUMENT_STATUS,
        H.ORIGINAL_EDO,
        H.COMMENT,
        H.DOCUMENT_HOLDER,
        H.DOCUMENT_LOCATION,
        H.TRANSFERRED_BY,
        H.PAID,
        H.DOCUMENT_AMOUNT,
        H.GU_FLAG,
        H.GU_TERMS,
        H.OBJECT_STATUS,
        H.WORK_START_DATE,
        H.WORK_END_PLAN,
        H.WORK_END_FACT,
        H.CREATED_AT,
        H.UPDATED_AT,
        H.DOCUMENT_TYPE_ID,
        H.UPDATED_BY_EMAIL,
        H.RESPONSIBLE_FOREMAN,
        H.HOLDER_EMPLOYEE_ID,
        H.TRANSFERRED_BY_EMPLOYEE_ID,
        H.RESPONSIBLE_FOREMAN_ID,
        H.DOCUMENT_STATUS_CHANGED_AT,
        H.REPORTING_PERIOD,
        H.CREATION_SOURCE,
        H.RECORD_STATUS
      ]
    },

    CARD_DICTIONARY: {
      name: 'Справочник для КО',
      headerRow: 2,
      dataStartRow: 3,
      requiredHeaders: [
        H.DOCUMENT_STATUS,
        H.DOCUMENT_LOCATION,
        H.PAID,
        H.GU_FLAG,
        H.ORIGINAL_EDO,
        H.RECORD_STATUS
      ]
    },

    OBJECTS: {
      name: 'Объекты',
      headerRow: 2,
      dataStartRow: 4,
      requiredHeaders: [
        H.OBJECT_ID,
        H.OBJECT_NAME,
        H.OBJECT_ADDRESS,
        H.CONTRACT_NUMBER,
        H.RESPONSIBLE_FOREMAN_ID,
        H.RESPONSIBLE_FOREMAN,
        H.OBJECT_STATUS,
        H.WORK_START_DATE,
        H.WORK_END_PLAN,
        H.WORK_END_FACT
      ]
    },

    EMPLOYEES: {
      name: 'Справочник сотрудников',
      headerRow: 4,
      dataStartRow: 5,
      requiredHeaders: [
        H.EMPLOYEE_ID,
        H.EMPLOYEE_NAME,
        H.EMPLOYEE_POSITION,
        H.WORK_PHONE,
        H.TELEGRAM_NICK,
        H.EMAIL,
        H.EMPLOYEE_GROUP,
        H.HIRE_DATE,
        H.DISMISSAL_DATE
      ]
    },

    CLIENTS: {
      name: 'Справочник клиентов',
      headerRow: 4,
      dataStartRow: 5,
      requiredHeaders: [
        H.CLIENT_ID,
        H.CLIENT_NAME,
        H.PHONE,
        H.EMAIL,
        H.CLIENT_STATUS
      ]
    },

    DOCUMENT_TYPES: {
      name: 'Справочник документов',
      headerRow: 4,
      dataStartRow: 5,
      requiredHeaders: [
        H.DOCUMENT_TYPE_ID,
        H.DOCUMENT_TYPE,
        H.MANDATORY,
        H.CREATE_ON_OBJECT_CREATION,
        H.EXECUTION_DEADLINE,
        H.REPEATABILITY
      ]
    }
  },

  /**
   * ==========================================================
   * КАРТА ПОЛЕЙ КАРТОЧКИ ОПЕРАЦИОНИСТА
   * ==========================================================
   *
   * editable: true — оператор может изменить поле.
   * editable: false — поле предназначено только для просмотра.
   *
   * sourceSheetKey: OBJECTS означает, что актуальное значение
   * при загрузке карточки нужно брать из листа «Объекты».
   */
  CARD_FIELD_MAP: [
    {
      cardHeader: H.DOCUMENT_ID,
      factHeader: H.DOCUMENT_ID,
      editable: false
    },
    {
      cardHeader: H.OBJECT_ID,
      factHeader: H.OBJECT_ID,
      editable: false
    },
    {
      cardHeader: H.DOCUMENT_TYPE,
      factHeader: H.DOCUMENT_TYPE,
      editable: false
    },
    {
      cardHeader: H.CONTRACT_NUMBER,
      factHeader: H.CONTRACT_NUMBER,
      editable: true
    },
    {
      cardHeader: H.DOCUMENT_DATE,
      factHeader: H.DOCUMENT_DATE,
      editable: true
    },
    {
      cardHeader: H.DOCUMENT_STATUS,
      factHeader: H.DOCUMENT_STATUS,
      editable: true
    },
    {
      cardHeader: H.ORIGINAL_EDO,
      factHeader: H.ORIGINAL_EDO,
      editable: true
    },
    {
      cardHeader: H.COMMENT,
      factHeader: H.COMMENT,
      editable: true
    },
    {
      cardHeader: H.DOCUMENT_HOLDER,
      factHeader: H.DOCUMENT_HOLDER,
      editable: true
    },
    {
      cardHeader: H.DOCUMENT_LOCATION,
      factHeader: H.DOCUMENT_LOCATION,
      editable: true
    },
    {
      cardHeader: H.TRANSFERRED_BY,
      factHeader: H.TRANSFERRED_BY,
      editable: true
    },
    {
      cardHeader: H.PAID,
      factHeader: H.PAID,
      editable: true
    },
    {
      cardHeader: H.DOCUMENT_AMOUNT,
      factHeader: H.DOCUMENT_AMOUNT,
      editable: true
    },
    {
      cardHeader: H.GU_FLAG,
      factHeader: H.GU_FLAG,
      editable: true
    },
    {
      cardHeader: H.GU_TERMS,
      factHeader: H.GU_TERMS,
      editable: true
    },

    // Актуальные данные объекта подставляются из «Объекты»
    {
      cardHeader: H.OBJECT_STATUS,
      factHeader: H.OBJECT_STATUS,
      sourceSheetKey: 'OBJECTS',
      sourceHeader: H.OBJECT_STATUS,
      editable: false
    },
    {
      cardHeader: H.WORK_START_DATE,
      factHeader: H.WORK_START_DATE,
      sourceSheetKey: 'OBJECTS',
      sourceHeader: H.WORK_START_DATE,
      editable: false
    },
    {
      cardHeader: H.WORK_END_PLAN,
      factHeader: H.WORK_END_PLAN,
      sourceSheetKey: 'OBJECTS',
      sourceHeader: H.WORK_END_PLAN,
      editable: false
    },
    {
      cardHeader: H.WORK_END_FACT,
      factHeader: H.WORK_END_FACT,
      sourceSheetKey: 'OBJECTS',
      sourceHeader: H.WORK_END_FACT,
      editable: false
    },

    // Системные поля документа
    {
      cardHeader: H.CREATED_AT,
      factHeader: H.CREATED_AT,
      editable: false
    },
    {
      cardHeader: H.UPDATED_AT,
      factHeader: H.UPDATED_AT,
      editable: false
    },
    {
      cardHeader: H.DOCUMENT_TYPE_ID,
      factHeader: H.DOCUMENT_TYPE_ID,
      editable: false
    },
    {
      cardHeader: H.UPDATED_BY_EMAIL,
      factHeader: H.UPDATED_BY_EMAIL,
      editable: false
    },

    // Ответственный также загружается из «Объекты»
    {
      cardHeader: H.RESPONSIBLE_FOREMAN,
      factHeader: H.RESPONSIBLE_FOREMAN,
      sourceSheetKey: 'OBJECTS',
      sourceHeader: H.RESPONSIBLE_FOREMAN,
      editable: false
    }
  ]
};


/**
 * ============================================================
 * 3. ГЛАВНАЯ ПРОВЕРКА СТРУКТУРЫ
 * ============================================================
 */

/**
 * Запускает проверку и показывает результат пользователю.
 */
function validateSystemStructure() {
  const result = runSystemStructureValidation_();

  showSystemValidationDialog_(result);

  return result;
}


/**
 * Выполняет проверку структуры без показа окна.
 *
 * Будущие рабочие скрипты смогут вызывать эту функцию
 * перед созданием или изменением документов.
 */
function runSystemStructureValidation_() {
  const ss = getSystemSpreadsheet_();

  const result = {
    version: SYSTEM_CONFIG.VERSION,
    checkedAt: new Date(),
    ok: true,
    errors: [],
    warnings: [],
    sheets: []
  };

  validateConfiguration_(result);

  Object.keys(SYSTEM_CONFIG.SHEETS).forEach(function (sheetKey) {
    const config = SYSTEM_CONFIG.SHEETS[sheetKey];
    const sheetResult = validateConfiguredSheet_(
      ss,
      sheetKey,
      config
    );

    result.sheets.push(sheetResult);

    sheetResult.errors.forEach(function (error) {
      result.errors.push(error);
    });

    sheetResult.warnings.forEach(function (warning) {
      result.warnings.push(warning);
    });
  });

  validateCardFieldMap_(result);

  result.ok = result.errors.length === 0;

  return result;
}


/**
 * Выполняет проверку и останавливает рабочий скрипт,
 * если структура системы нарушена.
 */
function assertSystemStructure_() {
  const result = runSystemStructureValidation_();

  if (!result.ok) {
    throw new Error(
      'Структура системы содержит ошибки:\n\n' +
      result.errors.join('\n')
    );
  }

  return result;
}


/**
 * Проверяет структуру только перечисленных системных листов.
 * Используется независимыми рабочими модулями, которым не нужна
 * полная проверка всех листов системы.
 */
function assertSystemSheetsStructure_(sheetKeys) {
  if (!Array.isArray(sheetKeys) || sheetKeys.length === 0) {
    throw new TypeError(
      'assertSystemSheetsStructure_ ожидает непустой массив ключей листов.'
    );
  }

  const ss = getSystemSpreadsheet_();
  const result = {
    version: SYSTEM_CONFIG.VERSION,
    checkedAt: new Date(),
    ok: true,
    errors: [],
    warnings: [],
    sheets: []
  };
  const seenKeys = {};

  sheetKeys.forEach(function (sheetKey) {
    if (seenKeys[sheetKey]) {
      throw new Error(
        'Системный ключ листа «' + sheetKey + '» передан повторно.'
      );
    }

    seenKeys[sheetKey] = true;

    const config = SYSTEM_CONFIG.SHEETS[sheetKey];

    if (!config) {
      throw new Error(
        'Неизвестный системный ключ листа: ' + sheetKey
      );
    }

    const sheetResult = validateConfiguredSheet_(
      ss,
      sheetKey,
      config
    );

    result.sheets.push(sheetResult);
    result.errors = result.errors.concat(sheetResult.errors);
    result.warnings = result.warnings.concat(sheetResult.warnings);
  });

  result.ok = result.errors.length === 0;

  if (!result.ok) {
    throw new Error(
      'Структура используемых листов содержит ошибки:\n\n' +
      result.errors.join('\n')
    );
  }

  return result;
}


/**
 * Проверяет внутреннюю конфигурацию ядра.
 */
function validateConfiguration_(result) {
  const sheetNames = {};
  const sheetKeys = Object.keys(SYSTEM_CONFIG.SHEETS);

  sheetKeys.forEach(function (sheetKey) {
    const config = SYSTEM_CONFIG.SHEETS[sheetKey];

    if (!config.name) {
      result.errors.push(
        'Для системного ключа «' + sheetKey +
        '» не указано название листа.'
      );
      return;
    }

    if (sheetNames[config.name]) {
      result.errors.push(
        'Название листа «' + config.name +
        '» используется в конфигурации несколько раз.'
      );
    }

    sheetNames[config.name] = true;

    const normalizedRequiredHeaders =
      config.requiredHeaders.map(sysNormalizeHeader_);

    const duplicateConfigHeaders =
      findDuplicateValues_(normalizedRequiredHeaders);

    if (duplicateConfigHeaders.length > 0) {
      result.errors.push(
        'В конфигурации листа «' + config.name +
        '» повторяются обязательные заголовки: ' +
        duplicateConfigHeaders.join(', ') + '.'
      );
    }
  });
}


/**
 * Проверяет один лист.
 */
function validateConfiguredSheet_(ss, sheetKey, config) {
  const sheetResult = {
    key: sheetKey,
    name: config.name,
    errors: [],
    warnings: [],
    headers: [],
    headerMap: {}
  };

  const sheet = ss.getSheetByName(config.name);

  if (!sheet) {
    sheetResult.errors.push(
      'Не найден лист «' + config.name + '».'
    );

    return sheetResult;
  }

  if (
    !Number.isInteger(config.headerRow) ||
    config.headerRow < 1
  ) {
    sheetResult.errors.push(
      'Для листа «' + config.name +
      '» некорректно указана строка заголовков.'
    );

    return sheetResult;
  }

  if (
    !Number.isInteger(config.dataStartRow) ||
    config.dataStartRow <= config.headerRow
  ) {
    sheetResult.errors.push(
      'Для листа «' + config.name +
      '» строка начала данных должна находиться ниже ' +
      'строки заголовков.'
    );
  }

  const lastColumn = Math.max(
    sheet.getLastColumn(),
    1
  );

  const rawHeaders = sheet
    .getRange(config.headerRow, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const headers = rawHeaders.map(sysNormalizeHeader_);

  sheetResult.headers = headers;

  const headerPositions = {};

  headers.forEach(function (header, index) {
    if (!header) {
      return;
    }

    if (!headerPositions[header]) {
      headerPositions[header] = [];
    }

    headerPositions[header].push(index + 1);
  });

  Object.keys(headerPositions).forEach(function (header) {
    const columns = headerPositions[header];

    if (columns.length > 1) {
      const columnNames = columns.map(sysColumnToLetter_);

      sheetResult.errors.push(
        'На листе «' + config.name +
        '» заголовок «' + header +
        '» найден несколько раз: ' +
        columnNames.join(', ') + '.'
      );
    }
  });

  config.requiredHeaders.forEach(function (requiredHeader) {
    const normalizedRequiredHeader =
      sysNormalizeHeader_(requiredHeader);

    const columns =
      headerPositions[normalizedRequiredHeader] || [];

    if (columns.length === 0) {
      sheetResult.errors.push(
        'На листе «' + config.name +
        '» отсутствует заголовок «' +
        requiredHeader + '».'
      );
    }

    if (columns.length === 1) {
      sheetResult.headerMap[requiredHeader] = columns[0];
    }
  });

  return sheetResult;
}


/**
 * Проверяет карту карточки операциониста.
 */
function validateCardFieldMap_(result) {
  const cardResult = findSheetValidationResult_(
    result,
    'OPERATOR_CARD'
  );

  const documentsResult = findSheetValidationResult_(
    result,
    'DOCUMENTS'
  );

  if (!cardResult || !documentsResult) {
    return;
  }

  SYSTEM_CONFIG.CARD_FIELD_MAP.forEach(function (field) {
    const cardHeader = sysNormalizeHeader_(
      field.cardHeader
    );

    const factHeader = sysNormalizeHeader_(
      field.factHeader
    );

    if (!cardResult.headers.includes(cardHeader)) {
      result.errors.push(
        'В карте карточки указано поле «' +
        field.cardHeader +
        '», но такого заголовка нет на листе «' +
        SYSTEM_CONFIG.SHEETS.OPERATOR_CARD.name +
        '».'
      );
    }

    if (!documentsResult.headers.includes(factHeader)) {
      result.errors.push(
        'В карте карточки указано поле факта «' +
        field.factHeader +
        '», но такого заголовка нет на листе «' +
        SYSTEM_CONFIG.SHEETS.DOCUMENTS.name +
        '».'
      );
    }

    if (field.sourceSheetKey) {
      const sourceResult = findSheetValidationResult_(
        result,
        field.sourceSheetKey
      );

      if (!sourceResult) {
        result.errors.push(
          'В карте поля «' + field.cardHeader +
          '» указан неизвестный источник «' +
          field.sourceSheetKey + '».'
        );

        return;
      }

      const sourceHeader = sysNormalizeHeader_(
        field.sourceHeader || field.cardHeader
      );

      if (!sourceResult.headers.includes(sourceHeader)) {
        result.errors.push(
          'Поле «' + field.cardHeader +
          '» должно загружаться с листа «' +
          sourceResult.name +
          '», но в источнике отсутствует заголовок «' +
          (field.sourceHeader || field.cardHeader) +
          '».'
        );
      }
    }
  });
}


/**
 * ============================================================
 * 4. ДОСТУП К ЛИСТАМ И СТОЛБЦАМ
 * ============================================================
 */

/**
 * Возвращает текущую таблицу.
 */
function getSystemSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Не удалось определить активную Google-таблицу. ' +
      'Проект Apps Script должен быть привязан к таблице.'
    );
  }

  return ss;
}


/**
 * Возвращает готовый контекст листа:
 * - лист;
 * - конфигурацию;
 * - массив заголовков;
 * - карту «заголовок → номер столбца».
 */
function getSystemSheetContext_(sheetKey) {
  const config = SYSTEM_CONFIG.SHEETS[sheetKey];

  if (!config) {
    throw new Error(
      'Неизвестный системный ключ листа: ' + sheetKey
    );
  }

  const ss = getSystemSpreadsheet_();
  const sheet = ss.getSheetByName(config.name);

  if (!sheet) {
    throw new Error(
      'Не найден обязательный лист «' +
      config.name + '».'
    );
  }

  const lastColumn = Math.max(
    sheet.getLastColumn(),
    1
  );

  const headers = sheet
    .getRange(config.headerRow, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(sysNormalizeHeader_);

  const headerMap = {};
  const duplicatePositions = {};

  headers.forEach(function (header, index) {
    if (!header) {
      return;
    }

    if (Object.prototype.hasOwnProperty.call(
      headerMap,
      header
    )) {
      if (!duplicatePositions[header]) {
        duplicatePositions[header] = [
          headerMap[header]
        ];
      }

      duplicatePositions[header].push(index + 1);
      return;
    }

    headerMap[header] = index + 1;
  });

  const duplicateHeaders =
    Object.keys(duplicatePositions);

  if (duplicateHeaders.length > 0) {
    throw new Error(
      'На листе «' + config.name +
      '» найдены повторяющиеся заголовки: ' +
      duplicateHeaders.join(', ') + '.'
    );
  }

  const missingHeaders = config.requiredHeaders.filter(
    function (requiredHeader) {
      const normalizedHeader =
        sysNormalizeHeader_(requiredHeader);

      return !Object.prototype.hasOwnProperty.call(
        headerMap,
        normalizedHeader
      );
    }
  );

  if (missingHeaders.length > 0) {
    throw new Error(
      'На листе «' + config.name +
      '» отсутствуют заголовки: ' +
      missingHeaders.join(', ') + '.'
    );
  }

  return {
    key: sheetKey,
    sheet: sheet,
    config: config,
    headers: headers,
    headerMap: headerMap
  };
}


/**
 * Получает номер столбца по системному ключу листа
 * и названию заголовка.
 */
function getSystemColumn_(sheetKey, headerName) {
  const context = getSystemSheetContext_(sheetKey);
  const normalizedHeader =
    sysNormalizeHeader_(headerName);

  const column =
    context.headerMap[normalizedHeader];

  if (!column) {
    throw new Error(
      'На листе «' + context.config.name +
      '» не найден столбец «' +
      headerName + '».'
    );
  }

  return column;
}


/**
 * Возвращает массив редактируемых полей карточки.
 */
function getEditableCardFields_() {
  return SYSTEM_CONFIG.CARD_FIELD_MAP.filter(
    function (field) {
      return field.editable === true;
    }
  );
}


/**
 * ============================================================
 * 5. EMAIL ПОЛЬЗОВАТЕЛЯ
 * ============================================================
 */

/**
 * Возвращает email пользователя, запустившего операцию.
 *
 * Не подставляем email владельца скрипта автоматически,
 * чтобы не записать неверного автора изменения.
 */
function getActiveUserEmail_() {
  const email = String(
    Session.getActiveUser().getEmail() || ''
  ).trim();

  return email || 'EMAIL_NOT_AVAILABLE';
}


/**
 * Резервирует следующий читаемый ID операции.
 * Вызывать только внутри withDocumentLock_: истории и служебный максимум
 * вместе не позволяют повторно использовать номер после удаления строк.
 */
function generateOperationId_(at) {
  const date = at instanceof Date ? at : new Date();
  const timezone = getSystemSpreadsheet_().getSpreadsheetTimeZone();
  const day = Utilities.formatDate(date, timezone, 'yyyyMMdd');
  const pattern = new RegExp('^OP-' + day + '-(\\d{4})$');
  let maximum = 0;

  ['OPERATION_HISTORY', 'CHANGE_HISTORY'].forEach(function (sheetKey) {
    const context = getSystemSheetContext_(sheetKey);
    const column = getSystemColumn_(sheetKey, H.OPERATION_ID);
    const lastRow = context.sheet.getLastRow();
    if (lastRow < context.config.dataStartRow) {
      return;
    }
    const ids = context.sheet.getRange(
      context.config.dataStartRow,
      column,
      lastRow - context.config.dataStartRow + 1,
      1
    ).getValues();
    ids.forEach(function (row) {
      const match = String(row[0] == null ? '' : row[0]).trim().match(pattern);
      if (match) {
        maximum = Math.max(maximum, Number(match[1]));
      }
    });
  });

  const properties = PropertiesService.getDocumentProperties();
  const propertyKey = 'SYSTEM_OPERATION_SEQUENCE_' + day;
  maximum = Math.max(maximum, Number(properties.getProperty(propertyKey)) || 0);
  if (maximum >= 9999) {
    throw new Error(
      'Исчерпан предел 9999 операций за ' + day + '. Запись данных отменена.'
    );
  }
  const next = maximum + 1;
  properties.setProperty(propertyKey, String(next));
  return SYSTEM_CONFIG.ID_PREFIXES.OPERATION + day + '-' +
    String(next).padStart(4, '0');
}


/** Формирует читаемый ID изменения, связанный с операцией. */
function generateChangeId_(operationId, sequence) {
  const match = String(operationId || '').match(/^OP-(\d{8})-(\d{4})$/);
  if (!match) {
    throw new Error('Невозможно создать ID изменения: некорректный ID операции.');
  }
  if (!Number.isInteger(sequence) || sequence < 1 || sequence > 9999) {
    throw new Error(
      'Исчерпан предел 9999 изменений внутри операции ' + operationId +
      '. Запись данных отменена.'
    );
  }
  return SYSTEM_CONFIG.ID_PREFIXES.CHANGE + match[1] + '-' + match[2] + '-' +
    String(sequence).padStart(4, '0');
}


/**
 * Ручная проверка получения email.
 */
function testActiveUserEmail() {
  const email = getActiveUserEmail_();

  const message =
    email === 'EMAIL_NOT_AVAILABLE'
      ? (
        'Apps Script не вернул email текущего пользователя.\n\n' +
        'Нужно проверить эту функцию на аккаунте операциониста.'
      )
      : (
        'Email пользователя успешно определён:\n\n' +
        email
      );

  SpreadsheetApp.getUi().alert(
    'Проверка email',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return email;
}


/**
 * ============================================================
 * 6. БЛОКИРОВКА ОДНОВРЕМЕННОЙ ЗАПИСИ
 * ============================================================
 */

/**
 * Выполняет функцию под общей блокировкой.
 *
 * Если проект привязан к таблице, используется DocumentLock.
 * В остальных случаях используется ScriptLock.
 */
function withDocumentLock_(callback) {
  if (typeof callback !== 'function') {
    throw new TypeError(
      'withDocumentLock_ ожидает функцию.'
    );
  }

  const lock =
    LockService.getDocumentLock() ||
    LockService.getScriptLock();

  const locked = lock.tryLock(
    SYSTEM_CONFIG.LOCK_WAIT_MS
  );

  if (!locked) {
    throw new Error(
      'Система сейчас обрабатывает изменения другого ' +
      'пользователя. Повторите попытку через несколько секунд.'
    );
  }

  try {
    const result = callback();

    /**
     * Принудительно отправляем накопленные изменения
     * до освобождения блокировки.
     */
    SpreadsheetApp.flush();

    return result;
  } finally {
    lock.releaseLock();
  }
}


/**
 * ============================================================
 * 7. СЛУЖЕБНЫЕ ФУНКЦИИ
 * ============================================================
 */

/**
 * Находит результат проверки конкретного листа.
 */
function findSheetValidationResult_(result, sheetKey) {
  return result.sheets.find(function (item) {
    return item.key === sheetKey;
  });
}


/**
 * Находит повторяющиеся значения массива.
 */
function findDuplicateValues_(values) {
  const seen = {};
  const duplicates = {};

  values.forEach(function (value) {
    if (!value) {
      return;
    }

    if (seen[value]) {
      duplicates[value] = true;
    }

    seen[value] = true;
  });

  return Object.keys(duplicates);
}


/**
 * Нормализует заголовок:
 * - заменяет неразрывные пробелы;
 * - убирает повторяющиеся пробелы;
 * - убирает пробелы по краям.
 *
 * Регистр букв сохраняется.
 */
function sysNormalizeHeader_(value) {
  return String(
    value == null ? '' : value
  )
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


/**
 * Преобразует номер столбца в буквенное обозначение.
 *
 * 1  → A
 * 2  → B
 * 27 → AA
 */
function sysColumnToLetter_(columnNumber) {
  let column = Number(columnNumber);
  let result = '';

  while (column > 0) {
    const remainder = (column - 1) % 26;

    result =
      String.fromCharCode(65 + remainder) +
      result;

    column = Math.floor(
      (column - 1) / 26
    );
  }

  return result;
}


/**
 * Экранирует текст перед выводом в HTML.
 */
function sysEscapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}


/**
 * ============================================================
 * 8. ОКНО С РЕЗУЛЬТАТАМИ ПРОВЕРКИ
 * ============================================================
 */

function showSystemValidationDialog_(result) {
  const lines = [];

  lines.push(
    result.ok
      ? 'СТРУКТУРА СИСТЕМЫ КОРРЕКТНА'
      : 'НАЙДЕНЫ ОШИБКИ В СТРУКТУРЕ'
  );

  lines.push('');
  lines.push('Версия ядра: ' + result.version);
  lines.push('Проверено листов: ' + result.sheets.length);
  lines.push('Ошибок: ' + result.errors.length);
  lines.push(
    'Предупреждений: ' +
    result.warnings.length
  );

  lines.push('');
  lines.push('ПРОВЕРЕННЫЕ ЛИСТЫ:');

  result.sheets.forEach(function (sheetResult) {
    const marker =
      sheetResult.errors.length === 0
        ? 'OK'
        : 'ОШИБКА';

    lines.push(
      marker + ' — ' + sheetResult.name
    );
  });

  if (result.errors.length > 0) {
    lines.push('');
    lines.push('ОШИБКИ:');

    result.errors.forEach(function (error, index) {
      lines.push(
        (index + 1) + '. ' + error
      );
    });
  }

  if (result.warnings.length > 0) {
    lines.push('');
    lines.push('ПРЕДУПРЕЖДЕНИЯ:');

    result.warnings.forEach(
      function (warning, index) {
        lines.push(
          (index + 1) + '. ' + warning
        );
      }
    );
  }

  if (result.ok) {
    lines.push('');
    lines.push(
      'Можно переходить к созданию документов по объекту.'
    );
  }

  const escapedReport = sysEscapeHtml_(
    lines.join('\n')
  );

  const html = HtmlService
    .createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<base target="_top">' +
      '<style>' +
      'body {' +
      '  font-family: Arial, sans-serif;' +
      '  padding: 18px;' +
      '  color: #202124;' +
      '}' +
      'pre {' +
      '  white-space: pre-wrap;' +
      '  word-break: break-word;' +
      '  font-family: Arial, sans-serif;' +
      '  line-height: 1.5;' +
      '  font-size: 13px;' +
      '}' +
      '</style>' +
      '</head>' +
      '<body>' +
      '<pre>' + escapedReport + '</pre>' +
      '</body>' +
      '</html>'
    )
    .setWidth(760)
    .setHeight(560);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Проверка структуры системы'
  );
}

/** Показывает критическую ошибку в заметном безопасном HTML-диалоге. */
function showCriticalOperationError_(title, summary, reason) {
  const safeTitle = sysEscapeHtml_(String(title || 'Ошибка'));
  const safeSummary = sysEscapeHtml_(String(summary || ''));
  const safeReason = sysEscapeHtml_(String(reason || 'Неизвестная ошибка'));
  const html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_top"><style>' +
    'body{font:14px Arial,sans-serif;color:#202124;padding:20px}' +
    'pre{white-space:pre-wrap;font:14px Arial,sans-serif;line-height:1.45}' +
    '.critical{border-top:2px solid #8b0000;margin-top:18px;padding-top:16px}' +
    '.critical h2{color:#8b0000;font-size:18px;margin:0 0 10px;font-weight:700}' +
    '.reason{font-weight:700;white-space:pre-wrap}' +
    'button{margin-top:22px;padding:8px 20px}</style></head><body>' +
    '<pre>' + safeSummary + '</pre><div class="critical">' +
    '<h2>КРИТИЧЕСКАЯ ОШИБКА</h2><div class="reason">' + safeReason + '</div>' +
    '</div><button onclick="google.script.host.close()">Закрыть</button>' +
    '</body></html>'
  ).setWidth(720).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, safeTitle);
}

/**
 * Необязательная идемпотентная настройка двух системных ячеек справочника.
 * Защита работает только в warning-only режиме и не ограничивает редакторов.
 */
function setupSystemDictionaryGuards() {
  assertSystemSheetsStructure_(['CARD_DICTIONARY']);
  const context = getSystemSheetContext_('CARD_DICTIONARY');
  const guards = [
    {
      header: H.DOCUMENT_STATUS,
      value: SYSTEM_CONFIG.VALUES.INITIAL_DOCUMENT_STATUS
    },
    {
      header: H.RECORD_STATUS,
      value: SYSTEM_CONFIG.VALUES.ACTIVE_RECORD_STATUS
    }
  ];
  const lastRow = context.sheet.getLastRow();
  if (lastRow < context.config.dataStartRow) {
    throw new Error('В справочнике отсутствуют обязательные системные значения.');
  }
  guards.forEach(function (guard) {
    const column = getSystemColumn_('CARD_DICTIONARY', guard.header);
    const range = context.sheet.getRange(
      context.config.dataStartRow,
      column,
      lastRow - context.config.dataStartRow + 1,
      1
    );
    const rows = range.getValues();
    let found = false;
    rows.forEach(function (row, offset) {
      if (String(row[0] == null ? '' : row[0]).trim() !== guard.value) {
        return;
      }
      found = true;
      const cell = context.sheet.getRange(context.config.dataStartRow + offset, column);
      cell.setNote(
        'Системно обязательное значение. Не удаляйте и не переименовывайте его.'
      ).setBackground('#fce8e6').setFontColor('#8b0000').setFontWeight('bold');
      const existing = cell.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      const protection = existing.length > 0
        ? existing[0]
        : cell.protect().setDescription('Системное значение документооборота');
      protection.setWarningOnly(true);
    });
    if (!found) {
      throw new Error(
        'Не найдено системное значение «' + guard.value + '» в поле «' +
        guard.header + '».'
      );
    }
  });
}
