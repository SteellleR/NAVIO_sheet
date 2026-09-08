/**
 * Utils.gs — общие чистые функции и поиск структуры листов.
 *
 * Нормализует текст, формулы, даты и номера ВАТС, находит столбцы по заголовкам,
 * проверяет диапазоны и строит стабильные хеши справочников. Здесь нет записи
 * бизнес-данных: функции переиспользуются всеми профильными модулями.
 */
function normalizeText_(value) {
  return String(value == null ? "" : value)
    .replace(/\u00a0/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}


/** Нормализует формулу, не меняя формулу в самой таблице. */
function normalizeFormula_(formula) {
  return String(formula == null ? "" : formula)
    .replace(/\s+/g, "")
    .toLowerCase();
}


/** Возвращает true только для действительно пустого значения ячейки. */
function isBlankCellValue_(value) {
  return value === "" || value === null || value === undefined;
}


/** Возвращает связанную таблицу и поддерживает запуск по таймеру без UI. */
function getProjectSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(
    APP_CONFIG_.properties.spreadsheetId
  );
  if (!spreadsheetId) {
    throw new Error(
      "Не сохранён ID связанной таблицы. Запустите installProject()."
    );
  }

  return SpreadsheetApp.openById(spreadsheetId);
}


/** Проверяет, пересекает ли изменённый диапазон указанный столбец. */
function rangeTouchesColumn_(range, column) {
  return (
    column >= range.getColumn() &&
    column <= range.getLastColumn()
  );
}


/** Преобразует номер столбца в буквенное обозначение A1. */
function columnToLetter_(column) {
  let value = Number(column);
  let result = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }

  return result;
}


/**
 * Ищет ровно один столбец по одному или нескольким допустимым заголовкам.
 * Пустой результат допустим только при optional=true.
 */
function findUniqueHeaderColumn_(sheet, expectedHeaders, optional) {
  return findHeaderColumns_(sheet, {
    result: {headers: expectedHeaders, optional: Boolean(optional)}
  }).result;
}


/** Ищет несколько заголовков за одно чтение первой строки листа. */
function findHeaderColumns_(sheet, definitions) {
  // getLastColumn() иногда падает на новых нативных таблицах Google с
  // типизированными столбцами. getMaxColumns() не запрашивает их числовой
  // формат и безопасен: строка заголовков всё равно читается только один раз.
  const columnCount = sheet.getMaxColumns();
  if (columnCount < 1) {
    throw new Error("На листе «" + sheet.getName() + "» нет заголовков.");
  }

  const headers = sheet
    .getRange(APP_CONFIG_.headerRow, 1, 1, columnCount)
    .getDisplayValues()[0]
    .map(normalizeText_);
  const result = {};

  Object.keys(definitions).forEach(function(key) {
    const definition = definitions[key];
    const expectedHeaders = definition.headers;
    const candidates = (Array.isArray(expectedHeaders)
      ? expectedHeaders
      : [expectedHeaders]
    ).map(normalizeText_);
    const matches = [];

    headers.forEach(function(header, index) {
      if (candidates.includes(header)) matches.push(index + 1);
    });

    if (matches.length > 1) {
      throw new Error(
        "На листе «" + sheet.getName() + "» найдено несколько столбцов «" +
        expectedHeaders + "»: " + matches.join(", ") + "."
      );
    }
    if (!matches.length && !definition.optional) {
      throw new Error(
        "На листе «" + sheet.getName() + "» не найден столбец «" +
        expectedHeaders + "»."
      );
    }

    result[key] = matches[0] || 0;
  });

  return result;
}


/** Находит критичные столбцы рабочего листа по заголовкам. */
function getRequiredColumns_(sheet) {
  try {
    return findHeaderColumns_(sheet, {
      date: {headers: APP_CONFIG_.headers.date},
      vehicle: {headers: APP_CONFIG_.headers.vehicle},
      teleoperator: {headers: APP_CONFIG_.headers.teleoperator},
      logviewer: {headers: APP_CONFIG_.headers.logviewer},
      logCheckDate: {headers: APP_CONFIG_.headers.logCheckDate}
    });
  } catch (error) {
    console.warn(error.message);
    return null;
  }
}


/** Возвращает стабильный SHA-256-хеш списка строк. */
function buildHash_(values) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    values.join("\n"),
    Utilities.Charset.UTF_8
  );

  return bytes.map(function(value) {
    return (value + 256).toString(16).slice(-2);
  }).join("");
}


/** Нормализует идентификатор ВАТС для URL Logviewer. */
function normalizeVehicle_(value) {
  const text = String(value == null ? "" : value)
    .trim()
    .toLowerCase()
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, "-")
    .replace(/_/g, "-")
    .replace(/\s+/g, "")
    .replace(/к/g, "k")
    .replace(/с/g, "c");

  if (!text) return "";

  let match = text.match(/^kc2-?(\d{1,3})$/);
  if (match) return "kc2-" + match[1].padStart(3, "0");

  match = text.match(/^sh-?(\d{1,3})$/);
  if (match) return "sh-" + match[1].padStart(3, "0");

  console.warn("Неизвестный формат ВАТС для Logviewer: " + value);
  return text;
}


/** Преобразует значение даты в ключ yyyy-MM-dd. */
function dateKey_(value, timeZone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timeZone, "yyyy-MM-dd");
  }

  // Google Sheets хранит даты как количество суток с 30.12.1899. Если у
  // ячейки слетел числовой формат, getValues() возвращает именно это число.
  // Распознаём его, чтобы старые значения можно было безопасно восстановить.
  if (typeof value === "number" && isFinite(value) && value >= 1) {
    const millisecondsPerDay = 24 * 60 * 60 * 1000;
    const sheetsEpoch = Date.UTC(1899, 11, 30);
    const date = new Date(sheetsEpoch + Math.floor(value) * millisecondsPerDay);
    return Utilities.formatDate(date, "UTC", "yyyy-MM-dd");
  }

  const text = String(value == null ? "" : value).trim();
  let match = text.match(
    /^(\d{1,2})\.(\d{1,2})(?:\.(\d{2}|\d{4}))?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/
  );

  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    let year = match[3] || Utilities.formatDate(new Date(), timeZone, "yyyy");
    if (year.length === 2) year = "20" + year;

    return isValidDateParts_(Number(year), month, day)
      ? String(year) + "-" + String(month).padStart(2, "0") + "-" +
        String(day).padStart(2, "0")
      : "";
  }

  match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return "";

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  return isValidDateParts_(year, month, day)
    ? String(year) + "-" + String(month).padStart(2, "0") + "-" +
      String(day).padStart(2, "0")
    : "";
}


/** Проверяет существование календарной даты независимо от локальной TZ среды. */
function isValidDateParts_(year, month, day) {
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}
