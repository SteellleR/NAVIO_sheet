/**
 * Violators.gs - быстрый обратимый отчёт по ошибкам проверки логов.
 *
 * Обычный onEdit пересчитывает только фактически затронутые строки и обновляет
 * их адреса в готовом отчёте. Контрольный триггер раз в 10 минут выполняет
 * полный проход по границам нативных таблиц, а не по всему размеру листов.
 * Каждый проход пишет в журнал режим, длительность и число просмотренных строк.
 */
function syncViolators(manageNativeTable) {
  const startedAt = Date.now();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    return {
      updated: false,
      reason: "locked",
      performance: finishViolatorsPerformance_("full", startedAt, {
        rowsScanned: 0,
        sheetsScanned: 0
      })
    };
  }

  try {
    const spreadsheet = getProjectSpreadsheet_();
    const entries = getTeleoperatorEntries_(spreadsheet);
    const sheet = spreadsheet.getSheetByName(APP_CONFIG_.violators.sheet);
    if (!sheet) {
      throw new Error(
        "Не найден лист «" + APP_CONFIG_.violators.sheet + "»."
      );
    }

    validateViolatorsHeaders_(sheet);

    // Полный авторизованный проход получает метаданные один раз: они нужны и
    // для точных границ рабочих таблиц, и для расширения «Нарушителей».
    const metadata = manageNativeTable === false
      ? null
      : fetchSpreadsheetTableMetadata_(spreadsheet.getId());
    const layouts = metadata
      ? buildWorkingTableLayouts_(metadata)
      : null;
    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    const todayKey = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    const snapshot = buildViolatorsSnapshot_(
      spreadsheet,
      entries,
      todayKey,
      timeZone,
      layouts
    );
    let dataRows = Math.max(
      entries.length,
      sheet.getLastRow() - APP_CONFIG_.violators.firstDataRow + 1
    );

    if (manageNativeTable !== false) {
      const table = ensureViolatorsTableCapacity_(
        spreadsheet,
        sheet,
        entries.length,
        metadata
      );
      dataRows = table.range.endRowIndex -
        (APP_CONFIG_.violators.firstDataRow - 1);
    }

    const values = Array.from({length: dataRows}, function() {
      return ["", "", ""];
    });
    snapshot.rows.forEach(function(row, index) {
      values[index] = row;
    });

    sheet.getRange(
      APP_CONFIG_.violators.firstDataRow,
      1,
      dataRows,
      APP_CONFIG_.violators.columns
    ).setValues(values);

    const performance = finishViolatorsPerformance_("full", startedAt, {
      rowsScanned: snapshot.rowsScanned,
      sheetsScanned: snapshot.sheetsScanned,
      operators: entries.length,
      staleRows: snapshot.staleRows
    });
    PropertiesService.getDocumentProperties().setProperty(
      APP_CONFIG_.properties.violatorsLastFullSync,
      JSON.stringify(performance)
    );

    return {
      updated: true,
      operators: entries.length,
      staleRows: snapshot.staleRows,
      performance: performance
    };
  } finally {
    lock.releaseLock();
  }
}


/** Контрольный полный проход, запускаемый таймером раз в 10 минут. */
function handleViolatorsPeriodicSync() {
  return syncViolators();
}


/** Совместимость со старым названием ручной и триггерной функции. */
function handleViolatorsDailySync() {
  return syncViolators();
}


/** Возвращает метрики последнего успешного полного прохода. */
function getLastViolatorsFullSyncMetrics() {
  const value = PropertiesService.getDocumentProperties().getProperty(
    APP_CONFIG_.properties.violatorsLastFullSync
  );
  return value ? JSON.parse(value) : null;
}


/** Пересчитывает только строки, которых коснулось ручное изменение. */
function handleViolatorsEdit_(event) {
  if (!event || !event.range) return false;

  const sheet = event.range.getSheet();
  if (!isWorkingSheetTracked_(sheet.getName())) return false;
  if (event.range.getLastRow() < APP_CONFIG_.firstDataRow) return false;

  const columns = getRequiredColumns_(sheet);
  if (!columns) return false;
  const relevantColumns = [
    columns.date,
    columns.vehicle,
    columns.teleoperator,
    columns.logviewer,
    columns.logCheckDate
  ];
  if (!relevantColumns.some(function(column) {
    return rangeTouchesColumn_(event.range, column);
  })) {
    return false;
  }

  const startRow = Math.max(APP_CONFIG_.firstDataRow, event.range.getRow());
  const endRow = event.range.getLastRow();
  return syncViolatorRows_(sheet, columns, startRow, endRow);
}


/**
 * Удаляет старое состояние затронутых адресов из B:C и добавляет новое.
 * Поэтому смена логина, исправление ошибки и многострочная вставка обратимы без
 * повторного чтения тысяч незатронутых рабочих строк.
 */
function syncViolatorRows_(sheet, columns, startRow, endRow) {
  const startedAt = Date.now();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return {
      updated: false,
      reason: "locked",
      performance: finishViolatorsPerformance_("incremental", startedAt, {
        rowsScanned: 0,
        sheetsScanned: 0
      })
    };
  }

  try {
    const spreadsheet = sheet.getParent();
    const entries = getTeleoperatorEntries_(spreadsheet);
    const reportSheet = spreadsheet.getSheetByName(
      APP_CONFIG_.violators.sheet
    );
    if (!reportSheet) {
      throw new Error(
        "Не найден лист «" + APP_CONFIG_.violators.sheet + "»."
      );
    }

    const reportValues = reportSheet.getRange(
      APP_CONFIG_.violators.headerRow,
      1,
      entries.length + 1,
      APP_CONFIG_.violators.columns
    ).getDisplayValues();
    validateViolatorsHeaderValues_(reportSheet, reportValues[0]);

    const states = buildCurrentViolatorStates_(entries, reportValues.slice(1));
    if (!states) {
      console.warn(
        "Инкрементальный пересчёт пропущен: список «Нарушителей» не " +
        "синхронизирован со справочником. Таймер восстановит его полностью."
      );
      return {
        updated: false,
        reason: "report_out_of_sync",
        performance: finishViolatorsPerformance_("incremental", startedAt, {
          rowsScanned: 0,
          sheetsScanned: 0,
          operators: entries.length
        })
      };
    }

    const workingRows = readWorkingRowsForViolators_(
      sheet,
      columns,
      startRow,
      endRow
    );
    const touchedAddresses = Object.create(null);
    workingRows.forEach(function(row) {
      touchedAddresses[formatViolationAddress_(sheet.getName(), row.row)] = true;
    });

    states.forEach(function(state) {
      state.stale = state.stale.filter(function(address) {
        return !touchedAddresses[address];
      });
    });

    const byLogin = Object.create(null);
    states.forEach(function(state) {
      byLogin[normalizeText_(state.login)] = state;
    });

    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    const todayKey = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    workingRows.forEach(function(row) {
      const state = byLogin[normalizeText_(row.teleoperator)];
      if (!state) return;

      const kind = classifyLogDateViolation_(
        row.eventDate,
        row.status,
        todayKey,
        timeZone
      );
      if (!kind) return;

      state.stale.push(formatViolationAddress_(sheet.getName(), row.row));
    });

    const nextValues = states.map(function(state) {
      state.stale = sortViolationAddresses_(state.stale);
      return [
        state.stale.length ? APP_CONFIG_.violators.marker : "",
        formatViolationAddresses_(state.stale)
      ];
    });
    const currentValues = reportValues.slice(1).map(function(row) {
      return row.slice(1, APP_CONFIG_.violators.columns);
    });
    const changed = JSON.stringify(nextValues) !== JSON.stringify(currentValues);

    if (changed) {
      reportSheet.getRange(
        APP_CONFIG_.violators.firstDataRow,
        2,
        entries.length,
        APP_CONFIG_.violators.columns - 1
      ).setValues(nextValues);
    }

    return {
      updated: changed,
      performance: finishViolatorsPerformance_("incremental", startedAt, {
        rowsScanned: workingRows.length,
        sheetsScanned: 1,
        operators: entries.length
      })
    };
  } finally {
    lock.releaseLock();
  }
}


/** Строит текущее состояние B:C; возвращает null при рассинхронизации A. */
function buildCurrentViolatorStates_(entries, reportRows) {
  const states = [];

  for (let index = 0; index < entries.length; index++) {
    const row = reportRows[index] || [];
    if (normalizeText_(row[0]) !== normalizeText_(entries[index].login)) {
      return null;
    }

    states.push({
      login: entries[index].login,
      stale: parseViolationAddresses_(row[2])
    });
  }

  return states;
}


/** Строит полный снимок A:C в порядке справочника телеоператоров. */
function buildViolatorsSnapshot_(
  spreadsheet,
  entries,
  todayKey,
  timeZone,
  layouts
) {
  const byLogin = Object.create(null);
  let staleRows = 0;
  let rowsScanned = 0;
  let sheetsScanned = 0;

  entries.forEach(function(entry) {
    byLogin[normalizeText_(entry.login)] = {
      login: entry.login,
      stale: []
    };
  });

  getActiveWorkingSheets_().forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    const layout = layouts ? layouts[sheetName] : null;
    const columns = layout ? layout.columns : getRequiredColumns_(sheet);
    if (!columns) return;

    const lastRow = layout ? layout.lastDataRow : sheet.getLastRow();
    if (lastRow < APP_CONFIG_.firstDataRow) return;

    const rows = readWorkingRowsForViolators_(
      sheet,
      columns,
      APP_CONFIG_.firstDataRow,
      lastRow
    );
    rowsScanned += rows.length;
    sheetsScanned++;

    rows.forEach(function(row) {
      const state = byLogin[normalizeText_(row.teleoperator)];
      if (!state) return;

      const kind = classifyLogDateViolation_(
        row.eventDate,
        row.status,
        todayKey,
        timeZone
      );
      if (!kind) return;

      const address = formatViolationAddress_(sheetName, row.row);
      state.stale.push(address);
      staleRows++;
    });
  });

  return {
    staleRows: staleRows,
    rowsScanned: rowsScanned,
    sheetsScanned: sheetsScanned,
    rows: entries.map(function(entry) {
      const state = byLogin[normalizeText_(entry.login)];
      return [
        entry.login,
        state.stale.length ? APP_CONFIG_.violators.marker : "",
        formatViolationAddresses_(state.stale)
      ];
    })
  };
}


/** Читает нужные поля любого числа строк одним обращением к SpreadsheetApp. */
function readWorkingRowsForViolators_(sheet, columns, startRow, endRow) {
  if (endRow < startRow) return [];

  const requiredColumns = [
    columns.date,
    columns.teleoperator,
    columns.logCheckDate
  ];
  const firstColumn = Math.min.apply(null, requiredColumns);
  const lastColumn = Math.max.apply(null, requiredColumns);
  const values = sheet.getRange(
    startRow,
    firstColumn,
    endRow - startRow + 1,
    lastColumn - firstColumn + 1
  ).getDisplayValues();

  return values.map(function(row, offset) {
    return {
      row: startRow + offset,
      eventDate: row[columns.date - firstColumn],
      teleoperator: row[columns.teleoperator - firstColumn],
      status: row[columns.logCheckDate - firstColumn]
    };
  });
}


/**
 * Находит рабочую нативную таблицу по обязательным заголовкам и возвращает её
 * фактическую последнюю строку. Пустой хвост листа больше не сканируется.
 */
function buildWorkingTableLayouts_(metadata) {
  const layouts = Object.create(null);
  const requiredHeaders = {
    date: APP_CONFIG_.headers.date,
    vehicle: APP_CONFIG_.headers.vehicle,
    teleoperator: APP_CONFIG_.headers.teleoperator,
    logviewer: APP_CONFIG_.headers.logviewer,
    logCheckDate: APP_CONFIG_.headers.logCheckDate
  };

  getActiveWorkingSheets_().forEach(function(sheetName) {
    const matchingSheets = (metadata.sheets || []).filter(function(item) {
      return item.properties && item.properties.title === sheetName;
    });
    if (matchingSheets.length !== 1) {
      throw new Error(
        "Для листа «" + sheetName + "» найдено листов в API: " +
        matchingSheets.length + "."
      );
    }

    const matches = [];
    (matchingSheets[0].tables || []).forEach(function(table) {
      const columns = Object.create(null);
      const startColumn = (table.range || {}).startColumnIndex || 0;

      (table.columnProperties || []).forEach(function(column, index) {
        Object.keys(requiredHeaders).forEach(function(key) {
          if (
            normalizeText_(column.columnName) ===
            normalizeText_(requiredHeaders[key])
          ) {
            columns[key] = startColumn + index + 1;
          }
        });
      });

      if (Object.keys(requiredHeaders).every(function(key) {
        return Boolean(columns[key]);
      })) {
        matches.push({table: table, columns: columns});
      }
    });

    if (matches.length !== 1) {
      throw new Error(
        "На листе «" + sheetName + "» найдено рабочих нативных таблиц: " +
        matches.length + "."
      );
    }

    const match = matches[0];
    const range = match.table.range || {};
    const startRowIndex = range.startRowIndex || 0;
    if (startRowIndex !== APP_CONFIG_.headerRow - 1) {
      throw new Error(
        "Рабочая таблица листа «" + sheetName +
        "» должна начинаться со строки заголовков 1."
      );
    }

    layouts[sheetName] = {
      columns: match.columns,
      lastDataRow: range.endRowIndex
    };
  });

  return layouts;
}


/** Возвращает stale для просроченной контрольной даты или пустую строку. */
function classifyLogDateViolation_(
  eventDate,
  status,
  todayKey,
  timeZone
) {
  const eventKey = dateKey_(eventDate, timeZone);
  const eventAge = calendarDaysBetween_(eventKey, todayKey);

  if (
    !Number.isFinite(eventAge) ||
    eventAge < 0 ||
    eventAge > APP_CONFIG_.violators.activeWindowDays
  ) {
    return "";
  }

  if (
    normalizeText_(status) ===
    normalizeText_(APP_CONFIG_.logviewer.filledLabel)
  ) {
    return "";
  }

  const statusKey = dateKey_(status, timeZone);
  if (statusKey) {
    const statusAge = calendarDaysBetween_(statusKey, todayKey);
    return Number.isFinite(statusAge) &&
      statusAge >= APP_CONFIG_.violators.staleAfterDays
      ? "stale"
      : "";
  }

  return "";
}


/** Формирует читаемый адрес одной ошибочной строки. */
function formatViolationAddress_(sheetName, row) {
  return sheetName + " / " + row;
}


/** Разбирает сохранённую строку адресов и убирает пустоты и дубли. */
function parseViolationAddresses_(value) {
  const seen = Object.create(null);
  return String(value == null ? "" : value)
    .split(";")
    .map(function(address) {
      return address.trim();
    })
    .filter(function(address) {
      if (!address || seen[address]) return false;
      seen[address] = true;
      return true;
    });
}


/** Сортирует адреса так же, как их создаёт полный проход. */
function sortViolationAddresses_(addresses) {
  const sheetOrder = Object.create(null);
  APP_CONFIG_.workingSheets.forEach(function(sheetName, index) {
    sheetOrder[normalizeText_(sheetName)] = index;
  });

  return parseViolationAddresses_(addresses.join("; ")).sort(function(a, b) {
    const left = parseViolationAddress_(a, sheetOrder);
    const right = parseViolationAddress_(b, sheetOrder);
    if (left.sheetOrder !== right.sheetOrder) {
      return left.sheetOrder - right.sheetOrder;
    }
    if (left.row !== right.row) return left.row - right.row;
    return a.localeCompare(b);
  });
}


function parseViolationAddress_(address, sheetOrder) {
  const match = String(address).match(/^(.*?)\s*\/\s*(\d+)$/);
  if (!match) {
    return {sheetOrder: APP_CONFIG_.workingSheets.length, row: Infinity};
  }

  const normalizedSheet = normalizeText_(match[1]);
  return {
    sheetOrder: Object.prototype.hasOwnProperty.call(
      sheetOrder,
      normalizedSheet
    ) ? sheetOrder[normalizedSheet] : APP_CONFIG_.workingSheets.length,
    row: Number(match[2])
  };
}


/** Объединяет любое количество адресов в одну строку отчёта. */
function formatViolationAddresses_(addresses) {
  return addresses.join("; ");
}


/** Считает разницу календарных дней между двумя ключами yyyy-MM-dd. */
function calendarDaysBetween_(earlierKey, laterKey) {
  const earlier = calendarDayNumber_(earlierKey);
  const later = calendarDayNumber_(laterKey);
  return Number.isFinite(earlier) && Number.isFinite(later)
    ? later - earlier
    : NaN;
}


function calendarDayNumber_(dateKey) {
  const match = String(dateKey || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return NaN;

  return Date.UTC(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3])
  ) / 86400000;
}


/** Проверяет заголовки управляемой таблицы A5:C5. */
function validateViolatorsHeaders_(sheet) {
  const actual = sheet.getRange(
    APP_CONFIG_.violators.headerRow,
    1,
    1,
    APP_CONFIG_.violators.columns
  ).getDisplayValues()[0];
  validateViolatorsHeaderValues_(sheet, actual);
}


function validateViolatorsHeaderValues_(sheet, actual) {
  APP_CONFIG_.violators.headers.forEach(function(expected, index) {
    if (normalizeText_(actual[index]) !== normalizeText_(expected)) {
      throw new Error(
        "На листе «" + sheet.getName() + "» в столбце " +
        columnToLetter_(index + 1) + " ожидался заголовок «" + expected +
        "», найдено «" + actual[index] + "»."
      );
    }
  });
}


/** При необходимости расширяет нативную таблицу под новых сотрудников. */
function ensureViolatorsTableCapacity_(
  spreadsheet,
  sheet,
  operatorCount,
  knownMetadata
) {
  const metadata = knownMetadata ||
    fetchSpreadsheetTableMetadata_(spreadsheet.getId());
  const matchingSheets = (metadata.sheets || []).filter(function(item) {
    return item.properties &&
      item.properties.title === APP_CONFIG_.violators.sheet;
  });
  if (matchingSheets.length !== 1) {
    throw new Error(
      "Для листа «" + APP_CONFIG_.violators.sheet +
      "» найдено листов в API: " + matchingSheets.length + "."
    );
  }

  const matches = (matchingSheets[0].tables || []).filter(function(table) {
    return normalizeText_(table.name) ===
      normalizeText_(APP_CONFIG_.violators.table);
  });
  if (matches.length !== 1) {
    throw new Error(
      "На листе «" + APP_CONFIG_.violators.sheet +
      "» найдено нативных таблиц «" + APP_CONFIG_.violators.table +
      "»: " + matches.length + "."
    );
  }

  const table = matches[0];
  const range = table.range || {};
  const expectedStartRow = APP_CONFIG_.violators.headerRow - 1;
  if (
    range.sheetId !== sheet.getSheetId() ||
    range.startRowIndex !== expectedStartRow ||
    (range.startColumnIndex || 0) !== 0 ||
    range.endColumnIndex !== APP_CONFIG_.violators.columns
  ) {
    throw new Error(
      "Нативная таблица «" + APP_CONFIG_.violators.table +
      "» должна занимать столбцы A:C и начинаться со строки 5."
    );
  }

  const requiredEndRow = expectedStartRow + Math.max(operatorCount + 1, 2);
  if (requiredEndRow <= range.endRowIndex) return table;

  if (requiredEndRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      requiredEndRow - sheet.getMaxRows()
    );
  }

  const expandedRange = Object.assign({}, range, {
    endRowIndex: requiredEndRow
  });
  sendSheetsBatchUpdate_(spreadsheet.getId(), [{
    updateTable: {
      table: {
        tableId: table.tableId,
        range: expandedRange
      },
      fields: "range"
    }
  }]);
  table.range = expandedRange;
  return table;
}


/** Формирует и пишет компактную диагностическую запись одного прохода. */
function finishViolatorsPerformance_(mode, startedAt, details) {
  const performance = Object.assign({
    mode: mode,
    durationMs: Date.now() - startedAt,
    finishedAt: new Date().toISOString()
  }, details || {});
  console.log("[Violators performance] " + JSON.stringify(performance));
  return performance;
}
