/**
 * Logviewer.gs — ссылки Logviewer и связанный статус проверки.
 *
 * Строит служебную ссылку «Добавить лог» по дате и ВАТС, распознаёт настоящие
 * /logs/-ссылки, динамически ставит или снимает «заполнен» и содержит ручные
 * операции для восстановления ссылок и сверки уже существующих статусов.
 */
function handleLogviewerEdit_(event) {
  if (!event || !event.range) return;

  const sheet = event.range.getSheet();
  if (!isWorkingSheetTracked_(sheet.getName())) return;
  if (event.range.getLastRow() < APP_CONFIG_.firstDataRow) return;

  const columns = getRequiredColumns_(sheet);
  if (!columns) return;

  const startRow = Math.max(APP_CONFIG_.firstDataRow, event.range.getRow());
  const rowCount = event.range.getLastRow() - startRow + 1;
  const dates = sheet.getRange(startRow, columns.date, rowCount, 1).getValues();
  const vehicles = sheet.getRange(
    startRow,
    columns.vehicle,
    rowCount,
    1
  ).getValues();
  const logDataRows = readLogCellData_(
    sheet.getRange(startRow, columns.logviewer, rowCount, 1)
  );
  const statuses = sheet.getRange(
    startRow,
    columns.logCheckDate,
    rowCount,
    1
  ).getValues();
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();

  for (let offset = 0; offset < rowCount; offset++) {
    const row = startRow + offset;
    const logData = logDataRows[offset];

    // Любое редактирование строки — повод восстановить служебную ссылку.
    // Это закрывает порядок заполнения, при котором дата и ВАТС уже стояли,
    // а пользователь менял только остальные поля кейса.
    updateAddLogLink_(
      sheet,
      row,
      columns.logviewer,
      dates[offset][0],
      vehicles[offset][0],
      logData,
      timeZone
    );

    // Статус динамически следует за настоящей ссылкой: вместе с /logs/
    // появляется «заполнен», после удаления возвращается дата события.
    syncFilledStatusWithLog_(
      sheet,
      row,
      columns.logCheckDate,
      statuses[offset][0],
      logData,
      dates[offset][0],
      timeZone
    );
  }
}


/**
 * Ручная служебная операция для восстановления ссылок и форматирования.
 * Настоящие ссылки и произвольный текст сохраняются. Контрольная дата уступает
 * настоящему логу, а после удаления лога восстанавливается из даты события.
 */
function fillMissingLinks() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];

  getActiveWorkingSheets_().forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    const columns = getRequiredColumns_(sheet);
    if (!columns) return;

    ensureLogCheckDateFormatting_(sheet, columns);

    const lastRow = sheet.getLastRow();
    if (lastRow < APP_CONFIG_.firstDataRow) return;

    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    let linksUpdated = 0;
    let statusesFilled = 0;
    let statusesCleared = 0;

    for (
      let startRow = APP_CONFIG_.firstDataRow;
      startRow <= lastRow;
      startRow += APP_CONFIG_.logviewer.maintenanceChunkRows
    ) {
      const rowCount = Math.min(
        APP_CONFIG_.logviewer.maintenanceChunkRows,
        lastRow - startRow + 1
      );
      const dates = sheet.getRange(
        startRow,
        columns.date,
        rowCount,
        1
      ).getValues();
      const vehicles = sheet.getRange(
        startRow,
        columns.vehicle,
        rowCount,
        1
      ).getValues();
      const logDataRows = readLogCellData_(sheet.getRange(
        startRow,
        columns.logviewer,
        rowCount,
        1
      ));
      const statuses = sheet.getRange(
        startRow,
        columns.logCheckDate,
        rowCount,
        1
      ).getValues();

      for (let offset = 0; offset < rowCount; offset++) {
        const row = startRow + offset;
        const logData = logDataRows[offset];

        if (hasRealLogLink_(logData)) {
          if (setFilledStatusForLog_(
            sheet,
            row,
            columns.logCheckDate,
            statuses[offset][0],
            timeZone
          )) {
            statusesFilled++;
          }
          continue;
        }

        if (clearFilledStatusWithoutLog_(
          sheet,
          row,
          columns.logCheckDate,
          statuses[offset][0],
          dates[offset][0],
          timeZone
        )) {
          statusesCleared++;
        }

        if (!isEmptyLogCell_(logData) && !isAddLogLink_(logData)) continue;

        const changed = updateAddLogLink_(
          sheet,
          row,
          columns.logviewer,
          dates[offset][0],
          vehicles[offset][0],
          logData,
          timeZone
        );
        if (changed) linksUpdated++;
      }
    }

    result.push({
      sheet: sheetName,
      linksUpdated: linksUpdated,
      statusesFilled: statusesFilled,
      statusesCleared: statusesCleared
    });
  });

  console.log(JSON.stringify(result));
  return result;
}


/**
 * Приводит «заполнен» в соответствие с реальными /logs/-ссылками. Большие
 * листы читает порциями, чтобы Spreadsheet service не обрывал запрос таймаутом.
 */
function reconcileLogviewerStatuses() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];

  getActiveWorkingSheets_().forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    const columns = getRequiredColumns_(sheet);
    if (!columns) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < APP_CONFIG_.firstDataRow) return;

    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    let filled = 0;
    let cleared = 0;

    for (
      let startRow = APP_CONFIG_.firstDataRow;
      startRow <= lastRow;
      startRow += APP_CONFIG_.logviewer.maintenanceChunkRows
    ) {
      const rowCount = Math.min(
        APP_CONFIG_.logviewer.maintenanceChunkRows,
        lastRow - startRow + 1
      );
      const dates = sheet.getRange(
        startRow,
        columns.date,
        rowCount,
        1
      ).getValues();
      const logDataRows = readLogCellData_(sheet.getRange(
        startRow,
        columns.logviewer,
        rowCount,
        1
      ));
      const statuses = sheet.getRange(
        startRow,
        columns.logCheckDate,
        rowCount,
        1
      ).getValues();

      logDataRows.forEach(function(logData, offset) {
        const row = startRow + offset;
        const action = syncFilledStatusWithLog_(
          sheet,
          row,
          columns.logCheckDate,
          statuses[offset][0],
          logData,
          dates[offset][0],
          timeZone
        );
        if (action === "filled") filled++;
        if (action === "cleared") cleared++;
      });
    }

    result.push({sheet: sheetName, filled: filled, cleared: cleared});
  });

  console.log(JSON.stringify(result));
  return result;
}


/** Синхронизирует одну ячейку статуса с настоящей /logs/-ссылкой. */
function syncFilledStatusWithLog_(
  sheet,
  row,
  statusColumn,
  currentValue,
  logData,
  eventDate,
  timeZone
) {
  if (hasRealLogLink_(logData)) {
    return setFilledStatusForLog_(
      sheet,
      row,
      statusColumn,
      currentValue,
      timeZone
    ) ? "filled" : "unchanged";
  }

  return clearFilledStatusWithoutLog_(
    sheet,
    row,
    statusColumn,
    currentValue,
    eventDate,
    timeZone
  ) ? "cleared" : "unchanged";
}


/** Ставит «заполнен» поверх пустоты или календарной контрольной даты. */
function setFilledStatusForLog_(
  sheet,
  row,
  statusColumn,
  currentValue,
  timeZone
) {
  if (
    normalizeText_(currentValue) ===
    normalizeText_(APP_CONFIG_.logviewer.filledLabel)
  ) {
    return false;
  }
  if (!isBlankCellValue_(currentValue) && !dateKey_(currentValue, timeZone)) {
    return false;
  }

  sheet.getRange(row, statusColumn).setValue(
    APP_CONFIG_.logviewer.filledLabel
  );
  return true;
}


/** При удалении настоящего лога возвращает актуальную дату события. */
function clearFilledStatusWithoutLog_(
  sheet,
  row,
  statusColumn,
  currentValue,
  eventDate,
  timeZone
) {
  if (
    normalizeText_(currentValue) !==
    normalizeText_(APP_CONFIG_.logviewer.filledLabel)
  ) {
    return false;
  }

  const eventKey = dateKey_(eventDate, timeZone);
  if (eventKey) {
    setLogCheckDate_(sheet, row, statusColumn, eventKey, timeZone);
  } else {
    sheet.getRange(row, statusColumn).clearContent();
  }
  return true;
}


/** Создаёт, обновляет или очищает только служебную ссылку «Добавить лог». */
function updateAddLogLink_(
  sheet,
  row,
  logviewerColumn,
  date,
  vehicle,
  logData,
  timeZone
) {
  if (hasRealLogLink_(logData)) return false;
  if (!isEmptyLogCell_(logData) && !isAddLogLink_(logData)) return false;

  const formula = buildAddLogFormula_(date, vehicle, timeZone);
  const cell = sheet.getRange(row, logviewerColumn);

  if (formula) {
    if (logData.formula !== formula) {
      cell.setFormula(formula);
      return true;
    }
  } else if (isAddLogLink_(logData)) {
    cell.clearContent();
    return true;
  }

  return false;
}


/** Строит локализованную формулу HYPERLINK для русской таблицы. */
function buildAddLogFormula_(date, vehicle, timeZone) {
  const dateKey = dateKey_(date, timeZone);
  const normalizedVehicle = normalizeVehicle_(vehicle);
  if (!dateKey || !normalizedVehicle) return "";

  const parts = dateKey.split("-");
  const formattedDate = parts[2] + "." + parts[1] + "." + parts[0].slice(-2);
  const url = APP_CONFIG_.logviewer.baseUrl + "?rovers_regexp=" +
    encodeURIComponent(normalizedVehicle) + "&date_range=" + formattedDate +
    "%2C" + formattedDate;

  return '=HYPERLINK("' + url + '";"' +
    APP_CONFIG_.logviewer.addLogLabel + '")';
}


/** Пакетно читает все представления небольшого диапазона Logviewer. */
function readLogCellData_(range) {
  const values = range.getValues();
  const displays = range.getDisplayValues();
  const formulas = range.getFormulas();
  const richTexts = range.getRichTextValues();

  return values.map(function(row, index) {
    return {
      value: row[0],
      display: displays[index][0],
      formula: formulas[index][0],
      richText: richTexts[index][0]
    };
  });
}


/** Проверяет наличие настоящей ссылки /logs/ в любом представлении ячейки. */
function hasRealLogLink_(logData) {
  const realPrefix = (
    APP_CONFIG_.logviewer.baseUrl + APP_CONFIG_.logviewer.realLogPath
  ).toLowerCase();
  return collectLogCellText_(logData).toLowerCase().includes(realPrefix);
}


/** Отличает служебную ссылку от произвольного текста. */
function isAddLogLink_(logData) {
  const formula = normalizeFormula_(logData && logData.formula);
  const baseUrl = APP_CONFIG_.logviewer.baseUrl.toLowerCase();
  if (
    formula.includes("hyperlink(") &&
    formula.includes(baseUrl) &&
    formula.includes("rovers_regexp=")
  ) {
    return true;
  }

  return collectLogCellLinks_(logData).some(function(url) {
    const normalizedUrl = String(url).toLowerCase();
    return normalizedUrl.startsWith(baseUrl) &&
      normalizedUrl.includes("rovers_regexp=");
  });
}


function isEmptyLogCell_(logData) {
  return collectLogCellText_(logData).trim() === "";
}


/** Собирает текст, формулу и URL для распознавания содержимого Logviewer. */
function collectLogCellText_(logData) {
  const parts = [
    logData && logData.value,
    logData && logData.display,
    logData && logData.formula
  ].concat(collectLogCellLinks_(logData));

  return parts.filter(function(value) {
    return value !== null && value !== undefined && value !== "";
  }).map(String).join(" ");
}


/** Возвращает прямые и посегментные RichText-ссылки. */
function collectLogCellLinks_(logData) {
  const links = [];
  if (!logData || !logData.richText) return links;

  const directLink = logData.richText.getLinkUrl();
  if (directLink) links.push(directLink);

  logData.richText.getRuns().forEach(function(run) {
    const link = run.getLinkUrl();
    if (link) links.push(link);
  });

  return links;
}
