/**
 * LogCheckDate.gs - жизненный цикл значения «Дата проверки лога».
 *
 * Один раз копирует дату события в контрольную дату, переносит исправление даты,
 * пока значение не актуализировали вручную, превращает «Сегодня» в фиксированную
 * дату и синхронизирует dropdown «Сегодня | заполнен» на рабочих листах. Даты
 * передаёт как дд.мм.гггг, чтобы нативный тип столбца не получил голый serial.
 */
function handleLogCheckDateEdit_(event) {
  if (!event || !event.range) return;

  const sheet = event.range.getSheet();
  if (!isWorkingSheetTracked_(sheet.getName())) return;
  if (event.range.getLastRow() < APP_CONFIG_.firstDataRow) return;

  const columns = getRequiredColumns_(sheet);
  if (!columns) return;

  const dateEdited = rangeTouchesColumn_(event.range, columns.date);
  const statusEdited = rangeTouchesColumn_(event.range, columns.logCheckDate);

  const startRow = Math.max(APP_CONFIG_.firstDataRow, event.range.getRow());
  const endRow = event.range.getLastRow();
  const rowCount = endRow - startRow + 1;
  const dates = sheet.getRange(startRow, columns.date, rowCount, 1).getValues();
  const statusRange = sheet.getRange(
    startRow,
    columns.logCheckDate,
    rowCount,
    1
  );
  const statuses = statusRange.getValues();
  const statusDisplays = statusRange.getDisplayValues();
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const isSingleDateCell = dateEdited &&
    event.range.getRow() === startRow &&
    event.range.getNumRows() === 1 &&
    event.range.getColumn() === columns.date &&
    event.range.getNumColumns() === 1;

  for (let offset = 0; offset < rowCount; offset++) {
    const row = startRow + offset;
    const status = statuses[offset][0];

    if (
      statusEdited &&
      normalizeText_(status) ===
        normalizeText_(APP_CONFIG_.logviewer.todayLabel)
    ) {
      setTodayDate_(sheet, row, columns.logCheckDate, timeZone);
      continue;
    }

    syncEventDateToLogCheckDate_(
      sheet,
      row,
      columns.logCheckDate,
      dates[offset][0],
      status,
      isSingleDateCell ? event.oldValue : null,
      timeZone,
      statusDisplays[offset][0]
    );
  }
}


/**
 * Заполняет пустые контрольные даты и нормализует существующие даты во всех
 * строках. Читает листы порциями; «заполнен» и произвольный текст не трогает.
 */
function syncLogCheckDatesFromEvents() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const result = [];

  getActiveWorkingSheets_().forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    const columns = getRequiredColumns_(sheet);
    if (!columns) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < APP_CONFIG_.firstDataRow) {
      result.push({sheet: sheetName, initialized: 0, normalized: 0});
      return;
    }

    let initialized = 0;
    let normalized = 0;

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
      const statusRange = sheet.getRange(
        startRow,
        columns.logCheckDate,
        rowCount,
        1
      );
      const statuses = statusRange.getValues();
      const statusDisplays = statusRange.getDisplayValues();
      let chunkChanged = false;

      for (let offset = 0; offset < rowCount; offset++) {
        const currentStatus = statuses[offset][0];
        const eventKey = dateKey_(dates[offset][0], timeZone);

        if (isBlankCellValue_(currentStatus)) {
          if (!eventKey) continue;
          statuses[offset][0] = formatLogCheckDateKey_(eventKey);
          initialized++;
          chunkChanged = true;
          continue;
        }

        const statusKey = dateKey_(currentStatus, timeZone);
        if (!statusKey) continue;

        const normalizedStatus = formatLogCheckDateKey_(statusKey);
        if (statusDisplays[offset][0].trim() === normalizedStatus) continue;

        statuses[offset][0] = normalizedStatus;
        normalized++;
        chunkChanged = true;
      }

      if (chunkChanged) statusRange.setValues(statuses);
    }
    result.push({
      sheet: sheetName,
      initialized: initialized,
      normalized: normalized
    });
  });

  console.log(JSON.stringify(result));
  return result;
}


/**
 * Обновляет автоматически связанную дату. Ручная актуализация и «заполнен»
 * защищены от изменения даты события.
 */
function syncEventDateToLogCheckDate_(
  sheet,
  row,
  statusColumn,
  eventDate,
  currentStatus,
  previousEventDate,
  timeZone,
  currentStatusDisplay
) {
  const eventKey = dateKey_(eventDate, timeZone);
  const statusKey = dateKey_(currentStatus, timeZone);
  const previousEventKey = dateKey_(previousEventDate, timeZone);
  const followsPreviousEventDate = previousEventKey &&
    statusKey === previousEventKey;

  if (!eventKey) {
    if (!followsPreviousEventDate) return "unchanged";
    sheet.getRange(row, statusColumn).clearContent();
    return "cleared";
  }

  if (!isBlankCellValue_(currentStatus) && !followsPreviousEventDate) {
    if (!statusKey) return "unchanged";

    const normalizedStatus = formatLogCheckDateKey_(statusKey);
    const displayedStatus = currentStatusDisplay === undefined
      ? String(currentStatus).trim()
      : String(currentStatusDisplay).trim();
    if (displayedStatus === normalizedStatus) return "unchanged";

    setLogCheckDate_(sheet, row, statusColumn, statusKey, timeZone);
    return "normalized";
  }
  if (
    statusKey === eventKey &&
    String(
      currentStatusDisplay === undefined
        ? currentStatus
        : currentStatusDisplay
    ).trim() === formatLogCheckDateKey_(eventKey)
  ) {
    return "unchanged";
  }

  setLogCheckDate_(sheet, row, statusColumn, eventKey, timeZone);
  return isBlankCellValue_(currentStatus) ? "initialized" : "corrected";
}


/**
 * Приводит нативную шторку «Дата проверки лога» на всех рабочих листах
 * к единому набору вариантов. Существующие значения ячеек не переписывает.
 */
function syncLogCheckDateDropdowns() {
  return syncNativeTableDropdown_({
    targetSheets: getActiveWorkingSheets_(),
    targetHeader: APP_CONFIG_.headers.logCheckDate,
    values: [
      APP_CONFIG_.logviewer.todayLabel,
      APP_CONFIG_.logviewer.filledLabel
    ]
  });
}


/** Записывает текущую дату обычным значением, а не формулой TODAY(). */
function setTodayDate_(sheet, row, statusColumn, timeZone) {
  const dateText = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  return setLogCheckDate_(sheet, row, statusColumn, dateText, timeZone);
}


/** Передаёт календарную дату в Sheets локализованным значением дд.мм.гггг. */
function setLogCheckDate_(sheet, row, statusColumn, dateKey, timeZone) {
  const dateText = formatLogCheckDateKey_(dateKey);
  sheet.getRange(row, statusColumn).setValue(dateText);
  return dateText;
}


/** Преобразует yyyy-MM-dd в стабильное отображаемое значение дд.мм.гггг. */
function formatLogCheckDateKey_(dateKey) {
  const parts = String(dateKey).split("-");
  if (parts.length !== 3) return "";
  return parts[2] + "." + parts[1] + "." + parts[0];
}
