/**
 * SheetTracking.gs - включение и приостановка рабочих листов без правки кода.
 *
 * Хранит список активных листов в Document Properties, отдаёт состояние
 * HTML-диалогу и после изменения приводит отчёт и включённые листы в порядок.
 * Полностью отключить все рабочие листы нельзя.
 */
let ACTIVE_WORKING_SHEETS_CACHE_ = null;


/** Возвращает активные листы в стабильном порядке общей конфигурации. */
function getActiveWorkingSheets_() {
  if (ACTIVE_WORKING_SHEETS_CACHE_) {
    return ACTIVE_WORKING_SHEETS_CACHE_.slice();
  }

  const value = PropertiesService.getDocumentProperties().getProperty(
    APP_CONFIG_.properties.activeWorkingSheets
  );
  let saved = null;

  if (value) {
    try {
      saved = JSON.parse(value);
    } catch (error) {
      console.warn("Не удалось прочитать список активных листов: " + error);
    }
  }

  const selected = Array.isArray(saved) ? saved.map(String) : null;
  const active = selected
    ? APP_CONFIG_.workingSheets.filter(function(sheetName) {
      return selected.includes(sheetName);
    })
    : APP_CONFIG_.workingSheets.slice();

  // Повреждённое или устаревшее пустое свойство не отключает проект целиком.
  ACTIVE_WORKING_SHEETS_CACHE_ = active.length
    ? active
    : APP_CONFIG_.workingSheets.slice();
  return ACTIVE_WORKING_SHEETS_CACHE_.slice();
}


/** Быстрая проверка для onEdit-профильных обработчиков. */
function isWorkingSheetTracked_(sheetName) {
  return getActiveWorkingSheets_().includes(sheetName);
}


/** Состояние для HTML-интерфейса. */
function getWorkingSheetTrackingState() {
  const active = getActiveWorkingSheets_();
  return APP_CONFIG_.workingSheets.map(function(sheetName) {
    return {
      name: sheetName,
      active: active.includes(sheetName)
    };
  });
}


/** Открывает понятное окно управления из меню NAVIO. */
function showWorkingSheetTrackingDialog() {
  const html = HtmlService.createHtmlOutputFromFile("SheetTrackingDialog")
    .setWidth(560)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, "Управление листами");
}


/**
 * Сохраняет выбор, синхронизирует заново включённые листы и сразу очищает
 * отчёт от адресов приостановленных листов.
 */
function saveWorkingSheetTrackingState(activeSheets) {
  if (!Array.isArray(activeSheets)) {
    throw new Error("Некорректный список листов.");
  }

  const requested = activeSheets.map(String);
  const unknown = requested.filter(function(sheetName) {
    return !APP_CONFIG_.workingSheets.includes(sheetName);
  });
  if (unknown.length) {
    throw new Error("Неизвестные рабочие листы: " + unknown.join(", "));
  }

  const next = APP_CONFIG_.workingSheets.filter(function(sheetName) {
    return requested.includes(sheetName);
  });
  if (!next.length) {
    throw new Error("Нужно оставить включённым хотя бы один рабочий лист.");
  }

  const previous = getActiveWorkingSheets_();
  const changed = previous.join("\n") !== next.join("\n");
  if (!changed) {
    return {
      changed: false,
      sheets: getWorkingSheetTrackingState()
    };
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Проект сейчас занят. Попробуйте сохранить ещё раз.");
  }

  try {
    PropertiesService.getDocumentProperties().setProperty(
      APP_CONFIG_.properties.activeWorkingSheets,
      JSON.stringify(next)
    );
    ACTIVE_WORKING_SHEETS_CACHE_ = next.slice();
  } finally {
    lock.releaseLock();
  }

  const newlyEnabled = next.filter(function(sheetName) {
    return !previous.includes(sheetName);
  });
  const maintenance = newlyEnabled.length
    ? synchronizeNewlyEnabledWorkingSheets_()
    : null;
  const violators = syncViolators();

  return {
    changed: true,
    newlyEnabled: newlyEnabled,
    sheets: getWorkingSheetTrackingState(),
    maintenance: maintenance,
    violators: violators
  };
}


/** Подтягивает на включённые листы актуальные справочники, даты и ссылки. */
function synchronizeNewlyEnabledWorkingSheets_() {
  const spreadsheet = getProjectSpreadsheet_();

  setupLogCheckDateFormatting();
  const logCheckDate = syncLogCheckDateDropdowns();
  const dates = syncLogCheckDatesFromEvents();
  const teleoperators = syncTeleoperatorDropdowns(true);
  const vehiclesSheet = spreadsheet.getSheetByName(
    APP_CONFIG_.references.vehiclesSheet
  );
  const vehicles = vehiclesSheet
    ? syncVehicleDropdowns(true)
    : {updated: false, reason: "sheet_missing"};
  const logviewer = fillMissingLinks();

  return {
    logCheckDate: logCheckDate,
    dates: dates,
    teleoperators: teleoperators,
    vehicles: vehicles,
    logviewer: logviewer
  };
}
