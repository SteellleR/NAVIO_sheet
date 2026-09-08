/**
 * Setup.gs — установка и обслуживание проекта.
 *
 * Создаёт installable onEdit/onChange и десятиминутный контрольный триггер,
 * выполняет первую синхронизацию справочников, дат, статусов и отчёта, а также
 * удаляет только те
 * триггеры, которые принадлежат этому проекту.
 */
function installProject() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty(
    APP_CONFIG_.properties.spreadsheetId,
    spreadsheet.getId()
  );
  removeProjectTriggers();

  setupLogCheckDateFormatting();
  const logCheckDate = syncLogCheckDateDropdowns();
  const initializedLogCheckDates = syncLogCheckDatesFromEvents();
  const teleoperators = syncTeleoperatorDropdowns(true);
  const vehiclesSheet = spreadsheet.getSheetByName(
    APP_CONFIG_.references.vehiclesSheet
  );
  const vehicles = vehiclesSheet
    ? syncVehicleDropdowns(true)
    : {updated: false, reason: "sheet_missing"};
  // Полная сверка заодно восстанавливает пропущенные «Добавить лог».
  const logviewerStatuses = fillMissingLinks();
  const violators = syncViolators();

  // Триггеры создаются после массовой синхронизации. Иначе изменения нативных
  // таблиц во время установки порождают пачку параллельных onChange-запусков.
  ScriptApp.newTrigger(APP_CONFIG_.triggers.editHandler)
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
  ScriptApp.newTrigger(APP_CONFIG_.triggers.changeHandler)
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();
  ScriptApp.newTrigger(APP_CONFIG_.triggers.violatorsPeriodicHandler)
    .timeBased()
    .everyMinutes(APP_CONFIG_.violators.fullSyncIntervalMinutes)
    .create();

  spreadsheet.toast("Проект установлен и синхронизирован", "Готово", 5);

  return {
    logCheckDate: logCheckDate,
    initializedLogCheckDates: initializedLogCheckDates,
    teleoperators: teleoperators,
    vehicles: vehicles,
    logviewerStatuses: logviewerStatuses,
    violators: violators
  };
}


/** Удаляет только триггеры, принадлежащие этому проекту. */
function removeProjectTriggers() {
  const handlers = [
    APP_CONFIG_.triggers.editHandler,
    APP_CONFIG_.triggers.changeHandler,
    APP_CONFIG_.triggers.violatorsPeriodicHandler
  ].concat(APP_CONFIG_.triggers.legacyHandlers);
  let removed = 0;

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (!handlers.includes(trigger.getHandlerFunction())) return;
    ScriptApp.deleteTrigger(trigger);
    removed++;
  });

  return removed;
}


/** Обратная совместимость со старым названием установочной функции. */
function installTeleoperatorSync() {
  return installProject();
}


/** Добавляет в таблицу безопасные ручные команды установки и диагностики. */
function addProjectMenu_() {
  const ui = SpreadsheetApp.getUi();
  const documentationMenu = ui.createMenu("Документация")
    .addItem("Для пользователя", "showUserDocumentation")
    .addItem("Для разработчика", "showDeveloperDocumentation");

  ui
    .createMenu("NAVIO")
    .addItem("Установить или обновить проект", "installProject")
    .addItem("Полный пересчёт нарушителей", "syncViolators")
    .addSeparator()
    .addItem("Управление листами", "showWorkingSheetTrackingDialog")
    .addSubMenu(documentationMenu)
    .addToUi();
}


/** Поддерживает старый installable onEdit до запуска новой установки. */
function handleTeleoperatorsEdit(event) {
  return handleReferenceSheetsEdit(event);
}


/** Поддерживает старый installable onChange до запуска новой установки. */
function handleTeleoperatorsChange(event) {
  return handleReferenceSheetsChange(event);
}


/** Общая точка входа установочного onEdit-триггера для справочников. */
function handleReferenceSheetsEdit(event) {
  ensureViolatorsPeriodicTrigger_();
  if (handleTeleoperatorsReferenceEdit_(event)) return;
  handleVehiclesReferenceEdit_(event);
}


/** Общая точка входа onChange для вставки и удаления строк/столбцов. */
function handleReferenceSheetsChange(event) {
  if (!event) return;

  ensureViolatorsPeriodicTrigger_();

  const relevantChanges = [
    "INSERT_ROW",
    "REMOVE_ROW",
    "INSERT_COLUMN",
    "REMOVE_COLUMN"
  ];
  if (event.changeType && !relevantChanges.includes(event.changeType)) return;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet.getSheetByName(APP_CONFIG_.references.teleoperatorsSheet)) {
    syncTeleoperatorDropdowns(true);
  }
  if (spreadsheet.getSheetByName(APP_CONFIG_.references.vehiclesSheet)) {
    syncVehicleDropdowns(false);
  }
  if (spreadsheet.getSheetByName(APP_CONFIG_.violators.sheet)) {
    syncViolators();
  }
}


/** Самовосстанавливает контрольный триггер после ближайшего edit/change. */
function ensureViolatorsPeriodicTrigger_() {
  const handler = APP_CONFIG_.triggers.violatorsPeriodicHandler;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    PropertiesService.getScriptProperties().setProperty(
      APP_CONFIG_.properties.spreadsheetId,
      spreadsheet.getId()
    );
  }

  const exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handler;
  });
  if (exists) return false;

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(APP_CONFIG_.violators.fullSyncIntervalMinutes)
    .create();
  return true;
}
