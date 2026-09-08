/**
 * Constants.gs - единая конфигурация проекта.
 *
 * Содержит имена рабочих и справочных листов, заголовки столбцов, URL
 * Logviewer, оформление, правила отчёта «Нарушители», ключи PropertiesService и
 * имена триггеров. Остальные файлы получают служебные значения только отсюда.
 */
const APP_CONFIG_ = Object.freeze({
  workingSheets: Object.freeze(["ПАО", "Такси", "Шаттл"]),
  headerRow: 1,
  firstDataRow: 2,
  timeZone: "Europe/Moscow",

  headers: Object.freeze({
    date: "Дата",
    vehicle: "ВАТС",
    teleoperator: "Телеоператор",
    logviewer: "Logviewer",
    logCheckDate: "Дата проверки лога"
  }),

  logviewer: Object.freeze({
    baseUrl: "https://logviewer.df.sbauto.tech/",
    realLogPath: "logs/",
    addLogLabel: "Добавить лог",
    todayLabel: "Сегодня",
    filledLabel: "заполнен",
    maintenanceChunkRows: 500
  }),

  formatting: Object.freeze({
    filledBackground: "#d9ecc0",
    filledText: "#35714e"
  }),

  references: Object.freeze({
    teleoperatorsSheet: "Телеоператоры",
    vehiclesSheet: "ВАТС",
    teleoperatorLoginHeaders: Object.freeze([
      "Логин",
      "Телеоператор",
      "Телеоператоры"
    ]),
    teleoperatorNameHeader: "ФИО",
    vehicleNameHeader: "Название",
    vehicleNumberHeader: "Номер ВАТС",
    vehicleAddedAtHeader: "Дата добавления",
    vehicleStatusHeader: "Статус",
    activeVehicleStatus: "Active",
    inactiveVehicleStatus: "Inactive"
  }),

  violators: Object.freeze({
    sheet: "Нарушители",
    table: "Нарушители",
    headerRow: 5,
    firstDataRow: 6,
    columns: 3,
    headers: Object.freeze([
      "Телеоператор",
      "Дата не обновлена",
      "Адрес ошибки (Не обновлена)\nЛист / строка"
    ]),
    marker: "😡",
    staleAfterDays: 4,
    activeWindowDays: 30,
    fullSyncIntervalMinutes: 10
  }),

  properties: Object.freeze({
    spreadsheetId: "bound_spreadsheet_id_v1",
    activeWorkingSheets: "active_working_sheets_v1",
    documentationPromptHtml: "documentation_prompt_html_v1",
    documentationPromptBackend: "documentation_prompt_backend_v1",
    teleoperatorsHash: "teleoperators_sync_hash_v2",
    vehiclesHash: "vehicles_sync_hash_v1",
    violatorsLastFullSync: "violators_last_full_sync_v1"
  }),

  sheetsApiUrl: "https://sheets.googleapis.com/v4/spreadsheets/",

  triggers: Object.freeze({
    editHandler: "handleReferenceSheetsEdit",
    changeHandler: "handleReferenceSheetsChange",
    violatorsPeriodicHandler: "handleViolatorsPeriodicSync",
    legacyHandlers: Object.freeze([
      "handleTeleoperatorsEdit",
      "handleTeleoperatorsChange",
      "handleViolatorsDailySync"
    ])
  })
});
