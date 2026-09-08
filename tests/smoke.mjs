/**
 * smoke.mjs - быстрые локальные тесты чистой бизнес-логики Apps Script.
 * Выполняет .gs-файлы в VM и проверяет даты, ссылки, статусы и нарушения.
 */
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";

const fileOrder = [
  "Constants.gs",
  "Utils.gs",
  "SheetTracking.gs",
  "Documentation.gs",
  "SheetsApi.gs",
  "DropdownSync.gs",
  "Logviewer.gs",
  "LogCheckDate.gs",
  "Teleoperators.gs",
  "Violators.gs",
  "Vehicles.gs",
  "Formatting.gs",
  "Setup.gs",
  "EntryPoints.gs"
];
const documentProperties = new Map();
const context = vm.createContext({
  console: {log() {}, warn() {}, error() {}},
  PropertiesService: {
    getDocumentProperties() {
      return {
        getProperty(key) { return documentProperties.get(key) || null; },
        setProperty(key, value) {
          documentProperties.set(key, String(value));
          return this;
        },
        getProperties() { return Object.fromEntries(documentProperties); },
        setProperties(values) {
          Object.entries(values).forEach(([key, value]) => {
            documentProperties.set(key, String(value));
          });
          return this;
        },
        deleteProperty(key) {
          documentProperties.delete(key);
          return this;
        }
      };
    }
  },
  LockService: {
    getDocumentLock() {
      return {
        tryLock() { return true; },
        releaseLock() {}
      };
    }
  },
  Utilities: {
    formatDate(date) { return date.toISOString().slice(0, 10); },
    parseDate(value) { return new Date(`${value}T00:00:00.000Z`); }
  }
});

for (const fileName of fileOrder) {
  const source = fs.readFileSync(path.resolve("src", fileName), "utf8");
  vm.runInContext(source, context, {filename: fileName});
}

const setupSource = fs.readFileSync(path.resolve("src", "Setup.gs"), "utf8");
const documentationHtml = fs.readFileSync(
  path.resolve("src", "DocumentationDialog.html"),
  "utf8"
);
assert.ok(
  setupSource.indexOf("const violators = syncViolators();") <
    setupSource.indexOf("ScriptApp.newTrigger"),
  "installProject должен создавать триггеры после массовой синхронизации"
);
assert.doesNotMatch(
  setupSource.match(/const relevantChanges = \[[\s\S]*?\];/)[0],
  /"OTHER"/,
  "onChange не должен реагировать на технические изменения OTHER"
);
assert.match(
  setupSource,
  /\.addSubMenu\(documentationMenu\)/,
  "меню NAVIO должно содержать вложенную документацию"
);
assert.match(
  setupSource,
  /Для пользователя[\s\S]*Для разработчика/,
  "в документации должны быть отдельные пользовательский и технический пункты"
);
assert.doesNotMatch(
  documentationHtml,
  /id="expandAll"/,
  "глобальная кнопка раскрытия документации должна быть удалена"
);
assert.match(
  documentationHtml,
  /link\.addEventListener\("click", function\(event\)[\s\S]*event\.preventDefault\(\)[\s\S]*window\.scrollTo/,
  "пункты содержания должны явно прокручивать iframe к выбранному разделу"
);
assert.match(
  documentationHtml,
  /data-prompt-card="html"[\s\S]*data-prompt-card="backend"/,
  "developer-документация должна содержать два раздельных ИИ-промпта"
);
assert.match(
  documentationHtml,
  /Скачать TXT[\s\S]*ограничение по длине сообщения/,
  "developer-документация должна объяснять запасной путь для длинных промптов"
);
assert.match(
  documentationHtml,
  /ПОДРОБНАЯ КАРТА ФАЙЛОВ И ОТВЕТСТВЕННОСТИ[\s\S]*ЖИЗНЕННЫЙ ЦИКЛ СТРОКИ[\s\S]*PERFORMАNCE И КВОТЫ/,
  "backend-промпт должен содержать самостоятельный технический контекст"
);
assert.match(
  documentationHtml,
  /Если изменение мелкое[\s\S]*Если backend-изменение несущественное/,
  "оба промпта должны различать существенные и косметические изменения"
);
assert.match(
  documentationHtml,
  /После плана отдельной строкой спроси[^\n]*Начинаем\?/,
  "ИИ должен получить согласование плана до начала реализации"
);
assert.match(
  documentationHtml,
  /новый рабочий лист[\s\S]*NAVIO → Управление листами/,
  "backend-промпт должен включать новый лист в управление отслеживанием"
);
assert.match(
  documentationHtml,
  /class="callout danger no-symbol"[\s\S]*Разработку безопаснее вести на копии таблицы/,
  "красное предупреждение должно требовать сначала тестировать на копии"
);
assert.match(
  documentationHtml,
  /data-prompt-action="copy"[\s\S]*data-prompt-action="download"[\s\S]*data-prompt-action="edit"/,
  "ИИ-промпты должны поддерживать копирование, TXT и редактирование"
);
assert.match(
  documentationHtml,
  /https:\/\/github\.com\/SteellleR\/NAVIO_sheet/,
  "developer-документация должна вести в открытый GitHub-репозиторий"
);

function evaluate(expression) {
  return vm.runInContext(expression, context);
}

assert.equal(
  evaluate("JSON.stringify(getActiveWorkingSheets_())"),
  JSON.stringify(["ПАО", "Такси", "Шаттл"])
);
documentProperties.set(
  "active_working_sheets_v1",
  JSON.stringify(["Такси", "Шаттл"])
);
evaluate("ACTIVE_WORKING_SHEETS_CACHE_ = null");
assert.equal(
  evaluate("JSON.stringify(getActiveWorkingSheets_())"),
  JSON.stringify(["Такси", "Шаттл"])
);
assert.equal(evaluate('isWorkingSheetTracked_("ПАО")'), false);
assert.throws(
  () => evaluate("saveWorkingSheetTrackingState([])"),
  /хотя бы один рабочий лист/
);
assert.throws(
  () => evaluate('saveWorkingSheetTrackingState(["Такси", "Чужой лист"])'),
  /Неизвестные рабочие листы/
);
documentProperties.delete("active_working_sheets_v1");
evaluate("ACTIVE_WORKING_SHEETS_CACHE_ = null");

const longDocumentationPrompt = "Обновлённый промпт для проверки. ".repeat(220);
vm.runInContext(
  `var __longDocumentationPrompt = ${JSON.stringify(longDocumentationPrompt)};`,
  context
);
assert.equal(
  evaluate('saveDocumentationPrompt("html", __longDocumentationPrompt).customized'),
  true
);
assert.equal(evaluate('readDocumentationPrompt_("html")'), longDocumentationPrompt.trim());
assert.ok(
  Number(documentProperties.get("documentation_prompt_html_v1_parts")) > 1,
  "длинный ИИ-промпт должен храниться несколькими частями"
);
assert.equal(
  evaluate('getDocumentationPromptOverrides().html'),
  longDocumentationPrompt.trim()
);
assert.equal(evaluate('resetDocumentationPrompt("html").customized'), false);
assert.equal(evaluate('getDocumentationPromptOverrides().html'), null);
assert.throws(
  () => evaluate('saveDocumentationPrompt("unknown", "текст")'),
  /Неизвестный тип промпта/
);

assert.equal(evaluate('normalizeVehicle_("кс2-021")'), "kc2-021");
assert.equal(evaluate('normalizeVehicle_("КС2 21")'), "kc2-021");
assert.equal(evaluate('normalizeVehicle_("KC2-21")'), "kc2-021");
assert.equal(evaluate('normalizeVehicle_("kc2_021")'), "kc2-021");
assert.equal(evaluate('normalizeVehicle_("SH-2")'), "sh-002");
assert.equal(evaluate('dateKey_("31.02.2026", "Europe/Moscow")'), "");
assert.equal(
  evaluate('dateKey_("27.08.26", "Europe/Moscow")'),
  "2026-08-27"
);
assert.equal(evaluate('dateKey_(46272, "Europe/Moscow")'), "2026-09-07");
assert.equal(evaluate('formatLogCheckDateKey_("2026-09-07")'), "07.09.2026");
assert.equal(
  evaluate('buildAddLogFormula_("27.08.26", "КС2 21", "Europe/Moscow")'),
  '=HYPERLINK("https://logviewer.df.sbauto.tech/?rovers_regexp=kc2-021&date_range=27.08.26%2C27.08.26";"Добавить лог")'
);
assert.equal(
  evaluate('isAddLogLink_({value: "24.08.26 logov net", display: "24.08.26 logov net", formula: "", richText: null})'),
  false
);
assert.equal(evaluate('isBlankCellValue_(new Date())'), false);
assert.equal(evaluate('isBlankCellValue_("заполнен")'), false);
assert.equal(evaluate('isBlankCellValue_("")'), true);
assert.equal(
  evaluate('classifyLogDateViolation_("20.04.2026", "20.04.2026", "2026-04-23", "Europe/Moscow")'),
  ""
);
assert.equal(
  evaluate('classifyLogDateViolation_("20.04.2026", "20.04.2026", "2026-04-24", "Europe/Moscow")'),
  "stale"
);
assert.equal(
  evaluate('classifyLogDateViolation_("20.04.2026", "21.04.2026", "2026-04-24", "Europe/Moscow")'),
  ""
);
assert.equal(
  evaluate('classifyLogDateViolation_("20.03.2026", "20.03.2026", "2026-04-24", "Europe/Moscow")'),
  ""
);
assert.equal(
  evaluate('classifyLogDateViolation_("20.04.2026", "заполнен", "2026-04-24", "Europe/Moscow")'),
  ""
);
assert.equal(
  evaluate('classifyLogDateViolation_("20.04.2026", "", "2026-04-24", "Europe/Moscow")'),
  ""
);
assert.equal(evaluate('formatViolationAddress_("Такси", 1852)'), "Такси / 1852");
assert.equal(
  evaluate('formatViolationAddresses_(["Такси / 1852", "Такси / 1876", "Шаттл / 76", "ПАО / 987"])'),
  "Такси / 1852; Такси / 1876; Шаттл / 76; ПАО / 987"
);
assert.equal(
  evaluate('JSON.stringify(parseViolationAddresses_("Такси / 1852; Такси / 1852; Шаттл / 76"))'),
  JSON.stringify(["Такси / 1852", "Шаттл / 76"])
);
assert.equal(
  evaluate('JSON.stringify(sortViolationAddresses_(["Шаттл / 76", "Такси / 1876", "ПАО / 987", "Такси / 1852"]))'),
  JSON.stringify([
    "ПАО / 987",
    "Такси / 1852",
    "Такси / 1876",
    "Шаттл / 76"
  ])
);

const layoutMetadata = {
  sheets: ["ПАО", "Такси", "Шаттл"].map((title, index) => ({
    properties: {title},
    tables: [{
      range: {
        endRowIndex: 100 + index,
        endColumnIndex: 14
      },
      columnProperties: [
        {columnName: "Дата"},
        {columnName: "Время"},
        {columnName: "Локация"},
        {columnName: "ВАТС"},
        {columnName: "Телеоператор"},
        {columnName: "Качество LTE"},
        {columnName: "Результат"},
        {columnName: "Нарушение ПДД"},
        {columnName: "Причина"},
        {columnName: "Классификатор"},
        {columnName: "Проблема"},
        {columnName: "Комментарий"},
        {columnName: "Logviewer"},
        {columnName: "Дата проверки лога"}
      ]
    }]
  }))
};
vm.runInContext(
  `var __layoutMetadata = ${JSON.stringify(layoutMetadata)};`,
  context
);
assert.equal(
  evaluate('buildWorkingTableLayouts_(__layoutMetadata)["ПАО"].lastDataRow'),
  100
);
assert.equal(
  evaluate('buildWorkingTableLayouts_(__layoutMetadata)["Шаттл"].columns.logCheckDate'),
  14
);

vm.runInContext(`
  var __entries = [{login: "operator-a"}, {login: "operator-b"}];
  var __reportRows = [
    ["operator-a", "😡", "Такси / 10"],
    ["operator-b", "", ""]
  ];
`, context);
assert.equal(
  evaluate('JSON.stringify(buildCurrentViolatorStates_(__entries, __reportRows))'),
  JSON.stringify([
    {login: "operator-a", stale: ["Такси / 10"]},
    {login: "operator-b", stale: []}
  ])
);

vm.runInContext(`
  var __statusWrites = 0;
  var __numberFormatWrites = 0;
  var __statusValue = null;
  var __statusSheet = {
    getRange: function() {
      return {
        setValue: function(value) {
          __statusWrites++;
          __statusValue = value;
          return this;
        },
        setNumberFormat: function() { __numberFormatWrites++; return this; },
        clearContent: function() { __statusValue = "cleared"; return this; }
      };
    }
  };
`, context);
assert.equal(
  evaluate('syncEventDateToLogCheckDate_(__statusSheet, 2, 14, "20.04.2026", "", null, "Europe/Moscow")'),
  "initialized"
);
assert.equal(evaluate("__statusWrites"), 1);
assert.equal(evaluate("__statusValue"), "20.04.2026");
assert.equal(evaluate("__numberFormatWrites"), 0);
assert.equal(
  evaluate('syncEventDateToLogCheckDate_(__statusSheet, 2, 14, "21.04.2026", "20.04.2026", "20.04.2026", "Europe/Moscow")'),
  "corrected"
);
assert.equal(evaluate("__statusWrites"), 2);
assert.equal(evaluate("__statusValue"), "21.04.2026");
assert.equal(
  evaluate('syncEventDateToLogCheckDate_(__statusSheet, 2, 14, "22.04.2026", "21.04.2026", "20.04.2026", "Europe/Moscow")'),
  "unchanged"
);
assert.equal(evaluate("__statusWrites"), 2);
assert.equal(
  evaluate('setFilledStatusForLog_(__statusSheet, 2, 14, "21.04.2026", "Europe/Moscow")'),
  true
);
assert.equal(evaluate('__statusValue'), "заполнен");
assert.equal(
  evaluate('clearFilledStatusWithoutLog_(__statusSheet, 2, 14, "заполнен", "22.04.2026", "Europe/Moscow")'),
  true
);
assert.equal(evaluate('__statusValue'), "22.04.2026");

assert.equal(
  evaluate('syncEventDateToLogCheckDate_(__statusSheet, 2, 14, "07.09.2026", 46272, null, "Europe/Moscow")'),
  "normalized"
);
assert.equal(evaluate('__statusValue'), "07.09.2026");
assert.equal(
  evaluate('syncEventDateToLogCheckDate_(__statusSheet, 2, 14, "07.09.2026", 46272, null, "Europe/Moscow", "07.09.2026")'),
  "unchanged"
);

vm.runInContext(`
  var __formulaWritten = "";
  var __emptyRichText = {
    getLinkUrl: function() { return null; },
    getRuns: function() { return []; }
  };
  var __rowSheet = {
    getName: function() { return "Такси"; },
    getParent: function() {
      return {getSpreadsheetTimeZone: function() { return "Europe/Moscow"; }};
    },
    getRange: function(row, column) {
      if (column === 1) return {getValues: function() { return [["07.09.2026"]]; }};
      if (column === 4) return {getValues: function() { return [["kc2-085"]]; }};
      if (column === 13) return {
        getValues: function() { return [[""]]; },
        getDisplayValues: function() { return [[""]]; },
        getFormulas: function() { return [[""]]; },
        getRichTextValues: function() { return [[__emptyRichText]]; },
        setFormula: function(formula) { __formulaWritten = formula; return this; }
      };
      if (column === 14) return {
        getValues: function() { return [["07.09.2026"]]; },
        setValue: function() { return this; },
        clearContent: function() { return this; }
      };
      throw new Error("Unexpected column " + column);
    }
  };
  var __otherColumnEdit = {
    range: {
      getSheet: function() { return __rowSheet; },
      getRow: function() { return 2174; },
      getLastRow: function() { return 2174; },
      getColumn: function() { return 12; },
      getLastColumn: function() { return 12; }
    }
  };
  getRequiredColumns_ = function() {
    return {date: 1, vehicle: 4, teleoperator: 5, logviewer: 13, logCheckDate: 14};
  };
  handleLogviewerEdit_(__otherColumnEdit);
`, context);
assert.ok(evaluate('__formulaWritten.includes("rovers_regexp=kc2-085")'));

console.log(
  "Smoke-тесты пройдены: даты, ссылки, stale-нарушения и динамический статус."
);
