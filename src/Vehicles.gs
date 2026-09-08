/**
 * Vehicles.gs — справочник ВАТС и безопасная миграция значений.
 *
 * Читает номера и статусы со справочного листа, отправляет в рабочие шторки
 * только активные ВАТС и проверяет допустимость статусов. Также содержит
 * предварительный просмотр и явный запуск миграции старых названий в номера.
 */
function syncVehicleDropdowns(force) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    return {updated: false, reason: "locked"};
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const entries = getVehicleEntries_(spreadsheet);
    const activeStatus = normalizeText_(
      APP_CONFIG_.references.activeVehicleStatus
    );
    const numbers = [];
    const seen = Object.create(null);

    entries.forEach(function(entry) {
      if (normalizeText_(entry.status) !== activeStatus) return;
      if (!entry.number) {
        throw new Error(
          "У активного ВАТС в строке " + entry.row + " не указан номер."
        );
      }

      const key = normalizeText_(entry.number);
      if (seen[key]) return;
      seen[key] = true;
      numbers.push(entry.number);
    });

    if (!numbers.length) {
      throw new Error("В справочнике ВАТС нет активных номеров.");
    }

    const hash = buildHash_(numbers);
    const properties = PropertiesService.getDocumentProperties();
    if (!force && properties.getProperty(
      APP_CONFIG_.properties.vehiclesHash
    ) === hash) {
      return {
        updated: false,
        reason: "unchanged",
        vehicles: numbers.length
      };
    }

    const syncResult = syncNativeTableDropdown_({
      spreadsheet: spreadsheet,
      targetSheets: getActiveWorkingSheets_(),
      targetHeader: APP_CONFIG_.headers.vehicle,
      values: numbers
    });

    properties.setProperty(APP_CONFIG_.properties.vehiclesHash, hash);

    return {
      updated: true,
      vehicles: numbers.length,
      tables: syncResult.tables
    };
  } finally {
    lock.releaseLock();
  }
}


/** Читает и валидирует справочник ВАТС по заголовкам. */
function getVehicleEntries_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(APP_CONFIG_.references.vehiclesSheet);
  if (!sheet) {
    throw new Error(
      "Не найден лист «" + APP_CONFIG_.references.vehiclesSheet + "»."
    );
  }

  const columns = findHeaderColumns_(sheet, {
    name: {headers: APP_CONFIG_.references.vehicleNameHeader},
    number: {headers: APP_CONFIG_.references.vehicleNumberHeader},
    status: {headers: APP_CONFIG_.references.vehicleStatusHeader}
  });
  const nameColumn = columns.name;
  const numberColumn = columns.number;
  const statusColumn = columns.status;

  const lastRow = sheet.getLastRow();
  if (lastRow < APP_CONFIG_.firstDataRow) return [];

  const rowCount = lastRow - APP_CONFIG_.firstDataRow + 1;
  const names = sheet.getRange(
    APP_CONFIG_.firstDataRow,
    nameColumn,
    rowCount,
    1
  ).getDisplayValues();
  const numbers = sheet.getRange(
    APP_CONFIG_.firstDataRow,
    numberColumn,
    rowCount,
    1
  ).getDisplayValues();
  const statuses = sheet.getRange(
    APP_CONFIG_.firstDataRow,
    statusColumn,
    rowCount,
    1
  ).getDisplayValues();
  const allowedStatuses = [
    normalizeText_(APP_CONFIG_.references.activeVehicleStatus),
    normalizeText_(APP_CONFIG_.references.inactiveVehicleStatus)
  ];
  const entries = [];

  for (let index = 0; index < rowCount; index++) {
    const name = String(names[index][0] == null ? "" : names[index][0]).trim();
    const number = String(numbers[index][0] == null ? "" : numbers[index][0]).trim();
    const status = String(statuses[index][0] == null ? "" : statuses[index][0]).trim();

    if (!name && !number && !status) continue;
    if (status && !allowedStatuses.includes(normalizeText_(status))) {
      throw new Error(
        "Неизвестный статус ВАТС «" + status + "» в строке " +
        (APP_CONFIG_.firstDataRow + index) + "."
      );
    }

    entries.push({
      id: APP_CONFIG_.firstDataRow + index,
      row: APP_CONFIG_.firstDataRow + index,
      name: name,
      number: number,
      status: status
    });
  }

  return entries;
}


/** Вызывается общим установочным onEdit-триггером. */
function handleVehiclesReferenceEdit_(event) {
  if (!event || !event.range) return false;

  const sheet = event.range.getSheet();
  if (sheet.getName() !== APP_CONFIG_.references.vehiclesSheet) return false;
  if (event.range.getRow() <= APP_CONFIG_.headerRow) {
    syncVehicleDropdowns(false);
    return true;
  }

  const columns = findHeaderColumns_(sheet, {
    number: {headers: APP_CONFIG_.references.vehicleNumberHeader},
    status: {headers: APP_CONFIG_.references.vehicleStatusHeader}
  });
  const numberColumn = columns.number;
  const statusColumn = columns.status;

  if (
    !rangeTouchesColumn_(event.range, numberColumn) &&
    !rangeTouchesColumn_(event.range, statusColumn)
  ) {
    return false;
  }

  syncVehicleDropdowns(false);
  return true;
}


/**
 * Безопасная ручная миграция «Название» → «Номер ВАТС».
 * Без аргумента работает только как предварительная проверка.
 * Для записи следует явно вызвать migrateVehicleNamesToNumbers(false).
 */
function migrateVehicleNamesToNumbers(dryRun) {
  const previewOnly = dryRun !== false;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const entries = getVehicleEntries_(spreadsheet);
  const numbers = Object.create(null);
  const mapping = Object.create(null);

  entries.forEach(function(entry) {
    if (entry.number) numbers[normalizeText_(entry.number)] = true;
    if (!entry.name || !entry.number) return;

    const key = normalizeText_(entry.name);
    if (!mapping[key]) mapping[key] = [];
    if (!mapping[key].includes(entry.number)) mapping[key].push(entry.number);
  });

  const report = {
    dryRun: previewOnly,
    replacements: 0,
    sheets: [],
    problems: []
  };

  getActiveWorkingSheets_().forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    const vehicleColumn = findUniqueHeaderColumn_(
      sheet,
      APP_CONFIG_.headers.vehicle,
      false
    );
    const lastRow = sheet.getLastRow();
    if (lastRow < APP_CONFIG_.firstDataRow) return;

    const rowCount = lastRow - APP_CONFIG_.firstDataRow + 1;
    const range = sheet.getRange(
      APP_CONFIG_.firstDataRow,
      vehicleColumn,
      rowCount,
      1
    );
    const values = range.getDisplayValues();
    const formulas = range.getFormulas();
    let sheetReplacements = 0;

    values.forEach(function(row, index) {
      const source = String(row[0] == null ? "" : row[0]).trim();
      if (!source || numbers[normalizeText_(source)]) return;

      const targetNumbers = mapping[normalizeText_(source)];
      const rowNumber = APP_CONFIG_.firstDataRow + index;

      if (!targetNumbers) {
        report.problems.push({
          sheet: sheetName,
          row: rowNumber,
          value: source,
          reason: "not_found"
        });
        return;
      }

      if (targetNumbers.length !== 1) {
        report.problems.push({
          sheet: sheetName,
          row: rowNumber,
          value: source,
          reason: "ambiguous",
          candidates: targetNumbers
        });
        return;
      }

      if (formulas[index][0]) {
        report.problems.push({
          sheet: sheetName,
          row: rowNumber,
          value: source,
          reason: "formula"
        });
        return;
      }

      sheetReplacements++;
      report.replacements++;
      if (!previewOnly) {
        sheet.getRange(rowNumber, vehicleColumn).setValue(targetNumbers[0]);
      }
    });

    report.sheets.push({sheet: sheetName, replacements: sheetReplacements});
  });

  if (!previewOnly && report.replacements) fillMissingLinks();

  console.log(JSON.stringify(report));
  return report;
}


/** Удобная точка запуска предварительной проверки из редактора Apps Script. */
function previewVehicleNamesToNumbersMigration() {
  return migrateVehicleNamesToNumbers(true);
}


/** Явная точка запуска записи после проверки отчёта preview-функции. */
function applyVehicleNamesToNumbersMigration() {
  return migrateVehicleNamesToNumbers(false);
}
