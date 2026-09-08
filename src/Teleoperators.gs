/**
 * Teleoperators.gs — справочник и шторки телеоператоров.
 *
 * Читает и проверяет логины на листе «Телеоператоры», устраняет дубли и
 * синхронизирует список с нативными таблицами рабочих листов. Добавления и
 * удаления каскадно отражаются в шторках и в отчёте «Нарушители».
 */
function syncTeleoperatorDropdowns(force) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    return {updated: false, reason: "locked"};
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const entries = getTeleoperatorEntries_(spreadsheet);
    const operators = entries.map(function(entry) {
      return entry.login;
    });
    const hash = buildHash_(operators);
    const properties = PropertiesService.getDocumentProperties();

    if (!force && properties.getProperty(
      APP_CONFIG_.properties.teleoperatorsHash
    ) === hash) {
      return {
        updated: false,
        reason: "unchanged",
        operators: operators.length
      };
    }

    const syncResult = syncNativeTableDropdown_({
      spreadsheet: spreadsheet,
      targetSheets: getActiveWorkingSheets_(),
      targetHeader: APP_CONFIG_.headers.teleoperator,
      values: operators
    });

    properties.setProperty(APP_CONFIG_.properties.teleoperatorsHash, hash);

    return {
      updated: true,
      operators: operators.length,
      tables: syncResult.tables
    };
  } finally {
    lock.releaseLock();
  }
}


/** Читает справочник, поддерживая и новый «Логин», и старый заголовок. */
function getTeleoperatorEntries_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(
    APP_CONFIG_.references.teleoperatorsSheet
  );

  if (!sheet) {
    throw new Error(
      "Не найден лист «" + APP_CONFIG_.references.teleoperatorsSheet + "»."
    );
  }

  const columns = findHeaderColumns_(sheet, {
    login: {headers: APP_CONFIG_.references.teleoperatorLoginHeaders},
    fullName: {
      headers: APP_CONFIG_.references.teleoperatorNameHeader,
      optional: true
    }
  });
  const loginColumn = columns.login;
  const nameColumn = columns.fullName;

  const lastRow = sheet.getLastRow();
  if (lastRow < APP_CONFIG_.firstDataRow) {
    throw new Error("В справочнике телеоператоров нет логинов.");
  }

  const rowCount = lastRow - APP_CONFIG_.firstDataRow + 1;
  const logins = sheet.getRange(
    APP_CONFIG_.firstDataRow,
    loginColumn,
    rowCount,
    1
  ).getDisplayValues();
  const names = nameColumn
    ? sheet.getRange(
      APP_CONFIG_.firstDataRow,
      nameColumn,
      rowCount,
      1
    ).getDisplayValues()
    : null;

  const entries = [];
  const byLogin = Object.create(null);

  logins.forEach(function(row, index) {
    const login = String(row[0] == null ? "" : row[0]).trim();
    if (!login) return;

    const fullName = names
      ? String(names[index][0] == null ? "" : names[index][0]).trim()
      : "";
    const key = normalizeText_(login);

    if (byLogin[key]) {
      const knownName = normalizeText_(byLogin[key].fullName);
      const nextName = normalizeText_(fullName);
      if (knownName && nextName && knownName !== nextName) {
        throw new Error(
          "Логину «" + login + "» соответствуют разные ФИО: «" +
          byLogin[key].fullName + "» и «" + fullName + "»."
        );
      }
      if (!byLogin[key].fullName && fullName) byLogin[key].fullName = fullName;
      return;
    }

    const entry = {
      login: login,
      fullName: fullName,
      row: APP_CONFIG_.firstDataRow + index
    };
    byLogin[key] = entry;
    entries.push(entry);
  });

  if (!entries.length) {
    throw new Error("В справочнике телеоператоров нет логинов.");
  }

  return entries;
}


/** Вызывается общим установочным onEdit-триггером. */
function handleTeleoperatorsReferenceEdit_(event) {
  if (!event || !event.range) return false;

  const sheet = event.range.getSheet();
  if (sheet.getName() !== APP_CONFIG_.references.teleoperatorsSheet) {
    return false;
  }
  if (event.range.getRow() <= APP_CONFIG_.headerRow) {
    syncTeleoperatorDropdowns(true);
    syncViolators();
    return true;
  }

  const columns = findHeaderColumns_(sheet, {
    login: {headers: APP_CONFIG_.references.teleoperatorLoginHeaders},
    fullName: {
      headers: APP_CONFIG_.references.teleoperatorNameHeader,
      optional: true
    }
  });
  const loginColumn = columns.login;
  const nameColumn = columns.fullName;

  if (
    !rangeTouchesColumn_(event.range, loginColumn) &&
    (!nameColumn || !rangeTouchesColumn_(event.range, nameColumn))
  ) {
    return false;
  }

  syncTeleoperatorDropdowns(true);
  syncViolators();
  return true;
}
