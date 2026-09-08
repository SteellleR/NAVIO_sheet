/**
 * DropdownSync.gs — универсальная синхронизация нативных шторок.
 *
 * По метаданным Google Sheets находит ровно одну таблицу и целевой столбец на
 * каждом листе, собирает updateTable-запросы и полностью заменяет ONE_OF_LIST.
 * За счёт полной замены поддерживает и каскадное добавление, и удаление вариантов.
 */
function syncNativeTableDropdown_(options) {
  const values = options.values || [];
  if (!values.length) {
    throw new Error("Нельзя синхронизировать пустой список вариантов.");
  }

  const spreadsheet = options.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const metadata = fetchSpreadsheetTableMetadata_(spreadsheet.getId());
  const requests = buildNativeTableDropdownRequests_(metadata, {
    targetSheets: options.targetSheets,
    targetHeader: options.targetHeader,
    values: values
  });

  sendSheetsBatchUpdate_(spreadsheet.getId(), requests);

  return {
    values: values.length,
    tables: requests.length
  };
}


/** Строит updateTable-запросы и строго проверяет структуру листов. */
function buildNativeTableDropdownRequests_(metadata, options) {
  const requests = [];
  const normalizedHeader = normalizeText_(options.targetHeader);

  options.targetSheets.forEach(function(sheetName) {
    const matchingSheets = (metadata.sheets || []).filter(function(sheet) {
      return sheet.properties && sheet.properties.title === sheetName;
    });

    if (matchingSheets.length !== 1) {
      throw new Error(
        "Для листа «" + sheetName + "» найдено листов в API: " +
        matchingSheets.length + "."
      );
    }

    const matches = [];
    (matchingSheets[0].tables || []).forEach(function(table) {
      const columns = table.columnProperties || [];
      const indexes = [];

      columns.forEach(function(column, index) {
        if (normalizeText_(column.columnName) === normalizedHeader) {
          indexes.push(index);
        }
      });

      if (indexes.length > 1) {
        throw new Error(
          "В таблице " + table.tableId + " листа «" + sheetName +
          "» найдено несколько столбцов «" + options.targetHeader + "»."
        );
      }

      if (indexes.length === 1) {
        matches.push({table: table, columnIndex: indexes[0]});
      }
    });

    if (matches.length !== 1) {
      throw new Error(
        "На листе «" + sheetName + "» найдено подходящих нативных таблиц " +
        "со столбцом «" + options.targetHeader + "»: " + matches.length + "."
      );
    }

    const match = matches[0];
    const updatedColumns = match.table.columnProperties.map(function(column, index) {
      if (index !== match.columnIndex) return column;

      const updatedColumn = Object.assign({}, column);
      updatedColumn.columnType = "DROPDOWN";
      updatedColumn.dataValidationRule = {
        condition: {
          type: "ONE_OF_LIST",
          values: options.values.map(function(value) {
            return {userEnteredValue: value};
          })
        }
      };
      return updatedColumn;
    });

    requests.push({
      updateTable: {
        table: {
          tableId: match.table.tableId,
          columnProperties: updatedColumns
        },
        fields: "columnProperties"
      }
    });
  });

  return requests;
}
