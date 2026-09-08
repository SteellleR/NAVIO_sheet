/**
 * SheetsApi.gs - минимальный HTTP-клиент Google Sheets API v4.
 *
 * Получает метаданные нативных таблиц и отправляет пакетные updateTable-
 * запросы с OAuth-токеном текущего Apps Script. Ошибки API преобразуются в
 * исключения с HTTP-кодом и исходным ответом для диагностики.
 */
function fetchSpreadsheetTableMetadata_(spreadsheetId) {
  const fields =
    "sheets(properties(sheetId,title)," +
    "tables(tableId,name,range,columnProperties))";
  const url = APP_CONFIG_.sheetsApiUrl + encodeURIComponent(spreadsheetId) +
    "?includeGridData=false&fields=" + encodeURIComponent(fields);

  return fetchSheetsJson_(url, {method: "get"});
}


/** Отправляет пакет обновлений в Google Sheets API. */
function sendSheetsBatchUpdate_(spreadsheetId, requests) {
  if (!requests.length) return {};

  const url = APP_CONFIG_.sheetsApiUrl + encodeURIComponent(spreadsheetId) +
    ":batchUpdate";

  return fetchSheetsJson_(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({requests: requests})
  });
}


/** Выполняет авторизованный запрос и приводит ошибки API к читаемому виду. */
function fetchSheetsJson_(url, options) {
  const requestOptions = Object.assign({}, options, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  const response = UrlFetchApp.fetch(url, requestOptions);
  const status = response.getResponseCode();
  const content = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error("Ошибка Google Sheets API " + status + ": " + content);
  }

  return content ? JSON.parse(content) : {};
}
