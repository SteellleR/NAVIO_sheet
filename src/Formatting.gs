/**
 * Formatting.gs — служебное форматирование «Дата проверки лога».
 *
 * Поддерживает зелёное оформление статуса «заполнен» и при первой установке
 * удаляет старое красное правило пустой даты. Посторонние правила листов
 * сохраняются, а служебное правило обновляется идемпотентно.
 */
function setupLogCheckDateFormatting() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  getActiveWorkingSheets_().forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    const columns = getRequiredColumns_(sheet);
    if (columns) ensureLogCheckDateFormatting_(sheet, columns);
  });
}


/** Оставляет одно зелёное правило проекта и сохраняет посторонние правила. */
function ensureLogCheckDateFormatting_(sheet, columns) {
  const formulas = buildLogFormattingFormulas_(columns);
  const targetRange = sheet.getRange(
    APP_CONFIG_.firstDataRow,
    columns.logCheckDate,
    sheet.getMaxRows() - APP_CONFIG_.firstDataRow + 1,
    1
  );
  const currentRules = sheet.getConditionalFormatRules();
  const ownedRules = currentRules.filter(function(rule) {
    return isOwnedLogFormattingRule_(rule, targetRange, formulas.known);
  });

  const hasExactGreen = ownedRules.some(function(rule) {
    return conditionalRuleMatches_(
      rule,
      targetRange,
      formulas.filled,
      APP_CONFIG_.formatting.filledBackground,
      APP_CONFIG_.formatting.filledText
    );
  });

  if (ownedRules.length === 1 && hasExactGreen) return;

  const unrelatedRules = currentRules.filter(function(rule) {
    return !isOwnedLogFormattingRule_(rule, targetRange, formulas.known);
  });
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulas.filled)
    .setBackground(APP_CONFIG_.formatting.filledBackground)
    .setFontColor(APP_CONFIG_.formatting.filledText)
    .setRanges([targetRange])
    .build();

  sheet.setConditionalFormatRules([greenRule].concat(unrelatedRules));
}


function buildLogFormattingFormulas_(columns) {
  const logColumn = columnToLetter_(columns.logviewer);
  const statusColumn = columnToLetter_(columns.logCheckDate);
  const row = APP_CONFIG_.firstDataRow;
  const label = APP_CONFIG_.logviewer.addLogLabel;
  const filled = APP_CONFIG_.logviewer.filledLabel;
  const filledFormula = '=$' + statusColumn + row + '="' + filled + '"';
  // Эти формулы больше не создаются. Они перечислены только для удаления
  // правил, оставшихся после старых версий проекта.
  const obsoleteEmptyDateRules = [
    '=($' + logColumn + row + '="' + label + '")*($' +
      statusColumn + row + '="")',
    '=AND($' + logColumn + row + '="' + label + '",$' +
      statusColumn + row + '="")',
    '=AND($' + logColumn + row + '="' + label + '";$' +
      statusColumn + row + '="")'
  ];
  const known = [
  ].concat(obsoleteEmptyDateRules).concat([
    filledFormula,
    '=($' + statusColumn + row + '="' + filled + '")'
  ]).map(normalizeFormula_);

  return {filled: filledFormula, known: known};
}


function isOwnedLogFormattingRule_(rule, targetRange, knownFormulas) {
  const condition = rule.getBooleanCondition();
  if (!condition) return false;
  if (
    condition.getCriteriaType() !==
    SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA
  ) {
    return false;
  }

  const formula = normalizeFormula_(condition.getCriteriaValues()[0]);
  if (!knownFormulas.includes(formula)) return false;

  const ranges = rule.getRanges();
  if (ranges.length !== 1) return false;

  return ranges.some(function(range) {
    return (
      range.getSheet().getSheetId() === targetRange.getSheet().getSheetId() &&
      range.getRow() === APP_CONFIG_.firstDataRow &&
      range.getColumn() === targetRange.getColumn() &&
      range.getNumColumns() === 1
    );
  });
}


function conditionalRuleMatches_(rule, targetRange, formula, background, font) {
  const condition = rule.getBooleanCondition();
  if (!condition) return false;
  if (normalizeFormula_(condition.getCriteriaValues()[0]) !==
    normalizeFormula_(formula)) {
    return false;
  }

  const ranges = rule.getRanges();
  if (ranges.length !== 1) return false;
  const range = ranges[0];
  const sameRange =
    range.getSheet().getSheetId() === targetRange.getSheet().getSheetId() &&
    range.getRow() === targetRange.getRow() &&
    range.getColumn() === targetRange.getColumn() &&
    range.getNumRows() === targetRange.getNumRows() &&
    range.getNumColumns() === targetRange.getNumColumns();
  if (!sameRange) return false;

  if (normalizeColor_(condition.getBackground()) !== normalizeColor_(background)) {
    return false;
  }

  return font == null ||
    normalizeColor_(condition.getFontColor()) === normalizeColor_(font);
}


function normalizeColor_(color) {
  return String(color == null ? "" : color).trim().toLowerCase();
}
