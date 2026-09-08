/**
 * EntryPoints.gs - простые точки входа Google Sheets.
 *
 * onOpen восстанавливает служебное форматирование без тяжёлого полного
 * сканирования. onEdit последовательно обрабатывает дату события, «Сегодня»,
 * Logviewer и только затронутые строки отчёта; полный контроль выполняет таймер.
 * Здесь нет бизнес-логики: функции только передают событие профильным модулям.
 */
function onOpen(event) {
  addProjectMenu_();
  setupLogCheckDateFormatting();
}


/** Простая точка входа после ручного изменения рабочих ячеек. */
function onEdit(event) {
  if (!event || !event.range) return;

  // Порядок принципиален: сначала фиксируется дата, затем меняется Logviewer,
  // а отчёт читает уже итоговое состояние строки.
  handleLogCheckDateEdit_(event);
  handleLogviewerEdit_(event);
  handleViolatorsEdit_(event);
}
