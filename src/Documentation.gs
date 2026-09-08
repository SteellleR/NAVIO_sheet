/**
 * Documentation.gs — встроенная документация проекта.
 *
 * Открывает два режима HTML-справки из вложенного меню NAVIO: короткую
 * пользовательскую инструкцию и подробную техническую карту для разработчика.
 * Контент и интерфейс хранятся рядом с кодом и обновляются вместе с проектом.
 * Отредактированные ИИ-промпты сохраняются для конкретной таблицы в
 * Document Properties безопасными частями и могут быть сброшены к встроенным.
 */
function showUserDocumentation() {
  showDocumentationDialog_("user", "Документация · Для пользователя");
}


/** Открывает подробную техническую документацию. */
function showDeveloperDocumentation() {
  showDocumentationDialog_("developer", "Документация · Для разработчика");
}


/** Создаёт большое адаптивное окно документации в нужном режиме. */
function showDocumentationDialog_(mode, title) {
  const allowedModes = ["user", "developer"];
  if (!allowedModes.includes(mode)) {
    throw new Error("Неизвестный раздел документации: " + mode);
  }

  const template = HtmlService.createTemplateFromFile("DocumentationDialog");
  template.documentationMode = mode;
  const html = template.evaluate()
    .setWidth(1180)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, title);
}


/** Возвращает пользовательские версии ИИ-промптов или null для встроенных. */
function getDocumentationPromptOverrides() {
  return {
    html: readDocumentationPrompt_("html"),
    backend: readDocumentationPrompt_("backend")
  };
}


/** Сохраняет отредактированный ИИ-промпт для текущей таблицы. */
function saveDocumentationPrompt(kind, prompt) {
  const value = String(prompt == null ? "" : prompt).trim();
  if (!value) throw new Error("Промпт не может быть пустым.");
  if (value.length > 40000) {
    throw new Error("Промпт слишком большой. Максимум — 40 000 символов.");
  }

  const prefix = getDocumentationPromptPrefix_(kind);
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Документация сейчас обновляется. Попробуйте ещё раз.");
  }

  try {
    writeChunkedDocumentationProperty_(prefix, value);
  } finally {
    lock.releaseLock();
  }

  return {kind: kind, prompt: value, customized: true};
}


/** Удаляет пользовательскую версию и возвращает встроенный промпт. */
function resetDocumentationPrompt(kind) {
  const prefix = getDocumentationPromptPrefix_(kind);
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Документация сейчас обновляется. Попробуйте ещё раз.");
  }

  try {
    deleteChunkedDocumentationProperty_(prefix);
  } finally {
    lock.releaseLock();
  }

  return {kind: kind, customized: false};
}


/** Сопоставляет публичный тип промпта с безопасным ключом PropertiesService. */
function getDocumentationPromptPrefix_(kind) {
  const keys = {
    html: APP_CONFIG_.properties.documentationPromptHtml,
    backend: APP_CONFIG_.properties.documentationPromptBackend
  };
  if (!Object.prototype.hasOwnProperty.call(keys, kind)) {
    throw new Error("Неизвестный тип промпта: " + kind);
  }
  return keys[kind];
}


/** Читает длинный текст, разбитый на безопасные для Properties части. */
function readDocumentationPrompt_(kind) {
  const prefix = getDocumentationPromptPrefix_(kind);
  const properties = PropertiesService.getDocumentProperties();
  const count = Number(properties.getProperty(prefix + "_parts"));
  if (!Number.isInteger(count) || count < 1 || count > 20) return null;

  const parts = [];
  for (let index = 0; index < count; index++) {
    const part = properties.getProperty(prefix + "_" + index);
    if (part == null) return null;
    parts.push(part);
  }
  return parts.join("");
}


/** Записывает длинный русский текст частями, не упираясь в лимит одного value. */
function writeChunkedDocumentationProperty_(prefix, value) {
  const properties = PropertiesService.getDocumentProperties();
  deleteChunkedDocumentationProperty_(prefix);

  // 2400 символов безопасно укладываются в 9 КБ даже для кириллицы UTF-8.
  const chunkSize = 2400;
  const chunks = [];
  for (let offset = 0; offset < value.length; offset += chunkSize) {
    chunks.push(value.slice(offset, offset + chunkSize));
  }

  const data = {};
  data[prefix + "_parts"] = String(chunks.length);
  chunks.forEach(function(chunk, index) {
    data[prefix + "_" + index] = chunk;
  });
  properties.setProperties(data, false);
}


/** Удаляет все части сохранённого промпта указанного типа. */
function deleteChunkedDocumentationProperty_(prefix) {
  const properties = PropertiesService.getDocumentProperties();
  const all = properties.getProperties();
  Object.keys(all).forEach(function(key) {
    if (key === prefix + "_parts" || key.indexOf(prefix + "_") === 0) {
      properties.deleteProperty(key);
    }
  });
}
