/**
 * check-structure.mjs - проверка состава модулей и глобального пространства.
 * Контролирует обязательные файлы, заголовочные комментарии и дубли имён.
 */
import fs from "node:fs";
import path from "node:path";

const expectedFiles = [
  "appsscript.json",
  "Constants.gs",
  "EntryPoints.gs",
  "Setup.gs",
  "SheetTracking.gs",
  "SheetTrackingDialog.html",
  "Documentation.gs",
  "DocumentationDialog.html",
  "Logviewer.gs",
  "LogCheckDate.gs",
  "Teleoperators.gs",
  "Violators.gs",
  "Vehicles.gs",
  "DropdownSync.gs",
  "Formatting.gs",
  "SheetsApi.gs",
  "Utils.gs"
];
const sourceDirectory = path.resolve("src");
const failures = [];
const typographyFiles = [
  "README.md",
  "eslint.config.mjs",
  "jsconfig.json",
  "package.json",
  "scripts/check-structure.mjs",
  "tests/smoke.mjs"
].concat(expectedFiles.map((fileName) => path.join("src", fileName)));

for (const fileName of typographyFiles) {
  const source = fs.readFileSync(path.resolve(fileName), "utf8");
  if (/[\u2013\u2014]/u.test(source)) {
    failures.push(`В ${fileName} найдено запрещённое длинное тире`);
  }
}

for (const fileName of expectedFiles) {
  if (!fs.existsSync(path.join(sourceDirectory, fileName))) {
    failures.push(`Отсутствует src/${fileName}`);
  }
}

const declarations = new Map();
for (const fileName of expectedFiles.filter((name) => name.endsWith(".gs"))) {
  const fullPath = path.join(sourceDirectory, fileName);
  if (!fs.existsSync(fullPath)) continue;
  const source = fs.readFileSync(fullPath, "utf8");

  if (!source.startsWith("/**") || !source.slice(0, 400).includes(fileName)) {
    failures.push(
      `В начале src/${fileName} ожидается комментарий с назначением файла`
    );
  }

  for (const match of source.matchAll(
    /^(?:function\s+|const\s+|let\s+|var\s+)([A-Za-z_$][\w$]*)/gm
  )) {
    if (declarations.has(match[1])) {
      failures.push(
        `Глобальное имя ${match[1]} повторяется в ${declarations.get(match[1])} и ${fileName}`
      );
    } else {
      declarations.set(match[1], fileName);
    }
  }
}

const entrySource = fs.readFileSync(
  path.join(sourceDirectory, "EntryPoints.gs"),
  "utf8"
);
for (const entryName of ["onOpen", "onEdit"]) {
  const count = [...entrySource.matchAll(
    new RegExp(`^function\\s+${entryName}\\s*\\(`, "gm")
  )].length;
  if (count !== 1) failures.push(`Ожидалась одна функция ${entryName}, найдено: ${count}`);
}

if (failures.length) {
  console.error(failures.join("\n"));
  process.exitCode = 1;
} else {
  console.log(`Структура корректна: ${expectedFiles.length} файлов, конфликтов имён нет.`);
}
