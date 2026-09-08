/**
 * eslint.config.mjs - правила статического анализа локальных .gs-файлов.
 * Собирает глобальные объявления Apps Script и проверяет ошибки JavaScript.
 */
import fs from "node:fs";
import path from "node:path";
import js from "@eslint/js";
import globals from "globals";

const sourceDirectory = path.resolve("src");
const sharedNames = {};

for (const fileName of fs.readdirSync(sourceDirectory)) {
  if (!fileName.endsWith(".gs")) continue;
  const source = fs.readFileSync(path.join(sourceDirectory, fileName), "utf8");

  for (const match of source.matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(/gm)) {
    sharedNames[match[1]] = "readonly";
  }
  for (const match of source.matchAll(/^(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=/gm)) {
    sharedNames[match[1]] = "readonly";
  }
}

const appsScriptGlobals = {
  console: "readonly",
  SpreadsheetApp: "readonly",
  HtmlService: "readonly",
  ScriptApp: "readonly",
  LockService: "readonly",
  PropertiesService: "readonly",
  Utilities: "readonly",
  UrlFetchApp: "readonly",
  Logger: "readonly"
};

export default [
  {
    ignores: ["node_modules/**"]
  },
  {
    files: ["src/**/*.gs"],
    ...js.configs.recommended,
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: "script",
      globals: {
        ...globals.es2021,
        ...appsScriptGlobals,
        ...sharedNames
      }
    },
    rules: {
      "no-unused-vars": "off",
      "no-undef": "error",
      "no-unreachable": "error",
      "no-dupe-keys": "error",
      "eqeqeq": ["error", "always", {"null": "ignore"}],
      "curly": ["error", "multi-line"],
      "semi": ["error", "always"]
    }
  }
];
