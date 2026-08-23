import js from "@eslint/js";
import astro from "eslint-plugin-astro";
import globals from "globals";

export default [
  {
    ignores: [".astro/**", "dist/**", "node_modules/**", "public/**", "src/Code.local.gs"]
  },
  js.configs.recommended,
  ...astro.configs.recommended,
  {
    files: ["**/*.astro"],
    languageOptions: {
      globals: globals.browser
    }
  },
  {
    files: ["**/*.mjs"],
    languageOptions: {
      globals: globals.node
    }
  },
  {
    files: ["src/**/*.gs"],
    languageOptions: {
      sourceType: "script",
      globals: {
        CalendarApp: "readonly",
        ContentService: "readonly",
        HtmlService: "readonly",
        MailApp: "readonly",
        PropertiesService: "readonly",
        ScriptApp: "readonly",
        Session: "readonly",
        SpreadsheetApp: "readonly",
        Utilities: "readonly"
      }
    },
    rules: {
      "no-unused-vars": "off"
    }
  }
];
