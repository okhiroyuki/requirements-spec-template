import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.join(__dirname, '..', '..');

// Mirrors the Apps Script project's file list. Within one Apps Script
// project every file shares a single global scope, so load order does not
// affect which functions can call which — this list just documents the
// project's files in one place.
const GAS_SOURCE_FILES = [
  'template-setup.gs',
  'validation.gs',
  'template-sheets.gs',
  'ids.gs',
  'menu.gs',
  'markdown-export.gs',
];

/**
 * Loads the project's .gs files, unmodified, into a single vm context
 * stubbed with just enough of the Apps Script globals for their pure-logic
 * functions to run. This tests the actual shipped files instead of a copy,
 * the same way Apps Script itself merges multiple files into one global
 * scope.
 */
export function loadGasContext() {
  const sandbox = {
    Utilities: {
      formatDate(date, timeZone, format) {
        if (format !== 'yyyy-MM-dd') {
          throw new Error(`gasSandbox stub: unsupported Utilities.formatDate format "${format}"`);
        }
        const y = date.getUTCFullYear();
        const m = String(date.getUTCMonth() + 1).padStart(2, '0');
        const d = String(date.getUTCDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      },
    },
    Session: {
      getScriptTimeZone() {
        return 'Asia/Tokyo';
      },
    },
    Logger: {
      log() {},
    },
  };
  vm.createContext(sandbox);
  for (const fileName of GAS_SOURCE_FILES) {
    const source = fs.readFileSync(path.join(PROJECT_ROOT, fileName), 'utf8');
    vm.runInContext(source, sandbox, { filename: fileName });
  }

  // Built-ins of a vm context (Date, Array, ...) live in that context's own
  // realm and are not exposed as ordinary properties on the sandbox object.
  // Functions defined in these .gs files resolve `Date` against that realm,
  // so tests must construct dates via this same constructor for
  // `instanceof Date` checks inside the script to succeed.
  sandbox.Date = vm.runInContext('Date', sandbox);

  return sandbox;
}
