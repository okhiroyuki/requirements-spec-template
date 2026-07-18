import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const GAS_SOURCE_PATH = path.join(__dirname, '..', '..', 'create-spreadsheet.gs');

/**
 * Loads create-spreadsheet.gs, unmodified, into a vm context stubbed with
 * just enough of the Apps Script globals for its pure-logic functions to
 * run. This lets us test the actual shipped file instead of a copy.
 */
export function loadGasContext() {
  const source = fs.readFileSync(GAS_SOURCE_PATH, 'utf8');

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
  vm.runInContext(source, sandbox, { filename: 'create-spreadsheet.gs' });

  // Built-ins of a vm context (Date, Array, ...) live in that context's own
  // realm and are not exposed as ordinary properties on the sandbox object.
  // Functions defined in create-spreadsheet.gs resolve `Date` against that
  // realm, so tests must construct dates via this same constructor for
  // `instanceof Date` checks inside the script to succeed.
  sandbox.Date = vm.runInContext('Date', sandbox);

  return sandbox;
}
