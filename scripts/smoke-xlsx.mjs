import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';

assert.equal(XLSX.version, '0.20.3', `Unexpected SheetJS version: ${XLSX.version}`);
const workbook = XLSX.read('barcode,quantity\n123,2', { type: 'string' });
assert.equal(workbook.Sheets.Sheet1?.A2?.v, 123, 'CSV parser smoke test failed');
