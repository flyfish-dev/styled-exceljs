import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { resolve } from 'node:path';

const require = createRequire(import.meta.url);
const selected = process.env.XLSX_TEST_MODULE;
const XLSX = selected
  ? await import(pathToFileURL(resolve(selected)).href).then(m => m.default || m)
  : require('../xlsx.js');
const rows = Array.from({ length: 800 }, (_, index) => ['Row ' + index, index * 7]);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), 'Data');
const original = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'biff8' }));
const directory = (original.readUInt32LE(48) + 1) * 512;
const unused = Buffer.from(original);
// Remove the synthetic writer marker, leaving only the unchanged, FAT-backed
// Workbook stream. The allocated MiniFAT is now unused by every live stream.
unused[directory + 128 + 66] = 0;
unused.writeUInt32LE(0, directory + 128 + 120);
unused.writeInt32LE(-2, directory + 116);
unused.writeUInt32LE(0, directory + 120);
const baseline = XLSX.CFB.read(original, { type: 'buffer' }).FileIndex.find(x => x.name === 'Workbook').content;
for (const start of [-1, -2, 0x7fffffff]) {
  const input = Buffer.from(unused);
  input.writeInt32LE(start, 60);
  const parsed = XLSX.read(input, { type: 'buffer' });
  assert.deepEqual(parsed.SheetNames, ['Data']);
  assert.deepEqual(XLSX.utils.sheet_to_json(parsed.Sheets.Data, { header: 1 }), rows);
  const stream = XLSX.CFB.read(input, { type: 'buffer' }).FileIndex.find(x => x.name === 'Workbook').content;
  assert.deepEqual(Buffer.from(stream), Buffer.from(baseline), 'Workbook payload changed');
}
// A required mini stream must never silently become empty to avoid the crash.
const missing = Buffer.from(original);
missing.writeInt32LE(-1, 60);
assert.throws(() => XLSX.CFB.read(missing, { type: 'buffer' }), /CFB MiniFAT chain is missing/);
const missingRoot = Buffer.from(original);
missingRoot.writeInt32LE(0x7fffffff, directory + 116);
assert.throws(() => XLSX.CFB.read(missingRoot, { type: 'buffer' }), /CFB root mini stream is missing/);
assert.deepEqual(XLSX.utils.sheet_to_json(XLSX.read(original, { type: 'buffer' }).Sheets.Data, { header: 1 }), rows);
console.log('CFB unused MiniFAT recovery preserves all 1600 cells and exact Workbook bytes; required missing streams reject.');
