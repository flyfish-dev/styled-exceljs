import assert from 'node:assert/strict'

import * as XLSXNamespace from '../xlsx.mjs'

const XLSX = XLSXNamespace.default || XLSXNamespace
const inheritedKey = '!autofilter'

const worksheet = XLSX.utils.aoa_to_sheet([['safe']])
const workbook = XLSX.utils.book_new()
XLSX.utils.book_append_sheet(workbook, worksheet, 'Safe')
const fods = XLSX.write(workbook, { type: 'string', bookType: 'fods' })
const databaseRange = [
  '<table:database-ranges>',
  '<table:database-range table:name="unsafe" table:target-range-address="__proto__.A1:A1"/>',
  '</table:database-ranges>',
].join('')
const malicious = fods.replace('</office:spreadsheet>', `${databaseRange}</office:spreadsheet>`)

delete Object.prototype[inheritedKey]
try {
  XLSX.read(malicious, { type: 'string' })
  assert.equal(
    Object.prototype.hasOwnProperty.call(Object.prototype, inheritedKey),
    false,
    'ODS parsing must not write !autofilter onto Object.prototype',
  )
} finally {
  delete Object.prototype[inheritedKey]
}

const protoWorkbook = XLSX.utils.book_new()
XLSX.utils.book_append_sheet(protoWorkbook, XLSX.utils.aoa_to_sheet([['prototype-safe']]), '__proto__')
assert.equal(Object.prototype.hasOwnProperty.call(protoWorkbook.Sheets, '__proto__'), true)
assert.equal(protoWorkbook.Sheets.__proto__.A1.v, 'prototype-safe')
const protoRoundTrip = XLSX.read(XLSX.write(protoWorkbook, { type: 'buffer', bookType: 'xlsx' }), {
  type: 'buffer',
})
assert.deepEqual(protoRoundTrip.SheetNames, ['__proto__'])
assert.equal(Object.prototype.hasOwnProperty.call(protoRoundTrip.Sheets, '__proto__'), true)
assert.equal(protoRoundTrip.Sheets.__proto__.A1.v, 'prototype-safe')

assert.throws(
  () => XLSX.utils.sheet_add_json(null, [{ safe: true }], { dense: true, origin: { r: '__proto__', c: 0 } }),
  /Invalid origin/,
)

const unsafeHtmlSheet = {
  A1: {
    t: 's',
    v: 'unsafe',
    h: '<img src=x onerror="globalThis.__styledExcelSentinel=1"><b>visible</b>',
    l: { Target: 'java\nscript:globalThis.__styledExcelSentinel=2' },
    s: { font: { name: "Font\\';color:red;/*", sz: '12;position:fixed' } },
  },
  '!ref': 'A1',
}
const unsafeHtml = XLSX.utils.sheet_to_html(unsafeHtmlSheet, {
  cellStyles: true,
  id: 'sheet\"><img src=x onerror=globalThis.__styledExcelSentinel=3>',
})
assert.equal(unsafeHtml.includes('<img src=x'), false)
assert.equal(unsafeHtml.includes('javascript:'), false)
assert.equal(unsafeHtml.includes('position:fixed'), false)
assert.equal(unsafeHtml.includes('<b>visible</b>'), true)
assert.equal(unsafeHtml.includes('&lt;img'), true)

const safeLinkSheet = XLSX.utils.aoa_to_sheet([['safe link']])
safeLinkSheet.A1.l = { Target: 'https://file-viewer.app/docs' }
const safeHtml = XLSX.utils.sheet_to_html(safeLinkSheet)
assert.equal(safeHtml.includes('<a href="https://file-viewer.app/docs">safe link</a>'), true)

const legacyMergeSheet = XLSX.utils.aoa_to_sheet([['legacy merge']])
legacyMergeSheet['!merges'] = [XLSX.utils.decode_range('A1:B2')]
const legacyMergeWorkbook = XLSX.utils.book_new()
XLSX.utils.book_append_sheet(legacyMergeWorkbook, legacyMergeSheet, 'Merge')
const legacyMergeBuffer = XLSX.write(legacyMergeWorkbook, { type: 'buffer', bookType: 'xlsx' })
const legacyMergeRoundTrip = XLSX.read(legacyMergeBuffer, { type: 'buffer', WTF: true })
assert.equal(legacyMergeRoundTrip.Sheets.Merge['!mergeErrors'][0].code, 'E_MERGE_BOUNDS')
assert.throws(
  () => XLSX.read(legacyMergeBuffer, { type: 'buffer', validateMerges: true }),
  /Merge range exceeds worksheet range/,
)
assert.throws(
  () => XLSX.write(legacyMergeWorkbook, { type: 'buffer', bookType: 'xlsx', validateMerges: true }),
  /Merge range exceeds worksheet range/,
)

console.log('styled-exceljs ODS prototype-pollution regression passed')
