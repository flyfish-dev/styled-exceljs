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

console.log('styled-exceljs ODS prototype-pollution regression passed')
