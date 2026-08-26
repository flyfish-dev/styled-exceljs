import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'

import { strToU8, zipSync } from 'fflate'
import * as XLSXNamespace from '../xlsx.mjs'

const XLSX = XLSXNamespace.default || XLSXNamespace

const xml = (value) => strToU8(value)

const createFixture = (tableRef = 'A1:E4', customStyle = false) => zipSync({
  '[Content_Types].xml': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`),
  '_rels/.rels': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
  'xl/workbook.xml': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Table" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
  'xl/_rels/workbook.xml.rels': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`),
  'xl/worksheets/sheet1.xml': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:E4"/>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Column 1</t></is></c><c r="B1" t="inlineStr"><is><t>Column 2</t></is></c><c r="C1" t="inlineStr"><is><t>Count</t></is></c><c r="D1" t="inlineStr"><is><t>Total</t></is></c><c r="E1" t="inlineStr"><is><t>Average</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2" t="inlineStr"><is><t>One</t></is></c><c r="C2"><v>2</v></c><c r="D2"><v>10</v></c><c r="E2"><v>5</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>B</t></is></c><c r="B4" t="inlineStr"><is><t>Two</t></is></c><c r="C4"><v>4</v></c><c r="D4"><v>20</v></c><c r="E4"><v>5</v></c></row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`),
  'xl/worksheets/_rels/sheet1.xml.rels': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`),
  'xl/styles.xml': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="${customStyle ? 4 : 2}"><dxf><numFmt numFmtId="1" formatCode="0"/></dxf><dxf><alignment vertical="center" wrapText="1"/></dxf>${customStyle ? '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFF4B183"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFC6E0B4"/></patternFill></fill></dxf>' : ''}</dxfs>
  ${customStyle ? '<tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"><tableStyle name="CustomTableStyle" pivot="0" table="1" count="2"><tableStyleElement type="wholeTable" dxfId="2"/><tableStyleElement type="firstRowStripe" size="1" dxfId="3"/></tableStyle></tableStyles>' : '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'}
</styleSheet>`),
  'xl/theme/theme1.xml': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements><a:clrScheme name="Office">
    <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
    <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
    <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
    <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
    <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
  <a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>`),
  'xl/tables/table1.xml': xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="${tableRef}" totalsRowShown="0" headerRowDxfId="1">
  <tableColumns count="5">
    <tableColumn id="1" name="Column 1"/><tableColumn id="2" name="Column 2"/>
    <tableColumn id="3" name="Count"/><tableColumn id="4" name="Total"/>
    <tableColumn id="5" name="Average" dataDxfId="0"/>
  </tableColumns>
  <tableStyleInfo name="${customStyle ? 'CustomTableStyle' : 'TableStyleMedium13'}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`),
})

const fixturePath = process.env.STYLED_EXCELJS_TABLE_STYLE_FIXTURE
const fixture = fixturePath ? readFileSync(fixturePath) : createFixture()
const expectedRef = fixturePath ? 'A1:E10' : 'A1:E4'
const workbook = XLSX.read(fixture, { type: 'buffer', dense: true, cellStyles: true })
const worksheet = workbook.Sheets[workbook.SheetNames[0]]

assert.equal(worksheet['!tables'].length, 1)
assert.equal(worksheet['!tables'][0].ref, expectedRef)
assert.equal(worksheet['!tables'][0].styleInfo.name, 'TableStyleMedium13')
assert.equal(worksheet['!tables'][0].styleInfo.showRowStripes, true)

const cell = (row, col) => worksheet['!data']?.[row]?.[col]
const style = (row, col) => XLSX.utils.resolve_table_cell_style(worksheet, row, col, cell(row, col)?.s)

const header = style(0, 0)
assert.equal(header.fill.fgColor.rgb, '5B9BD5')
assert.equal(header.font.color.rgb, 'FFFFFF')
assert.equal(header.font.bold, 1)
assert.equal(header.alignment.wrapText, true)

const firstDataRow = style(1, 0)
const secondDataRow = style(2, 0)
assert.equal(firstDataRow.fill.fgColor.rgb, 'BDD7EE')
assert.equal(secondDataRow.fill.fgColor.rgb, 'DEEBF7')
assert.equal(cell(2, 0), undefined, 'lazy table styling must not materialize blank cells')

const columnDxf = style(1, 4)
assert.equal(columnDxf.numFmt.numFmtId, 1)
assert.equal(columnDxf.numFmt.formatCode, '0')

const html = XLSX.utils.sheet_to_html(worksheet, { cellStyles: true, header: '', footer: '' })
assert.match(html, /background-color:#5B9BD5/)
assert.match(html, /background-color:#BDD7EE/)
assert.match(html, /background-color:#DEEBF7/)

const largeWorkbook = XLSX.read(createFixture('A1:XFD1048576'), {
  type: 'array',
  dense: true,
  cellStyles: true,
})
assert.equal(largeWorkbook.Sheets.Table['!data'].length, 4)
assert.equal(largeWorkbook.Sheets.Table['!tables'][0].ref, 'A1:XFD1048576')

const customWorkbook = XLSX.read(createFixture('A1:E4', true), {
  type: 'array',
  dense: true,
  cellStyles: true,
})
const customSheet = customWorkbook.Sheets.Table
const customStyle = (row, col) => XLSX.utils.resolve_table_cell_style(
  customSheet,
  row,
  col,
  customSheet['!data']?.[row]?.[col]?.s,
)
assert.equal(customSheet['!tables'][0].styleInfo.name, 'CustomTableStyle')
assert.equal(customStyle(0, 0).fill.fgColor.rgb, 'F4B183')
assert.equal(customStyle(1, 0).fill.fgColor.rgb, 'C6E0B4')
assert.equal(customStyle(2, 0).fill.fgColor.rgb, 'F4B183')

const stylesDisabledWorkbook = XLSX.read(createFixture(), {
  type: 'array',
  dense: true,
  cellStyles: false,
})
assert.equal(stylesDisabledWorkbook.Sheets.Table['!tables'], undefined)

console.log(`styled-exceljs table-style regression passed (${fixturePath || 'generated fixture'})`)
