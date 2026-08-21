# CHANGELOG

This log is intended to keep track of backwards-incompatible changes, including
but not limited to API changes and file location changes.  Minor behavioral
changes may not be included if they are not expected to break existing code.

* Sheet Visibility for ODS / FODS (h/t @edemaine)
* HTML DOM ingress support formulae (`data-f`)
* Proper handling of XLSX encoded entities (h/t @inreoh)
* Proper handling of invalid DIF sheets that match heuristics (h/t @lowkeyfish)

## v0.21.2

* Preserve TIFF drawing MIME metadata for XLSX images so browser renderers can
  select an explicit decoder instead of receiving `application/octet-stream`
* Keep workbook column-width measurement stable across sheets and tolerate
  missing Numbers format tables

中文摘要：

* XLSX 中的 TIFF 图片现在会保留 `image/tiff` 类型，便于浏览器渲染器按需解码
* 稳定多工作表列宽测量，并容错缺失的 Numbers 格式表

## v0.21.1

* Added browser-aware text measurement and `XLSX.utils.auto_fit_columns` for
  best-fit column widths with wrap, overflow, shrink-to-fit, and merge support
* `sheet_to_html` visual mode now supports `autoFit`, `overflow`, fixed-layout
  `<colgroup>` output, and Canvas-backed shrink-to-fit sizing
* Updated TypeScript definitions, English documentation, Chinese documentation,
  and unit tests for the auto-fit rendering path

中文摘要：

* 新增浏览器文字测量和 `XLSX.utils.auto_fit_columns`，可结合换行、溢出、
  收缩适应和合并跨度计算最优列宽
* `sheet_to_html` 视觉模式支持 `autoFit`、`overflow`、固定布局 `<colgroup>`
  和基于 Canvas 测量的收缩适应字号
* 更新 TypeScript 类型、英文文档、中文文档和自动列宽单元测试

## v0.21.0

* XLSX style parsing resolves fonts, fills, borders, alignment, protection,
  colors, theme tint, and number formats into full `cell.s` objects
* XLS BIFF8 style parsing maps Font / XF / XFExt / Palette / Theme records to
  the same style model used for XLSX
* `sheet_to_html` visual mode emits browser pixel row/column dimensions,
  styles, merged cells, drawing images, and chart SVG
* XLSX and XLS drawing/chart parsing is gated behind explicit `drawings` and
  `charts` options to preserve default read performance
* Added `ws["!drawings"]`, `ws["!charts"]`, chartsheet `ws["!chart"]`,
  `ws["!mergeErrors"]`, `XLSX.utils.validate_merges`, and browser pixel
  conversion helpers
* Added complete public TypeScript definitions for styles, colors, drawings,
  charts, merge errors, and visual HTML options
* Added `docs/visual-fidelity.md` and `misc/visual_matrix.js` for integration
  guidance and XLS/XLSX visual completeness checks

中文摘要：

* 完整扩展 XLSX / XLS 视觉读取能力，包括完整样式、颜色、边框、对齐、
  行列尺寸、合并校验、图片、绘图和常见图表模型
* `sheet_to_html` 在显式开启视觉选项后可输出浏览器像素尺寸、样式、
  合并单元格、图片层和 SVG 图表，默认读取路径仍保持轻量兼容
* 新增 `cell.s` 完整样式对象、`ws["!drawings"]`、`ws["!charts"]`、
  图表工作表 `ws["!chart"]`、`ws["!mergeErrors"]`、
  `XLSX.utils.validate_merges` 和像素换算工具
* 补全 TypeScript 类型声明，覆盖样式、颜色、边框、图表、绘图、合并错误
  和浏览器渲染选项
* 新增英文文档 `docs/visual-fidelity.md`、中文文档
  `docs/visual-fidelity.zh-CN.md` 和完整度报告脚本 `misc/visual_matrix.js`

## v0.20.3

* Correct parsing of NUMBERS and ODS merge cells (h/t @s-ashwin)
* More precise treatment of infinite and NaN values
* XLML Streaming Write
* Parse `Int8Array` objects (for compatibility with JS engines in Java)
* CSV Export only quote leading ID (h/t @lako12)

## v0.20.2

* Reworked parsing methods to avoid slow regexes (CVE-2024-22363)
* HTML properly encode data-v attribute
* SYLK read and write error cells

## v0.20.1

* `init` use packaged test files to work around GitHub breaking changes
* SSF date code rounding to 15 decimal digits (h/t @davidtamaki)
* `sheet_to_json` force UTC interpretation for formatted strings (h/t @Blanay)
* QPW extract result of string formula
* XLSX parse non-compliant merge cell expressions
* NUMBERS correctly handle rows omitted from official exports
* DBF parse empty logical field (h/t @Roman91)
* `dense` option added to types
* package.json add mini and core scripts to export map (h/t @stof)

## v0.20.0

* Use UTC interpretation of Date objects for date cells (potentially breaking)
* API functions support UTC and local time value interpretations
* Export `NaN` values to `#NUM!` and infinite values to `#DIV/0!`

## v0.19.3

* XLSX Ensure comment address is valid (h/t @slonser)
* Enforce Excel worksheet name restrictions
* Fixed "Prototype Pollution" vulnerability (CVE-2023-30533)

## v0.19.2

* XLSX proper decoding of hyperlinks (h/t @tw-yaxu)
* XLSX ignore unexpected attributes in rich text (h/t @colin4)
* `sheet_to_json` type fix (h/t @chsdwn)

## v0.19.1

* Fixed types issue in strict mode (h/t @younes-io)
* Numbers 12.2 parsing skip ActivityStream.iwa

## v0.19.0

* XLSX export hyperlinks compatible with google sheets (h/t Evan Bovie)
* NUMBERS export multiple sheets, full worksheet range
* formalized `dense` mode

## v0.18.12

* `package.json` added types in `exports` structure
* uncapped NUMBERS single-sheet single-table export
* DBF export records using supported codepages

## v0.18.11

* Base64 input ignore data URI wrapper
* Parse ZIP files that use ZIP64 extended information field
* More precise handling of time-only values
* Threaded Comment fallback text for older Excel

## v0.18.10

* `exports` field in package.json to satiate ViteJS and newer tooling
* JSC (Safari / Bun) perf, see <https://bugs.webkit.org/show_bug.cgi?id=243148>
* workbook `bookType` property to denote the origin format when parsed from file
* XLSX force export of stub cells with number formats when `sheetStubs` is set

## v0.18.9

* XLSX / ODS write defined names
* sync defined names to AutoFilter setting on export
* 1904 date system setting properly roundtripped
* ODS read/write number formats

## v0.18.8

* Plaintext parsing of dateless meridien time values (`1:23:45 PM`)
* Legacy format (SYLK / WK# / Multiplan) minutiae

## v0.18.7

* Normalized handling of `\r` and `\n` newline characters

## v0.18.6

* Removed all npm dependencies
* Auto-correct bad Google Sheets format `d.m`
* NUMBERS write merge cells, cells up to column "ALL"

## v0.18.5

* Enabled `sideEffects: false` in package.json
* Basic NUMBERS write support

## v0.18.4

* CSV output omits trailing record separator
* Properly terminate NodeJS Streams
* DBF preserve column types on import and use when applicable on export

## v0.18.3

* Removed references to `require` and `process` in browser builds

## v0.18.2

* Hotfix for unicode processing of XLSX exports

## v0.18.1

* Removed Node ESM build script and folded into standard ESM build
* Removed undocumented aliases including `make_formulae` and `get_formulae`

## v0.18.0

* Browser scripts only expose `XLSX` variable
* Module no longer ships with `dist/jszip.js` browser script

## v0.17.4

* CLI script moved to `xlsx-cli` package

## v0.17.3

* `window.XLSX` explicit assignment to satiate LWC
* CSV Proper formatting of errors
* HTML emit data-\* attributes

## v0.17.2

* Browser and Node optional ESM support
* DSV correct handling of bare quotes (h/t @bgamrat)

## v0.17.1

* `XLSB` writer uses short cell form when viable

## 0.17.0:

* mini build includes ODS parse/write support
* DBF explicitly cap worksheet to 1<<20 rows
* XLS throw errors on truncated records

## v0.16.2

* Disabled `PRN` parsing by default (better support for CSV without delimeters)

## v0.16.1

* skip empty custom property tags if data is absent (fixes DocSecurity issue)
* HTML output add raw value, type, number format
* DOM parse look for `v` / `t` / `z` attributes when determining value
* double quotes in properties escaped using `_x0022_`
* changed AMD structure for NetSuite and other RequireJS implementations
- `encode_cell` and `decode_cell` do not rely on `encode_col` / `decode_col`

## v0.16.0

* Date handling changed
* XLML certain tag tests are now case insensitive
* Fixed potentially vulnerable regular expressions

## v0.15.6

* CFB prevent infinite loop
* ODS empty cells marked as stub (type "z")
* `cellStyles` option implies `sheetStubs`

## v0.15.5

* `sheets` parse option to specify which sheets to parse

## v0.15.4

* AOA utilities properly preserve number formats
* Number formats captured in stub cells

## v0.15.3

* Properties and Custom Properties properly XML-encoded

## v0.15.2

- `sheet_get_cell` utility function
- `sheet_to_json` explicitly support `null` as alias for default behavior
- `encode_col` throw on negative column index
- HTML properly handle whitespace around tags in a run
- HTML use `id` option on write
- Files starting with `0x09` followed by a display character are now TSV files
- XLS parse references col/row indices mod by the correct number for BIFF ver
- XLSX comments moved to avoid overlapping cell
- XLSB outline level
- AutoFilter update `_FilterDatabase` defined name on write
- XLML skip CDATA blocks

## v0.15.1 (2019-08-14)

* XLSX ignore XML artifacts
* HTML capture and persist merges

## v0.15.0

* `dist/xlsx.mini.min.js` mini build with XLSX read/write and some utilities
* Removed legacy conversion utility functions

## v0.14.5

* XLS PtgNameX lookup
* XLS always create stub cells for blank cells with comments


## v0.14.4

* Better treatment of `skipHidden` in CSV output
* Ignore CLSID in XLS
* SYLK 7-bit character encoding
* SYLK and DBF codepage support

## v0.14.3

* Proper shifting of addresses in Shared Formulae

## v0.14.2

* Proper XML encoding of comments

## v0.14.1

* raw cell objects can be passed to `sheet_add_aoa`
* `_FilterDatabase` fix for AutoFilter-related crashes
* `stream.to_json` doesn't end up accidentally scanning to max row

## 0.14.0 (2018-09-06)

* `sheet_to_json` default flipped to `raw: true`

## 0.13.5 (2018-08-25)

* HTML output generates `<br/>` instead of encoded newline character

## 0.13.2 (2018-07-08)

* Buffer.from shim replaced, will not be defined in node `<=0.12`

## 0.13.0 (2018-06-01)

* Library reshaped to support AMD out of the box

## 0.12.11 (2018-04-27)

* XLS/XLSX/XLSB range truncation (errors in `WTF` mode)

## 0.12.4 (2018-03-04)

* `JSZip` renamed to `JSZipSync`

## 0.12.0 (2018-02-08)

* Extendscript target script in NPM package 

## 0.11.19 (2018-02-03)

* Error on empty workbook

## 0.11.16 (2017-12-30)

* XLS ANSI/CP separation
* 'array' write type and ArrayBuffer processing

## 0.11.6 (2017-10-16)

* Semicolon-delimited files are detected

## 0.11.5 (2017-09-30)

* Bower main script shifted to full version
* 'binary' / 'string' encoding

## 0.11.3 (2017-08-19)

* XLS cell ixfe/XF removed

## 0.11.0 (2017-07-31)

* Strip `require` statements from minified version
* minifier mangler enabled

## 0.10.9 (2017-07-28)

* XLML/HTML resolution logic looks further into the data stream to decide type
* Errors thrown on suspected RTF files

## 0.10.5 (2017-06-09)

* HTML Table output header/footer should not include `<table>` tag

## 0.10.2 (2017-05-16)

* Dates are converted to numbers by default (set `cellDates:true` to emit Dates)
* Module does not export CFB

## 0.9.10 (2017-04-08)

* `--perf` renamed to `--read-only`

## 0.9.9 (2017-04-03)

* default output format changed to XLSB
* comment text line endings are now normalized
* errors thrown on write when worksheets have invalid names

## 0.9.7 (2017-03-28)

* XLS legacy `!range` field removed
* Hyperlink tooltip is stored in the `Tooltip` field

## 0.9.6 (2017-03-25)

* `sheet_to_json` now passes `null` values when `raw` is set to `true`
* `sheet_to_json` treats `null` stub cells as values in conjunction with `raw`

## 0.9.5 (2017-03-22)

* `cellDates` affects parsing in non-XLSX formats

## 0.9.3 (2017-03-15)

* XLML property names are more closely mapped to the XLSX equivalent
* Stub cells are now cell type `z`

## 0.9.2 (2017-03-13)

* Removed stale TypeScript definition files.  Flowtype comments are used in the
  `xlsx.flow.js` source and stripped to produce `xlsx.js`.
* sed usage reworked to support GNU sed in-place form.  BSD sed seems to work,
  but the build script has not been tested on other sed variants:

```bash
$ sed -i.ext  [...] # GNU
$ sed -i .ext [...] # bsd
```

## 0.9.0 (2017-03-09)

* Removed ods.js source.  The xlsx.js source absorbed the ODS logic and exposes
  the ODS variable, so projects should remove references to ods.js
