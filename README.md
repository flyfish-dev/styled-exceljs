# [SheetJS](https://sheetjs.com)

## Fork Notice

`styled-exceljs` is a fork of the original SheetJS Community Edition source
code.  This fork keeps the SheetJS-compatible workbook and worksheet model, then
adds browser-oriented visual fidelity extensions for styles, dimensions,
drawings, charts, merge validation, and HTML rendering.

中文说明：`styled-exceljs` 基于 SheetJS Community Edition 原始代码 fork
并继续增强，保持 SheetJS 兼容的数据模型，同时扩展样式、行列尺寸、图片绘图、
图表、合并校验和浏览器 HTML 渲染能力。

The SheetJS Community Edition offers battle-tested open-source solutions for
extracting useful data from almost any complex spreadsheet and generating new
spreadsheets that will work with legacy and modern software alike.

[SheetJS Pro](https://sheetjs.com/pro) offers solutions beyond data processing:
Edit complex templates with ease; let out your inner Picasso with styling; make
custom sheets with images/graphs/PivotTables; evaluate formula expressions and
port calculations to web apps; automate common spreadsheet tasks, and much more!

## Documentation

- [API and Usage Documentation](https://docs.sheetjs.com)

- [Downloadable Scripts and Modules](https://cdn.sheetjs.com)

- [XLSX / XLS Visual Fidelity Extensions](docs/visual-fidelity.md)

- [XLSX / XLS 视觉保真扩展中文文档](docs/visual-fidelity.zh-CN.md)

## Visual Fidelity

Install this build from the public npm registry:

```bash
npm install styled-exceljs
```

This build extends the SheetJS worksheet model for browser-oriented rendering of
XLSX/XLS workbooks.  The default read path remains lightweight and compatible.
Visual metadata is exposed when explicit options are enabled:

```js
const wb = XLSX.read(data, {
  type: "buffer",
  cellStyles: true,
  browserPixels: true,
  charts: true,
  drawings: true,
  validateMerges: true
});

const ws = wb.Sheets[wb.SheetNames[0]];
const html = XLSX.utils.sheet_to_html(ws, {
  cellStyles: true,
  browserPixels: true,
  autoFit: true,
  charts: true,
  drawings: true
});
```

The following public worksheet fields are populated when requested:

- `cell.s`: resolved style object with `font`, `fill`, `border`,
  `alignment`, `protection`, `numFmt`, `fgColor`, and `bgColor`.
- `ws["!cols"]` / `ws["!rows"]`: column widths and row heights including
  browser pixel fields (`wpx`, `hpx`) and default row/column styles.
- `ws["!mergeErrors"]`: non-fatal merge validation issues in tolerant reads.
- `ws["!drawings"]`: parsed images, shapes, anchors, raw fallback records.
- `ws["!charts"]` and chartsheet `ws["!chart"]`: normalized chart models used
  by HTML SVG rendering.
- `XLSX.utils.auto_fit_columns(ws, opts)`: browser-measured best-fit column
  widths with wrap, overflow, shrink-to-fit, and merge-span handling.

`XLSX.utils.validate_merges(ws, opts)` validates merge ranges and reports
duplicate, overlapping, invalid, or out-of-bounds ranges.  Passing
`{WTF:true}` or reading with `{validateMerges:true}` throws on invalid merges.

See [docs/visual-fidelity.md](docs/visual-fidelity.md) for the full data
structure reference and release notes for unsupported fallback behavior.

## 中文说明

从官方 npm registry 安装：

```bash
npm install styled-exceljs
```

这个版本面向浏览器预览和报表渲染扩展了 XLSX / XLS 的视觉读取能力。
默认读取路径仍然保持轻量和兼容；只有显式开启 `cellStyles`、
`browserPixels`、`autoFit`、`charts`、`drawings` 和 `validateMerges` 等选项时，
才会解析完整样式、浏览器像素尺寸、自动列宽、图片、绘图、图表和合并单元格
校验结果。

常用入口如下：

- `cell.s`：解析后的完整样式对象，包括字体、填充、边框、对齐、保护和数字格式。
- `ws["!cols"]` / `ws["!rows"]`：列宽、行高、隐藏状态、层级和浏览器像素值。
- `ws["!drawings"]`：图片、形状、锚点和复杂绘图的 raw fallback。
- `ws["!charts"]` / `ws["!chart"]`：工作表嵌入图表和图表工作表模型，可用于 SVG 渲染。
- `XLSX.utils.validate_merges(ws, opts)`：检查非法、重复、重叠和越界的合并单元格。
- `XLSX.utils.auto_fit_columns(ws, opts)`：结合浏览器文字测量、换行、溢出、
  收缩适应和合并跨度计算最优列宽。

完整中文数据结构、选项说明和渲染行为请参考
[docs/visual-fidelity.zh-CN.md](docs/visual-fidelity.zh-CN.md)。

## Constellation

- <https://oss.sheetjs.com/notes/>: File Format Notes

- [`ssf`](packages/ssf): Format data using ECMA-376 spreadsheet format codes

- [`xlsx-cli`](packages/xlsx-cli): NodeJS command-line tool for processing files

- [`cfb`](https://git.sheetjs.com/SheetJS/js-cfb): Container (OLE/ZIP) file
processing library

- [`codepage`](https://git.sheetjs.com/SheetJS/js-codepage): Legacy text
encodings for XLS and other legacy spreadsheet formats

- [`dta`](packages/dta): Stata DTA file processor

## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 License are reserved by the Original Author.
