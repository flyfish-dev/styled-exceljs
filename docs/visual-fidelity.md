# XLSX / XLS Visual Fidelity Extensions

`styled-exceljs` is built from a fork of the original SheetJS Community Edition
source code.  The fork keeps SheetJS-compatible APIs and extends the original
reader / writer foundation with visual parsing and browser rendering features.

This document describes the browser rendering metadata exposed by this build.
The public model extends the existing SheetJS worksheet object and keeps default
read behavior compatible.  Expensive visual parsing is enabled with explicit
options.

Install from the public npm registry:

```bash
npm install styled-exceljs
```

## Options

```js
const wb = XLSX.read(data, {
  type: "buffer",
  cellStyles: true,
  browserPixels: true,
  charts: true,
  drawings: true,
  validateMerges: true
});

const html = XLSX.utils.sheet_to_html(ws, {
  cellStyles: true,
  browserPixels: true,
  autoFit: true,
  charts: true,
  drawings: true
});
```

| Option | Applies To | Effect |
| --- | --- | --- |
| `cellStyles` | read / HTML | Resolves full styles into `cell.s`, row styles, and column styles. |
| `browserPixels` | read / HTML | Converts XLSX/XLS column widths and row heights to CSS pixels. |
| `autoFit` | HTML / utilities | Measures formatted cell text and computes best-fit column widths. |
| `overflow` | HTML | Controls unwrapped text overflow (`excel` / `visible` / `clip` / `hidden`). |
| `charts` | read / HTML | Parses chart models and renders supported chart types as inline SVG. |
| `drawings` | read / HTML | Parses image/drawing anchors and renders supported images as `<img>`. |
| `validateMerges` | read | Throws on invalid, duplicate, overlapping, or out-of-bounds merge ranges. |
| `WTF` | read / utilities | Throws on unsupported or invalid records instead of preserving fallbacks. |

## Worksheet Fields

### Styles

When `cellStyles:true` is set, cells expose a resolved `cell.s` object:

```ts
interface CellStyle {
  id?: number;
  xf?: any;
  numFmtId?: number;
  numFmt?: string;
  font?: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  alignment?: AlignmentStyle;
  protection?: ProtectionStyle;
  patternType?: string;
  fgColor?: StyleColor;
  bgColor?: StyleColor;
}
```

`font`, `fill`, `border`, `alignment`, and `protection` are normalized across
XLSX and BIFF8 XLS.  The `fgColor`, `bgColor`, and `patternType` fields are
preserved as compatibility aliases for older integrations.

Colors may include resolved RGB as well as the original theme or palette data:

```ts
interface StyleColor {
  rgb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
  index?: number;
  auto?: boolean;
  raw_rgb?: string;
}
```

### Row and Column Dimensions

`ws["!cols"]` and `ws["!rows"]` expose browser-friendly measurements:

```ts
interface ColInfo {
  hidden?: boolean;
  width?: number;
  wpx?: number;
  wch?: number;
  MDW?: number;
  level?: number;
  style?: number | string;
  s?: CellStyle;
}

interface RowInfo {
  hidden?: boolean;
  hpx?: number;
  hpt?: number;
  level?: number;
  ixfe?: number;
  s?: CellStyle;
}
```

The conversion helpers are public:

```js
XLSX.utils.col_width_to_px(width);
XLSX.utils.px_to_col_width(px);
XLSX.utils.row_height_to_px(points);
XLSX.utils.px_to_row_height(px);
```

Best-fit widths can be computed with browser text measurement:

```js
XLSX.utils.auto_fit_columns(ws, {
  set: true,
  measureText(text, font, style) {
    ctx.font = font;
    return ctx.measureText(text).width;
  }
});
```

`auto_fit_columns` scans formatted display text (`cell.w` after formatting),
composes column, row, and cell styles, and applies Excel-like layout rules:

- `alignment.wrapText` measures the longest explicit line or word segment
  instead of expanding to the full paragraph width.
- `alignment.shrinkToFit` does not force a wider column when an existing width
  is available.
- merged cells distribute additional width across the merged column span.
- widths are clamped to the Excel column width range and converted through MDW.

`XLSX.utils.measure_text_width(text, style, opts)` is also public.  In browsers
it uses `opts.measureText` or Canvas `measureText`; in Node or headless
environments it uses a deterministic font-aware approximation.  For pixel-close
Excel rendering, pass a browser Canvas callback using the same font stack as the
rendered table.

### Drawings

When `drawings:true` is set, worksheets may expose `ws["!drawings"]`:

```ts
interface DrawingInfo {
  raw?: any;
  chart?: string;
  charts?: ChartInfo[];
  images?: DrawingImage[];
  shapes?: DrawingShape[];
  groups?: boolean;
}
```

Image entries include anchors and browser-ready `dataURI` values when a
supported embedded image is available:

```ts
interface DrawingImage {
  id?: string;
  rel?: DrawingRelationship;
  objectId?: number;
  biffType?: string;
  target?: string;
  path?: string;
  anchor?: DrawingAnchor;
  dataURI?: string;
  contentType?: string;
  raw?: any;
}
```

Complex shapes, unsupported image formats, WMF/EMF data, and OfficeArt records
are preserved in `raw` / `shapes` fallbacks for downstream renderers.

### Charts

When `charts:true` is set, worksheets may expose `ws["!charts"]` and chart
sheets may expose `ws["!chart"]`:

```ts
interface ChartInfo {
  id?: string;
  rel?: DrawingRelationship;
  objectId?: number;
  biffType?: string;
  target?: string;
  path?: string;
  title?: string;
  anchor?: DrawingAnchor;
  model?: ChartModel;
  data?: WorkSheet;
  raw?: any;
}

interface ChartModel {
  target?: string;
  raw?: any;
  rels?: any;
  type?: string;
  grouping?: string;
  title?: string;
  legend?: ChartLegend;
  series?: ChartSeries[];
}
```

Supported SVG rendering covers common bar/column, line, area, scatter, pie,
doughnut, bubble, and mixed chart models.  Advanced 3D, radar, surface, chartEx,
and complex OfficeArt cases retain raw metadata and fallback information.

### Merge Validation

`XLSX.utils.validate_merges(ws, opts)` returns a list of merge errors:

```ts
interface MergeError {
  code: string;
  message: string;
  index: number;
  other?: number;
  range?: string | Range;
  otherRange?: string;
  ref?: string;
}
```

Tolerant reads attach non-fatal errors to `ws["!mergeErrors"]`.  Reads with
`validateMerges:true` or utilities called with `{WTF:true}` throw.

## HTML Output

`XLSX.utils.sheet_to_html(ws, {cellStyles:true, browserPixels:true,
autoFit:true, charts:true, drawings:true})` emits:

- `<colgroup>` entries with browser pixel widths
- inline row heights
- font, fill, border, alignment, wrap, shrink-to-fit, overflow, and rotation CSS
- hidden row/column behavior
- `rowspan` / `colspan` for merged cells
- an absolutely positioned drawing layer with images and chart SVG

The renderer does not materialize empty grid cells for row/column styles.  Row
and column styles are composed with real cell styles at render time.

When `autoFit` is enabled, the generated table uses fixed layout and a computed
`<colgroup>`.  Unwrapped text defaults to Excel-like visible overflow.  Use
`overflow:"clip"` or `overflow:"hidden"` when the host page must constrain cell
content to the calculated column width.

## Completeness Report

The release includes a local report helper:

```bash
node misc/visual_matrix.js
```

The script reads representative XLS and XLSX fixtures with the visual options
enabled and reports support states for values, formulas, styles, dimensions,
merges, links/comments, drawings, charts, and HTML output.
