# XLSX / XLS 视觉保真扩展

`styled-exceljs` 基于 SheetJS Community Edition 原始代码 fork 并继续增强。
本 fork 保持 SheetJS 兼容 API 和读写基础，在此之上扩展视觉解析和浏览器渲染能力。

本文档说明本版本在 SheetJS 兼容工作表模型上新增的浏览器渲染元数据。
默认读取行为保持轻量和兼容；完整视觉信息需要通过显式选项启用。

从官方 npm registry 安装：

```bash
npm install styled-exceljs
```

## 快速使用

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

这些选项同样适用于 XLSX / XLSM 和 BIFF8 XLS。BIFF2-5 会尽量映射已有
样式和尺寸信息，但重点支持目标是 BIFF8 / Excel 97-2003。

## 选项

| 选项 | 适用位置 | 作用 |
| --- | --- | --- |
| `cellStyles` | 读取 / HTML | 将完整样式解析到 `cell.s`，并读取行样式和列样式。 |
| `browserPixels` | 读取 / HTML | 按 XLSX / XLS 规范将列宽、行高换算为浏览器 CSS 像素。 |
| `autoFit` | HTML / 工具函数 | 按格式化后的文字宽度计算最优列宽。 |
| `overflow` | HTML | 控制非换行文字溢出方式：`excel` / `visible` / `clip` / `hidden`。 |
| `charts` | 读取 / HTML | 解析图表模型，并将支持的图表类型以内联 SVG 渲染。 |
| `drawings` | 读取 / HTML | 解析图片、绘图锚点，支持的内嵌图片渲染为 `<img>`。 |
| `validateMerges` | 读取 | 读取时校验合并单元格，发现非法、重复、重叠或越界范围时抛错。 |
| `WTF` | 读取 / 工具函数 | 遇到不支持或非法记录时抛错，而不是保留 fallback 信息。 |

默认不开启 `charts` / `drawings`，避免普通读取路径解析重型绘图和图表关系。

## 工作表字段

### 样式 `cell.s`

开启 `cellStyles:true` 后，单元格会暴露解析后的完整样式对象：

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

`font`、`fill`、`border`、`alignment`、`protection` 和 `numFmt` 在 XLSX
和 XLS BIFF8 之间归一化。`fgColor`、`bgColor`、`patternType` 保留为
兼容旧集成的别名字段。

颜色结构会尽量同时保留解析后的 RGB 和原始来源信息：

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

XLSX 颜色支持 `rgb`、`theme`、`indexed`、`tint`；XLS 颜色按 palette /
theme / XFExt 尽量解析并映射到同一结构。

### 字体

```ts
interface FontStyle {
  name?: string;
  sz?: number;
  color?: StyleColor;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  strike?: boolean;
  outline?: boolean;
  shadow?: boolean;
  vertAlign?: "superscript" | "subscript" | "baseline" | string;
  family?: number;
  scheme?: "major" | "minor" | "none" | string;
  charset?: number;
}
```

### 填充

```ts
interface FillStyle {
  patternType?: string;
  fgColor?: StyleColor;
  bgColor?: StyleColor;
  gradient?: any;
}
```

普通纯色、图案填充和可解析的渐变信息会保留在 `fill` 中；无法可靠还原的
高级填充会保留原始信息，供上层渲染器继续处理。

### 边框

```ts
interface BorderStyle {
  left?: BorderPr;
  right?: BorderPr;
  top?: BorderPr;
  bottom?: BorderPr;
  diagonal?: BorderPr;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
  outline?: boolean;
}

interface BorderPr {
  style?: string;
  color?: StyleColor;
}
```

HTML 渲染会将常见边框样式转换为 CSS border。无法精确表达的边框会选择
浏览器中接近的样式，同时保留原始样式名。

### 对齐和保护

```ts
interface AlignmentStyle {
  horizontal?: string;
  vertical?: string;
  textRotation?: number;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  readingOrder?: number;
}

interface ProtectionStyle {
  locked?: boolean;
  hidden?: boolean;
}
```

HTML 输出支持水平 / 垂直对齐、换行、缩小字体、缩进、文字旋转等常见浏览器
可表达属性。

## 行列尺寸

`ws["!cols"]` 和 `ws["!rows"]` 会暴露浏览器友好的尺寸字段：

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

列宽遵循 XLSX 列宽公式和 Normal 字体 MDW。行高按 points / twips 转换为
96 CSS px/in 下的像素值。

公开换算工具：

```js
XLSX.utils.col_width_to_px(width);
XLSX.utils.px_to_col_width(px);
XLSX.utils.row_height_to_px(points);
XLSX.utils.px_to_row_height(px);
```

自动列宽可以接入浏览器真实文字测量：

```js
XLSX.utils.auto_fit_columns(ws, {
  set: true,
  measureText(text, font, style) {
    ctx.font = font;
    return ctx.measureText(text).width;
  }
});
```

`auto_fit_columns` 会扫描格式化后的显示文字（`cell.w`），合成列、行、
单元格样式，并尽量遵循 Excel 的布局行为：

- `alignment.wrapText` 开启时，按最长显式行或最长词段计算，而不是按整段
  文本强行拉宽列。
- `alignment.shrinkToFit` 开启且已有列宽时，不再强制扩大列宽。
- 合并单元格会把额外宽度分摊到覆盖的列。
- 最终列宽会限制在 Excel 允许范围内，并通过 MDW 公式转换。

`XLSX.utils.measure_text_width(text, style, opts)` 也可单独使用。在浏览器中
会优先使用 `opts.measureText` 或 Canvas `measureText`；在 Node / 无 DOM
环境中使用稳定的字体近似算法。若需要尽量贴近 Excel，请传入与页面实际
字体一致的 Canvas 测量回调。

## 图片和绘图

开启 `drawings:true` 后，工作表可能暴露 `ws["!drawings"]`：

```ts
interface DrawingInfo {
  raw?: any;
  chart?: string;
  charts?: ChartInfo[];
  images?: DrawingImage[];
  shapes?: DrawingShape[];
  groups?: boolean;
}

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

interface DrawingShape {
  id?: string;
  objectId?: number;
  biffType?: string;
  anchor?: DrawingAnchor;
  text?: string;
  svg?: string;
  raw?: any;
}
```

支持的 PNG / JPEG / DIB 内嵌图片会生成浏览器可直接使用的 `dataURI`。
WMF / EMF、复杂 OfficeArt 形状、组合形状和无法稳定转换的绘图会保留在
`raw` / `shapes` fallback 中，便于上层项目二次渲染。

## 图表

开启 `charts:true` 后：

- 普通工作表上的嵌入图表保存在 `ws["!charts"]`
- 图表工作表的主图表保存在 `ws["!chart"]`

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

interface ChartSeries {
  name?: string;
  color?: string;
  cat?: ChartCache;
  val?: ChartCache;
  xVal?: ChartCache;
  yVal?: ChartCache;
  bubbleSize?: ChartCache;
}
```

浏览器 SVG 渲染覆盖常见柱状图 / 条形图、折线图、面积图、散点图、饼图、
圆环图、气泡图和组合图。3D、雷达图、曲面图、chartEx 和复杂 OfficeArt
图表保留 raw 元数据和 fallback 信息。

## 合并单元格校验

`XLSX.utils.validate_merges(ws, opts)` 会返回合并范围错误列表：

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

可检测问题包括：

- 非法范围
- 重复范围
- 范围重叠
- 超出 `!ref` 或格式限制

宽容读取会把非致命错误放在 `ws["!mergeErrors"]`。读取时使用
`validateMerges:true` 或工具函数使用 `{WTF:true}` 会直接抛错。

## HTML 输出

```js
XLSX.utils.sheet_to_html(ws, {
  cellStyles: true,
  browserPixels: true,
  autoFit: true,
  charts: true,
  drawings: true
});
```

开启视觉选项后，HTML 输出会尽量包含：

- `<colgroup>` 列宽
- 行高 CSS
- 字体、字号、颜色、填充、边框
- 水平 / 垂直对齐、换行、缩进、旋转、收缩适应、溢出控制
- 隐藏行列
- 合并单元格的 `rowspan` / `colspan`
- 绝对定位图片层
- 内联 SVG 图表

渲染器不会为了行样式或列样式物化整张空白网格。真实单元格渲染时会按
cell XF > row XF > column XF > default XF 的思路合成最终样式。

开启 `autoFit` 后，HTML 表格会使用固定布局和计算后的 `<colgroup>`。
未换行文本默认使用接近 Excel 的可见溢出效果；如果宿主页面必须把内容
限制在列宽内，可以设置 `overflow:"clip"` 或 `overflow:"hidden"`。

## 完整度报告

仓库提供本地完整度报告脚本：

```bash
node misc/visual_matrix.js
```

脚本会用视觉选项读取代表性 XLS / XLSX fixture，并输出以下能力状态：

- 值和公式
- 样式
- 行列尺寸
- 合并单元格
- 超链接和批注
- 图片和绘图
- 图表模型
- HTML 渲染

状态包括 `supported`、`partial`、`raw-fallback` 和 `unsupported`。默认读取
路径也会被检查，确保重型绘图和图表不会在未显式开启选项时被解析。
