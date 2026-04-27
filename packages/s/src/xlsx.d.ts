/* index.d.ts (C) 2015-present SheetJS and contributors */
// TypeScript Version: 2.2
declare namespace XLSX {

/** Version string */
export const version: string;

/** NODE ONLY! Attempts to read filename and parse */
export function readFile(filename: string, opts?: ParsingOptions): WorkBook;
/** Attempts to parse data */
export function read(data: any, opts?: ParsingOptions): WorkBook;
/** Attempts to write or download workbook data to file */
export function writeFile(data: WorkBook, filename: string, opts?: WritingOptions): any;
/** Attempts to write the workbook data */
export function write(data: WorkBook, opts?: WritingOptions): any;

/** Utility Functions */
export const utils: XLSX$Utils;
/** Stream Utility Functions */
export const stream: StreamUtils;

/** Number Format (either a string or an index to the format table) */
export type NumberFormat = string | number;

/** Basic File Properties */
export interface Properties {
    /** Summary tab "Title" */
    Title?: string;
    /** Summary tab "Subject" */
    Subject?: string;
    /** Summary tab "Author" */
    Author?: string;
    /** Summary tab "Manager" */
    Manager?: string;
    /** Summary tab "Company" */
    Company?: string;
    /** Summary tab "Category" */
    Category?: string;
    /** Summary tab "Keywords" */
    Keywords?: string;
    /** Summary tab "Comments" */
    Comments?: string;
    /** Statistics tab "Last saved by" */
    LastAuthor?: string;
    /** Statistics tab "Created" */
    CreatedDate?: Date;
}

/** Other supported properties */
export interface FullProperties extends Properties {
    ModifiedDate?: Date;
    Application?: string;
    AppVersion?: string;
    DocSecurity?: string;
    HyperlinksChanged?: boolean;
    SharedDoc?: boolean;
    LinksUpToDate?: boolean;
    ScaleCrop?: boolean;
    Worksheets?: number;
    SheetNames?: string[];
    ContentStatus?: string;
    LastPrinted?: string;
    Revision?: string | number;
    Version?: string;
    Identifier?: string;
    Language?: string;
}

export interface CommonOptions {
    /**
     * If true, throw errors when features are not understood
     * @default false
     */
    WTF?: boolean;

    /**
     * When reading a file with VBA macros, expose CFB blob to `vbaraw` field
     * When writing BIFF8/XLSB/XLSM, reseat `vbaraw` and export to file
     * @default false
     */
    bookVBA?: boolean;

    /**
     * When reading a file, store dates as type d (default is n)
     * When writing XLSX/XLSM file, use native date (default uses date codes)
     * @default false
     */
    cellDates?: boolean;

    /**
     * When reading a file, save style/theme info to the .s field
     * When writing a file, export style/theme info
     * @default false
     */
    cellStyles?: boolean;

    /**
     * If true, preserve row heights and column widths as browser CSS pixels
     * @default false
     */
    browserPixels?: boolean;

    /**
     * If true, parse supported chart records and expose worksheet `!charts`
     * @default false
     */
    charts?: boolean;

    /**
     * If true, parse supported drawing/image records and expose worksheet `!drawings`
     * @default false
     */
    drawings?: boolean;

    /**
     * If true, validate merge ranges and throw on invalid or overlapping ranges
     * @default false
     */
    validateMerges?: boolean;
}

export interface DateNFOption {
    /** Use specified date format */
    dateNF?: NumberFormat;
}

/** Options for read and readFile */
export interface ParsingOptions extends CommonOptions {
    /** Input data encoding */
    type?: 'base64' | 'binary' | 'buffer' | 'file' | 'array' | 'string';

    /** Default codepage */
    codepage?: number;

    /**
     * Save formulae to the .f field
     * @default true
     */
    cellFormula?: boolean;

    /**
     * Parse rich text and save HTML to the .h field
     * @default true
     */
    cellHTML?: boolean;

    /**
     * Save number format string to the .z field
     * @default false
     */
    cellNF?: boolean;

    /**
     * Generate formatted text to the .w field
     * @default true
     */
    cellText?: boolean;

    /** Override default date format (code 14) */
    dateNF?: string;

    /**
     * Create cell objects for stub cells
     * @default false
     */
    sheetStubs?: boolean;

    /**
     * If >0, read the first sheetRows rows
     * @default 0
     */
    sheetRows?: number;

    /**
     * If true, parse calculation chains
     * @default false
     */
    bookDeps?: boolean;

    /**
     * If true, add raw files to book object
     * @default false
     */
    bookFiles?: boolean;

    /**
     * If true, only parse enough to get book metadata
     * @default false
     */
    bookProps?: boolean;

    /**
     * If true, only parse enough to get the sheet names
     * @default false
     */
    bookSheets?: boolean;

    /**
     * If defined and file is encrypted, use password
     * @default ''
     */
    password?: string;

    /* If true, plaintext parsing will not parse values */
    raw?: boolean;

    dense?: boolean;
}

/** Options for write and writeFile */
export interface WritingOptions extends CommonOptions {
    /** Output data encoding */
    type?: 'base64' | 'binary' | 'buffer' | 'file' | 'array' | 'string';

    /**
     * Generate Shared String Table
     * @default false
     */
    bookSST?: boolean;

    /**
     * File format of generated workbook
     * @default 'xlsx'
     */
    bookType?: BookType;

    /**
     * Name of Worksheet (for single-sheet formats)
     * @default ''
     */
    sheet?: string;

    /**
     * Use ZIP compression for ZIP-based formats
     * @default false
     */
    compression?: boolean;

    /**
     * Suppress "number stored as text" errors in generated files
     * @default true
     */
    ignoreEC?: boolean;

    /** Override workbook properties on save */
    Props?: Properties;
}

/** Workbook Object */
export interface WorkBook {
    /**
     * A dictionary of the worksheets in the workbook.
     * Use SheetNames to reference these.
     */
    Sheets: { [sheet: string]: WorkSheet };

    /** Ordered list of the sheet names in the workbook */
    SheetNames: string[];

    /** Standard workbook Properties */
    Props?: FullProperties;

    /** Custom workbook Properties */
    Custprops?: object;

    Workbook?: WBProps;

    vbaraw?: any;
}

export interface SheetProps {
    /** Sheet Visibility (0=Visible 1=Hidden 2=VeryHidden) */
    Hidden?: 0 | 1 | 2;

    /** Name of Document Module in associated VBA Project */
    CodeName?: string;
}

/** Defined Name Object */
export interface DefinedName {
    /** Name */
    Name: string;

    /** Reference */
    Ref: string;

    /** Scope (undefined for workbook scope) */
    Sheet?: number;

    /** Name comment */
    Comment?: string;
}

/** Workbook-Level Attributes */
export interface WBProps {
    /** Sheet Properties */
    Sheets?: SheetProps[];

    /** Defined Names */
    Names?: DefinedName[];

    /** Workbook Views */
    Views?: WBView[];

    /** Other Workbook Properties */
    WBProps?: WorkbookProperties;
}

/** Workbook View */
export interface WBView {
    /** Right-to-left mode */
    RTL?: boolean;
}

/** Other Workbook Properties */
export interface WorkbookProperties {
    /** Worksheet Epoch (1904 if true, 1900 if false) */
    date1904?: boolean;

    /** Warn or strip personally identifying info on save */
    filterPrivacy?: boolean;

    /** Name of Document Module in associated VBA Project */
    CodeName?: string;
}

/** Column Properties Object */
export interface ColInfo {
    /* --- visibility --- */

    /** if true, the column is hidden */
    hidden?: boolean;

    /* --- column width --- */

    /** width in Excel's "Max Digit Width", width*256 is integral */
    width?: number;

    /** width in screen pixels */
    wpx?: number;

    /** width in "characters" */
    wch?: number;

    /** outline / group level */
    level?: number;

    /** XLSX column style index */
    style?: number | string;

    /** Column best-fit flag */
    bestFit?: boolean | string;
    bestfit?: boolean | string;

    /** Column custom-width flag */
    customWidth?: boolean | string;
    customwidth?: boolean | string;

    /** Excel's "Max Digit Width" unit, always integral */
    MDW?: number;

    /** default column style */
    s?: CellStyle | any;
}

/** Row Properties Object */
export interface RowInfo {
    /* --- visibility --- */

    /** if true, the column is hidden */
    hidden?: boolean;

    /* --- row height --- */

    /** height in screen pixels */
    hpx?: number;

    /** height in points */
    hpt?: number;

    /** outline / group level */
    level?: number;

    /** XLS row style index */
    ixfe?: number;

    /** Row custom-height flag */
    customHeight?: boolean | string;
    customheight?: boolean | string;

    /** default row style */
    s?: CellStyle | any;
}

export type FontVertAlign = "baseline" | "superscript" | "subscript";
export type FontScheme = "major" | "minor" | "none";
export type FillPatternType = "none" | "solid" | "mediumGray" | "darkGray" | "lightGray" | "darkHorizontal" | "darkVertical" | "darkDown" | "darkUp" | "darkGrid" | "darkTrellis" | "lightHorizontal" | "lightVertical" | "lightDown" | "lightUp" | "lightGrid" | "lightTrellis" | "gray125" | "gray0625" | string;
export type BorderStyleName = "none" | "thin" | "medium" | "dashed" | "dotted" | "thick" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot" | string;
export type HorizontalAlign = "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed" | string;
export type VerticalAlign = "top" | "center" | "bottom" | "justify" | "distributed" | string;
export type DrawingAnchorType = "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | string;
export type ChartType = "barChart" | "lineChart" | "areaChart" | "scatterChart" | "pieChart" | "doughnutChart" | "bubbleChart" | "radarChart" | "surfaceChart" | string;

export interface StyleColor { rgb?: string; theme?: number; tint?: number; indexed?: number; index?: number; auto?: boolean; raw_rgb?: string; }
export interface FontStyle { name?: string; sz?: number; color?: StyleColor; bold?: boolean | number; italic?: boolean | number; underline?: boolean | number | string; strike?: boolean | number; outline?: boolean | number; shadow?: boolean | number; condense?: boolean | number; extend?: boolean | number; vertAlign?: FontVertAlign | string; family?: number; charset?: number; scheme?: FontScheme | string; }
export interface FillStyle { patternType?: FillPatternType; fgColor?: StyleColor; bgColor?: StyleColor; gradientFill?: any; }
export interface BorderPr { style?: BorderStyleName; color?: StyleColor; }
export interface BorderStyle { left?: BorderPr; right?: BorderPr; top?: BorderPr; bottom?: BorderPr; diagonal?: BorderPr; horizontal?: BorderPr; vertical?: BorderPr; start?: BorderPr; end?: BorderPr; diagonalUp?: boolean; diagonalDown?: boolean; }
export interface AlignmentStyle { horizontal?: HorizontalAlign; vertical?: VerticalAlign; textRotation?: number; indent?: number; relativeIndent?: number; readingOrder?: number; wrapText?: boolean; shrinkToFit?: boolean; justifyLastLine?: boolean; }
export interface ProtectionStyle { locked?: boolean; hidden?: boolean; }
export interface CellStyle extends FillStyle { id?: number; xf?: any; numFmtId?: number; numFmt?: string; font?: FontStyle; fill?: FillStyle; border?: BorderStyle; alignment?: AlignmentStyle; protection?: ProtectionStyle; fgColor?: StyleColor; bgColor?: StyleColor; }
export interface TextMeasureOpts { measureText?: (text: string, font: string, style: CellStyle) => number; canvas?: any; MDW?: number; padding?: number; }
export interface AutoFitColumnOpts extends TextMeasureOpts { range?: string | Range; min?: number; max?: number; minPx?: number; maxPx?: number; includeMerged?: boolean; skipHidden?: boolean; set?: boolean; }
export type HTMLOverflowMode = "excel" | "visible" | "clip" | "hidden";
export interface DrawingRelationship { Type?: string; Target?: string; Id?: string; TargetMode?: string; }
export interface DrawingMarker { col: number; colOff: number; row: number; rowOff: number; }
export interface DrawingAnchor { type?: DrawingAnchorType; flags?: number; from?: DrawingMarker; to?: DrawingMarker; pos?: { x: number; y: number; }; ext?: { cx: number; cy: number; }; }
export interface ChartCache { values?: Array<number | string>; formatCode?: string; formula?: string; }
export interface ChartSeries { name?: string; idx?: number; order?: number; chartType?: ChartType; cat?: ChartCache; val?: ChartCache; xVal?: ChartCache; yVal?: ChartCache; bubbleSize?: ChartCache; data?: Array<number | string>; raw?: any; }
export interface ChartLegend { position?: string; raw?: any; }
export interface ChartModel { target?: string; raw?: any; rels?: any; type?: ChartType; grouping?: string; title?: string; legend?: ChartLegend; series?: ChartSeries[]; }
export interface ChartInfo { id?: string; rel?: DrawingRelationship; objectId?: number; biffType?: string; target?: string; path?: string; title?: string; anchor?: DrawingAnchor; model?: ChartModel; data?: WorkSheet; raw?: any; }
export interface DrawingImage { id?: string; rel?: DrawingRelationship; objectId?: number; biffType?: string; target?: string; path?: string; anchor?: DrawingAnchor; dataURI?: string; contentType?: string; raw?: any; }
export interface DrawingShape { id?: string; rel?: DrawingRelationship; objectId?: number; biffType?: string; target?: string; anchor?: DrawingAnchor; text?: string; props?: any; raw?: any; }
export interface DrawingInfo { raw?: any; chart?: string; charts?: ChartInfo[]; images?: DrawingImage[]; shapes?: DrawingShape[]; groups?: boolean; }
export interface MergeError { code: string; message: string; index: number; other?: number; range?: string | Range; otherRange?: string; ref?: string; }

/**
 * Write sheet protection properties.
 */
export interface ProtectInfo {
    /**
     * The password for formats that support password-protected sheets
     * (XLSX/XLSB/XLS). The writer uses the XOR obfuscation method.
     */
    password?: string;
    /**
     * Select locked cells
     * @default: true
     */
    selectLockedCells?: boolean;
    /**
     * Select unlocked cells
     * @default: true
     */
    selectUnlockedCells?: boolean;
    /**
     * Format cells
     * @default: false
     */
    formatCells?: boolean;
    /**
     * Format columns
     * @default: false
     */
    formatColumns?: boolean;
    /**
     * Format rows
     * @default: false
     */
    formatRows?: boolean;
    /**
     * Insert columns
     * @default: false
     */
    insertColumns?: boolean;
    /**
     * Insert rows
     * @default: false
     */
    insertRows?: boolean;
    /**
     * Insert hyperlinks
     * @default: false
     */
    insertHyperlinks?: boolean;
    /**
     * Delete columns
     * @default: false
     */
    deleteColumns?: boolean;
    /**
     * Delete rows
     * @default: false
     */
    deleteRows?: boolean;
    /**
     * Sort
     * @default: false
     */
    sort?: boolean;
    /**
     * Filter
     * @default: false
     */
    autoFilter?: boolean;
    /**
     * Use PivotTable reports
     * @default: false
     */
    pivotTables?: boolean;
    /**
     * Edit objects
     * @default: true
     */
    objects?: boolean;
    /**
     * Edit scenarios
     * @default: true
     */
    scenarios?: boolean;
}

/** Page Margins -- see Excel Page Setup .. Margins diagram for explanation */
export interface MarginInfo {
    /** Left side margin (inches) */
    left?: number;
    /** Right side margin (inches) */
    right?: number;
    /** Top side margin (inches) */
    top?: number;
    /** Bottom side margin (inches) */
    bottom?: number;
    /** Header top margin (inches) */
    header?: number;
    /** Footer bottom height (inches) */
    footer?: number;
}
export type SheetType = 'sheet' | 'chart';
export type SheetKeys = string | MarginInfo | SheetType;
/** General object representing a Sheet (worksheet or chartsheet) */
export interface Sheet {
    /**
     * Indexing with a cell address string maps to a cell object
     * Special keys start with '!'
     */
    [cell: string]: CellObject | SheetKeys | any;

    /** Sheet type */
    '!type'?: SheetType;

    /** Sheet Range */
    '!ref'?: string;

    /** Page Margins */
    '!margins'?: MarginInfo;
}

/** AutoFilter properties */
export interface AutoFilterInfo {
    /** Range of the AutoFilter table */
    ref: string;
}
export type WSKeys = SheetKeys | ColInfo[] | RowInfo[] | Range[] | ProtectInfo | AutoFilterInfo | ChartModel | ChartInfo[] | DrawingInfo | MergeError[];

/** Worksheet Object */
export interface WorkSheet extends Sheet {
    /**
     * Indexing with a cell address string maps to a cell object
     * Special keys start with '!'
     */
    [cell: string]: CellObject | WSKeys | any;

    /** Column Info */
    '!cols'?: ColInfo[];

    /** Row Info */
    '!rows'?: RowInfo[];

    /** Merge Ranges */
    '!merges'?: Range[];

    /** Merge validation errors */
    '!mergeErrors'?: MergeError[];

    /** Parsed chart drawings */
    '!charts'?: ChartInfo[];

    /** Parsed chart model for chartsheets */
    '!chart'?: ChartModel;

    /** Parsed drawing information */
    '!drawings'?: DrawingInfo;

    /** Worksheet Protection info */
    '!protect'?: ProtectInfo;

    /** AutoFilter info */
    '!autofilter'?: AutoFilterInfo;
}

/**
 * The Excel data type for a cell.
 * b Boolean, n Number, e error, s String, d Date, z Stub
 */
export type ExcelDataType = 'b' | 'n' | 'e' | 's' | 'd' | 'z';

/**
 * Type of generated workbook
 * @default 'xlsx'
 */
export type BookType = 'xlsx' | 'xlsm' | 'xlsb' | 'xls' | 'xla' | 'biff8' | 'biff5' | 'biff2' | 'xlml' | 'ods' | 'fods' | 'csv' | 'txt' | 'sylk' | 'slk' | 'html' | 'dif' | 'rtf' | 'prn' | 'eth' | 'dbf' | 'numbers';

/** Comment element */
export interface Comment {
    /** Author of the comment block */
    a?: string;

    /** Plaintext of the comment */
    t: string;

    /** If true, mark the comment as a part of a thread */
    T?: boolean;
}

/** Cell comments */
export interface Comments extends Array<Comment> {
    /** Hide comment by default */
    hidden?: boolean;
}

/** Link object */
export interface Hyperlink {
    /** Target of the link (HREF) */
    Target: string;

    /** Plaintext tooltip to display when mouse is over cell */
    Tooltip?: string;
}

/** Worksheet Cell Object */
export interface CellObject {
    /** The raw value of the cell.  Can be omitted if a formula is specified */
    v?: string | number | boolean | Date;

    /** Formatted text (if applicable) */
    w?: string;

    /**
     * The Excel Data Type of the cell.
     * b Boolean, n Number, e Error, s String, d Date, z Empty
     */
    t: ExcelDataType;

    /** Cell formula (if applicable) */
    f?: string;

    /** Range of enclosing array if formula is array formula (if applicable) */
    F?: string;

    /** If true, cell is a dynamic array formula (for supported file formats) */
    D?: boolean;

    /** Rich text encoding (if applicable) */
    r?: any;

    /** HTML rendering of the rich text (if applicable) */
    h?: string;

    /** Comments associated with the cell */
    c?: Comments;

    /** Number format string associated with the cell (if requested) */
    z?: NumberFormat;

    /** Cell hyperlink object (.Target holds link, .tooltip is tooltip) */
    l?: Hyperlink;

    /** The style/theme of the cell (if applicable) */
    s?: CellStyle | any;
}

/** Simple Cell Address */
export interface CellAddress {
    /** Column number */
    c: number;
    /** Row number */
    r: number;
}

/**
 * Range object (representing ranges like "A1:B2")
 */
export interface Range {
    /** Starting cell */
    s: CellAddress;
    /** Ending cell */
    e: CellAddress;
}

export interface Sheet2CSVOpts extends DateNFOption {
    /** Field Separator ("delimiter") */
    FS?: string;

    /** Record Separator ("row separator") */
    RS?: string;

    /** Remove trailing field separators in each record */
    strip?: boolean;

    /** Include blank lines in the CSV output */
    blankrows?: boolean;

    /** Skip hidden rows and columns in the CSV output */
    skipHidden?: boolean;
}

export interface OriginOption {
    /** Top-Left cell for operation (CellAddress or A1 string or row) */
    origin?: number | string | CellAddress;
}

export interface Sheet2HTMLOpts {
    /** TABLE element id attribute */
    id?: string;

    /** Add contenteditable to every cell */
    editable?: boolean;

    /** Header HTML */
    header?: string;

    /** Footer HTML */
    footer?: string;

    /** If true, remove javascript: URLs from hyperlinks */
    sanitizeLinks?: boolean;

    /** If true, emit CSS for parsed cell styles */
    cellStyles?: boolean;

    /** If true, emit browser pixel column widths and row heights */
    browserPixels?: boolean;

    /** If true, render parsed charts as inline SVG */
    charts?: boolean;

    /** If true, render parsed drawing images */
    drawings?: boolean;

    /** If true or an options object, auto-fit columns using measured text widths */
    autoFit?: boolean | AutoFitColumnOpts;

    /** Unwrapped text overflow behavior for browser rendering */
    overflow?: HTMLOverflowMode;

    /** Browser-compatible text measurement callback used by autoFit and shrink-to-fit */
    measureText?: (text: string, font: string, style: CellStyle) => number;

    /** Optional canvas-like object used for text measurement */
    canvas?: any;
}

export interface Sheet2JSONOpts extends DateNFOption {
    /** Output format */
    header?: "A"|number|string[];

    /** Override worksheet range */
    range?: any;

    /** Include or omit blank lines in the output */
    blankrows?: boolean;

    /** Default value for null/undefined values */
    defval?: any;

    /** if true, return raw data; if false, return formatted text */
    raw?: boolean;
}

export interface AOA2SheetOpts extends CommonOptions, DateNFOption {
    /**
     * Create cell objects for stub cells
     * @default false
     */
    sheetStubs?: boolean;
}

export interface SheetAOAOpts extends AOA2SheetOpts, OriginOption {}

export interface JSON2SheetOpts extends CommonOptions, DateNFOption {
    /** Use specified column order */
    header?: string[];

    /** Skip header row in generated sheet */
    skipHeader?: boolean;
}

export interface SheetJSONOpts extends JSON2SheetOpts, OriginOption {}

export interface Table2SheetOpts extends CommonOptions, DateNFOption {
    /* If true, plaintext parsing will not parse values */
    raw?: boolean;

    /**
     * If >0, read the first sheetRows rows
     * @default 0
     */
    sheetRows?: number;

    /** If true, hidden rows and cells will not be parsed */
    display?: boolean;
}

/** General utilities */
export interface XLSX$Utils {
    /* --- Import Functions --- */

    /** Converts an array of arrays of JS data to a worksheet. */
    aoa_to_sheet<T>(data: T[][], opts?: AOA2SheetOpts): WorkSheet;
    aoa_to_sheet(data: any[][], opts?: AOA2SheetOpts): WorkSheet;

    /** Converts an array of JS objects to a worksheet. */
    json_to_sheet<T>(data: T[], opts?: JSON2SheetOpts): WorkSheet;
    json_to_sheet(data: any[], opts?: JSON2SheetOpts): WorkSheet;

    /** BROWSER ONLY! Converts a TABLE DOM element to a worksheet. */
    table_to_sheet(data: any,  opts?: Table2SheetOpts): WorkSheet;
    table_to_book(data: any,  opts?: Table2SheetOpts): WorkBook;

    /* --- Export Functions --- */

    /** Converts a worksheet object to an array of JSON objects */
    sheet_to_json<T>(worksheet: WorkSheet, opts?: Sheet2JSONOpts): T[];
    sheet_to_json(worksheet: WorkSheet, opts?: Sheet2JSONOpts): any[][];
    sheet_to_json(worksheet: WorkSheet, opts?: Sheet2JSONOpts): any[];

    /** Generates delimiter-separated-values output */
    sheet_to_csv(worksheet: WorkSheet, options?: Sheet2CSVOpts): string;

    /** Generates UTF16 Formatted Text */
    sheet_to_txt(worksheet: WorkSheet, options?: Sheet2CSVOpts): string;

    /** Generates HTML */
    sheet_to_html(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /** Validate worksheet merge ranges */
    validate_merges(worksheet: WorkSheet, opts?: CommonOptions): MergeError[];

    /** Measure text width in browser pixels using Canvas when available */
    measure_text_width(text: string, style?: CellStyle, opts?: TextMeasureOpts): number;

    /** Auto-fit worksheet columns based on formatted cell text */
    auto_fit_columns(worksheet: WorkSheet, opts?: AutoFitColumnOpts): ColInfo[];

    /** Alias for auto_fit_columns */
    autofit_columns(worksheet: WorkSheet, opts?: AutoFitColumnOpts): ColInfo[];

    /** Convert XLSX column width to browser pixels */
    col_width_to_px(width: number): number;

    /** Convert browser pixels to XLSX column width */
    px_to_col_width(px: number): number;

    /** Convert row height points to browser pixels */
    row_height_to_px(hpt: number): number;

    /** Convert browser pixels to row height points */
    px_to_row_height(hpx: number): number;

    /** Generates a list of the formulae (with value fallbacks) */
    sheet_to_formulae(worksheet: WorkSheet): string[];

    /** Generates DIF */
    sheet_to_dif(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /** Generates SYLK (Symbolic Link) */
    sheet_to_slk(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /** Generates ETH */
    sheet_to_eth(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /* --- Cell Address Utilities --- */

    /** Converts 0-indexed cell address to A1 form */
    encode_cell(cell: CellAddress): string;

    /** Converts 0-indexed row to A1 form */
    encode_row(row: number): string;

    /** Converts 0-indexed column to A1 form */
    encode_col(col: number): string;

    /** Converts 0-indexed range to A1 form */
    encode_range(s: CellAddress, e: CellAddress): string;
    encode_range(r: Range): string;

    /** Converts A1 cell address to 0-indexed form */
    decode_cell(address: string): CellAddress;

    /** Converts A1 row to 0-indexed form */
    decode_row(row: string): number;

    /** Converts A1 column to 0-indexed form */
    decode_col(col: string): number;

    /** Converts A1 range to 0-indexed form */
    decode_range(range: string): Range;

    /** Format cell */
    format_cell(cell: CellObject, v?: any, opts?: any): string;

    /* --- General Utilities --- */

    /** Creates a new workbook */
    book_new(): WorkBook;

    /** Append a worksheet to a workbook */
    book_append_sheet(workbook: WorkBook, worksheet: WorkSheet, name?: string): void;

    /** Set sheet visibility (visible/hidden/very hidden) */
    book_set_sheet_visibility(workbook: WorkBook, sheet: number|string, visibility: number): void;

    /** Set number format for a cell */
    cell_set_number_format(cell: CellObject, fmt: string|number): CellObject;

    /** Set hyperlink for a cell */
    cell_set_hyperlink(cell: CellObject, target: string, tooltip?: string): CellObject;

    /** Set internal link for a cell */
    cell_set_internal_link(cell: CellObject, target: string, tooltip?: string): CellObject;

    /** Add comment to a cell */
    cell_add_comment(cell: CellObject, text: string, author?: string): void;

    /** Assign an Array Formula to a range */
    sheet_set_array_formula(ws: WorkSheet, range: Range|string, formula: string): WorkSheet;

    /** Add an array of arrays of JS data to a worksheet */
    sheet_add_aoa<T>(ws: WorkSheet, data: T[][], opts?: SheetAOAOpts): WorkSheet;
    sheet_add_aoa(ws: WorkSheet, data: any[][], opts?: SheetAOAOpts): WorkSheet;

    /** Add an array of JS objects to a worksheet */
    sheet_add_json(ws: WorkSheet, data: any[], opts?: SheetJSONOpts): WorkSheet;
    sheet_add_json<T>(ws: WorkSheet, data: T[], opts?: SheetJSONOpts): WorkSheet;


    consts: XLSX$Consts;
}

export interface XLSX$Consts {
    /* --- Sheet Visibility --- */

    /** Visibility: Visible */
    SHEET_VISIBLE: 0;

    /** Visibility: Hidden */
    SHEET_HIDDEN: 1;

    /** Visibility: Very Hidden */
    SHEET_VERYHIDDEN: 2;
}

/** NODE ONLY! these return Readable Streams */
export interface StreamUtils {
    /** CSV output stream, generate one line at a time */
    to_csv(sheet: WorkSheet, opts?: Sheet2CSVOpts): any;
    /** HTML output stream, generate one line at a time */
    to_html(sheet: WorkSheet, opts?: Sheet2HTMLOpts): any;
    /** JSON object stream, generate one row at a time */
    to_json(sheet: WorkSheet, opts?: Sheet2JSONOpts): any;
}
}
