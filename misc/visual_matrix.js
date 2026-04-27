#!/usr/bin/env node
/* visual_matrix.js -- summarize XLS/XLSX visual fidelity support */
var fs = require("fs");
var path = require("path");
var XLSX = require("../xlsx.js");

var root = path.resolve(__dirname, "..");
var defaultFiles = [
	"test_files/cell_style_simple.xls",
	"test_files/cell_style_simple.xlsx",
	"test_files/column_width.xls",
	"test_files/row_height.xls",
	"test_files/row_height.xlsx",
	"test_files/merge_cells.xls",
	"test_files/merge_cells.xlsx",
	"test_files/hyperlink_stress_test_2011.xls",
	"test_files/hyperlink_stress_test_2011.xlsx",
	"test_files/comments_stress_test.xls",
	"test_files/comments_stress_test.xlsx",
	"test_files/wps/image.xls",
	"test_files/wps/image.xlsx",
	"test_files/apachepoi_SimpleWithImages.xls",
	"test_files/apachepoi_SimpleWithImages.xlsx",
	"test_files/jxls-src_chart.xls",
	"test_files/pyExcelerator_chart1v8.xls",
	"test_files/apachepoi_WithTwoCharts.xls",
	"test_files/apachepoi_WithThreeCharts.xlsx",
	"test_files/apachepoi_SimpleScatterChart.xlsx"
];

var args = process.argv.slice(2);
if(args.indexOf("--help") !== -1 || args.indexOf("-h") !== -1) {
	console.log("Usage: node misc/visual_matrix.js [files...]");
	console.log("Reads XLS/XLSX files with visual options and emits a JSON support matrix.");
	process.exit(0);
}

var featureNames = [
	"values", "formulas", "styles", "dimensions", "merges",
	"linksComments", "drawings", "charts", "html"
];
var statusRank = {unsupported:0, partial:1, "raw-fallback":2, supported:3};

function resolveDefault(f) { return path.resolve(root, f); }
function resolveArg(f) { return path.resolve(process.cwd(), f); }
function exists(f) { try { return fs.statSync(f).isFile(); } catch(e) { return false; } }
function rel(f) { return path.relative(root, f).replace(/\\/g, "/"); }
function fmtOf(f) {
	var ext = path.extname(f).toLowerCase().replace(/^\./, "");
	if(ext === "xlsm" || ext === "xltx" || ext === "xltm") return "xlsx";
	return ext || "unknown";
}
function initFeature() { return {status:"unsupported", evidence:[]}; }
function initFormat() {
	var out = {files:0, timings:{defaultMs:[], visualMs:[]}, features:{}};
	for(var i = 0; i < featureNames.length; ++i) out.features[featureNames[i]] = initFeature();
	return out;
}
function mark(formatReport, feature, status, file, sheet, detail) {
	var slot = formatReport.features[feature];
	if(!slot) return;
	if(statusRank[status] > statusRank[slot.status]) slot.status = status;
	if(slot.evidence.length < 8) slot.evidence.push({file:rel(file), sheet:sheet, status:status, detail:detail});
}
function hasStyle(s) {
	return !!(s && (s.font || s.fill || s.border || s.alignment || s.protection || s.numFmt || s.patternType || s.fgColor || s.bgColor));
}
function htmlHas(str, token) { return str.indexOf(token) !== -1; }
function elapsedMs(start) {
	var diff = process.hrtime(start);
	return Math.round((diff[0] * 1000 + diff[1] / 1e6) * 1000) / 1000;
}
function avg(arr) {
	if(!arr.length) return 0;
	var sum = 0;
	for(var i = 0; i < arr.length; ++i) sum += arr[i];
	return Math.round(sum / arr.length * 1000) / 1000;
}
function max(arr) {
	var m = 0;
	for(var i = 0; i < arr.length; ++i) if(arr[i] > m) m = arr[i];
	return m;
}
function workbookHasHeavyVisuals(wb) {
	for(var i = 0; i < wb.SheetNames.length; ++i) {
		var ws = wb.Sheets[wb.SheetNames[i]];
		if(ws && (ws["!charts"] || ws["!chart"] || ws["!drawings"])) return true;
	}
	return false;
}
function readWorkbook(file, opts) {
	var data = fs.readFileSync(file);
	var start = process.hrtime();
	var wb = XLSX.read(data, opts);
	return {wb:wb, ms:elapsedMs(start)};
}
function scanSheet(ws) {
	var out = {
		cells:0, formulas:0, styledCells:0, styledRows:0, styledCols:0,
		dimensionItems:0, merges:0, mergeErrors:0, links:0, comments:0,
		drawingImages:0, drawingShapes:0, drawingRaw:0,
		charts:0, chartModels:0, chartRaw:0,
		html:{table:false, style:false, colgroup:false, span:false, svg:false, img:false}
	};
	var keys = Object.keys(ws);
	for(var i = 0; i < keys.length; ++i) {
		var k = keys[i];
		if(k.charAt(0) === "!") continue;
		var cell = ws[k];
		if(!cell) continue;
		++out.cells;
		if(cell.f != null) ++out.formulas;
		if(hasStyle(cell.s)) ++out.styledCells;
		if(cell.l) ++out.links;
		if(cell.c && cell.c.length) out.comments += cell.c.length;
	}
	if(ws["!cols"]) for(i = 0; i < ws["!cols"].length; ++i) {
		var col = ws["!cols"][i];
		if(!col) continue;
		if(col.wpx != null || col.wch != null || col.width != null || col.hidden) ++out.dimensionItems;
		if(hasStyle(col.s)) ++out.styledCols;
	}
	if(ws["!rows"]) for(i = 0; i < ws["!rows"].length; ++i) {
		var row = ws["!rows"][i];
		if(!row) continue;
		if(row.hpx != null || row.hpt != null || row.hidden) ++out.dimensionItems;
		if(hasStyle(row.s)) ++out.styledRows;
	}
	if(ws["!merges"]) {
		out.merges = ws["!merges"].length;
		try { out.mergeErrors = XLSX.utils.validate_merges(ws).length; }
		catch(e) { out.mergeErrors = 1; }
	}
	var drawings = ws["!drawings"];
	if(drawings) {
		if(drawings.images) out.drawingImages += drawings.images.length;
		if(drawings.shapes) out.drawingShapes += drawings.shapes.length;
		if(drawings.raw) out.drawingRaw += drawings.raw.length;
	}
	var charts = ws["!charts"] || (ws["!chart"] ? [ws["!chart"]] : []);
	out.charts = charts.length;
	for(i = 0; i < charts.length; ++i) {
		if(charts[i] && charts[i].model) ++out.chartModels;
		if(charts[i] && charts[i].raw) ++out.chartRaw;
	}
	try {
		var html = XLSX.utils.sheet_to_html(ws, {cellStyles:true, browserPixels:true, charts:true, drawings:true});
		out.html.table = htmlHas(html, "<table");
		out.html.style = htmlHas(html, "style=\"");
		out.html.colgroup = htmlHas(html, "<colgroup>");
		out.html.span = htmlHas(html, "rowspan=") || htmlHas(html, "colspan=");
		out.html.svg = htmlHas(html, "<svg");
		out.html.img = htmlHas(html, "<img");
	} catch(e) {
		out.html.error = String(e && e.message || e);
	}
	return out;
}
function markSheet(report, fmt, file, sheetName, scan) {
	var fr = report.formats[fmt];
	if(scan.cells) mark(fr, "values", "supported", file, sheetName, scan.cells + " cells");
	if(scan.formulas) mark(fr, "formulas", "supported", file, sheetName, scan.formulas + " formulas");
	if(scan.styledCells || scan.styledRows || scan.styledCols) mark(fr, "styles", "supported", file, sheetName, scan.styledCells + " cells, " + scan.styledRows + " rows, " + scan.styledCols + " cols");
	if(scan.dimensionItems) mark(fr, "dimensions", "supported", file, sheetName, scan.dimensionItems + " row/col items");
	if(scan.merges) mark(fr, "merges", scan.mergeErrors ? "partial" : "supported", file, sheetName, scan.merges + " merges, " + scan.mergeErrors + " errors");
	if(scan.links || scan.comments) mark(fr, "linksComments", "supported", file, sheetName, scan.links + " links, " + scan.comments + " comments");
	if(scan.drawingImages) mark(fr, "drawings", "supported", file, sheetName, scan.drawingImages + " images");
	else if(scan.drawingShapes || scan.drawingRaw) mark(fr, "drawings", "raw-fallback", file, sheetName, scan.drawingShapes + " shapes, " + scan.drawingRaw + " raw records");
	if(scan.charts && scan.chartModels && scan.html.svg) mark(fr, "charts", "supported", file, sheetName, scan.chartModels + " models rendered to SVG");
	else if(scan.charts && scan.chartModels) mark(fr, "charts", "partial", file, sheetName, scan.chartModels + " models");
	else if(scan.chartRaw) mark(fr, "charts", "raw-fallback", file, sheetName, scan.chartRaw + " raw charts");
	if(scan.html.table && (scan.html.style || scan.html.colgroup || scan.html.span || scan.html.svg || scan.html.img)) {
		mark(fr, "html", "supported", file, sheetName, JSON.stringify(scan.html));
	} else if(scan.html.table) mark(fr, "html", "partial", file, sheetName, JSON.stringify(scan.html));
}

var files = (args.length ? args.map(resolveArg) : defaultFiles.map(resolveDefault)).filter(exists);
if(!files.length) {
	console.error("No XLS/XLSX files found. Pass files explicitly or run from a checkout with test_files.");
	process.exit(1);
}

var report = {
	generatedAt: new Date().toISOString(),
	options: {cellStyles:true, browserPixels:true, charts:true, drawings:true},
	formats: {xls:initFormat(), xlsx:initFormat()},
	files: []
};

for(var fidx = 0; fidx < files.length; ++fidx) {
	var file = files[fidx], fmt = fmtOf(file);
	if(fmt !== "xls" && fmt !== "xlsx") continue;
	var fileReport = {file:rel(file), format:fmt, sheets:[]};
	try {
		var base = readWorkbook(file, {type:"buffer"});
		var visual = readWorkbook(file, {type:"buffer", cellStyles:true, browserPixels:true, charts:true, drawings:true, cellFormula:true, cellNF:true});
		report.formats[fmt].files++;
		report.formats[fmt].timings.defaultMs.push(base.ms);
		report.formats[fmt].timings.visualMs.push(visual.ms);
		fileReport.defaultMs = base.ms;
		fileReport.visualMs = visual.ms;
		fileReport.defaultHeavyVisuals = workbookHasHeavyVisuals(base.wb);
		for(var sidx = 0; sidx < visual.wb.SheetNames.length; ++sidx) {
			var sheetName = visual.wb.SheetNames[sidx];
			var ws = visual.wb.Sheets[sheetName];
			var scan = scanSheet(ws);
			fileReport.sheets.push({name:sheetName, scan:scan});
			markSheet(report, fmt, file, sheetName, scan);
		}
	} catch(e) {
		fileReport.error = String(e && e.message || e);
	}
	report.files.push(fileReport);
}

Object.keys(report.formats).forEach(function(fmt) {
	var fr = report.formats[fmt];
	fr.timings.avgDefaultMs = avg(fr.timings.defaultMs);
	fr.timings.avgVisualMs = avg(fr.timings.visualMs);
	fr.timings.maxDefaultMs = max(fr.timings.defaultMs);
	fr.timings.maxVisualMs = max(fr.timings.visualMs);
	fr.timings.avgVisualOverDefault = fr.timings.avgDefaultMs ? Math.round(fr.timings.avgVisualMs / fr.timings.avgDefaultMs * 1000) / 1000 : 0;
});

console.log(JSON.stringify(report, null, 2));
