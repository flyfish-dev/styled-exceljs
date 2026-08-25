function hex2RGB(h) {
	var o = h.slice(h[0]==="#"?1:0).slice(0,6);
	return [parseInt(o.slice(0,2),16),parseInt(o.slice(2,4),16),parseInt(o.slice(4,6),16)];
}
function rgb2Hex(rgb) {
	for(var i=0,o=1; i!=3; ++i) o = o*256 + (rgb[i]>255?255:rgb[i]<0?0:rgb[i]);
	return o.toString(16).toUpperCase().slice(1);
}

function rgb2HSL(rgb) {
	var R = rgb[0]/255, G = rgb[1]/255, B=rgb[2]/255;
	var M = Math.max(R, G, B), m = Math.min(R, G, B), C = M - m;
	if(C === 0) return [0, 0, R];

	var H6 = 0, S = 0, L2 = (M + m);
	S = C / (L2 > 1 ? 2 - L2 : L2);
	switch(M){
		case R: H6 = ((G - B) / C + 6)%6; break;
		case G: H6 = ((B - R) / C + 2); break;
		case B: H6 = ((R - G) / C + 4); break;
	}
	return [H6 / 6, S, L2 / 2];
}

function hsl2RGB(hsl){
	var H = hsl[0], S = hsl[1], L = hsl[2];
	var C = S * 2 * (L < 0.5 ? L : 1 - L), m = L - C/2;
	var rgb = [m,m,m], h6 = 6*H;

	var X;
	if(S !== 0) switch(h6|0) {
		case 0: case 6: X = C * h6; rgb[0] += C; rgb[1] += X; break;
		case 1: X = C * (2 - h6);   rgb[0] += X; rgb[1] += C; break;
		case 2: X = C * (h6 - 2);   rgb[1] += C; rgb[2] += X; break;
		case 3: X = C * (4 - h6);   rgb[1] += X; rgb[2] += C; break;
		case 4: X = C * (h6 - 4);   rgb[2] += C; rgb[0] += X; break;
		case 5: X = C * (6 - h6);   rgb[2] += X; rgb[0] += C; break;
	}
	for(var i = 0; i != 3; ++i) rgb[i] = Math.round(rgb[i]*255);
	return rgb;
}

/* 18.8.3 bgColor tint algorithm */
function rgb_tint(hex, tint) {
	if(tint === 0) return hex;
	var hsl = rgb2HSL(hex2RGB(hex));
	if (tint < 0) hsl[2] = hsl[2] * (1 + tint);
	else hsl[2] = 1 - (1 - hsl[2]) * (1 - tint);
	return rgb2Hex(hsl2RGB(hsl));
}

function style_color_from_attrs(y, themes) {
	var o = ({}/*:any*/);
	if(y == null) return o;
	if(y.auto != null) o.auto = parsexmlbool(y.auto);
	if(y.rgb != null) o.rgb = y.rgb.slice(-6).toUpperCase();
	if(y.indexed != null) {
		o.indexed = parseInt(y.indexed, 10); o.index = o.indexed;
		var icv = XLSIcv[o.indexed];
		if(o.indexed == 81) icv = XLSIcv[1];
		if(!icv) icv = XLSIcv[1];
		if(icv) o.rgb = rgb2Hex(icv);
	}
	if(y.theme != null) {
		o.theme = parseInt(y.theme, 10);
		if(y.tint != null) o.tint = parseFloat(y.tint);
		if(themes && themes.themeElements && themes.themeElements.clrScheme && themes.themeElements.clrScheme[o.theme]) {
			o.raw_rgb = themes.themeElements.clrScheme[o.theme].rgb;
			o.rgb = rgb_tint(o.raw_rgb, o.tint || 0);
		}
	}
	if(y.tint != null && o.tint == null) o.tint = parseFloat(y.tint);
	return o;
}

function resolve_style_color(color, themes) {
	if(!color) return color;
	var o = dup(color);
	if(o.rgb) o.rgb = ("" + o.rgb).slice(-6).toUpperCase();
	if(o.indexed != null && !o.rgb) {
		var icv = XLSIcv[o.indexed];
		if(o.indexed == 81) icv = XLSIcv[1];
		if(!icv) icv = XLSIcv[1];
		if(icv) o.rgb = rgb2Hex(icv);
	}
	if(o.index != null && o.indexed == null) o.indexed = o.index;
	if(o.theme != null && themes && themes.themeElements && themes.themeElements.clrScheme && themes.themeElements.clrScheme[o.theme]) {
		o.raw_rgb = themes.themeElements.clrScheme[o.theme].rgb;
		o.rgb = rgb_tint(o.raw_rgb, o.tint || 0);
	}
	return o;
}

function resolve_style_obj_color(obj, themes) {
	if(!obj) return obj;
	var o = dup(obj);
	if(o.color) o.color = resolve_style_color(o.color, themes);
	if(o.fgColor) o.fgColor = resolve_style_color(o.fgColor, themes);
	if(o.bgColor) o.bgColor = resolve_style_color(o.bgColor, themes);
	["left","right","top","bottom","diagonal","horizontal","vertical","start","end"].forEach(function(k) {
		if(o[k] && o[k].color) o[k].color = resolve_style_color(o[k].color, themes);
	});
	if(o.gradientFill && o.gradientFill.stops) o.gradientFill.stops.forEach(function(stop) {
		if(stop.color) stop.color = resolve_style_color(stop.color, themes);
	});
	return o;
}

/* 18.3.1.13 width calculations */
/* [MS-OI29500] 2.1.595 Column Width & Formatting */
var DEF_MDW = 7, MAX_MDW = 15, MIN_MDW = 1, MDW = DEF_MDW;
function width2px(width) { return Math.floor(( width + (Math.round(128/MDW))/256 )* MDW ); }
function px2char(px) { return (Math.floor((px - 5)/MDW * 100 + 0.5))/100; }
function char2width(chr) { return (Math.round((chr * MDW + 5)/MDW*256))/256; }
//function px2char_(px) { return (((px - 5)/MDW * 100 + 0.5))/100; }
//function char2width_(chr) { return (((chr * MDW + 5)/MDW*256))/256; }
function cycle_width(collw) { return char2width(px2char(width2px(collw))); }
/* XLSX/XLSB/XLS specify width in units of MDW */
function find_mdw_colw(collw) {
	var delta = Math.abs(collw - cycle_width(collw)), _MDW = MDW;
	if(delta > 0.005) for(MDW=MIN_MDW; MDW<MAX_MDW; ++MDW) if(Math.abs(collw - cycle_width(collw)) <= delta) { delta = Math.abs(collw - cycle_width(collw)); _MDW = MDW; }
	MDW = _MDW;
}
/* XLML specifies width in terms of pixels */
/*function find_mdw_wpx(wpx) {
	var delta = Infinity, guess = 0, _MDW = MIN_MDW;
	for(MDW=MIN_MDW; MDW<MAX_MDW; ++MDW) {
		guess = char2width_(px2char_(wpx))*256;
		guess = (guess) % 1;
		if(guess > 0.5) guess--;
		if(Math.abs(guess) < delta) { delta = Math.abs(guess); _MDW = MDW; }
	}
	MDW = _MDW;
}*/

function process_col(coll/*:ColInfo*/) {
	if(coll.width) {
		coll.wpx = width2px(coll.width);
		coll.wch = px2char(coll.wpx);
		coll.MDW = MDW;
	} else if(coll.wpx) {
		coll.wch = px2char(coll.wpx);
		coll.width = char2width(coll.wch);
		coll.MDW = MDW;
	} else if(typeof coll.wch == 'number') {
		coll.width = char2width(coll.wch);
		coll.wpx = width2px(coll.width);
		coll.MDW = MDW;
	}
	if(coll.customWidth) delete coll.customWidth;
}

var XLSX_COL_WIDTH_MAX = 255, XLSX_COL_WIDTH_PADDING = 5;
function clamp_col_width(width) {
	return width > XLSX_COL_WIDTH_MAX ? XLSX_COL_WIDTH_MAX : width < 0 ? 0 : width;
}
function px2col(px) {
	var width = char2width(px2char(Math.max(0, px)));
	return clamp_col_width(width);
}
function col2px(col) {
	if(!col) return 64;
	if(col.wpx != null) return col.wpx;
	if(col.width != null) return width2px(col.width);
	if(col.wch != null) return width2px(char2width(col.wch));
	return 64;
}
function set_col_width_from_px(col, px) {
	if(!col) col = ({}/*:any*/);
	col.wpx = Math.max(0, Math.ceil(px));
	col.wch = px2char(col.wpx);
	col.width = px2col(col.wpx);
	col.MDW = MDW;
	col.bestFit = true;
	col.customWidth = true;
	return col;
}
function style_font_size_pt(style) {
	var font = style && style.font || {};
	var sz = +font.sz;
	return sz > 0 ? sz : 11;
}
function style_font_family(style) {
	var font = style && style.font || {};
	return font.name || "Calibri";
}
function css_string_escape(value/*:any*/, quote/*:string*/)/*:string*/ {
	var str = String(value), out = [];
	for(var i = 0; i < str.length; ++i) {
		var cc = str.charCodeAt(i), ch = str.charAt(i);
		if(ch == quote || ch == "\\" || cc < 32 || cc == 127 || ch == "<" || ch == ">" || ch == "&") out.push("\\" + cc.toString(16) + " ");
		else out.push(ch);
	}
	return out.join("");
}
function css_font_from_style(style) {
	var font = style && style.font || {};
	var parts = [];
	if(font.italic) parts.push("italic");
	if(font.bold) parts.push("bold");
	parts.push(style_font_size_pt(style) + "pt");
	var name = style_font_family(style);
	if(/[,\s'"]/.test(name)) name = '"' + css_string_escape(name, '"') + '"';
	else name = css_string_escape(name, '"');
	parts.push(name);
	return parts.join(" ");
}
function fallback_char_width(ch, fpx) {
	var cc = ch.charCodeAt(0);
	if(cc == 9 || cc == 32) return fpx * 0.33;
	if(cc >= 48 && cc <= 57) return fpx * 0.52;
	if(cc >= 65 && cc <= 90) return fpx * 0.62;
	if(cc >= 97 && cc <= 122) {
		if("iljtfr".indexOf(ch) > -1) return fpx * 0.28;
		if("mw".indexOf(ch) > -1) return fpx * 0.82;
		return fpx * 0.50;
	}
	if(cc >= 0x2E80) return fpx;
	return fpx * 0.52;
}
function fallback_text_width(text, style) {
	var fpx = pt2px_browser(style_font_size_pt(style)), width = 0;
	for(var i = 0; i < text.length; ++i) width += fallback_char_width(text.charAt(i), fpx);
	return width;
}
function measure_text_width(text/*:string*/, style/*:?CellStyle*/, opts/*:?any*/)/*:number*/ {
	var o = opts || {};
	var str = text == null ? "" : String(text);
	var lines = str.split(/\r\n|\n|\r/g), max = 0, i = 0, w = 0;
	if(o.measureText) for(i = 0; i < lines.length; ++i) {
		w = +o.measureText(lines[i], css_font_from_style(style), style || {});
		if(isFinite(w) && w > max) max = w;
	}
	else {
		var canvas = o.canvas, ctx = null;
		if(!canvas && typeof document !== "undefined" && document.createElement) try { canvas = document.createElement("canvas"); } catch(e) {}
		try { ctx = canvas && canvas.getContext && canvas.getContext("2d"); } catch(e) { ctx = null; }
		if(ctx && ctx.measureText) {
			ctx.font = css_font_from_style(style);
			for(i = 0; i < lines.length; ++i) {
				w = ctx.measureText(lines[i]).width;
				if(isFinite(w) && w > max) max = w;
			}
		} else for(i = 0; i < lines.length; ++i) {
			w = fallback_text_width(lines[i], style || {});
			if(isFinite(w) && w > max) max = w;
		}
	}
	return max;
}
function auto_fit_cell_text(cell/*:Cell*/, opts/*:?any*/)/*:string*/ {
	if(!cell || cell.v == null) return "";
	if(cell.w == null && cell.t != 'z') format_cell(cell);
	return String(cell.w != null ? cell.w : cell.v);
}
function text_width_segments(text, wrap) {
	var lines = String(text == null ? "" : text).split(/\r\n|\n|\r/g), out = [], i = 0, j = 0, parts = null;
	if(!wrap) return lines;
	for(i = 0; i < lines.length; ++i) {
		parts = lines[i].split(/[\t \u00A0\-\/]+/);
		for(j = 0; j < parts.length; ++j) if(parts[j]) out.push(parts[j]);
		if(!parts.length || (parts.length == 1 && !parts[0])) out.push(lines[i]);
	}
	return out.length ? out : [""];
}
function effective_cell_style(ws/*:Worksheet*/, cell/*:?Cell*/, R/*:number*/, C/*:number*/) {
	var out = ({}/*:any*/), cols = ws["!cols"] || [], rows = ws["!rows"] || [];
	if(cols[C] && cols[C].s) extend_style_obj(out, cols[C].s);
	if(rows[R] && rows[R].s) extend_style_obj(out, rows[R].s);
	if(cell && cell.s) extend_style_obj(out, cell.s);
	return out;
}
function measure_cell_text_width(cell/*:Cell*/, style/*:CellStyle*/, opts/*:?any*/)/*:number*/ {
	var o = opts || {}, text = auto_fit_cell_text(cell, o), align = style && style.alignment || {};
	var wrap = !!align.wrapText, segments = text_width_segments(text, wrap), max = 0, w = 0;
	for(var i = 0; i < segments.length; ++i) {
		w = measure_text_width(segments[i], style, o);
		if(w > max) max = w;
	}
	var indent = +align.indent || 0;
	if(indent > 0) max += indent * 3 * Math.max(1, MDW);
	var rotation = +align.textRotation || 0;
	if(rotation == 255) rotation = 90;
	if(rotation > 90) rotation = 90 - rotation;
	if(rotation < 0) rotation = -rotation;
	if(rotation > 0 && rotation < 90) {
		var rad = rotation * Math.PI / 180, fpx = pt2px_browser(style_font_size_pt(style));
		max = Math.abs(max * Math.cos(rad)) + Math.abs(fpx * Math.sin(rad));
	} else if(rotation >= 90) max = Math.max(MDW + XLSX_COL_WIDTH_PADDING, pt2px_browser(style_font_size_pt(style)) + 2);
	return max + (o.padding != null ? +o.padding : XLSX_COL_WIDTH_PADDING);
}
function merge_start_map(merges) {
	var out = {};
	for(var i = 0; i < merges.length; ++i) out[merges[i].s.r + ":" + merges[i].s.c] = merges[i];
	return out;
}
function is_covered_merge(merges, R, C) {
	for(var i = 0; i < merges.length; ++i) {
		var m = merges[i];
		if(m.s.r <= R && R <= m.e.r && m.s.c <= C && C <= m.e.c) return !(m.s.r == R && m.s.c == C);
	}
	return false;
}
function auto_fit_columns(ws/*:Worksheet*/, opts/*:?any*/)/*:Array<ColInfo>*/ {
	var o = opts || {}, oldMDW = MDW;
	if(o.MDW) MDW = +o.MDW || MDW;
	var r = o.range ? (typeof o.range == "string" ? decode_range(o.range) : o.range) : decode_range(ws["!ref"] || "A1");
	var base = ws["!cols"] || [], rows = ws["!rows"] || [], cols = [], widths = [];
	var minpx = o.minPx != null ? +o.minPx : o.min != null ? width2px(char2width(+o.min)) : 0;
	var maxpx = o.maxPx != null ? +o.maxPx : width2px(XLSX_COL_WIDTH_MAX);
	var merges = o.includeMerged === false ? [] : (ws["!merges"] || []);
	var starts = merge_start_map(merges), dense = ws["!data"] != null;
	var C = 0, R = 0, span = 0, m = null, j = 0, cell = null, style = null, need = 0, have = 0;
	for(C = r.s.c; C <= r.e.c; ++C) {
		cols[C] = dup(base[C] || {});
		if(cols[C] && cols[C].hidden && o.skipHidden) continue;
		widths[C] = cols[C] && (cols[C].wpx != null || cols[C].width != null || cols[C].wch != null) ? col2px(cols[C]) : minpx;
	}
	for(R = r.s.r; R <= r.e.r; ++R) {
		if(rows[R] && rows[R].hidden && o.skipHidden) continue;
		for(C = r.s.c; C <= r.e.c; ++C) {
			if(cols[C] && cols[C].hidden && o.skipHidden) continue;
			if(is_covered_merge(merges, R, C)) continue;
			cell = dense ? (ws["!data"][R]||[])[C] : ws[encode_cell({r:R,c:C})];
			if(!cell || cell.v == null) continue;
			style = effective_cell_style(ws, cell, R, C);
			need = measure_cell_text_width(cell, style, o);
			if(style && style.alignment && style.alignment.shrinkToFit && widths[C]) need = Math.min(need, widths[C]);
			m = starts[R + ":" + C];
			span = m ? m.e.c - m.s.c + 1 : 1;
			if(span > 1) {
				have = 0;
				for(j = 0; j < span; ++j) have += widths[C+j] || minpx;
				if(need > have) for(j = 0; j < span; ++j) widths[C+j] = Math.min(maxpx, Math.max(widths[C+j] || minpx, (widths[C+j] || minpx) + (need - have) / span));
			} else widths[C] = Math.min(maxpx, Math.max(widths[C] || minpx, need));
		}
	}
	for(C = r.s.c; C <= r.e.c; ++C) if(widths[C] != null) cols[C] = set_col_width_from_px(cols[C] || {}, Math.min(maxpx, Math.max(minpx, widths[C])));
	if(o.set === false) { MDW = oldMDW; return cols; }
	ws["!cols"] = cols;
	MDW = oldMDW;
	return cols;
}

var DEF_PPI = 96, PPI = DEF_PPI;
function px2pt(px) { return px * 96 / PPI; }
function pt2px(pt) { return pt * PPI / 96; }
function pt2px_browser(pt) { return pt * 96 / 72; }
function px2pt_browser(px) { return px * 72 / 96; }

/* [MS-EXSPXML3] 2.4.54 ST_enmPattern */
var XLMLPatternTypeMap = {
	"None": "none",
	"Solid": "solid",
	"Gray50": "mediumGray",
	"Gray75": "darkGray",
	"Gray25": "lightGray",
	"HorzStripe": "darkHorizontal",
	"VertStripe": "darkVertical",
	"ReverseDiagStripe": "darkDown",
	"DiagStripe": "darkUp",
	"DiagCross": "darkGrid",
	"ThickDiagCross": "darkTrellis",
	"ThinHorzStripe": "lightHorizontal",
	"ThinVertStripe": "lightVertical",
	"ThinReverseDiagStripe": "lightDown",
	"ThinHorzCross": "lightGrid"
};
