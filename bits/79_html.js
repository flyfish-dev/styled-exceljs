/* note: browser DOM element cannot see mso- style attrs, must parse */
function html_to_sheet(str/*:string*/, _opts)/*:Workbook*/ {
	var opts = _opts || {};
	var dense = (opts.dense != null) ? opts.dense : DENSE;
	var ws/*:Worksheet*/ = ({}/*:any*/); if(dense) ws["!data"] = [];
	str = str_remove_ng(str, "<!--", "-->");
	var mtch/*:any*/ = str.match(/<table/i);
	if(!mtch) throw new Error("Invalid HTML: could not find <table>");
	var mtch2/*:any*/ = str.match(/<\/table/i);
	var i/*:number*/ = mtch.index, j/*:number*/ = mtch2 && mtch2.index || str.length;
	var rows = split_regex(str.slice(i, j), /(:?<tr[^<>]*>)/i, "<tr>");
	var R = -1, C = 0, RS = 0, CS = 0;
	var range/*:Range*/ = {s:{r:10000000, c:10000000},e:{r:0,c:0}};
	var merges/*:Array<Range>*/ = [];
	for(i = 0; i < rows.length; ++i) {
		var row = rows[i].trim();
		var hd = row.slice(0,3).toLowerCase();
		if(hd == "<tr") { ++R; if(opts.sheetRows && opts.sheetRows <= R) { --R; break; } C = 0; continue; }
		if(hd != "<td" && hd != "<th") continue;
		var cells = row.split(/<\/t[dh]>/i);
		for(j = 0; j < cells.length; ++j) {
			var cell = cells[j].trim();
			if(!cell.match(/<t[dh]/i)) continue;
			var m = cell, cc = 0;
			/* TODO: parse styles etc */
			while(m.charAt(0) == "<" && (cc = m.indexOf(">")) > -1) m = m.slice(cc+1);
			for(var midx = 0; midx < merges.length; ++midx) {
				var _merge/*:Range*/ = merges[midx];
				if(_merge.s.c == C && _merge.s.r < R && R <= _merge.e.r) { C = _merge.e.c + 1; midx = -1; }
			}
			var tag = parsexmltag(cell.slice(0, cell.indexOf(">")));
			CS = tag.colspan ? +tag.colspan : 1;
			if((RS = +tag.rowspan)>1 || CS>1) merges.push({s:{r:R,c:C},e:{r:R + (RS||1) - 1, c:C + CS - 1}});
			var _t/*:string*/ = tag.t || tag["data-t"] || "";
			/* TODO: generate stub cells */
			if(!m.length) { C += CS; continue; }
			m = htmldecode(m);
			if(range.s.r > R) range.s.r = R; if(range.e.r < R) range.e.r = R;
			if(range.s.c > C) range.s.c = C; if(range.e.c < C) range.e.c = C;
			if(!m.length) { C += CS; continue; }
			var o/*:Cell*/ = {t:'s', v:m};
			if(opts.raw || !m.trim().length || _t == 's'){}
			else if(m === 'TRUE') o = {t:'b', v:true};
			else if(m === 'FALSE') o = {t:'b', v:false};
			else if(!isNaN(fuzzynum(m))) o = {t:'n', v:fuzzynum(m)};
			else if(!isNaN(fuzzydate(m).getDate())) {
				o = ({t:'d', v:parseDate(m)}/*:any*/);
				if(opts.UTC === false) o.v = utc_to_local(o.v);
				if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)}/*:any*/);
				o.z = opts.dateNF || table_fmt[14];
			} else if(m.charCodeAt(0) == 35 /* # */ && RBErr[m] != null) {
				o.t = 'e'; o.w = m; o.v = RBErr[m];
			}
			if(o.cellText !== false) o.w = m;
			if(dense) { if(!ws["!data"][R]) ws["!data"][R] = []; ws["!data"][R][C] = o; }
			else ws[encode_cell({r:R, c:C})] = o;
			C += CS;
		}
	}
	ws['!ref'] = encode_range(range);
	if(merges.length) ws["!merges"] = merges;
	return ws;
}

function css_color(color) {
	if(!color) return "";
	if(color.rgb && /^[0-9A-Fa-f]{6}(?:[0-9A-Fa-f]{2})?$/.test(String(color.rgb))) return "#" + String(color.rgb).slice(-6);
	return "";
}
function css_font_family(name) {
	if(!name) return "";
	return "'" + css_string_escape(name, "'") + "'";
}

function html_attr_escape(value/*:any*/)/*:string*/ {
	return String(value).replace(/[&<>"']/g, function(ch) {
		return ch == "&" ? "&amp;" : ch == "<" ? "&lt;" : ch == ">" ? "&gt;" : ch == '"' ? "&quot;" : "&#39;";
	});
}
function html_writextag(tag/*:string*/, content/*:?string*/, attrs)/*:string*/ {
	var out = ["<", tag];
	if(attrs) keys(attrs).forEach(function(key) { out.push(" ", key, '="', html_attr_escape(attrs[key]), '"'); });
	if(content == null) { out.push("/>"); return out.join(""); }
	out.push(">", content, "</", tag, ">");
	return out.join("");
}
function safe_html_href(value/*:any*/)/*:?string*/ {
	var href = String(value == null ? "" : value).trim();
	if(!href || href.slice(0, 2) == "//") return null;
	var colon = href.indexOf(":"), prefix = colon == -1 ? "" : href.slice(0, colon), compact = "";
	for(var i = 0; i < prefix.length; ++i) if(prefix.charCodeAt(i) > 32 && prefix.charCodeAt(i) != 127) compact += prefix.charAt(i);
	if(colon != -1 && !/^(?:https?|mailto|tel)$/i.test(compact)) return null;
	return href;
}
function safe_html_image_src(value/*:any*/)/*:?string*/ {
	var src = String(value == null ? "" : value);
	return /^data:image\/[A-Za-z0-9.+-]+;base64,[A-Za-z0-9+/=\r\n]+$/.test(src) ? src : null;
}
function sanitize_cell_html(value/*:any*/)/*:string*/ {
	var html = String(value == null ? "" : value), out = [], pos = 0;
	while(pos < html.length) {
		var start = html.indexOf("<", pos);
		if(start == -1) { out.push(html.slice(pos)); break; }
		out.push(html.slice(pos, start));
		var end = html.indexOf(">", start + 1);
		if(end == -1) { out.push("&lt;", html.slice(start + 1)); break; }
		var raw = html.slice(start + 1, end).trim(), lower = raw.toLowerCase();
		if(/^(?:\/?(?:b|i|s|sup|sub)|br\s*\/?)$/.test(lower)) out.push("<", lower == "br" ? "br/" : lower, ">");
		else if(lower == "/span") out.push("</span>");
		else if(lower.slice(0, 12) == 'span style="' && raw.charAt(raw.length - 1) == '"') {
			var declarations = raw.slice(12, -1).split(";"), safe = [];
			declarations.forEach(function(decl) {
				var colon = decl.indexOf(":"), key = colon == -1 ? "" : decl.slice(0, colon).trim().toLowerCase(), val = colon == -1 ? "" : decl.slice(colon + 1).trim().toLowerCase();
				if(key == "text-decoration" && val == "underline") safe.push("text-decoration:underline");
				else if(key == "text-underline-style" && /^(?:single|double|single-accounting|double-accounting)$/.test(val)) safe.push("text-underline-style:" + val);
				else if(key == "font-size" && /^\d+(?:\.\d+)?pt$/.test(val)) safe.push("font-size:" + val);
				else if(key == "text-effect" && val == "outline") safe.push("text-effect:outline");
				else if(key == "text-shadow" && val == "auto") safe.push("text-shadow:auto");
			});
			out.push('<span style="', safe.join(";"), safe.length ? ";" : "", '">');
		} else out.push(html_attr_escape(html.slice(start, end + 1)));
		pos = end + 1;
	}
	return out.join("");
}
function html_border_style(style) {
	switch(style) {
		case "dashDot": case "dashDotDot": case "dashed": case "mediumDashed": return "dashed";
		case "dotted": case "hair": return "dotted";
		case "double": return "double";
		case "none": return "none";
		default: return "solid";
	}
}
function html_border_width(style) {
	switch(style) {
		case "medium": case "mediumDashDot": case "mediumDashDotDot": case "mediumDashed": return "2px";
		case "thick": return "3px";
		case "hair": return "1px";
		default: return "1px";
	}
}
function add_html_border(css, side, border) {
	if(!border || !border.style || border.style == "none") return;
	var color = css_color(border.color) || "#000000";
	css.push("border-" + side + ":" + html_border_width(border.style) + " " + html_border_style(border.style) + " " + color);
}
function html_cell_style(cell/*:Cell*/, opts/*:Sheet2HTMLOpts*/) {
	if(!opts || !opts.cellStyles || !cell || !cell.s) return "";
	var s = cell.s, css = [];
	var font = s.font || {};
	if(font.name) css.push("font-family:" + css_font_family(font.name));
	var fontSize = +font.sz;
	if(isFinite(fontSize) && fontSize > 0) css.push("font-size:" + Math.min(fontSize, 409) + "pt");
	if(font.bold) css.push("font-weight:bold");
	if(font.italic) css.push("font-style:italic");
	var deco = [];
	if(font.underline) deco.push("underline");
	if(font.strike) deco.push("line-through");
	if(deco.length) css.push("text-decoration:" + deco.join(" "));
	if(font.color && css_color(font.color)) css.push("color:" + css_color(font.color));
	var fill = s.fill || s;
	var fillColor = "";
	if(fill.patternType != "none" && fill.patternType != "gray125") fillColor = css_color(fill.fgColor) || css_color(fill.bgColor);
	if(fillColor) css.push("background-color:" + fillColor);
	var alignment = s.alignment || {};
	if(/^(?:left|right|center|justify|fill|distributed)$/.test(alignment.horizontal || "")) css.push("text-align:" + alignment.horizontal);
	var vertical = alignment.vertical == "center" ? "middle" : alignment.vertical;
	if(/^(?:top|middle|bottom|baseline)$/.test(vertical || "")) css.push("vertical-align:" + vertical);
	if(alignment.textRotation != null && alignment.textRotation !== 0) {
		var deg = alignment.textRotation == 255 ? 90 : alignment.textRotation > 90 ? 90 - alignment.textRotation : alignment.textRotation;
		css.push("transform:rotate(" + deg + "deg)");
		css.push("transform-origin:center");
	}
	var border = s.border || {};
	add_html_border(css, "left", border.left || border.start);
	add_html_border(css, "right", border.right || border.end);
	add_html_border(css, "top", border.top);
	add_html_border(css, "bottom", border.bottom);
	return css.join(";");
}
function html_cell_layout_style(cell/*:Cell*/, opts/*:Sheet2HTMLOpts*/, C/*:number*/, CS/*:number*/, cols/*:Array<ColInfo>*/) {
	if(!opts || !cell) return "";
	var s = cell.s || {}, alignment = s.alignment || {}, css = [];
	if(opts.browserPixels || opts.autoFit) css.push("box-sizing:border-box;padding:0 2px;min-width:0");
	if(alignment.wrapText) css.push("white-space:pre-wrap;overflow-wrap:normal;word-break:normal");
	else {
		css.push("white-space:pre");
		if(alignment.shrinkToFit || opts.overflow == "clip" || opts.overflow == "hidden") css.push("overflow:hidden;text-overflow:clip");
		else css.push("overflow:visible");
	}
	if(alignment.shrinkToFit && cols && C != null) {
		var have = 0, span = CS || 1;
		for(var j = 0; j < span; ++j) have += html_col_width(cols[C+j]);
		var need = measure_cell_text_width(cell, s, opts);
		if(have > 0 && need > have) {
			var scale = Math.max(0.25, Math.min(1, have / need));
			css.push("font-size:" + Math.max(1, style_font_size_pt(s) * scale) + "pt");
		}
	}
	return css.join(";");
}
function html_col_width(col) {
	return col2px(col);
}
function html_row_height(row) {
	if(!row) return 20;
	if(row.hpx != null) return row.hpx;
	if(row.hpt != null) return pt2px_browser(row.hpt);
	return 20;
}

function make_html_row(ws/*:Worksheet*/, r/*:Range*/, R/*:number*/, o/*:Sheet2HTMLOpts*/)/*:string*/ {
	var M/*:Array<Range>*/ = (ws['!merges'] ||[]);
	var oo/*:Array<string>*/ = [];
	var sp = ({}/*:any*/);
	var dense = ws["!data"] != null;
	var row = (ws["!rows"]||[])[R];
	var cols = o && o._htmlCols || ws["!cols"] || [];
	for(var C = r.s.c; C <= r.e.c; ++C) {
		var RS = 0, CS = 0;
		for(var j = 0; j < M.length; ++j) {
			if(M[j].s.r > R || M[j].s.c > C) continue;
			if(M[j].e.r < R || M[j].e.c < C) continue;
			if(M[j].s.r < R || M[j].s.c < C) { RS = -1; break; }
			RS = M[j].e.r - M[j].s.r + 1; CS = M[j].e.c - M[j].s.c + 1; break;
		}
		if(RS < 0) continue;
		var coord = encode_col(C) + encode_row(R);
		var cell = dense ? (ws["!data"][R]||[])[C] : ws[coord];
		var stylecell = cell;
		if(cell && cell.t == 'n' && cell.v != null && !isFinite(cell.v)) {
			if(isNaN(cell.v)) cell = ({t:'e', v:0x24, w:BErr[0x24]});
			else cell = ({t:'e', v:0x07, w:BErr[0x07]});
			stylecell = cell;
		}
		if(o.cellStyles) {
			var inherited = ({}/*:any*/);
			if(cols[C] && cols[C].s) extend_style_obj(inherited, cols[C].s);
			if(row && row.s) extend_style_obj(inherited, row.s);
			if(cell && cell.s) extend_style_obj(inherited, cell.s);
			inherited = resolve_table_cell_style(ws, R, C, inherited) || inherited;
			if(keys(inherited).length) {
				stylecell = cell ? dup(cell) : {t:'z'};
				stylecell.s = inherited;
			}
		}
		/* TODO: html entities */
		var w = (cell && cell.v != null) && (cell.h ? sanitize_cell_html(cell.h) : escapehtml(cell.w || (format_cell(cell), cell.w) || "")) || "";
		sp = ({}/*:any*/);
		if(RS > 1) sp.rowspan = RS;
		if(CS > 1) sp.colspan = CS;
		if(o.editable) w = '<span contenteditable="true">' + w + '</span>';
		else if(cell) {
			sp["data-t"] = cell && cell.t || 'z';
			// note: data-v is unaffected by the timezone interpretation
			if(cell.v != null) sp["data-v"] = cell.v instanceof Date ? cell.v.toISOString() : cell.v;
			if(cell.z != null) sp["data-z"] = cell.z;
			if(cell.f != null) sp["data-f"] = cell.f;
			var href = cell.l && safe_html_href(cell.l.Target);
			if(href) w = '<a href="' + html_attr_escape(href) +'">' + w + '</a>';
		}
		var cstyle = html_cell_style(stylecell, o);
		var lstyle = html_cell_layout_style(stylecell, o, C, CS, cols);
		if(lstyle) cstyle = cstyle ? cstyle + ";" + lstyle : lstyle;
		if(cstyle) sp.style = cstyle;
		sp.id = (o.id || "sjs") + "-" + coord;
		oo.push(html_writextag('td', w, sp));
	}
	var rsp = ({}/*:any*/), rstyle = [];
	if(row) {
		if(row.hidden) rstyle.push("display:none");
		if(o.browserPixels) rstyle.push("height:" + html_row_height(row) + "px");
	}
	if(rstyle.length) rsp.style = rstyle.join(";");
	return html_writextag('tr', oo.join(""), rsp);
}

var HTML_BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
var HTML_END = '</body></html>';

function html_to_workbook(str/*:string*/, opts)/*:Workbook*/ {
	var mtch = str_match_xml_ig(str, "table");
	if(!mtch || mtch.length == 0) throw new Error("Invalid HTML: could not find <table>");
	if(mtch.length == 1) {
		var w = sheet_to_workbook(html_to_sheet(mtch[0], opts), opts);
		w.bookType = "html";
		return w;
	}
	var wb = book_new();
	mtch.forEach(function(s, idx) { book_append_sheet(wb, html_to_sheet(s, opts), "Sheet" + (idx+1)); });
	wb.bookType = "html";
	return wb;
}

var HTML_EMU_PER_PIXEL = 9525;
function html_anchor_pos(ws, anchor, opts) {
	var out = {left:0, top:0, width:480, height:288};
	var cols = opts && opts._htmlCols || ws["!cols"] || [], rows = ws["!rows"] || [];
	var from = anchor && anchor.from || {col:0,row:0,colOff:0,rowOff:0};
	var to = anchor && anchor.to;
	for(var C = 0; C < (from.col||0); ++C) out.left += html_col_width(cols[C]);
	for(var R = 0; R < (from.row||0); ++R) out.top += html_row_height(rows[R]);
	out.left += (from.colOff||0) / HTML_EMU_PER_PIXEL;
	out.top += (from.rowOff||0) / HTML_EMU_PER_PIXEL;
	if(to) {
		var right = 0, bottom = 0;
		for(C = 0; C < (to.col||0); ++C) right += html_col_width(cols[C]);
		for(R = 0; R < (to.row||0); ++R) bottom += html_row_height(rows[R]);
		right += (to.colOff||0) / HTML_EMU_PER_PIXEL;
		bottom += (to.rowOff||0) / HTML_EMU_PER_PIXEL;
		out.width = Math.max(1, right - out.left);
		out.height = Math.max(1, bottom - out.top);
	} else if(anchor && anchor.ext) {
		if(anchor.ext.cx) out.width = anchor.ext.cx / HTML_EMU_PER_PIXEL;
		if(anchor.ext.cy) out.height = anchor.ext.cy / HTML_EMU_PER_PIXEL;
	}
	return out;
}
function html_abs_style(pos) {
	return "position:absolute;left:" + Math.round(pos.left) + "px;top:" + Math.round(pos.top) + "px;width:" + Math.round(pos.width) + "px;height:" + Math.round(pos.height) + "px";
}
function chart_series_values(ser) {
	if(!ser) return [];
	if(ser.val && ser.val.values) return ser.val.values;
	if(ser.yVal && ser.yVal.values) return ser.yVal.values;
	if(ser.data && ser.data.length) return ser.data;
	return [];
}
function chart_series_labels(ser) {
	if(!ser) return [];
	if(ser.cat && ser.cat.values) return ser.cat.values;
	if(ser.xVal && ser.xVal.values) return ser.xVal.values;
	return [];
}
function render_chart_svg(chart, width, height) {
	var model = chart && (chart.model || chart["!chart"] || chart);
	if(!model) return "";
	var series = model.series || [], type = (model.type || "").replace(/Chart$/, "");
	var w = Math.max(160, Math.round(width || 480)), h = Math.max(120, Math.round(height || 288));
	var title = model.title || chart.title || "";
	var colors = ["#4F81BD","#C0504D","#9BBB59","#8064A2","#4BACC6","#F79646"];
	var out = ['<svg class="sjs-chart-svg" xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + h + '" viewBox="0 0 ' + w + ' ' + h + '">'];
	out.push('<rect x="0" y="0" width="' + w + '" height="' + h + '" fill="#fff" stroke="#d0d7de"/>');
	if(title) out.push('<text x="' + (w/2) + '" y="20" text-anchor="middle" font-family="Arial" font-size="14">' + escapehtml(title) + '</text>');
	var top = title ? 34 : 16, left = 42, right = 14, bottom = 28;
	var pw = w - left - right, ph = h - top - bottom;
	var all = [];
	series.forEach(function(ser) { chart_series_values(ser).forEach(function(v) { if(typeof v == "number" && isFinite(v)) all.push(v); }); });
	var max = Math.max.apply(Math, all.concat([0])), min = Math.min.apply(Math, all.concat([0]));
	if(min > 0) min = 0;
	if(max == min) max = min + 1;
	out.push('<line x1="' + left + '" y1="' + (top+ph) + '" x2="' + (left+pw) + '" y2="' + (top+ph) + '" stroke="#444"/>');
	out.push('<line x1="' + left + '" y1="' + top + '" x2="' + left + '" y2="' + (top+ph) + '" stroke="#444"/>');
	if(type == "pie" || type == "doughnut") {
		var vals = chart_series_values(series[0] || {});
		var total = vals.reduce(function(a,b) { return a + (typeof b == "number" ? Math.max(0,b) : 0); }, 0) || 1;
		var cx = w/2, cy = top + ph/2, r = Math.max(10, Math.min(pw, ph) / 2 - 8), a0 = -Math.PI/2;
		vals.forEach(function(v, i) {
			var a1 = a0 + (Math.max(0, +v || 0) / total) * Math.PI * 2;
			var x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
			var x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
			var large = a1 - a0 > Math.PI ? 1 : 0;
			out.push('<path d="M' + cx + ',' + cy + ' L' + x0 + ',' + y0 + ' A' + r + ',' + r + ' 0 ' + large + ',1 ' + x1 + ',' + y1 + ' Z" fill="' + colors[i%colors.length] + '"/>');
			a0 = a1;
		});
		if(type == "doughnut") out.push('<circle cx="' + cx + '" cy="' + cy + '" r="' + (r*0.45) + '" fill="#fff"/>');
		out.push('</svg>');
		return out.join("");
	}
	var count = 0;
	series.forEach(function(ser) { count = Math.max(count, chart_series_values(ser).length); });
	if(!count) { out.push('</svg>'); return out.join(""); }
	var xstep = pw / count;
	series.forEach(function(ser, si) {
		var vals = chart_series_values(ser), color = colors[si % colors.length];
		if(type == "line" || type == "scatter" || type == "area") {
			var pts = [];
			vals.forEach(function(v, i) {
				if(typeof v != "number" || !isFinite(v)) return;
				var x = left + xstep * (i + 0.5);
				var y = top + ph - (v - min) / (max - min) * ph;
				pts.push([x,y]);
			});
			if(type == "area" && pts.length) out.push('<polygon points="' + [[pts[0][0],top+ph]].concat(pts).concat([[pts[pts.length-1][0],top+ph]]).map(function(p) { return p[0] + "," + p[1]; }).join(" ") + '" fill="' + color + '" opacity="0.35"/>');
			if(pts.length) out.push('<polyline points="' + pts.map(function(p) { return p[0] + "," + p[1]; }).join(" ") + '" fill="none" stroke="' + color + '" stroke-width="2"/>');
			pts.forEach(function(p) { out.push('<circle cx="' + p[0] + '" cy="' + p[1] + '" r="2.5" fill="' + color + '"/>'); });
		} else {
			var bw = Math.max(1, xstep / Math.max(series.length, 1) * 0.72);
			vals.forEach(function(v, i) {
				if(typeof v != "number" || !isFinite(v)) return;
				var x = left + xstep * i + (xstep - bw * series.length) / 2 + bw * si;
				var y = top + ph - (v - min) / (max - min) * ph;
				var y0 = top + ph - (0 - min) / (max - min) * ph;
				out.push('<rect x="' + x + '" y="' + Math.min(y,y0) + '" width="' + bw + '" height="' + Math.max(1, Math.abs(y0-y)) + '" fill="' + color + '"/>');
			});
		}
		if(ser.name) out.push('<text x="' + (left + 8) + '" y="' + (top + 14 + si * 14) + '" font-family="Arial" font-size="11" fill="' + color + '">' + escapehtml(ser.name) + '</text>');
	});
	var labels = chart_series_labels(series[0] || {});
	labels.slice(0, Math.min(labels.length, count)).forEach(function(lbl, i) {
		if(i % Math.ceil(count / 8) != 0) return;
		out.push('<text x="' + (left + xstep * (i + 0.5)) + '" y="' + (top + ph + 16) + '" text-anchor="middle" font-family="Arial" font-size="10">' + escapehtml(String(lbl)) + '</text>');
	});
	out.push('</svg>');
	return out.join("");
}
function render_html_drawings(ws, opts) {
	var out = [], drawings = ws["!drawings"] || {}, charts = ws["!charts"] || [];
	if(opts && opts.drawings && drawings.images) drawings.images.forEach(function(img) {
		if(!img || !img.dataURI) return;
		var src = safe_html_image_src(img.dataURI);
		if(!src) return;
		var pos = html_anchor_pos(ws, img.anchor, opts);
		out.push('<img class="sjs-drawing-image" src="' + html_attr_escape(src) + '" style="' + html_abs_style(pos) + '"/>');
	});
	if(opts && opts.charts) charts.forEach(function(chart) {
		var pos = html_anchor_pos(ws, chart.anchor, opts);
		out.push('<div class="sjs-chart" style="' + html_abs_style(pos) + '">' + render_chart_svg(chart, pos.width, pos.height) + '</div>');
	});
	if(!out.length) return "";
	return '<div class="sjs-drawing-layer" style="position:absolute;left:0;top:0;pointer-events:none">' + out.join("") + '</div>';
}
function make_html_preamble(ws/*:Worksheet*/, R/*:Range*/, o/*:Sheet2HTMLOpts*/)/*:string*/ {
	var out/*:Array<string>*/ = [];
	var tattr = ({}/*:any*/), tstyle = [];
	if(o && o.id) tattr.id = o.id;
	if(o && (o.browserPixels || o.autoFit)) tstyle.push("border-collapse:collapse;table-layout:fixed");
	if(tstyle.length) tattr.style = tstyle.join(";");
	var table = html_writextag("table", "", tattr).replace(/<\/table>$/, "");
	var cols = o && o._htmlCols || ws["!cols"];
	if(o && (o.browserPixels || o.autoFit) && cols) {
		out.push("<colgroup>");
		for(var C = R.s.c; C <= R.e.c; ++C) {
			var col = cols[C], style = [];
			style.push("width:" + html_col_width(col) + "px");
			if(col && col.hidden) style.push("display:none");
			out.push(html_writextag("col", null, {style:style.join(";")}));
		}
		out.push("</colgroup>");
	}
	return table + out.join("");
}

function sheet_to_html(ws/*:Worksheet*/, opts/*:?Sheet2HTMLOpts*//*, wb:?Workbook*/)/*:string*/ {
	var o = {};
	if(opts) for(var k in opts) if(Object.prototype.hasOwnProperty.call(opts, k)) o[k] = opts[k];
	if(o.autoFit) {
		var af = {set:false};
		if(typeof o.autoFit == "object") for(k in o.autoFit) if(Object.prototype.hasOwnProperty.call(o.autoFit, k)) af[k] = o.autoFit[k];
		if(o.measureText && af.measureText == null) af.measureText = o.measureText;
		if(o.canvas && af.canvas == null) af.canvas = o.canvas;
		o._htmlCols = auto_fit_columns(ws, af);
		if(o.browserPixels == null) o.browserPixels = true;
	}
	var header = o.header != null ? o.header : HTML_BEGIN;
	var footer = o.footer != null ? o.footer : HTML_END;
	var out/*:Array<string>*/ = [header];
	var r = decode_range(ws['!ref'] || "A1");
	var draw = (o.charts && ws["!charts"] && ws["!charts"].length) || (o.drawings && ws["!drawings"]);
	if(draw) out.push('<div class="sjs-sheet" style="position:relative;display:inline-block">');
	out.push(make_html_preamble(ws, r, o));
	if(ws["!ref"]) for(var R = r.s.r; R <= r.e.r; ++R) out.push(make_html_row(ws, r, R, o));
	if(draw) {
		out.push("</table>");
		out.push(render_html_drawings(ws, o));
		out.push("</div>");
		out.push(footer);
	} else out.push("</table>" + footer);
	return out.join("");
}

function sheet_add_dom(ws/*:Worksheet*/, table/*:HTMLElement*/, _opts/*:?any*/)/*:Worksheet*/ {
	var opts = _opts || {};
	var dense = ws["!data"] != null;
	var or_R = 0, or_C = 0;
	if(opts.origin != null) {
		if(typeof opts.origin == 'number') or_R = opts.origin;
		else {
			var _origin/*:CellAddress*/ = typeof opts.origin == "string" ? decode_cell(opts.origin) : opts.origin;
			or_R = _origin.r; or_C = _origin.c;
		}
	}

	var range/*:Range*/ = {s:{r:0,c:0},e:{r:or_R,c:or_C}};
	if(ws["!ref"]) {
		var _range/*:Range*/ = decode_range(ws["!ref"]);
		range.s.r = Math.min(range.s.r, _range.s.r);
		range.s.c = Math.min(range.s.c, _range.s.c);
		range.e.r = Math.max(range.e.r, _range.e.r);
		range.e.c = Math.max(range.e.c, _range.e.c);
		if(or_R == -1) range.e.r = or_R = _range.e.r + 1;
	}


	var rows/*:HTMLCollection<HTMLTableRowElement>*/ = table.rows;
	if(!rows) {
		/* not an HTML TABLE */
		throw "Unsupported origin when " + table.tagName + " is not a TABLE";
	}
	var sheetRows = Math.min(opts.sheetRows||10000000, rows.length);

	var merges/*:Array<Range>*/ = [], midx = 0;
	var rowinfo/*:Array<RowInfo>*/ = ws["!rows"] || (ws["!rows"] = []);
	var _R = 0, R = 0, _C = 0, C = 0, RS = 0, CS = 0;
	if(!ws["!cols"]) ws['!cols'] = [];
	for(; _R < rows.length && R < sheetRows; ++_R) {
		var row/*:HTMLTableRowElement*/ = rows[_R];
		if (is_dom_element_hidden(row)) {
			if (opts.display) continue;
			rowinfo[R] = {hidden: true};
		}
		var elts/*:HTMLCollection<HTMLTableCellElement>*/ = (row.cells);
		for(_C = C = 0; _C < elts.length; ++_C) {
			var elt/*:HTMLTableCellElement*/ = elts[_C];
			if (opts.display && is_dom_element_hidden(elt)) continue;
			var v/*:?string*/ = elt.hasAttribute('data-v') ? elt.getAttribute('data-v') : elt.hasAttribute('v') ? elt.getAttribute('v') : htmldecode(elt.innerHTML);
			var z/*:?string*/ = elt.getAttribute('data-z') || elt.getAttribute('z');
			var f/*:?string*/ = elt.hasAttribute('data-f') ? elt.getAttribute('data-f') : elt.hasAttribute('f') ? elt.getAttribute('f') : null;
			for(midx = 0; midx < merges.length; ++midx) {
				var m/*:Range*/ = merges[midx];
				if(m.s.c == C + or_C && m.s.r < R + or_R && R + or_R <= m.e.r) { C = m.e.c+1 - or_C; midx = -1; }
			}
			/* TODO: figure out how to extract nonstandard mso- style */
			CS = +elt.getAttribute("colspan") || 1;
			if( ((RS = (+elt.getAttribute("rowspan") || 1)))>1 || CS>1) {
				merges.push({s:{r:R + or_R,c:C + or_C},e:{r:R + or_R + (RS||1) - 1, c:C + or_C + (CS||1) - 1}});
			}
			var o/*:Cell*/ = {t:'s', v:v};
			var _t/*:string*/ = elt.getAttribute("data-t") || elt.getAttribute("t") || "";
			if(v != null) {
				if(v.length == 0) o.t = _t || 'z';
				else if(opts.raw || v.trim().length == 0 || _t == "s"){}
				else if(_t == "e" && BErr[+v]) o = {t:'e', v:+v, w: BErr[+v]};
				else if(v === 'TRUE') o = {t:'b', v:true};
				else if(v === 'FALSE') o = {t:'b', v:false};
				else if(!isNaN(fuzzynum(v))) o = {t:'n', v:fuzzynum(v)};
				else if(!isNaN(fuzzydate(v).getDate())) {
					o = ({t:'d', v:parseDate(v)}/*:any*/);
					if(opts.UTC) o.v = local_to_utc(o.v);
					if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)}/*:any*/);
					o.z = opts.dateNF || table_fmt[14];
				} else if(v.charCodeAt(0) == 35 /* # */ && RBErr[v] != null) o = ({t:'e', v: RBErr[v], w: v});
			}
			if(o.z === undefined && z != null) o.z = z;
			/* The first link is used.  Links are assumed to be fully specified.
			 * TODO: The right way to process relative links is to make a new <a> */
			var l = "", Aelts = elt.getElementsByTagName("A");
			if(Aelts && Aelts.length) for(var Aelti = 0; Aelti < Aelts.length; ++Aelti)	if(Aelts[Aelti].hasAttribute("href")) {
				l = Aelts[Aelti].getAttribute("href"); if(l.charAt(0) != "#") break;
			}
			if(l && l.charAt(0) != "#" &&	l.slice(0, 11).toLowerCase() != 'javascript:') o.l = ({ Target: l });
			if(f != null) o.f = f;
			if(dense) { if(!ws["!data"][R + or_R]) ws["!data"][R + or_R] = []; ws["!data"][R + or_R][C + or_C] = o; }
			else ws[encode_cell({c:C + or_C, r:R + or_R})] = o;
			if(range.e.c < C + or_C) range.e.c = C + or_C;
			C += CS;
		}
		++R;
	}
	if(merges.length) ws['!merges'] = (ws["!merges"] || []).concat(merges);
	range.e.r = Math.max(range.e.r, R - 1 + or_R);
	ws['!ref'] = encode_range(range);
	if(R >= sheetRows) ws['!fullref'] = encode_range((range.e.r = rows.length-_R+R-1 + or_R,range)); // We can count the real number of rows to parse but we don't to improve the performance
	return ws;
}

function parse_dom_table(table/*:HTMLElement*/, _opts/*:?any*/)/*:Worksheet*/ {
	var opts = _opts || {};
	var ws/*:Worksheet*/ = ({}/*:any*/); if(opts.dense) ws["!data"] = [];
	return sheet_add_dom(ws, table, _opts);
}

function table_to_book(table/*:HTMLElement*/, opts/*:?any*/)/*:Workbook*/ {
	var o = sheet_to_workbook(parse_dom_table(table, opts), opts);
	//o.bookType = "dom"; // TODO: define a type for this
	return o;
}

function is_dom_element_hidden(element/*:HTMLElement*/)/*:boolean*/ {
	var display/*:string*/ = '';
	var get_computed_style/*:?function*/ = get_get_computed_style_function(element);
	if(get_computed_style) display = get_computed_style(element).getPropertyValue('display');
	if(!display) display = element.style && element.style.display;
	return display === 'none';
}

/* global getComputedStyle */
function get_get_computed_style_function(element/*:HTMLElement*/)/*:?function*/ {
	// The proper getComputedStyle implementation is the one defined in the element window
	if(element.ownerDocument.defaultView && typeof element.ownerDocument.defaultView.getComputedStyle === 'function') return element.ownerDocument.defaultView.getComputedStyle;
	// If it is not available, try to get one from the global namespace
	if(typeof getComputedStyle === 'function') return getComputedStyle;
	return null;
}
