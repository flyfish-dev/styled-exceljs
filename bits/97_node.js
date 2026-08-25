var _Readable;
function set_readable(R) { _Readable = R; }

function write_csv_stream(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
	var stream = _Readable();
	var o = opts == null ? {} : opts;
	if(sheet == null || sheet["!ref"] == null) { stream.push(null); return stream; }
	var r = safe_decode_range(sheet["!ref"]);
	var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
	var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
	var row/*:?string*/ = "", cols/*:Array<string>*/ = [];
	var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
	var rowinfo/*:Array<RowInfo>*/ = o.skipHidden && sheet["!rows"] || [];
	for(var C = r.s.c; C <= r.e.c; ++C) if (!((colinfo[C]||{}).hidden)) cols[C] = encode_col(C);
	var R = r.s.r;
	var BOM = false, w = 0;
	stream._read = function() {
		if(!BOM) { BOM = true; return stream.push("\uFEFF"); }
		while(R <= r.e.r) {
			++R;
			if ((rowinfo[R-1]||{}).hidden) continue;
			row = make_csv_row(sheet, r, R-1, cols, fs, rs, FS, w, o);
			if(row != null) {
				if(row || (o.blankrows !== false)) return stream.push((w++ ? RS : "") + row);
			}
		}
		return stream.push(null);
	};
	return stream;
}

function write_html_stream(ws/*:Worksheet*/, opts/*:?Sheet2HTMLOpts*/) {
	var stream = _Readable();

	var o = opts || {};
	var header = o.header != null ? o.header : HTML_BEGIN;
	var footer = o.footer != null ? o.footer : HTML_END;
	stream.push(header);
	var r = decode_range(ws['!ref']);
	stream.push(make_html_preamble(ws, r, o));
	var R = r.s.r;
	var end = false;
	stream._read = function() {
		if(R > r.e.r) {
			if(!end) { end = true; stream.push("</table>" + footer); }
			return stream.push(null);
		}
		while(R <= r.e.r) {
			stream.push(make_html_row(ws, r, R, o));
			++R;
			break;
		}
	};
	return stream;
}

function write_json_stream(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
	var stream = _Readable({objectMode:true});

	if(sheet == null || sheet["!ref"] == null) { stream.push(null); return stream; }
	var val = {t:'n',v:0}, header = 0, offset = 1, hdr/*:Array<any>*/ = [], v=0, vv="";
	var r = {s:{r:0,c:0},e:{r:0,c:0}};
	var o = opts || {};
	var range = o.range != null ? o.range : sheet["!ref"];
	if(o.header === 1) header = 1;
	else if(o.header === "A") header = 2;
	else if(Array.isArray(o.header)) header = 3;
	switch(typeof range) {
		case 'string': r = safe_decode_range(range); break;
		case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
		default: r = range;
	}
	if(header > 0) offset = 0;
	var rr = encode_row(r.s.r);
	var cols/*:Array<string>*/ = [];
	var counter = 0;
	var dense = sheet["!data"] != null;
	var R = r.s.r, C = 0;
	var header_cnt = {};
	if(dense && !sheet["!data"][R]) sheet["!data"][R] = [];
	var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
	var rowinfo/*:Array<RowInfo>*/ = o.skipHidden && sheet["!rows"] || [];
	for(C = r.s.c; C <= r.e.c; ++C) {
		if(((colinfo[C]||{}).hidden)) continue;
		cols[C] = encode_col(C);
		val = dense ? sheet["!data"][R][C] : sheet[cols[C] + rr];
		switch(header) {
			case 1: hdr[C] = C - r.s.c; break;
			case 2: hdr[C] = cols[C]; break;
			case 3: hdr[C] = o.header[C - r.s.c]; break;
			default:
				if(val == null) val = {w: "__EMPTY", t: "s"};
				vv = v = format_cell(val, null, o);
				counter = header_cnt[v] || 0;
				if(!counter) header_cnt[v] = 1;
				else {
					do { vv = v + "_" + (counter++); } while(header_cnt[vv]); header_cnt[v] = counter;
					header_cnt[vv] = 1;
				}
				hdr[C] = vv;
		}
	}
	R = r.s.r + offset;
	stream._read = function() {
		while(R <= r.e.r) {
			if ((rowinfo[R]||{}).hidden) {
				++R;
				continue;
			}
			var row = make_json_row(sheet, r, R, cols, header, hdr, o);
			++R;
			if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) {
				stream.push(row.row);
				return;
			}
		}
		return stream.push(null);
	};
	return stream;
}

function write_xlml_stream(wb/*:Workbook*/, o/*:?Sheet2XLMLOpts*/) {
	var stream = _Readable();
	var opts = o == null ? {} : o;
	var stride = +opts.stride || 10;
	if(!wb.SSF) wb.SSF = dup(table_fmt);
	if(wb.SSF) {
		make_ssf(); SSF_load_table(wb.SSF);
		// $FlowIgnore
		opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
		opts.ssf = wb.SSF;
		opts.cellXfs = [];
		get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});
	}

	/* do one pass to determine styles since they must be added before tables */
	wb.SheetNames.forEach(function(n) {
		var ws = sheet_map_get(wb.Sheets, n);
		if(!ws || !ws["!ref"]) return;
		var range = decode_range(ws["!ref"]);
		var dense = ws["!data"] != null;
		var ddata = dense ? ws["!data"] : [];
		var addr = {r:0,c:0};
		for(var R = range.s.r; R <= range.e.r; ++R) {
			addr.r = R;
			if(dense && !ddata[R]) continue;
			for(var C = range.s.c; C <= range.e.c; ++C) {
				addr.c = C;
				var cell = dense ? ddata[R][C] : ws[encode_col(C) + encode_row(R)];
				if(!cell) continue;
				if(cell.t == "d" && cell.z == null) { cell = dup(cell); cell.z = table_fmt[14]; }
				void get_cell_style(opts.cellXfs, cell, opts);
			}
		}
	});
	var sty = write_sty_xlml(wb, opts);

	var stage = 0, wsidx = 0, ws = sheet_map_get(wb.Sheets, wb.SheetNames[wsidx]), range = safe_decode_range(ws), R = -1, T = false;

	var marr = [], mi = 0, dense = false, darr = [], addr = {r:0,c:0};

	stream._read = function() { switch(stage) {
		/* header */
		case 0: {
			stage = 1;
			stream.push(XML_HEADER);
			stream.push("<Workbook" + wxt_helper({
				'xmlns':      XLMLNS.ss,
				'xmlns:o':    XLMLNS.o,
				'xmlns:x':    XLMLNS.x,
				'xmlns:ss':   XLMLNS.ss,
				'xmlns:dt':   XLMLNS.dt,
				'xmlns:html': XLMLNS.html
			}) + ">");
		} break;

		/* preamble */
		case 1: {
			stage = 2;
			stream.push(write_props_xlml(wb, opts));
			stream.push(write_wb_xlml(wb, opts));
		} break;

		/* style and name tables */
		case 2: {
			stage = 3;
			stream.push(sty);
			stream.push(write_names_xlml(wb, opts));
		} break;

		/* worksheet preamble */
		case 3: {
			T = false;
			if(wsidx >= wb.SheetNames.length) { stage = -1; stream.push(""); break; }

			stream.push("<Worksheet" + wxt_helper({ "ss:Name": escapexml(wb.SheetNames[wsidx])}) + ">");

			ws = sheet_map_get(wb.Sheets, wb.SheetNames[wsidx]);
			if(!ws) { stream.push("</Worksheet>"); return void ++wsidx; }

			var names = write_ws_xlml_names(ws, opts, wsidx, wb);
			if(names.length) stream.push("<Names>" + names + "</Names>");

			if(!ws["!ref"]) return (stage = 5);
			range = safe_decode_range(ws["!ref"]);
			R = range.s.r;
			stage = 4;
		} break;

		/* worksheet intramble */
		case 4: {
			if(R < 0 || R > range.e.r) { if(T) stream.push("</Table>"); return void (stage = 5); }

			if(R <= range.s.r) {
				if(ws['!cols']) ws['!cols'].forEach(function(n, i) {
					process_col(n);
					var w = !!n.width;
					var p = col_obj_w(i, n);
					var k/*:any*/ = {"ss:Index":i+1};
					if(w) k['ss:Width'] = width2px(p.width);
					if(n.hidden) k['ss:Hidden']="1";
					if(!T) { T = true; stream.push("<Table>"); }
					stream.push(writextag("Column",null,k));
				});
				dense = ws["!data"] != null;
				if(dense) darr = ws["!data"];
				addr.r = addr.c = 0;
			}

			/* process `stride` rows per invocation */
			for(var cnt = 0; R <= range.e.r && cnt < stride; ++R, ++cnt) {
				var row = [write_ws_xlml_row(R, (ws['!rows']||[])[R])];
				addr.r = R;
				if(!(dense && !darr[R])) for(var C = range.s.c; C <= range.e.c; ++C) {
					addr.c = C;
					var skip = false;
					for(mi = 0; mi != marr.length; ++mi) {
						if(marr[mi].s.c > C) continue;
						if(marr[mi].s.r > R) continue;
						if(marr[mi].e.c < C) continue;
						if(marr[mi].e.r < R) continue;
						if(marr[mi].s.c != C || marr[mi].s.r != R) skip = true;
						break;
					}
					if(skip) continue;
					var ref = encode_col(C) + encode_row(R), cell = dense ? darr[R][C] : ws[ref];
					row.push(write_ws_xlml_cell(cell, ref, ws, opts, wsidx, wb, addr));
				}
				row.push("</Row>");
				if(!T) { T = true; stream.push("<Table>"); }
				stream.push(row.join(""));
			}
		} break;

		/* worksheet postamble */
		case 5: {
			stream.push(write_ws_xlml_wsopts(ws, opts, wsidx, wb));
			if(ws && ws["!autofilter"]) stream.push('<AutoFilter x:Range="' + a1_to_rc(fix_range(ws["!autofilter"].ref), {r:0,c:0}) + '" xmlns="urn:schemas-microsoft-com:office:excel"></AutoFilter>');
			stream.push("</Worksheet>");
			wsidx++; R = -1;
			return void (stage = 3);
		}

		/* footer */
		case -1: {
			stage = -2;
			stream.push("</Workbook>");
		} break;

		/* exeunt */
		case -2: stream.push(null); break;
	}};
	return stream;
}

var __stream = {
	to_json: write_json_stream,
	to_html: write_html_stream,
	to_csv: write_csv_stream,
	to_xlml: write_xlml_stream,
	set_readable: set_readable
};
