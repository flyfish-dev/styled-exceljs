var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options


/*global Map */
var browser_has_Map = typeof Map !== 'undefined';

function get_sst_id(sst/*:SST*/, str/*:string*/, rev)/*:number*/ {
	var i = 0, len = sst.length;
	if(rev) {
		if(browser_has_Map ? rev.has(str) : Object.prototype.hasOwnProperty.call(rev, str)) {
			var revarr = browser_has_Map ? rev.get(str) : rev[str];
			for(; i < revarr.length; ++i) {
				if(sst[revarr[i]].t === str) { sst.Count ++; return revarr[i]; }
			}
		}
	} else for(; i < len; ++i) {
		if(sst[i].t === str) { sst.Count ++; return i; }
	}
	sst[len] = ({t:str}/*:any*/); sst.Count ++; sst.Unique ++;
	if(rev) {
		if(browser_has_Map) {
			if(!rev.has(str)) rev.set(str, []);
			rev.get(str).push(len);
		} else {
			if(!Object.prototype.hasOwnProperty.call(rev, str)) rev[str] = [];
			rev[str].push(len);
		}
	}
	return len;
}

function col_obj_w(C/*:number*/, col) {
	var p = ({min:C+1,max:C+1}/*:any*/);
	/* wch (chars), wpx (pixels) */
	var wch = -1;
	if(col.MDW) MDW = col.MDW;
	if(col.width != null) p.customWidth = 1;
	else if(col.wpx != null) wch = px2char(col.wpx);
	else if(col.wch != null) wch = col.wch;
	if(wch > -1) { p.width = char2width(wch); p.customWidth = 1; }
	else if(col.width != null) p.width = col.width;
	if(p.width != null) p.width = clamp_col_width(p.width);
	if(col.hidden) p.hidden = true;
	if(col.level != null) { p.outlineLevel = p.level = col.level; }
	if(col.bestFit || col.bestfit) p.bestFit = true;
	if(col.customWidth || col.customwidth) p.customWidth = 1;
	return p;
}

function default_margins(margins/*:Margins*/, mode/*:?string*/) {
	if(!margins) return;
	var defs = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3];
	if(mode == 'xlml') defs = [1, 1, 1, 1, 0.5, 0.5];
	if(margins.left   == null) margins.left   = defs[0];
	if(margins.right  == null) margins.right  = defs[1];
	if(margins.top    == null) margins.top    = defs[2];
	if(margins.bottom == null) margins.bottom = defs[3];
	if(margins.header == null) margins.header = defs[4];
	if(margins.footer == null) margins.footer = defs[5];
}

function get_cell_style(styles/*:Array<any>*/, cell/*:Cell*/, opts) {
	var z = opts.revssf[cell.z != null ? cell.z : "General"];
	var i = 0x3c, len = styles.length;
	if(z == null && opts.ssf) {
		for(; i < 0x188; ++i) if(opts.ssf[i] == null) {
			SSF__load(cell.z, i);
			// $FlowIgnore
			opts.ssf[i] = cell.z;
			opts.revssf[cell.z] = z = i;
			break;
		}
	}
	for(i = 0; i != len; ++i) if(styles[i].numFmtId === z) return i;
	styles[len] = {
		numFmtId:z,
		fontId:0,
		fillId:0,
		borderId:0,
		xfId:0,
		applyNumberFormat:1
	};
	return len;
}

function extend_style_obj(dst, src) {
	if(!src) return dst;
	keys(src).forEach(function(k) { dst[k] = dup(src[k]); });
	return dst;
}

function fill_style_aliases(out, fill) {
	if(!fill) return out;
	if(fill.patternType != null) out.patternType = fill.patternType;
	if(fill.fgColor != null) out.fgColor = dup(fill.fgColor);
	if(fill.bgColor != null) out.bgColor = dup(fill.bgColor);
	if(fill.gradientFill != null) out.gradientFill = dup(fill.gradientFill);
	return out;
}

function resolve_cell_style(styles, cf, themes, styleid) {
	if(!styles || !cf) return null;
	var xf = ({}/*:any*/);
	if(cf.xfId != null && styles.CellStyleXf && styles.CellStyleXf[cf.xfId]) extend_style_obj(xf, styles.CellStyleXf[cf.xfId]);
	extend_style_obj(xf, cf);
	var out = ({id: styleid, xf: dup(xf)}/*:any*/);
	if(xf.numFmtId != null) {
		out.numFmtId = xf.numFmtId;
		if(styles.NumberFmt && styles.NumberFmt[xf.numFmtId] != null) out.numFmt = styles.NumberFmt[xf.numFmtId];
	}
	if(xf.fontId != null && styles.Fonts && styles.Fonts[xf.fontId]) out.font = resolve_style_obj_color(styles.Fonts[xf.fontId], themes);
	if(xf.fillId != null && styles.Fills && styles.Fills[xf.fillId]) {
		out.fill = resolve_style_obj_color(styles.Fills[xf.fillId], themes);
		fill_style_aliases(out, out.fill);
	}
	if(xf.borderId != null && styles.Borders && styles.Borders[xf.borderId]) out.border = resolve_style_obj_color(styles.Borders[xf.borderId], themes);
	if(xf.alignment) out.alignment = dup(xf.alignment);
	if(xf.protection) out.protection = dup(xf.protection);
	return out;
}

function safe_format(p/*:Cell*/, fmtid/*:number*/, fillid/*:?number*/, opts, themes, styles, date1904, cf/*:?any*/, styleid/*:?number*/) {
	try {
		if(opts.cellNF) p.z = table_fmt[fmtid];
	} catch(e) { if(opts.WTF) throw e; }
	if(p.t === 'z' && !opts.cellStyles) return;
	if(p.t === 'd' && typeof p.v === 'string') p.v = parseDate(p.v);
	if((!opts || opts.cellText !== false) && p.t !== 'z') try {
		if(table_fmt[fmtid] == null) SSF__load(SSFImplicit[fmtid] || "General", fmtid);
		if(p.t === 'e') p.w = p.w || BErr[p.v];
		else if(fmtid === 0) {
			if(p.t === 'n') {
				if((p.v|0) === p.v) p.w = p.v.toString(10);
				else p.w = SSF_general_num(p.v);
			}
			else if(p.t === 'd') {
				var dd = datenum(p.v, !!date1904);
				if((dd|0) === dd) p.w = dd.toString(10);
				else p.w = SSF_general_num(dd);
			}
			else if(p.v === undefined) return "";
			else p.w = SSF_general(p.v,_ssfopts);
		}
		else if(p.t === 'd') p.w = SSF_format(fmtid,datenum(p.v, !!date1904),_ssfopts);
		else p.w = SSF_format(fmtid,p.v,_ssfopts);
	} catch(e) { if(opts.WTF) throw e; }
	if(!opts.cellStyles) return;
	if(cf != null) try {
		var resolved = resolve_cell_style(styles, cf, themes, styleid);
		if(resolved) { p.s = resolved; return; }
	} catch(e) { if(opts.WTF) throw e; }
	if(fillid != null) try {
		p.s = resolve_style_obj_color(styles.Fills[fillid], themes);
	} catch(e) { if(opts.WTF && styles.Fills) throw e; }
}

function merge_range_overlap(a, b) {
	return !(a.e.r < b.s.r || b.e.r < a.s.r || a.e.c < b.s.c || b.e.c < a.s.c);
}

function validate_merges(ws/*:Worksheet*/, opts/*:?any*/) {
	var merges = (ws && ws["!merges"]) || [];
	var errors = [];
	var ref = ws && ws["!ref"] ? safe_decode_range(ws["!ref"]) : null;
	var seen = {};
	for(var i = 0; i < merges.length; ++i) {
		var m = merges[i];
		var enc = "";
		if(!m || !m.s || !m.e) {
			errors.push({code:"E_MERGE_RANGE", message:"Merge range is malformed", index:i});
			continue;
		}
		if(m.s.r < 0 || m.s.c < 0 || m.e.r < m.s.r || m.e.c < m.s.c) {
			errors.push({code:"E_MERGE_RANGE", message:"Merge range is invalid", index:i, range:m});
			continue;
		}
		enc = encode_range(m);
		if(seen[enc] != null) errors.push({code:"E_MERGE_DUP", message:"Merge range is duplicated", index:i, other:seen[enc], range:enc});
		seen[enc] = i;
		if(ref && (m.s.r < ref.s.r || m.s.c < ref.s.c || m.e.r > ref.e.r || m.e.c > ref.e.c))
			errors.push({code:"E_MERGE_BOUNDS", message:"Merge range exceeds worksheet range", index:i, range:enc, ref:encode_range(ref)});
		for(var j = 0; j < i; ++j) {
			if(!merges[j] || !merges[j].s || !merges[j].e) continue;
			if(merge_range_overlap(m, merges[j]) && encode_range(merges[j]) != enc)
				errors.push({code:"E_MERGE_OVERLAP", message:"Merge ranges overlap", index:i, other:j, range:enc, otherRange:encode_range(merges[j])});
		}
	}
	if(errors.length && opts && opts.WTF) throw new Error(errors[0].message + " (" + (errors[0].range || errors[0].index) + ")");
	return errors;
}

function check_ws(ws/*:Worksheet*/, sname/*:string*/, i/*:number*/) {
	if(ws && ws['!ref']) {
		var range = safe_decode_range(ws['!ref']);
		if(range.e.c < range.s.c || range.e.r < range.s.r) throw new Error("Bad range (" + i + "): " + ws['!ref']);
	}
}
