/* 20.5 DrawingML - SpreadsheetML Drawing */
/* 20.5.2.35 wsDr CT_Drawing */
function parse_drawing(data, rels/*:any*/) {
	if(!data) return ({charts:[], images:[], shapes:[], raw:""}/*:any*/);
	/*
	  Chartsheet Drawing:
	   - 20.5.2.35 wsDr CT_Drawing
	    - 20.5.2.1  absoluteAnchor CT_AbsoluteAnchor
	     - 20.5.2.16 graphicFrame CT_GraphicalObjectFrame
	      - 20.1.2.2.16 graphic CT_GraphicalObject
	       - 20.1.2.2.17 graphicData CT_GraphicalObjectData
	       - chart reference
	   the actual type is based on the URI of the graphicData
		TODO: handle embedded charts and other types of graphics
	*/
	var out = ({charts:[], images:[], shapes:[], raw:data}/*:any*/);
	if(!rels) rels = {'!id':{}};
	var anchorRegex = /<(?:\w+:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/(?:\w+:)?\1>/g;
	var anchors = data.match(anchorRegex) || [data];
	anchors.forEach(function(anchorXml) {
		var anchor = parse_drawing_anchor(anchorXml);
		(anchorXml.match(/<c:chart\b[^<>]*r:id="([^<>"]*)"/g)||[]).forEach(function(chartTag) {
			var idm = chartTag.match(/r:id="([^<>"]*)"/);
			if(!idm) return;
			var rel = rels['!id'][idm[1]] || {};
			out.charts.push({id:idm[1], rel:rel, target:rel.Target, anchor:anchor});
		});
		(anchorXml.match(/<a:blip\b[^<>]*(?:r:embed|r:link)="([^<>"]*)"/g)||[]).forEach(function(blipTag) {
			var idm = blipTag.match(/(?:r:embed|r:link)="([^<>"]*)"/);
			if(!idm) return;
			var rel = rels['!id'][idm[1]] || {};
			out.images.push({id:idm[1], rel:rel, target:rel.Target, anchor:anchor});
		});
		var tx = parse_drawing_text(anchorXml);
		if(tx) out.shapes.push({text:tx, anchor:anchor, raw:anchorXml});
	});
	out.chart = out.charts[0] && out.charts[0].target;
	return out;
}

function parse_drawing_marker(data, tag) {
	var m = str_match_xml_ns(data, tag);
	var body = m && m[1] || "";
	function num(name) {
		var v = str_match_xml_ns(body, name);
		return v && v[1] != null ? parseInt(v[1], 10) : 0;
	}
	return {col:num("col"), colOff:num("colOff"), row:num("row"), rowOff:num("rowOff")};
}

function parse_drawing_anchor(data) {
	var type = (data.match(/^<(?:\w+:)?(\w+)/)||[])[1] || "anchor";
	var anchor = ({type:type}/*:any*/);
	if(data.indexOf("<xdr:from") >= 0 || data.indexOf("<from") >= 0) anchor.from = parse_drawing_marker(data, "from");
	if(data.indexOf("<xdr:to") >= 0 || data.indexOf("<to") >= 0) anchor.to = parse_drawing_marker(data, "to");
	var pos = data.match(/<(?:\w+:)?pos\b[^<>]*\/>/);
	if(pos) {
		var p = parsexmltag(pos[0]);
		anchor.pos = {x:p.x ? parseInt(p.x, 10) : 0, y:p.y ? parseInt(p.y, 10) : 0};
	}
	var ext = data.match(/<(?:\w+:)?ext\b[^<>]*\/>/);
	if(ext) {
		var e = parsexmltag(ext[0]);
		anchor.ext = {cx:e.cx ? parseInt(e.cx, 10) : 0, cy:e.cy ? parseInt(e.cy, 10) : 0};
	}
	return anchor;
}

function parse_drawing_text(data) {
	var out = [];
	(data.match(/<a:t\b[^>]*>[\s\S]*?<\/a:t>/g)||[]).forEach(function(t) {
		out.push(unescapexml(t.replace(/<[^>]*>/g, "")));
	});
	return out.join("");
}

/* [MS-ODRAW] OfficeArt containers used in BIFF8 MsoDrawing records */
function parse_MsoDrawingGroup(blob, length, opts) {
	var data = blob.slice(blob.l, blob.l + length);
	blob.l += length;
	if(!opts || (!opts.drawings && !opts.charts)) return {raw:data, blips:[], shapes:[], images:[], charts:[], groups:true};
	return parse_xls_officeart(data, true);
}

function parse_MsoDrawing(blob, length, opts) {
	var data = blob.slice(blob.l, blob.l + length);
	blob.l += length;
	if(!opts || (!opts.drawings && !opts.charts)) return {raw:data, blips:[], shapes:[], images:[], charts:[], groups:false};
	return parse_xls_officeart(data, false);
}

function xls_officeart_u16(data, off) { return data[off] | (data[off+1] << 8); }
function xls_officeart_u32(data, off) { return (data[off] | (data[off+1] << 8) | (data[off+2] << 16) | (data[off+3] << 24)) >>> 0; }

function parse_xls_officeart(data, isGroup) {
	var out = ({raw:data, blips:[], shapes:[], images:[], charts:[], groups:!!isGroup}/*:any*/);
	var state = {out:out, current:null};
	try { walk_xls_officeart(data, 0, data.length, state); } catch(e) { out.error = e.message || String(e); }
	return out;
}

function walk_xls_officeart(data, start, end, state) {
	var pos = start;
	while(pos + 8 <= end && pos + 8 <= data.length) {
		var verinst = xls_officeart_u16(data, pos), fbt = xls_officeart_u16(data, pos + 2), cb = xls_officeart_u32(data, pos + 4);
		var ver = verinst & 0x000F, inst = verinst >> 4, body = pos + 8, next = body + cb;
		if(next > data.length) next = data.length;
		if(fbt == 0xF004 /* OfficeArtSpContainer */) {
			var shape = ({raw:data.slice(pos, next), props:{}}/*:any*/);
			state.out.shapes.push(shape);
			var old = state.current; state.current = shape;
			walk_xls_officeart(data, body, next, state);
			state.current = old;
		} else if(fbt == 0xF007 /* OfficeArtBStoreContainerFileBlock */) {
			state.out.blips.push(parse_xls_officeart_bse(data, body, next, inst));
		} else if(fbt == 0xF00A /* OfficeArtFSP */ && state.current) {
			state.current.spid = xls_officeart_u32(data, body);
			state.current.flags = xls_officeart_u32(data, body + 4);
		} else if(fbt == 0xF00B /* OfficeArtFOPT */ && state.current) {
			state.current.props = parse_xls_officeart_props(data, body, next, inst);
			if(state.current.props.pib != null) state.current.blipId = state.current.props.pib;
		} else if(fbt == 0xF010 /* OfficeArtClientAnchor */ && state.current) {
			state.current.anchor = parse_xls_officeart_anchor(data, body, next);
		} else if(fbt == 0xF011 /* OfficeArtClientData */ && state.current) {
			state.current.clientData = data.slice(body, next);
		} else if(ver == 0x0F) walk_xls_officeart(data, body, next, state);
		pos = next;
	}
}

function parse_xls_officeart_props(data, start, end, count) {
	var props = {}, complex = [], pos = start, i = 0;
	for(; i < count && pos + 6 <= end; ++i, pos += 6) {
		var opid = xls_officeart_u16(data, pos), op = xls_officeart_u32(data, pos + 2), id = opid & 0x3FFF;
		props[id] = op;
		if(opid & 0x4000) complex.push([id, op]);
		if(id == 260) props.pib = op;
		if(id == 261) props.pibName = op;
		if(id == 896) props.fillColor = op;
		if(id == 897) props.fillOpacity = op;
		if(id == 959) props.lineColor = op;
	}
	complex.forEach(function(c) {
		if(pos + c[1] <= end) props["complex_" + c[0]] = data.slice(pos, pos + c[1]);
		pos += c[1];
	});
	return props;
}

function parse_xls_officeart_anchor(data, start, end) {
	if(end - start < 18) return {};
	var pos = start;
	var flags = xls_officeart_u16(data, pos); pos += 2;
	var c1 = xls_officeart_u16(data, pos), dx1 = xls_officeart_u16(data, pos + 2), r1 = xls_officeart_u16(data, pos + 4), dy1 = xls_officeart_u16(data, pos + 6); pos += 8;
	var c2 = xls_officeart_u16(data, pos), dx2 = xls_officeart_u16(data, pos + 2), r2 = xls_officeart_u16(data, pos + 4), dy2 = xls_officeart_u16(data, pos + 6);
	return {type:"twoCellAnchor", flags:flags, from:{col:c1, colOff:dx1 * 9525 / 1024, row:r1, rowOff:dy1 * 9525 / 256}, to:{col:c2, colOff:dx2 * 9525 / 1024, row:r2, rowOff:dy2 * 9525 / 256}};
}

function xls_find_image_magic(data, start, end) {
	for(var i = start; i + 8 <= end; ++i) {
		if(data[i] == 0x89 && data[i+1] == 0x50 && data[i+2] == 0x4E && data[i+3] == 0x47) return [i, "image/png"];
		if(data[i] == 0xFF && data[i+1] == 0xD8 && data[i+2] == 0xFF) return [i, "image/jpeg"];
		if(data[i] == 0x42 && data[i+1] == 0x4D) return [i, "image/bmp"];
	}
	return [start, ""];
}

function parse_xls_officeart_bse(data, start, end, inst) {
	var out = ({index:inst, raw:data.slice(start, end)}/*:any*/);
	if(end - start < 36) return out;
	out.btWin32 = data[start]; out.btMacOS = data[start + 1];
	out.size = xls_officeart_u32(data, start + 20);
	out.cRef = xls_officeart_u32(data, start + 24);
	var blip = start + 36;
	if(blip + 8 <= end) {
		out.blipType = xls_officeart_u16(data, blip + 2);
		var bstart = blip + 8, bend = end;
		var magic = xls_find_image_magic(data, bstart, bend);
		if(magic[1]) {
			out.contentType = magic[1];
			out.data = data.slice(magic[0], bend);
			out.dataURI = "data:" + out.contentType + ";base64," + Base64_encode_arr(out.data);
		}
	}
	return out;
}
