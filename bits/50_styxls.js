/* [MS-XLS] 2.4.326 TODO: payload is a zip file */
function parse_Theme(blob, length, opts) {
	var end = blob.l + length;
	var dwThemeVersion = blob.read_shift(4);
	if(dwThemeVersion === 124226) return;
	if(!opts.cellStyles) { blob.l = end; return; }
	var data = blob.slice(blob.l);
	blob.l = end;
	var zip; try { zip = zip_read(data, {type: "array"}); } catch(e) { return; }
	var themeXML = getzipstr(zip, "theme/theme/theme1.xml", true);
	if(!themeXML) return;
	return parse_theme_xml(themeXML, opts);
}

/* 2.5.49 */
function parse_ColorTheme(blob/*::, length*/) { return blob.read_shift(4); }

/* 2.5.155 */
function parse_FullColorExt(blob/*::, length*/) {
	var o = {};
	o.xclrType = blob.read_shift(2);
	o.nTintShade = blob.read_shift(2, 'i');
	switch(o.xclrType) {
		case 0: blob.l += 4; break;
		case 1: o.xclrValue = parse_IcvXF(blob, 4); break;
		case 2: o.xclrValue = parse_LongRGBA(blob, 4); break;
		case 3: o.xclrValue = parse_ColorTheme(blob, 4); break;
		case 4: blob.l += 4; break;
	}
	blob.l += 8;
	return o;
}

/* 2.5.164 */
function parse_IcvXF(blob, length) {
	var end = blob.l + length;
	var icv = blob.read_shift(2) & 0x7F;
	blob.l = end;
	return icv;
}

/* 2.5.280 */
function parse_XFExtGradient(blob, length) {
	return parsenoop(blob, length);
}

/* [MS-XLS] 2.5.108 */
function parse_ExtProp(blob/*::, length*/)/*:Array<any>*/ {
	var extType = blob.read_shift(2);
	var cb = blob.read_shift(2) - 4;
	var o = [extType];
	switch(extType) {
		case 0x04: case 0x05: case 0x07: case 0x08:
		case 0x09: case 0x0A: case 0x0B: case 0x0D:
			o[1] = parse_FullColorExt(blob, cb); break;
		case 0x06: o[1] = parse_XFExtGradient(blob, cb); break;
		case 0x0E: case 0x0F: o[1] = blob.read_shift(cb === 1 ? 1 : 2); break;
		default: throw new Error("Unrecognized ExtProp type: " + extType + " " + cb);
	}
	return o;
}

/* 2.4.355 */
function parse_XFExt(blob, length) {
	var end = blob.l + length;
	blob.l += 2;
	var ixfe = blob.read_shift(2);
	blob.l += 2;
	var cexts = blob.read_shift(2);
	var ext/*:AOA*/ = [];
	while(cexts-- > 0) ext.push(parse_ExtProp(blob, end-blob.l));
	return {ixfe:ixfe, ext:ext};
}

/* xf is an XF, see parse_XFExt for xfext */
function update_xfext(xf, xfext) {
	if(!xf) return;
	if(!xf.data) xf.data = {};
	xfext.forEach(function(xfe) {
		switch(xfe[0]) { /* 2.5.108 extPropData */
			case 0x04: xf.data.xfextFore = xfe[1]; break; /* foreground color */
			case 0x05: xf.data.xfextBack = xfe[1]; break; /* background color */
			case 0x06: xf.data.gradientFill = xfe[1]; break; /* gradient fill */
			case 0x07: xf.data.xfextTop = xfe[1]; break; /* top cell border color */
			case 0x08: xf.data.xfextBottom = xfe[1]; break; /* bottom cell border color */
			case 0x09: xf.data.xfextLeft = xfe[1]; break; /* left cell border color */
			case 0x0a: xf.data.xfextRight = xfe[1]; break; /* right cell border color */
			case 0x0b: xf.data.xfextDiag = xfe[1]; break; /* diagonal cell border color */
			case 0x0d: xf.xfextFont = xfe[1]; break; /* text color */
			case 0x0e: xf.fontScheme = xfe[1]; break; /* font scheme */
			case 0x0f: xf.data.cIndent = xfe[1]; break; /* indentation level */
		}
	});
}
