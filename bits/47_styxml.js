/* 18.8.5 borders CT_Borders */
function parse_borders(t, styles, themes, opts) {
	styles.Borders = [];
	var border = {};
	var side = "";
	var pass = false;
	(t.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<borders': case '<borders>': case '</borders>': break;

			/* 18.8.4 border CT_Border */
			case '<border': case '<border>': case '<border/>':
				border = /*::(*/{}/*:: :any)*/;
				if(y.diagonalUp) border.diagonalUp = parsexmlbool(y.diagonalUp);
				if(y.diagonalDown) border.diagonalDown = parsexmlbool(y.diagonalDown);
				styles.Borders.push(border);
				break;
			case '</border>': break;

			/* note: not in spec, appears to be CT_BorderPr */
			case '<left/>':
				border.left = {};
				if(y.style) border.left.style = y.style;
				break;
			case '<left': case '<left>':
				side = "left"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</left>': side = ""; break;

			/* note: not in spec, appears to be CT_BorderPr */
			case '<right/>':
				border.right = {};
				if(y.style) border.right.style = y.style;
				break;
			case '<right': case '<right>':
				side = "right"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</right>': side = ""; break;

			/* 18.8.43 top CT_BorderPr */
			case '<top/>':
				border.top = {};
				if(y.style) border.top.style = y.style;
				break;
			case '<top': case '<top>':
				side = "top"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</top>': side = ""; break;

			/* 18.8.6 bottom CT_BorderPr */
			case '<bottom/>':
				border.bottom = {};
				if(y.style) border.bottom.style = y.style;
				break;
			case '<bottom': case '<bottom>':
				side = "bottom"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</bottom>': side = ""; break;

			/* 18.8.13 diagonal CT_BorderPr */
			case '<diagonal/>':
				border.diagonal = {};
				if(y.style) border.diagonal.style = y.style;
				break;
			case '<diagonal': case '<diagonal>':
				side = "diagonal"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</diagonal>': side = ""; break;

			/* 18.8.25 horizontal CT_BorderPr */
			case '<horizontal/>':
				border.horizontal = {};
				if(y.style) border.horizontal.style = y.style;
				break;
			case '<horizontal': case '<horizontal>':
				side = "horizontal"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</horizontal>': side = ""; break;

			/* 18.8.44 vertical CT_BorderPr */
			case '<vertical/>':
				border.vertical = {};
				if(y.style) border.vertical.style = y.style;
				break;
			case '<vertical': case '<vertical>':
				side = "vertical"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</vertical>': side = ""; break;

			/* 18.8.37 start CT_BorderPr */
			case '<start/>':
				border.start = {};
				if(y.style) border.start.style = y.style;
				break;
			case '<start': case '<start>':
				side = "start"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</start>': side = ""; break;

			/* 18.8.16 end CT_BorderPr */
			case '<end/>':
				border.end = {};
				if(y.style) border.end.style = y.style;
				break;
			case '<end': case '<end>':
				side = "end"; border[side] = border[side] || {};
				if(y.style) border[side].style = y.style;
				break;
			case '</end>': side = ""; break;

			/* 18.8.? color CT_Color */
			case '<color': case '<color>':
				if(side) border[side].color = style_color_from_attrs(y, themes);
				break;
			case '<color/>': case '</color>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in borders');
			}
		}
	});
}

/* 18.8.21 fills CT_Fills */
function parse_fills(t, styles, themes, opts) {
	styles.Fills = [];
	var fill = {};
	var gradient = null, stop = null;
	var pass = false;
	(t.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<fills': case '<fills>': case '</fills>': break;

			/* 18.8.20 fill CT_Fill */
			case '<fill>': case '<fill': case '<fill/>':
				fill = {}; styles.Fills.push(fill); break;
			case '</fill>': break;

			/* 18.8.24 gradientFill CT_GradientFill */
			case '<gradientFill>':
				gradient = {stops:[]}; fill.gradientFill = gradient; break;
			case '<gradientFill':
				gradient = {stops:[]};
				if(y.type) gradient.type = y.type;
				if(y.degree) gradient.degree = parseFloat(y.degree);
				["left","right","top","bottom"].forEach(function(k) { if(y[k] != null) gradient[k] = parseFloat(y[k]); });
				fill.gradientFill = gradient;
				break;
			case '</gradientFill>': gradient = null; break;

			/* 18.8.32 patternFill CT_PatternFill */
			case '<patternFill': case '<patternFill>':
				if(y.patternType) fill.patternType = y.patternType;
				break;
			case '<patternFill/>': case '</patternFill>': break;

			/* 18.8.3 bgColor CT_Color */
			case '<bgColor':
				fill.bgColor = style_color_from_attrs(y, themes);
				break;
			case '<bgColor/>': case '</bgColor>': break;

			/* 18.8.19 fgColor CT_Color */
			case '<fgColor':
				fill.fgColor = style_color_from_attrs(y, themes);
				break;
			case '<fgColor/>': case '</fgColor>': break;

			/* 18.8.38 stop CT_GradientStop */
			case '<stop':
				stop = {};
				if(y.position != null) stop.position = parseFloat(y.position);
				if(gradient) gradient.stops.push(stop);
				break;
			case '<stop/>':
				if(gradient) {
					stop = {};
					if(y.position != null) stop.position = parseFloat(y.position);
					gradient.stops.push(stop);
				}
				stop = null;
				break;
			case '</stop>': stop = null; break;

			/* 18.8.? color CT_Color */
			case '<color':
				if(stop) stop.color = style_color_from_attrs(y, themes);
				break;
			case '<color/>':
				if(stop) stop.color = style_color_from_attrs(y, themes);
				break;
			case '</color>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in fills');
			}
		}
	});
}

/* 18.8.23 fonts CT_Fonts */
function parse_fonts(t, styles, themes, opts) {
	styles.Fonts = [];
	var font = {};
	var pass = false;
	(t.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<fonts': case '<fonts>': case '</fonts>': break;

			/* 18.8.22 font CT_Font */
			case '<font': case '<font>': break;
			case '</font>': case '<font/>':
				styles.Fonts.push(font);
				font = {};
				break;

			/* 18.8.29 name CT_FontName */
			case '<name': if(y.val) font.name = utf8read(y.val); break;
			case '<name/>': case '</name>': break;

			/* 18.8.2  b CT_BooleanProperty */
			case '<b': font.bold = y.val ? parsexmlbool(y.val) : 1; break;
			case '<b/>': font.bold = 1; break;
			case '</b>': case '</b': break;

			/* 18.8.26 i CT_BooleanProperty */
			case '<i': font.italic = y.val ? parsexmlbool(y.val) : 1; break;
			case '<i/>': font.italic = 1; break;
			case '</i>': case '</i': break;

			/* 18.4.13 u CT_UnderlineProperty */
			case '<u':
				switch(y.val) {
					case "none": font.underline = 0x00; break;
					case "single": font.underline = 0x01; break;
					case "double": font.underline = 0x02; break;
					case "singleAccounting": font.underline = 0x21; break;
					case "doubleAccounting": font.underline = 0x22; break;
				} break;
			case '<u/>': font.underline = 1; break;
			case '</u>': case '</u': break;

			/* 18.4.10 strike CT_BooleanProperty */
			case '<strike': font.strike = y.val ? parsexmlbool(y.val) : 1; break;
			case '<strike/>': font.strike = 1; break;
			case '</strike>': case '</strike': break;

			/* 18.4.2  outline CT_BooleanProperty */
			case '<outline': font.outline = y.val ? parsexmlbool(y.val) : 1; break;
			case '<outline/>': font.outline = 1; break;
			case '</outline>': case '</outline': break;

			/* 18.8.36 shadow CT_BooleanProperty */
			case '<shadow': font.shadow = y.val ? parsexmlbool(y.val) : 1; break;
			case '<shadow/>': font.shadow = 1; break;
			case '</shadow>': case '</shadow': break;

			/* 18.8.12 condense CT_BooleanProperty */
			case '<condense': font.condense = y.val ? parsexmlbool(y.val) : 1; break;
			case '<condense/>': font.condense = 1; break;
			case '</condense>': case '</condense': break;

			/* 18.8.17 extend CT_BooleanProperty */
			case '<extend': font.extend = y.val ? parsexmlbool(y.val) : 1; break;
			case '<extend/>': font.extend = 1; break;
			case '</extend>': case '</extend': break;

			/* 18.4.11 sz CT_FontSize */
			case '<sz': if(y.val) font.sz = +y.val; break;
			case '<sz/>': case '</sz>': case '</sz': break;

			/* 18.4.14 vertAlign CT_VerticalAlignFontProperty */
			case '<vertAlign': if(y.val) font.vertAlign = y.val; break;
			case '<vertAlign/>': case '</vertAlign>': case '</vertAlign': break;

			/* 18.8.18 family CT_FontFamily */
			case '<family': if(y.val) font.family = parseInt(y.val,10); break;
			case '<family/>': case '</family>': case '</family': break;

			/* 18.8.35 scheme CT_FontScheme */
			case '<scheme': if(y.val) font.scheme = y.val; break;
			case '<scheme/>': case '</scheme>': case '</scheme': break;

			/* 18.4.1 charset CT_IntProperty */
			case '<charset':
				if(y.val == '1') break;
				y.codepage = CS2CP[parseInt(y.val, 10)];
				break;
			case '<charset/>': case '</charset>': case '</charset': break;

			/* 18.?.? color CT_Color */
			case '<color':
				font.color = style_color_from_attrs(y, themes);
				break;
			case '<color/>': case '</color>': case '</color': break;

			/* note: sometimes mc:AlternateContent appears bare */
			case '<AlternateContent': pass = true; break;
			case '</AlternateContent>': case '</AlternateContent': pass = false; break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in fonts');
			}
		}
	});
}

/* 18.8.31 numFmts CT_NumFmts */
function parse_numFmts(t, styles, opts) {
	styles.NumberFmt = [];
	var k/*Array<number>*/ = (keys(table_fmt)/*:any*/);
	for(var i=0; i < k.length; ++i) styles.NumberFmt[k[i]] = table_fmt[k[i]];
	var m = t.match(tagregex);
	if(!m) return;
	for(i=0; i < m.length; ++i) {
		var y = parsexmltag(m[i]);
		switch(strip_ns(y[0])) {
			case '<numFmts': case '</numFmts>': case '<numFmts/>': case '<numFmts>': break;
			case '<numFmt': {
				var f=unescapexml(utf8read(y.formatCode)), j=parseInt(y.numFmtId,10);
				styles.NumberFmt[j] = f;
				if(j>0) {
					if(j > 0x188) {
						for(j = 0x188; j > 0x3c; --j) if(styles.NumberFmt[j] == null) break;
						styles.NumberFmt[j] = f;
					}
					SSF__load(f,j);
				}
			} break;
			case '</numFmt>': break;
			default: if(opts.WTF) throw new Error('unrecognized ' + y[0] + ' in numFmts');
		}
	}
}

function write_numFmts(NF/*:{[n:number|string]:string}*//*::, opts*/) {
	var o = ["<numFmts>"];
	[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
		for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) o[o.length] = (writextag('numFmt',null,{numFmtId:i,formatCode:escapexml(NF[i])}));
	});
	if(o.length === 1) return "";
	o[o.length] = ("</numFmts>");
	o[0] = writextag('numFmts', null, { count:o.length-2 }).replace("/>", ">");
	return o.join("");
}

/* 18.8.10 cellXfs CT_CellXfs */
var cellXF_uint = [ "numFmtId", "fillId", "fontId", "borderId", "xfId" ];
var cellXF_bool = [ "applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat", "applyProtection", "pivotButton", "quotePrefix" ];
function parse_xfs(t, styles, opts, key) {
	styles[key] = [];
	var xf;
	var pass = false;
	(t.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x), i = 0;
		switch(strip_ns(y[0])) {
			case '<cellXfs': case '<cellXfs>': case '<cellXfs/>': case '</cellXfs>':
			case '<cellStyleXfs': case '<cellStyleXfs>': case '<cellStyleXfs/>': case '</cellStyleXfs>':
				break;

			/* 18.8.45 xf CT_Xf */
			case '<xf': case '<xf/>': case '<xf>':
				xf = y;
				delete xf[0];
				for(i = 0; i < cellXF_uint.length; ++i) if(xf[cellXF_uint[i]])
					xf[cellXF_uint[i]] = parseInt(xf[cellXF_uint[i]], 10);
				for(i = 0; i < cellXF_bool.length; ++i) if(xf[cellXF_bool[i]])
					xf[cellXF_bool[i]] = parsexmlbool(xf[cellXF_bool[i]]);
				if(styles.NumberFmt && xf.numFmtId > 0x188) {
					for(i = 0x188; i > 0x3c; --i) if(styles.NumberFmt[xf.numFmtId] == styles.NumberFmt[i]) { xf.numFmtId = i; break; }
				}
				styles[key].push(xf); break;
			case '</xf>': break;

			/* 18.8.1 alignment CT_CellAlignment */
			case '<alignment': case '<alignment/>': case '<alignment>':
				var alignment = {};
				if(y.vertical) alignment.vertical = y.vertical;
				if(y.horizontal) alignment.horizontal = y.horizontal;
				if(y.textRotation != null) alignment.textRotation = parseInt(y.textRotation, 10);
				if(y.indent) alignment.indent = parseInt(y.indent, 10);
				if(y.relativeIndent) alignment.relativeIndent = parseInt(y.relativeIndent, 10);
				if(y.readingOrder) alignment.readingOrder = parseInt(y.readingOrder, 10);
				if(y.wrapText) alignment.wrapText = parsexmlbool(y.wrapText);
				if(y.shrinkToFit) alignment.shrinkToFit = parsexmlbool(y.shrinkToFit);
				if(y.justifyLastLine) alignment.justifyLastLine = parsexmlbool(y.justifyLastLine);
				xf.alignment = alignment;
				break;
			case '</alignment>': break;

			/* 18.8.33 protection CT_CellProtection */
			case '<protection': case '<protection>':
				var protection = {};
				if(y.locked != null) protection.locked = parsexmlbool(y.locked);
				if(y.hidden != null) protection.hidden = parsexmlbool(y.hidden);
				xf.protection = protection;
				break;
			case '<protection/>':
				protection = {};
				if(y.locked != null) protection.locked = parsexmlbool(y.locked);
				if(y.hidden != null) protection.hidden = parsexmlbool(y.hidden);
				xf.protection = protection;
				break;
			case '</protection>': break;

			/* note: sometimes mc:AlternateContent appears bare */
			case '<AlternateContent': case '<AlternateContent>': pass = true; break;
			case '</AlternateContent>': pass = false; break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in cellXfs');
			}
		}
	});
}
function parse_cellXfs(t, styles, opts) { parse_xfs(t, styles, opts, "CellXf"); }
function parse_cellStyleXfs(t, styles, opts) { parse_xfs(t, styles, opts, "CellStyleXf"); }

function parse_cellStyles(t, styles, opts) {
	styles.CellStyles = [];
	(t.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<cellStyles': case '<cellStyles>': case '<cellStyles/>': case '</cellStyles>': break;
			case '<cellStyle': case '<cellStyle/>':
				delete y[0];
				if(y.xfId != null) y.xfId = parseInt(y.xfId, 10);
				if(y.builtinId != null) y.builtinId = parseInt(y.builtinId, 10);
				if(y.iLevel != null) y.iLevel = parseInt(y.iLevel, 10);
				if(y.customBuiltin != null) y.customBuiltin = parsexmlbool(y.customBuiltin);
				if(y.hidden != null) y.hidden = parsexmlbool(y.hidden);
				if(y.name) y.name = utf8read(y.name);
				styles.CellStyles.push(y);
				break;
			case '</cellStyle>': break;
			default: if(opts && opts.WTF) throw new Error('unrecognized ' + y[0] + ' in cellStyles');
		}
	});
}

function parse_dxfs(t, styles, themes, opts) {
	styles.Dxfs = [];
	var dxf = null, font = null, fill = null, border = null, side = "";
	var pass = false;
	(t.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<dxfs': case '<dxfs>': case '<dxfs/>': case '</dxfs>': break;
			case '<dxf': case '<dxf>': dxf = {}; styles.Dxfs.push(dxf); break;
			case '<dxf/>': styles.Dxfs.push({}); dxf = null; break;
			case '</dxf>': dxf = null; break;
				case '<font': case '<font>': font = {}; if(dxf) dxf.font = font; break;
				case '</font>': case '<font/>': font = null; break;
				case '<name': if(font && y.val) font.name = utf8read(y.val); break;
				case '<name/>': case '</name>': break;
				case '<b': if(font) font.bold = y.val ? parsexmlbool(y.val) : 1; break;
				case '<b/>': if(font) font.bold = 1; break;
				case '</b>': break;
				case '<i': if(font) font.italic = y.val ? parsexmlbool(y.val) : 1; break;
				case '<i/>': if(font) font.italic = 1; break;
				case '</i>': break;
				case '<u': if(font) font.underline = y.val || 1; break;
				case '<u/>': if(font) font.underline = 1; break;
				case '</u>': break;
				case '<strike': if(font) font.strike = y.val ? parsexmlbool(y.val) : 1; break;
				case '<strike/>': if(font) font.strike = 1; break;
				case '</strike>': break;
				case '<sz': if(font && y.val) font.sz = +y.val; break;
				case '<sz/>': case '</sz>': break;
				case '<outline': if(font) font.outline = y.val ? parsexmlbool(y.val) : 1; break;
				case '<outline/>': if(font) font.outline = 1; break;
				case '</outline>': break;
				case '<shadow': if(font) font.shadow = y.val ? parsexmlbool(y.val) : 1; break;
				case '<shadow/>': if(font) font.shadow = 1; break;
				case '</shadow>': break;
				case '<condense': if(font) font.condense = y.val ? parsexmlbool(y.val) : 1; break;
				case '<condense/>': if(font) font.condense = 1; break;
				case '</condense>': break;
				case '<extend': if(font) font.extend = y.val ? parsexmlbool(y.val) : 1; break;
				case '<extend/>': if(font) font.extend = 1; break;
				case '</extend>': break;
				case '<vertAlign': if(font && y.val) font.vertAlign = y.val; break;
				case '<vertAlign/>': case '</vertAlign>': break;
				case '<family': if(font && y.val) font.family = parseInt(y.val, 10); break;
				case '<family/>': case '</family>': break;
				case '<scheme': if(font && y.val) font.scheme = y.val; break;
				case '<scheme/>': case '</scheme>': break;
				case '<charset': if(font && y.val != null) font.charset = parseInt(y.val, 10); break;
				case '<charset/>': case '</charset>': break;
				case '<color': if(font) font.color = style_color_from_attrs(y, themes); else if(side && border) border[side].color = style_color_from_attrs(y, themes); break;
				case '<color/>': case '</color>': break;
				case '<fill': case '<fill>': fill = {}; if(dxf) dxf.fill = fill; break;
				case '<fill/>': fill = null; break;
				case '</fill>': fill = null; break;
				case '<patternFill': case '<patternFill>': if(fill && y.patternType) fill.patternType = y.patternType; break;
				case '<patternFill/>': case '</patternFill>': break;
				case '<fgColor': if(fill) fill.fgColor = style_color_from_attrs(y, themes); break;
				case '<fgColor/>': case '</fgColor>': break;
				case '<bgColor': if(fill) fill.bgColor = style_color_from_attrs(y, themes); break;
				case '<bgColor/>': case '</bgColor>': break;
				case '<border': case '<border>': border = {}; if(dxf) dxf.border = border; break;
			case '<border/>': border = null; break;
			case '</border>': border = null; break;
			case '<left': case '<left>': case '<left/>': side = "left"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
			case '</left>': side = ""; break;
			case '<right': case '<right>': case '<right/>': side = "right"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
			case '</right>': side = ""; break;
				case '<top': case '<top>': case '<top/>': side = "top"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
				case '</top>': side = ""; break;
				case '<bottom': case '<bottom>': case '<bottom/>': side = "bottom"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
				case '</bottom>': side = ""; break;
				case '<diagonal': case '<diagonal>': case '<diagonal/>': side = "diagonal"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
				case '</diagonal>': side = ""; break;
				case '<horizontal': case '<horizontal>': case '<horizontal/>': side = "horizontal"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
				case '</horizontal>': side = ""; break;
				case '<vertical': case '<vertical>': case '<vertical/>': side = "vertical"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
				case '</vertical>': side = ""; break;
				case '<start': case '<start>': case '<start/>': side = "start"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
				case '</start>': side = ""; break;
				case '<end': case '<end>': case '<end/>': side = "end"; if(border) { border[side] = border[side] || {}; if(y.style) border[side].style = y.style; } if(y[0].slice(-2) == "/>") side = ""; break;
				case '</end>': side = ""; break;
				case '<alignment': if(dxf) {
					dxf.alignment = {};
					if(y.vertical) dxf.alignment.vertical = y.vertical;
					if(y.horizontal) dxf.alignment.horizontal = y.horizontal;
					if(y.wrapText) dxf.alignment.wrapText = parsexmlbool(y.wrapText);
				} break;
				case '<alignment/>': case '</alignment>': break;
				case '<numFmt': if(dxf) {
					dxf.numFmt = {numFmtId:y.numFmtId != null ? parseInt(y.numFmtId, 10) : void 0, formatCode:y.formatCode ? unescapexml(utf8read(y.formatCode)) : void 0};
				} break;
				case '<numFmt/>': case '</numFmt>': break;
				case '<protection': if(dxf) {
					dxf.protection = {};
					if(y.locked != null) dxf.protection.locked = parsexmlbool(y.locked);
					if(y.hidden != null) dxf.protection.hidden = parsexmlbool(y.hidden);
				} break;
				case '<protection/>': case '</protection>': break;
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in dxfs');
			}
		}
	});
}

function parse_colors(t, styles, themes, opts) {
	styles.Colors = {indexedColors:[], mruColors:[], themeColors:[]};
	var target = null, pass = false;
	(t.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<colors': case '<colors>': case '</colors>': break;
			case '<indexedColors': case '<indexedColors>': target = styles.Colors.indexedColors; break;
			case '</indexedColors>': target = null; break;
			case '<themeColors': case '<themeColors>': target = styles.Colors.themeColors; break;
			case '</themeColors>': target = null; break;
			case '<mruColors': case '<mruColors>': target = styles.Colors.mruColors; break;
			case '</mruColors>': target = null; break;
			case '<rgbColor': case '<rgbColor/>':
				if(target) target.push(style_color_from_attrs(y, themes));
				break;
			case '</rgbColor>': break;
			case '<color': case '<color/>':
				if(target) target.push(style_color_from_attrs(y, themes));
				break;
			case '</color>': break;
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in colors');
			}
		}
	});
}

function write_cellXfs(cellXfs)/*:string*/ {
	var o/*:Array<string>*/ = [];
	o[o.length] = (writextag('cellXfs',null));
	cellXfs.forEach(function(c) {
		o[o.length] = (writextag('xf', null, c));
	});
	o[o.length] = ("</cellXfs>");
	if(o.length === 2) return "";
	o[0] = writextag('cellXfs',null, {count:o.length-2}).replace("/>",">");
	return o.join("");
}

/* 18.8 Styles CT_Stylesheet*/
var parse_sty_xml= /*#__PURE__*/(function make_pstyx() {

return function parse_sty_xml(data, themes, opts) {
	var styles = {};
	if(!data) return styles;
	data = remove_doctype(str_remove_ng(data, "<!--", "-->"));
	/* 18.8.39 styleSheet CT_Stylesheet */
	var t;

	/* 18.8.31 numFmts CT_NumFmts ? */
	if((t=str_match_xml_ns(data, "numFmts"))) parse_numFmts(t[0], styles, opts);

	/* 18.8.23 fonts CT_Fonts ? */
	if((t=str_match_xml_ns(data, "fonts"))) parse_fonts(t[0], styles, themes, opts);

	/* 18.8.21 fills CT_Fills ? */
	if((t=str_match_xml_ns(data, "fills"))) parse_fills(t[0], styles, themes, opts);

	/* 18.8.5  borders CT_Borders ? */
	if((t=str_match_xml_ns(data, "borders"))) parse_borders(t[0], styles, themes, opts);

	/* 18.8.9  cellStyleXfs CT_CellStyleXfs ? */
	if((t=str_match_xml_ns(data, "cellStyleXfs"))) parse_cellStyleXfs(t[0], styles, opts);

	/* 18.8.8  cellStyles CT_CellStyles ? */
	if((t=str_match_xml_ns(data, "cellStyles"))) parse_cellStyles(t[0], styles, opts);

	/* 18.8.10 cellXfs CT_CellXfs ? */
	if((t=str_match_xml_ns(data, "cellXfs"))) parse_cellXfs(t[0], styles, opts);

	/* 18.8.15 dxfs CT_Dxfs ? */
	if((t=str_match_xml_ns(data, "dxfs"))) parse_dxfs(t[0], styles, themes, opts);

	/* 18.8.42 tableStyles CT_TableStyles ? */
	/* 18.8.11 colors CT_Colors ? */
	if((t=str_match_xml_ns(data, "colors"))) parse_colors(t[0], styles, themes, opts);

	/* 18.2.10 extLst CT_ExtensionList ? */

	return styles;
};
})();

function write_sty_xml(wb/*:Workbook*/, opts)/*:string*/ {
	var o = [XML_HEADER, writextag('styleSheet', null, {
		'xmlns': XMLNS_main[0],
		'xmlns:vt': XMLNS.vt
	})], w;
	if(wb.SSF && (w = write_numFmts(wb.SSF)) != null) o[o.length] = w;
	o[o.length] = ('<fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>');
	o[o.length] = ('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
	o[o.length] = ('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');
	o[o.length] = ('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
	if((w = write_cellXfs(opts.cellXfs))) o[o.length] = (w);
	o[o.length] = ('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
	o[o.length] = ('<dxfs count="0"/>');
	o[o.length] = ('<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>');

	if(o.length>2){ o[o.length] = ('</styleSheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
