var TABLE_STYLE_MAX_TABLES = 1024;
var TABLE_STYLE_MAX_XML_LENGTH = 5 * 1024 * 1024;

function table_style_attr(y, name) {
	if(y[name] != null) return y[name];
	return y[name.toLowerCase()];
}

function table_style_bool(y, name, fallback) {
	var value = table_style_attr(y, name);
	return value == null ? fallback : parsexmlbool(value);
}

function table_style_int(y, name) {
	var value = table_style_attr(y, name);
	if(value == null || value === "") return void 0;
	value = parseInt(value, 10);
	return isFinite(value) && value >= 0 ? value : void 0;
}

function table_style_merge(target, source) {
	if(!source) return target;
	if(!target) target = {};
	keys(source).forEach(function(k) {
		var value = source[k];
		if(value && typeof value == "object" && !Array.isArray(value)) {
			target[k] = table_style_merge(
				target[k] && typeof target[k] == "object" && !Array.isArray(target[k]) ? dup(target[k]) : {},
				value
			);
		} else target[k] = value;
	});
	return target;
}

function table_style_color(themes, theme, tint) {
	return style_color_from_attrs({theme:String(theme), tint:String(tint || 0)}, themes);
}

function table_style_fill(color) {
	return {fill:{patternType:"solid", fgColor:color}};
}

function table_style_font(color, bold) {
	var font = {};
	if(color) font.color = color;
	if(bold) font.bold = 1;
	return {font:font};
}

function table_style_border(side, color, style) {
	var border = {}, out = {border:border};
	border[side] = {style:style || "thin", color:color};
	return out;
}

function table_style_accent(themes, index) {
	var position = (index - 1) % 7;
	/* The first style in each seven-style family is neutral. The remaining
	 * six entries use the workbook's accent1..accent6 colors. */
	return table_style_color(themes, position ? 3 + position : 1, 0);
}

function built_in_table_style_rules(name, themes) {
	var match = /^TableStyle(Light|Medium|Dark)(\d+)$/i.exec(name || "");
	if(!match) return null;
	var family = match[1].toLowerCase(), index = parseInt(match[2], 10);
	var maximum = family == "light" ? 21 : family == "medium" ? 28 : 11;
	if(index < 1 || index > maximum) return null;
	var accent = table_style_accent(themes, index);
	var white = table_style_color(themes, 0, 0);
	var dark = table_style_color(themes, 1, 0);
	var rules = {};

	if(family == "light") {
		rules.headerRow = table_style_merge(
			table_style_font(accent, true),
			table_style_border("bottom", accent, "medium")
		);
		rules.totalRow = table_style_merge(
			table_style_font(accent, true),
			table_style_border("top", accent, "double")
		);
		rules.firstRowStripe = table_style_fill(table_style_color(themes, accent.theme, 0.9));
		rules.firstColumnStripe = rules.firstRowStripe;
	} else if(family == "medium") {
		var group = Math.floor((index - 1) / 7);
		rules.headerRow = table_style_merge(table_style_fill(accent), table_style_font(white, true));
		rules.totalRow = table_style_merge(
			table_style_font(accent, true),
			table_style_border("top", accent, "double")
		);
		if(group == 1) {
			/* Medium 8..14 use two theme-derived body fills. This is the family
			 * used by Excel's TableStyleMedium13. */
			rules.wholeTable = table_style_fill(table_style_color(themes, accent.theme, 0.8));
			rules.firstRowStripe = table_style_fill(table_style_color(themes, accent.theme, 0.6));
			rules.firstColumnStripe = rules.firstRowStripe;
		} else if(group == 2) {
			rules.firstRowStripe = table_style_fill(table_style_color(themes, accent.theme, 0.8));
			rules.firstColumnStripe = rules.firstRowStripe;
			rules.wholeTable = table_style_border("bottom", table_style_color(themes, accent.theme, 0.4), "thin");
		} else if(group == 3) {
			rules.wholeTable = table_style_fill(table_style_color(themes, accent.theme, 0.9));
			rules.firstRowStripe = table_style_fill(table_style_color(themes, accent.theme, 0.7));
			rules.firstColumnStripe = rules.firstRowStripe;
		} else {
			rules.firstRowStripe = table_style_fill(table_style_color(themes, accent.theme, 0.8));
			rules.firstColumnStripe = rules.firstRowStripe;
		}
	} else {
		rules.wholeTable = table_style_merge(table_style_fill(dark), table_style_font(white, false));
		rules.headerRow = table_style_merge(table_style_fill(accent), table_style_font(white, true));
		rules.firstRowStripe = table_style_fill(table_style_color(themes, accent.theme, -0.35));
		rules.firstColumnStripe = rules.firstRowStripe;
		rules.totalRow = table_style_merge(
			table_style_font(white, true),
			table_style_border("top", accent, "double")
		);
	}

	rules.firstColumn = table_style_font(null, true);
	rules.lastColumn = table_style_font(null, true);
	return rules;
}

function custom_table_style_rules(name, styles) {
	var tableStyles = styles && styles.TableStyles;
	if(!tableStyles || !tableStyles.styles) return null;
	for(var i = 0; i < tableStyles.styles.length; ++i) {
		var tableStyle = tableStyles.styles[i];
		if(tableStyle.name != name) continue;
		var rules = {};
		(tableStyle.elements || []).forEach(function(element) {
			var dxf = styles.Dxfs && styles.Dxfs[element.dxfId];
			if(dxf) rules[element.type] = {style:dup(dxf), size:element.size || 1};
		});
		return rules;
	}
	return null;
}

function parse_table_xml(data, path, themes, styles, opts) {
	if(!data || data.length > TABLE_STYLE_MAX_XML_LENGTH) return null;
	data = remove_doctype(str_remove_ng(data, "<!--", "-->"));
	var table = {columns:[]}, column = null, pass = false;
	(data.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x), tag = strip_ns(y[0]);
		switch(tag) {
			case '<table': case '<table>':
				table.id = table_style_int(y, "id");
				table.name = utf8read(unescapexml(table_style_attr(y, "name") || ""));
				table.displayName = utf8read(unescapexml(table_style_attr(y, "displayName") || table.name));
				table.ref = table_style_attr(y, "ref") || "";
				table.headerRowCount = table_style_int(y, "headerRowCount");
				table.totalsRowCount = table_style_int(y, "totalsRowCount");
				table.totalsRowShown = table_style_bool(y, "totalsRowShown", false);
				table.headerRowDxfId = table_style_int(y, "headerRowDxfId");
				table.dataDxfId = table_style_int(y, "dataDxfId");
				table.totalsRowDxfId = table_style_int(y, "totalsRowDxfId");
				table.path = path;
				break;
			case '<tableColumn': case '<tableColumn>': case '<tableColumn/>':
				column = {
					id:table_style_int(y, "id"),
					name:utf8read(unescapexml(table_style_attr(y, "name") || "")),
					headerRowDxfId:table_style_int(y, "headerRowDxfId"),
					dataDxfId:table_style_int(y, "dataDxfId"),
					totalsRowDxfId:table_style_int(y, "totalsRowDxfId")
				};
				table.columns.push(column);
				if(tag.slice(-2) == "/>" ) column = null;
				break;
			case '</tableColumn>': column = null; break;
			case '<tableStyleInfo': case '<tableStyleInfo>': case '<tableStyleInfo/>':
				table.styleInfo = {
					name:utf8read(unescapexml(table_style_attr(y, "name") || "")),
					showFirstColumn:table_style_bool(y, "showFirstColumn", false),
					showLastColumn:table_style_bool(y, "showLastColumn", false),
					showRowStripes:table_style_bool(y, "showRowStripes", false),
					showColumnStripes:table_style_bool(y, "showColumnStripes", false)
				};
				break;
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF && !pass) {
				/* Table formulas and filter metadata are intentionally preserved by
				 * the raw XML parser but do not affect visual style resolution. */
				if(/^<\/?(?:tableColumns|autoFilter|sortState|calculatedColumnFormula|totalsRowFormula|xmlColumnPr)/.test(tag)) break;
			}
		}
	});
	if(!table.ref) return null;
	try {
		table.range = safe_decode_range(table.ref);
		if(table.range.s.r < 0 || table.range.s.c < 0 || table.range.e.r > 1048575 || table.range.e.c > 16383) return null;
	} catch(e) { return null; }
	var styleName = table.styleInfo && table.styleInfo.name;
	table.styleRules = custom_table_style_rules(styleName, styles) || built_in_table_style_rules(styleName, themes) || {};
	table.dxfs = styles && styles.Dxfs ? styles.Dxfs : [];
	return table;
}

function parse_sheet_tables(sheet, zip, path, rels, opts, themes, styles) {
	if(!sheet || !rels || !rels['!id'] || !opts || !opts.cellStyles) return;
	var tables = [], seen = Object.create(null);
	keys(rels['!id']).forEach(function(id) {
		if(tables.length >= TABLE_STYLE_MAX_TABLES) return;
		var rel = rels['!id'][id];
		if(!rel || rel.TargetMode == "External" || typeof rel.Target != "string" || !rel.Target ||
			(rel.Type != RELS.TABLE && !/\/table$/.test(rel.Type || ""))) return;
		var tablePath = resolve_path(rel.Target, path);
		if(seen[tablePath]) return;
		seen[tablePath] = true;
		var table = parse_table_xml(getzipstr(zip, tablePath, true), tablePath, themes, styles, opts);
		if(table) tables.push(table);
	});
	if(tables.length) sheet['!tables'] = tables;
}

function table_style_rule_value(rule) {
	return rule && rule.style ? rule.style : rule;
}

function table_style_stripe_rule(first, second, offset) {
	var firstSize = first && first.size || 1, secondSize = second && second.size || 1;
	var period = firstSize + secondSize;
	if(!period) return null;
	return offset % period < firstSize ? first : second;
}

function table_style_dxf(table, id) {
	return id == null || !table.dxfs ? null : table.dxfs[id];
}

function resolve_table_cell_style(ws, row, col, baseStyle) {
	var tables = ws && ws['!tables'];
	var resolved = baseStyle ? dup(baseStyle) : {};
	if(!tables || !tables.length) return keys(resolved).length ? resolved : void 0;
	for(var i = 0; i < tables.length; ++i) {
		var table = tables[i], range = table.range;
		if(!range || row < range.s.r || row > range.e.r || col < range.s.c || col > range.e.c) continue;
		var rules = table.styleRules || {}, info = table.styleInfo || {};
		var rowOffset = row - range.s.r, colOffset = col - range.s.c;
		var headerRows = table.headerRowCount == null ? 1 : table.headerRowCount;
		var totalRows = table.totalsRowCount != null ? table.totalsRowCount : table.totalsRowShown ? 1 : 0;
		var dataStart = headerRows, dataEnd = range.e.r - range.s.r - totalRows;
		var isHeader = rowOffset < headerRows;
		var isTotal = totalRows > 0 && rowOffset > dataEnd;
		var dataOffset = rowOffset - dataStart;
		var column = table.columns && table.columns[colOffset];

		resolved = table_style_merge(resolved, table_style_rule_value(rules.wholeTable));
		if(!isHeader && !isTotal && dataOffset >= 0) {
			if(info.showRowStripes) resolved = table_style_merge(resolved, table_style_rule_value(
				table_style_stripe_rule(rules.firstRowStripe, rules.secondRowStripe, dataOffset)
			));
			if(info.showColumnStripes) resolved = table_style_merge(resolved, table_style_rule_value(
				table_style_stripe_rule(rules.firstColumnStripe, rules.secondColumnStripe, colOffset)
			));
		}
		if(info.showFirstColumn && col == range.s.c) resolved = table_style_merge(resolved, table_style_rule_value(rules.firstColumn));
		if(info.showLastColumn && col == range.e.c) resolved = table_style_merge(resolved, table_style_rule_value(rules.lastColumn));
		if(isHeader) {
			resolved = table_style_merge(resolved, table_style_rule_value(rules.headerRow));
			if(col == range.s.c) resolved = table_style_merge(resolved, table_style_rule_value(rules.firstHeaderCell));
			if(col == range.e.c) resolved = table_style_merge(resolved, table_style_rule_value(rules.lastHeaderCell));
			resolved = table_style_merge(resolved, table_style_dxf(table, table.headerRowDxfId));
			resolved = table_style_merge(resolved, table_style_dxf(table, column && column.headerRowDxfId));
		} else if(isTotal) {
			resolved = table_style_merge(resolved, table_style_rule_value(rules.totalRow));
			if(col == range.s.c) resolved = table_style_merge(resolved, table_style_rule_value(rules.firstTotalCell));
			if(col == range.e.c) resolved = table_style_merge(resolved, table_style_rule_value(rules.lastTotalCell));
			resolved = table_style_merge(resolved, table_style_dxf(table, table.totalsRowDxfId));
			resolved = table_style_merge(resolved, table_style_dxf(table, column && column.totalsRowDxfId));
		} else {
			resolved = table_style_merge(resolved, table_style_dxf(table, table.dataDxfId));
			resolved = table_style_merge(resolved, table_style_dxf(table, column && column.dataDxfId));
		}
	}
	return keys(resolved).length ? resolved : void 0;
}
