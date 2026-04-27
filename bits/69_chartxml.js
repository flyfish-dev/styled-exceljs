function parse_Cache(data/*:string*/)/*:[Array<number|string>, string, ?string]*/ {
	var col/*:Array<number|string>*/ = [];
	var num = data.match(/^<c:numCache>/);
	var f;

	/* 21.2.2.150 pt CT_NumVal */
	(data.match(/<c:pt idx="(\d*)"[^<>\/]*><c:v>([^<]*)<\/c:v><\/c:pt>/mg)||[]).forEach(function(pt) {
		var q = pt.match(/<c:pt idx="(\d*)"[^<>\/]*><c:v>([^<]*)<\/c:v><\/c:pt>/);
		if(!q) return;
		col[+q[1]] = num ? +q[2] : q[2];
	});

	/* 21.2.2.71 formatCode CT_Xstring */
	var nf = unescapexml((str_match_xml(data, "c:formatCode") || ["","General"])[1]);

	(str_match_ng(data, "<c:f>", "</c:f>")||[]).forEach(function(F) { f = F.replace(/<[^<>]*>/g,""); });

	return [col, nf, f];
}

function parse_chart_cache(data) {
	var nc = str_match_ng(data, "<c:numCache>", "</c:numCache>");
	if(nc && nc.length) {
		var n = parse_Cache(nc[0]);
		return {values:n[0], formatCode:n[1], formula:n[2]};
	}
	var sc = str_match_ng(data, "<c:strCache>", "</c:strCache>");
	if(sc && sc.length) {
		var s = parse_Cache(sc[0]);
		return {values:s[0], formatCode:s[1], formula:s[2]};
	}
	var f = (str_match_ng(data, "<c:f>", "</c:f>")||[])[0];
	return {values:[], formula:f ? f.replace(/<[^<>]*>/g,"") : void 0};
}

function parse_chart_tx(data) {
	var tx = str_match_ng(data, "<c:tx>", "</c:tx>");
	if(!tx || !tx.length) return "";
	var v = str_match_xml_ns(tx[0], "v");
	if(v && v[1]) return unescapexml(v[1]);
	var f = str_match_xml_ns(tx[0], "f");
	if(f && f[1]) return unescapexml(f[1]);
	return "";
}

function parse_chart_series(data) {
	var ser = ({name:parse_chart_tx(data)}/*:any*/);
	var idx = str_match_xml_ns(data, "idx");
	if(idx) ser.idx = +(parsexmltag(idx[0]).val || 0);
	var order = str_match_xml_ns(data, "order");
	if(order) ser.order = +(parsexmltag(order[0]).val || 0);
	["cat","val","xVal","yVal","bubbleSize"].forEach(function(k) {
		var m = str_match_xml_ns(data, k);
		if(m) ser[k] = parse_chart_cache(m[0]);
	});
	if(ser.val && ser.val.values) ser.data = ser.val.values;
	else if(ser.yVal && ser.yVal.values) ser.data = ser.yVal.values;
	return ser;
}

function parse_chart_title(data) {
	var title = str_match_xml_ns(data, "title");
	if(!title) return "";
	var out = [];
	(title[0].match(/<a:t\b[^>]*>[\s\S]*?<\/a:t>/g)||[]).forEach(function(t) { out.push(unescapexml(t.replace(/<[^>]*>/g, ""))); });
	return out.join("");
}

function parse_chart_model(data, name, rels) {
	var model = ({target:name, raw:data, rels:rels, series:[]}/*:any*/);
	model.title = parse_chart_title(data);
	var plot = str_match_xml_ns(data, "plotArea");
	var body = plot ? plot[1] : data;
	["barChart","lineChart","areaChart","scatterChart","pieChart","doughnutChart","bubbleChart"].forEach(function(type) {
		(str_match_ng(body, "<c:" + type + ">", "</c:" + type + ">")||[]).forEach(function(chartXml) {
			if(!model.type) model.type = type;
			var grouping = str_match_xml_ns(chartXml, "grouping");
			if(grouping) model.grouping = parsexmltag(grouping[0]).val;
			(str_match_ng(chartXml, "<c:ser>", "</c:ser>")||[]).forEach(function(serXml) {
				var ser = parse_chart_series(serXml);
				ser.chartType = type;
				model.series.push(ser);
			});
		});
	});
	var legend = str_match_xml_ns(data, "legend");
	if(legend) {
		model.legend = {};
		var pos = str_match_xml_ns(legend[0], "legendPos");
		if(pos) model.legend.position = parsexmltag(pos[0]).val;
	}
	return model;
}

/* 21.2 DrawingML - Charts */
function parse_chart(data/*:?string*/, name/*:string*/, opts, rels, wb, csheet) {
	var cs/*:Worksheet*/ = ((csheet || {"!type":"chart"})/*:any*/);
	if(!data) return csheet;
	/* 21.2.2.27 chart CT_Chart */
	cs["!chart"] = parse_chart_model(data, name, rels);

	var C = 0, R = 0, col = "A";
	var refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };

	/* 21.2.2.120 numCache CT_NumData */
	(str_match_ng(data, "<c:numCache>", "</c:numCache>")||[]).forEach(function(nc) {
		var cache = parse_Cache(nc);
		refguess.s.r = refguess.s.c = 0;
		refguess.e.c = C;
		col = encode_col(C);
		cache[0].forEach(function(n,i) {
			if(cs["!data"]) {
				if(!cs["!data"][i]) cs["!data"][i] = [];
				cs["!data"][i][C] = {t:'n', v:n, z:cache[1] };
			} else cs[col + encode_row(i)] = {t:'n', v:n, z:cache[1] };
			R = i;
		});
		if(refguess.e.r < R) refguess.e.r = R;
		++C;
	});
	if(C > 0) cs["!ref"] = encode_range(refguess);
	return cs;
}
