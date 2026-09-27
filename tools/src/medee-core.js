/* Засварын мэдээ (Mine2TL repair log): browser + Node.
 * Port of .claude/skills/zasvariin-medee/scripts/medee_process.py; the output
 * matches it cell for cell. The source package is reused so every cell keeps
 * its own font, border and number format; only the sheet, styles and a table
 * part are rewritten. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("jszip"), require("xlsx"));
  } else {
    root.MedeeCore = factory(root.JSZip, root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, XLSX) {
  "use strict";

  var UX = "Ухаа худаг", CX = "Цагаан хад", LOCS = [UX, CX];
  var BLUE = "FF2EAAFC", GRAY = "FFD6D6D6";
  var ACC_ID = 43;   // built-in _(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)
  var WIDTHS = [15.1, 12.6, 17.4, 8.6, 12.8, 17.1, 10.6, 12.1, 12.3, 25.3, 15.9, 9.2, 12.7, 14.9, 15.4, 12.6, 17.2, 13.4, 13.8, 13.4, 14.0];
  var TU_RE = /т[үуy]\s*-?\s*(\d(?:\.\d)?)/i;
  var NEED = ["засварын төрөл", "техникийн №", "жолооч/нэр, овог/", "засварын байршил", "үйлчилгээ", "дис тайлбар",
    "мех тайлбар", "механик", "хүлээсэн цаг", "зарцуулсан цаг", "хүсэлт өгсөн огноо", "эхэлсэн огноо",
    "дууссан огноо", "засварын код", "машины төлөв"];
  var NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  var REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  function norm(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase(); }
  function empty(v) { return v == null || v === "" || v === 0 || v === false; }   // Python truthiness
  function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function colLetter(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function pad(n) { return String(n).padStart(2, "0"); }
  function attr(tag, name) { var m = new RegExp("\\s" + name + '="([^"]*)"').exec(tag); return m ? m[1] : null; }
  function setAttr(tag, name, val) {
    return new RegExp("\\s" + name + '="').test(tag) ? tag.replace(new RegExp("(\\s" + name + '=")[^"]*"'), "$1" + val + '"')
      : tag.replace(/^<(\w+)/, "<$1 " + name + '="' + val + '"');
  }
  function serialText(v) {
    var d = new Date(Math.round((v - 25569) * 86400000));
    return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate()) + " " + pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes());
  }
  function parseEnd(s) {
    var m = /^\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?\s*$/.exec(String(s || ""));
    if (!m) return null;
    return Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) / 86400000 + 25569;
  }
  /* Excel's coercion of "hh:mm" text times 24 (blank -> 0). */
  function hmHours(t) {
    if (t == null || t === "") return 0;
    if (typeof t === "number") return t * 24;
    var m = /^\s*(\d+):(\d{1,2})(?::(\d{1,2}))?\s*$/.exec(String(t));
    if (!m) throw new Error("цагийн утга танигдсангүй: " + t);
    return +m[1] + +m[2] / 60 + +(m[3] || 0) / 3600;
  }
  function isReddish(rgb) {
    if (!rgb || rgb.length < 6) return false;
    var r = parseInt(rgb.slice(-6, -4), 16), g = parseInt(rgb.slice(-4, -2), 16), b = parseInt(rgb.slice(-2), 16);
    return r > 180 && g < 120 && b < 120;
  }

  /* ---------- package reading ---------- */
  async function openSource(buf) {
    var zip = await JSZip.loadAsync(buf);
    var wbx = await zip.file("xl/workbook.xml").async("string");
    var rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    var first = /<sheet\b[^>]*>/.exec(wbx)[0], rid = attr(first, "r:id");
    var rel = new RegExp('<Relationship\\b[^>]*Id="' + rid + '"[^>]*>').exec(rels)[0];
    var target = attr(rel, "Target").replace(/^\//, "");
    var sheetPath = /^xl\//.test(target) ? target : "xl/" + target;
    var sheetXml = await zip.file(sheetPath).async("string");
    var styles = await zip.file("xl/styles.xml").async("string");
    var st = parseStyles(styles);
    // style index of every cell (values come from SheetJS)
    var cellS = {};
    sheetXml.replace(/<c\b[^>]*?\/?>/g, function (tag) {
      var r = attr(tag, "r"), s = attr(tag, "s");
      if (r) cellS[r] = s == null ? 0 : +s;
      return tag;
    });
    var wb = XLSX.read(buf, { type: "array", cellFormula: false });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
    var grid = [];
    for (var r = range.s.r; r <= range.e.r; r++) {
      var row = [];
      for (var c = 0; c <= range.e.c; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        row.push(cell && cell.v !== undefined && cell.v !== "" ? cell.v : null);
      }
      grid.push(row);
    }
    return { zip: zip, sheetPath: sheetPath, styles: styles, st: st, cellS: cellS, grid: grid, ncol: range.e.c + 1 };
  }

  function parseStyles(xml) {
    var block = function (tag) { var m = new RegExp("<" + tag + "\\b[^>]*>([\\s\\S]*?)</" + tag + ">").exec(xml); return m ? m[1] : ""; };
    var xfs = block("cellXfs").match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g) || [];
    var fills = block("fills").match(/<fill\b[^>]*?(?:\/>|>[\s\S]*?<\/fill>)/g) || [];
    var borders = block("borders").match(/<border\b[^>]*?(?:\/>|>[\s\S]*?<\/border>)/g) || [];
    return { xfs: xfs, fills: fills, borders: borders };
  }
  function fillRgb(st, s) {
    var xf = st.xfs[s] || "";
    var fid = +(attr(xf.match(/^<xf\b[^>]*>/)[0], "fillId") || 0);
    var f = st.fills[fid] || "";
    if (!/patternType="(solid|[a-z]+)"/i.test(f) || /patternType="none"/.test(f)) return null;
    var m = /<fgColor\b[^>]*rgb="([0-9A-Fa-f]{6,8})"/.exec(f);
    return m ? m[1].toUpperCase() : null;
  }

  /* ---------- analysis (steps 1-8) ---------- */
  function analyse(src, opts) {
    opts = opts || {};
    var grid = src.grid, header = grid[0] || [];
    var names = header.map(norm);
    var missing = NEED.filter(function (n) { return names.indexOf(n) < 0; });
    if (missing.length) throw new Error("Багана олдсонгүй: " + missing.join(", "));
    var drop = names.indexOf("хэсгийн №");
    var keep = []; for (var i = 0; i < src.ncol; i++) if (i !== drop) keep.push(i);
    var H = keep.map(function (i) { return header[i]; });
    var I = {}; H.forEach(function (h, j) { I[norm(h)] = j; });

    var rows = [], dropped = 0;
    for (var r = 1; r < grid.length; r++) {
      var raw = grid[r], vals = keep.map(function (i) { return raw[i]; });
      if (empty(vals[I["засварын төрөл"]]) || empty(vals[I["техникийн №"]])) {
        if (raw.some(function (v) { return v != null; })) dropped++;
        continue;
      }
      var s0 = src.cellS["A" + (r + 1)];
      rows.push({ v: vals, src: r + 1, s: keep.map(function (i) { var k = src.cellS[colLetter(i + 1) + (r + 1)]; return k == null ? 0 : k; }),
        fill: s0 == null ? null : fillRgb(src.st, s0) });
    }
    var g = function (row, k) { return row.v[I[k]]; };
    var set = function (row, k, v, why) {
      var old = row.v[I[k]];
      if (old === v) return;
      (row.changes || (row.changes = [])).push(H[I[k]] + ": " + (old == null || old === "" ? "(хоосон)" : old) + " → " + v + (why ? " (" + why + ")" : ""));
      row.v[I[k]] = v;
    };

    var warnings = [];
    rows.forEach(function (row) {
      var end = g(row, "дууссан огноо");
      row.red = isReddish(row.fill);
      row.blue = end == null;
      if (row.fill === GRAY && end == null) warnings.push(g(row, "техникийн №") + ": саарал өнгөтэй боловч Дууссан огноо хоосон, дуусаагүй гэж тооцов");
      if ((row.fill === BLUE || row.red) && end != null) warnings.push(g(row, "техникийн №") + ": цэнхэр/улаан боловч Дууссан огноотой, дууссан гэж тооцов");
    });

    // mechanic locations: auto from non-Замд work, overridable
    var seen = {}, order = [];
    rows.forEach(function (row) {
      var loc = g(row, "засварын байршил"), m = g(row, "механик");
      if (LOCS.indexOf(loc) < 0) return;
      if (!seen[m]) { seen[m] = {}; seen[m]._order = []; order.push(m); }
      if (!seen[m][loc]) { seen[m][loc] = 0; seen[m]._order.push(loc); }
      seen[m][loc]++;
    });
    var auto = {};
    order.forEach(function (m) {
      if (empty(m)) return;
      var best = null; seen[m]._order.forEach(function (l) { if (best == null || seen[m][l] > seen[m][best]) best = l; });
      auto[m] = best;
    });
    var overrides = opts.mech || {}, mech = Object.assign({}, auto, overrides);
    var needed = [];
    rows.forEach(function (row) {
      if (g(row, "засварын төрөл") === "Дуудлага" || g(row, "засварын байршил") === "Замд") {
        var m = g(row, "механик"); if (needed.indexOf(m) < 0) needed.push(m);
      }
    });
    needed.sort(function (a, b) { a = String(a); b = String(b); return a < b ? -1 : a > b ? 1 : 0; });
    var unknown = needed.filter(function (m) { return !mech[m]; });

    var suggestions = [];
    rows.forEach(function (row) {
      if (g(row, "засварын төрөл") === "Төлөвлөгөөт" && empty(g(row, "жолооч/нэр, овог/")) && empty(g(row, "үйлчилгээ")) && empty(g(row, "дис тайлбар"))) {
        var m = TU_RE.exec(String(g(row, "мех тайлбар") || ""));
        suggestions.push({ tech: g(row, "техникийн №"), mex: g(row, "мех тайлбар"), suggest: m ? "ТҮ-" + m[1] : null });
      }
    });
    var mechanics = {};
    Object.keys(mech).concat(needed).forEach(function (m) {
      if (mechanics[m]) return;
      var counts = {}; if (seen[m]) seen[m]._order.forEach(function (l) { counts[l] = seen[m][l]; });
      mechanics[m] = { location: mech[m] || null, auto: auto[m] || null, override: m in overrides, counts: counts, needed: needed.indexOf(m) >= 0 };
    });
    var gray1 = rows.filter(function (r) { return !r.blue; })[0];
    return {
      H: H, I: I, rows: rows, g: g, set: set, mech: mech, warnings: warnings,
      report: { rows: rows.length, dropped_rows: dropped, mechanics: mechanics, mechanics_needed: needed, mechanics_unknown: unknown,
        tu_suggestions: suggestions, red_rows: rows.filter(function (r) { return r.red; }).length, warnings: warnings,
        default_end: gray1 ? serialText(Math.floor(g(gray1, "дууссан огноо")) + 0.25) : null }
    };
  }

  async function checkMedee(buf, opts) {
    var src = await openSource(buf);
    return analyse(src, opts).report;
  }

  /* ---------- build (steps 2-14) ---------- */
  async function buildMedee(buf, opts) {
    opts = opts || {};
    var src = await openSource(buf);
    var A = analyse(src, opts), rows = A.rows, g = A.g, set = A.set, H = A.H, I = A.I, mech = A.mech, warnings = A.warnings;
    var report = A.report;
    if (report.mechanics_unknown.length) throw new Error("Байршил тодорхойгүй механик: " + report.mechanics_unknown.join(", "));
    var tu = opts.tu || {};

    rows.forEach(function (row) { if (g(row, "засварын төрөл") === "Дуудлага") { set(row, "засварын байршил", mech[g(row, "механик")], "2. Дуудлага: механик " + g(row, "механик") + "-ийн байршил"); set(row, "үйлчилгээ", "Дуудлага", "2. Дуудлага"); } });
    rows.forEach(function (row) { if (g(row, "засварын төрөл") === "Төлөвлөгөөт") { set(row, "дис тайлбар", g(row, "үйлчилгээ"), "4. Төлөвлөгөөт: үйлчилгээг Дис тайлбар руу"); set(row, "үйлчилгээ", "Төлөвлөгөөт", "4. Төлөвлөгөөт"); } });
    rows.forEach(function (row) { if (empty(g(row, "жолооч/нэр, овог/"))) set(row, "үйлчилгээ", "Ачигч", "5. Жолооч хоосон"); });
    rows.forEach(function (row) {
      var t = tu[g(row, "техникийн №")];
      if (t && g(row, "засварын төрөл") === "Төлөвлөгөөт" && empty(g(row, "дис тайлбар"))) set(row, "дис тайлбар", t, "5а. Мех тайлбараас, та баталсан");
    });
    rows.forEach(function (row) { if (g(row, "засварын байршил") === "Замд") set(row, "засварын байршил", mech[g(row, "механик")], "6. Замд: механик " + g(row, "механик") + "-ийн байршил"); });

    // step 7
    var services = {}, locs = {}, problems = [];
    rows.forEach(function (r) {
      var sv = g(r, "үйлчилгээ"), lc = g(r, "засварын байршил");
      services[sv == null ? "" : sv] = (services[sv == null ? "" : sv] || 0) + 1;
      locs[lc == null ? "" : lc] = (locs[lc == null ? "" : lc] || 0) + 1;
    });
    ["Ачигч", "Дуудлага", "Төлөвлөгөөт"].forEach(function (w) { if (!services[w]) warnings.push('Үйлчилгээнд "' + w + '" алга'); });
    if (services[""]) problems.push("Үйлчилгээ хоосон мөр үлдсэн");
    var badLocs = Object.keys(locs).filter(function (l) { return LOCS.indexOf(l) < 0; });
    if (badLocs.length) problems.push("Байршилд УХ/ЦХ-ээс өөр утга: " + badLocs.join(", "));

    // step 8
    var blue = rows.filter(function (r) { return r.blue; }), gray = rows.filter(function (r) { return !r.blue; });
    var end;
    if (opts.end) { end = parseEnd(opts.end); if (end == null) throw new Error("Дуусах цагийг YYYY-MM-DD HH:MM хэлбэрээр бичнэ үү"); }
    else if (gray.length) end = Math.floor(g(gray[0], "дууссан огноо")) + 0.25;
    else throw new Error("Дууссан засвар (саарал мөр) алга. Дуусах цагийг өөрөө бичнэ үү");
    blue.forEach(function (r) { var st = g(r, "эхэлсэн огноо"); if (st != null && st > end) warnings.push(g(r, "техникийн №") + ": " + serialText(end).slice(11) + "-аас хойш эхэлсэн"); });
    var ordered = blue.concat(gray);

    // steps 9-14
    var ref = function (k) { return "Table1[[#This Row],[" + H[I[k]] + "]]"; };
    blue.forEach(function (r) {
      var st = g(r, "эхэлсэн огноо"), rq = g(r, "хүсэлт өгсөн огноо");
      if (st == null || rq == null) { problems.push(g(r, "техникийн №") + ": Эхэлсэн эсвэл Хүсэлт өгсөн огноо хоосон"); return; }
      var wait = (st - rq) * 24, spent = (end - st) * 24;
      r.calc = {
        "механик": ["=" + ref("эхэлсэн огноо") + "*24-" + ref("хүсэлт өгсөн огноо") + "*24", wait],
        "хүлээсэн цаг": ["=" + ref("дууссан огноо") + "*24-" + ref("эхэлсэн огноо") + "*24", (spent - st) * 24],
        "зарцуулсан цаг": [null, wait], "дууссан огноо": [null, spent], "засварын код": [null, wait],
        "машины төлөв": ["=" + ref("дууссан огноо") + "+" + ref("засварын код"), spent + wait]
      };
    });
    gray.forEach(function (r) {
      var sp, wt;
      try { sp = hmHours(g(r, "зарцуулсан цаг")); wt = hmHours(g(r, "хүлээсэн цаг")); }
      catch (e) { problems.push(g(r, "техникийн №") + ": " + e.message); return; }
      r.calc = {
        "дууссан огноо": ["=" + ref("зарцуулсан цаг") + "*24", sp],
        "засварын код": ["=" + ref("хүлээсэн цаг") + "*24", wt],
        "машины төлөв": ["=" + ref("дууссан огноо") + "+" + ref("засварын код"), sp + wt]
      };
    });
    if (problems.length) { var err = new Error(problems.join("; ")); err.problems = problems; throw err; }

    var bytes = await write(src, H, ordered);
    report.end = serialText(end); report.blue = blue.length; report.gray = gray.length;
    report.services = services; report.locations = locs; report.warnings = warnings;
    report.fileName = "ZM_" + serialText(end).slice(8, 10) + ".xlsx";
    // explanation: one line per output row, and the formulas used
    report.explain = {
      headers: H, mech: mech, end: serialText(end),
      rows: ordered.map(function (r, i) {
        return { row: i + 2, tech: g(r, "техникийн №"), color: r.blue ? (r.red ? "улаан → цэнхэр" : "цэнхэр") : "саарал",
          changes: (r.changes || []).join("; ") || "өөрчлөгдөөгүй",
          calc: Object.keys(r.calc).map(function (k) {
            var f = r.calc[k][0], v = Math.round(r.calc[k][1] * 100) / 100;
            return H[I[k]] + ": " + (f ? f.slice(1).replace(/Table1\[\[#This Row\],(\[[^\]]+\])\]/g, "$1") + " = " + v : "тоон утга " + v);
          }).join("; ") };
      }),
      formulas: {
        blue: Object.keys(blue[0] ? blue[0].calc : {}).map(function (k) { return [H[I[k]], blue[0].calc[k][0] || "тоон утга"]; }),
        gray: Object.keys(gray[0] ? gray[0].calc : {}).map(function (k) { return [H[I[k]], gray[0].calc[k][0] || "тоон утга"]; })
      }
    };
    return { bytes: bytes, report: report };
  }

  /* ---------- writing ---------- */
  async function write(src, H, ordered) {
    var zip = src.zip, styles = src.styles, st = src.st;
    var fills = st.fills.slice(), xfs = st.xfs.slice(), cache = {};
    var blueFill = fills.findIndex(function (f) { return new RegExp('rgb="' + BLUE + '"', "i").test(f) && /solid/.test(f); });
    if (blueFill < 0) { fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + BLUE + '"/><bgColor indexed="64"/></patternFill></fill>'); blueFill = fills.length - 1; }
    var borders = st.borders.slice(), noTop = {};
    var grayFill = fills.findIndex(function (f) { return new RegExp('rgb="' + GRAY + '"', "i").test(f) && /solid/.test(f); });
    if (grayFill < 0) { fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + GRAY + '"/><bgColor indexed="64"/></patternFill></fill>'); grayFill = fills.length - 1; }
    // o: {fill, numFmt, left: true (horizontal left), noTop: true (no top border)}
    function derive(s, o) {
      var key = s + "|" + JSON.stringify(o);
      if (cache[key] != null) return cache[key];
      var xf = xfs[s] || xfs[0], open = xf.match(/^<xf\b[^>]*?(\/?)>/)[0], selfClose = /\/>$/.test(open);
      var tag = open.replace(/\s*\/?>$/, ""), body = selfClose ? "" : xf.slice(open.length).replace(/<\/xf>$/, "");
      if (o.fill != null) { tag = setAttr(tag, "fillId", o.fill); tag = setAttr(tag, "applyFill", "1"); }
      if (o.numFmt != null) { tag = setAttr(tag, "numFmtId", o.numFmt); tag = setAttr(tag, "applyNumberFormat", "1"); }
      if (o.noTop) {
        var bid = +(attr(tag, "borderId") || 0);
        if (noTop[bid] == null) {
          var b = borders[bid] || "<border/>";
          b = /<top\b/.test(b) ? b.replace(/<top\b[^>]*?(?:\/>|>[\s\S]*?<\/top>)/, "<top/>") : b;
          borders.push(b); noTop[bid] = borders.length - 1;
        }
        tag = setAttr(tag, "borderId", noTop[bid]); tag = setAttr(tag, "applyBorder", "1");
      }
      if (o.left) {
        body = /<alignment\b/.test(body) ? body.replace(/<alignment\b[^>]*?\/?>/, function (a) { return setAttr(a.replace(/\s*\/?>$/, ""), "horizontal", "left") + (/\/>$/.test(a) ? "/>" : ">"); })
          : '<alignment horizontal="left"/>' + body;
        tag = setAttr(tag, "applyAlignment", "1");
      }
      var out = body ? tag + ">" + body + "</xf>" : tag + "/>";
      if (out === xf) return (cache[key] = s);
      xfs.push(out);
      return (cache[key] = xfs.length - 1);
    }
    var CALC = ["дууссан огноо", "засварын код", "машины төлөв"];
    var calcFmt = function (k) { return CALC.indexOf(k) >= 0 ? ACC_ID : 0; };
    var IDX = {}; H.forEach(function (h, j) { IDX[norm(h)] = j; });
    var ncol = H.length, last = colLetter(ncol), nrow = ordered.length + 1;
    var cellXml = function (ref, s, v, f) {
      var sa = s ? ' s="' + s + '"' : "";
      if (f) return '<c r="' + ref + '"' + sa + "><f>" + esc(f.slice(1)) + "</f><v>" + v + "</v></c>";
      if (v == null || v === "") return '<c r="' + ref + '"' + sa + "/>";
      if (typeof v === "number") return '<c r="' + ref + '"' + sa + "><v>" + v + "</v></c>";
      if (typeof v === "boolean") return '<c r="' + ref + '"' + sa + ' t="b"><v>' + (v ? 1 : 0) + "</v></c>";
      return '<c r="' + ref + '"' + sa + ' t="inlineStr"><is><t xml:space="preserve">' + esc(v) + "</t></is></c>";
    };
    // header keeps the source header cells' styles
    var keepCols = [], header = src.grid[0].map(norm), drop = header.indexOf("хэсгийн №");
    for (var i = 0; i < src.ncol; i++) if (i !== drop) keepCols.push(i);
    // as in ЗМ 27: header without a top border, the hour columns' headers in General
    var sd = ['<row r="1" ht="12.75" customHeight="1">' + H.map(function (h, j) {
      var o = { noTop: true }; if (CALC.indexOf(norm(h)) >= 0) o.numFmt = 0;
      return cellXml(colLetter(j + 1) + "1", derive(src.cellS[colLetter(keepCols[j] + 1) + "1"] || 0, o), h);
    }).join("") + "</row>"];
    ordered.forEach(function (r, n) {
      var rn = n + 2, cells = [];
      for (var j = 0; j < ncol; j++) {
        var s = r.s[j], v = r.v[j], f = null, fmt = null;
        var k = Object.keys(r.calc || {}).filter(function (x) { return IDX[x] === j; })[0];
        if (k) { f = r.calc[k][0]; v = r.calc[k][1]; fmt = calcFmt(k); }
        var o = {};
        if (r.blue) o.fill = blueFill;
        if (fmt != null) o.numFmt = fmt;
        // as in ЗМ 27: on blue rows the hour columns are gray (Засварын код left-aligned),
        // and so is Үйлчилгээ when it became Төлөвлөгөөт
        if (r.blue && CALC.indexOf(k) >= 0) { o.fill = grayFill; if (k === "засварын код") o.left = true; }
        if (r.blue && j === IDX["үйлчилгээ"] && v === "Төлөвлөгөөт") o.fill = grayFill;
        s = derive(s, o);
        cells.push(cellXml(colLetter(j + 1) + rn, s, v, f));
      }
      sd.push('<row r="' + rn + '" ht="15.75" customHeight="1">' + cells.join("") + "</row>");
    });
    var cols = "<cols>" + H.map(function (h, j) { return '<col min="' + (j + 1) + '" max="' + (j + 1) + '" width="' + (WIDTHS[j] || 12) + '" customWidth="1"/>'; }).join("") + "</cols>";
    var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ' + NS + '><dimension ref="A1:' + last + nrow + '"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
      cols + "<sheetData>" + sd.join("") + '</sheetData><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><tableParts count="1"><tablePart r:id="rIdTbl1"/></tableParts></worksheet>';

    // drop any old tables, add Table1
    Object.keys(zip.files).forEach(function (p) { if (/^xl\/tables\//.test(p) || p === "xl/calcChain.xml") zip.remove(p); });
    var ct = await zip.file("[Content_Types].xml").async("string");
    ct = ct.replace(/<Override\b[^>]*PartName="\/xl\/(tables\/[^"]*|calcChain\.xml)"[^>]*\/>/g, "");
    ct = ct.replace("</Types>", '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>');
    zip.file("[Content_Types].xml", ct);
    var wrels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    zip.file("xl/_rels/workbook.xml.rels", wrels.replace(/<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/g, ""));
    var table = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' + last + nrow + '" totalsRowShown="0"><autoFilter ref="A1:' + last + nrow + '"/>' +
      '<tableColumns count="' + ncol + '">' + H.map(function (h, j) { return '<tableColumn id="' + (j + 1) + '" name="' + esc(h) + '"/>'; }).join("") + "</tableColumns>" +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
    zip.file("xl/tables/table1.xml", table);
    var relPath = src.sheetPath.replace(/([^/]+)$/, "_rels/$1.rels");
    var up = src.sheetPath.split("/").length - 2;   // xl/worksheets/sheet1.xml -> ../tables
    zip.file(relPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTbl1" Type="' + REL + '/table" Target="' + "../".repeat(up) + 'tables/table1.xml"/></Relationships>');
    zip.file(src.sheetPath, sheet);
    // styles: extra fill and derived xfs
    styles = styles.replace(/<fills\b[^>]*>[\s\S]*?<\/fills>/, '<fills count="' + fills.length + '">' + fills.join("") + "</fills>")
      .replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/, '<borders count="' + borders.length + '">' + borders.join("") + "</borders>")
      .replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, '<cellXfs count="' + xfs.length + '">' + xfs.join("") + "</cellXfs>");
    zip.file("xl/styles.xml", styles);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  return { checkMedee: checkMedee, buildMedee: buildMedee, parseEnd: parseEnd, LOCS: LOCS };
});
