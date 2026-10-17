/* Repair request report (Mine2TL export): browser + Node.
 * Port of .claude/skills/zasvariin-huselt/scripts/zasvar_report.py.
 * The export workbook is kept byte-identical; two sheets are injected in front
 * of it at the XML level, laid out like the team's final report:
 *   Sheet1  pivot layout (Count of Засварын төрөл: байршил > үйлчилгээ × өдөр)
 *   Sheet2  report table at D5 (navy / blue / light), "Нийт хүсэлт" last */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("jszip"), require("xlsx"));
  } else {
    root.ZasvarCore = factory(root.JSZip, root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, XLSX) {
  "use strict";

  var BLANK = "(хоосон)";
  var BUCKETS = [["Өнөөдөр", 0, 0], ["1–3 хоног", 1, 3], ["4–7 хоног", 4, 7], ["8+ хоног", 8, 1e9]];
  var MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var REP_TOP = 5, REP_LEFT = 4; // report table starts at D5
  var NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  function norm(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase(); }
  function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function colLetter(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  // Mongolian alphabetical order (Ө after О, Ү after У), as Excel sorts it
  var MN_ABC = "абвгдеёжзийклмноөпрстуүфхцчшщъыьэюя";
  function mnKey(s) { return String(s).toLowerCase().replace(/[а-яёөү]/g, function (c) { return String.fromCharCode(0x2000 + MN_ABC.indexOf(c)); }); }
  function cmp(a, b) { a = mnKey(a); b = mnKey(b); return a < b ? -1 : a > b ? 1 : 0; }

  /* Dates as "YYYY-MM-DD" keys; Excel serials are converted without time zones. */
  function dayKey(v) {
    if (typeof v === "number" && v > 1) {
      var d = new Date(Math.round((Math.floor(v) - 25569) * 86400000));
      return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
    }
    if (v instanceof Date) return v.getFullYear() + "-" + pad(v.getMonth() + 1) + "-" + pad(v.getDate());
    var s = String(v == null ? "" : v).trim(), m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return m[1] + "-" + pad(m[2]) + "-" + pad(m[3]);
    if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s))) return m[3] + "-" + pad(m[1]) + "-" + pad(m[2]);
    if ((m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})/.exec(s))) return m[3] + "-" + pad(m[2]) + "-" + pad(m[1]);
    return null;
  }
  function pad(n) { return String(n).padStart(2, "0"); }
  function dayNum(k) { var p = k.split("-"); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000; }
  function dayLabel(k) { var p = k.split("-"); return p[2] + "-" + MON[+p[1] - 1]; }

  function readRequests(buf) {
    var wb = XLSX.read(buf, { type: "array" });
    var found = null;
    for (var si = 0; si < wb.SheetNames.length && !found; si++) {
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[si]], { header: 1, raw: true, defval: null });
      for (var hi = 0; hi < Math.min(rows.length, 30); hi++) {
        var names = rows[hi].map(norm);
        if (names.indexOf("засварын төрөл") >= 0 && names.indexOf("үйлчилгээ") >= 0 && names.indexOf("засварын байршил") >= 0) {
          found = { rows: rows, hi: hi, names: names };
          break;
        }
      }
    }
    if (!found) throw new Error('Засварын хүсэлтийн гарчиг ("Засварын төрөл", "Засварын байршил", "Үйлчилгээ") олдсонгүй');
    function find(pred, label) {
      for (var i = 0; i < found.names.length; i++) if (pred(found.names[i])) return i;
      throw new Error('"' + label + '" багана олдсонгүй');
    }
    var ix = {
      type: find(function (n) { return n === "засварын төрөл"; }, "Засварын төрөл"),
      tech: find(function (n) { return n.indexOf("техни") === 0; }, "Техникийн №"),
      loc: find(function (n) { return n === "засварын байршил"; }, "Засварын байршил"),
      srv: find(function (n) { return n === "үйлчилгээ"; }, "Үйлчилгээ"),
      req: find(function (n) { return n.indexOf("хүсэлт гаргасан") === 0; }, "Хүсэлт гаргасан огноо")
    };
    var descIx = found.names.indexOf("тайлбар");
    var out = [], skipped = 0;
    found.rows.slice(found.hi + 1).forEach(function (r) {
      if (r[ix.tech] == null || r[ix.type] == null) { if (r.some(function (c) { return c != null; })) skipped++; return; }
      var day = dayKey(r[ix.req]);
      if (!day) { skipped++; return; }
      out.push({
        type: String(r[ix.type]).trim() || BLANK,
        tech: String(r[ix.tech]).trim(),
        loc: String(r[ix.loc] == null ? "" : r[ix.loc]).trim() || BLANK,
        srv: String(r[ix.srv] == null ? "" : r[ix.srv]).trim() || BLANK,
        day: day,
        desc: descIx >= 0 && r[descIx] != null ? String(r[descIx]).trim() : ""
      });
    });
    return { reqs: out, skipped: skipped };
  }

  /* ---------- styles appended to the export's styles.xml ---------- */
  function addStyles(styles) {
    function append(tag, items) {
      var re = new RegExp("<" + tag + ' count="(\\d+)"([^>]*)>([\\s\\S]*?)</' + tag + ">");
      var m = re.exec(styles);
      if (!m) throw new Error("styles.xml-д <" + tag + "> хэсэг олдсонгүй");
      var base = parseInt(m[1], 10);
      styles = styles.replace(m[0], "<" + tag + ' count="' + (base + items.length) + '"' + m[2] + ">" + m[3] + items.join("") + "</" + tag + ">");
      return base;
    }
    function font(bold, color) { return "<font>" + (bold ? "<b/>" : "") + '<sz val="11"/><color rgb="FF' + color + '"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'; }
    function fill(c) { return '<fill><patternFill patternType="solid"><fgColor rgb="FF' + c + '"/><bgColor indexed="64"/></patternFill></fill>'; }
    var f0 = append("fonts", [font(true, "FFFFFF"), font(false, "000000"), font(true, "000000")]);
    var F = { boldWhite: f0, plain: f0 + 1, bold: f0 + 2 };
    // Office accent 5 (4472C4) with the report's tints: −50 %, −25 %, +40 %; pivot header light blue
    var colors = { navy: "203864", blue: "2F5597", light: "8FAADC", pv: "DDEBF7" };
    var keys = Object.keys(colors);
    var p0 = append("fills", keys.map(function (k) { return fill(colors[k]); }));
    var P = {}; keys.forEach(function (k, i) { P[k] = p0 + i; });
    var side = function (n, c) { return "<" + n + ' style="thin"><color rgb="FF' + c + '"/></' + n + ">"; };
    var b0 = append("borders", [
      "<border>" + side("left", "000000") + side("right", "000000") + side("top", "000000") + side("bottom", "000000") + "<diagonal/></border>",
      "<border><left/><right/><top/>" + side("bottom", "9BC2E6") + "<diagonal/></border>",
      "<border><left/><right/>" + side("top", "9BC2E6") + "<bottom/><diagonal/></border>"
    ]);
    var B = { grid: b0, under: b0 + 1, over: b0 + 2 };
    // [font, fill, border, horizontal, indent]
    var defs = {
      navy: [F.boldWhite, P.navy, B.grid, "center"], blue: [F.boldWhite, P.blue, B.grid, "center"], light: [F.plain, P.light, B.grid, "center"],
      pvHead: [F.bold, P.pv, B.under, "left"], pvLoc: [F.bold, 0, 0, "left"], pvLocN: [F.bold, 0, 0, "general"],
      pvSrv: [F.plain, 0, 0, "left", 1], pvN: [F.plain, 0, 0, "general"], pvTot: [F.bold, P.pv, B.over, "left"], pvTotN: [F.bold, P.pv, B.over, "general"]
    };
    var names = Object.keys(defs);
    var x0 = append("cellXfs", names.map(function (k) {
      var d = defs[k], al = d[3] === "center" ? '<alignment horizontal="center" vertical="center"/>' : d[3] === "left" ? '<alignment horizontal="left"' + (d[4] ? ' indent="' + d[4] + '"' : "") + "/>" : "";
      return '<xf numFmtId="0" fontId="' + d[0] + '" fillId="' + d[1] + '" borderId="' + d[2] + '" xfId="0" applyFont="1" applyFill="1" applyBorder="1"' + (al ? ' applyAlignment="1">' + al + "</xf>" : "/>");
    }));
    var X = {}; names.forEach(function (k, i) { X[k] = x0 + i; });
    return { styles: styles, X: X };
  }

  /* ---------- sheet XML ---------- */
  function Sheet() { this.rows = {}; this.cols = []; this.freeze = false; }
  Sheet.prototype.set = function (r, c, v, s) {
    (this.rows[r] = this.rows[r] || {})[c] = { v: v, s: s };
  };
  Sheet.prototype.xml = function (selected) {
    var rows = Object.keys(this.rows).map(Number).sort(function (a, b) { return a - b; });
    var body = rows.map(function (r) {
      var cells = this.rows[r], cs = Object.keys(cells).map(Number).sort(function (a, b) { return a - b; });
      return '<row r="' + r + '">' + cs.map(function (c) {
        var x = cells[c], ref = colLetter(c) + r, s = x.s != null ? ' s="' + x.s + '"' : "";
        if (x.v == null || x.v === "") return '<c r="' + ref + '"' + s + "/>";
        if (typeof x.v === "number") return '<c r="' + ref + '"' + s + "><v>" + x.v + "</v></c>";
        return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(x.v) + "</t></is></c>";
      }).join("") + "</row>";
    }, this).join("");
    var pane = this.freeze ? '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/>' : "";
    var cols = this.cols.length ? "<cols>" + this.cols.map(function (w, i) { return w ? '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>' : ""; }).join("") + "</cols>" : "";
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ' + NS + '><sheetViews><sheetView workbookViewId="0"' + (selected ? ' tabSelected="1"' : "") + ">" + pane +
      '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>' + cols + "<sheetData>" + body + '</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  };

  function counter(list, key) { var c = {}; list.forEach(function (q) { var k = key(q); c[k] = (c[k] || 0) + 1; }); return c; }
  function uniq(list) { return list.filter(function (x, i) { return list.indexOf(x) === i; }).sort(cmp); }

  function reportSheet(reqs, days, X) {
    var sh = new Sheet(), top = REP_TOP, left = REP_LEFT;
    sh.cols[left - 1] = 15.88671875;
    var ncol = days.length + 2;
    function put(r, label, list, style) {
      var c = counter(list, function (q) { return q.day; });
      sh.set(r, left, label, style);
      days.forEach(function (d, i) { sh.set(r, left + 1 + i, c[d] || null, style); });
      sh.set(r, left + ncol - 1, list.length, style);
    }
    sh.set(top, left, "Засварын төрөл", X.navy);
    days.forEach(function (d, i) { sh.set(top, left + 1 + i, dayLabel(d), X.navy); });
    sh.set(top, left + ncol - 1, "Хүсэлт", X.navy);
    var r = top + 1, table = [];
    uniq(reqs.map(function (q) { return q.loc; })).forEach(function (loc) {
      var inLoc = reqs.filter(function (q) { return q.loc === loc; });
      var services = uniq(inLoc.map(function (q) { return q.srv; })).filter(function (s) { return s !== BLANK; });
      put(r, loc, inLoc, services.length ? X.blue : X.light);
      table.push({ label: loc, loc: true, list: inLoc }); r++;
      services.forEach(function (s) {
        var l = inLoc.filter(function (q) { return q.srv === s; });
        put(r, s, l, X.light); table.push({ label: s, list: l }); r++;
      });
    });
    put(r, "Нийт хүсэлт", reqs, X.navy);
    table.push({ label: "Нийт хүсэлт", total: true, list: reqs });
    return { sheet: sh, table: table };
  }

  /* Values laid out as the Excel PivotTable "Count of Засварын төрөл"
     (rows: байршил > үйлчилгээ, columns: request day). */
  function pivotSheet(reqs, days, X) {
    var sh = new Sheet(), last = days.length + 2;
    sh.cols = [23.21875, 15.5546875].concat(days.slice(1).map(function () { return 6.5546875; })).concat([10.77734375]);
    sh.set(1, 1, "Count of Засварын төрөл", X.pvHead); sh.set(1, 2, "Column Labels", X.pvHead);
    for (var c = 3; c <= last; c++) sh.set(1, c, null, X.pvHead);
    sh.set(2, 1, "Row Labels", X.pvHead);
    days.forEach(function (d, i) { sh.set(2, i + 2, dayLabel(d), X.pvHead); });
    sh.set(2, last, "Grand Total", X.pvHead);
    var r = 3;
    function put(label, list, L, N) {
      var cnt = counter(list, function (q) { return q.day; });
      sh.set(r, 1, label, L);
      days.forEach(function (d, i) { sh.set(r, i + 2, cnt[d] || null, N); });
      sh.set(r, last, list.length, N); r++;
    }
    uniq(reqs.map(function (q) { return q.loc; })).forEach(function (loc) {
      var inLoc = reqs.filter(function (q) { return q.loc === loc; });
      put(loc, inLoc, X.pvLoc, X.pvLocN);
      var srv = uniq(inLoc.map(function (q) { return q.srv; })).filter(function (s) { return s !== BLANK; });
      if (inLoc.some(function (q) { return q.srv === BLANK; })) srv.push(BLANK);
      srv.forEach(function (s) { put(s === BLANK ? "(blank)" : s, inLoc.filter(function (q) { return q.srv === s; }), X.pvSrv, X.pvN); });
    });
    put("Grand Total", reqs, X.pvTot, X.pvTotN);
    return sh;
  }

  /* Same truck (Техникийн №) with more than one request */
  function duplicates(reqs) {
    var by = {};
    reqs.forEach(function (q) { (by[q.tech] = by[q.tech] || []).push(q); });
    return Object.keys(by).filter(function (t) { return by[t].length > 1; }).sort(cmp).map(function (t) { return { tech: t, list: by[t] }; });
  }

  async function buildZasvar(opts) {
    var read = readRequests(opts.input);
    var reqs = read.reqs;
    if (!reqs.length) throw new Error("Засварын хүсэлт олдсонгүй");
    var days = uniq(reqs.map(function (q) { return q.day; }));
    var today = opts.today || days[days.length - 1];
    reqs.forEach(function (q) { q.days = dayNum(today) - dayNum(q.day); });
    var bucketOf = function (q) { for (var i = 0; i < BUCKETS.length; i++) if (q.days >= BUCKETS[i][1] && q.days <= BUCKETS[i][2]) return BUCKETS[i][0]; return BUCKETS[0][0]; };

    var zip;
    try { zip = await JSZip.loadAsync(opts.input); } catch (e) { throw new Error("Файл .xlsx биш байна. Excel дээр нээгээд .xlsx хэлбэрээр хадгалаад дахин оруулна уу."); }
    var st = addStyles(await zip.file("xl/styles.xml").async("string"));
    var X = st.X;
    var rep = reportSheet(reqs, days, X);
    var sheets = [
      { name: "Sheet1", xml: pivotSheet(reqs, days, X).xml(false) },
      { name: "Sheet2", xml: rep.sheet.xml(true) }
    ];

    var wbx = await zip.file("xl/workbook.xml").async("string");
    var rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    var ct = await zip.file("[Content_Types].xml").async("string");
    var existing = []; wbx.replace(/<sheet [^>]*name="([^"]*)"/g, function (_, n) { existing.push(n.replace(/&amp;/g, "&")); return _; });
    var sid = Math.max.apply(null, (wbx.match(/sheetId="(\d+)"/g) || ['sheetId="0"']).map(function (t) { return +/(\d+)/.exec(t)[1]; }));
    var rid = Math.max.apply(null, (rels.match(/Id="rId(\d+)"/g) || ['Id="rId0"']).map(function (t) { return +/(\d+)/.exec(t)[1]; }));
    var n = 1, tags = [], newPaths = {};
    sheets.forEach(function (s) {
      var name = s.name, k = 2;
      while (existing.indexOf(name) >= 0) name = s.name + " (" + (k++) + ")";
      existing.push(name); s.finalName = name;
      while (zip.file("xl/worksheets/sheet" + n + ".xml")) n++;
      var path = "xl/worksheets/sheet" + n + ".xml";
      zip.file(path, s.xml);
      newPaths[path] = 1;
      sid++; rid++;
      tags.push('<sheet name="' + esc(name) + '" sheetId="' + sid + '" r:id="rId' + rid + '"/>');
      rels = rels.replace("</Relationships>", '<Relationship Id="rId' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/></Relationships>');
      ct = ct.replace("</Types>", '<Override PartName="/' + path + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
      n++;
    });
    // New sheets go first; shift sheet-local defined names by the number of inserted sheets.
    wbx = wbx.replace(/<sheets>/, "<sheets>" + tags.join(""));
    wbx = wbx.replace(/localSheetId="(\d+)"/g, function (_, i) { return 'localSheetId="' + (+i + sheets.length) + '"'; });
    if (/activeTab="\d+"/.test(wbx)) wbx = wbx.replace(/activeTab="\d+"/, 'activeTab="1"'); else wbx = wbx.replace(/<workbookView /, '<workbookView activeTab="1" ');
    wbx = wbx.replace(/firstSheet="\d+"/, 'firstSheet="0"');
    zip.file("xl/workbook.xml", wbx);
    zip.file("xl/_rels/workbook.xml.rels", rels);
    zip.file("[Content_Types].xml", ct);
    zip.file("xl/styles.xml", st.styles);
    var others = zip.file(/^xl\/worksheets\/sheet\d+\.xml$/);
    for (var i = 0; i < others.length; i++) {
      if (newPaths[others[i].name]) continue;
      var txt = await others[i].async("string");
      if (txt.indexOf(' tabSelected="1"') >= 0) zip.file(others[i].name, txt.replace(' tabSelected="1"', ""));
    }
    var bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return {
      bytes: bytes, reqs: reqs, days: days, today: today, skipped: read.skipped, table: rep.table,
      sheetNames: sheets.map(function (s) { return s.finalName; }),
      waiting: counter(reqs, bucketOf), bucketOf: bucketOf,
      techCount: uniq(reqs.map(function (q) { return q.tech; })).length, dups: duplicates(reqs)
    };
  }

  async function verifyZasvar(original, res) {
    var problems = [], N = res.reqs.length;
    var za = await JSZip.loadAsync(original), zb = await JSZip.loadAsync(res.bytes);
    var allowed = { "[Content_Types].xml": 1, "xl/workbook.xml": 1, "xl/_rels/workbook.xml.rels": 1, "xl/styles.xml": 1 };
    var names = Object.keys(za.files).filter(function (k) { return !za.files[k].dir; });
    for (var i = 0; i < names.length; i++) {
      var f = names[i];
      if (!zb.file(f)) { problems.push("Эх файлын хэсэг алга: " + f); continue; }
      var a = await za.file(f).async("string"), b = await zb.file(f).async("string");
      if (a === b || allowed[f]) continue;
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(f) && a.replace(' tabSelected="1"', "") === b) continue;
      problems.push("Эх өгөгдөл өөрчлөгдсөн: " + f);
    }
    var wo = XLSX.read(original, { type: "array" }), wn = XLSX.read(res.bytes, { type: "array" });
    wo.SheetNames.forEach(function (s) {
      var A = wo.Sheets[s], B = wn.Sheets[s];
      if (!B) { problems.push("Эх sheet алга: " + s); return; }
      Object.keys(A).forEach(function (k) { if (k[0] !== "!" && (A[k].v !== (B[k] || {}).v)) problems.push(s + " " + k + " өөрчлөгдсөн"); });
    });
    var t = wn.Sheets[res.sheetNames[1]];
    if (!t) { problems.push("Тайлан sheet олдсонгүй"); return problems; }
    var v = function (col, row) { var c = t[colLetter(col) + row]; return c ? c.v : null; };
    var T = REP_TOP, L = REP_LEFT, ncol = res.days.length + 2, lastCol = L + ncol - 1, locSum = 0, r = T + 1, total = null;
    if (v(L, T) !== "Засварын төрөл" || v(lastCol, T) !== "Хүсэлт") problems.push("Тайлан: гарчгийн мөр буруу");
    for (; r < T + 1 + res.table.length; r++) {
      var sum = 0;
      for (var col = L + 1; col < lastCol; col++) sum += v(col, r) || 0;
      var entry = res.table[r - T - 1];
      if (v(L, r) !== entry.label) problems.push("Тайлан: " + r + "-р мөрийн нэр буруу");
      if (sum !== v(lastCol, r) || v(lastCol, r) !== entry.list.length) problems.push('Тайлан: "' + entry.label + '" мөрийн нийлбэр буруу');
      if (entry.loc) locSum += v(lastCol, r);
      if (entry.total) total = v(lastCol, r);
    }
    if (total !== N) problems.push("Тайлан: нийт хүсэлт " + total + " ≠ " + N);
    if (locSum !== N) problems.push("Тайлан: байршлуудын нийлбэр " + locSum + " ≠ " + N);
    [0].forEach(function (k) {
      var s = wn.Sheets[res.sheetNames[k]], rr = XLSX.utils.sheet_to_json(s, { header: 1, raw: true, defval: null });
      var lastRow = rr[rr.length - 1];
      if (lastRow[lastRow.length - 1] !== N) problems.push(res.sheetNames[k] + ": нийт " + lastRow[lastRow.length - 1] + " ≠ " + N);
    });
    return problems;
  }

  return { buildZasvar: buildZasvar, verifyZasvar: verifyZasvar, readRequests: readRequests };
});
