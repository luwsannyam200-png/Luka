/* Weekly stock report builder: browser + Node.
 * Port of .claude/skills/weekly/scripts/build_weekly.py and verify_weekly.py.
 * The draft workbook is patched at the XML level so existing sheets, cached
 * formula values and external links stay byte-identical. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("jszip"), require("xlsx"));
  } else {
    root.WeeklyCore = factory(root.JSZip, root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, XLSX) {
  "use strict";

  var SUPPLIERS = ["AODE", "Очлуур од", "Parts and oil"];
  var REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  function norm(s) {
    return String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase();
  }

  function num(x) {
    var r = Math.round(Number(x) * 1e6) / 1e6;
    return String(r);
  }

  function escRe(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function colLetter(n) {
    var s = "";
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function cellVal(ws, ref) {
    var c = ws[ref];
    return c ? c.v : undefined;
  }

  /* Warehouse → УХ / ЦХ. TL* and the Cyrillic "ТЛ ши" are УХ; TKH* and
   * TG_TKH-Sub are ЦХ. */
  function whLoc(sub) {
    var u = String(sub == null ? "" : sub).trim().toUpperCase();
    if (/(^|[_\s-])TKH/.test(u)) return "TKH";
    if (/^(TL|ТЛ)/.test(u)) return "TL";
    return "";
  }

  /* Oracle stock: sum Quantity per item code for TL* (УХ) and TKH* (ЦХ).
   * The report may sit on any sheet (a TL_TKH summary sheet is often added in
   * front of it). Without the detailed report, a summary sheet with
   * Item Code / TL / TKH columns is read as is. */
  function readOracleStock(buf) {
    var wb = XLSX.read(buf, { type: "array" });
    var books = wb.SheetNames.map(function (n) { return XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }); });
    var found = null, summary = null;
    books.forEach(function (rows) {
      for (var i = 0; i < Math.min(rows.length, 30) && !found; i++) {
        var names = rows[i].map(norm);
        if (names.indexOf("item code") < 0) continue;
        if (names.indexOf("subinventory") >= 0 && names.indexOf("quantity") >= 0) {
          found = { rows: rows, start: i + 1, ix: { code: names.indexOf("item code"), sub: names.indexOf("subinventory"), q: names.indexOf("quantity") } };
        } else if (!summary && names.indexOf("tl") >= 0 && names.indexOf("tkh") >= 0) {
          summary = { rows: rows, start: i + 1, ix: { code: names.indexOf("item code"), tl: names.indexOf("tl"), tkh: names.indexOf("tkh") } };
        }
      }
    });
    var tl = {}, tkh = {}, n = 0, j, r, code;
    if (found) {
      for (j = found.start; j < found.rows.length; j++) {
        r = found.rows[j]; code = r[found.ix.code]; var q = r[found.ix.q];
        if (code == null || typeof q !== "number") continue;
        code = String(code).trim();
        var sub = String(r[found.ix.sub] || "").trim().toUpperCase();
        var w = whLoc(sub);
        if (w === "TKH") { tkh[code] = (tkh[code] || 0) + q; n++; }
        else if (w === "TL") { tl[code] = (tl[code] || 0) + q; n++; }
      }
      return { tl: tl, tkh: tkh, rows: n, source: "detail" };
    }
    if (summary) {
      for (j = summary.start; j < summary.rows.length; j++) {
        r = summary.rows[j]; code = r[summary.ix.code];
        if (code == null || /total|нийт|дүн/i.test(String(code)) || !String(code).trim()) continue; // skip totals
        code = String(code).trim();
        var a = r[summary.ix.tl], b = r[summary.ix.tkh];
        if (typeof a === "number") tl[code] = (tl[code] || 0) + a;
        if (typeof b === "number") tkh[code] = (tkh[code] || 0) + b;
        n++;
      }
      return { tl: tl, tkh: tkh, rows: n, source: "summary" };
    }
    throw new Error("Oracle үлдэгдлийн файлд Item Code / Subinventory / Quantity (эсвэл Item Code / TL / TKH) гарчиг олдсонгүй");
  }

  /* Supplier files differ in layout: take the 13-digit item code in each row
   * and the last numeric cell after it as the quantity. */
  function readSupplier(buf) {
    var wb = XLSX.read(buf, { type: "array" });
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
    var out = {}, n = 0;
    rows.forEach(function (r) {
      var ci = -1;
      for (var i = 0; i < r.length; i++) {
        if (r[i] != null && /^\d{13}$/.test(String(r[i]).trim())) { ci = i; break; }
      }
      if (ci < 0) return;
      var q = null;
      for (var k = ci + 1; k < r.length; k++) {
        var c = r[k];
        if (typeof c === "number") q = c;
        else if (typeof c === "string" && /^[\d,]+(\.\d*)?$/.test(c.trim())) q = parseFloat(c.trim().replace(/,/g, ""));
      }
      if (q != null) {
        var code = String(r[ci]).trim();
        out[code] = (out[code] || 0) + q;
        n++;
      }
    });
    return { values: out, rows: n };
  }

  async function buildWeekly(opts) {
    var stock = readOracleStock(opts.stock);
    var supplier = {};
    SUPPLIERS.forEach(function (label) {
      if (opts.suppliers && opts.suppliers[label]) supplier[label] = readSupplier(opts.suppliers[label]).values;
    });

    var zip = await JSZip.loadAsync(opts.draft);
    var wbx = await zip.file("xl/workbook.xml").async("string");
    var rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    var sheets = [];
    wbx.replace(/<sheet [^>]*?>/g, function (tag) {
      var name = /name="([^"]*)"/.exec(tag), rid = /r:id="(rId\d+)"/.exec(tag);
      if (name && rid) sheets.push({ name: name[1].replace(/&amp;/g, "&"), rid: rid[1] });
      return tag;
    });
    if (sheets.some(function (s) { return s.name === opts.date; })) {
      throw new Error('"' + opts.date + '" нэртэй sheet аль хэдийн байна. Энэ sheet-гүй Weekly draft-ийг оруулна уу.');
    }
    var tpl = sheets[sheets.length - 1];
    var relTag = new RegExp('<Relationship [^>]*Id="' + tpl.rid + '"[^>]*>').exec(rels)[0];
    var tplPath = "xl/" + /Target="([^"]+)"/.exec(relTag)[1].replace(/^\/?(xl\/)?/, "");

    var wb = XLSX.read(opts.draft, { type: "array", sheets: [tpl.name] });
    var ws = wb.Sheets[tpl.name];
    var range = XLSX.utils.decode_range(ws["!ref"]);
    var col = {};
    for (var c = 1; c <= range.e.c + 1; c++) {
      [1, 2].forEach(function (r) {
        var h = norm(cellVal(ws, colLetter(c) + r));
        if (h && !(h in col)) col[h] = c;
      });
    }
    function need(name) {
      var k = norm(name);
      if (!(k in col)) throw new Error('Загвар sheet "' + tpl.name + '"-д "' + name + '" багана олдсонгүй');
      return colLetter(col[k]);
    }
    var C = {
      code: need("item code"), zbn: need("ЗБН"), ux: need("УХ"), cx: need("ЦХ"),
      tlw: need("TL Warehouse"), total: need("Нийт агуулах"), all: need("All warehouse"),
      name: need("Сэлбэгийн мэдээлэл")
    };
    var SUPCOL = {};
    SUPPLIERS.forEach(function (l) { SUPCOL[l] = need(l); });
    var lastSup = colLetter(Math.max.apply(null, SUPPLIERS.map(function (l) { return col[norm(l)]; })));
    var firstRow = 3, lastRow = firstRow;
    for (var rr = firstRow; rr <= range.e.r + 1; rr++) if (cellVal(ws, C.code + rr) != null) lastRow = rr;

    var sheet = await zip.file(tplPath).async("string");
    var styles = await zip.file("xl/styles.xml").async("string");
    var fills = /<fills count="(\d+)">([\s\S]*?)<\/fills>/.exec(styles);
    var nfill = parseInt(fills[1], 10);
    styles = styles.replace(fills[0], '<fills count="' + (nfill + 1) + '">' + fills[2] +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>');
    var cx = /<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/.exec(styles);
    var xfs = cx[2].match(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g);
    if (xfs.length !== parseInt(cx[1], 10)) throw new Error("styles.xml бүтэц таарахгүй байна");
    var yellow = {};
    function ystyle(s) {
      if (!(s in yellow)) {
        var x = xfs[parseInt(s, 10)].replace(/ fillId="\d+"/, ' fillId="' + nfill + '"').replace(/ applyFill="\d"/, "").replace("<xf ", '<xf applyFill="1" ');
        yellow[s] = String(xfs.length);
        xfs.push(x);
      }
      return yellow[s];
    }
    function setcell(ref, inner, fillYellow) {
      var re = new RegExp('<c r="' + ref + '"( [^>]*?)?(/>|>.*?</c>)');
      var m = re.exec(sheet);
      if (!m) throw new Error("Загвар sheet-д " + ref + " нүд алга");
      var s = / s="(\d+)"/.exec(m[1] || "");
      var sattr = s ? ' s="' + (fillYellow ? ystyle(s[1]) : s[1]) + '"' : "";
      var neu = inner ? '<c r="' + ref + '"' + sattr + ">" + inner + "</c>" : '<c r="' + ref + '"' + sattr + "/>";
      sheet = sheet.slice(0, m.index) + neu + sheet.slice(m.index + m[0].length);
    }

    var report = [];
    for (var r = firstRow; r <= lastRow; r++) {
      var code = String(cellVal(ws, C.code + r)).trim();
      var zbn = Number(cellVal(ws, C.zbn + r)) || 0;
      var ux = stock.tl[code] || 0, cxv = stock.tkh[code] || 0;
      setcell(C.ux + r, "<v>" + num(ux) + "</v>");
      setcell(C.cx + r, "<v>" + num(cxv) + "</v>");
      var total = ux + cxv;
      var row = { row: r, name: cellVal(ws, C.name + r), code: code, zbn: zbn, "УХ": ux, "ЦХ": cxv, prev: {} };
      row.prev["УХ"] = cellVal(ws, C.ux + r);
      row.prev["ЦХ"] = cellVal(ws, C.cx + r);
      SUPPLIERS.forEach(function (label) {
        var L = SUPCOL[label];
        row.prev[label] = cellVal(ws, L + r);
        if (supplier[label]) {
          var v = supplier[label][code] || 0;
          setcell(L + r, "<v>" + num(v) + "</v>");
          total += v;
          row[label] = v;
        } else {
          setcell(L + r, "", true);
          row[label] = null;
        }
      });
      var tlw = zbn ? (ux + cxv) / zbn : 0;
      setcell(C.tlw + r, "<f>SUM(" + C.ux + r + ":" + C.cx + r + ")/" + C.zbn + r + "</f><v>" + num(tlw) + "</v>");
      setcell(C.total + r, "<f>SUM(" + C.ux + r + ":" + lastSup + r + ")</f><v>" + num(total) + "</v>");
      setcell(C.all + r, "<f>" + C.total + r + "/" + C.zbn + r + "</f><v>" + num(zbn ? total / zbn : 0) + "</v>");
      row.total = total;
      row.all = zbn ? total / zbn : 0;
      report.push(row);
    }

    sheet = sheet.replace(/ xr:uid="\{[^}]*\}"/, "");
    sheet = sheet.replace(/(<pageSetup[^>]*?) r:id="rId\d+"/, "$1");
    sheet = sheet.replace(/<selection [^>]*\/>/, '<selection activeCell="A1" sqref="A1"/>');
    styles = styles.replace(/<cellXfs count="\d+">[\s\S]*?<\/cellXfs>/, function () {
      return '<cellXfs count="' + xfs.length + '">' + xfs.join("") + "</cellXfs>";
    });

    var n = 1;
    while (zip.file("xl/worksheets/sheet" + n + ".xml")) n++;
    var newPath = "xl/worksheets/sheet" + n + ".xml";
    var others = zip.file(/^xl\/worksheets\/sheet\d+\.xml$/);
    for (var i = 0; i < others.length; i++) {
      var txt = await others[i].async("string");
      if (txt.indexOf('<sheetView tabSelected="1" ') >= 0) zip.file(others[i].name, txt.replace('<sheetView tabSelected="1" ', "<sheetView "));
    }
    zip.file(newPath, sheet);
    zip.file("xl/styles.xml", styles);

    var sid = Math.max.apply(null, (wbx.match(/<sheet [^>]*sheetId="(\d+)"/g) || []).map(function (t) { return parseInt(/sheetId="(\d+)"/.exec(t)[1], 10); })) + 1;
    var rid = Math.max.apply(null, (rels.match(/Id="rId(\d+)"/g) || []).map(function (t) { return parseInt(/(\d+)/.exec(t)[1], 10); })) + 1;
    wbx = wbx.replace("</sheets>", '<sheet name="' + opts.date + '" sheetId="' + sid + '" r:id="rId' + rid + '"/></sheets>');
    wbx = wbx.replace(/activeTab="\d+"/, 'activeTab="' + sheets.length + '"');
    zip.file("xl/workbook.xml", wbx);
    zip.file("xl/_rels/workbook.xml.rels", rels.replace("</Relationships>",
      '<Relationship Id="rId' + rid + '" Type="' + REL + '/worksheet" Target="worksheets/sheet' + n + '.xml"/></Relationships>'));
    var ct = await zip.file("[Content_Types].xml").async("string");
    zip.file("[Content_Types].xml", ct.replace("</Types>",
      '<Override PartName="/' + newPath + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'));

    var bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return {
      bytes: bytes, template: tpl.name, date: opts.date, rows: report, stockRows: stock.rows,
      suppliersFilled: SUPPLIERS.filter(function (l) { return supplier[l]; }),
      suppliersBlank: SUPPLIERS.filter(function (l) { return !supplier[l]; })
    };
  }

  /* Independent checks on the result; returns a list of problems (empty = OK). */
  async function verifyWeekly(original, result, date) {
    var problems = [];
    var za = await JSZip.loadAsync(original), zb = await JSZip.loadAsync(result);
    var allowed = { "[Content_Types].xml": 1, "xl/workbook.xml": 1, "xl/_rels/workbook.xml.rels": 1, "xl/styles.xml": 1 };
    var names = Object.keys(za.files).filter(function (k) { return !za.files[k].dir; });
    for (var i = 0; i < names.length; i++) {
      var f = names[i];
      if (!zb.file(f)) { problems.push("Файлын хэсэг алга: " + f); continue; }
      var a = await za.file(f).async("uint8array"), b = await zb.file(f).async("uint8array");
      var same = a.length === b.length && a.every(function (x, k) { return x === b[k]; });
      if (same || allowed[f]) continue;
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(f)) {
        var ta = new TextDecoder().decode(a), tb = new TextDecoder().decode(b);
        if (ta.replace('<sheetView tabSelected="1" ', "<sheetView ") === tb) continue;
      }
      problems.push("Хуучин хэсэг өөрчлөгдсөн: " + f);
    }

    var wo = XLSX.read(original, { type: "array", cellFormula: true });
    var wn = XLSX.read(result, { type: "array", cellFormula: true });
    if (wn.SheetNames.length !== wo.SheetNames.length + 1 || wn.SheetNames[wn.SheetNames.length - 1] !== date) {
      problems.push("Sheet-ийн жагсаалт буруу");
    }
    wo.SheetNames.forEach(function (s) {
      var A = wo.Sheets[s], B = wn.Sheets[s];
      if (!B) { problems.push("Sheet алга: " + s); return; }
      Object.keys(A).forEach(function (k) {
        if (k[0] === "!") return;
        var x = A[k], y = B[k] || {};
        if (x.v !== y.v || (x.f || "") !== (y.f || "")) problems.push(s + " " + k + " өөрчлөгдсөн");
      });
    });

    var t = wn.Sheets[date];
    if (t) {
      var rg = XLSX.utils.decode_range(t["!ref"]), hdr = {};
      for (var c = 1; c <= rg.e.c + 1; c++) {
        [1, 2].forEach(function (r) {
          var cell = t[colLetter(c) + r], h = norm(cell && cell.v);
          if (h && !(h in hdr)) hdr[h] = colLetter(c);
        });
      }
      var v = function (L, r) { var cell = t[L + r]; return cell && typeof cell.v === "number" ? cell.v : 0; };
      for (var r = 3; r <= rg.e.r + 1; r++) {
        if (!t[hdr["item code"] + r]) continue;
        var parts = [hdr["ух"], hdr["цх"], hdr["aode"], hdr["очлуур од"], hdr["parts and oil"]].map(function (L) { return v(L, r); });
        var sum = parts.reduce(function (p, q) { return p + q; }, 0), e = v(hdr["збн"], r);
        var checks = [[hdr["нийт агуулах"], sum, "Нийт агуулах"]];
        if (e) checks.push([hdr["all warehouse"], sum / e, "All warehouse"], [hdr["tl warehouse"], (parts[0] + parts[1]) / e, "TL Warehouse"]);
        checks.forEach(function (ch) {
          var cell = t[ch[0] + r];
          if (!cell || typeof cell.v !== "number" || Math.abs(cell.v - ch[1]) > 1e-4) problems.push(r + "-р мөр: " + ch[2] + " буруу");
        });
      }
    } else {
      problems.push("Шинэ sheet олдсонгүй");
    }
    return problems;
  }

  return { buildWeekly: buildWeekly, verifyWeekly: verifyWeekly, readOracleStock: readOracleStock, readSupplier: readSupplier, SUPPLIERS: SUPPLIERS };
});
