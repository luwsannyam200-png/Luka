/* Offline assistant: answers Mongolian questions from the loaded files.
 * Rule based (no AI, no network). Browser + Node. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("xlsx"));
  } else {
    root.AssistantCore = factory(root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (XLSX) {
  "use strict";

  var LATIN = { A: "А", B: "В", E: "Е", K: "К", M: "М", H: "Н", O: "О", P: "Р", C: "С", T: "Т", X: "Х", Y: "У" };
  var LOCS = { "ухаа худаг": "Ухаа худаг", "ух": "Ухаа худаг", "цагаан хад": "Цагаан хад", "цх": "Цагаан хад", "замд": "Замд" };
  var STOP = ["зөвхөн", "дээрх", "дахь", "хэд", "хэдэн", "байна", "байгаа", "бий", "юу", "уу", "үү", "вэ", "бэ", "нь", "хаана", "үлдэгдэл", "үлдэгдэлтэй",
    "агуулахад", "агуулах", "ширхэг", "ш", "тоо", "хэмжээ", "сэлбэг", "код", "кодтой", "item", "ямар", "олох", "хай", "хайх",
    "харуул", "мэдээлэл", "дээр", "бол", "тэр", "энэ", "за", "надад", "тэгвэл", "одоо", "дахиад", "бас", "харуулаач", "хайгаач", "олоод", "өгөөч", "?", "."];

  function fold(s) {
    return String(s == null ? "" : s).toLowerCase().replace(/ө/g, "о").replace(/ү/g, "у").replace(/ё/g, "е").replace(/\s+/g, " ").trim();
  }
  function plateKey(s) {
    return String(s || "").toUpperCase().replace(/\s+/g, "").replace(/[ABEKMHOPCTXY]/g, function (c) { return LATIN[c]; })
      .replace(/Ө/g, "О").replace(/Ү/g, "У");
  }
  function norm(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase(); }
  function rowsOf(buf) {
    var wb = XLSX.read(buf, { type: "array" });
    return wb.SheetNames.map(function (n) { return XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }); });
  }
  function findHeader(sheets, required) {
    for (var s = 0; s < sheets.length; s++) {
      for (var i = 0; i < Math.min(sheets[s].length, 30); i++) {
        var names = sheets[s][i].map(norm);
        if (required.every(function (r) { return names.some(function (n) { return n === r || n.indexOf(r) === 0; }); })) {
          return { rows: sheets[s], hi: i, names: names };
        }
      }
    }
    return null;
  }
  function col(h, name) {
    for (var i = 0; i < h.names.length; i++) if (h.names[i] === name) return i;
    for (i = 0; i < h.names.length; i++) if (h.names[i].indexOf(name) === 0) return i;
    return -1;
  }
  function fmt(n) { return Number(n).toLocaleString("en-US", { maximumFractionDigits: 2 }); }
  function pad(n) { return String(n).padStart(2, "0"); }
  function dayKey(v) {
    if (typeof v === "number" && v > 1) {
      var d = new Date(Math.round((Math.floor(v) - 25569) * 86400000));
      return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
    }
    var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(v || "").trim());
    return m ? m[1] + "-" + pad(m[2]) + "-" + pad(m[3]) : null;
  }
  function dayNum(k) { var p = k.split("-"); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000; }

  /* ---------- loaders: detect the file by its header ---------- */
  function load(buf, name) {
    var sheets = rowsOf(buf);
    var h = findHeader(sheets, ["item code", "subinventory", "quantity"]);
    if (h) return { kind: "stock", data: loadStock(h), name: name };
    h = findSummaryHeader(sheets);
    if (h) return { kind: "stock", data: loadStock(h), name: name };
    h = findHeader(sheets, ["засварын төрөл", "засварын байршил", "үйлчилгээ", "хүсэлт гаргасан"]);
    if (h) return { kind: "requests", data: loadRequests(h), name: name };
    var kits = [].concat.apply([], sheets.map(parseMaint));
    if (kits.length) return { kind: "maint", data: kits, name: name };
    h = findCatalogHeader(sheets);
    if (h) return { kind: "catalog", data: loadCatalog(h), name: name };
    var sup = loadSupplier(sheets[0]);
    if (Object.keys(sup).length) return { kind: "supplier", data: sup, name: name };
    return { kind: "unknown", name: name };
  }
  /* A TL_TKH summary sheet (Item Code | Item Description | … | TL | TKH) with no
   * Oracle detail: turn each TL / TKH cell into a row loadStock understands. */
  function findSummaryHeader(sheets) {
    for (var s = 0; s < sheets.length; s++) {
      for (var i = 0; i < Math.min(sheets[s].length, 30); i++) {
        var names = sheets[s][i].map(norm), code = names.indexOf("item code"), tl = names.indexOf("tl"), tkh = names.indexOf("tkh");
        if (code < 0 || tl < 0 || tkh < 0) continue;
        var en = names.indexOf("item description"), mn = names.indexOf("item mongolian description"), uom = names.indexOf("uom");
        var rows = [["item code", "item description", "item mongolian description", "uom", "subinventory", "quantity"]];
        sheets[s].slice(i + 1).forEach(function (r) {
          if (r[code] == null || /total|нийт|дүн/i.test(String(r[code]))) return;
          [["TL", tl], ["TKH", tkh]].forEach(function (w) {
            if (typeof r[w[1]] === "number" && r[w[1]]) rows.push([r[code], en >= 0 ? r[en] : "", mn >= 0 ? r[mn] : "", uom >= 0 ? r[uom] : "", w[0], r[w[1]]]);
          });
        });
        return { rows: rows, hi: 0, names: rows[0] };
      }
    }
    return null;
  }
  /* Spare-part list (e.g. "2025 item.xlsx": Item дугаар | Сэлбэгийн нэр, or the
   * older Item.xlsx: Item | Description | Эдийн дугаар | Техникийн төрөл | Бүлэг). */
  function findCatalogHeader(sheets) {
    for (var s = 0; s < sheets.length; s++) {
      for (var i = 0; i < Math.min(sheets[s].length, 10); i++) {
        var names = sheets[s][i].map(norm);
        var code = names.findIndex(function (n) { return n === "item" || n.indexOf("item дугаар") === 0 || n === "item code"; });
        var desc = names.findIndex(function (n) { return n.indexOf("сэлбэгийн нэр") === 0 || n === "description"; });
        if (code >= 0 && desc >= 0) return { rows: sheets[s], hi: i, names: names, code: code, desc: desc };
      }
    }
    return null;
  }
  function loadCatalog(h) {
    var extra = { type: col(h, "техникийн төрөл"), group: col(h, "бүлэг"), en: col(h, "english short name") };
    var parts = ["эдийн дугаар", "эдийн дугаар 2", "эдийн дугаар 3"].map(function (n) { return col(h, n); }).filter(function (i) { return i >= 0; });
    var out = {};
    h.rows.slice(h.hi + 1).forEach(function (r) {
      var code = r[h.code];
      if (code == null || !/^\d{6,}$/.test(String(code).trim())) return;
      code = String(code).trim();
      var desc = String(r[h.desc] || "").trim();
      var segs = desc.split("|").map(function (x) { return x.trim(); }).filter(Boolean);
      var mn = segs.filter(function (x) { return /[А-Яа-яӨөҮү]/.test(x); }).pop() || "";
      var pn = parts.map(function (i) { return r[i]; }).filter(Boolean).join(" ");
      out[code] = { code: code, desc: desc, en: (extra.en >= 0 && r[extra.en]) || segs[0] || "", mn: mn,
        type: extra.type >= 0 ? r[extra.type] || "" : "", group: extra.group >= 0 ? r[extra.group] || "" : "",
        search: fold(code + " " + desc + " " + pn + " " + (extra.en >= 0 ? r[extra.en] || "" : "")) };
    });
    return out;
  }
  function loadStock(h) {
    var c = { code: col(h, "item code"), en: col(h, "item description"), mn: col(h, "item mongolian description"),
      sub: col(h, "subinventory"), loc: col(h, "locator"), q: col(h, "quantity"), uom: col(h, "uom"),
      supplier: col(h, "supplier"), date: col(h, "origination date") };
    var items = {}, raw = [];
    h.rows.slice(h.hi + 1).forEach(function (r) {
      var code = r[c.code], q = r[c.q];
      if (code == null || typeof q !== "number") return;
      var sub = String(r[c.sub] || "").trim(), u = sub.toUpperCase();
      if (u.indexOf("TL") !== 0 && u.indexOf("TKH") !== 0) return; // only УХ (TL) and ЦХ (TKH) are used
      code = String(code).trim();
      var it = items[code] || (items[code] = { code: code, en: r[c.en] || "", mn: r[c.mn] || "", uom: r[c.uom] || "", subs: {}, total: 0 });
      var key = sub + (r[c.loc] ? " (" + r[c.loc] + ")" : "");
      it.subs[key] = (it.subs[key] || 0) + q;
      it.total += q;
      raw.push({ code: code, en: String(r[c.en] || ""), mn: String(r[c.mn] || ""), uom: r[c.uom] || "", sub: sub, loc: r[c.loc] || "",
        qty: q, supplier: c.supplier >= 0 ? r[c.supplier] || "" : "", date: c.date >= 0 ? dayKey(r[c.date]) || "" : "" });
    });
    Object.defineProperty(items, "_rows", { value: raw, enumerable: false });
    Object.keys(items).forEach(function (k) {
      var it = items[k];
      it.tl = 0; it.tkh = 0;
      Object.keys(it.subs).forEach(function (s) {
        var u = s.toUpperCase();
        if (u.indexOf("TKH") === 0) it.tkh += it.subs[s]; else if (u.indexOf("TL") === 0) it.tl += it.subs[s];
      });
      it.search = fold(it.code + " " + it.en + " " + it.mn);
    });
    return items;
  }
  function loadSupplier(rows) {
    var out = {};
    (rows || []).forEach(function (r) {
      var ci = -1;
      for (var i = 0; i < r.length; i++) if (r[i] != null && /^\d{13}$/.test(String(r[i]).trim())) { ci = i; break; }
      if (ci < 0) return;
      var q = null;
      for (var k = ci + 1; k < r.length; k++) {
        if (typeof r[k] === "number") q = r[k];
        else if (typeof r[k] === "string" && /^[\d,]+(\.\d*)?$/.test(r[k].trim())) q = parseFloat(r[k].trim().replace(/,/g, ""));
      }
      if (q != null) { var code = String(r[ci]).trim(); out[code] = (out[code] || 0) + q; }
    });
    return out;
  }
  function loadRequests(h) {
    var c = { type: col(h, "засварын төрөл"), tech: col(h, "техни"), loc: col(h, "засварын байршил"), srv: col(h, "үйлчилгээ"),
      desc: col(h, "тайлбар"), req: col(h, "хүсэлт гаргасан") };
    var out = [];
    h.rows.slice(h.hi + 1).forEach(function (r) {
      if (r[c.tech] == null || r[c.type] == null) return;
      var day = dayKey(r[c.req]);
      if (!day) return;
      out.push({ type: String(r[c.type]).trim(), tech: String(r[c.tech]).trim(), loc: String(r[c.loc] || "").trim() || "(хоосон)",
        srv: String(r[c.srv] || "").trim() || "(хоосон)", desc: String(r[c.desc] || "").trim(), day: day });
    });
    var last = out.reduce(function (m, q) { return q.day > m ? q.day : m; }, "");
    out.forEach(function (q) { q.days = dayNum(last) - dayNum(q.day); });
    return { list: out, today: last };
  }

  /* ---------- answering ---------- */
  function table(head, rows, limit) {
    limit = limit || 20;
    return { head: head, rows: rows.slice(0, limit), more: Math.max(0, rows.length - limit), all: rows };
  }
  /* ---------- ТҮ (техник үйлчилгээ) сэлбэгийн хүснэгт ----------
   * Built in from the user's sheet; an Excel file in the same layout adds or
   * replaces machines. Layout: a row with ТҮ-1 (10000) … headers from column D,
   * Activity row(s), a "№ | Item code | …" row, then one row per part. */
  var MAINT_TSV = [
    "Howo 371 Засвар\t\t\tTY-1 (10000)\tTY-2 (30000)\tTY-3, УЗ-1 (50000)\tТҮ-3, УЗ-2 (100000)\tТҮ-3, ИЗ (200000)",
    "Activity\t\t\tTV001-MAI-004\tTV001-MAI-006\tTV001-MAI-005\tTV001-MAI-005\tTV001-MAI-005",
    "№\tItem code\t\t\t\tTV001-MAI-021\tTV001-MAI-022\tTV001-MAI-023",
    "1\t1512150100049\toil\t24\t24\t24\t24\t24",
    "2\t4016150400048\toil filter Маслын шүүр\t2\t2\t2\t2\t2",
    "3\t4016150500154\tair filter\t1\t1\t1\t1\t1",
    "4\t4016151300136\tfuel filter Түлшний тунгаагуур\t1\t1\t1\t1\t1",
    "5\t4016151300043\tfuel filter 2 Түлшний шүүр жижиг\t1\t1\t1\t1\t1",
    "6\t1512150300019\tgear oil\t0\t18\t36,55\t36,55\t36,55",
    "",
    "HOWO T7H Төлөвлөгөөт засвар",
    "Засвар\t\t\tTY-1 (10000)\tTY-2 (30000)\tTY-3, УЗ-1 (50000)\tТҮ-3, УЗ-2 (100000)\tТҮ-3, ИЗ (200000)",
    "Activity\t\t\tTV001-MAI-030\tTV001-MAI-031\tTV001-MAI-032\tTV001-MAI-032\tTV001-MAI-032",
    "№\tItem code\tDiscription\t\t\tTV001-MAI-021\tTV001-MAI-022\tTV001-MAI-023",
    "1\t4016151300197\tТүлшний тунгаагуур /цахилгаан халаалтын/\t1\t1\t1\t1\t1",
    "2\t4016151300198\tТүлшний тунгаагуур /энгийн/\t1\t1\t1\t1\t1",
    "3\t4016151300199\tТүлшний тунгаагуурын шүүр\t1\t1\t1\t1\t1",
    "4\t4016151300200\tТүлшний шүүр /цаасан/\t1\t1\t1\t1\t1",
    "5\t4016150500154\tАгаар шүүгч\t1\t1\t1\t1\t1",
    "6\t4016150400213\tТосны шүүр\t1\t1\t1\t1\t1",
    "7\t1512150100093\tХөдөлгүүрийн тос\t42\t42\t42\t42\t42",
    "8\t4016150500190\tЭйр кондейшны шүүр\t1\t1\t1\t1\t1",
    "9\t1512150300019\tХүч дамжуулах ангийн тос\t0\t18\t54\t54\t54"
  ].map(function (l) { return l.split("\t"); });
  var LEVEL = /^\s*[tт][yуү]\s*-?\s*(\d)/i;
  function maintQty(v) {
    if (typeof v === "number") return v;
    var s = String(v == null ? "" : v).trim();
    if (/^\d+,\d{1,2}$/.test(s)) return parseFloat(s.replace(",", "."));   // 36,55 = 36.55 (decimal comma)
    s = s.replace(/,/g, "");
    return /^\d+(\.\d+)?$/.test(s) ? parseFloat(s) : 0;
  }
  function parseMaint(rows) {
    var kits = [], txt = function (v) { return String(v == null ? "" : v).trim(); };
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || [], cols = [];
      for (var c = 1; c < r.length; c++) if (LEVEL.test(txt(r[c]))) cols.push(c);
      if (cols.length < 2) continue;
      var clean = function (s) { return txt(s).replace(/төлөвлөгөөт|засвар/gi, "").replace(/\s+/g, " ").trim(); };
      var name = clean(r[0]);
      for (var p = i - 1; !name && p >= 0 && p >= i - 3; p--) name = clean((rows[p] || [])[0]);
      var kit = { name: name || "Машин " + (kits.length + 1), levels: cols.map(function (c) {
        return { label: txt(r[c]).replace(/^\s*[tт][yуү]/i, "ТҮ").replace(/,\s*[yу]з/i, ", УЗ"), n: +LEVEL.exec(txt(r[c]))[1], activity: [] };
      }), parts: [] };
      for (var j = i + 1; j < rows.length; j++) {
        var row = rows[j] || [], code = txt(row[1]);
        if (/^\d{10,13}$/.test(code)) {
          kit.parts.push({ code: code, name: txt(row[2]), qty: cols.map(function (c) { return maintQty(row[c]); }) });
        } else if (kit.parts.length && !row.some(function (x) { return txt(x); })) {
          break;                                                           // blank line ends the table
        } else if (kit.parts.length && cols.filter(function (c) { return LEVEL.test(txt(row[c])); }).length >= 2) {
          break;                                                           // next machine
        } else if (!kit.parts.length) {
          cols.forEach(function (c, k) { var v = txt(row[c]); if (v) kit.levels[k].activity.push(v); });
        }
      }
      if (kit.parts.length) kits.push(kit);
      i = j - 1;
    }
    return kits;
  }
  var MAINT = parseMaint(MAINT_TSV);
  function maintKey(s) {
    // Latin and Cyrillic look-alikes: "т7н", "Т7Х" = "t7h"; "хово" = "howo"
    return fold(s).replace(/хово/g, "howo").replace(/т/g, "t").replace(/[нх]/g, "h").replace(/[оo]/g, "o");
  }
  function maintKits(kb) {
    var byName = {};
    MAINT.concat(kb.maint || []).forEach(function (k) { byName[maintKey(k.name)] = k; });
    return Object.keys(byName).map(function (k) { return byName[k]; });
  }
  function maintAnswer(question, q, kb) {
    var kits = maintKits(kb), mq = " " + maintKey(question).replace(/[?!.,()]/g, " ") + " ";
    var picked = kits.filter(function (k) {
      return maintKey(k.name).split(" ").filter(function (t) { return t !== "howo" && t.length > 1; }).some(function (t) { return mq.indexOf(" " + t + " ") >= 0; });
    });
    var lv = /(?:^|[\s(])(?:ту|ty|tu)\s*-?\s*(\d)/.exec(q.replace(/ү/g, "у"));
    var aboutKit = /(ту|ty|tu)[\s-]?(\d|н|ний|ны|ийн)|техник(ийн)? уйлчилгээ/.test(q) && /(сэлбэг|материал|юу|ямар|хэрэг|жагсаалт|орох|kit|тос|шуур)/.test(q);
    if (!picked.length && !aboutKit) return null;
    if (!picked.length) {
      if (/howo|хово/.test(q)) picked = kits.filter(function (k) { return /howo/.test(maintKey(k.name)); });
      if (!picked.length) picked = kits;
    }
    var stock = kb.stock, cat = kb.catalog;
    var sub = /(уз|yz)\s*-?\s*(\d)/.exec(q) || (/(^|\s)(из|их засвар)(\s|$)/.test(q) ? [0, "из"] : null);
    var km = /(\d{2,3})\s?000/.exec(q.replace(/(\d)\s(\d{3})/g, "$1$2"));
    function pickLevels(k) {
      return k.levels.map(function (l, i) { return i; }).filter(function (i) {
        var l = k.levels[i], lab = fold(l.label).replace(/ty/g, "ту");
        if (km) return lab.indexOf("(" + km[1] + "000)") >= 0;
        if (lv && l.n !== +lv[1]) return false;
        if (sub && sub[0] && lab.replace(/\s/g, "").indexOf("уз-" + sub[2]) < 0 && lab.replace(/\s/g, "").indexOf("уз" + sub[2]) < 0) return false;
        if (sub && !sub[0] && lab.indexOf("из") < 0) return false;
        return true;
      });
    }
    function one(k) {
      var idx = pickLevels(k);
      if (!idx.length) idx = k.levels.map(function (l, i) { return i; });
      var parts = k.parts.filter(function (p) { return idx.some(function (i) { return p.qty[i] > 0; }); });
      var short = { TL: [], TKH: [] };
      var rows = parts.map(function (p) {
        var it = stock && stock[p.code];
        var row = [p.code, p.name || (it ? it.mn || it.en : "")].concat(idx.map(function (i) { return p.qty[i]; }));
        if (stock) {
          var need = Math.max.apply(null, idx.map(function (i) { return p.qty[i]; }));
          var tl = it ? it.tl : 0, tkh = it ? it.tkh : 0;
          if (tl < need) short.TL.push(p.name || p.code);
          if (tkh < need) short.TKH.push(p.name || p.code);
          row = row.concat([fmt(tl), fmt(tkh), (tl >= need ? "✅" : "❌") + " / " + (tkh >= need ? "✅" : "❌")]);
        }
        if (cat) row.push(cat[p.code] ? "⭐" : "");
        return row;
      });
      var head = ["Item code", "Сэлбэг"].concat(idx.map(function (i) { return k.levels[i].label; }))
        .concat(stock ? ["TL (УХ)", "TKH (ЦХ)", "Хүрэлцэх УХ / ЦХ"] : []).concat(cat ? ["Гэрээт"] : []);
      var acts = idx.map(function (i) { return k.levels[i].label + ": " + (k.levels[i].activity.join(", ") || "—"); });
      var txt = "**" + k.name + "**, техник үйлчилгээний сэлбэг (" + parts.length + ")\nActivity: " + acts.join("; ");
      if (stock) {
        var sh = function (w, nm) { return short[w].length ? nm + "-д **хүрэлцэхгүй " + short[w].length + "**: " + short[w].join(", ") : nm + "-д бүгд хүрэлцэнэ ✅"; };
        txt += "\n" + sh("TL", "УХ") + "\n" + sh("TKH", "ЦХ");
      } else txt += "\nOracle-ийн үлдэгдэл оруулбал УХ, ЦХ-д хүрэлцэх эсэхийг харуулна.";
      return { text: txt, table: table(head, rows, 100), exportName: k.name + "_TU" };
    }
    if (picked.length === 1) {
      var a = one(picked[0]);
      a.examples = picked[0].levels.map(function (l) { return picked[0].name + " " + l.label.replace(/\s*\(.*$/, "").replace(/,/, ""); }).filter(function (x, i, arr) { return arr.indexOf(x) === i; });
      return a;
    }
    // several machines: first as the table, the rest listed as example chips
    var first = one(picked[0]);
    first.text = "ТҮ-ний сэлбэгийн хүснэгт **" + picked.length + "** машинд байна: " + picked.map(function (k) { return "**" + k.name + "**"; }).join(", ") + ".\nДоор **" + picked[0].name + "**-ийнх. Бусдыг нь доорх товчоор сонгоно уу.\n\n" + first.text;
    var second = one(picked[1]);
    first.detailTitle = picked[1].name + ":";
    first.detail = second.table;
    first.examples = picked.map(function (k) { return k.name + " ТҮ-1"; }).concat(picked.map(function (k) { return k.name + " ТҮ-2"; }));
    return first;
  }

  function need(kind) {
    var what = { stock: "Oracle-ийн үлдэгдлийн тайлан (жишээ нь 09-26 үлдэгдэл.xlsx)", requests: "Mine2TL-ийн засварын хүсэлтийн файл" }[kind];
    return { text: "Энэ асуултад хариулахын тулд **" + what + "** хэрэгтэй. Дээрх \"Файл нэмэх\" хэсэгт оруулна уу." };
  }
  var HELP = {
    text: "Би оруулсан файлуудаас хайж хариулна. Жишээ асуултууд:",
    examples: ["агаар шүүгч", "зөвхөн ЦХ", "ТҮ-ний сэлбэг", "Howo 371 ТҮ-2", "турбин гэж юу вэ", "маслын шүүр хэд байна", "4016150400048 хаана байна", "VG1540080311", "5840ӨМА", "хүсэлт хэд байна", "цагаан хадын гагнуур хэд", "8 хоногоос удсан хүсэлт", "гэрээт сэлбэгээс нийлүүлэгчид байхгүй"]
  };

  /* ---------- conversation: small talk, knowledge, follow-ups ---------- */
  var GLOSSARY = [
    [["ту", "техник уйлчилгээ", "техникийн уйлчилгээ"], "**ТҮ (Техник үйлчилгээ)** нь машиныг эвдрэхээс нь өмнө тогтмол давтамжтай хийдэг урьдчилан сэргийлэх үйлчилгээ. Тос, шүүр солих, тослох, тохируулах зэрэг ажил багтана. **ТҮ-1, ТҮ-2, ТҮ-3** нь явсан км эсвэл мото цагаас хамаарсан шатууд бөгөөд дугаар ихсэх тусам ажлын хүрээ өргөн болно. Давтамжийг компанийн засвар үйлчилгээний журмаар тогтоодог."],
    [["уз"], "**УЗ** нь ТҮ-тэй хамт хийгддэг нэмэлт ажлын шатыг заадаг. Жишээ нь ТҮ-3 УЗ-1. Нарийн утгыг танай засварын журмаас шалгана уу. Та туслахад `заа: УЗ гэж юу вэ = …` гэж зааж өгч болно."],
    [["wo", "work order", "ажлын захиалга"], "**WO (Work Order, ажлын захиалга)** нь ERP (Oracle)-д засварын ажлыг бүртгэдэг баримт. Ямар машинд, ямар ажил хийх, ямар сэлбэг, хэдэн цагийн хөдөлмөр орохыг бүртгэдэг. Сэлбэг агуулахаас WO-оор гардаг. ER2286811 гэх мэт дугаартай."],
    [["erp", "oracle", "ebs"], "**ERP (Oracle E-Business Suite)** бол компанийн нэгдсэн удирдлагын систем. Санхүү, агуулах, худалдан авалт, засвар үйлчилгээ (eAM) бүгд нэг дор бүртгэгддэг. Сэлбэгийн үлдэгдлийн тайлан, WO хоёулаа эндээс гардаг."],
    [["eam"], "**eAM (Enterprise Asset Management)** бол Oracle-ийн тоног төхөөрөмжийн засвар үйлчилгээний модуль. WO, засварын төлөвлөгөө энд бүртгэгддэг."],
    [["mine2tl"], "**Mine2TL** бол уурхайн тээврийн хэрэгслийн засварын хүсэлт, засварын мэдээ, жолооч, рейсийг бүртгэдэг програм. Засварын хүсэлт, засварын мэдээний Excel файлууд эндээс гардаг."],
    [["subinventory", "агуулахын бус"], "**Subinventory** гэдэг нь Oracle-ийн агуулахын **бүс**. Жишээ нь TL_ZONE_A, TKH_LD. Манай тайланд **TL…** = Ухаа худаг (УХ), **TKH…** = Цагаан хад (ЦХ)."],
    [["locator", "тавиур"], "**Locator** гэдэг нь агуулахын бүс доторх **тавиурын байршил**. Жишээ нь TLA10-1. Сэлбэгийг яг хаанаас олохыг заана."],
    [["uom", "хэмжих нэгж"], "**UOM** гэдэг нь хэмжих нэгж: **EA** = ширхэг, **L** = литр, **KG** = килограмм, **M** = метр, **PR** = хос, **ST** = иж бүрдэл."],
    [["tl"], "**TL** гэж эхэлсэн агуулахууд (TL_ZONE_A, TL_LD…) нь **Ухаа худаг (УХ)**-ийн агуулах."],
    [["tkh"], "**TKH** гэж эхэлсэн агуулахууд (TKH_ZONE_A, TKH_LD…) нь **Цагаан хад (ЦХ)**-ын агуулах."],
    [["збн"], "**ЗБН** бол Weekly тайланд сэлбэг бүрийн **байх ёстой нөөцийн хэмжээ**. **All warehouse** = нийт үлдэгдэл ÷ ЗБН. 1-ээс бага бол нөөц хүрэлцэхгүй байна гэсэн үг."],
    [["all warehouse"], "**All warehouse** = (УХ + ЦХ + гэрээт агуулахуудын үлдэгдэл) ÷ ЗБН. 0.5-аас бага бол байх ёстой нөөцийн талаас ч хүрэхгүй байна."],
    [["гэрээт агуулах", "гэрээт"], "**Гэрээт агуулах** гэдэг нь гэрээт нийлүүлэгчдийн агуулах: **AODE**, **Очлуур од**, **Parts and oil**. 2025 item жагсаалтын сэлбэгүүдийг эдгээрээс авч ашигладаг."],
    [["дуудлага"], "**Дуудлага** гэдэг нь зам дээр эвдэрсэн машинд засварчин очиж хийдэг засвар. Засварын мэдээнд механик нь аль байршлынх вэ, түүгээр УХ эсвэл ЦХ-д тооцдог."],
    [["төлөвлөгөөт бус"], "**Төлөвлөгөөт бус засвар** гэдэг нь урьдчилан төлөвлөөгүй, гэнэт гарсан эвдрэлийн засвар."],
    [["төлөвлөгөөт"], "**Төлөвлөгөөт засвар** гэдэг нь урьдчилан төлөвлөсөн засвар үйлчилгээ. Ихэвчлэн ТҮ-1, ТҮ-2, ТҮ-3."],
    [["хүлээсэн цаг"], "**Хүлээсэн цаг** = засвар эхэлсэн − хүсэлт өгсөн. Машин засвар эхлэхийг хэр удаан хүлээснийг заана."],
    [["зарцуулсан цаг"], "**Зарцуулсан цаг** = засвар дууссан − засвар эхэлсэн. Засварт хэр удсаныг заана."],
    [["турбин", "турбо", "turbocharger"], "**Турбин (турбо хөөрөгч)** нь яндангийн хийн энергиэр эргэж, хөдөлгүүр рүү илүү их агаар шахдаг. Ингэснээр хүчин чадлыг нэмэгдүүлдэг. Тос гоожих, исгэрэх, хүч суларвал шалгах хэрэгтэй."],
    [["форсунк", "injector"], "**Форсунк (injector)** нь түлшийг хөдөлгүүрийн цилиндрт нарийн шүршиж өгдөг хэсэг. Муудвал утаа ихэсч, хүч сулардаг."],
    [["сальник", "oil seal"], "**Сальник (oil seal)** нь эргэлддэг голын дагуу тос гоожихоос хамгаалдаг резинэн битүүмжлэгч."],
    [["жийрэг", "gasket", "прокладк"], "**Жийрэг (gasket)** нь хоёр эд ангийн холбоосын завсрыг битүүмжилдэг хавтгай дэвсгэр. Тос, ус, хий алдагдахаас хамгаална."],
    [["холхивч", "подшипник", "bearing", "шарик"], "**Холхивч (bearing, шарик)** нь эргэлддэг голыг тулж, үрэлтийг багасгадаг. Ступиц, кардан, редукторт их хэрэглэгдэнэ."],
    [["кардан", "propeller shaft"], "**Кардан (propeller shaft)** нь хурдны хайрцгаас хүчийг ар тэнхлэг (редуктор) рүү дамжуулдаг гол."],
    [["редуктор"], "**Редуктор (гол дамжуулга)** нь карданаас ирсэн эргэлтийг дугуй руу дамжуулж, эргэлтийн хурдыг бууруулж хүчийг нэмэгдүүлдэг тэнхлэгийн механизм."],
    [["дифференциал"], "**Дифференциал** нь эргэх үед зүүн, баруун дугуйг өөр хурдтай эргэх боломж олгодог механизм."],
    [["стартер"], "**Стартер** нь хөдөлгүүрийг асаахдаа эхний эргэлтийг өгдөг цахилгаан мотор."],
    [["динам", "генератор", "alternator"], "**Динам (генератор, alternator)** нь хөдөлгүүр ажиллах үед цахилгаан үйлдвэрлэж, аккумляторыг цэнэглэдэг."],
    [["аккумлятор", "батерей"], "**Аккумлятор** нь асаалт болон цахилгаан хэрэглэгчдэд эрчим хүч хадгалж өгдөг. Howo-д ихэвчлэн 24V (2 ширхэг 12V)."],
    [["тунгаагуур"], "**Түлшний тунгаагуур** нь түлшнээс ус, том хольцыг ялгаж авдаг анхан шатны шүүлтүүр. Үндсэн (нарийн) шүүрийн өмнө байрладаг."],
    [["агаар шуугч", "агаар шүүгч"], "**Агаар шүүгч** нь хөдөлгүүрт орох агаараас тоос шороог шүүдэг. Уурхайн тоостой нөхцөлд ойр ойрхон шалгаж солих шаардлагатай."],
    [["маслын шуур", "маслын шүүр", "тосны шүүр"], "**Маслын (тосны) шүүр** нь хөдөлгүүрийн тосноос металлын үртэс, хольцыг шүүдэг. Ихэвчлэн тос солихтой хамт солигдоно."],
    [["мембрам", "мембран", "brake chamber"], "**Мембрам (тоормосны камер)** нь хийн даралтыг механик хүч болгож тоормосыг ажиллуулдаг. Давхар мембрам нь зогсоолын тоормостой хамт байдаг."],
    [["конс", "clutch", "шүүрэлт"], "**Конс (шүүрэлт, clutch)** нь хөдөлгүүр ба хурдны хайрцгийг холбож, салгадаг механизм. Дээд, доод консны аппарат нь түүнийг гидравликаар удирддаг."],
    [["гильз", "cylinder liner"], "**Гильз (cylinder liner)** нь хөдөлгүүрийн цилиндрийн дотор ханын солигддог ган хоолой. Поршен түүний дотор хөдөлдөг."],
    [["поршен", "кольц"], "**Поршен** нь цилиндр дотор дээш доош хөдөлж, шаталтын даралтыг тахир гол руу дамжуулна. **Кольц (поршений цагираг)** нь битүүмжлэлийг хангана."],
    [["втулк", "бушинг", "bushing"], "**Втулк (bushing)** нь гол, пальцны эргэн тойрон дахь элэгддэг хамгаалалтын цагираг."],
    [["пальц", "pin"], "**Пальц (pin)** нь хоёр хэсгийг нугастай холбодог гол. Жишээ нь эмээлийн пальц, дүүжин пальц."],
    [["шпилька", "stud"], "**Шпилька (stud)** нь хоёр үзүүртээ эрэгтэй боолт. Дугуйн шпилька, гайк нь обудыг ступицад бэхэлдэг."],
    [["ступиц", "ступец", "hub"], "**Ступиц (hub)** нь дугуйг тэнхлэгт холбож, холхивчоор эргүүлдэг зангилаа."],
    [["эмээл", "fifth wheel"], "**Эмээл (fifth wheel)** нь чирэгч дээрх чиргүүл холбох механизм. Чиргүүлийн king pin энд түгжигддэг."]
  ];
  var JOKES = [
    "Механик эмчид очоод: \"Эмчээ, миний ажил ч ялгаагүй. Хөдөлгүүрийг ажиллаж байхад нь засдаг\" гэхэд эмч: \"Тэгвэл миний ажлыг хийгээд үз дээ\" гэжээ 😄",
    "Агаар шүүгч хүүхдэдээ: \"Амьдралд тоос шороо их, гэхдээ бүгдийг нь дотроо хадгалах хэрэггүй шүү\" гэж хэлдэг гэнэ 😄",
    "Excel-ийн шүүлтүүр ба би хоёрын ялгаа юу вэ? Би \"баярлалаа\" гэж хэлэхэд чинь хариулдаг 😄"
  ];
  function greet() {
    var h = new Date().getHours();
    return h < 11 ? "Өглөөний мэнд" : h < 18 ? "Өдрийн мэнд" : "Оройн мэнд";
  }
  function words(q) { return q.replace(/[?!.,]/g, " ").split(" ").filter(Boolean); }
  function smallTalk(q, kb) {
    var name = kb.user ? ", " + kb.user : "";
    var w = words(q), short = w.length <= 6 && !/\d{4}/.test(q);
    var m = /^(?:намайг|миний нэр(?:ийг)?)\s+(.+?)\s+(?:гэдэг|гэнэ|гэж дууд\S*)/.exec(q);
    if (m) {
      var nm = m[1].replace(/^./, function (c) { return c.toUpperCase(); });
      kb.user = nm; kb.dirty = true;
      return { text: "Танилцсандаа баяртай байна, **" + nm + "**! Одоо юу хайх вэ?" };
    }
    if (!short) return null;
    if (/^(сайн байна уу|сайн уу|сайн|мэнд|мэндээ|сайхан амарсан уу|hi|hello|hey|сайн байцгаана уу)$/.test(w.join(" ")) || /^(оглооний|одрийн|оройн) мэнд/.test(q))
      return { text: greet() + name + "! 👋 Сэлбэгийн үлдэгдэл, засварын хүсэлт, эсвэл нэр томьёоны талаар асуугаарай.", examples: HELP.examples };
    if (/^(юу байна|юу сонин|яаж байна|сайн байна уу чи|чи сайн уу|ямар байна)/.test(q))
      return { text: "Сайн, баярлалаа" + name + "! Файлууд бэлэн бол хайлтад туслахад бэлэн байна. Та юу хайж байна?" };
    if (/(баярлалаа|баярлаа|гоё|гое|сайхан|зүгээр|зугээр|мундаг|goy|thanks|thank)/.test(q) && w.length <= 4)
      return { text: "Зүгээр ээ" + name + " 😊 Өөр асуух зүйл байвал хэлээрэй." };
    if (/^(баяртай|дараа уулзъя|дараа уулзая|бай|bye|амжилт)/.test(q))
      return { text: "Баяртай" + name + "! Амжилт хүсье 👋" };
    if (/(чи хэн|та хэн|нэр чинь|нэрийг чинь|юу хийж чадах|юу хийдэг)/.test(q))
      return { text: "Би **Ажлын хэрэгслийн туслах**. Интернэтгүй ажилладаг. Оруулсан файлуудаас сэлбэгийн үлдэгдэл, эдийн дугаар, засварын хүсэлтийг хайж, засварын нэр томьёог тайлбарлана. Та надад шинэ зүйл зааж ч болно: `заа: асуулт = хариулт`.", examples: HELP.examples };
    if (/(хэдний одор|хэдэн сарын|ямар одор|огноо|хэдэн цаг)/.test(q)) {
      var d = new Date(), days = ["Ням", "Даваа", "Мягмар", "Лхагва", "Пүрэв", "Баасан", "Бямба"];
      return { text: "Өнөөдөр **" + d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) + "**, " + days[d.getDay()] + " гараг. Цаг **" + pad(d.getHours()) + ":" + pad(d.getMinutes()) + "**." };
    }
    if (/(онигоо|инээд|хошин)/.test(q)) return { text: JOKES[Math.floor(Math.random() * JOKES.length)] };
    if (/^(за|ок|ok|okay|тийм|ойлголоо|ойлгосон|болно|zaa)$/.test(w.join(" "))) return { text: "За" + name + ". Дараагийн асуултаа бичээрэй." };
    if (/(ойлгохгуй|буруу|муу|алдаа|ажиллахгуй)/.test(q))
      return { text: "Уучлаарай" + name + " 🙏 Би урьдчилан заасан төрлийн асуултад л хариулдаг. Сэлбэгийн нэр, 13 оронтой код, машины дугаар, эсвэл \"… гэж юу вэ\" гэж асуугаарай.", examples: HELP.examples };
    return null;
  }
  var ASKS = /(гэж юу вэ|гэж юу бэ|гэдэг нь юу|гэдэг юу вэ|юу гэсэн уг|гэж юуг|тухай тайлбарла|тайлбарлаж ог|тайлбарла|юу вэ|юу бэ)\s*\??$/;
  // key under which a taught question is stored (and looked up)
  function customKey(q) { return fold(q).replace(/[?!.]+$/, "").trim(); }
  function teach(question, kb) {
    var m = /^\s*заа\s*:?\s*(.+?)\s*=\s*(.+)$/i.exec(question);
    if (m) {
      kb.custom = kb.custom || {};
      kb.custom[customKey(m[1])] = { q: m[1].trim(), a: m[2].trim() };
      kb.dirty = true;
      return { text: "Ойлголоо, санаж авлаа ✅\n**" + m[1].trim() + "** → " + m[2].trim() };
    }
    m = /^\s*март(?:аа|)\s*:?\s*(.+)$/i.exec(question);
    if (m && kb.custom) {
      var k = customKey(m[1]);
      if (kb.custom[k]) { delete kb.custom[k]; kb.dirty = true; return { text: "Мартлаа: **" + m[1].trim() + "**" }; }
      return { text: "\"" + m[1].trim() + "\" гэсэн заасан хариулт олдсонгүй." };
    }
    if (/^(заасан|заасныг|миний заасан)/.test(fold(question))) {
      var keys = Object.keys(kb.custom || {});
      return keys.length ? { text: "Та надад **" + keys.length + "** зүйл заасан:", table: table(["Асуулт", "Хариулт"], keys.map(function (x) { return [kb.custom[x].q, kb.custom[x].a]; }), 100) }
        : { text: "Та надад одоогоор юу ч заагаагүй байна. Жишээ нь: `заа: ЗБН гэж юу вэ = байх ёстой нөөц`" };
    }
    return null;
  }
  function knowledge(q, kb) {
    var clean = q.replace(/[?!.]+$/, "").trim();
    if (kb.custom) {
      if (kb.custom[clean]) return { text: kb.custom[clean].a };
      var ck = Object.keys(kb.custom).filter(function (k) { return k.length > 2 && (clean.indexOf(k) >= 0 || k.indexOf(clean) >= 0); })[0];
      if (ck) return { text: kb.custom[ck].a };
    }
    if (!ASKS.test(q)) return null;
    var term = q.replace(ASKS, "").replace(/[?!.,]/g, " ").trim();
    if (!term) return null;
    var shown = lowerText(kb._question || term).replace(/[?!.,]/g, " ").trim().split(" ").slice(0, words(term).length).join(" ");
    var best = null, bestLen = 0;
    GLOSSARY.forEach(function (g) {
      g[0].forEach(function (alias) {
        var a = fold(alias);
        var hit = term === a || new RegExp("(^|\\s)" + a.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&") + "").test(term);
        if (hit && a.length > bestLen) { best = g; bestLen = a.length; }
      });
    });
    if (best) return { text: best[1] };
    return { text: "\"" + shown + "\"-ийн тухай мэдлэг надад алга. Би интернэтгүй тул зөвхөн бэлэн мэдлэгийн сангаас хариулдаг.\nТа надад зааж өгч болно: `заа: " + shown + " гэж юу вэ = …`" };
  }
  function followUp(q, kb) {
    var last = kb.last || {};
    if (last.answer && (/^(excel|эксел|татах|татаж ав|excel болго|эксел болго|файл болго)/.test(q) || /(excel|эксел) (болго|татах|гарга)/.test(q)))
      return Object.assign({}, last.answer, { text: "Сүүлийн хайлт (**" + last.phrase + "**)-ыг Excel болгох товч доор байна." });
    // "зөвхөн ЦХ", "УХ,ЦХ нь хэрэгтэй", "TL . TKH үлдэгдэл л хэрэгтэй": narrow the last search by warehouse
    var t = " " + q.replace(/[,.;:/\\|+&()!?-]/g, " ") + " ";
    var wantTkh = /\s(цх|tkh|цагаан хад[а-я]*)\s/.test(t), wantTl = /\s(ух|tl|ухаа худаг[а-я]*)\s/.test(t);
    var rest = t.replace(/\s(цх|tkh|цагаан хад[а-я]*|ух|tl|ухаа худаг[а-я]*)(?=\s)/g, " ")
      .replace(/\s(зовхон|дээр|дээрх|дээрхи|дахь|дахи|агуулах|агуулахын|тэр|уунээс|эндээс|нь|л|хэрэгтэй|байна|улдэгдэл|улдэгдлийг|улдэгдэлийг|харуул|харуулаач|гарга|байгаа|ба|болон|мон|and|хоёр|2|хоер|дагуу|дээрхийг|шуу)(?=\s)/g, " ").trim();
    // "цагаан хадын гагнуур" with a request file is a question about repair requests, not parts
    var isService = kb.requests && kb.requests.list.some(function (r) { var v = fold(r.srv); return v && rest.indexOf(v) >= 0; });
    if ((wantTkh || wantTl) && rest && kb.stock && !kb._narrowing && !isService) {
      // "ЦХ дээр маслын шүүр": search the part, then narrow to the warehouse
      var part = (" " + String(kb._question || q).toLowerCase().replace(/[,.;:/\\|+&()!?-]/g, " ") + " ")
        .replace(/\s(цх|tkh|цагаан хад[а-яөү]*|ух|tl|ухаа худаг[а-яөү]*)(?=\s)/g, " ").trim();
      var ph = part && phraseAnswer(part, kb);
      if (!ph || !ph.table || !ph.table.all) return null;
      kb.last = { phrase: phraseOf(part), rows: ph.table.all, answer: ph };
      kb._narrowing = true;
      try { return followUp(wantTl && wantTkh ? "ух цх" : wantTl ? "ух" : "цх", kb); } finally { kb._narrowing = false; }
    }
    if ((wantTkh || wantTl) && last.rows && !rest) {
      var locs = (wantTl ? ["TL"] : []).concat(wantTkh ? ["TKH"] : []);
      var where = function (r) { var u = String(r[3]).toUpperCase(); return u.indexOf("TKH") === 0 ? "TKH" : u.indexOf("TL") === 0 ? "TL" : ""; };
      var rows = last.rows.filter(function (r) { return locs.indexOf(where(r)) >= 0; });
      var sums = { TL: 0, TKH: 0 }, per = {};
      rows.forEach(function (r) {
        var w = where(r), n = Number(r[5]) || 0; sums[w] += n;
        per[r[0]] = per[r[0]] || { name: r[1], TL: 0, TKH: 0 }; per[r[0]][w] += n;
      });
      var NAME = { TL: "Ухаа худаг (TL)", TKH: "Цагаан хад (TKH)" };
      var perRows = Object.keys(per).map(function (k) { return [k, per[k].name].concat(locs.map(function (l) { return per[k][l]; })); })
        .sort(function (a, b) { return b.slice(2).reduce(function (x, y) { return x + y; }, 0) - a.slice(2).reduce(function (x, y) { return x + y; }, 0); });
      return {
        text: "**\"" + last.phrase + "\"**, зөвхөн " + locs.map(function (l) { return "**" + NAME[l] + "**: " + fmt(sums[l]); }).join(", ") + ". Нийт " + rows.length + " мөр, " + perRows.length + " сэлбэг.",
        table: table(last.answer.table.head, rows, 500),
        detailTitle: "Сэлбэг тус бүрээр:",
        // keep the contract-warehouse columns (⭐ list, supplier stock) of the full search
        detail: table(["Item code", "Нэр"].concat(kb.catalog ? ["Гэрээт"] : []).concat(locs.map(function (l) { return NAME[l]; })).concat(Object.keys(kb.suppliers || {})),
          perRows.map(function (r) {
            return r.slice(0, 2).concat(kb.catalog ? [kb.catalog[r[0]] ? "⭐" : ""] : []).concat(r.slice(2).map(fmt))
              .concat(Object.keys(kb.suppliers || {}).map(function (sn) { return fmt(kb.suppliers[sn][r[0]] || 0); }));
          }), 200),
        exportName: last.phrase + "_" + locs.join("_")
      };
    }
    return null;
  }

  function answer(question, kb) {
    var q = fold(question);
    if (!q) return HELP;
    var t = teach(question, kb) || smallTalk(q, kb);
    if (t) return t;
    if (/тусла|юу асуу|жишээ|help|заавар/.test(q)) return HELP;
    kb._question = question;
    var mt = maintAnswer(question, q, kb);
    if (mt) return mt;
    var k = knowledge(q, kb);
    if (k) return k;
    var f = followUp(q, kb);
    if (f) return f;

    var reqs = kb.requests;
    // machine number
    var plate = /(?:^|[^\d])(\d{4})\s*([a-zа-яөү]{2,3})(?![a-zа-яөү])/i.exec(question) || (/^\s*(\d{4})\s*$/.exec(question));
    if (plate && !/^\d{13}$/.test(question.trim())) {
      if (!reqs) return need("requests");
      var key = plateKey(plate[1] + (plate[2] || ""));
      var hits = reqs.list.filter(function (r) { var k = plateKey(r.tech); return plate[2] ? k === key : k.indexOf(plate[1]) === 0; });
      if (!hits.length) return { text: "**" + question.trim() + "** дугаартай машины засварын хүсэлт энэ файлд алга." };
      return {
        text: "**" + hits[0].tech + "**: " + hits.length + " хүсэлт",
        table: table(["Техникийн №", "Төрөл", "Байршил", "Үйлчилгээ", "Хүсэлт", "Хоног", "Тайлбар"],
          hits.map(function (r) { return [r.tech, r.type, r.loc, r.srv, r.day, r.days, r.desc]; }))
      };
    }

    // waiting / old requests
    if (/хоног|удс|удаж|хулээ|хуучин/.test(q) && (reqs || /хусэлт/.test(q))) {
      if (!reqs) return need("requests");
      var n = /(\d+)\s*\+?\s*хоног/.exec(q);
      var min = n ? parseInt(n[1], 10) : 8;
      var old = reqs.list.filter(function (r) { return r.days >= min; }).sort(function (a, b) { return b.days - a.days; });
      return {
        text: min + "+ хоног хүлээж буй хүсэлт: **" + old.length + "** (" + reqs.today + "-ний байдлаар)",
        table: table(["Техникийн №", "Байршил", "Үйлчилгээ", "Хүсэлт", "Хоног", "Тайлбар"],
          old.map(function (r) { return [r.tech, r.loc, r.srv, r.day, r.days, r.desc]; }))
      };
    }

    // request counts with optional location / service filters
    var loc = null, srv = null;
    Object.keys(LOCS).forEach(function (k) {
      var re = k.length <= 2 ? new RegExp("(^|\\s)" + fold(k) + "(\\s|$)") : new RegExp("(^|\\s)" + fold(k));
      if (re.test(q)) loc = LOCS[k];
    });
    if (reqs) {
      uniq(reqs.list.map(function (r) { return r.srv; })).forEach(function (s) {
        if (s !== "(хоосон)" && new RegExp("(^|\\s)" + fold(s).slice(0, 5)).test(q)) srv = s;
      });
    }
    // A service word alone ("гагнуур хэд") means requests; with other words
    // ("гагнуурын утас") it is a spare part name.
    if (srv && !loc && !/хусэлт|засвар/.test(q)) {
      var rest = q.replace(/[?.,!]/g, " ").split(" ").filter(function (w) {
        return w && STOP.map(fold).indexOf(w) < 0 && w.indexOf(fold(srv).slice(0, 5)) !== 0;
      });
      if (rest.length) srv = null;
    }
    if ((/хусэлт|засвар/.test(q) || loc || srv) && !/шуур|сэлбэг|улдэгдэл/.test(q)) {
      if (!reqs) return need("requests");
      var list = reqs.list.filter(function (r) { return (!loc || r.loc === loc) && (!srv || r.srv === srv); });
      var label = [loc, srv].filter(Boolean).join(" / ") || "Нийт";
      if (loc || srv) {
        return {
          text: label + ": **" + list.length + "** хүсэлт",
          table: table(["Техникийн №", "Байршил", "Үйлчилгээ", "Хүсэлт", "Хоног", "Тайлбар"],
            list.map(function (r) { return [r.tech, r.loc, r.srv, r.day, r.days, r.desc]; }))
        };
      }
      var byLoc = {};
      list.forEach(function (r) { var k = r.loc + "|" + r.srv; byLoc[k] = (byLoc[k] || 0) + 1; });
      return {
        text: "Нийт **" + list.length + "** засварын хүсэлт (" + reqs.today + "-ний байдлаар)",
        table: table(["Байршил", "Үйлчилгээ", "Тоо"], Object.keys(byLoc).sort().map(function (k) { var p = k.split("|"); return [p[0], p[1], byLoc[k]]; }), 50)
      };
    }

    // spare parts: Oracle stock and/or the spare-part list (catalog)
    var stock = kb.stock, cat = kb.catalog;
    if (!stock && !cat) return /\d{13}|шуур|улдэгдэл|сэлбэг|хэд байна|хаана/.test(q) ? need("stock") : notUnderstood(kb);

    // contract items (the list = parts taken from the contract suppliers' warehouses)
    // that none of the loaded suppliers has in stock
    if (cat && /(гэрээт|2025|жагсаалт)/.test(q) && /(байхгуй|алга|дууссан|дутуу|0)/.test(q)) {
      var supN = Object.keys(kb.suppliers || {});
      if (!supN.length) return { text: "Гэрээт сэлбэгийн үлдэгдлийг шалгахын тулд **нийлүүлэгчийн файл** (AODE, Очлуур од, Parts and oil) хэрэгтэй. Дээрх \"Файлууд\" хэсэгт оруулна уу." };
      var none = Object.keys(cat).map(function (k) { return cat[k]; }).filter(function (c) {
        return supN.every(function (sn) { return !(kb.suppliers[sn][c.code] > 0); });
      });
      return {
        text: "Гэрээт **" + Object.keys(cat).length + "** сэлбэгээс **" + none.length + "** нь оруулсан нийлүүлэгчдэд (" + supN.join(", ") + ") алга." +
          (supN.length < 3 ? "\nБүх нийлүүлэгчийн файлыг оруулаагүй тул тоо өсөж харагдаж болно." : ""),
        table: table(["Item code", "Нэр"].concat(stock ? ["TL (УХ)", "TKH (ЦХ)"] : []), none.map(function (c) {
          var it = stock && stock[c.code]; return [c.code, c.mn || c.en].concat(stock ? [fmt(it ? it.tl : 0), fmt(it ? it.tkh : 0)] : []);
        }), 30)
      };
    }

    var code = /\b\d{13}\b/.exec(question);
    var fuzzyNote = "";
    if (!code) {
      var ph = phraseAnswer(question, kb);
      if (ph) { kb.last = { phrase: phraseOf(question), rows: kb.stock ? ph.table.all : null, answer: ph }; return ph; }
      fuzzyNote = "Яг таарах илэрц олдсонгүй. Ойролцоо илэрц: ";
    }
    var codes;
    if (code) {
      codes = (stock && stock[code[0]]) || (cat && cat[code[0]]) ? [code[0]] : [];
    } else {
      var words = q.replace(/[?.,!]/g, " ").split(" ").filter(function (w) { return w && STOP.map(fold).indexOf(w) < 0; });
      if (!words.length) return notUnderstood(kb);
      var stems = words.map(function (w) { return w.length > 5 && /[а-я]/.test(w) ? w.slice(0, w.length - 2) : w; });
      var hit = function (text) { return stems.every(function (st) { return text.indexOf(st) >= 0; }); };
      var set = {};
      if (stock) Object.keys(stock).forEach(function (k) { if (hit(stock[k].search) || (cat && cat[k] && hit(cat[k].search))) set[k] = 1; });
      if (cat) Object.keys(cat).forEach(function (k) { if (hit(cat[k].search)) set[k] = 1; });
      codes = Object.keys(set);
    }
    if (!codes.length) return { text: "\"" + question.trim() + "\" гэсэн сэлбэг " + (stock ? "Oracle-ийн үлдэгдэлд" : "") + (stock && cat ? " болон " : "") + (cat ? "гэрээт сэлбэгийн жагсаалтад" : "") + " олдсонгүй. Өөр үгээр, эдийн дугаараар эсвэл 13 оронтой Item code-оор хайгаад үзээрэй." };
    var qty = function (k) { var it = stock && stock[k]; return it ? it.tl + it.tkh : 0; };
    codes.sort(function (a, b) {
      return ((cat && cat[b] ? 1 : 0) - (cat && cat[a] ? 1 : 0)) || (qty(b) - qty(a)) || ((stock && stock[b] ? stock[b].total : 0) - (stock && stock[a] ? stock[a].total : 0));
    });
    var sup = kb.suppliers || {};
    var supNames = Object.keys(sup);
    var name = function (k) { var it = stock && stock[k], c = cat && cat[k]; return (it && (it.mn || it.en)) || (c && (c.mn || c.en)) || ""; };
    var head = ["Item code", "Нэр"].concat(cat ? ["Гэрээт"] : []).concat(stock ? ["TL (УХ)", "TKH (ЦХ)"] : ["Тайлбар"]).concat(supNames);
    var rows = codes.map(function (k) {
      var it = stock && stock[k], c = cat && cat[k];
      return [k, name(k)].concat(cat ? [c ? "⭐" : ""] : [])
        .concat(stock ? [fmt(it ? it.tl : 0), fmt(it ? it.tkh : 0)] : [c ? c.desc : ""])
        .concat(supNames.map(function (sn) { return fmt(sup[sn][k] || 0); }));
    });
    var inCat = cat ? codes.filter(function (k) { return cat[k]; }).length : 0;
    var text = fuzzyNote + (codes.length === 1 ? "**" + name(codes[0]) + "** (" + codes[0] + ")" : "**" + codes.length + "** сэлбэг олдлоо" + (cat ? ", үүнээс **" + inCat + "** нь гэрээт ⭐" : "") + ":");
    if (!stock) text += "\nOracle-ийн үлдэгдлийн файл оруулбал агуулахын тоо харагдана.";
    var res = { text: text, table: table(head, rows, 15) };
    if (codes.length === 1) {
      var one = stock && stock[codes[0]], c1 = cat && cat[codes[0]];
      if (c1) res.text += "\n⭐ Гэрээт агуулахаас авдаг сэлбэг." + (supNames.length ? "" : " Нийлүүлэгчийн файл оруулбал гэрээт агуулахын үлдэгдэл харагдана.") +
        "\n" + c1.desc + (c1.type ? " · " + c1.type : "") + (c1.group ? " · " + c1.group : "");
      if (one) res.detail = table(["Агуулах (тавиур)", "Тоо (" + one.uom + ")"], Object.keys(one.subs).sort().map(function (s2) { return [s2, fmt(one.subs[s2])]; }), 30);
      else if (stock && !c1) res.text += "\n⚠️ Энэ сэлбэг Oracle-ийн үлдэгдэлд алга (агуулахад 0).";
    }
    return res;
  }
  /* Excel filter "contains": the phrase as typed (case-insensitive), row by row. */
  function lowerText(t) { return String(t == null ? "" : t).toLowerCase().replace(/\s+/g, " "); }
  function phraseOf(question) {
    var stop = STOP.map(fold);
    return lowerText(question).replace(/[?!]/g, " ").split(" ").filter(function (w) { return w && stop.indexOf(fold(w.replace(/[.,]$/, ""))) < 0; }).join(" ").replace(/[.,]$/, "");
  }
  function phraseAnswer(question, kb) {
    var phrase = phraseOf(question), stock = kb.stock, cat = kb.catalog;
    if (phrase.length < 2) return null;
    var rows = stock ? stock._rows.filter(function (r) { return lowerText(r.en).indexOf(phrase) >= 0 || lowerText(r.mn).indexOf(phrase) >= 0; }) : [];
    var catHits = cat ? Object.keys(cat).filter(function (k) { return lowerText(cat[k].desc).indexOf(phrase) >= 0; }) : [];
    if (!rows.length && !catHits.length) return null;
    var sup = kb.suppliers || {}, supNames = Object.keys(sup);
    if (!stock) {
      return {
        text: "**\"" + phrase + "\"** гэсэн сэлбэг гэрээт жагсаалтад **" + catHits.length + "** байна.\nOracle-ийн үлдэгдлийн файл оруулбал агуулахын мөрүүд харагдана.",
        table: table(["Item дугаар", "Сэлбэгийн нэр"].concat(supNames), catHits.map(function (k) { return [k, cat[k].desc].concat(supNames.map(function (sn) { return fmt(sup[sn][k] || 0); })); }), 500),
        exportName: phrase
      };
    }
    var codes = uniq(rows.map(function (r) { return r.code; }));
    var tl = 0, tkh = 0;
    rows.forEach(function (r) { var u = r.sub.toUpperCase(); if (u.indexOf("TKH") === 0) tkh += r.qty; else if (u.indexOf("TL") === 0) tl += r.qty; });
    var head = ["Item Code", "Item Mongolian Description", "UOM", "Subinventory", "Locator", "Quantity", "Supplier", "Origination date"].concat(cat ? ["Гэрээт"] : []);
    var list = rows.map(function (r) {
      return [r.code, r.mn || r.en, r.uom, r.sub, r.loc, r.qty, r.supplier, r.date].concat(cat ? [cat[r.code] ? "⭐" : ""] : []);
    });
    var perItem = codes.map(function (k) {
      var it = stock[k];
      return [k, it.mn || it.en].concat(cat ? [cat[k] ? "⭐" : ""] : []).concat([fmt(it.tl), fmt(it.tkh)])
        .concat(supNames.map(function (sn) { return fmt(sup[sn][k] || 0); }));
    }).sort(function (a, b) { return parseFloat(String(b[cat ? 3 : 2]).replace(/,/g, "")) + parseFloat(String(b[cat ? 4 : 3]).replace(/,/g, "")) - parseFloat(String(a[cat ? 3 : 2]).replace(/,/g, "")) - parseFloat(String(a[cat ? 4 : 3]).replace(/,/g, "")); });
    var extra = cat ? catHits.filter(function (k) { return !stock[k]; }) : [];
    return {
      text: "**\"" + phrase + "\"**: Oracle-ийн үлдэгдэлд **" + rows.length + " мөр** (" + codes.length + " сэлбэг). TL нийт **" + fmt(tl) + "**, TKH нийт **" + fmt(tkh) + "**." +
        (cat ? "\nГэрээт жагсаалтад **" + catHits.length + "**" + (extra.length ? ", үүнээс " + extra.length + " нь Oracle-ийн үлдэгдэлд алга." : ".") : ""),
      table: table(head, list, 500),
      detailTitle: "Сэлбэг тус бүрээр:",
      detail: table(["Item code", "Нэр"].concat(cat ? ["Гэрээт"] : []).concat(["TL (УХ)", "TKH (ЦХ)"]).concat(supNames), perItem, 200),
      exportName: phrase
    };
  }
  function uniq(a) { return a.filter(function (x, i) { return a.indexOf(x) === i; }); }
  function notUnderstood(kb) {
    if (!kb.stock && !kb.catalog && !kb.requests)
      return { text: "Энэ асуултад хариулахад файл хэрэгтэй. Oracle үлдэгдэл, гэрээт жагсаалт эсвэл засварын хүсэлтийн файлаа **нэг удаа** оруулна уу. Дараа нь энэ компьютерт хадгалагдах тул дахин оруулах шаардлагагүй.\nФайлгүйгээр нэр томьёоны тайлбар, таны заасан хариулт, энгийн яриа ажиллана.", examples: HELP.examples };
    return { text: "Уучлаарай, асуултыг ойлгосонгүй. Жишээ асуултуудаас сонгоод үзээрэй.", examples: HELP.examples };
  }

  return { load: load, answer: answer, fold: fold, plateKey: plateKey, customKey: customKey, HELP: HELP };
});
