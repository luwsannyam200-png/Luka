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
  var STOP = ["хэд", "хэдэн", "байна", "байгаа", "бий", "юу", "уу", "үү", "вэ", "бэ", "нь", "хаана", "үлдэгдэл", "үлдэгдэлтэй",
    "агуулахад", "агуулах", "ширхэг", "ш", "тоо", "хэмжээ", "сэлбэг", "код", "кодтой", "item", "ямар", "олох", "хай", "хайх",
    "харуул", "мэдээлэл", "дээр", "бол", "тэр", "энэ", "?", "."];

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
    h = findHeader(sheets, ["засварын төрөл", "засварын байршил", "үйлчилгээ", "хүсэлт гаргасан"]);
    if (h) return { kind: "requests", data: loadRequests(h), name: name };
    var sup = loadSupplier(sheets[0]);
    if (Object.keys(sup).length) return { kind: "supplier", data: sup, name: name };
    return { kind: "unknown", name: name };
  }
  function loadStock(h) {
    var c = { code: col(h, "item code"), en: col(h, "item description"), mn: col(h, "item mongolian description"),
      sub: col(h, "subinventory"), loc: col(h, "locator"), q: col(h, "quantity"), uom: col(h, "uom") };
    var items = {};
    h.rows.slice(h.hi + 1).forEach(function (r) {
      var code = r[c.code], q = r[c.q];
      if (code == null || typeof q !== "number") return;
      code = String(code).trim();
      var it = items[code] || (items[code] = { code: code, en: r[c.en] || "", mn: r[c.mn] || "", uom: r[c.uom] || "", subs: {}, total: 0 });
      var sub = String(r[c.sub] || "").trim() || "?";
      var key = sub + (r[c.loc] ? " (" + r[c.loc] + ")" : "");
      it.subs[key] = (it.subs[key] || 0) + q;
      it.total += q;
    });
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
    return { head: head, rows: rows.slice(0, limit), more: Math.max(0, rows.length - limit) };
  }
  function need(kind) {
    var what = { stock: "Oracle-ийн үлдэгдлийн тайлан (жишээ нь 09-26 үлдэгдэл.xlsx)", requests: "Mine2TL-ийн засварын хүсэлтийн файл" }[kind];
    return { text: "Энэ асуултад хариулахын тулд **" + what + "** хэрэгтэй. Дээрх \"Файл нэмэх\" хэсэгт оруулна уу." };
  }
  var HELP = {
    text: "Би оруулсан файлуудаас хайж хариулна. Жишээ асуултууд:",
    examples: ["маслын шүүр хэд байна", "4016150400048 хаана байна", "5840ӨМА", "хүсэлт хэд байна", "цагаан хадын гагнуур хэд", "8 хоногоос удсан хүсэлт"]
  };

  function answer(question, kb) {
    var q = fold(question);
    if (!q) return HELP;
    if (/тусла|юу асуу|жишээ|help|заавар/.test(q) || /^(сайн байна уу|сайн уу|сайн|hi|hello|мэнд|байна уу)[?!. ]*$/.test(q)) return HELP;

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

    // stock lookup
    var stock = kb.stock;
    if (!stock) return /\d{13}|шуур|улдэгдэл|сэлбэг|хэд байна|хаана/.test(q) ? need("stock") : notUnderstood(kb);
    var code = /\b\d{13}\b/.exec(question);
    var found;
    if (code) {
      found = stock[code[0]] ? [stock[code[0]]] : [];
    } else {
      var words = q.replace(/[?.,!]/g, " ").split(" ").filter(function (w) { return w && STOP.map(fold).indexOf(w) < 0; });
      if (!words.length) return notUnderstood(kb);
      var stems = words.map(function (w) { return w.length > 5 ? w.slice(0, w.length - 2) : w; });
      found = Object.keys(stock).map(function (k) { return stock[k]; })
        .filter(function (it) { return stems.every(function (s) { return it.search.indexOf(s) >= 0; }); })
        .sort(function (a, b) { return (b.tl + b.tkh) - (a.tl + a.tkh) || b.total - a.total; });
    }
    if (!found.length) return { text: "\"" + question.trim() + "\" гэсэн сэлбэг Oracle-ийн үлдэгдэлд олдсонгүй. Өөр үгээр, эсвэл 13 оронтой Item code-оор хайгаад үзээрэй." };
    var sup = kb.suppliers || {};
    var supNames = Object.keys(sup);
    var head = ["Item code", "Нэр", "TL (УХ)", "TKH (ЦХ)", "Бусад агуулах"].concat(supNames);
    var rows = found.map(function (it) {
      return [it.code, it.mn || it.en, fmt(it.tl), fmt(it.tkh), fmt(it.total - it.tl - it.tkh)].concat(supNames.map(function (s) { return fmt(sup[s][it.code] || 0); }));
    });
    var res = { text: found.length === 1 ? "**" + (found[0].mn || found[0].en) + "** (" + found[0].code + ")" : "**" + found.length + "** сэлбэг олдлоо:", table: table(head, rows, 15) };
    if (found.length === 1) {
      var it = found[0];
      res.detail = table(["Агуулах (тавиур)", "Тоо (" + it.uom + ")"], Object.keys(it.subs).sort().map(function (s) { return [s, fmt(it.subs[s])]; }), 30);
    }
    return res;
  }
  function uniq(a) { return a.filter(function (x, i) { return a.indexOf(x) === i; }); }
  function notUnderstood(kb) {
    return { text: "Уучлаарай, асуултыг ойлгосонгүй. Жишээ асуултуудаас сонгоод үзээрэй.", examples: HELP.examples };
  }

  return { load: load, answer: answer, fold: fold, plateKey: plateKey, HELP: HELP };
});
