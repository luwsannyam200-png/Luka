/* Loader PM: browser + Node.
 * Хэрэглэгчийн «Loader PM*.xlsx» (LONKING хуудас) файлд PUU-ийн мото цагийг бичиж, засварын мэдээнээс ТҮ-д орсон
 * ачигчийг архивлана (E, U багана + «<он> loader archive.xlsx»-д мөр). Файлын бусад бүх зүйл (томьёо, нөхцөлт формат,
 * бусад хуудас) хэвээр: зөвхөн өөрчилсөн нүднүүдийн XML-г солино, Excel нээхдээ дахин тооцно (fullCalcOnLoad).
 *
 * LONKING: нэг мөр = нэг техник, F = =Z<мөр>. B улсын дугаар, C байршил, D загвар /парк/, E сүүлийн ТҮ
 * ("MM/dd/yyyy <мот/цаг>"), R asset, S WO, U сүүлд хийсэн ТҮ-ийн цэг, V = U+250 (Heli 200, Howo 500), X улсын дугаар
 * (өнгөөр төлөв: шар = өнөөдөр шинэчлэгдсэн, улаан = шинэ утга бага → бичээгүй, цагаан = ирээгүй),
 * Z мото цаг: тоо, эсвэл тоолуур сольсон бол =AA+AB(+AC) — PUU одоогийн тоолуурын утгыг ирүүлдэг тул тэр нүдэнд бичнэ. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("jszip"), require("xlsx"));
  } else {
    root.LoaderCore = factory(root.JSZip, root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, XLSX) {
  "use strict";

  var REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  var YELLOW = "FFFFFF00", RED = "FFFF0000";

  function pad(n) { return String(n).padStart(2, "0"); }
  function xesc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function low(s) { return String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase(); }
  function num(v) {
    if (typeof v === "number") return isFinite(v) ? v : null;
    var n = parseFloat(String(v == null ? "" : v).replace(/\s/g, "").replace(",", "."));
    return isFinite(n) ? n : null;
  }
  function r1(v) { return Math.round(v * 10) / 10; }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function colLetter(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function splitRef(ref) { var m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref); return m ? { c: m[1], r: +m[2] } : null; }
  function usDate(d) { return pad(d.getMonth() + 1) + "/" + pad(d.getDate()) + "/" + d.getFullYear(); }

  /* ---------- улсын дугаар ---------- */
  var LAT = { A: "А", B: "В", E: "Е", K: "К", M: "М", H: "Н", O: "О", P: "Р", C: "С", T: "Т", Y: "У", X: "Х", U: "У" };
  function normPlate(s) { return String(s == null ? "" : s).toUpperCase().replace(/[ABEKMHOPCTYXU]/g, function (c) { return LAT[c]; }).replace(/[\s\-–._]/g, ""); }
  function digits(s) { return (String(s == null ? "" : s).match(/\d+/g) || []).join(""); }
  /* Цифр нь ижил, үсэг нь ижил (эсвэл бичээгүй) — ачигч 4984ММА ба машин 4984ӨМА-г андуурахгүй */
  function findMachine(machines, plate) {
    var n = normPlate(plate); if (!n) return null;
    for (var i = 0; i < machines.length; i++) if (normPlate(machines[i].plate) === n) return machines[i];
    var d = digits(n), let_ = n.replace(/\d/g, "");
    if (d.length < 3) return null;
    var c = machines.filter(function (x) { return digits(x.plate) === d && (!let_ || normPlate(x.plate).replace(/\d/g, "") === let_); });
    return c.length === 1 ? c[0] : null;
  }
  function findByPark(machines, park) {
    var p = String(park == null ? "" : park).trim().replace(/^(ГР|GR)\s*-?\s*/i, "");
    if (!/^(\d{1,3}|EV-?\d)$/i.test(p)) return null;
    var c = machines.filter(function (x) { return x.park && normPlate(x.park) === normPlate(p); });
    return c.length === 1 ? c[0] : null;
  }

  /* ---------- ТҮ-ийн дүрэм ---------- */
  function kindOf(model) {
    var m = String(model || "");
    if (/heli|cpcd/i.test(m)) return "heli";
    if (/AS4006|apol/i.test(m)) return "apollo";
    if (/howo|огтор/i.test(m)) return "howo";
    return "std";
  }
  function levelAt(x, p) {
    if (p == null) return "";
    if (x.kind === "heli") return p % 1200 === 0 ? "PM1200" : p % 600 === 0 ? "PM600" : "PM200";
    if (x.kind === "apollo") return p % 2400 === 0 ? "ТҮ-2" : "ТҮ-1";
    if (x.kind === "howo") return p % 200000 === 0 ? "ТҮ-3/ИЗ" : p % 100000 === 0 ? "ТҮ-3/УЗ-2" : p % 50000 === 0 ? "ТҮ-3/УЗ-1" : p % 30000 === 0 ? "ТҮ-2" : "ТҮ-1";
    return p % 1000 === 0 ? "ТҮ-3" : p % 500 === 0 ? "ТҮ-2" : "ТҮ-1";
  }

  /* ---------- багц (xlsx) ---------- */
  async function openPkg(bytes) {
    var zip = await JSZip.loadAsync(bytes);
    var wb = await zip.file("xl/workbook.xml").async("string");
    var rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    var sheets = [], re = /<sheet\b[^>]*>/g, m;
    while ((m = re.exec(wb))) {
      var name = /\sname="([^"]*)"/.exec(m[0])[1].replace(/&amp;/g, "&");
      var rid = /\sr:id="([^"]*)"/.exec(m[0]) || /\sid="([^"]*)"/.exec(m[0]);
      var rel = new RegExp('<Relationship\\b[^>]*Id="' + rid[1] + '"[^>]*>').exec(rels);
      var target = /Target="([^"]*)"/.exec(rel[0])[1];
      sheets.push({ name: name, path: target.charAt(0) === "/" ? target.slice(1) : "xl/" + target.replace(/^\.\//, "") });
    }
    return { zip: zip, wb: wb, sheets: sheets };
  }
  function findCell(xml, ref) {
    var m = new RegExp('<c r="' + ref + '"(?=[\\s>/])[^>]*?(?:/>|>[\\s\\S]*?</c>)').exec(xml);
    return m ? { start: m.index, end: m.index + m[0].length, xml: m[0] } : null;
  }
  function cellStyle(xml, ref) { var c = findCell(xml, ref); var s = c && /\ss="(\d+)"/.exec(c.xml.slice(0, c.xml.indexOf(">") + 1)); return s ? +s[1] : 0; }
  /* Нүдийг солих (байхгүй бол мөрөнд баганын дарааллаар нэмнэ). make(oldTag|null) → шинэ <c> */
  function putCell(xml, ref, make) {
    var c = findCell(xml, ref);
    if (c) return xml.slice(0, c.start) + make(c.xml) + xml.slice(c.end);
    var p = splitRef(ref), rowRe = new RegExp('<row r="' + p.r + '"(?=[\\s>/])[^>]*?(/>|>)');
    var rm = rowRe.exec(xml), cell = make(null);
    if (!rm) {   // мөр байхгүй: дарааллаар нь шинээр нэмнэ
      var after = /<row r="(\d+)"/g, mm, pos = xml.indexOf("</sheetData>");
      while ((mm = after.exec(xml))) if (+mm[1] > p.r) { pos = mm.index; break; }
      if (pos < 0) throw new Error("sheetData олдсонгүй");
      return xml.slice(0, pos) + '<row r="' + p.r + '">' + cell + "</row>" + xml.slice(pos);
    }
    if (rm[1] === "/>") return xml.slice(0, rm.index) + rm[0].slice(0, -2) + ">" + cell + "</row>" + xml.slice(rm.index + rm[0].length);
    var bodyStart = rm.index + rm[0].length, bodyEnd = xml.indexOf("</row>", bodyStart);
    var body = xml.slice(bodyStart, bodyEnd), cre = /<c r="([A-Z]+)\d+"/g, cm, at = body.length, want = colNum(p.c);
    while ((cm = cre.exec(body))) if (colNum(cm[1]) > want) { at = cm.index; break; }
    return xml.slice(0, bodyStart) + body.slice(0, at) + cell + body.slice(at) + xml.slice(bodyEnd);
  }
  function sAttr(old, s) { if (s != null) return ' s="' + s + '"'; var m = old && /\ss="(\d+)"/.exec(old.slice(0, old.indexOf(">") + 1)); return m ? ' s="' + m[1] + '"' : ""; }
  function numCell(ref, v, s) { return function (old) { return '<c r="' + ref + '"' + sAttr(old, s) + "><v>" + v + "</v></c>"; }; }
  function strCell(ref, t, s) { return function (old) { return '<c r="' + ref + '"' + sAttr(old, s) + ' t="inlineStr"><is><t>' + xesc(t) + "</t></is></c>"; }; }
  /* Томьёот нүдний хадгалсан утгыг шинэчилнэ (томьёо хэвээр) */
  function cachedCell(v) { return function (old) { if (!old || old.indexOf("<f") < 0) return old || ""; var o = old.replace(/\st="[^"]*"/, "").replace(/<v>[\s\S]*?<\/v>/, ""); return o.replace("</c>", "<v>" + v + "</v></c>"); }; }
  function hasFormula(xml, ref) { var c = findCell(xml, ref); return !!(c && c.xml.indexOf("<f") >= 0); }

  /* styles.xml: X нүдний дүүргэлтийг солих (шар / улаан / хоосон) */
  function Styles(xml) {
    var fillsM = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(xml), xfsM = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
    var fills = fillsM[1].match(/<fill\b[\s\S]*?<\/fill>|<fill\s*\/>/g) || [];
    var xfs = xfsM[1].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g) || [];
    var cache = {}, fillIdx = {};
    function fillOf(s) { var x = xfs[s] || xfs[0]; var f = /\sfillId="(\d+)"/.exec(x); return fills[f ? +f[1] : 0] || ""; }
    return {
      fillOf: fillOf,
      withFill: function (s, rgb) {
        var k = s + "|" + rgb; if (k in cache) return cache[k];
        var fid = 0;
        if (rgb) {
          if (!(rgb in fillIdx)) { fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + rgb + '"/><bgColor indexed="64"/></patternFill></fill>'); fillIdx[rgb] = fills.length - 1; }
          fid = fillIdx[rgb];
        }
        var x = (xfs[s] || xfs[0]).replace(/\sfillId="\d+"/, "").replace(/\sapplyFill="\d"/, "").replace(/^<xf\b/, '<xf fillId="' + fid + '" applyFill="1"');
        xfs.push(x); cache[k] = xfs.length - 1; return cache[k];
      },
      xml: function () {
        return xml.replace(/<fills\b[^>]*>[\s\S]*?<\/fills>/, '<fills count="' + fills.length + '">' + fills.join("") + "</fills>")
          .replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, '<cellXfs count="' + xfs.length + '">' + xfs.join("") + "</cellXfs>");
      }
    };
  }
  /* Одоогийн тоолуурын цайвар цэнхэр нүд (#B7DEE8 орчим, эсвэл accent-ийн цайвар өнгө) */
  function isLightBlue(fill) {
    var m = /fgColor\b[^>]*rgb="([0-9A-F]{6}(?:[0-9A-F]{2})?)"/i.exec(fill);
    if (m) { m[1] = m[1].slice(-6); var r = parseInt(m[1].slice(0, 2), 16), g = parseInt(m[1].slice(2, 4), 16), b = parseInt(m[1].slice(4), 16); return b > 200 && g > 180 && r < b - 20 && r > 120; }
    return /fgColor\b[^>]*theme="(4|8)"[^>]*tint="0\.[3-9]/.test(fill);
  }

  /* ---------- Loader PM файл унших ---------- */
  function lonkingSheet(pkg) { return pkg.sheets.find(function (s) { return /^lonking$/i.test(s.name.trim()); }); }
  async function readLoader(bytes) {
    var pkg = await openPkg(bytes), sh = lonkingSheet(pkg);
    if (!sh) throw new Error("LONKING хуудас олдсонгүй — Loader PM файл мөн эсэхийг шалгана уу");
    var wbx = XLSX.read(bytes, { type: "array", cellFormula: true, cellDates: false, sheets: [sh.name] });
    var ws = wbx.Sheets[sh.name], xml = await pkg.zip.file(sh.path).async("string");
    var styles = Styles(await pkg.zip.file("xl/styles.xml").async("string"));
    function cell(c, r) { return ws[c + r] || null; }
    function val(c, r) { var x = cell(c, r); return x ? x.v : null; }
    var range = XLSX.utils.decode_range(ws["!ref"]), out = [];
    for (var r = 1; r <= Math.min(range.e.r + 1, 300); r++) {
      var f = cell("F", r); if (!f || !f.f || !/^\$?Z\$?\d+$/i.test(f.f.replace(/^=/, ""))) continue;   // техникийн мөр: F = =Z<мөр>
      var plate = String(val("B", r) == null ? "" : val("B", r)).trim() || String(val("X", r) == null ? "" : val("X", r)).trim();
      var model = String(val("D", r) == null ? "" : val("D", r)).trim();
      if (!plate && !model) continue;
      var pm = /\/\s*([^/]+?)\s*\//.exec(model);
      var z = cell("Z", r), Z = num(z && z.v), refs = [];
      var zf = z && z.f ? z.f.replace(/[=\s$]/g, "") : "";
      if (zf && /^[A-Z]{1,3}\d+(\+[A-Z]{1,3}\d+)+$/.test(zf)) {
        zf.split("+").forEach(function (ref) {
          var c = cell(splitRef(ref).c, splitRef(ref).r);
          refs.push({ ref: ref, value: num(c && c.v), blue: isLightBlue(styles.fillOf(cellStyle(xml, ref))) });
        });
      } else if (!zf) refs.push({ ref: "Z" + r, value: Z, blue: false });
      var v = cell("V", r), vf = v && v.f ? /U\$?\d+\s*\+\s*(\d+)/.exec(v.f) : null;
      var x = {
        row: r, plate: plate || model, model: model, park: pm ? pm[1] : String(val("Y", r) == null ? "" : val("Y", r)).replace(/^(ГР|GR)\s*-?\s*/i, ""),
        loc: String(val("C", r) == null ? "" : val("C", r)).trim(), asset: String(val("R", r) == null ? "" : val("R", r)).trim(),
        wo: String(val("S", r) == null ? "" : val("S", r)).trim(), E: val("E", r) == null ? "" : String(val("E", r)),
        Z: Z, U: num(val("U", r)), V: num(v && v.v), step: vf ? +vf[1] : null, refs: refs, zFormula: zf,
        uFormula: !!(cell("U", r) && cell("U", r).f)
      };
      x.kind = kindOf(model);
      if (x.step == null) x.step = x.kind === "heli" ? 200 : x.kind === "howo" ? 500 : x.kind === "apollo" ? 300 : 250;
      if (x.V == null && x.U != null) x.V = x.U + x.step;
      x.level = levelAt(x, x.V);
      out.push(x);
    }
    if (!out.length) throw new Error("LONKING хуудаснаас техникийн мөр (F = =Z…) олдсонгүй");
    return { sheet: sh.name, machines: out };
  }

  /* ---------- мото цаг: аль нүдэнд, ямар төлөвтэй ---------- */
  // PUU одоогийн тоолуурын утгыг ирүүлнэ: тоолуур сольсон бол цайвар цэнхэр нүд, эсвэл утгаараа хамгийн ойр нүд
  function resolveHours(x, v) {
    if (v == null || !x.refs.length) return null;
    var blue = x.refs.filter(function (r) { return r.blue; });
    var t = blue.length === 1 ? blue[0] : x.refs.slice().sort(function (a, b) { return Math.abs((a.value || 0) - v) - Math.abs((b.value || 0) - v); })[0];
    var old = t.value, z = x.Z != null && old != null ? r1(x.Z - old + v) : v;
    var status = old == null ? "upd" : v < old ? "low" : v - old > 400 ? "jump" : "upd";
    return { ref: t.ref, old: old, value: v, newZ: status === "low" ? x.Z : z, status: status };
  }
  /* OCR / хуулсан текстийн тоог боломжит утга руу ойртуулна (цэг, таслал алдсан бол /10, /100) */
  function plausible(x, nums) {
    var ref = x.refs.length ? x.refs.slice().sort(function (a, b) { return (b.blue ? 1 : 0) - (a.blue ? 1 : 0); }) : [];
    var olds = ref.map(function (r) { return r.value; }).filter(function (v) { return v != null; });
    if (!olds.length) return nums.length ? nums[nums.length - 1] : null;
    var best = null, score = Infinity;
    nums.forEach(function (n) {
      [n, n / 10, n / 100].forEach(function (c, k) {
        olds.forEach(function (o) {
          var d = c - o, s = (d >= 0 ? d : -d * 3) + k * 50;
          if (d > 600 || d < -o * 0.5) return;
          if (s < score) { score = s; best = r1(c); }
        });
      });
    });
    return best;
  }
  /* Мөр бүрээс техникийг (улсын дугаарын цифрээр) таньж, мото цагийг авна. UHG PUU-ийн имэйлийн хүснэгт, OCR текст */
  function parsePuuText(text, machines) {
    var out = {};
    String(text || "").split(/\r?\n/).forEach(function (line) {
      var l = line.replace(/(\d)\s*[-–]\s*(\d)/g, "$1$2").replace(/(\d),(\d{3})(?=[.\s]|$)/g, "$1$2");
      var toks = l.match(/\d+(?:[.,]\d+)?/g) || [];
      var x = null, used = -1;
      for (var i = 0; i < toks.length && !x; i++) {
        var d = toks[i].replace(/\D/g, "");
        if (d.length < 3 || /[.,]/.test(toks[i])) continue;
        var cand = machines.filter(function (m) { return digits(m.plate) === d; });
        if (cand.length === 1) { x = cand[0]; used = i; }
        else if (cand.length > 1) { x = cand.find(function (m) { return normPlate(l).indexOf(normPlate(m.plate)) >= 0; }) || null; used = i; }
      }
      if (!x) return;
      var nums = toks.filter(function (t, i) { return i !== used; }).map(num).filter(function (n) { return n != null && n >= 100; });
      var v = plausible(x, nums);
      if (v != null) out[x.row] = { row: x.row, value: v, line: line.trim() };
    });
    return Object.keys(out).map(function (k) { return out[k]; });
  }
  /* Timemotion.xlsx (tkh.puu21, Terminal Puu1) эсвэл имэйлийн хүснэгтийг хадгалсан Excel: «Улсын дугаар» + «Эцсийн мото цаг» / «Сүүлийн м/цаг».
     Техник бүрийн хамгийн сүүлийн огнооны хамгийн их утга; 0 = өөр газар ажиллаж байгаа → алгасна */
  function parsePuuBook(bytes, machines) {
    var wb = XLSX.read(bytes, { type: "array", cellDates: true }), best = {}, found = false;
    wb.SheetNames.forEach(function (sn) {
      var rs = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
      for (var h = 0; h < Math.min(rs.length, 15); h++) {
        var hd = (rs[h] || []).map(low);
        var ip = hd.findIndex(function (c) { return c.indexOf("улсын дугаар") >= 0; });
        var iv = hd.findIndex(function (c) { return /эцсийн мото|сүүлийн м\/?ц|сүүлийн мото/.test(c); });
        if (ip < 0 || iv < 0) continue;
        var idt = hd.findIndex(function (c) { return c.indexOf("огноо") >= 0; });
        found = true;
        rs.slice(h + 1).forEach(function (r, i) {
          var x = findMachine(machines, r[ip]), v = num(r[iv]); if (!x || !v) return;
          var d = idt >= 0 && r[idt] instanceof Date ? r[idt].getTime() : i;
          var b = best[x.row]; if (!b || d > b.d || (d === b.d && v > b.value)) best[x.row] = { row: x.row, value: v, d: d, line: sn };
        });
        break;
      }
    });
    return found ? Object.keys(best).map(function (k) { return best[k]; }) : null;
  }

  /* ---------- засварын мэдээ / huselt: ТҮ-д орсон ачигч ---------- */
  var TY_RE = /(?:^|[^а-яёөүa-z])(?:[тt][үуy]|техник\s*үйлчилгээ)\s*[-–—]?\s*([123])/i;
  function parseWhen(v) {
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === "number" && v > 30000) return new Date(Math.round((v - 25569) * 864e5) + new Date().getTimezoneOffset() * 6e4);
    var m = /(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/.exec(String(v || ""));
    return m ? new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : null;
  }
  function isRepairBook(bytes) {
    var wb = XLSX.read(bytes, { type: "array", sheetRows: 15 });
    return wb.SheetNames.some(function (sn) { return repairHeader(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null })) >= 0; });
  }
  function repairHeader(rs) {
    for (var i = 0; i < Math.min(rs.length, 15); i++) { var h = (rs[i] || []).map(low); if (h.some(function (c) { return c.indexOf("засварын төрөл") >= 0; }) && h.some(function (c) { return c.indexOf("техникийн") >= 0; })) return i; }
    return -1;
  }
  function parseEDate(e) { var m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(e || "")); return m ? new Date(+m[3], +m[1] - 1, +m[2]) : null; }
  function parseRepairs(bytes, machines) {
    var wb = XLSX.read(bytes, { type: "array", cellDates: true }), out = [];
    wb.SheetNames.forEach(function (sn) {
      var rs = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null }), h = repairHeader(rs);
      if (h < 0) return;
      var hd = rs[h].map(low);
      function col(re) { return hd.findIndex(function (c) { return re.test(c); }); }
      var C = { a: col(/засварын төрөл/), b: col(/техникийн/), c: col(/хэсгийн/), d: col(/жолооч/), svc: col(/^үйлчилгээ/), i: col(/дис тайлбар/), j: col(/мех тайлбар/),
        k: col(/шилжүүлсэн/), m: col(/^механик/), t: col(/дууссан/), s: col(/эхэлсэн/) };
      function g(r, k) { return C[k] >= 0 ? r[C[k]] : null; }
      rs.slice(h + 1).forEach(function (r, ri) {
        var drv = String(g(r, "d") == null ? "" : g(r, "d")).trim(), svc = low(g(r, "svc"));
        var x = findMachine(machines, g(r, "b"));
        if (!x && (/ачигч/.test(svc) || !drv)) x = findByPark(machines, g(r, "c"));
        if (!x) return;
        var typ = low(g(r, "a")), plan = typ.indexOf("төлөвлөгөөт") >= 0;
        var ij = [g(r, "i"), g(r, "j")].map(function (v) { return String(v == null ? "" : v); }).filter(Boolean).join(" · ");
        var kk = String(g(r, "k") == null ? "" : g(r, "k"));
        var m = TY_RE.exec(" " + ij), mk = !m && TY_RE.exec(" " + kk);
        var maybe = !m && !mk && (plan || /маслын шүүр|тос\s*сол/i.test(ij));
        if (!m && !mk && !maybe) return;
        var end = parseWhen(g(r, "t")), lvW = m ? "ТҮ-" + m[1] : mk ? "ТҮ-" + mk[1] : "";
        var warn = [];
        if (!plan) warn.push("засварын төрөл «" + (g(r, "a") || "") + "» (Төлөвлөгөөт биш)");
        if (drv) warn.push("жолоочийн нэр бичигдсэн");
        if (lvW && x.kind === "std" && x.level && x.level !== lvW) warn.push("бичсэн " + lvW + " ≠ хуваарь " + x.level + " (" + x.V + ")");
        if (mk) warn.push("зөвхөн шилжүүлсэн тэмдэглэлд — хийгдээгүй байж магадгүй");
        if (!end) warn.push("дуусаагүй");
        if (!m && !mk) warn.push("ТҮ гэж бичээгүй — ТҮ байж магадгүй");
        if (x.Z != null && x.V != null && x.Z < x.V - 60) warn.push("мото цаг (" + x.Z + ") ТҮ-ийн цэгт (" + x.V + ") хүрээгүй");
        var ed = parseEDate(x.E);
        var done = !!(ed && end && ed >= new Date(end.getFullYear(), end.getMonth(), end.getDate()) && x.V != null && x.U != null && x.U >= x.V - x.step);
        if (done) warn.push("E-д " + x.E + " — аль хэдийн архивлагдсан бололтой");
        out.push({ key: sn + ":" + ri, row: x.row, plate: x.plate, model: x.model, text: ij || kk, written: lvW, level: x.level, point: x.V, end: end,
          sure: !!m, ok: !!m && plan && !!end && !mk && !done, warn: warn, done: done });
      });
    });
    return out;
  }

  /* ---------- Loader PM файл үүсгэх ----------
     hours: [{row, value}] (PUU-ийн утга), pms: [{row}] (архивлах техник), today: Date */
  async function buildLoader(bytes, opts) {
    var info = await readLoader(bytes), pkg = await openPkg(bytes), sh = lonkingSheet(pkg);
    var xml = await pkg.zip.file(sh.path).async("string");
    var styles = Styles(await pkg.zip.file("xl/styles.xml").async("string"));
    var today = opts.today || new Date(), byRow = {}, log = [], archive = [];
    info.machines.forEach(function (x) { byRow[x.row] = x; });
    var hv = {}; (opts.hours || []).forEach(function (h) { if (h.value != null) hv[h.row] = h.value; });
    var pmRows = {}; (opts.pms || []).forEach(function (p) { pmRows[p.row] = p; });
    info.machines.forEach(function (x) {
      var r = x.row, h = hv[r] != null ? resolveHours(x, hv[r]) : null, state = h ? (h.status === "low" ? "low" : "upd") : "";
      if (h && h.status !== "low") {
        if (hasFormula(xml, h.ref)) { log.push(x.plate + ": " + h.ref + " томьёотой тул бичсэнгүй"); state = ""; }
        else {
          xml = putCell(xml, h.ref, numCell(h.ref, h.value));
          if (h.ref !== "Z" + r) xml = putCell(xml, "Z" + r, cachedCell(h.newZ));
          xml = putCell(xml, "F" + r, cachedCell(h.newZ));
          if (x.U != null) xml = putCell(xml, "W" + r, cachedCell(r1(h.newZ - x.U)));
        }
      }
      // X: өнөөдрийн мото цагийн төлөв
      var xs = cellStyle(xml, "X" + r);
      xml = putCell(xml, "X" + r, function (old) { return old ? old.replace(/\ss="\d+"/, "").replace(/^<c r="X\d+"/, '<c r="X' + r + '" s="' + styles.withFill(xs, state === "upd" ? YELLOW : state === "low" ? RED : null) + '"') : '<c r="X' + r + '" s="' + styles.withFill(xs, state === "upd" ? YELLOW : state === "low" ? RED : null) + '"/>'; });
      // ТҮ архивлах: E = "<өнөөдөр> <шинэчлэлтээс өмнөх мото цаг>", U ← V
      var p = pmRows[r];
      if (p) {
        var prevZ = x.Z != null ? r1(x.Z) : null;
        var e = usDate(today) + (x.kind === "howo" || prevZ == null ? "" : " " + prevZ);
        xml = putCell(xml, "E" + r, strCell("E" + r, e));
        var point = p.point != null ? p.point : x.V;
        if (x.uFormula) log.push(x.plate + ": U томьёотой тул шилжүүлсэнгүй");
        else if (point != null && (x.U == null || point > x.U)) {
          xml = putCell(xml, "U" + r, numCell("U" + r, point));
          xml = putCell(xml, "V" + r, cachedCell(point + x.step));
          var zNow = h && h.status !== "low" ? h.newZ : x.Z;
          if (zNow != null) xml = putCell(xml, "W" + r, cachedCell(r1(zNow - point)));
        }
        archive.push({ plate: x.plate, b: e, point: point, level: p.level || levelAt(x, point), wo: x.wo });
      }
    });
    pkg.zip.file(sh.path, xml);
    pkg.zip.file("xl/styles.xml", styles.xml());
    pkg.zip.file("xl/workbook.xml", fullCalc(pkg.wb));
    var out = await pkg.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return { bytes: out, archive: archive, log: log };
  }
  function fullCalc(wb) {
    if (/<calcPr\b/.test(wb)) return wb.replace(/<calcPr\b([^>]*?)(\/?)>/, function (m, a, sl) { return "<calcPr" + a.replace(/\sfullCalcOnLoad="[^"]*"/, "") + ' fullCalcOnLoad="1"' + sl + ">"; });
    var at = wb.indexOf("</definedNames>"); at = at >= 0 ? at + 15 : wb.indexOf("</sheets>") + 9;
    return wb.slice(0, at) + '<calcPr fullCalcOnLoad="1"/>' + wb.slice(at);
  }

  /* ---------- архив: «<он> loader archive.xlsx»-д мөр нэмэх ---------- */
  var AHEAD = ["Улсын дугаар", "РМ орсон сүүлийн огноо, м/ц", "Техник Үйлчилгээний М/Цаг", "ТҮ1-250 ТҮ2-500 ТҮ3-1000", "Work order"];
  function isArchiveBook(bytes) {
    var wb = XLSX.read(bytes, { type: "array", sheetRows: 5 });
    return wb.SheetNames.some(function (n) { return /архив/i.test(n); }) && !wb.SheetNames.some(function (n) { return /^lonking$/i.test(n.trim()); });
  }
  async function appendArchive(bytes, rows) {
    var pkg = await openPkg(bytes), sh = pkg.sheets.find(function (s) { return /архив/i.test(s.name); }) || pkg.sheets[0];
    var ws = XLSX.read(bytes, { type: "array", sheets: [sh.name] }).Sheets[sh.name];
    var rs = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    var range = XLSX.utils.decode_range(ws["!ref"] || "A1"), last = 0;
    rs.forEach(function (r, i) { if (r && r.slice(0, 5).some(function (v) { return v != null && v !== ""; })) last = i + 1; });
    last += range.s.r;
    var have = {}; rs.forEach(function (r) { if (r && r[0]) have[normPlate(r[0]) + "|" + num(r[2])] = 1; });
    var xml = await pkg.zip.file(sh.path).async("string"), st = {}, added = [], skipped = [];
    ["A", "B", "C", "D", "E"].forEach(function (c) { st[c] = cellStyle(xml, c + last); });
    var rowTag = new RegExp('<row r="' + last + '"(?=[\\s>/])[^>]*?>').exec(xml), attrs = rowTag ? rowTag[0].replace(/^<row r="\d+"/, "").replace(/\/?>$/, "") : "";
    attrs = attrs.replace(/\sspans="[^"]*"/, "");
    rows.forEach(function (a) {
      if (have[normPlate(a.plate) + "|" + a.point]) { skipped.push(a); return; }
      have[normPlate(a.plate) + "|" + a.point] = 1;
      var r = ++last, cells = strCell("A" + r, a.plate, st.A)(null) + strCell("B" + r, a.b, st.B)(null) +
        (a.point != null ? numCell("C" + r, a.point, st.C)(null) : "") + strCell("D" + r, a.level || "", st.D)(null) + (a.wo ? strCell("E" + r, a.wo, st.E)(null) : "");
      var pos = xml.indexOf("</sheetData>"), nxt = new RegExp('<row r="(\\d+)"', "g"), mm;
      while ((mm = nxt.exec(xml))) if (+mm[1] > r) { pos = mm.index; break; }
      xml = xml.slice(0, pos) + '<row r="' + r + '"' + attrs + ">" + cells + "</row>" + xml.slice(pos);
      added.push(a);
    });
    if (added.length) {
      xml = xml.replace(/<dimension ref="([A-Z]+\d+):([A-Z]+)(\d+)"/, function (m, a, c, n) { return '<dimension ref="' + a + ":" + c + Math.max(+n, last) + '"'; });
      xml = xml.replace(/<autoFilter ref="([A-Z]+\d+):([A-Z]+)(\d+)"/, function (m, a, c, n) { return '<autoFilter ref="' + a + ":" + c + Math.max(+n, last) + '"'; });
      pkg.zip.file(sh.path, xml);
      // Excel-ийн хүснэгт (Table) байвал хүрээг нь сунгана
      var relsPath = sh.path.replace(/worksheets\/([^/]+)$/, "worksheets/_rels/$1.rels"), rf = pkg.zip.file(relsPath);
      if (rf) {
        var rels = await rf.async("string"), tre = /Target="([^"]*tables\/[^"]+)"/g, tm;
        while ((tm = tre.exec(rels))) {
          var tp = tm[1].charAt(0) === "/" ? tm[1].slice(1) : "xl/" + tm[1].replace(/^\.\.\//, ""), tf = pkg.zip.file(tp); if (!tf) continue;
          var tx = (await tf.async("string")).replace(/(<(?:table|autoFilter)\b[^>]*\sref="[A-Z]+\d+:[A-Z]+)(\d+)"/g, function (m, a, n) { return a + Math.max(+n, last) + '"'; });
          pkg.zip.file(tp, tx);
        }
      }
    }
    return { bytes: await pkg.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), added: added, skipped: skipped };
  }
  function newArchive(rows) {
    var ws = XLSX.utils.aoa_to_sheet([AHEAD].concat(rows.map(function (a) { return [a.plate, a.b, a.point, a.level, a.wo || ""]; })));
    ws["!cols"] = [14, 26, 14, 16, 14].map(function (w) { return { wch: w }; });
    var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Архив");
    return new Uint8Array(XLSX.write(wb, { bookType: "xlsx", type: "array" }));
  }
  async function isLoaderBook(bytes) { try { return !!lonkingSheet(await openPkg(bytes)); } catch (e) { return false; } }

  return {
    readLoader: readLoader, resolveHours: resolveHours, parsePuuText: parsePuuText, parsePuuBook: parsePuuBook,
    parseRepairs: parseRepairs, buildLoader: buildLoader, appendArchive: appendArchive, newArchive: newArchive,
    isLoaderBook: isLoaderBook, isArchiveBook: isArchiveBook, isRepairBook: isRepairBook,
    normPlate: normPlate, levelAt: levelAt, findMachine: findMachine
  };
});
