/* =====================================================================
   BRK Genius · 매출 공용 모듈  (/assets/brk-sales.js)

   하는 일
   1) 매출 읽기   BRKSales.load()
      - /data/sales.json (예전 방식, 기본값)
      - 구글시트 «sales» 탭 (엑셀로 올린 값, 이게 우선)
      두 곳을 합쳐서 달별·일별 매출을 돌려줍니다.
   2) 엑셀 읽기   BRKSales.parseFiles(files)
      - .xlsx / .xls / .csv 를 읽어 날짜별 매출·고객수·전년 동요일·제품군을 뽑습니다.
      - 제목 줄의 이름을 보고 칸을 알아서 찾습니다. (양식이 조금 달라도 됩니다)
   3) 올리기 칸   BRKSales.mountUploader(상자, {onApplied})
      - 연간 마케팅 플랜 페이지 맨 아래에 붙어 있습니다.

   구글시트 주소(DATA_URL)는 IP 관리 페이지와 같은 Apps Script 를 씁니다.
   ===================================================================== */
(function (global) {
  "use strict";

  var DATA_URL = "https://script.google.com/macros/s/AKfycbwa42Avu6K35gUEB1MW2e-5Pcu6WKnHaL9Qh_QgfhxQ0myuNHGjIyHuu2emqcl7Sd-JPQ/exec";
  try { var _u = localStorage.getItem("brkg_data_url"); if (_u) DATA_URL = _u; } catch (e) {}

  var XLSX_SRC = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  var CATS_DEFAULT = ["아이스크림", "케이크", "디저트", "음료/커피", "기타"];

  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function dim(y, m) { return new Date(y, m, 0).getDate(); }
  function eok(v, dp) { return (v / 1e8).toFixed(dp == null ? 1 : dp); }
  function comma(n) { return Math.round(n).toLocaleString("ko-KR"); }

  /* ── 시트 읽기 (JSONP : 브라우저 보안에 막히지 않음) ── */
  var _cb = 0;
  function jsonp(params, ms) {
    return new Promise(function (res, rej) {
      if (!DATA_URL) { rej(new Error("DATA_URL 비어 있음")); return; }
      var cb = "brksales" + (++_cb) + "_" + Date.now(), sc = document.createElement("script"), done = false;
      var t = setTimeout(function () { fin(); rej(new Error("시간 초과")); }, ms || 15000);
      function fin() {
        if (done) return; done = true; clearTimeout(t);
        try { delete global[cb]; } catch (e) { global[cb] = undefined; }
        if (sc.parentNode) sc.parentNode.removeChild(sc);
      }
      global[cb] = function (d) { fin(); res(d); };
      var q = [];
      for (var k in params) q.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
      q.push("callback=" + cb, "_=" + Date.now());
      sc.src = DATA_URL + (DATA_URL.indexOf("?") < 0 ? "?" : "&") + q.join("&");
      sc.onerror = function () { fin(); rej(new Error("주소를 열지 못했습니다")); };
      document.head.appendChild(sc);
    });
  }

  /* ─────────────────────────────────────────────
     1) 매출 읽기
     ───────────────────────────────────────────── */
  var cache = null;
  function load(opt) {
    if (cache && !(opt && opt.fresh)) return cache;
    var base = (location.protocol === "file:") ? Promise.resolve(null)
      : fetch("/data/sales.json", { cache: "no-store" })
          .then(function (r) { return r.ok ? r.json() : null; })
          .catch(function () { return null; });
    var sheet = jsonp({ action: "sales" })
      .catch(function (e) { return { ok: false, error: String(e && e.message || e) }; });
    cache = Promise.all([base, sheet]).then(function (r) { return build(r[0], r[1]); });
    return cache;
  }

  function num(v) { var n = +v; return isFinite(n) ? n : 0; }

  function build(base, sheet) {
    var M = { days: {}, mon: {}, base: base || null, sheetOk: !!(sheet && sheet.ok),
              sheetErr: (sheet && !sheet.ok) ? (sheet.error || "읽기 실패") : null,
              sheetRows: 0, lastUpload: null, lastBy: "" };

    /* 예전 방식 sales.json 의 일매출을 바탕으로 깐다 */
    if (base && base.daily) {
      Object.keys(base.daily).forEach(function (mk) {
        var D = base.daily[mk] || {};
        Object.keys(D).forEach(function (d) {
          var v = D[d];
          M.days[mk + "-" + pad(+d)] = { s: num(v.s), c: num(v.c), pw: num(v.pw), pwc: num(v.pwc), cats: null, src: "base" };
        });
      });
    }
    /* 시트에 올라간 값이 있으면 그걸로 덮는다 */
    var rows = (sheet && sheet.rows) || [];
    rows.forEach(function (r) {
      var id = String(r.id || "");
      if (!id) return;
      var cats = r.cats;
      if (typeof cats === "string") { try { cats = JSON.parse(cats); } catch (e) { cats = null; } }
      var v = { s: num(r.s), c: num(r.c), pw: num(r.pw), pwc: num(r.pwc),
                cats: (cats && typeof cats === "object") ? cats : null, src: "sheet" };
      if (id.charAt(0) === "d" && /^\d{4}-\d{2}-\d{2}$/.test(id.slice(1))) M.days[id.slice(1)] = v;
      else if (id.charAt(0) === "m" && /^\d{4}-\d{2}$/.test(id.slice(1))) M.mon[id.slice(1)] = v;
      else return;
      M.sheetRows++;
      if (r.updated && (!M.lastUpload || String(r.updated) > M.lastUpload)) { M.lastUpload = String(r.updated); M.lastBy = r.by || ""; }
    });

    M.month = function (key) { return monthOf(M, key); };
    M.keys = function () {
      var k = {};
      Object.keys(M.days).forEach(function (d) { k[d.slice(0, 7)] = 1; });
      Object.keys(M.mon).forEach(function (m) { k[m] = 1; });
      return Object.keys(k).sort();
    };
    /* 시트에서 온 값이 들어 있는 달만 (페이지에 새로 칠할 달) */
    M.sheetKeys = function () {
      var k = {};
      Object.keys(M.days).forEach(function (d) { if (M.days[d].src === "sheet") k[d.slice(0, 7)] = 1; });
      Object.keys(M.mon).forEach(function (m) { k[m] = 1; });
      return Object.keys(k).sort();
    };
    M.dailyOf = function (key) {           /* {1:{s,c,aov,pw,pwc}, ...} 플랜 페이지 달력용 */
      var y = +key.slice(0, 4), m = +key.slice(5, 7), out = {};
      for (var d = 1; d <= dim(y, m); d++) {
        var v = M.days[key + "-" + pad(d)];
        if (!v) continue;
        out[d] = { s: v.s, c: v.c, aov: v.c ? Math.round(v.s / v.c) : 0, pw: v.pw, pwc: v.pwc };
      }
      return out;
    };
    return M;
  }

  function monthOf(M, key) {
    var y = +key.slice(0, 4), m = +key.slice(5, 7), D = dim(y, m);
    var s = 0, c = 0, pw = 0, pwc = 0, n = 0, last = 0, cats = {}, okC = true, okPw = true, okPwc = true, anyCat = false, fromSheet = false;
    for (var d = 1; d <= D; d++) {
      var v = M.days[key + "-" + pad(d)];
      if (!v) continue;
      n++; last = d; s += v.s;
      if (v.src === "sheet") fromSheet = true;
      if (v.c) c += v.c; else okC = false;
      if (v.pw) pw += v.pw; else okPw = false;
      if (v.pwc) pwc += v.pwc; else okPwc = false;
      if (v.cats) { anyCat = true; for (var k in v.cats) cats[k] = (cats[k] || 0) + num(v.cats[k]); }
    }
    var mo = M.mon[key];
    if (mo) {           /* 월 마감 합계가 따로 올라와 있으면 그게 우선 */
      return { key: key, s: mo.s, c: mo.c, pw: mo.pw, pwc: mo.pwc, n: D, last: D, dim: D, full: true,
               gaps: false, daily: n > 0, cats: mo.cats || (anyCat ? cats : null), fromSheet: true, closed: true };
    }
    if (!n) return null;
    return { key: key, s: s, c: okC ? c : 0, pw: okPw ? pw : 0, pwc: okPwc ? pwc : 0, n: n, last: last, dim: D,
             full: n === D, gaps: n !== last, daily: true, cats: anyCat ? cats : null, fromSheet: fromSheet, closed: false };
  }

  /* ─────────────────────────────────────────────
     2) 엑셀 읽기
     ───────────────────────────────────────────── */
  var _xlsx = null;
  function ensureXLSX() {
    if (global.XLSX) return Promise.resolve(global.XLSX);
    if (_xlsx) return _xlsx;
    _xlsx = new Promise(function (res, rej) {
      var s = document.createElement("script");
      s.src = XLSX_SRC;
      s.onload = function () { global.XLSX ? res(global.XLSX) : rej(new Error("엑셀 읽기 도구를 불러오지 못했습니다")); };
      s.onerror = function () { _xlsx = null; rej(new Error("엑셀 읽기 도구를 불러오지 못했습니다 (인터넷 연결 확인)")); };
      document.head.appendChild(s);
    });
    return _xlsx;
  }

  function readBuf(f) {
    return new Promise(function (res, rej) {
      var fr = new FileReader();
      fr.onload = function () { res(fr.result); };
      fr.onerror = function () { rej(new Error(f.name + " 을(를) 열지 못했습니다")); };
      fr.readAsArrayBuffer(f);
    });
  }
  function decodeText(buf) {
    try { return new TextDecoder("utf-8", { fatal: true }).decode(buf).replace(/^\uFEFF/, ""); }
    catch (e) { try { return new TextDecoder("euc-kr").decode(buf); } catch (e2) { return new TextDecoder().decode(buf); } }
  }

  /* 셀 값 → 날짜 */
  function ymd(y, m, d) {
    if (y < 100) y += 2000;
    if (!(y >= 2000 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= dim(y, m))) return null;
    return y + "-" + pad(m) + "-" + pad(d);
  }
  function serialToYmd(n) {
    var t = new Date(Math.round((n - 25569) * 86400000));
    return ymd(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
  }
  function toDay(v, yh) {
    if (v == null || v === "") return null;
    if (v instanceof Date) return ymd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    if (typeof v === "number") {
      if (v > 36000 && v < 80000) return serialToYmd(v);
      if (v >= 20000101 && v <= 21001231 && v === Math.floor(v)) return ymd(Math.floor(v / 10000), Math.floor(v / 100) % 100, v % 100);
      return null;
    }
    var t = String(v).trim().replace(/\(.*?\)/g, "").replace(/\s+/g, " ").trim();
    var a;
    if ((a = /^(\d{4})\s*[-.\/년]\s*(\d{1,2})\s*[-.\/월]\s*(\d{1,2})\s*일?\.?$/.exec(t))) return ymd(+a[1], +a[2], +a[3]);
    if ((a = /^(\d{2})\s*[-.\/]\s*(\d{1,2})\s*[-.\/]\s*(\d{1,2})$/.exec(t))) return ymd(+a[1], +a[2], +a[3]);
    if ((a = /^(20\d{2})(\d{2})(\d{2})$/.exec(t))) return ymd(+a[1], +a[2], +a[3]);
    if (yh && (a = /^(\d{1,2})\s*[\/.월-]\s*(\d{1,2})\s*일?$/.exec(t))) return ymd(yh, +a[1], +a[2]);
    return null;
  }
  function toMonth(v, yh) {
    if (v == null || v === "") return null;
    if (typeof v === "number") return null;
    var t = String(v).trim().replace(/\s+/g, "");
    var a, y, m;
    if ((a = /^(\d{4})[-.\/년](\d{1,2})월?\.?$/.exec(t))) { y = +a[1]; m = +a[2]; }
    else if ((a = /^(\d{2})[-.\/](\d{1,2})$/.exec(t)) && +a[1] > 12) { y = 2000 + +a[1]; m = +a[2]; }
    else if ((a = /^(\d{2})년(\d{1,2})월$/.exec(t))) { y = 2000 + +a[1]; m = +a[2]; }
    else if (yh && (a = /^(\d{1,2})월$/.exec(t))) { y = yh; m = +a[1]; }
    else return null;
    if (!(y >= 2000 && y <= 2100 && m >= 1 && m <= 12)) return null;
    return y + "-" + pad(m);
  }
  function toNum(v) {
    if (v == null || v === "") return null;
    if (typeof v === "number") return isFinite(v) ? v : null;
    var t = String(v).trim();
    if (!t || /^[-–—]$/.test(t)) return null;
    var neg = /^\(.*\)$/.test(t) || /^-/.test(t) || /^△|^▼/.test(t);
    t = t.replace(/[^0-9.]/g, "");
    if (!t || t === ".") return null;
    var n = parseFloat(t);
    return isFinite(n) ? (neg ? -n : n) : null;
  }

  /* 제목 줄 글자 → 무슨 칸인지 */
  var RX = {
    date:  /일자|날짜|영업일|^date$|기준일|매출일|년월일|^일$|^일별$/i,
    month: /^(년월|월|월별|기간|month|연월)$/i,
    cat:   /제품군|카테고리|품목군|대분류|중분류|^구분$|^분류$|category|상품군/i,
    bad:   /비중|구성|%|률|율|比|대비|증감|객단가|단가|수량|목표|누계|누적|평균|순번|^no\.?$|매장수|점포수|순위|신장|차이|성장|계획|예상|전망|기여|index/i,
    prev:  /전년|작년|전기|^ly|py$|^py|동요일/i,
    cust:  /고객|객수|영수증|구매건수|결제건수/,
    total: /합계|총계|총매출|전체|실매출|순매출|^pos|pos매출|^매출액?(\(.*\))?$|^금액(\(.*\))?$|^실적$|토탈|total|^계$|^sales$/i
  };
  var CUR = /금년|당년|올해|this|^실적/gi;

  function classify(lab, yrCur) {
    var L = lab.replace(/\s+/g, "");
    if (!L) return null;
    if (RX.date.test(L)) return { k: "date" };
    if (RX.month.test(L)) return { k: "month" };
    if (RX.cat.test(L)) return { k: "cat" };
    if (RX.bad.test(L)) return null;
    var prev = RX.prev.test(L);
    var yr = /(20\d{2}|\d{2})년/.exec(L);
    if (yr && yrCur) { var yv = +yr[1]; if (yv < 100) yv += 2000; if (yv < yrCur) prev = true; }
    if (RX.cust.test(L)) return { k: prev ? "pwc" : "c" };
    if (prev) return { k: "pw" };
    var L2 = L.replace(CUR, "").replace(/(20\d{2}|\d{2})년/, "");
    if (!L2 || RX.total.test(L2)) return { k: "s" };
    var name = L2.replace(/매출액?|금액|실적|\(.*?\)|원$/g, "");
    if (!name) return { k: "s" };
    return { k: "catcol", name: name };
  }

  function unitFrom(text) {
    var a = /단위\s*[:：]?\s*\(?\s*(백만원|천원|억원|억|만원|원)/.exec(text) || /\(\s*(백만원|천원|억원|만원|억)\s*\)/.exec(text);
    if (!a) return null;
    return { "원": 1, "천원": 1e3, "만원": 1e4, "백만원": 1e6, "억원": 1e8, "억": 1e8 }[a[1]] || null;
  }
  var UNIT_NAME = { 1: "원", 1000: "천원", 10000: "만원", 1000000: "백만원", 100000000: "억원" };

  /* 병합된 칸은 왼쪽 위 값을 나머지 칸에도 채운다 (그래야 제목·날짜가 빠지지 않음) */
  function fillMerges(X, ws) {
    (ws["!merges"] || []).forEach(function (mg) {
      var src = ws[X.utils.encode_cell(mg.s)];
      if (!src) return;
      for (var r = mg.s.r; r <= mg.e.r; r++) for (var c = mg.s.c; c <= mg.e.c; c++) {
        if (r === mg.s.r && c === mg.s.c) continue;
        ws[X.utils.encode_cell({ r: r, c: c })] = { t: src.t, v: src.v, w: src.w };
      }
    });
  }

  function yearHintOf(name, rows) {
    var txt = rows.slice(0, 8).map(function (r) { return (r || []).join(" "); }).join(" ") + " " + name;
    var a = /(20\d{2})\s*년/.exec(txt) || /(20\d{2})[.\-_ ]?\d{2}/.exec(txt) || /(?:^|\D)(2\d)[.년_ ]\s?\d{1,2}(?:\D|$)/.exec(name);
    if (a) { var y = +a[1]; return y < 100 ? 2000 + y : y; }
    return new Date().getFullYear();
  }

  /* 한 장(시트) 읽기 */
  function parseSheet(rows, sheetName, fileName) {
    var out = { days: {}, mon: {}, map: [], sheet: sheetName, file: fileName, unit: null, mode: "" };
    if (!rows.length) return out;
    var yh = yearHintOf(fileName + " " + sheetName, rows);
    var headText = rows.slice(0, 12).map(function (r) { return (r || []).join(" "); }).join(" ");
    out.unit = unitFrom(headText);
    var maxC = 0; rows.forEach(function (r) { if (r && r.length > maxC) maxC = r.length; });
    var cell = function (r, c) { return rows[r] ? rows[r][c] : null; };
    var txt = function (v) { return v == null ? "" : String(v).trim(); };

    /* ── A. 날짜가 세로로 내려가는 표 ── */
    var best = null;
    for (var c = 0; c < Math.min(maxC, 8); c++) {
      var hits = 0, first = -1, monthHits = 0, mFirst = -1;
      for (var r = 0; r < rows.length; r++) {
        if (toDay(cell(r, c), yh)) { hits++; if (first < 0) first = r; }
        else if (toMonth(cell(r, c), yh)) { monthHits++; if (mFirst < 0) mFirst = r; }
      }
      if (hits >= 1 && (!best || hits > best.hits)) best = { c: c, hits: hits, first: first, kind: "d" };
      if (monthHits >= 2 && monthHits > hits && (!best || monthHits > best.hits)) best = { c: c, hits: monthHits, first: mFirst, kind: "m" };
    }
    /* ── B. 날짜가 가로로 늘어선 표 (제목 줄에 날짜가 여러 개) ── */
    var across = null;
    for (var r2 = 0; r2 < Math.min(rows.length, 30); r2++) {
      var dc = [];
      for (var c2 = 0; c2 < maxC; c2++) { var dd = toDay(cell(r2, c2), yh); if (dd) dc.push({ c: c2, d: dd }); }
      if (dc.length >= 3 && (!across || dc.length > across.cols.length)) across = { r: r2, cols: dc };
    }
    if (across && (!best || across.cols.length > best.hits)) return parseAcross(rows, across, out, yh);
    if (!best) return out;
    out.mode = best.kind === "m" ? "월별" : "일별";

    /* 제목 줄 : 첫 자료 줄 위로 최대 3줄을 합쳐 칸 이름을 만든다 */
    var hTop = Math.max(0, best.first - 3), labels = [];
    for (var cc = 0; cc < maxC; cc++) {
      var parts = [];
      for (var hr = hTop; hr < best.first; hr++) {
        var t = txt(cell(hr, cc));
        if (t && !toDay(t, yh) && parts.indexOf(t) < 0 && t.length < 40 && !/단위|^\(.*\)$|^20\d{2}년\s*\d{1,2}월$/.test(t)) parts.push(t);
      }
      labels.push(parts.join(" "));
    }
    /* 제목 줄에 연도가 둘 이상이면 큰 쪽이 올해 */
    var yrs = []; labels.forEach(function (l) { var a = /(20\d{2}|\d{2})년/.exec(l.replace(/\s/g, "")); if (a) { var y = +a[1]; yrs.push(y < 100 ? y + 2000 : y); } });
    var yrCur = yrs.length ? Math.max.apply(null, yrs) : 0;

    var colK = [];
    for (var ci = 0; ci < maxC; ci++) {
      if (ci === best.c) { colK.push({ k: best.kind === "m" ? "month" : "date" }); continue; }
      var k = classify(labels[ci], yrCur);
      if (k && (k.k === "date" || k.k === "month")) k = null;
      /* 숫자가 하나도 없는 칸은 버린다 (단, 제품군 칸은 글자) */
      if (k && k.k !== "cat") {
        var any = false;
        for (var rr = best.first; rr < rows.length && !any; rr++) if (toNum(cell(rr, ci)) !== null) any = true;
        if (!any) k = null;
      }
      colK.push(k);
    }
    var catCol = -1; colK.forEach(function (k, i) { if (k && k.k === "cat" && catCol < 0) catCol = i; });
    var has = function (kk) { return colK.some(function (k) { return k && k.k === kk; }); };
    var colsOf = function (kk) { var a = []; colK.forEach(function (k, i) { if (k && k.k === kk) a.push(i); }); return a; };

    /* 같은 종류 칸이 여럿이면 첫 칸만 쓴다 (금년·전년이 둘 다 '매출'로만 적힌 경우 두 번째를 전년으로) */
    var sCols = colsOf("s"), cCols = colsOf("c");
    if (sCols.length >= 2 && !has("pw")) { colK[sCols[1]] = { k: "pw" }; }
    if (cCols.length >= 2 && !has("pwc")) { colK[cCols[1]] = { k: "pwc" }; }
    colK.forEach(function (k, i) {
      if (!k) return;
      if (["s", "c", "pw", "pwc"].indexOf(k.k) > -1 && colsOf(k.k)[0] !== i) colK[i] = null;
    });

    /* 읽은 칸 목록 (확인용) */
    var NAMES = { date: "날짜", month: "월", cat: "제품군 이름", s: "매출", c: "고객수", pw: "전년 동요일 매출", pwc: "전년 동요일 고객수" };
    var catNames = [];
    colK.forEach(function (k, i) {
      if (!k) return;
      if (k.k === "catcol") catNames.push(k.name);
      else out.map.push(NAMES[k.k] + " = " + (labels[i] || colName(i)));
    });
    if (catNames.length) out.map.push("제품군 " + catNames.length + "칸 (" + catNames.join(", ") + ")");

    var lastKey = null;
    for (var r3 = best.first; r3 < rows.length; r3++) {
      var raw = cell(r3, best.c);
      var key = best.kind === "m" ? toMonth(raw, yh) : toDay(raw, yh);
      if (!key) {
        if (raw != null && txt(raw)) { lastKey = null; continue; }   /* 합계·소계 줄 */
        if (catCol >= 0 && lastKey && txt(cell(r3, catCol))) key = lastKey;   /* 날짜가 첫 줄에만 있는 표 */
        else continue;
      }
      lastKey = key;
      var bucket = best.kind === "m" ? out.mon : out.days;
      var rec = bucket[key] || (bucket[key] = { s: 0, c: 0, pw: 0, pwc: 0, cats: {}, _tot: false, _pwTot: false });
      var get = function (kk) { var i = colK.findIndex(function (k) { return k && k.k === kk; }); return i < 0 ? null : toNum(cell(r3, i)); };

      if (catCol >= 0) {                       /* 한 줄에 제품군 하나 */
        var cn = txt(cell(r3, catCol));
        var sv = get("s"); if (sv === null) { var cs = colsOf("catcol"); if (cs.length) sv = toNum(cell(r3, cs[0])); }
        if (/^(합계|총계|전체|계|소계|total|총합계)$/i.test(cn.replace(/\s/g, ""))) {
          if (sv !== null) { rec.s = sv; rec._tot = true; }
          var cv = get("c"), pv = get("pw"), pcv = get("pwc");
          if (cv !== null) rec.c = cv; if (pv !== null) { rec.pw = pv; rec._pwTot = true; } if (pcv !== null) rec.pwc = pcv;
        } else if (cn && sv !== null) {
          rec.cats[cn] = (rec.cats[cn] || 0) + sv;
          if (!rec._tot) rec._sumCats = true;
          var pv2 = get("pw"); if (pv2 !== null && !rec._pwTot) rec._pwCats = (rec._pwCats || 0) + pv2;
          var cv2 = get("c"); if (cv2 !== null) rec._cCats = Math.max(rec._cCats || 0, cv2);
        }
      } else {                                  /* 한 줄에 하루 */
        var s1 = get("s");
        colK.forEach(function (k, i) {
          if (k && k.k === "catcol") { var v = toNum(cell(r3, i)); if (v !== null) rec.cats[k.name] = (rec.cats[k.name] || 0) + v; }
        });
        if (s1 !== null) { rec.s = s1; rec._tot = true; }
        var c1 = get("c"), p1 = get("pw"), pc1 = get("pwc");
        if (c1 !== null) rec.c = c1; if (p1 !== null) rec.pw = p1; if (pc1 !== null) rec.pwc = pc1;
      }
    }
    finish(out.days); finish(out.mon);
    return out;
  }

  function colName(i) { var s = ""; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s + "열"; }

  function finish(bucket) {
    Object.keys(bucket).forEach(function (k) {
      var r = bucket[k];
      var catSum = 0, nc = 0; for (var c in r.cats) { catSum += r.cats[c]; nc++; }
      if (!r._tot) r.s = catSum;
      if (!r.pw && r._pwCats) r.pw = r._pwCats;
      if (!r.c && r._cCats) r.c = r._cCats;
      if (!nc) r.cats = null;
      delete r._tot; delete r._pwTot; delete r._sumCats; delete r._pwCats; delete r._cCats;
      if (!r.s) delete bucket[k];
    });
  }

  /* 날짜가 가로로 늘어선 표 */
  function parseAcross(rows, across, out, yh) {
    out.mode = "일별(가로)";
    var yrCur = 0, group = "";
    var catNames = [], seen = {};
    for (var r = across.r + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var lab = "", c0 = row[0] == null ? "" : String(row[0]).trim();
      if (c0 && toNum(c0) === null) group = c0;
      for (var c = 0; c < across.cols[0].c; c++) { var t = row[c] == null ? "" : String(row[c]).trim(); if (t && toNum(t) === null) lab += (lab ? " " : "") + t; }
      if (!c0 && lab && group && across.cols[0].c > 1) lab = group + " " + lab;
      if (!lab) continue;
      var k = classify(lab, yrCur);
      if (!k || k.k === "date" || k.k === "month" || k.k === "cat") continue;
      if (seen[k.k] && k.k !== "catcol") continue;
      seen[k.k] = 1;
      if (k.k === "catcol") catNames.push(k.name);
      else out.map.push({ s: "매출", c: "고객수", pw: "전년 동요일 매출", pwc: "전년 동요일 고객수" }[k.k] + " = " + lab);
      across.cols.forEach(function (dc) {
        var v = toNum(row[dc.c]); if (v === null) return;
        var rec = out.days[dc.d] || (out.days[dc.d] = { s: 0, c: 0, pw: 0, pwc: 0, cats: {}, _tot: false });
        if (k.k === "catcol") rec.cats[k.name] = (rec.cats[k.name] || 0) + v;
        else { rec[k.k] = v; if (k.k === "s") rec._tot = true; }
      });
    }
    if (catNames.length) out.map.push("제품군 " + catNames.length + "줄 (" + catNames.join(", ") + ")");
    finish(out.days);
    return out;
  }

  /* 파일 여러 개 → 하나로 */
  function parseFiles(files) {
    files = Array.prototype.slice.call(files || []);
    if (!files.length) return Promise.reject(new Error("파일이 없습니다"));
    return ensureXLSX().then(function (X) {
      return Promise.all(files.map(function (f) {
        return readBuf(f).then(function (buf) {
          var isText = /\.(csv|tsv|txt)$/i.test(f.name);
          var wb = isText ? X.read(decodeText(buf), { type: "string", raw: true })
                          : X.read(new Uint8Array(buf), { type: "array", cellDates: false });
          return wb.SheetNames.map(function (sn) {
            var ws = wb.Sheets[sn];
            if (!ws || !ws["!ref"]) return null;
            fillMerges(X, ws);
            var rows = X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
            try { return parseSheet(rows, sn, f.name); } catch (e) { return { days: {}, mon: {}, map: [], sheet: sn, file: f.name, err: String(e) }; }
          }).filter(Boolean);
        });
      })).then(function (lists) {
        var R = { days: {}, mon: {}, sheets: [], notes: [], unitGuess: 1, unitFound: null };
        var all = [].concat.apply([], lists);
        all.forEach(function (sh) {
          var nd = Object.keys(sh.days).length, nm = Object.keys(sh.mon).length;
          if (!nd && !nm) return;
          if (sh.unit) R.unitFound = sh.unit;
          R.sheets.push({ file: sh.file, sheet: sh.sheet, mode: sh.mode, days: nd, months: nm, map: sh.map });
          Object.keys(sh.days).forEach(function (k) { R.days[k] = sh.days[k]; });
          Object.keys(sh.mon).forEach(function (k) { R.mon[k] = sh.mon[k]; });
        });
        if (!R.sheets.length) throw new Error("날짜와 매출이 들어 있는 표를 찾지 못했습니다. 첫 칸이 날짜(예: 2026-09-01)이고 제목 줄에 «매출»이 있는지 확인해 주세요.");
        /* 단위 짐작 : 하루 매출은 보통 10억~40억 원, 한 달은 500억~1,100억 원 */
        R.unitGuess = R.unitFound || guessUnit(R);
        return R;
      });
    });
  }

  function guessUnit(R) {
    var vals = Object.keys(R.days).map(function (k) { return R.days[k].s; }).sort(function (a, b) { return a - b; });
    var expect = 2e9;
    if (!vals.length) { vals = Object.keys(R.mon).map(function (k) { return R.mon[k].s; }).sort(function (a, b) { return a - b; }); expect = 7e10; }
    if (!vals.length) return 1;
    var med = vals[Math.floor(vals.length / 2)];
    if (med <= 0) return 1;
    var best = 1, bestD = Infinity;
    [1, 1e3, 1e4, 1e6, 1e8].forEach(function (f) { var d = Math.abs(Math.log10(med * f / expect)); if (d < bestD) { bestD = d; best = f; } });
    return best;
  }

  /* 파싱 결과 → 시트에 올릴 줄 (원 단위로 바꿈) */
  function toRecords(R, unit) {
    var f = unit || 1, recs = [];
    var mk = function (id, v) {
      var cats = null;
      if (v.cats) { cats = {}; for (var k in v.cats) cats[k] = Math.round(v.cats[k] * f); }
      return { id: id, kind: id.charAt(0), s: Math.round(v.s * f), c: Math.round(v.c || 0),
               pw: Math.round((v.pw || 0) * f), pwc: Math.round(v.pwc || 0), cats: cats };
    };
    Object.keys(R.days).sort().forEach(function (k) { recs.push(mk("d" + k, R.days[k])); });
    Object.keys(R.mon).sort().forEach(function (k) { recs.push(mk("m" + k, R.mon[k])); });
    return recs;
  }

  /* 시트에 올리기 */
  function upload(recs, key, by) {
    recs.forEach(function (r) { r.by = by || ""; });
    var chunks = [];
    for (var i = 0; i < recs.length; i += 250) chunks.push(recs.slice(i, i + 250));
    var hd = { "Content-Type": "text/plain;charset=utf-8" };
    var p = Promise.resolve({ ok: true, n: 0, blind: false });
    chunks.forEach(function (ch) {
      p = p.then(function (acc) {
        var body = JSON.stringify({ type: "sales", key: key, recs: ch });
        return fetch(DATA_URL, { method: "POST", headers: hd, body: body })
          .then(function (r) { return r.json(); })
          .then(function (j) {
            if (!j || !j.ok) {
              var m = (j && j.error) || "저장하지 못했습니다";
              if (/type/.test(m)) m = "Apps Script 가 아직 예전 버전입니다. brkg-data-apps-script.txt 를 다시 붙여넣고 «배포 관리 → 새 버전»으로 배포해 주세요.";
              var e = new Error(m); e.hard = true; throw e;
            }
            acc.n += ch.length; return acc;
          })
          .catch(function (e) {
            if (e && e.hard) throw e;
            /* 브라우저가 답을 못 읽는 경우 : 그냥 보내 두고 나중에 다시 읽어 확인 */
            return fetch(DATA_URL, { method: "POST", mode: "no-cors", headers: hd, body: body })
              .then(function () { acc.n += ch.length; acc.blind = true; return acc; });
          });
      });
    });
    return p;
  }

  /* 양식 파일 내려받기 */
  function template() {
    return ensureXLSX().then(function (X) {
      var head = ["날짜", "매출 합계", "고객수", "전년 동요일 매출", "전년 동요일 고객수"].concat(CATS_DEFAULT);
      var ex = [["2026-10-01", 1850000000, 138000, 1720000000, 131000, 1350000000, 330000000, 90000000, 70000000, 10000000],
                ["2026-10-02", 1790000000, 134500, 1810000000, 137200, 1300000000, 322000000, 88000000, 70000000, 10000000]];
      var ws1 = X.utils.aoa_to_sheet([["단위: 원 · 한 줄에 하루치 · 제품군 칸은 없어도 됩니다"], head].concat(ex));
      ws1["!cols"] = head.map(function (h, i) { return { wch: i ? 16 : 12 }; });
      var ws2 = X.utils.aoa_to_sheet([["단위: 원 · 한 달 마감 합계만 있을 때 이 장을 씁니다"], ["월", "매출 합계", "고객수"], ["2026-10", 80000000000, 5600000]]);
      ws2["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 12 }];
      var wb = X.utils.book_new();
      X.utils.book_append_sheet(wb, ws1, "일매출");
      X.utils.book_append_sheet(wb, ws2, "월 마감");
      X.writeFile(wb, "BRK_매출_올리기_양식.xlsx");
    });
  }

  /* ─────────────────────────────────────────────
     3) 올리기 칸 (화면)
     ───────────────────────────────────────────── */
  var CSS =
    ".brks{margin-top:34px;border:1.5px solid #EBD9CD;border-radius:16px;padding:24px 24px 22px;background:#fff}" +
    ".brks h3{font-size:17px;font-weight:900;margin:0 0 4px}" +
    ".brks .lead2{font-size:13px;color:#6E5E55;line-height:1.7;margin:0 0 14px}" +
    ".brks .now{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:12.5px;color:#6E5E55;padding:10px 14px;background:#FBF6F2;border-radius:10px;margin-bottom:14px}" +
    ".brks .now b{color:#000;font-weight:800}" +
    ".brks .now .bad{color:#C0483C}" +
    ".brks .drop{display:block;border:2px dashed #E3C9BB;border-radius:14px;padding:28px 18px;text-align:center;cursor:pointer;transition:border-color .15s,background .15s}" +
    ".brks .drop:hover,.brks .drop:focus-visible,.brks .drop.on{border-color:#E8447A;background:#FFF5F8;outline:none}" +
    ".brks .drop b{display:block;font-size:15px;font-weight:800;margin-bottom:4px}" +
    ".brks .drop span{font-size:12.5px;color:#8A7A70}" +
    ".brks .row{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-top:12px}" +
    ".brks .lnk{font:inherit;font-size:12.5px;color:#6E5E55;background:none;border:0;padding:0;text-decoration:underline;cursor:pointer}" +
    ".brks .lnk:hover{color:#E8447A}" +
    ".brks .msg{font-size:13px;line-height:1.7;margin-top:12px}" +
    ".brks .msg.err{color:#C0483C;font-weight:700}" +
    ".brks .msg.ok{color:#0F6E56;font-weight:700}" +
    ".brks .pv{margin-top:18px}" +
    ".brks .pv h4{font-size:14px;font-weight:900;margin:0 0 8px}" +
    ".brks .tw{overflow-x:auto;border:1px solid #EBD9CD;border-radius:12px}" +
    ".brks table{width:100%;border-collapse:collapse;font-size:13px}" +
    ".brks th{background:#000;color:#fff;font-weight:700;font-size:12px;padding:9px 12px;text-align:left;white-space:nowrap}" +
    ".brks td{padding:9px 12px;border-top:1px solid #F1E6DE;white-space:nowrap}" +
    ".brks td.n{text-align:right;font-variant-numeric:tabular-nums}" +
    ".brks .tag{display:inline-block;font-size:11px;font-weight:800;padding:2px 8px;border-radius:20px}" +
    ".brks .tag.new{background:#E7F2EE;color:#0F6E56}.brks .tag.chg{background:#FDE7EE;color:#C2185B}.brks .tag.same{background:#F1ECE8;color:#8A7A70}" +
    ".brks .up{color:#0F6E56;font-weight:700}.brks .dn{color:#C2185B;font-weight:700}" +
    ".brks .map{font-size:12px;color:#6E5E55;line-height:1.75;margin:10px 0 0;padding-left:0;list-style:none}" +
    ".brks .map li b{color:#000}" +
    ".brks .warn{font-size:12.5px;color:#9A5B00;background:#FFF6E0;border-radius:10px;padding:9px 13px;margin-top:10px;line-height:1.7}" +
    ".brks .act{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin-top:16px;padding-top:16px;border-top:1px solid #F1E6DE}" +
    ".brks label.u{font-size:12.5px;color:#6E5E55;display:flex;align-items:center;gap:6px}" +
    ".brks select,.brks input[type=password],.brks input[type=text]{font:inherit;font-size:13.5px;padding:8px 11px;border:1.5px solid #EBD9CD;border-radius:10px;background:#fff;color:#000}" +
    ".brks input[type=password]{width:120px}.brks input[type=text]{width:110px}" +
    ".brks select:focus,.brks input:focus{outline:none;border-color:#E8447A}" +
    ".brks .go{font:inherit;font-size:14px;font-weight:800;padding:10px 20px;border-radius:11px;border:0;background:#E8447A;color:#fff;cursor:pointer}" +
    ".brks .go:disabled{opacity:.5;cursor:default}" +
    ".brks .go:focus-visible,.brks .ghost:focus-visible{outline:2px solid #000;outline-offset:2px}" +
    ".brks .ghost{font:inherit;font-size:13.5px;font-weight:700;padding:9px 16px;border-radius:11px;border:1.5px solid #EBD9CD;background:#fff;color:#000;cursor:pointer}" +
    ".brks .sp{flex:1}" +
    "@media (max-width:640px){.brks{padding:18px 16px}.brks .sp{flex-basis:100%}}";

  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function pct(a, b) { if (!a || !b) return null; return (a / b - 1) * 100; }
  function pctHtml(v) { if (v === null || !isFinite(v)) return "—"; return "<span class='" + (v >= 0 ? "up" : "dn") + "'>" + (v >= 0 ? "+" : "") + v.toFixed(1) + "%</span>"; }
  function fmtTime(iso) {
    if (!iso) return "";
    var d = new Date(iso); if (isNaN(d)) return String(iso).slice(0, 16);
    return (d.getMonth() + 1) + "/" + d.getDate() + " " + pad(d.getHours()) + ":" + pad(d.getMinutes());
  }
  function rangeText(a) {
    if (!a) return "—";
    var m = +a.key.slice(5, 7);
    if (a.closed) return m + "월 마감";
    return m + "/1~" + a.last + (a.full ? " (마감)" : "") + (a.gaps ? " · 빠진 날 있음" : "");
  }

  function mountUploader(box, opt) {
    opt = opt || {};
    if (!box) return;
    if (!document.getElementById("brksCss")) {
      var st = document.createElement("style"); st.id = "brksCss"; st.textContent = CSS; document.head.appendChild(st);
    }
    box.classList.add("brks");
    box.innerHTML =
      "<h3>매출 엑셀 올리기</h3>" +
      "<p class='lead2'>BR 일매출 엑셀을 넣으면 이 페이지와 <a href='/sales/'>총 POS 매출</a> 페이지의 매출·고객수·그래프가 함께 바뀝니다. 같은 날짜를 다시 올리면 새 값으로 바뀝니다.</p>" +
      "<div class='now' id='brksNow'>지금 들어 있는 매출을 확인하는 중…</div>" +
      "<label class='drop' id='brksDrop' tabindex='0' role='button' aria-label='엑셀 파일 고르기'>" +
        "<b>엑셀 파일을 여기에 끌어다 놓거나 눌러서 고르세요</b>" +
        "<span>.xlsx · .xls · .csv · 여러 개를 한 번에 넣어도 됩니다</span>" +
        "<input type='file' id='brksFile' accept='.xlsx,.xls,.xlsm,.csv,.tsv,.txt' multiple hidden>" +
      "</label>" +
      "<div class='row'><button type='button' class='lnk' id='brksTpl'>양식 파일 받기</button>" +
        "<span style='font-size:12px;color:#8A7A70'>MKT 일매출 엑셀을 그대로 넣어도 제목 줄을 보고 칸을 찾습니다.</span></div>" +
      "<div class='msg' id='brksMsg' role='status' aria-live='polite'></div>" +
      "<div class='pv' id='brksPv' hidden></div>";

    var $ = function (id) { return document.getElementById(id); };
    var drop = $("brksDrop"), file = $("brksFile"), msg = $("brksMsg"), pv = $("brksPv");
    var parsed = null, model = null;

    function say(t, cls) { msg.className = "msg" + (cls ? " " + cls : ""); msg.innerHTML = t; }

    function showNow() {
      load().then(function (M) {
        model = M;
        var ks = M.keys(), lastK = ks[ks.length - 1], a = lastK ? M.month(lastK) : null;
        var html = "<span>들어 있는 마지막 달 <b>" + (a ? lastK.slice(0, 4) + "년 " + rangeText(a) : "—") + "</b></span>";
        if (M.sheetOk) html += "<span>엑셀로 올린 날 <b>" + M.sheetRows + "</b>건" + (M.lastUpload ? " · 마지막 " + fmtTime(M.lastUpload) + (M.lastBy ? " " + esc(M.lastBy) : "") : "") + "</span>";
        else html += "<span class='bad'>구글시트에 연결하지 못했습니다 (" + esc(M.sheetErr || "") + ") — Apps Script 새 버전 배포가 필요할 수 있습니다</span>";
        $("brksNow").innerHTML = html;
      });
    }
    showNow();

    function pick(files) {
      if (!files || !files.length) return;
      pv.hidden = true; parsed = null;
      say("읽는 중…");
      parseFiles(files).then(function (R) {
        parsed = R; say(""); preview();
      }).catch(function (e) { say(esc(e.message || e), "err"); });
    }

    function preview() {
      var R = parsed, unit = +($("brksUnit") && $("brksUnit").value) || R.unitGuess;
      /* 달별로 묶기 */
      var tmp = { days: {}, mon: {} };
      Object.keys(R.days).forEach(function (k) { var v = R.days[k]; tmp.days[k] = { s: v.s * unit, c: v.c, pw: v.pw * unit, pwc: v.pwc, cats: v.cats, src: "sheet" }; });
      Object.keys(R.mon).forEach(function (k) { var v = R.mon[k]; tmp.mon[k] = { s: v.s * unit, c: v.c, pw: v.pw * unit, pwc: v.pwc, cats: null }; });
      var mkeys = {}; Object.keys(tmp.days).forEach(function (k) { mkeys[k.slice(0, 7)] = 1; }); Object.keys(tmp.mon).forEach(function (k) { mkeys[k] = 1; });
      /* 이미 들어 있는 날과 합쳐서 '올린 뒤 모습'을 보여 준다 */
      var after = { days: {}, mon: {} };
      if (model) Object.keys(model.days).forEach(function (d) { if (mkeys[d.slice(0, 7)]) after.days[d] = model.days[d]; });
      Object.keys(tmp.days).forEach(function (d) { after.days[d] = tmp.days[d]; });
      Object.keys(tmp.mon).forEach(function (k) { after.mon[k] = tmp.mon[k]; });

      var rows = Object.keys(mkeys).sort().map(function (key) {
        var a = monthOf(after, key);
        /* 지금 사이트에 있는 값과 비교 */
        var old = null, state = "new";
        if (model) {
          var ks = Object.keys(tmp.days).filter(function (d) { return d.slice(0, 7) === key; });
          var diff = 0, same = 0, fresh = 0;
          ks.forEach(function (d) {
            var o = model.days[d];
            if (!o) fresh++; else if (Math.abs(o.s - tmp.days[d].s) < 1) same++; else diff++;
          });
          if (tmp.mon[key]) { var om = model.mon[key]; if (!om) fresh++; else if (Math.abs(om.s - tmp.mon[key].s) < 1) same++; else diff++; }
          state = diff ? "chg" : (fresh ? "new" : "same");
          old = { diff: diff, fresh: fresh, same: same };
        }
        var yoyTxt = a.pw ? pctHtml(pct(a.s, a.pw)) : "—";
        var tag = state === "new" ? "<span class='tag new'>새로 " + (old ? old.fresh : "") + "일</span>"
                : state === "chg" ? "<span class='tag chg'>" + old.diff + "일 숫자 바뀜</span>" + (old.fresh ? " <span class='tag new'>새로 " + old.fresh + "일</span>" : "")
                : "<span class='tag same'>그대로</span>";
        if (tmp.mon[key] && !Object.keys(tmp.days).some(function (d) { return d.slice(0, 7) === key; })) tag = state === "same" ? "<span class='tag same'>그대로</span>" : "<span class='tag " + state + "'>월 합계</span>";
        return "<tr><td><b>" + key.slice(0, 4) + "." + key.slice(5) + "</b></td><td>" + rangeText(a) + "</td>" +
               "<td class='n'><b>" + eok(a.s) + "억</b></td><td class='n'>" + (a.c ? comma(a.c) + "명" : "—") + "</td>" +
               "<td class='n'>" + yoyTxt + "</td><td>" + tag + "</td></tr>";
      });

      var catSet = {};
      Object.keys(R.days).forEach(function (k) { var c = R.days[k].cats; if (c) for (var x in c) catSet[x] = 1; });
      var warn = [];
      Object.keys(mkeys).forEach(function (key) { var a = monthOf(after, key); if (a && a.gaps) warn.push(key.slice(5) + "월에 빠진 날짜가 있습니다 — 빠진 날은 기존 값이 있으면 그대로 두고, 없으면 비어 있게 됩니다."); });
      if (!Object.keys(R.days).some(function (k) { return R.days[k].c; }) && Object.keys(R.days).length) warn.push("고객수 칸을 찾지 못했습니다. 매출만 바뀝니다.");
      if (!Object.keys(R.days).some(function (k) { return R.days[k].pw; }) && Object.keys(R.days).length) warn.push("전년 동요일 칸을 찾지 못했습니다. 전년비는 월 합계끼리 비교합니다.");

      var units = [[1, "원"], [1e3, "천원"], [1e4, "만원"], [1e6, "백만원"], [1e8, "억원"]];
      pv.innerHTML =
        "<h4>읽은 내용 확인 <span style='font-weight:500;font-size:12px;color:#8A7A70'>— 이미 들어 있는 날과 합친, 올린 뒤의 달 합계입니다</span></h4>" +
        "<div class='tw'><table><thead><tr><th>달</th><th>들어온 날짜</th><th style='text-align:right'>매출</th><th style='text-align:right'>고객수</th><th style='text-align:right'>전년 동요일 대비</th><th>바뀌는 점</th></tr></thead><tbody>" +
        rows.join("") + "</tbody></table></div>" +
        "<ul class='map'>" + R.sheets.map(function (s) {
          return "<li><b>" + esc(s.file) + " › " + esc(s.sheet) + "</b> (" + s.mode + ", " + (s.days ? s.days + "일" : s.months + "달") + ") — " + esc(s.map.join(" · ") || "칸 이름 없음") + "</li>";
        }).join("") + (Object.keys(catSet).length ? "<li><b>제품군</b> — " + esc(Object.keys(catSet).join(", ")) + "</li>" : "") + "</ul>" +
        (warn.length ? "<div class='warn'>" + warn.map(esc).join("<br>") + "</div>" : "") +
        "<div class='act'>" +
          "<label class='u'>금액 단위 <select id='brksUnit'>" + units.map(function (u) {
            return "<option value='" + u[0] + "'" + (u[0] === unit ? " selected" : "") + ">" + u[1] + (u[0] === R.unitGuess && !R.unitFound ? " (짐작)" : "") + "</option>";
          }).join("") + "</select></label>" +
          "<span class='sp'></span>" +
          "<input type='text' id='brksBy' placeholder='올린 사람' aria-label='올린 사람 이름' autocomplete='name'>" +
          "<input type='password' id='brksKey' placeholder='올리기 비밀번호' aria-label='올리기 비밀번호' autocomplete='off'>" +
          "<button type='button' class='ghost' id='brksCancel'>취소</button>" +
          "<button type='button' class='go' id='brksGo'>사이트에 반영하기</button>" +
        "</div>";
      pv.hidden = false;
      try { var nm = localStorage.getItem("brks_by"); if (nm) $("brksBy").value = nm; } catch (e) {}
      $("brksUnit").onchange = preview;
      $("brksCancel").onclick = function () { parsed = null; pv.hidden = true; say(""); };
      $("brksKey").onkeydown = function (e) { if (e.key === "Enter") $("brksGo").click(); };
      $("brksGo").onclick = go;
    }

    function go() {
      var key = $("brksKey").value.trim(), by = $("brksBy").value.trim();
      if (!key) { say("올리기 비밀번호를 넣어 주세요.", "err"); $("brksKey").focus(); return; }
      try { if (by) localStorage.setItem("brks_by", by); } catch (e) {}
      var unit = +$("brksUnit").value || 1;
      var recs = toRecords(parsed, unit);
      var btn = $("brksGo"); btn.disabled = true; btn.textContent = "올리는 중…";
      var before = model ? model.sheetRows : 0;
      upload(recs, key, by).then(function (res) {
        return load({ fresh: true }).then(function (M) {
          /* 답을 못 받고 보낸 경우 : 다시 읽어 실제로 들어갔는지 확인 */
          if (res.blind) {
            var hit = recs.filter(function (r) {
              var v = r.kind === "d" ? M.days[r.id.slice(1)] : M.mon[r.id.slice(1)];
              return v && v.src !== "base" && Math.abs(v.s - r.s) < 1;
            }).length;
            if (hit < recs.length * 0.9) throw new Error("시트에 들어가지 않았습니다. 비밀번호를 확인하거나, Apps Script 를 새 버전으로 배포했는지 확인해 주세요.");
          }
          model = M;
          var ks = Object.keys(parsed.days).map(function (k) { return k.slice(0, 7); }).concat(Object.keys(parsed.mon));
          ks = ks.filter(function (k, i) { return ks.indexOf(k) === i; }).sort();
          say("반영했습니다 — " + ks.map(function (k) { return k.slice(0, 4) + "." + k.slice(5); }).join(", ") + " · " + recs.length + "건. 매출 페이지도 새로 열면 같은 숫자가 나옵니다.", "ok");
          pv.hidden = true; parsed = null;
          showNow();
          if (opt.onApplied) opt.onApplied(M);
        });
      }).catch(function (e) {
        say(esc(e.message || e), "err");
        btn.disabled = false; btn.textContent = "사이트에 반영하기";
      });
    }

    drop.addEventListener("click", function (e) { if (e.target !== file) file.click(); e.preventDefault(); });
    drop.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); file.click(); } });
    file.addEventListener("change", function () { pick(file.files); file.value = ""; });
    ["dragenter", "dragover"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("on"); }); });
    ["dragleave", "drop"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove("on"); }); });
    drop.addEventListener("drop", function (e) { pick(e.dataTransfer && e.dataTransfer.files); });
    $("brksTpl").addEventListener("click", function () { template().catch(function (e) { say(esc(e.message || e), "err"); }); });
  }

  global.BRKSales = {
    load: load, parseFiles: parseFiles, toRecords: toRecords, upload: upload,
    mountUploader: mountUploader, template: template,
    util: { pad: pad, dim: dim, eok: eok, comma: comma, monthOf: monthOf, parseSheet: parseSheet, toDay: toDay }
  };
})(window);
