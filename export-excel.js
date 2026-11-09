/* STATUS SEAT PLAN A320-03 — EXPORT TO EXCEL
 * buildSeatWorkbook(ExcelJS, data) returns an ExcelJS workbook with 5 sheets (same as A340):
 *   ผังที่นั่ง | Galley และห้องน้ำ | รวมข้อขัดข้องคงค้าง | ทุกที่นั่ง และ Monitor | ประวัติการแก้ไข
 * (old note:
 *   ผังที่นั่ง | รวมข้อขัดข้องคงค้าง | ทุกที่นั่ง และ Monitor | ประวัติการแก้ไข
 * data = {
 *   state:    { id: {s:0|1|2, m:0|1|2, note:'', by:'', at:ISO} }   id = 'CC1A', '12E', 'MON-D12R' ...
 *   lastSave: { by, at:ISO } | null,
 *   inspector: '',
 *   history:  [[atISO|text, email, inspector, label, field, oldValue, newValue], ...] | null
 * }
 * The "ผังที่นั่ง" sheet is drawn from planSheet() in layout.js, the same plan the Apps Script uses
 * to build the Google Sheet, so both look the same.
 */
(function (root) {
  var FONT = 'Tahoma';
  var thin = { style: 'thin', color: { argb: 'FF000000' } };
  var BOX = { top: thin, left: thin, bottom: thin, right: thin };

  function argb(hex) { return 'FF' + String(hex || '#FFFFFF').replace('#', '').toUpperCase(); }
  function fill(hex) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(hex) } }; }
  function font(o) {
    o = o || {};
    var f = { name: FONT, size: o.sz || 9, color: { argb: argb(o.color || '#000000') } };
    if (o.b) f.bold = true; if (o.i) f.italic = true; return f;
  }
  /* Excel stores dates without a time zone: shift so the cell shows Thailand time (UTC+7) */
  function th(d) { if (!(d instanceof Date)) d = new Date(d); if (isNaN(d)) return null; return new Date(d.getTime() + 7 * 3600e3); }
  function thText(d) { /* 14 OCT26 (10:52 TH-LT) */
    var t = th(d); if (!t) return '-';
    var mon = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'][t.getUTCMonth()];
    var p = function (n) { return String(n).padStart(2, '0'); };
    return p(t.getUTCDate()) + ' ' + mon + p(t.getUTCFullYear() % 100) + ' (' + p(t.getUTCHours()) + ':' + p(t.getUTCMinutes()) + ' TH-LT)';
  }

  function buildSeatWorkbook(ExcelJS, data) {
    data = data || {};
    var state = data.state || {};
    var L = root.seatPlanLayout(), P = root.planSheet(L), S = root.SEAT_S, M = root.SEAT_M, E = root.SEAT_E, C = root.SEAT_COLOR;
    var wb = new ExcelJS.Workbook();
    wb.creator = 'STATUS SEAT PLAN A320-03';
    wb.calcProperties = { fullCalcOnLoad: true };
    var ws = wb.addWorksheet('ผังที่นั่ง', {
      views: [{ showGridLines: false }],
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    /* Excel width is in characters: a character is 7 px plus 5 px cell padding, so (px - 5) / 7 gives the same pixel width as the Google Sheet */
    Object.keys(P.widths).forEach(function (c) { ws.getColumn(Number(c)).width = Math.max(1, (P.widths[c] - 5) / 7); });
    Object.keys(P.heights).forEach(function (r) { ws.getRow(Number(r)).height = P.heights[r] * 0.75; });

    function count(q) {
      var n = 0;
      q.ids.forEach(function (id) { var d = state[id] || {}; if (q.vals.indexOf((d[q.f] | 0)) >= 0) n++; });
      return n;
    }
    P.ops.forEach(function (o) {
      var r2 = o.r + o.rs - 1, c2 = o.c + o.cs - 1, rr, cc;
      if (o.cabin) {
        for (rr = o.r; rr <= r2; rr++) for (cc = o.c; cc <= c2; cc++) ws.getCell(rr, cc).fill = fill(o.fill);
        return;
      }
      if (o.rs > 1 || o.cs > 1) ws.mergeCells(o.r, o.c, r2, c2);
      var cell = ws.getCell(o.r, o.c);
      if (o.v !== undefined && o.v !== null) {
        cell.value = String(o.v).charAt(0) === '=' ? { formula: o.v.slice(1), result: o.q ? count(o.q) : undefined } : o.v;
      }
      cell.font = font(o.font);
      var a = o.align || {};
      cell.alignment = { horizontal: a.h, vertical: a.v || 'middle', wrapText: !!a.wrap };
      if (o.fmt) cell.numFmt = o.fmt;
      for (rr = o.r; rr <= r2; rr++) for (cc = o.c; cc <= c2; cc++) {
        var x = ws.getCell(rr, cc);
        if (o.fill) x.fill = fill(o.fill);
        if (o.border) x.border = BOX;
      }
      if (o.dv) cell.dataValidation = { type: 'list', allowBlank: true, formulae: ['"' + (o.dv === 'M' ? M : o.dv === 'E' ? E : S).join(',') + '"'] };
    });
    var med = { style: 'medium', color: { argb: argb(C.MUTED) } };
    P.boxes.forEach(function (b) {
      for (var rr = b.r1; rr <= b.r2; rr++) for (var cc = b.c1; cc <= b.c2; cc++) {
        if (rr !== b.r1 && rr !== b.r2 && cc !== b.c1 && cc !== b.c2) continue;
        var x = ws.getCell(rr, cc), bd = Object.assign({}, x.border || {});
        if (rr === b.r1) bd.top = med; if (rr === b.r2) bd.bottom = med; if (cc === b.c1) bd.left = med; if (cc === b.c2) bd.right = med;
        x.border = bd;
      }
    });
    var rules = P.cf.map(function (c, i) {
      var st = { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(c.fill) } } };
      if (c.font) st.font = { color: { argb: argb(c.font) } };
      return c.text !== undefined
        ? { type: 'cellIs', operator: 'equal', priority: i + 1, formulae: ['"' + c.text + '"'], style: st }
        : { type: 'expression', priority: i + 1, formulae: [c.f.slice(1)], style: st };
    });
    ws.addConditionalFormatting({ ref: P.cf[0].range, rules: rules });

    /* header and current state */
    var ls = data.lastSave || {};
    ws.getCell('D2').value = ls.by || '-';
    ws.getCell('M2').value = ls.at ? thText(ls.at) : '-';
    if (data.inspector) ws.getCell('D3').value = data.inspector;
    ws.getCell('B4').value = 'บันทึกสภาพที่นั่ง, Monitor, Galley และห้องน้ำ · Export จากหน้าเว็บ ' + thText(new Date()) + ' · เวลาทั้งหมดเป็นเวลาประเทศไทย (TH-LT)';
    L.order.forEach(function (id) {
      var it = L.items[id], p = P.pos[id], d = state[id]; if (!d || !p) return;
      ws.getCell(p.status[0], p.status[1]).value = it.kind === 'mon' ? M[d.m | 0] : (it.kind === 'eq' || it.kind === 'lav') ? E[d.s | 0] : S[d.s | 0];
      if (d.note) ws.getCell(p.note[0], p.note[1]).value = d.note;
    });

    var ZN = {}; root.SEAT_ZONES.forEach(function (z) { ZN[z.key] = z.name; });
    /* ---------- Galley และห้องน้ำ (as A340): summary tables (COUNTIFS over the list below) + the list linked to the map ---------- */
    var isEq = function (id) { var k = L.items[id].kind; return k === 'eq' || k === 'lav'; };
    var eqIds = L.order.filter(isEq), MS = "'ผังที่นั่ง'!";
    var eqGroup = function (it) { return it.kind === 'eq' ? it.group : ZN[it.zone]; };
    var lavZones = root.SEAT_ZONES.filter(function (z) { return eqIds.some(function (id) { var it = L.items[id]; return it.kind === 'lav' && it.zone === z.key; }); });
    var galleys = Object.keys(root.SEAT_GAL).map(function (k) {
      return { key: k, name: root.SEAT_GAL[k], types: root.SEAT_ETYPE_ORDER.filter(function (t) { return eqIds.some(function (id) { var it = L.items[id]; return it.gal === k && it.type === t; }); }) };
    });
    var g = wb.addWorksheet('Galley และห้องน้ำ', { views: [{ showGridLines: false }], pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    [30, 18, 18, 18, 40, 22, 20].forEach(function (w, j) { g.getColumn(j + 1).width = w; });
    function gput(r, c, v, o) {
      o = o || {}; var cell = g.getCell(r, c);
      if (v !== undefined && v !== null) cell.value = v;
      cell.font = o.font || font(); cell.alignment = o.align || { horizontal: 'center', vertical: 'middle', wrapText: true }; cell.border = BOX;
      if (o.fill) cell.fill = fill(o.fill); if (o.fmt) cell.numFmt = o.fmt;
      return cell;
    }
    g.mergeCells('A1:G1'); g.getCell('A1').value = 'สถานภาพ Galley และห้องน้ำ (LAVATORY) A320-03';
    g.getCell('A1').font = font({ sz: 14, b: true }); g.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }; g.getCell('A1').fill = fill(C.TAG); g.getRow(1).height = 28;
    g.mergeCells('A2:G2'); g.getCell('A2').value = 'แก้สภาพ Galley / ห้องน้ำได้ที่แผ่นงาน "ผังที่นั่ง" (dropdown ใต้ชื่ออุปกรณ์) · รายการและตารางสรุปในแผ่นนี้ดึงค่ามาเอง · Export ' + thText(new Date());
    g.getCell('A2').font = font({ sz: 8, i: true }); g.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    var nGalRows = 0; galleys.forEach(function (G) { nGalRows += 1 + G.types.length; });
    var GS = 4, GE = GS + nGalRows + 1, LS = GE + 2, LE = LS + lavZones.length + 1, L0 = LE + 2, L1 = L0 + eqIds.length;
    var rng = function (c) { return '$' + c + '$' + (L0 + 1) + ':$' + c + '$' + Math.max(L1, L0 + 1); };
    var rB = rng('B'), rC = rng('C'), rD = rng('D');
    function q(t) { return '"' + String(t).replace(/"/g, '""') + '"'; }
    function cnt(f) { var c = { t: 0, s: [0, 0, 0] }; eqIds.forEach(function (id) { var it = L.items[id]; if (!f(it)) return; var d = state[id] || {}; c.t++; c.s[d.s | 0]++; }); return c; }
    function sumHeader(r, first) {
      [first, 'จำนวน', E[0], E[1], E[2]].forEach(function (t, j) { gput(r, j + 1, t, { font: font({ sz: 9, b: true }), fill: [C.ZONEH, C.ZONEH, C.GREEN, C.YEL, C.RED][j] }); });
      g.getRow(r).height = 22;
    }
    function sumRow(r, name, crit, c, o) {
      o = o || {};
      gput(r, 1, name, { font: font({ b: !!o.bold }), align: { vertical: 'middle', horizontal: 'left', indent: o.indent ? 2 : 0 }, fill: o.fill });
      gput(r, 2, { formula: 'COUNTIFS(' + crit + ')', result: c.t }, { font: font({ b: !!o.bold }), fill: o.fill });
      E.forEach(function (t, j) { gput(r, 3 + j, { formula: 'COUNTIFS(' + crit + ',' + rD + ',' + q(t) + ')', result: c.s[j] }, { font: font({ b: !!o.bold }), fill: o.fill }); });
    }
    function totRow(r, rows, c) {
      gput(r, 1, 'รวม', { font: font({ b: true }), align: { vertical: 'middle', horizontal: 'left' }, fill: C.GREY });
      [c.t, c.s[0], c.s[1], c.s[2]].forEach(function (v, j) {
        var Lc = col(2 + j); gput(r, 2 + j, { formula: rows.map(function (x) { return Lc + x; }).join('+') || '0', result: v }, { font: font({ b: true }), fill: C.GREY });
      });
    }
    sumHeader(GS, 'GALLEY / อุปกรณ์');
    var rr2 = GS + 1, galHeads = [];
    galleys.forEach(function (G) {
      galHeads.push(rr2);
      sumRow(rr2++, G.name, rB + ',' + q(G.name), cnt(function (it) { return it.kind === 'eq' && it.gal === G.key; }), { bold: true, fill: C.CABIN });
      G.types.forEach(function (t) {
        sumRow(rr2++, root.SEAT_ETYPE[t], rB + ',' + q(G.name) + ',' + rC + ',' + q(root.SEAT_ETYPE[t]), cnt(function (it) { return it.kind === 'eq' && it.gal === G.key && it.type === t; }), { indent: true });
      });
    });
    totRow(GE, galHeads, cnt(function (it) { return it.kind === 'eq'; }));
    sumHeader(LS, 'ห้องน้ำ (LAVATORY) · ZONE');
    var lavRows = [];
    lavZones.forEach(function (z, i) {
      var r = LS + 1 + i; lavRows.push(r);
      var names = eqIds.filter(function (id) { var it = L.items[id]; return it.kind === 'lav' && it.zone === z.key; }).map(function (id) { return L.items[id].label; });
      sumRow(r, z.name + ' (' + names.join(', ') + ')', rB + ',' + q(z.name) + ',' + rC + ',' + q('LAVATORY'), cnt(function (it) { return it.kind === 'lav' && it.zone === z.key; }));
    });
    totRow(LE, lavRows, cnt(function (it) { return it.kind === 'lav'; }));
    ['รายการ', 'GALLEY / โซน', 'ประเภท', 'สภาพ', 'หมายเหตุ', 'ผู้บันทึกล่าสุด', 'วัน/เวลาล่าสุด (TH-LT)'].forEach(function (t, j) {
      gput(L0, j + 1, t, { font: font({ sz: 9, b: true }), fill: C.ZONEH });
    });
    g.getRow(L0).height = 22;
    eqIds.forEach(function (id, i) {
      var r = L0 + 1 + i, it = L.items[id], d = state[id] || {}, p = P.pos[id];
      var stA = col(p.status[1]) + p.status[0], ntA = col(p.note[1]) + p.note[0];
      [it.full || it.label, eqGroup(it), it.tname, { formula: MS + stA, result: E[d.s | 0] }, { formula: 'IF(' + MS + ntA + '="","",' + MS + ntA + ')', result: d.note || '' }, d.by || '', d.at ? th(d.at) : ''].forEach(function (v, j) {
        gput(r, j + 1, v, { align: { horizontal: (j === 0 || j === 4 || j === 5) ? 'left' : 'center', vertical: 'top', wrapText: j === 4 } });
      });
      g.getCell(r, 7).numFmt = 'dd/mm/yyyy hh:mm';
    });
    if (eqIds.length) {
      g.addConditionalFormatting({ ref: 'D' + (L0 + 1) + ':D' + L1, rules: [
        { type: 'cellIs', operator: 'equal', priority: 1, formulae: ['"' + E[1] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(C.YEL) } } } },
        { type: 'cellIs', operator: 'equal', priority: 2, formulae: ['"' + E[2] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(C.RED) } }, font: { color: { argb: 'FFFFFFFF' } } } }
      ] });
      g.autoFilter = 'A' + L0 + ':G' + L1;
    }

    /* list sheets */
    var HDR = ['ที่นั่ง / จุด', 'ประเภท', 'โซน', 'สภาพที่นั่ง / อุปกรณ์', 'สภาพ Monitor', 'หมายเหตุ', 'ผู้บันทึกล่าสุด', 'วัน/เวลาล่าสุด (TH-LT)'];
    var WID = [52, 10, 18, 16, 14, 44, 24, 20];
    var CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };
    function col(n) { return root.seatColName(n); }
    function statusCF(s, end) {
      if (end < 2) return;
      var red = { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(C.RED) } }, font: { color: { argb: 'FFFFFFFF' } } };
      var yel = { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(C.YEL) } } };
      s.addConditionalFormatting({ ref: 'D2:E' + end, rules: [
        { type: 'cellIs', operator: 'equal', priority: 1, formulae: ['"' + S[1] + '"'], style: yel },
        { type: 'cellIs', operator: 'equal', priority: 2, formulae: ['"' + M[1] + '"'], style: yel },
        { type: 'cellIs', operator: 'equal', priority: 3, formulae: ['"' + S[2] + '"'], style: red },
        { type: 'cellIs', operator: 'equal', priority: 4, formulae: ['"' + M[2] + '"'], style: red },
        { type: 'cellIs', operator: 'equal', priority: 5, formulae: ['"' + E[1] + '"'], style: yel },
        { type: 'cellIs', operator: 'equal', priority: 6, formulae: ['"' + E[2] + '"'], style: red }
      ] });
    }
    function header(s, titles, widths) {
      titles.forEach(function (t, j) {
        var c = s.getCell(1, j + 1); c.value = t; c.font = font({ sz: 10, b: true }); c.fill = fill(C.ZONEH); c.border = BOX; c.alignment = CENTER;
        s.getColumn(j + 1).width = widths[j];
      });
      s.views = [{ state: 'frozen', ySplit: 1 }];
    }
    function listSheet(title, ids, live) {
      var s = wb.addWorksheet(title); header(s, HDR, WID);
      var q = "'ผังที่นั่ง'!";
      ids.forEach(function (id, i) {
        var row = 2 + i, it = L.items[id], p = P.pos[id], d = state[id] || {}, isMon = it.kind === 'mon';
        var st = col(p.status[1]) + p.status[0], nt = col(p.note[1]) + p.note[0];
        var vals = [it.name, isMon ? 'Monitor' : it.kind === 'eq' ? 'Galley' : it.kind === 'lav' ? 'ห้องน้ำ' : 'ที่นั่ง', it.kind === 'eq' ? it.group : ZN[it.zone]];
        if (isEq(id)) {
          vals.push(E[d.s | 0], '-', d.note || '');
        } else if (live) {
          vals.push(isMon ? '-' : { formula: q + st, result: S[d.s | 0] });
          vals.push(isMon ? { formula: q + st, result: M[d.m | 0] } : '-');
          vals.push({ formula: 'IF(' + q + nt + '="",""' + ',' + q + nt + ')', result: d.note || '' });
        } else {
          vals.push(isMon ? '-' : S[d.s | 0], isMon ? M[d.m | 0] : '-', d.note || '');
        }
        vals.push(d.by || '', d.at ? th(d.at) : '');
        vals.forEach(function (v, j) {
          var c = s.getCell(row, j + 1); c.value = v; c.font = font(); c.border = BOX;
          c.alignment = { horizontal: (j === 0 || j === 5 || j === 6) ? 'left' : 'center', vertical: 'top', wrapText: j === 0 || j === 5 };
        });
        s.getCell(row, 8).numFmt = 'dd/mm/yyyy hh:mm';
      });
      var end = 1 + ids.length; statusCF(s, end);
      if (end >= 2) s.autoFilter = 'A1:H' + end;
      return s;
    }
    var bad = L.order.filter(function (id) { var d = state[id] || {}; return (d.s | 0) || (d.m | 0) || d.note; });
    var sBad = listSheet('รวมข้อขัดข้องคงค้าง', bad, false);
    if (!bad.length) { sBad.getCell('A2').value = 'ไม่มีข้อขัดข้องคงค้าง'; sBad.getCell('A2').font = font(); }
    listSheet('ทุกที่นั่ง และ Monitor', L.order.filter(function (id) { return !isEq(id); }), true);

    /* history: every cell centred (horizontal + vertical) */
    var h = wb.addWorksheet('ประวัติการแก้ไข');
    header(h, ['วัน/เวลา (TH-LT)', 'อีเมลผู้แก้ไข', 'ผู้ตรวจ', 'ที่นั่ง / จุด', 'รายการ', 'ค่าเดิม', 'ค่าใหม่'], [20, 30, 24, 52, 14, 36, 36]);
    var hist = data.history;
    if (!hist) {
      h.mergeCells('A2:G2'); h.getCell('A2').value = 'ไม่สามารถดึงประวัติการแก้ไขได้ ดูในแท็บ "ประวัติการแก้ไข" ของ Google Sheet';
      h.getCell('A2').font = font(); h.getCell('A2').alignment = CENTER;
    } else {
      hist.forEach(function (row, i) {
        row.forEach(function (v, j) {
          var c = h.getCell(2 + i, j + 1);
          if (j === 0) { var t = (typeof v === 'string' && /^\d{4}-\d\d-\d\dT/.test(v)) ? th(v) : null; c.value = t || String(v || ''); if (t) c.numFmt = 'dd/mm/yyyy hh:mm:ss'; }
          else c.value = v === null || v === undefined ? '' : String(v);
          c.font = font(); c.border = BOX; c.alignment = CENTER;
        });
      });
      if (hist.length) h.autoFilter = 'A1:G' + (1 + hist.length);
    }
    return wb;
  }

  root.buildSeatWorkbook = buildSeatWorkbook;
  root.seatThText = thText;
})(typeof window !== 'undefined' ? window : this);
