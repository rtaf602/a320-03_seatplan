/* STATUS SEAT PLAN A320(03) — EXPORT TO EXCEL
 * buildSeatWorkbook(ExcelJS, data) returns an ExcelJS workbook with 4 sheets:
 *   ผังที่นั่ง | รวมข้อขัดข้องคงค้าง | ทุกที่นั่ง และ Monitor | ประวัติการแก้ไข
 * data = {
 *   state:   { id: {s:0|1|2, m:0|1|2, note:'', by:'', at:ISO} }   id = 'CC1A', '12E', 'MON-D12L' ...
 *   lastSave:{ by, at:ISO } | null,
 *   inspector: '',
 *   history: [[atISO|text, email, inspector, label, field, oldValue, newValue], ...] | null
 * }
 * The "ผังที่นั่ง" layout here is the same as the tab in the Google Sheet (two columns per seat),
 * and wb.seatCells / wb.seatOrder give every seat and Monitor with its cells.
 * apps-script/Code.gs carries a copy of this file (tools/build-code-gs.js) and builds the Google Sheet tab from it,
 * so after changing the layout here run:  node tools/build-code-gs.js
 */
(function (root) {
  var S = ['ปกติ', 'ขัดข้อง-ใช้ได้', 'ใช้ไม่ได้'];
  var M = ['จอปกติ', 'จอเสื่อม', 'จอเสีย'];
  var C = { GREEN: 'FF00FF00', MGREEN: 'FF8ED873', YEL: 'FFFFFF00', RED: 'FFFF0000', CABIN: 'FFCAEDFB', ROWNO: 'FF6D9EEB',
    ZONEH: 'FFC1E4F5', TAG: 'FF45B0E1', WHITE: 'FFFFFFFF', BLUE: 'FF00B0F0', GREY: 'FFF2F2F2', MUTED: 'FF566575',
    OVHD: 'FFB4A7D6', FURN: 'FFD9D9D9' };
  var ZN = { F: 'FRONT CREW ZONE', A: 'VVIP OA', B: 'VIP OB', C: 'VVIP OC', D: 'PASSENGER OD' };
  var ZKEYS = ['F', 'A', 'B', 'C', 'D'];
  var MTYPE = { W: 'Monitor', R: 'OVHD Monitor' };
  var MTYPE_CELL = { W: 'Monitor 21.5"', R: 'OVHD 12.0"' }; /* label on the plan, with screen size */
  var MTYPE_FULL = { W: 'Wall Mounted Monitor 21.5 Inch', R: 'Bin Mounted Retractable Monitors (OVHD Monitor) 12.0 Inch' };
  var LIST = 'ทุกที่นั่ง และ Monitor';
  var FONT = 'Tahoma';
  var CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };
  var thin = { style: 'thin', color: { argb: 'FF000000' } };
  var BOX = { top: thin, left: thin, bottom: thin, right: thin };
  /* seat columns: left A B C, right D E F (each seat spans 2 columns) */
  var COLS = { A: 3, B: 5, C: 7, D: 14, E: 16, F: 18 };
  var MON_L = 4, MON_R = 15; /* Monitor cells span 4 columns, centred over each side */

  function col(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
  function fill(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }
  function font(o) { var f = { name: FONT, size: 9, color: { argb: 'FF000000' } }; for (var k in (o || {})) f[k] = o[k]; return f; }

  /* Excel stores dates without a time zone: shift so the cell shows Thailand time (UTC+7) */
  function th(d) { if (!(d instanceof Date)) d = new Date(d); if (isNaN(d)) return null; return new Date(d.getTime() + 7 * 3600e3); }
  function thText(d) { /* 13 OCT26 (15:49 TH-LT) */
    var t = th(d); if (!t) return '-';
    var mon = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'][t.getUTCMonth()];
    var p = function (n) { return String(n).padStart(2, '0'); };
    return p(t.getUTCDate()) + ' ' + mon + p(t.getUTCFullYear() % 100) + ' (' + p(t.getUTCHours()) + ':' + p(t.getUTCMinutes()) + ' TH-LT)';
  }

  function buildSeatWorkbook(ExcelJS, data) {
    data = data || {};
    var state = data.state || {};
    var wb = new ExcelJS.Workbook();
    wb.creator = 'STATUS SEAT PLAN A320(03)';
    wb.calcProperties = { fullCalcOnLoad: true };
    var ws = wb.addWorksheet('ผังที่นั่ง', {
      views: [{ showGridLines: false }],
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    ws.getColumn(1).width = 2; ws.getColumn(2).width = 8;
    for (var c = 3; c <= 20; c++) ws.getColumn(c).width = 6;
    for (c = 9; c <= 13; c++) ws.getColumn(c).width = 3.5;
    ws.getColumn(20).width = 2.5;

    function put(r, c, v, o) {
      o = o || {};
      var span = o.span === undefined ? 2 : o.span, rows = o.rows || 1;
      if (span > 1 || rows > 1) ws.mergeCells(r, c, r + rows - 1, c + span - 1);
      var cell = ws.getCell(r, c);
      if (v !== undefined && v !== null) cell.value = v;
      cell.font = o.font || font();
      cell.alignment = o.align || CENTER;
      if (o.fmt) cell.numFmt = o.fmt;
      for (var rr = r; rr < r + rows; rr++) for (var cc = c; cc < c + span; cc++) {
        var x = ws.getCell(rr, cc);
        if (o.fill) x.fill = fill(o.fill);
        if (o.border !== false) x.border = BOX;
      }
      return cell;
    }
    function dv(cell, list) { ws.getCell(cell).dataValidation = { type: 'list', allowBlank: true, formulae: ['"' + list.join(',') + '"'] }; }
    function small(r, c, txt, valign) {
      ws.getCell(r, c).value = txt; ws.getCell(r, c).font = font({ size: 6, bold: txt !== 'NOTE', color: { argb: C.MUTED } });
      ws.getCell(r, c).alignment = { horizontal: 'right', vertical: valign || 'top' };
    }
    function statusColor(ref, statusCell, list) {
      ws.addConditionalFormatting({ ref: ref, rules: [
        { type: 'expression', priority: 1, formulae: [statusCell + '="' + list[2] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.RED } }, font: { color: { argb: 'FFFFFFFF' } } } },
        { type: 'expression', priority: 2, formulae: [statusCell + '="' + list[1] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.YEL } } } }
      ] });
    }

    /* header */
    ws.mergeCells('B1:T1');
    ws.getCell('B1').value = 'STATUS SEAT PLAN A320(03)';
    ws.getCell('B1').font = font({ size: 18, bold: true }); ws.getCell('B1').alignment = CENTER;
    for (c = 2; c <= 20; c++) ws.getCell(1, c).fill = fill(C.TAG);
    ws.getRow(1).height = 32;
    var ls = data.lastSave || {};
    put(2, 2, 'SAVE ล่าสุด', { font: font({ bold: true }), fill: C.ZONEH });
    put(2, 4, ls.by || '-', { span: 7, fill: C.GREY });
    put(2, 11, 'วัน/เวลาที่ SAVE (TH-LT)', { span: 3, font: font({ size: 8, bold: true }), fill: C.ZONEH });
    put(2, 14, ls.at ? thText(ls.at) : '-', { span: 6, fill: C.GREY });
    put(3, 2, 'ผู้ตรวจ', { font: font({ bold: true }), fill: C.ZONEH });
    put(3, 4, data.inspector || null, { span: 7, fill: C.YEL });
    put(3, 11, 'วัน/เวลาที่ตรวจ', { span: 3, font: font({ size: 8, bold: true }), fill: C.ZONEH });
    put(3, 14, null, { span: 6, fill: C.YEL, fmt: 'dd/mm/yyyy hh:mm' });
    ws.getRow(2).height = 26; ws.getRow(3).height = 26;
    ws.mergeCells('B4:T4');
    ws.getCell('B4').value = 'บันทึกสภาพที่นั่งและ Monitor · Export จากหน้าเว็บ ' + thText(new Date()) + ' · เวลาทั้งหมดเป็นเวลาประเทศไทย (TH-LT)';
    ws.getCell('B4').font = font({ size: 8, italic: true }); ws.getCell('B4').alignment = CENTER;

    var SUM0 = 5, TYP0 = 13;
    [['ZONE', 2, 3, null], ['ที่นั่งรวม', 5, 2, null], ['ใช้งานได้ปกติ', 7, 2, C.GREEN], ['มีข้อขัดข้อง\nแต่ใช้งานได้', 9, 2, C.YEL],
     ['ใช้งานไม่ได้', 11, 2, C.RED], ['Monitor รวม', 13, 2, C.BLUE], ['Monitor\nใช้งานได้', 15, 2, C.MGREEN],
     ['Monitor\nเสื่อมสภาพ', 17, 2, C.YEL], ['Monitor\nเสีย', 19, 2, C.RED]].forEach(function (h) {
      put(SUM0, h[1], h[0], { span: h[2], font: font({ size: 8, bold: true }), fill: h[3] });
    });
    ws.getRow(SUM0).height = 30;
    [['ชนิด Monitor', 2, 9, null], ['จำนวน', 11, 2, C.BLUE], ['ใช้งานได้', 13, 2, C.MGREEN], ['เสื่อมสภาพ', 15, 2, C.YEL], ['เสีย', 17, 2, C.RED]].forEach(function (h) {
      put(TYP0, h[1], h[0], { span: h[2], font: font({ size: 8, bold: true }), fill: h[3] });
    });
    ws.getRow(TYP0).height = 22;
    ws.getCell('B18').value = 'วิธีใช้'; ws.getCell('B18').font = font({ size: 8, bold: true });
    ws.mergeCells('C18:T18');
    ws.getCell('C18').value = 'เลือกสภาพจาก dropdown ใต้ชื่อที่นั่งและใต้ป้าย Monitor · พิมพ์รายละเอียดในช่อง NOTE · สีและตารางสรุปเปลี่ยนเอง · Monitor = จอติดผนัง 21.5 นิ้ว, OVHD Monitor = จอพับเหนือที่นั่ง 12.0 นิ้ว';
    ws.getCell('C18').font = font({ size: 8 }); ws.getCell('C18').alignment = { wrapText: true, vertical: 'middle' };
    ws.getRow(18).height = 30;

    var r = 20, cells = {}, order = [];
    function zoneHeader(title, tag) {
      ws.mergeCells(r, 2, r, 16); ws.getCell(r, 2).value = title;
      ws.getCell(r, 2).font = font({ size: 11, bold: true }); ws.getCell(r, 2).alignment = { vertical: 'middle' };
      for (var cc = 2; cc <= 16; cc++) ws.getCell(r, cc).fill = fill(C.ZONEH);
      ws.mergeCells(r, 17, r, 20); ws.getCell(r, 17).value = tag;
      ws.getCell(r, 17).font = font({ size: 10, bold: true }); ws.getCell(r, 17).alignment = CENTER;
      for (cc = 17; cc <= 20; cc++) ws.getCell(r, cc).fill = fill(C.TAG);
      ws.getRow(r).height = 22; r++; ws.getRow(r).height = 6; r++;
    }
    function cabin(r0, r1) {
      for (var rr = r0; rr <= r1; rr++) for (var cc = 2; cc <= 20; cc++) {
        var x = ws.getCell(rr, cc);
        if (!x.fill || !x.fill.fgColor) x.fill = fill(C.CABIN);
      }
    }
    function gapRow(h) { ws.getRow(r).height = h || 5; r++; }
    /* seats: [[id, column, span, text]] all in one cabin row */
    function seatRow(zone, num, seats) {
      if (num !== null) put(r, 2, num, { span: 1, font: font({ size: 10, bold: true, color: { argb: 'FFFFFFFF' } }), fill: C.ROWNO });
      seats.forEach(function (s) {
        var id = s[0], c = s[1], span = s[2] || 2, a = col(c);
        put(r, c, s[3] || id, { span: span, font: font({ size: s[3] ? 8 : 11, bold: true }), fill: C.GREEN });
        put(r + 1, c, S[0], { span: span, font: font({ size: 7 }), fill: C.WHITE }); dv(a + (r + 1), S);
        put(r + 2, c, null, { span: span, font: font({ size: 7 }), fill: C.WHITE, align: { wrapText: true, vertical: 'top', horizontal: 'left' } });
        statusColor(a + r + ':' + col(c + span - 1) + r, '$' + a + '$' + (r + 1), S);
        cells[id] = { zone: zone, kind: 'seat', label: s[3] ? 'Commander Center 1A' : id, seat: a + (r + 1), mon: null, note: a + (r + 2) };
        order.push(id);
      });
      ws.getRow(r).height = 22; ws.getRow(r + 1).height = 14; ws.getRow(r + 2).height = 30;
      small(r + 2, 2, 'NOTE');
      r += 3;
    }
    /* monitors: [left, right], each null or [id, type 'W'|'R', position text, first column (optional)]
     * the 4-column Monitor cell sits centred over the seats it serves */
    function monRow(zone, pair) {
      pair.forEach(function (m, i) {
        if (!m) return;
        var c = m[3] || (i ? MON_R : MON_L), a = col(c), id = m[0], label = MTYPE[m[1]] + ' ' + m[2];
        put(r, c, MTYPE_CELL[m[1]], { span: 4, font: font({ size: 7, bold: true }), fill: m[1] === 'R' ? C.OVHD : C.MGREEN });
        put(r + 1, c, M[0], { span: 4, font: font({ size: 7 }), fill: C.WHITE }); dv(a + (r + 1), M);
        put(r + 2, c, null, { span: 4, font: font({ size: 7 }), fill: C.WHITE, align: { wrapText: true, vertical: 'top', horizontal: 'left' } });
        statusColor(a + r + ':' + col(c + 3) + r, '$' + a + '$' + (r + 1), M);
        cells[id] = { zone: zone, kind: 'mon', type: m[1], label: label, seat: null, mon: a + (r + 1), note: a + (r + 2) };
        order.push(id);
      });
      small(r, 2, 'Monitor', 'middle');
      ws.getRow(r).height = 14; ws.getRow(r + 1).height = 14; ws.getRow(r + 2).height = 24;
      small(r + 2, 2, 'NOTE');
      r += 3;
    }
    function furniture(r0, c0, rows, span, txt) {
      put(r0, c0, txt, { span: span, rows: rows, font: font({ size: 9, bold: true, color: { argb: C.MUTED } }), fill: C.FURN });
    }
    function frame(r0, r1, c0, c1) {
      var med = { style: 'medium', color: { argb: C.MUTED } };
      frames.push([r0, r1, c0, c1]);
      for (var rr = r0; rr <= r1; rr++) for (var cc = c0; cc <= c1; cc++) {
        var b = Object.assign({}, ws.getCell(rr, cc).border || {});
        if (cc === c0) b.left = med; if (cc === c1) b.right = med; if (rr === r0) b.top = med; if (rr === r1) b.bottom = med;
        ws.getCell(rr, cc).border = b;
      }
    }
    var z0, zr = {}, rr0, frames = [];

    /* FRONT CREW ZONE: Commander Center room on the left, 1D/1F on the right */
    z0 = r; zoneHeader('Front Crew Zone (3 Seats)', 'CREW');
    rr0 = r;
    ws.mergeCells(r, 3, r, 8); ws.getCell(r, 3).value = 'ห้อง Commander Center';
    ws.getCell(r, 3).font = font({ size: 8, bold: true, color: { argb: C.MUTED } }); ws.getCell(r, 3).alignment = CENTER; r++;
    monRow('F', [['MON-F1', 'W', 'หน้าที่นั่ง CC 1A'], ['MON-F2', 'W', 'หน้าที่นั่ง 1D, 1F']]);
    seatRow('F', 1, [['CC1A', 3, 6, 'Commander Center 1A'], ['1D', COLS.D], ['1F', COLS.F]]);
    frame(rr0, r - 1, 3, 8);
    zr.F = [z0, r]; cabin(z0 + 1, r); gapRow(); r++;

    /* VVIP OA: 2A and 3A on the left, sofa on the right */
    z0 = r; zoneHeader('VVIP OA (2 Seats)', 'VVIP OA');
    monRow('A', [['MON-A1', 'W', 'หน้าที่นั่ง 2A'], null]);
    var sofa0 = r;
    /* 2A and 3A sit in the middle column of the left side so they line up under the Monitor */
    seatRow('A', 2, [['2A', COLS.B]]); gapRow();
    seatRow('A', 3, [['3A', COLS.B]]);
    furniture(sofa0, 14, r - sofa0, 6, 'SOFA');
    monRow('A', [['MON-A2', 'W', 'หลังที่นั่ง 3A'], ['MON-A3', 'W', 'หลังโซฟา']]);
    zr.A = [z0, r]; cabin(z0 + 1, r); gapRow(); r++;

    /* VIP OB: Meeting Area */
    z0 = r; zoneHeader('VIP OB · Meeting Area (6 Seats)', 'VIP OB');
    monRow('B', [['MON-B1', 'W', 'หน้าที่นั่ง 5B'], ['MON-B2', 'W', 'หน้าที่นั่ง 5D, 5F']]);
    /* right side: 5D/5F and 6D/6F sit side by side under their Monitor (columns 15-18), table as wide as the 2 seats */
    seatRow('B', 5, [['5B', COLS.B], ['5D', 15], ['5F', 17]]); gapRow();
    furniture(r, COLS.B, 1, 2, 'TABLE'); furniture(r, 15, 1, 4, 'TABLE'); ws.getRow(r).height = 22; r++; gapRow();
    seatRow('B', 6, [['6B', COLS.B], ['6D', 15], ['6F', 17]]);
    monRow('B', [['MON-B3', 'W', 'หลังที่นั่ง 6B'], ['MON-B4', 'W', 'หลังที่นั่ง 6D, 6F']]);
    zr.B = [z0, r]; cabin(z0 + 1, r); gapRow(); r++;

    /* VVIP OC: Business Class */
    z0 = r; zoneHeader('VVIP OC · Business Class (8 Seats)', 'VVIP OC');
    monRow('C', [['MON-C1', 'W', 'หน้าที่นั่ง 7B, 7C'], ['MON-C2', 'W', 'หน้าที่นั่ง 7D, 7F']]);
    /* seat pairs sit side by side under their Monitor, lined up with VVIP OA / OB: left columns 4-7, right columns 15-18 */
    seatRow('C', 7, [['7B', 4], ['7C', 6], ['7D', 15], ['7F', 17]]); gapRow();
    seatRow('C', 8, [['8B', 4], ['8C', 6], ['8D', 15], ['8F', 17]]);
    monRow('C', [['MON-C3', 'W', 'หลังที่นั่ง 8B, 8C'], ['MON-C4', 'W', 'หลังที่นั่ง 8D, 8F']]);
    zr.C = [z0, r]; cabin(z0 + 1, r); gapRow(); r++;

    /* PASSENGER OD: Economy, OVHD Monitors in front of rows 10, 12, 15, 17 */
    z0 = r; zoneHeader('Passenger OD · Economy Class (54 Seats)', 'ECONOMY');
    monRow('D', [['MON-D1', 'W', 'หน้าที่นั่ง 9A, 9B, 9C'], ['MON-D2', 'W', 'หน้าที่นั่ง 9D, 9E, 9F']]);
    [9, 10, 11, 12, 14, 15, 16, 17, 18].forEach(function (n, i) {
      if ([10, 12, 15, 17].indexOf(n) >= 0) monRow('D', [['MON-D' + n + 'L', 'R', 'หน้าที่นั่ง ' + n + 'A, ' + n + 'B, ' + n + 'C'], ['MON-D' + n + 'R', 'R', 'หน้าที่นั่ง ' + n + 'D, ' + n + 'E, ' + n + 'F']]);
      else if (i) gapRow();
      seatRow('D', n, 'ABCDEF'.split('').map(function (L) { return [n + L, COLS[L]]; }));
    });
    zr.D = [z0, r]; cabin(z0 + 1, r);
    var LAST = r;

    var area = 'C20:T' + LAST;
    ws.addConditionalFormatting({ ref: area, rules: [
      { type: 'cellIs', operator: 'equal', priority: 3, formulae: ['"' + S[1] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.YEL } } } },
      { type: 'cellIs', operator: 'equal', priority: 4, formulae: ['"' + M[1] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.YEL } } } },
      { type: 'cellIs', operator: 'equal', priority: 5, formulae: ['"' + S[2] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.RED } }, font: { color: { argb: 'FFFFFFFF' } } } },
      { type: 'cellIs', operator: 'equal', priority: 6, formulae: ['"' + M[2] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.RED } }, font: { color: { argb: 'FFFFFFFF' } } } }
    ] });

    /* summary formulas count the list sheet (with cached results so every viewer shows numbers at once) */
    var q = "'" + LIST + "'!";
    function cnt(pred) {
      var o = { t: 0, s: [0, 0, 0], mt: 0, m: [0, 0, 0] };
      order.forEach(function (id) { var it = cells[id], d = state[id] || {}; if (!pred(it)) return;
        if (it.seat) { o.t++; o.s[d.s | 0]++; } if (it.mon) { o.mt++; o.m[d.m | 0]++; } });
      return o;
    }
    var tot = cnt(function () { return true; });
    ZKEYS.forEach(function (k, i) {
      var row = SUM0 + 1 + i, n = cnt(function (it) { return it.zone === k; });
      put(row, 2, ZN[k], { span: 3, font: font({ size: 8, bold: true }), align: { vertical: 'middle' } });
      put(row, 5, { formula: 'G' + row + '+I' + row + '+K' + row, result: n.t });
      S.forEach(function (t, j) { put(row, 7 + 2 * j, { formula: 'COUNTIFS(' + q + '$C:$C,$B' + row + ',' + q + '$D:$D,"' + t + '")', result: n.s[j] }, { fill: [C.GREEN, C.YEL, C.RED][j] }); });
      put(row, 13, { formula: 'O' + row + '+Q' + row + '+S' + row, result: n.mt }, { fill: C.BLUE });
      M.forEach(function (t, j) { put(row, 15 + 2 * j, { formula: 'COUNTIFS(' + q + '$C:$C,$B' + row + ',' + q + '$E:$E,"' + t + '")', result: n.m[j] }, { fill: [C.MGREEN, C.YEL, C.RED][j] }); });
    });
    var rt = SUM0 + 6;
    put(rt, 2, 'รวม', { span: 3, font: font({ size: 8, bold: true }), align: { vertical: 'middle' } });
    [[5, null, tot.t], [7, C.GREEN, tot.s[0]], [9, C.YEL, tot.s[1]], [11, C.RED, tot.s[2]], [13, C.BLUE, tot.mt], [15, C.MGREEN, tot.m[0]], [17, C.YEL, tot.m[1]], [19, C.RED, tot.m[2]]].forEach(function (x) {
      put(rt, x[0], { formula: 'SUM(' + col(x[0]) + (SUM0 + 1) + ':' + col(x[0]) + (SUM0 + 5) + ')', result: x[2] }, { font: font({ bold: true }), fill: x[1] });
    });
    /* summary by Monitor type */
    ['W', 'R'].forEach(function (k, i) {
      var row = TYP0 + 1 + i, n = cnt(function (it) { return it.type === k; });
      put(row, 2, MTYPE_FULL[k] + ' (' + MTYPE[k] + ')', { span: 9, font: font({ size: 8, bold: true }), align: { vertical: 'middle', wrapText: true } });
      put(row, 11, { formula: 'M' + row + '+O' + row + '+Q' + row, result: n.mt }, { fill: C.BLUE });
      M.forEach(function (t, j) { put(row, 13 + 2 * j, { formula: 'COUNTIFS(' + q + '$B:$B,"' + MTYPE[k] + '",' + q + '$E:$E,"' + t + '")', result: n.m[j] }, { fill: [C.MGREEN, C.YEL, C.RED][j] }); });
      ws.getRow(row).height = 24;
    });
    var rt2 = TYP0 + 3;
    put(rt2, 2, 'รวม', { span: 9, font: font({ size: 8, bold: true }), align: { vertical: 'middle' } });
    [[11, C.BLUE, tot.mt], [13, C.MGREEN, tot.m[0]], [15, C.YEL, tot.m[1]], [17, C.RED, tot.m[2]]].forEach(function (x) {
      put(rt2, x[0], { formula: 'SUM(' + col(x[0]) + (TYP0 + 1) + ':' + col(x[0]) + (TYP0 + 2) + ')', result: x[2] }, { font: font({ bold: true }), fill: x[1] });
    });

    /* current state */
    order.forEach(function (id) {
      var it = cells[id], d = state[id]; if (!d) return;
      if (it.seat) ws.getCell(it.seat).value = S[d.s | 0];
      if (it.mon) ws.getCell(it.mon).value = M[d.m | 0];
      if (d.note) ws.getCell(it.note).value = d.note;
    });

    /* list sheets */
    var HDR = ['ที่นั่ง / จุด', 'ประเภท', 'โซน', 'สภาพที่นั่ง', 'สภาพ Monitor', 'หมายเหตุ', 'ผู้บันทึกล่าสุด', 'วัน/เวลาล่าสุด (TH-LT)'];
    var WID = [34, 14, 18, 16, 14, 44, 24, 20];
    function statusCF(s, end) {
      if (end < 2) return;
      s.addConditionalFormatting({ ref: 'D2:E' + end, rules: [
        { type: 'cellIs', operator: 'equal', priority: 1, formulae: ['"' + S[1] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.YEL } } } },
        { type: 'cellIs', operator: 'equal', priority: 2, formulae: ['"' + M[1] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.YEL } } } },
        { type: 'cellIs', operator: 'equal', priority: 3, formulae: ['"' + S[2] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.RED } }, font: { color: { argb: 'FFFFFFFF' } } } },
        { type: 'cellIs', operator: 'equal', priority: 4, formulae: ['"' + M[2] + '"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: C.RED } }, font: { color: { argb: 'FFFFFFFF' } } } }
      ] });
    }
    function header(s, titles, widths) {
      titles.forEach(function (t, j) {
        var c = s.getCell(1, j + 1); c.value = t; c.font = font({ size: 10, bold: true }); c.fill = fill(C.ZONEH); c.border = BOX; c.alignment = CENTER;
        s.getColumn(j + 1).width = widths[j];
      });
      s.views = [{ state: 'frozen', ySplit: 1 }];
    }
    function listSheet(title, ids, live) {
      var s = wb.addWorksheet(title); header(s, HDR, WID);
      var qm = "'ผังที่นั่ง'!";
      ids.forEach(function (id, i) {
        var row = 2 + i, it = cells[id], d = state[id] || {};
        var vals = [it.label, it.seat ? 'ที่นั่ง' : MTYPE[it.type], ZN[it.zone]];
        if (live) {
          vals.push(it.seat ? { formula: qm + it.seat, result: S[d.s | 0] } : '-');
          vals.push(it.mon ? { formula: qm + it.mon, result: M[d.m | 0] } : '-');
          vals.push({ formula: 'IF(' + qm + it.note + '="",""' + ',' + qm + it.note + ')', result: d.note || '' });
        } else {
          vals.push(it.seat ? S[d.s | 0] : '-', it.mon ? M[d.m | 0] : '-', d.note || '');
        }
        vals.push(d.by || '', d.at ? th(d.at) : '');
        vals.forEach(function (v, j) {
          var c = s.getCell(row, j + 1); c.value = v; c.font = font(); c.border = BOX;
          c.alignment = { horizontal: (j === 0 || j === 5 || j === 6) ? 'left' : 'center', vertical: 'top', wrapText: j === 5 };
        });
        s.getCell(row, 8).numFmt = 'dd/mm/yyyy hh:mm';
      });
      var end = 1 + ids.length; statusCF(s, end);
      if (end >= 2) s.autoFilter = 'A1:H' + end;
      return s;
    }
    var bad = order.filter(function (id) { var d = state[id] || {}; return (d.s | 0) || (d.m | 0) || d.note; });
    var sBad = listSheet('รวมข้อขัดข้องคงค้าง', bad, false);
    if (!bad.length) { sBad.getCell('A2').value = 'ไม่มีข้อขัดข้องคงค้าง'; sBad.getCell('A2').font = font(); }
    listSheet(LIST, order, true);

    /* history: every cell centred (horizontal + vertical) */
    var h = wb.addWorksheet('ประวัติการแก้ไข');
    header(h, ['วัน/เวลา (TH-LT)', 'อีเมลผู้แก้ไข', 'ผู้ตรวจ', 'ที่นั่ง / จุด', 'รายการ', 'ค่าเดิม', 'ค่าใหม่'], [20, 30, 24, 34, 14, 36, 36]);
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
    wb.seatCells = cells; wb.seatOrder = order; wb.seatZoneRows = zr; wb.seatFrames = frames; wb.seatLastRow = LAST;
    return wb;
  }

  root.buildSeatWorkbook = buildSeatWorkbook;
  root.seatThText = thText;
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
