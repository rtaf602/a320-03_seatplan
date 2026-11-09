/* STATUS SEAT PLAN A320-03 — ผังที่นั่งในชีท "ผังที่นั่ง"
 * ใช้ร่วมกัน 2 ที่: หน้าเว็บ (export-excel.js) และ apps-script/Code.gs (tools/build-code-gs.sh คัดลอกส่วนนี้ไปไว้ท้าย Code.gs)
 * ถ้าแก้ไฟล์นี้ ต้องแก้ส่วนเดียวกันใน Code.gs ให้ตรงกันด้วย
 *
 * seatPlanLayout() คืน { order:[id...], items:{id:{id,kind,zone,type,label,name}}, blocks:[...] }
 *   label = ข้อความในช่องหัวของที่นั่ง/จอในชีท ใช้หาตำแหน่งช่อง (จอชนิดเดียวกันใช้ชื่อซ้ำกันได้ จับคู่ตามลำดับซ้ายไปขวา บนลงล่าง)
 *   alias = ชื่อแบบเดิม (ข้อความเดียวหรือ array) ใช้อ่านชีทที่สร้างด้วยรุ่นก่อน เช่น 'OVHD 12.0"', 'Commander Center 1A' ของชีท A320-03 เดิม
 *   name  = ชื่อที่ใช้ในแท็บ "ทุกที่นั่ง และ Monitor" และ "ประวัติการแก้ไข"
 * planSheet(L) คืนรายการช่องที่ต้องเขียน/จัดรูปแบบ และตำแหน่งของแต่ละที่นั่ง/จอ
 */
var SEAT_S = ['ปกติ', 'ขัดข้อง-ใช้ได้', 'ใช้ไม่ได้'];
var SEAT_M = ['จอปกติ', 'จอเสื่อม', 'จอเสีย'];
var SEAT_E = ['ใช้งานได้', 'ใช้งานได้บางส่วน', 'ใช้งานไม่ได้'];   // Galley equipment and lavatories
var SEAT_GAL = { FWD: 'FWD GALLEY', AFT: 'AFT GALLEY' };
var SEAT_ETYPE = { OVEN: 'Oven', CM: 'Coffee Maker', MW: 'Microwave', HOTCUP: 'Hot Cup', BOILER: 'Water Boiler', CHILLER: 'Air Chiller', FRIDGE: 'Refrigerator', STORAGE: 'Storage', STOWAGE: 'Stowage', TROLLEY: 'Trolley Stowage', HEATER: 'Water Heater', FAUCET: 'Faucet', LIGHT: 'Work Light', OUTLET: '220VAC Power Outlet', WASTE: 'Waste' };
var SEAT_ETYPE_ORDER = ['OVEN', 'CM', 'MW', 'HOTCUP', 'BOILER', 'CHILLER', 'FRIDGE', 'STORAGE', 'STOWAGE', 'TROLLEY', 'HEATER', 'FAUCET', 'LIGHT', 'OUTLET', 'WASTE'];
/* Names the A320-03 sheet used in the history tab before Galley / lavatories were added (Oct 2026).
 * The history keeps those rows, so the last editor of each Monitor is still found under its old name. */
var SEAT_HIST_NAMES = {
  "MON-F1": "Monitor หน้าที่นั่ง CC 1A",
  "MON-F2": "Monitor หน้าที่นั่ง 1D, 1F",
  "MON-A1": "Monitor หน้าที่นั่ง 2A",
  "MON-A2": "Monitor หลังที่นั่ง 3A",
  "MON-A3": "Monitor หลังโซฟา",
  "MON-B1": "Monitor หน้าที่นั่ง 5B",
  "MON-B2": "Monitor หน้าที่นั่ง 5D, 5F",
  "MON-B3": "Monitor หลังที่นั่ง 6B",
  "MON-B4": "Monitor หลังที่นั่ง 6D, 6F",
  "MON-C1": "Monitor หน้าที่นั่ง 7B, 7C",
  "MON-C2": "Monitor หน้าที่นั่ง 7D, 7F",
  "MON-C3": "Monitor หลังที่นั่ง 8B, 8C",
  "MON-C4": "Monitor หลังที่นั่ง 8D, 8F",
  "MON-D1": "Monitor หน้าที่นั่ง 9A, 9B, 9C",
  "MON-D2": "Monitor หน้าที่นั่ง 9D, 9E, 9F",
  "MON-D10L": "OVHD Monitor หน้าที่นั่ง 10A, 10B, 10C",
  "MON-D10R": "OVHD Monitor หน้าที่นั่ง 10D, 10E, 10F",
  "MON-D12L": "OVHD Monitor หน้าที่นั่ง 12A, 12B, 12C",
  "MON-D12R": "OVHD Monitor หน้าที่นั่ง 12D, 12E, 12F",
  "MON-D15L": "OVHD Monitor หน้าที่นั่ง 15A, 15B, 15C",
  "MON-D15R": "OVHD Monitor หน้าที่นั่ง 15D, 15E, 15F",
  "MON-D17L": "OVHD Monitor หน้าที่นั่ง 17A, 17B, 17C",
  "MON-D17R": "OVHD Monitor หน้าที่นั่ง 17D, 17E, 17F"
};
var SEAT_MT = {
  w21: { name: 'Wall Mounted Monitor 21.5 Inch', short: 'Monitor 21.5"', color: '#8ED873' },   // เขียว
  b12: { name: 'Bin Mounted Retractable Monitors (OVHD Monitor) 12.0 Inch', short: 'OVHD Monitor 12.0"', legacy: 'OVHD 12.0"', color: '#B4A7D6' }   // legacy = ชื่อในชีท A320-03 รุ่นเดิม  // ม่วง
};
var SEAT_MT_ORDER = ['w21', 'b12'];
var SEAT_ZONES = [
  { key: 'F', name: 'FRONT CREW ZONE' }, { key: 'A', name: 'VVIP OA' }, { key: 'B', name: 'VIP OB' },
  { key: 'C', name: 'VVIP OC' }, { key: 'D', name: 'PASSENGER OD' }
];
var SEAT_COLOR = {
  GREEN: '#00FF00', MGREEN: '#8ED873', YEL: '#FFFF00', RED: '#FF0000', CABIN: '#CAEDFB', ROWNO: '#6D9EEB',
  ZONEH: '#C1E4F5', TAG: '#45B0E1', WHITE: '#FFFFFF', BLUE: '#00B0F0', GREY: '#F2F2F2', MUTED: '#566575', TABLE: '#E7E6E6',
  GAL: '#1EBDD2', LAV: '#DDE7FB', EMPTY: '#E7ECF1', DOOR: '#E1E6EB', EXIT: '#FCE4D6'
};

function seatPlanLayout() {
  var items = {}, order = [], blocks = [];
  function seat(id, zone, label, name) { items[id] = { id: id, kind: 'seat', zone: zone, label: label || id, name: name || id }; order.push(id); }
  function mon(id, zone, type, posLong, posShort) {
    items[id] = { id: id, kind: 'mon', zone: zone, type: type, label: SEAT_MT[type].short, alias: [SEAT_MT[type].short + ' · ' + posShort].concat(SEAT_MT[type].legacy || []), name: SEAT_MT[type].name + ' · ' + posLong };
    order.push(id);
  }
  function zone(k, title, tag) { blocks.push({ t: 'zone', zone: k, title: title, tag: tag }); }
  /* one block = 3 sheet rows (label / status / note). cells: {id,c,span} or {text,c,span} */
  function blk(k, b, cells, room) { blocks.push({ t: 'blk', zone: k, b: b, cells: cells, room: !!room }); }
  function title(k, text, c, span) { blocks.push({ t: 'title', zone: k, text: text, c: c, span: span, room: true }); }
  /* Galley equipment / lavatory / door (same ids as the web page: FWD-OVEN-1, LAV-A ...) */
  var ESHORT = { STOWAGE: 'Stow', TROLLEY: 'Trolley', CM: 'CM', OVEN: 'Oven', MW: 'Microwave', HOTCUP: 'Hot Cup', BOILER: 'Water Boiler', HEATER: 'Water Heater', FAUCET: 'Faucet (Sink)', LIGHT: 'Work Light', OUTLET: '220V Outlet', FRIDGE: 'Fridge', CHILLER: 'Air Chiller', WASTE: 'Waste' };
  function eqShort(type, label) { var m = type === 'OUTLET' ? null : /^[^(]*?(\d+)/.exec(label); return (type === 'OVEN' && !m ? 'Steam Oven' : ESHORT[type]) + (m ? ' ' + m[1] : ''); }
  function eq(gal, code, type, label) { var id = gal + '-' + code; items[id] = { id: id, kind: 'eq', zone: gal, gal: gal, type: type, label: eqShort(type, label), full: label, name: SEAT_GAL[gal] + ' · ' + label, group: SEAT_GAL[gal], tname: SEAT_ETYPE[type] }; order.push(id); return id; }
  function lav(k, zone, label) { var id = 'LAV-' + k; items[id] = { id: id, kind: 'lav', zone: zone, label: label, name: label, tname: 'LAVATORY' }; order.push(id); return id; }
  function door(k, left, right, exit, code) { blocks.push({ t: 'door', zone: k, left: left, right: right, exit: !!exit, code: code }); }
  function gtitle(k, text, c, span) { blocks.push({ t: 'gtitle', zone: k, text: text, c: c, span: span }); }
  /* one galley row spread over columns c0..c1 with every cell the same width (as near as the sheet columns allow) */
  var COLPX = function (c) { return c === 9 ? 22 : 52; };
  function even(ids, c0, c1) {
    var edges = [c0], acc = [0], tot = 0, c;
    for (c = c0; c <= c1; c++) { tot += COLPX(c); acc.push(tot); }
    var n = ids.length, cut = [0];
    for (var k = 1; k < n; k++) {
      var want = tot * k / n, best = cut[k - 1] + 1;
      for (var b = cut[k - 1] + 1; b <= (c1 - c0 + 1) - (n - k); b++) if (Math.abs(acc[b] - want) < Math.abs(acc[best] - want)) best = b;
      cut.push(best);
    }
    cut.push(c1 - c0 + 1);
    return ids.map(function (id, i) { return { id: id, c: c0 + cut[i], span: cut[i + 1] - cut[i] }; });
  }
  var C1 = 3, C2 = 5, C3 = 7, R1 = 10, R2 = 12, R3 = 14;  // C E G | J L N  (left side C..H, right side J..O)
  var L1 = 5, L2 = 4, R2S = 11;                             // centred: 1 seat at E, 2 seats at D-G, 2 seats right at K-N

  zone('F', 'Front Crew Zone (3 Seats)', 'CREW');
  mon('MON-F1', 'F', 'w21', 'หน้าที่นั่ง CC 1A', 'หน้า CC 1A');
  mon('MON-F2', 'F', 'w21', 'หน้าที่นั่ง 1D, 1F', 'หน้า 1D-1F');
  /* FWD GALLEY = G1 GALLEY LAYOUT (A320 CCOM-S 02-140) on the right side (columns of 1D-1F), LAV A on the left.
     Sheet columns J K L M N-O = Refrigerator/Steam Oven/Trolley | | CM 1, Stowage 1, MW, Stowage 2 | CM 2, Hot Cup, Sink, Waste | Air Chiller.
     rows: 2 = the cell runs down over two blocks */
  var EMP = SEAT_COLOR.EMPTY;
  gtitle('F', 'FWD GALLEY (G1)', R1, 6);
  blk('F', 'GALLEY', [{ id: lav('A', 'F', 'LAVATORY A'), c: L2, span: 4 }, { id: eq('FWD', 'STOWAGE-1', 'STOWAGE', 'Stowage 1'), c: R1, span: 1 }, { id: eq('FWD', 'STOWAGE-2', 'STOWAGE', 'Stowage 2'), c: R1 + 1, span: 1 }, { text: 'Elec. Panel', c: R1 + 2, span: 2, fill: EMP }, { text: '', c: R1 + 4, span: 2, fill: EMP }]);
  blk('F', '', [{ id: eq('FWD', 'FRIDGE-1', 'FRIDGE', 'Refrigerator'), c: R1, span: 1, rows: 2 }, { id: eq('FWD', 'OVEN-1', 'OVEN', 'Steam Oven'), c: R1 + 1, span: 1, rows: 2 }, { id: eq('FWD', 'CM-1', 'CM', 'Coffee Maker 1'), c: R1 + 2, span: 1 }, { id: eq('FWD', 'CM-2', 'CM', 'Coffee Maker 2'), c: R1 + 3, span: 1 }, { id: eq('FWD', 'LIGHT-1', 'LIGHT', 'Work Light'), c: R1 + 4, span: 2 }]);
  blk('F', '', [{ id: eq('FWD', 'STOWAGE-3', 'STOWAGE', 'Stowage 3 (ข้าง Oven ใต้ CM)'), c: R1 + 2, span: 2 }, { id: eq('FWD', 'HOTCUP-1', 'HOTCUP', 'Hot Cup'), c: R1 + 4, span: 1 }, { id: eq('FWD', 'FAUCET-1', 'FAUCET', 'Faucet (Sink)'), c: R1 + 5, span: 1 }]);
  blk('F', '', [{ id: eq('FWD', 'TROLLEY-1', 'TROLLEY', 'Trolley Stowage 1 (Half / Full Size)'), c: R1, span: 1, rows: 3 }, { id: eq('FWD', 'TROLLEY-2', 'TROLLEY', 'Trolley Stowage 2 (Half / Full Size)'), c: R1 + 1, span: 1, rows: 3 }, { id: eq('FWD', 'MW-1', 'MW', 'Microwave'), c: R1 + 2, span: 1 }, { text: 'Sink', c: R1 + 3, span: 1, fill: EMP }, { id: eq('FWD', 'OUTLET-1', 'OUTLET', '220VAC Power Outlet'), c: R1 + 4, span: 2 }]);
  blk('F', '', [{ id: eq('FWD', 'STOWAGE-4', 'STOWAGE', 'Stowage 4 (ช่องที่มี Water Heater)'), c: R1 + 2, span: 1 }, { id: eq('FWD', 'WASTE-1', 'WASTE', 'Waste Compartment (behind Stowage door)'), c: R1 + 3, span: 1, rows: 2 }, { id: eq('FWD', 'CHILLER-1', 'CHILLER', 'Air Chiller'), c: R1 + 4, span: 2, rows: 2 }]);
  blk('F', '', [{ id: eq('FWD', 'HEATER-1', 'HEATER', 'Water Heater (ใน Stowage 4)'), c: R1 + 2, span: 1 }]);
  door('F', '◀ DOOR 1L', 'DOOR 1R ▶');
  title('F', 'ห้อง Commander Center', C1, 6);
  blk('F', 'Monitor', [{ id: 'MON-F1', c: L2, span: 4 }, { id: 'MON-F2', c: R2S, span: 4 }], true);
  seat('CC1A', 'F', 'CC 1A', 'Commander Center 1A'); items.CC1A.alias = 'Commander Center 1A'; seat('1D', 'F'); seat('1F', 'F');   // alias = ชื่อในชีทรุ่นเดิม
  blk('F', '1', [{ id: 'CC1A', c: L2, span: 4 }, { id: '1D', c: R2S }, { id: '1F', c: R2S + 2 }], true);

  zone('A', 'VVIP OA (2 Seats)', 'VVIP OA');
  mon('MON-A1', 'A', 'w21', 'หน้าที่นั่ง 2A', 'หน้า 2A');
  mon('MON-A2', 'A', 'w21', 'หลังที่นั่ง 3A', 'หลัง 3A');
  mon('MON-A3', 'A', 'w21', 'หลังโซฟา', 'หลังโซฟา');
  blk('A', 'LAV', [{ id: lav('FV', 'A', 'FWD VIP LAVATORY'), c: R2S, span: 4 }]);
  blk('A', 'Monitor', [{ id: 'MON-A1', c: L1 }]);
  seat('2A', 'A'); blk('A', '2', [{ id: '2A', c: L1 }, { text: 'SOFA', c: R2S, span: 4, rows: 2 }]);
  seat('3A', 'A'); blk('A', '3', [{ id: '3A', c: L1 }]);
  blk('A', 'Monitor', [{ id: 'MON-A2', c: L1 }, { id: 'MON-A3', c: R2S, span: 4 }]);

  zone('B', 'VIP OB (6 Seats)', 'VIP OB');
  mon('MON-B1', 'B', 'w21', 'หน้าที่นั่ง 5B', 'หน้า 5B');
  mon('MON-B2', 'B', 'w21', 'หน้าที่นั่ง 5D, 5F', 'หน้า 5D-5F');
  mon('MON-B3', 'B', 'w21', 'หลังที่นั่ง 6B', 'หลัง 6B');
  mon('MON-B4', 'B', 'w21', 'หลังที่นั่ง 6D, 6F', 'หลัง 6D-6F');
  blk('B', 'Monitor', [{ id: 'MON-B1', c: L1 }, { id: 'MON-B2', c: R2S, span: 4 }]);
  seat('5B', 'B'); seat('5D', 'B'); seat('5F', 'B'); blk('B', '5', [{ id: '5B', c: L1 }, { id: '5D', c: R2S }, { id: '5F', c: R2S + 2 }]);
  blk('B', '', [{ text: 'TABLE', c: L1 }, { text: 'TABLE', c: R2S, span: 4 }]);
  seat('6B', 'B'); seat('6D', 'B'); seat('6F', 'B'); blk('B', '6', [{ id: '6B', c: L1 }, { id: '6D', c: R2S }, { id: '6F', c: R2S + 2 }]);
  blk('B', 'Monitor', [{ id: 'MON-B3', c: L1 }, { id: 'MON-B4', c: R2S, span: 4 }]);
  door('B', '', '', true, 'FWD');

  zone('C', 'VVIP OC (8 Seats)', 'VVIP OC');
  mon('MON-C1', 'C', 'w21', 'หน้าที่นั่ง 7B, 7C', 'หน้า 7B-7C');
  mon('MON-C2', 'C', 'w21', 'หน้าที่นั่ง 7D, 7F', 'หน้า 7D-7F');
  mon('MON-C3', 'C', 'w21', 'หลังที่นั่ง 8B, 8C', 'หลัง 8B-8C');
  mon('MON-C4', 'C', 'w21', 'หลังที่นั่ง 8D, 8F', 'หลัง 8D-8F');
  blk('C', '', [{ text: '◀ EMER\nEXIT\nAFT L', c: 1, span: 1, fill: SEAT_COLOR.EXIT, bold: true }, { id: lav('CV', 'C', 'CENTER VIP LAVATORY'), c: R2S, span: 4 }, { text: 'EMER\nEXIT\nAFT R ▶', c: 16, span: 1, fill: SEAT_COLOR.EXIT, bold: true }]);
  blk('C', 'Monitor', [{ id: 'MON-C1', c: L2, span: 4 }, { id: 'MON-C2', c: R2S, span: 4 }]);
  [7, 8].forEach(function (n) {
    ['B', 'C', 'D', 'F'].forEach(function (x) { seat(n + x, 'C'); });
    blk('C', String(n), [{ id: n + 'B', c: L2 }, { id: n + 'C', c: L2 + 2 }, { id: n + 'D', c: R2S }, { id: n + 'F', c: R2S + 2 }]);
  });
  blk('C', 'Monitor', [{ id: 'MON-C3', c: L2, span: 4 }, { id: 'MON-C4', c: R2S, span: 4 }]);

  zone('D', 'Passenger OD (54 Seats)', 'ECONOMY');
  mon('MON-D1', 'D', 'w21', 'หน้าที่นั่ง 9A, 9B, 9C', 'หน้า 9A-9C');
  mon('MON-D2', 'D', 'w21', 'หน้าที่นั่ง 9D, 9E, 9F', 'หน้า 9D-9F');
  blk('D', 'Monitor', [{ id: 'MON-D1', c: C1, span: 6 }, { id: 'MON-D2', c: R1, span: 6 }]);
  [9, 10, 11, 12, 14, 15, 16, 17, 18].forEach(function (n) {
    if ([10, 12, 15, 17].indexOf(n) >= 0) {
      mon('MON-D' + n + 'L', 'D', 'b12', 'หน้าที่นั่ง ' + n + 'A, ' + n + 'B, ' + n + 'C', 'หน้า ' + n + 'A-' + n + 'C');
      mon('MON-D' + n + 'R', 'D', 'b12', 'หน้าที่นั่ง ' + n + 'D, ' + n + 'E, ' + n + 'F', 'หน้า ' + n + 'D-' + n + 'F');
      blk('D', 'OVHD', [{ id: 'MON-D' + n + 'L', c: C1, span: 6 }, { id: 'MON-D' + n + 'R', c: R1, span: 6 }]);
    }
    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(function (x) { seat(n + x, 'D'); });
    blk('D', String(n), [{ id: n + 'A', c: C1 }, { id: n + 'B', c: C2 }, { id: n + 'C', c: C3 }, { id: n + 'D', c: R1 }, { id: n + 'E', c: R2 }, { id: n + 'F', c: R3 }]);
  });
  blk('D', 'LAV', [{ id: lav('D', 'D', 'LAVATORY D'), c: C1, span: 6 }, { id: lav('E', 'D', 'LAVATORY E'), c: R1, span: 6 }]);
  door('D', '◀ DOOR 4L', 'DOOR 4R ▶');
  /* AFT GALLEY = G5 GALLEY LAYOUT (A320 CCOM-S 02-140 P5), full cabin width; grid column -> sheet columns C D E F G H I-J K L-M N-O */
  gtitle('D', 'AFT GALLEY (G5)', C1, 13);
  blk('D', 'GALLEY', [{ text: '', c: 3, span: 1, fill: SEAT_COLOR.EMPTY, rows: 2 }, { text: 'Elec. Panel', c: 4, span: 3, fill: SEAT_COLOR.EMPTY }, { id: eq('AFT', 'STOWAGE-1', 'STOWAGE', 'Stowage 1'), c: 7, span: 1 }, { id: eq('AFT', 'STOWAGE-2', 'STOWAGE', 'Stowage 2'), c: 8, span: 1 }, { id: eq('AFT', 'STOWAGE-3', 'STOWAGE', 'Stowage 3'), c: 9, span: 2 }, { id: eq('AFT', 'STOWAGE-4', 'STOWAGE', 'Stowage 4'), c: 11, span: 1 }, { text: '', c: 12, span: 4, fill: SEAT_COLOR.EMPTY }]);
  blk('D', '', [{ id: eq('AFT', 'BOILER-1', 'BOILER', 'Water Boiler'), c: 4, span: 1 }, { id: eq('AFT', 'CM-1', 'CM', 'Coffee Maker 1 (Beverage Maker)'), c: 5, span: 1 }, { id: eq('AFT', 'CM-2', 'CM', 'Coffee Maker 2 (Beverage Maker)'), c: 6, span: 1 }, { id: eq('AFT', 'OVEN-1', 'OVEN', 'Steam Oven 1'), c: 7, span: 1, rows: 2 }, { id: eq('AFT', 'OVEN-2', 'OVEN', 'Steam Oven 2'), c: 8, span: 1, rows: 2 }, { id: eq('AFT', 'OVEN-3', 'OVEN', 'Steam Oven 3'), c: 9, span: 2, rows: 2 }, { id: eq('AFT', 'STOWAGE-5', 'STOWAGE', 'Stowage 5 (เหนือ Microwave)'), c: 11, span: 1 }, { id: eq('AFT', 'STOWAGE-6', 'STOWAGE', 'Stowage 6'), c: 12, span: 2, rows: 2 }, { text: '', c: 14, span: 2, fill: SEAT_COLOR.EMPTY }]);
  blk('D', '', [{ id: eq('AFT', 'OUTLET-1', 'OUTLET', '220VAC Power Outlet'), c: 3, span: 1 }, { id: eq('AFT', 'LIGHT-1', 'LIGHT', 'Work Light'), c: 4, span: 1 }, { id: eq('AFT', 'FAUCET-1', 'FAUCET', 'Faucet (Sink)'), c: 5, span: 2 }, { id: eq('AFT', 'MW-1', 'MW', 'Microwave'), c: 11, span: 1 }, { id: eq('AFT', 'STOWAGE-7', 'STOWAGE', 'Stowage 7'), c: 14, span: 2 }]);
  blk('D', '', [{ id: eq('AFT', 'STOWAGE-8', 'STOWAGE', 'Stowage 8'), c: 3, span: 1 }, { id: eq('AFT', 'WASTE-1', 'WASTE', 'Waste Compartment (Stowage)'), c: 4, span: 1, rows: 2 }, { id: eq('AFT', 'STOWAGE-9', 'STOWAGE', 'Stowage 9 (ช่องที่มี Water Heater)'), c: 5, span: 2 }, { id: eq('AFT', 'TROLLEY-1', 'TROLLEY', 'Full Size Trolley Stowage 1'), c: 7, span: 1, rows: 2 }, { id: eq('AFT', 'TROLLEY-2', 'TROLLEY', 'Full Size Trolley Stowage 2'), c: 8, span: 1, rows: 2 }, { id: eq('AFT', 'TROLLEY-3', 'TROLLEY', 'Full Size Trolley Stowage 3'), c: 9, span: 2, rows: 2 }, { id: eq('AFT', 'TROLLEY-4', 'TROLLEY', 'Full Size Trolley Stowage 4'), c: 11, span: 1, rows: 2 }, { id: eq('AFT', 'TROLLEY-5', 'TROLLEY', 'Full Size Trolley Stowage 5'), c: 12, span: 2, rows: 2 }, { id: eq('AFT', 'STOWAGE-10', 'STOWAGE', 'Stowage 10'), c: 14, span: 2 }]);
  blk('D', '', [{ text: '', c: 3, span: 1, fill: SEAT_COLOR.EMPTY }, { id: eq('AFT', 'HEATER-1', 'HEATER', 'Water Heater (ใน Stowage 9)'), c: 5, span: 2 }, { text: '', c: 14, span: 2, fill: SEAT_COLOR.EMPTY }]);
  return { order: order, items: items, blocks: blocks };
}

function seatColName(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

/* planSheet: the "ผังที่นั่ง" tab as a list of cell operations.
 * op = {r,c,rs,cs, v (text, or formula starting with '='), q (how to compute the formula's value), fill, font:{b,i,sz,color}, align:{h,v,wrap}, border, dv:'S'|'M', fmt}
 * widths/heights are in pixels. */
function planSheet(L) {
  var C = SEAT_COLOR, ops = [], heights = {}, cf = [], boxes = [], pos = {}, zoneRows = {}, typeCells = {};
  var LASTC = 15, A = seatColName;
  function op(o) { o.rs = o.rs || 1; o.cs = o.cs || 1; ops.push(o); return o; }
  var hdr = { b: true, sz: 8 }, mid = { h: 'center', v: 'middle', wrap: true };
  var widths = { 1: 60, 2: 60 };   // A: left-hand emergency exit marks (white, outside the cabin), as P on the right
  for (var c = 3; c <= 8; c++) widths[c] = 52;
  widths[9] = 22;
  for (c = 10; c <= 15; c++) widths[c] = 52;
  widths[16] = 60;   // P: right-hand emergency exit marks

  op({ r: 1, c: 2, cs: 14, v: 'STATUS SEAT PLAN A320-03', fill: C.TAG, font: { b: true, sz: 18 }, align: mid, border: false });
  heights[1] = 44;
  op({ r: 2, c: 2, cs: 2, v: 'SAVE ล่าสุด', fill: C.ZONEH, font: { b: true, sz: 9 }, align: mid, border: true });
  op({ r: 2, c: 4, cs: 6, v: '-', fill: C.GREY, font: { sz: 9 }, align: mid, border: true });
  op({ r: 2, c: 10, cs: 3, v: 'วัน/เวลาที่ SAVE (TH-LT)', fill: C.ZONEH, font: hdr, align: mid, border: true });
  op({ r: 2, c: 13, cs: 3, v: '-', fill: C.GREY, font: { sz: 9 }, align: mid, border: true });
  op({ r: 3, c: 2, cs: 2, v: 'ผู้ตรวจ', fill: C.ZONEH, font: { b: true, sz: 9 }, align: mid, border: true });
  op({ r: 3, c: 4, cs: 6, fill: C.YEL, font: { sz: 9 }, align: mid, border: true });
  op({ r: 3, c: 10, cs: 3, v: 'วัน/เวลาที่ตรวจ', fill: C.ZONEH, font: hdr, align: mid, border: true });
  op({ r: 3, c: 13, cs: 3, fill: C.YEL, font: { sz: 9 }, align: mid, border: true, fmt: 'dd/mm/yyyy hh:mm' });
  heights[2] = 30; heights[3] = 30;
  op({ r: 4, c: 2, cs: 14, v: 'บันทึกสภาพที่นั่ง, Monitor, Galley และห้องน้ำ · สำหรับ Admin ฝ่ายการช่าง และฝ่ายสื่อสาร · เวลาทั้งหมดเป็นเวลาประเทศไทย (TH-LT)', font: { i: true, sz: 8 }, align: mid, border: false });

  /* summary by zone: rows 6..12, by monitor type: rows 14..18 */
  var SZ = 6, ST = 14;
  /* headers broken by hand so no word is split in the 52 px columns (e.g. "เสื่อมสภา/พ") */
  var hdr7 = { b: true, sz: 7 };
  [['ZONE', 2, 3, null], ['ที่นั่ง\nรวม', 5, 1, null], ['ใช้งานได้\nปกติ', 6, 1, C.GREEN], ['มีข้อขัดข้อง\nแต่ใช้งานได้', 7, 1, C.YEL], ['ใช้งาน\nไม่ได้', 8, 1, C.RED],
   ['Monitor\nรวม', 10, 1, C.BLUE], ['Monitor\nใช้งานได้', 11, 1, C.MGREEN], ['Monitor\nเสื่อม\nสภาพ', 12, 1, C.YEL], ['Monitor\nเสีย', 13, 1, C.RED]].forEach(function (h) {
    op({ r: SZ, c: h[1], cs: h[2], v: h[0], fill: h[3], font: h[2] > 1 ? hdr : hdr7, align: mid, border: true });
  });
  heights[SZ] = 58;
  [['ชนิด Monitor', 2, 7, null], ['จำนวน', 10, 1, C.BLUE], ['Monitor\nใช้งานได้', 11, 1, C.MGREEN], ['Monitor\nเสื่อม\nสภาพ', 12, 1, C.YEL], ['Monitor\nเสีย', 13, 1, C.RED]].forEach(function (h) {
    op({ r: ST, c: h[1], cs: h[2], v: h[0], fill: h[3], font: h[2] > 1 ? hdr : hdr7, align: mid, border: true });
  });
  heights[ST] = 46;
  op({ r: 20, c: 2, v: 'วิธีใช้', font: { b: true, sz: 8 }, align: { v: 'middle' }, border: false });
  op({ r: 20, c: 3, cs: 13, v: 'เลือกสภาพจาก dropdown ในช่องใต้ชื่อที่นั่ง / Monitor / อุปกรณ์ Galley / ห้องน้ำ · พิมพ์รายละเอียดในช่องใต้สุด (NOTE) · สีและตารางสรุปเปลี่ยนเอง · ทุกการแก้ไขบันทึกในแท็บ ประวัติการแก้ไข', font: { sz: 8 }, align: { v: 'middle', wrap: true }, border: false });
  heights[20] = 30;

  var FIRST = 22, r = FIRST, roomTop = null, roomBot = null;
  var curZone = null, zoneStart = 0;
  function closeZone() {
    if (!curZone) return;
    zoneRows[curZone] = [zoneStart, r - 1];
    ops.unshift({ r: zoneStart, c: 2, rs: r - zoneStart, cs: 14, fill: C.CABIN, cabin: true });
  }
  L.blocks.forEach(function (b) {
    if (b.t === 'zone') {
      closeZone(); if (curZone) { heights[r] = 14; r++; }
      op({ r: r, c: 2, cs: 10, v: b.title, fill: C.ZONEH, font: { b: true, sz: 11 }, align: { v: 'middle' }, border: false });
      op({ r: r, c: 12, cs: 4, v: b.tag, fill: C.TAG, font: { b: true, sz: 10 }, align: mid, border: false });
      heights[r] = 28; r++; heights[r] = 8;
      curZone = b.zone; zoneStart = r; r++;
      return;
    }
    if (b.t === 'door' && b.exit) {
      var ef = { b: true, sz: 8, color: '#A33A00' };
      op({ r: r, c: 1, v: '◀ EMER\nEXIT\n' + b.code + ' L', fill: C.EXIT, font: ef, align: { h: 'left', v: 'middle', wrap: true }, border: false });
      op({ r: r, c: 16, v: 'EMER\nEXIT\n' + b.code + ' R ▶', fill: C.EXIT, font: ef, align: { h: 'right', v: 'middle', wrap: true }, border: false });
      heights[r] = 46; heights[r + 1] = 6; r += 2;   // 3 short lines, all visible
      return;
    }
    if (b.t === 'door') {
      var dc = b.exit ? C.EXIT : C.DOOR, dfont = { b: true, sz: 8, color: b.exit ? '#A33A00' : C.MUTED };
      /* door band runs the full width, edge to edge (A..P), like the emergency exit marks */
      op({ r: r, c: 1, cs: 8, v: b.left, fill: dc, font: dfont, align: { h: 'left', v: 'middle' }, border: false });
      op({ r: r, c: 9, cs: 1, fill: dc, border: false });
      op({ r: r, c: 10, cs: 7, v: b.right, fill: dc, font: dfont, align: { h: 'right', v: 'middle' }, border: false });
      heights[r] = 20; heights[r + 1] = 6; r += 2;
      return;
    }
    if (b.t === 'gtitle') {
      op({ r: r, c: b.c, cs: b.span, v: b.text, fill: C.GAL, font: { b: true, sz: 9 }, align: mid, border: true });
      heights[r] = 20; r++;
      return;
    }
    if (b.t === 'title') {
      op({ r: r, c: b.c, cs: b.span, v: b.text, font: { b: true, sz: 9, color: C.MUTED }, align: mid, border: false });
      if (roomTop === null) roomTop = r;
      heights[r] = 22; r++;
      return;
    }
    var isMonRow = b.b === 'Monitor' || b.b === 'OVHD' || b.b === 'LAV' || b.b === 'GALLEY' || b.b === 'EXIT' || b.b === '', isNum = /^\d+$/.test(b.b);
    if (b.b) op(isNum ? { r: r, c: 2, v: b.b, fill: C.ROWNO, font: { b: true, sz: 10, color: '#FFFFFF' }, align: mid, border: true }
                      : { r: r, c: 2, v: b.b, font: { b: true, sz: 7, color: C.MUTED }, align: { h: 'right', v: 'middle' }, border: false });
    if (b.cells.some(function (x) { return x.id; }) && !b.cells.some(function (x) { return x.c === 2; })) op({ r: r + 2, c: 2, v: 'NOTE', font: { sz: 7, color: C.MUTED }, align: { h: 'right', v: 'top' }, border: false });
    b.cells.forEach(function (x) {
      var span = x.span || 2;
      if (x.text !== undefined) { op({ r: r, c: x.c, rs: 3 + 4 * ((x.rows || 1) - 1), cs: span, v: x.text, fill: x.fill || C.TABLE, font: x.bold ? { b: true, sz: 8, color: '#A33A00' } : { i: true, sz: 8, color: C.MUTED }, align: mid, border: true }); return; }
      var it = L.items[x.id], isMon = it.kind === 'mon', isEq = it.kind === 'eq' || it.kind === 'lav';
      /* galley names: font small enough that the longest word fits the cell width (no "Microwav/e") */
      var lsz = isMon || isEq ? 8 : 11;
      if (it.kind === 'eq') {
        var wpx = 0; for (var cc = x.c; cc < x.c + span; cc++) wpx += widths[cc] || 52;
        var longest = Math.max.apply(null, String(it.label).split(/\s+/).map(function (w) { return w.length; }));
        lsz = Math.max(6, Math.min(8, Math.floor((wpx - 8) / (longest * 0.58 * 1.333) * 2) / 2));
      }
      op({ r: r, c: x.c, cs: span, v: it.label, fill: isMon ? SEAT_MT[it.type].color : it.kind === 'lav' ? C.LAV : C.GREEN, font: { b: true, sz: lsz }, align: mid, border: true });
      op({ r: r + 1, c: x.c, cs: span, v: isMon ? SEAT_M[0] : isEq ? SEAT_E[0] : SEAT_S[0], fill: C.WHITE, font: { sz: isEq ? 7 : 8 }, align: mid, border: true, dv: isMon ? 'M' : isEq ? 'E' : 'S' });
      op({ r: r + 2, c: x.c, rs: 1 + 4 * ((x.rows || 1) - 1), cs: span, fill: C.WHITE, font: { sz: 8 }, align: { h: 'left', v: 'top', wrap: true }, border: true });
      pos[x.id] = { label: [r, x.c], status: [r + 1, x.c], note: [r + 2, x.c] };
      if (isMon) (typeCells[it.type] = typeCells[it.type] || []).push(A(x.c) + (r + 1));
    });
    if (b.room) { if (roomTop === null) roomTop = r; roomBot = r + 2; }
    var eqRow = b.cells.some(function (x) { return x.id && (L.items[x.id].kind === 'eq' || L.items[x.id].kind === 'lav'); });
    heights[r] = isMonRow ? 34 : 26; heights[r + 1] = eqRow ? 30 : 18; heights[r + 2] = 34; heights[r + 3] = 6;
    r += 4;
  });
  closeZone();
  var LAST = r - 1;
  if (roomTop !== null) {
    boxes.push({ r1: roomTop, c1: 3, r2: roomBot, c2: 8 });
  }

  /* summary formulas; q tells the Excel export how to fill in the cached result */
  var all = L.order, seatsOf = function (z) { return all.filter(function (id) { return L.items[id].kind === 'seat' && (!z || L.items[id].zone === z); }); };
  var monsOf = function (z, t) { return all.filter(function (id) { var it = L.items[id]; return it.kind === 'mon' && (!z || it.zone === z) && (!t || it.type === t); }); };
  SEAT_ZONES.forEach(function (z, i) {
    var row = SZ + 1 + i, zr = zoneRows[z.key], rg = '$C$' + zr[0] + ':$O$' + zr[1];
    op({ r: row, c: 2, cs: 3, v: z.name, font: { b: true, sz: 8 }, align: { v: 'middle' }, border: true });
    op({ r: row, c: 5, v: '=F' + row + '+G' + row + '+H' + row, q: { ids: seatsOf(z.key), f: 's', vals: [0, 1, 2] }, font: { sz: 9 }, align: mid, border: true });
    SEAT_S.forEach(function (t, j) { op({ r: row, c: 6 + j, v: '=COUNTIF(' + rg + ',"' + t + '")', q: { ids: seatsOf(z.key), f: 's', vals: [j] }, fill: [C.GREEN, C.YEL, C.RED][j], font: { sz: 9 }, align: mid, border: true }); });
    op({ r: row, c: 10, v: '=K' + row + '+L' + row + '+M' + row, q: { ids: monsOf(z.key), f: 'm', vals: [0, 1, 2] }, fill: C.BLUE, font: { sz: 9 }, align: mid, border: true });
    SEAT_M.forEach(function (t, j) { op({ r: row, c: 11 + j, v: '=COUNTIF(' + rg + ',"' + t + '")', q: { ids: monsOf(z.key), f: 'm', vals: [j] }, fill: [C.MGREEN, C.YEL, C.RED][j], font: { sz: 9 }, align: mid, border: true }); });
    heights[row] = 22;
  });
  var tr = SZ + 1 + SEAT_ZONES.length;
  op({ r: tr, c: 2, cs: 3, v: 'รวม', font: { b: true, sz: 8 }, align: { v: 'middle' }, border: true });
  [[5, null, seatsOf(), 's', [0, 1, 2]], [6, C.GREEN, seatsOf(), 's', [0]], [7, C.YEL, seatsOf(), 's', [1]], [8, C.RED, seatsOf(), 's', [2]],
   [10, C.BLUE, monsOf(), 'm', [0, 1, 2]], [11, C.MGREEN, monsOf(), 'm', [0]], [12, C.YEL, monsOf(), 'm', [1]], [13, C.RED, monsOf(), 'm', [2]]].forEach(function (x) {
    op({ r: tr, c: x[0], v: '=SUM(' + A(x[0]) + (SZ + 1) + ':' + A(x[0]) + (tr - 1) + ')', q: { ids: x[2], f: x[3], vals: x[4] }, fill: x[1], font: { b: true, sz: 9 }, align: mid, border: true });
  });
  heights[tr] = 22;
  SEAT_MT_ORDER.forEach(function (t, i) {
    var row = ST + 1 + i, cells = typeCells[t] || [];
    op({ r: row, c: 2, cs: 7, v: SEAT_MT[t].name, fill: SEAT_MT[t].color, font: { b: true, sz: 8 }, align: { v: 'middle' }, border: true });
    op({ r: row, c: 10, v: '=K' + row + '+L' + row + '+M' + row, q: { ids: monsOf(null, t), f: 'm', vals: [0, 1, 2] }, fill: C.BLUE, font: { sz: 9 }, align: mid, border: true });
    SEAT_M.forEach(function (s, j) {
      op({ r: row, c: 11 + j, v: '=' + cells.map(function (a) { return 'N(' + a + '="' + s + '")'; }).join('+'), q: { ids: monsOf(null, t), f: 'm', vals: [j] }, fill: [C.MGREEN, C.YEL, C.RED][j], font: { sz: 9 }, align: mid, border: true });
    });
    heights[row] = 22;
  });
  var tt = ST + 1 + SEAT_MT_ORDER.length;
  op({ r: tt, c: 2, cs: 7, v: 'รวม', font: { b: true, sz: 8 }, align: { v: 'middle' }, border: true });
  [[10, C.BLUE, [0, 1, 2]], [11, C.MGREEN, [0]], [12, C.YEL, [1]], [13, C.RED, [2]]].forEach(function (x) {
    op({ r: tt, c: x[0], v: '=SUM(' + A(x[0]) + (ST + 1) + ':' + A(x[0]) + (tt - 1) + ')', q: { ids: monsOf(), f: 'm', vals: x[2] }, fill: x[1], font: { b: true, sz: 9 }, align: mid, border: true });
  });
  heights[tt] = 22;

  /* colours: a status cell by its own value, a name cell by the status just below it */
  var area = 'C' + FIRST + ':O' + LAST, nx = 'C' + (FIRST + 1);
  cf.push({ range: area, text: SEAT_S[2], fill: C.RED, font: '#FFFFFF' });
  cf.push({ range: area, text: SEAT_M[2], fill: C.RED, font: '#FFFFFF' });
  cf.push({ range: area, text: SEAT_S[1], fill: C.YEL });
  cf.push({ range: area, text: SEAT_M[1], fill: C.YEL });
  cf.push({ range: area, text: SEAT_E[2], fill: C.RED, font: '#FFFFFF' });
  cf.push({ range: area, text: SEAT_E[1], fill: C.YEL });
  cf.push({ range: area, f: '=OR(' + nx + '="' + SEAT_S[2] + '",' + nx + '="' + SEAT_M[2] + '",' + nx + '="' + SEAT_E[2] + '")', fill: C.RED, font: '#FFFFFF' });
  cf.push({ range: area, f: '=OR(' + nx + '="' + SEAT_S[1] + '",' + nx + '="' + SEAT_M[1] + '",' + nx + '="' + SEAT_E[1] + '")', fill: C.YEL });

  return { ops: ops, widths: widths, heights: heights, cf: cf, boxes: boxes, pos: pos, zoneRows: zoneRows, firstMapRow: FIRST, lastRow: LAST, lastCol: LASTC };
}

if (typeof window !== 'undefined') { window.seatPlanLayout = seatPlanLayout; window.planSheet = planSheet; }
