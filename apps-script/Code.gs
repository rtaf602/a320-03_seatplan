/**
 * Seat Plan 320(03) (For Admin) — Apps Script
 *
 * ทำ 3 อย่าง:
 *  1) setupSheet(): สร้างแท็บ "ผังที่นั่ง", "รวมข้อขัดข้องคงค้าง", "ทุกที่นั่ง และ Monitor", "ประวัติการแก้ไข"
 *     (รันครั้งแรกครั้งเดียว หรือรันซ้ำเพื่อจัดผังใหม่ ค่าสภาพและหมายเหตุเดิมยังอยู่)
 *  2) เป็น API ให้หน้าเว็บ (GitHub Pages) อ่าน/บันทึกสภาพที่นั่งลงแท็บ "ผังที่นั่ง"
 *     และบันทึกประวัติลงแท็บ "ประวัติการแก้ไข" ทุกครั้งที่กด SAVE บนเว็บ
 *  3) บันทึกประวัติอัตโนมัติเมื่อมีคนแก้ไขช่องในชีทโดยตรง (onEdit)
 *
 * ตำแหน่งช่องของแต่ละที่นั่งมาจากผังใน export-excel.js (คัดลอกไว้ท้ายไฟล์นี้)
 * ห้ามแทรก/ลบแถวหรือคอลัมน์ในแท็บ "ผังที่นั่ง" เอง ถ้าต้องการจัดผังใหม่ให้แก้ export-excel.js แล้วรัน setupSheet
 *
 * วิธีติดตั้ง: ดู README.md ใน repo
 */

// ---------- ตั้งค่า ----------
const MAP_SHEET = 'ผังที่นั่ง';
const BAD_SHEET = 'รวมข้อขัดข้องคงค้าง';
const LIST_SHEET = 'ทุกที่นั่ง และ Monitor';
const LOG_SHEET = 'ประวัติการแก้ไข';
const INSPECTOR_CELL = 'D3';   // ช่องชื่อผู้ตรวจในชีท
const WRITE_PIN = '';          // รหัสเข้าหน้าเว็บและ SAVE เช่น '6020' (เว้นว่าง = ไม่ต้องใช้รหัส)
const ADMIN_PIN = '';          // รหัสสำหรับปุ่ม "ล้างข้อมูลทั้งหมด" บนเว็บ (เว้นว่าง = ปิดปุ่มนี้)

const SEAT_TXT = ['ปกติ', 'ขัดข้อง-ใช้ได้', 'ใช้ไม่ได้'];
const MON_TXT = ['จอปกติ', 'จอเสื่อม', 'จอเสีย'];
const LOG_HDR = ['วัน/เวลา (TH-LT)', 'อีเมลผู้แก้ไข', 'ผู้ตรวจ (ช่อง D3)', 'ที่นั่ง / จุด', 'รายการ', 'ค่าเดิม', 'ค่าใหม่'];

// ---------- Web API ----------
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'load';
  try {
    if (action === 'load') {
      if (WRITE_PIN) return json_({ ok: false, error: 'pin' });  // มีรหัส: ต้องโหลดผ่าน POST พร้อมรหัส
      return json_(loadData_());
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err) });
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    if (body.action === 'load') {
      if (WRITE_PIN && String(body.pin || '') !== WRITE_PIN) return json_({ ok: false, error: 'pin' });
      return json_(loadData_());
    }
    if (body.action === 'history') {
      if (WRITE_PIN && String(body.pin || '') !== WRITE_PIN) return json_({ ok: false, error: 'pin' });
      return json_(historyData_());
    }
    if (body.action === 'save') return json_(saveData_(body));
    if (body.action === 'reset') return json_(resetData_(body));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- ผังและตำแหน่งช่อง ----------
let LAYOUT_ = null;
/** ผังจาก export-excel.js: ที่นั่ง/Monitor ทุกตัวพร้อมตำแหน่งช่อง [แถว, คอลัมน์] */
function layout_() {
  if (LAYOUT_) return LAYOUT_;
  const wb = buildSeatWorkbook(ShimExcel_, {});
  const items = {};
  wb.seatOrder.forEach(function (id) {
    const c = wb.seatCells[id];
    items[id] = { id: id, label: c.label, zone: c.zone, kind: c.kind, type: c.type || '',
      seat: c.seat ? a1rc_(c.seat) : null, mon: c.mon ? a1rc_(c.mon) : null, note: a1rc_(c.note) };
  });
  LAYOUT_ = { wb: wb, items: items, order: wb.seatOrder, last: wb.seatLastRow };
  return LAYOUT_;
}
function a1rc_(a1) {
  const m = /^([A-Z]+)(\d+)$/.exec(a1); let c = 0;
  for (let i = 0; i < m[1].length; i++) c = c * 26 + m[1].charCodeAt(i) - 64;
  return [Number(m[2]), c];
}
function rcA1_(r, c) { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s + r; }

/** อ่านค่าปัจจุบันในแท็บผังที่นั่ง */
function buildMap_(sh) {
  const L = layout_();
  const v = sh.getRange(1, 1, L.last, 20).getDisplayValues();
  const cell = function (r, c) { return (r >= 1 && r <= L.last) ? String(v[r - 1][c - 1]).trim() : ''; };
  return { items: L.items, cell: cell };
}
function mapSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(MAP_SHEET);
  if (!sh) throw new Error('ยังไม่มีแท็บ "' + MAP_SHEET + '" ให้รัน setupSheet ใน Apps Script ก่อน');
  return sh;
}

function loadData_() {
  const ss = SpreadsheetApp.getActive();
  const m = buildMap_(mapSheet_());
  const byLabel = {};
  Object.keys(m.items).forEach(function (id) { byLabel[m.items[id].label] = id; });

  // ผู้แก้ไข/เวลาล่าสุดของแต่ละที่นั่ง จากแท็บประวัติ
  const lastBy = {};
  const log = ss.getSheetByName(LOG_SHEET);
  const n = log ? log.getLastRow() : 0;
  if (n > 1) {
    const rows = log.getRange(2, 1, n - 1, 4).getValues();
    rows.forEach(function (row) {
      const id = byLabel[String(row[3])];
      if (!id) return;
      const at = row[0] instanceof Date ? row[0].toISOString() : '';
      const who = (row[2] && row[2] !== '-') ? String(row[2]) : String(row[1] || '');
      lastBy[id] = { by: who, at: at };
    });
  }

  const seats = {};
  Object.keys(m.items).forEach(function (id) {
    const it = m.items[id];
    const s = it.seat ? Math.max(0, SEAT_TXT.indexOf(m.cell(it.seat[0], it.seat[1]))) : 0;
    const mo = it.mon ? Math.max(0, MON_TXT.indexOf(m.cell(it.mon[0], it.mon[1]))) : 0;
    const note = m.cell(it.note[0], it.note[1]);
    const lb = lastBy[id] || {};
    seats[id] = { s: s, m: mo, note: note, by: lb.by || '', at: lb.at || '' };
  });
  const ls = PropertiesService.getScriptProperties().getProperty('lastSave');
  return { ok: true, seats: seats, lastSave: ls ? JSON.parse(ls) : null, pinRequired: !!WRITE_PIN, resetEnabled: !!ADMIN_PIN };
}

/** ทุกบรรทัดในแท็บประวัติ (สำหรับ EXPORT TO EXCEL) วันเวลาส่งเป็น ISO */
function historyData_() {
  const log = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
  const n = log ? log.getLastRow() : 0;
  if (n < 2) return { ok: true, rows: [] };
  const rows = log.getRange(2, 1, n - 1, 7).getValues().map(function (r) {
    return r.map(function (v, j) { return (j === 0 && v instanceof Date) ? v.toISOString() : (v === null ? '' : String(v)); });
  });
  return { ok: true, rows: rows };
}

/**
 * SAVE จากหน้าเว็บ
 * แต่ละรายการ: { id, s, m, note, fields: ['s','m','note'] ที่แก้จริง, base: {s, m, note} ค่าตอนเริ่มแก้ }
 * - เขียนเฉพาะช่องใน fields ช่องที่ไม่ได้แก้จะไม่ถูกแตะ (กันทับค่าที่คนอื่นเพิ่งบันทึก)
 * - ถ้าค่าในชีทตอนนี้ไม่ตรงกับ base และไม่ตรงกับค่าใหม่ = มีคนแก้ช่องเดียวกันไปก่อน
 *   จะไม่บันทึกช่องนั้น แต่ส่งกลับใน conflicts ให้หน้าเว็บถามผู้ใช้
 */
function saveData_(body) {
  if (WRITE_PIN && String(body.pin || '') !== WRITE_PIN) return { ok: false, error: 'pin' };
  const by = String(body.by || '').trim().slice(0, 60);
  if (!by) return { ok: false, error: 'no_name' };
  const changes = Array.isArray(body.changes) ? body.changes.slice(0, 500) : [];

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = mapSheet_();
    const m = buildMap_(sh);
    const now = new Date();
    const logRows = [];
    const conflicts = [];
    const current = {};
    let saved = 0;
    let logCache = null;

    // ผู้แก้ล่าสุดของช่องนี้ จากแท็บประวัติ (อ่านเฉพาะเมื่อมีการชน)
    const lastEdit = function (label, fieldTh) {
      if (!logCache) {
        const log = ss.getSheetByName(LOG_SHEET);
        const n = log ? log.getLastRow() : 0;
        logCache = n > 1 ? log.getRange(2, 1, n - 1, 5).getValues() : [];
      }
      for (let i = logCache.length - 1; i >= 0; i--) {
        const r = logCache[i];
        if (String(r[3]) === label && String(r[4]) === fieldTh) {
          return { by: (r[2] && r[2] !== '-') ? String(r[2]) : String(r[1] || ''), at: r[0] instanceof Date ? r[0].toISOString() : '' };
        }
      }
      return { by: '', at: '' };
    };

    changes.forEach(function (ch) {
      const it = m.items[String(ch.id)];
      if (!it) return;
      const fields = Array.isArray(ch.fields) ? ch.fields : ['s', 'm', 'note'];
      const base = ch.base && typeof ch.base === 'object' ? ch.base : null;
      const cur = {
        s: it.seat ? Math.max(0, SEAT_TXT.indexOf(m.cell(it.seat[0], it.seat[1]))) : 0,
        m: it.mon ? Math.max(0, MON_TXT.indexOf(m.cell(it.mon[0], it.mon[1]))) : 0,
        note: m.cell(it.note[0], it.note[1])
      };
      const spec = {
        s: it.seat ? { pos: it.seat, th: 'สภาพที่นั่ง', val: clamp_(ch.s), txt: function (x) { return SEAT_TXT[x]; } } : null,
        m: it.mon ? { pos: it.mon, th: 'สภาพ Monitor', val: clamp_(ch.m), txt: function (x) { return MON_TXT[x]; } } : null,
        note: { pos: it.note, th: 'หมายเหตุ', val: String(ch.note || '').trim().slice(0, 500), txt: function (x) { return x; } }
      };
      let touched = false;
      ['s', 'm', 'note'].forEach(function (f) {
        const sp = spec[f];
        if (!sp || fields.indexOf(f) < 0) return;
        if (cur[f] === sp.val) return;
        if (base && f in base) {
          const b = f === 'note' ? String(base[f] || '').trim() : clamp_(base[f]);
          if (cur[f] !== b) {
            const le = lastEdit(it.label, sp.th);
            conflicts.push({ id: it.id, label: it.label, field: f, cur: cur[f], mine: sp.val, by: le.by, at: le.at });
            return;
          }
        }
        const oldV = m.cell(sp.pos[0], sp.pos[1]);
        const newV = sp.txt(sp.val);
        sh.getRange(sp.pos[0], sp.pos[1]).setValue(newV);
        logRows.push([now, 'หน้าเว็บ', by, it.label, sp.th, oldV, newV]);
        cur[f] = sp.val;
        touched = true;
      });
      current[it.id] = cur;
      if (touched) saved++;
    });
    appendLog_(ss, logRows);
    const props = PropertiesService.getScriptProperties();
    let lastSave;
    if (saved) {
      lastSave = { by: by, at: now.toISOString(), n: saved };
      props.setProperty('lastSave', JSON.stringify(lastSave));
    } else {
      const ls = props.getProperty('lastSave');
      lastSave = ls ? JSON.parse(ls) : null;
    }
    return { ok: true, saved: saved, at: now.toISOString(), lastSave: lastSave, conflicts: conflicts, current: current };
  } finally {
    lock.releaseLock();
  }
}

function resetData_(body) {
  if (!ADMIN_PIN) return { ok: false, error: 'reset_disabled' };
  if (String(body.adminPin || '') !== ADMIN_PIN) return { ok: false, error: 'admin_pin' };
  const by = String(body.by || '').trim().slice(0, 60) || '-';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = mapSheet_();
    writeState_(sh, {});
    const now = new Date();
    appendLog_(ss, [[now, 'หน้าเว็บ', by, 'ทั้งลำ', 'ล้างข้อมูลเริ่มรอบตรวจใหม่', '', 'ปกติทั้งหมด']]);
    const lastSave = { by: by, at: now.toISOString(), n: 0 };
    PropertiesService.getScriptProperties().setProperty('lastSave', JSON.stringify(lastSave));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** เขียนสภาพทุกที่นั่งลงชีท state = { id: {s, m, note} } (ที่ไม่มีใน state = ปกติ, หมายเหตุว่าง) */
function writeState_(sh, state) {
  const L = layout_();
  Object.keys(L.items).forEach(function (id) {
    const it = L.items[id], d = state[id] || {};
    if (it.seat) sh.getRange(it.seat[0], it.seat[1]).setValue(SEAT_TXT[clamp_(d.s)]);
    if (it.mon) sh.getRange(it.mon[0], it.mon[1]).setValue(MON_TXT[clamp_(d.m)]);
    sh.getRange(it.note[0], it.note[1]).setValue(d.note || '');
  });
}

function clamp_(x) { x = Number(x) | 0; return x < 0 ? 0 : (x > 2 ? 2 : x); }

function appendLog_(ss, rows) {
  if (!rows.length) return;
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log) return;
  const start = log.getLastRow() + 1;
  log.getRange(start, 1, rows.length, 7).setValues(rows);
  log.getRange(start, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

// ---------- ประวัติเมื่อแก้ไขในชีทโดยตรง ----------
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== MAP_SHEET) return;

  const r0 = e.range.getRow(), c0 = e.range.getColumn();
  const nr = e.range.getNumRows(), nc = e.range.getNumColumns();
  const idx = cellIndex_();
  const values = e.range.getDisplayValues();
  const single = nr === 1 && nc === 1;
  const inspector = sh.getRange(INSPECTOR_CELL).getDisplayValue();
  let email = '';
  try { email = (e.user && e.user.getEmail()) || Session.getActiveUser().getEmail(); } catch (err) {}
  const now = new Date();

  const rows = [];
  for (let i = 0; i < nr; i++) {
    for (let j = 0; j < nc; j++) {
      const info = identify_(idx, r0 + i, c0 + j);
      if (!info) continue;
      const oldV = single ? (e.oldValue === undefined ? '' : e.oldValue) : '(แก้หลายช่องพร้อมกัน)';
      const newV = single ? (e.value === undefined ? values[0][0] : e.value) : values[i][j];
      rows.push([now, email || '-', inspector || '-', info.id, info.field, oldV, newV]);
    }
  }
  appendLog_(e.source, rows);
}

/** ช่อง -> ที่นั่ง/Monitor และชนิดช่อง */
function cellIndex_() {
  const L = layout_(), idx = {};
  Object.keys(L.items).forEach(function (id) {
    const it = L.items[id];
    if (it.seat) idx[it.seat.join(',')] = { id: it.label, field: 'สภาพที่นั่ง' };
    if (it.mon) idx[it.mon.join(',')] = { id: it.label, field: 'สภาพ Monitor' };
    idx[it.note.join(',')] = { id: it.label, field: 'หมายเหตุ' };
  });
  return idx;
}
function identify_(idx, r, c) {
  if (r === 3 && c >= 4 && c <= 10) return { id: 'ส่วนหัว', field: 'ผู้ตรวจ' };
  if (r === 3 && c >= 14) return { id: 'ส่วนหัว', field: 'วันที่ตรวจ' };
  const hit = idx[r + ',' + c];
  if (hit) return hit;
  if (r < 20) return null;
  return { id: rcA1_(r, c), field: 'อื่น ๆ' };
}

// ---------- เมนูในชีท ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Seat Plan')
    .addItem('สร้าง / จัดผังที่นั่งใหม่ (ค่าเดิมยังอยู่)', 'setupSheet')
    .addToUi();
}

// ---------- สร้างแท็บทั้งหมด ----------
/**
 * สร้างหรือจัดแท็บใหม่ตามผังใน export-excel.js
 * ถ้ามีแท็บผังที่นั่งอยู่แล้ว จะอ่านสภาพ/หมายเหตุเดิมไว้ แล้วเขียนกลับหลังจัดผัง
 * แท็บประวัติการแก้ไขไม่ถูกลบ
 */
function setupSheet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    ss.setSpreadsheetTimeZone('Asia/Bangkok');
    const L = layout_();
    const ws = L.wb.worksheets[0];

    // เก็บค่าเดิม
    let keep = null;
    let sh = ss.getSheetByName(MAP_SHEET);
    if (sh && sh.getLastRow() > 20) {
      // อ่านจากแท็บ "ทุกที่นั่ง และ Monitor" ตามชื่อที่นั่ง (สูตรในแท็บนั้นยังชี้ช่องตามผังเดิม จึงถูกต้องแม้ย้ายตำแหน่งช่อง)
      keep = keepFromList_(ss);
      if (!keep) { try { keep = loadData_().seats; } catch (err) { keep = null; } }
    }
    if (!sh) sh = ss.insertSheet(MAP_SHEET, 0);
    applyWorksheet_(sh, ws, L.wb.seatFrames);

    // ส่วนหัวในชีท: SAVE ล่าสุดดึงจากแท็บประวัติ
    const H = "'" + LOG_SHEET + "'!";
    sh.getRange('D2').setFormula('=IFERROR(INDEX(FILTER(' + H + 'C2:C,' + H + 'B2:B="หน้าเว็บ"),COUNTIF(' + H + 'B2:B,"หน้าเว็บ")),"-")');
    sh.getRange('N2').setFormula('=LET(t,IFERROR(INDEX(FILTER(' + H + 'A2:A,' + H + 'B2:B="หน้าเว็บ"),COUNTIF(' + H + 'B2:B,"หน้าเว็บ")),""),IF(t="","-",UPPER(TEXT(t,"dd mmm"))&TEXT(t,"yy")&" ("&TEXT(t,"hh:mm")&" TH-LT)"))');
    sh.getRange('N3').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange('B4').setValue('บันทึกสภาพที่นั่งและ Monitor · สำหรับ Admin ฝ่ายการช่าง และฝ่ายสื่อสาร · เวลาทั้งหมดเป็นเวลาประเทศไทย (TH-LT)');
    sh.getRange('C18').setValue('เลือกสภาพจาก dropdown ใต้ชื่อที่นั่งและใต้ป้าย Monitor · พิมพ์รายละเอียดในช่อง NOTE · สีและตารางสรุปเปลี่ยนเอง · ทุกการแก้ไขบันทึกในแท็บ ประวัติการแก้ไข · Monitor = จอติดผนัง 21.5 นิ้ว, OVHD Monitor = จอพับเหนือที่นั่ง 12.0 นิ้ว');
    if (keep) writeState_(sh, keep);

    setupLog_(ss);
    setupList_(ss);
    setupBad_(ss);
    // ลำดับแท็บ
    [MAP_SHEET, BAD_SHEET, LIST_SHEET, LOG_SHEET].forEach(function (name, i) {
      ss.setActiveSheet(ss.getSheetByName(name)); ss.moveActiveSheet(i + 1);
    });
    // ลบแท็บว่างที่มากับไฟล์ใหม่
    ss.getSheets().forEach(function (s) {
      if ([MAP_SHEET, BAD_SHEET, LIST_SHEET, LOG_SHEET].indexOf(s.getName()) < 0 && s.getLastRow() === 0 && s.getLastColumn() === 0) ss.deleteSheet(s);
    });
    ss.setActiveSheet(sh);
  } finally {
    lock.releaseLock();
  }
}

/** สภาพเดิมของทุกที่นั่งจากแท็บ "ทุกที่นั่ง และ Monitor" (จับคู่ด้วยชื่อในคอลัมน์ A) */
function keepFromList_(ss) {
  const s = ss.getSheetByName(LIST_SHEET);
  if (!s || s.getLastRow() < 2) return null;
  const L = layout_(), byLabel = {};
  Object.keys(L.items).forEach(function (id) { byLabel[L.items[id].label] = id; });
  const out = {};
  let n = 0;
  s.getRange(2, 1, s.getLastRow() - 1, 6).getDisplayValues().forEach(function (r) {
    const id = byLabel[String(r[0]).trim()];
    if (!id) return;
    out[id] = { s: Math.max(0, SEAT_TXT.indexOf(String(r[3]).trim())), m: Math.max(0, MON_TXT.indexOf(String(r[4]).trim())), note: String(r[5]).trim() };
    n++;
  });
  return n ? out : null;
}

function hex_(argb) { return argb ? '#' + String(argb).slice(-6) : null; }

/** วาดแท็บจากผัง (worksheet ที่ได้จาก ShimExcel_) */
function applyWorksheet_(sh, ws, frames) {
  const cells = ws.cells, nR = ws.maxRow(), nC = 20;
  // ล้างแท็บ
  if (sh.getMaxRows() < nR + 5) sh.insertRowsAfter(sh.getMaxRows(), nR + 5 - sh.getMaxRows());
  if (sh.getMaxRows() > nR + 5) sh.deleteRows(nR + 6, sh.getMaxRows() - nR - 5);
  if (sh.getMaxColumns() < nC + 1) sh.insertColumnsAfter(sh.getMaxColumns(), nC + 1 - sh.getMaxColumns());
  if (sh.getMaxColumns() > nC + 1) sh.deleteColumns(nC + 2, sh.getMaxColumns() - nC - 1);
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart(); all.clearDataValidations(); sh.clearConditionalFormatRules(); sh.clear();

  const grid = function (v) { const a = []; for (let r = 0; r < nR; r++) { const row = []; for (let c = 0; c < nC; c++) row.push(v); a.push(row); } return a; };
  const val = grid(''), bg = grid(null), fw = grid('normal'), fs = grid(10), fc = grid('#000000'), fst = grid('normal'),
    ha = grid('left'), va = grid('bottom'), wr = grid(false), dvs = grid(null);
  const dvCache = {}, thinA1 = [];
  const merged = {};
  ws.merges.forEach(function (m) { merged[m[0] + ',' + m[1]] = m; });
  const covered = {};
  ws.merges.forEach(function (m) { for (let r = m[0]; r <= m[2]; r++) for (let c = m[1]; c <= m[3]; c++) if (r !== m[0] || c !== m[1]) covered[r + ',' + c] = 1; });

  cells.forEach(function (x) {
    const r = x.r - 1, c = x.c - 1;
    if (r >= nR || c >= nC) return;
    let v = x.value;
    if (v && typeof v === 'object' && v.formula) v = '=' + v.formula;
    if (v !== undefined && v !== null) val[r][c] = v;
    if (x.fill && x.fill.fgColor) bg[r][c] = hex_(x.fill.fgColor.argb);
    if (x.font) {
      if (x.font.bold) fw[r][c] = 'bold';
      if (x.font.italic) fst[r][c] = 'italic';
      if (x.font.size) fs[r][c] = x.font.size + 1;
      if (x.font.color && x.font.color.argb) fc[r][c] = hex_(x.font.color.argb);
    }
    if (x.alignment) {
      if (x.alignment.horizontal) ha[r][c] = x.alignment.horizontal;
      if (x.alignment.vertical) va[r][c] = x.alignment.vertical === 'middle' ? 'middle' : x.alignment.vertical;
      if (x.alignment.wrapText) wr[r][c] = true;
    }
    if (x.dataValidation && x.dataValidation.type === 'list') {
      const key = x.dataValidation.formulae[0];
      if (!dvCache[key]) dvCache[key] = SpreadsheetApp.newDataValidation()
        .requireValueInList(key.replace(/^"|"$/g, '').split(','), true).setAllowInvalid(false).build();
      dvs[r][c] = dvCache[key];
    }
    const b = x.border;
    if (b && b.top && b.left && b.bottom && b.right && !covered[x.r + ',' + x.c]) {
      const m = merged[x.r + ',' + x.c];
      thinA1.push(m ? rcA1_(m[0], m[1]) + ':' + rcA1_(m[2], m[3]) : rcA1_(x.r, x.c));
    }
  });
  const rg = sh.getRange(1, 1, nR, nC);
  rg.setValues(val);
  rg.setBackgrounds(bg); rg.setFontWeights(fw); rg.setFontSizes(fs); rg.setFontColors(fc); rg.setFontStyles(fst);
  rg.setHorizontalAlignments(ha); rg.setVerticalAlignments(va); rg.setWraps(wr);
  rg.setDataValidations(dvs);
  ws.merges.forEach(function (m) { sh.getRange(m[0], m[1], m[2] - m[0] + 1, m[3] - m[1] + 1).merge(); });
  for (let i = 0; i < thinA1.length; i += 200) {
    sh.getRangeList(thinA1.slice(i, i + 200)).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }
  (frames || []).forEach(function (f) {
    sh.getRange(f[0], f[2], f[1] - f[0] + 1, f[3] - f[2] + 1).setBorder(true, true, true, true, null, null, '#566575', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  // สีตามสถานะ
  const rules = [];
  ws.cf.forEach(function (o) {
    const range = sh.getRange(o.ref);
    o.rules.forEach(function (ru) {
      let b = SpreadsheetApp.newConditionalFormatRule().setRanges([range]);
      const f = ru.formulae[0];
      if (ru.type === 'expression') b = b.whenFormulaSatisfied('=' + f);
      else if (ru.type === 'cellIs' && ru.operator === 'equal') b = b.whenTextEqualTo(String(f).replace(/^"|"$/g, ''));
      else return;
      if (ru.style.fill && ru.style.fill.bgColor) b = b.setBackground(hex_(ru.style.fill.bgColor.argb));
      if (ru.style.font && ru.style.font.color) b = b.setFontColor(hex_(ru.style.font.color.argb));
      rules.push(b.build());
    });
  });
  sh.setConditionalFormatRules(rules);

  for (let c = 1; c <= nC + 1; c++) sh.setColumnWidth(c, Math.round(((ws.cols[c] || {}).width || 2.5) * 7));
  for (let r = 1; r <= nR; r++) sh.setRowHeight(r, Math.max(6, Math.round(((ws.rows[r] || {}).height || 15) * 1.33)));
  sh.setHiddenGridlines(true);
}

function headerRow_(s, titles, widths) {
  s.getRange(1, 1, 1, titles.length).setValues([titles]).setFontWeight('bold').setBackground('#C1E4F5')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  widths.forEach(function (w, j) { s.setColumnWidth(j + 1, w); });
  s.setFrozenRows(1);
}
function statusRules_(s, a1) {
  const rg = s.getRange(a1);
  const mk = function (t, bg, fg) { let b = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setRanges([rg]); if (fg) b = b.setFontColor(fg); return b.build(); };
  s.setConditionalFormatRules([mk(SEAT_TXT[1], '#FFFF00'), mk(MON_TXT[1], '#FFFF00'), mk(SEAT_TXT[2], '#FF0000', '#FFFFFF'), mk(MON_TXT[2], '#FF0000', '#FFFFFF')]);
}
const LIST_HDR = ['ที่นั่ง / จุด', 'ประเภท', 'โซน', 'สภาพที่นั่ง', 'สภาพ Monitor', 'หมายเหตุ', 'ผู้บันทึกล่าสุด', 'วัน/เวลาล่าสุด (TH-LT)'];
const LIST_W = [250, 100, 130, 110, 100, 300, 170, 140];
const ZONE_NAME = { F: 'FRONT CREW ZONE', A: 'VVIP OA', B: 'VVIP OB', C: 'VVIP OC', D: 'PASSENGER OD' };

function setupLog_(ss) {
  let s = ss.getSheetByName(LOG_SHEET);
  if (!s) s = ss.insertSheet(LOG_SHEET);
  if (s.getLastRow() === 0) {
    headerRow_(s, LOG_HDR, [150, 220, 170, 250, 110, 220, 220]);
    s.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
}
function setupList_(ss) {
  let s = ss.getSheetByName(LIST_SHEET);
  if (!s) s = ss.insertSheet(LIST_SHEET);
  s.clear(); s.clearConditionalFormatRules();
  headerRow_(s, LIST_HDR, LIST_W);
  const L = layout_(), q = "'" + MAP_SHEET + "'!", h = "'" + LOG_SHEET + "'!";
  const rows = L.order.map(function (id, i) {
    const it = L.items[id], r = i + 2;
    return [it.label, it.seat ? 'ที่นั่ง' : (it.type === 'R' ? 'OVHD Monitor' : 'Monitor'), ZONE_NAME[it.zone],
      it.seat ? '=' + q + rcA1_(it.seat[0], it.seat[1]) : '-',
      it.mon ? '=' + q + rcA1_(it.mon[0], it.mon[1]) : '-',
      '=' + q + rcA1_(it.note[0], it.note[1]) + '&""',
      '=IFERROR(INDEX(FILTER(' + h + '$C$2:$C,' + h + '$D$2:$D=$A' + r + '),COUNTIF(' + h + '$D$2:$D,$A' + r + ')),"")',
      '=IFERROR(INDEX(FILTER(' + h + '$A$2:$A,' + h + '$D$2:$D=$A' + r + '),COUNTIF(' + h + '$D$2:$D,$A' + r + ')),"")'];
  });
  const rg = s.getRange(2, 1, rows.length, 8);
  rg.setValues(rows).setVerticalAlignment('top').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  s.getRange(2, 2, rows.length, 4).setHorizontalAlignment('center');
  s.getRange(2, 6, rows.length, 1).setWrap(true);
  s.getRange(2, 8, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm').setHorizontalAlignment('center');
  statusRules_(s, 'D2:E' + (rows.length + 1));
}
function setupBad_(ss) {
  let s = ss.getSheetByName(BAD_SHEET);
  if (!s) s = ss.insertSheet(BAD_SHEET);
  s.clear(); s.clearConditionalFormatRules();
  headerRow_(s, LIST_HDR, LIST_W);
  const q = "'" + LIST_SHEET + "'!";
  s.getRange('A2').setFormula('=IFERROR(FILTER(' + q + 'A2:H,(' + q + 'D2:D="' + SEAT_TXT[1] + '")+(' + q + 'D2:D="' + SEAT_TXT[2] + '")+(' + q + 'E2:E="' + MON_TXT[1] + '")+(' + q + 'E2:E="' + MON_TXT[2] + '")+(' + q + 'F2:F<>"")),"ไม่มีข้อขัดข้องคงค้าง")');
  s.getRange('A2:H200').setVerticalAlignment('top');
  s.getRange('B2:E200').setHorizontalAlignment('center');
  s.getRange('F2:F200').setWrap(true);
  s.getRange('H2:H200').setNumberFormat('dd/MM/yyyy HH:mm').setHorizontalAlignment('center');
  statusRules_(s, 'D2:E200');
}

// ---------- ตัวรับผังจาก export-excel.js (ทำหน้าที่แทน ExcelJS) ----------
const ShimExcel_ = (function () {
  function a1(s) { const m = /^([A-Z]+)(\d+)$/.exec(s); let c = 0; for (let i = 0; i < m[1].length; i++) c = c * 26 + m[1].charCodeAt(i) - 64; return [Number(m[2]), c]; }
  function WS(name) { this.name = name; this.map = {}; this.cells = []; this.cols = {}; this.rows = {}; this.merges = []; this.cf = []; this.views = []; this.autoFilter = null; }
  WS.prototype.getCell = function (a, b) {
    let r = a, c = b; if (typeof a === 'string') { const x = a1(a); r = x[0]; c = x[1]; }
    const k = r + ',' + c; if (!this.map[k]) { this.map[k] = { r: r, c: c }; this.cells.push(this.map[k]); } return this.map[k];
  };
  WS.prototype.getColumn = function (c) { return this.cols[c] = this.cols[c] || {}; };
  WS.prototype.getRow = function (r) { return this.rows[r] = this.rows[r] || {}; };
  WS.prototype.mergeCells = function (a, b, c, d) {
    if (typeof a === 'string') { const p = a.split(':'), x = a1(p[0]), y = a1(p[1]); a = x[0]; b = x[1]; c = y[0]; d = y[1]; }
    this.merges.push([a, b, c, d]);
  };
  WS.prototype.addConditionalFormatting = function (o) { this.cf.push(o); };
  WS.prototype.maxRow = function () { let m = 0; this.cells.forEach(function (x) { if (x.r > m) m = x.r; }); return m; };
  function Workbook() { this.worksheets = []; }
  Workbook.prototype.addWorksheet = function (n) { const w = new WS(n); this.worksheets.push(w); return w; };
  return { Workbook: Workbook };
})();

// ===== LAYOUT: copied from export-excel.js by tools/build-code-gs.js (do not edit here) =====
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
  var ZN = { F: 'FRONT CREW ZONE', A: 'VVIP OA', B: 'VVIP OB', C: 'VVIP OC', D: 'PASSENGER OD' };
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

    /* VVIP OB: Meeting Area */
    z0 = r; zoneHeader('VVIP OB · Meeting Area (6 Seats)', 'VVIP OB');
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
// ===== END LAYOUT =====
