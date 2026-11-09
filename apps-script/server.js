// ---------- ตั้งค่า ----------
const MAP_SHEET = 'ผังที่นั่ง';
const LOG_SHEET = 'ประวัติการแก้ไข';
const ALL_SHEET = 'ทุกที่นั่ง และ Monitor';
const BAD_SHEET = 'รวมข้อขัดข้องคงค้าง';
const GAL_SHEET = 'Galley และห้องน้ำ';
const INSPECTOR_CELL = 'D3';   // ช่องชื่อผู้ตรวจในชีท
const LASTSAVE_BY = 'D2', LASTSAVE_AT = 'M2';
const WRITE_PIN = '';          // รหัสเข้าหน้าเว็บและ SAVE เช่น '6020' (เว้นว่าง = ไม่ต้องใช้รหัส)
const ADMIN_PIN = '';          // รหัสสำหรับปุ่ม "ล้างข้อมูลทั้งหมด" บนเว็บ (เว้นว่าง = ปิดปุ่มนี้)
const TZ = 'Asia/Bangkok';
const INIT_BY = 'ร.อ.ธัญชนิต จินตภากร';   // ชื่อผู้บันทึกค่าเริ่มต้นของ Galley / ห้องน้ำ ตอนเพิ่มเข้าผังครั้งแรก

// ผังที่นั่งอยู่ท้ายไฟล์ (seatPlanLayout / planSheet) อ่านครั้งแรกที่ใช้
let LAYOUT_ = null, FIRST_ROW_ = null;
function layout_() { return LAYOUT_ || (LAYOUT_ = seatPlanLayout()); }
// อุปกรณ์ Galley และห้องน้ำใช้สภาพ 3 ระดับ SEAT_E (ใช้งานได้ / ใช้งานได้บางส่วน / ใช้งานไม่ได้) ในช่อง s เหมือนที่นั่ง
function isEq_(it) { return it.kind === 'eq' || it.kind === 'lav'; }
function stList_(it) { return it.kind === 'mon' ? SEAT_M : isEq_(it) ? SEAT_E : SEAT_S; }
function stTh_(it) { return it.kind === 'mon' ? 'สภาพ Monitor' : isEq_(it) ? (it.kind === 'lav' ? 'สภาพห้องน้ำ' : 'สภาพอุปกรณ์') : 'สภาพที่นั่ง'; }
function firstMapRow_() { return FIRST_ROW_ || (FIRST_ROW_ = planSheet(layout_()).firstMapRow); }

// ---------- Web API ----------
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'load';
  try {
    ensureSetup_();
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
    ensureSetup_();
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

/** หาตำแหน่งช่องของทุกที่นั่ง/จอ จากชื่อในช่องหัว (ช่องสถานะอยู่ใต้ชื่อ 1 แถว, NOTE อยู่ใต้ 2 แถว)
 *  เพิ่ม/ลบแถวในชีทได้ ตราบใดที่ยังเรียง ชื่อ / สถานะ / NOTE ไว้เหมือนเดิม */
function buildMap_(sh) {
  const last = Math.max(sh.getLastRow(), 1);
  const width = Math.max(sh.getLastColumn(), 15);
  const v = sh.getRange(1, 1, last + 2, width).getDisplayValues();
  const cell = function (r, c) { return (r >= 1 && r <= v.length && c >= 1 && c <= width) ? String(v[r - 1][c - 1]).trim() : ''; };
  // ชื่อเดียวกัน (เช่น จอ Monitor 21.5" หลายจอ) จับคู่ตามลำดับในผัง: ซ้ายไปขวา บนลงล่าง
  const byLabel = {};
  layout_().order.forEach(function (id) {
    const it = layout_().items[id];
    [it.label].concat(it.alias || []).forEach(function (t) { if (t) (byLabel[t] = byLabel[t] || []).push(id); });   // alias = ข้อความเดียวหรือหลายข้อความ
  });
  const items = {};
  for (let r = firstMapRow_(); r <= last; r++) {
    if (cell(r, 2) === 'NOTE') continue;
    for (let c = 3; c <= width; c++) {
      const cands = byLabel[cell(r, c)];
      if (!cands) continue;
      const id = cands.filter(function (x) { return !items[x]; })[0];
      if (!id) continue;
      const it = layout_().items[id];
      const st = cell(r + 1, c), list = stList_(it);
      if (st !== '' && list.indexOf(st) < 0) continue;
      items[id] = { id: id, kind: it.kind, label: it.name, status: [r + 1, c], note: [r + 2, c] };
    }
  }
  return { items: items, cell: cell };
}

function loadData_() {
  const ss = SpreadsheetApp.getActive();
  const m = buildMap_(ss.getSheetByName(MAP_SHEET));
  const byName = {};
  Object.keys(m.items).forEach(function (id) { byName[m.items[id].label] = id; });
  // ชื่อจอแบบเดิมในแท็บประวัติ (ก่อนเพิ่ม Galley / ห้องน้ำ) ให้ยังหาผู้บันทึกล่าสุดได้
  Object.keys(SEAT_HIST_NAMES).forEach(function (id) { if (m.items[id]) byName[SEAT_HIST_NAMES[id]] = id; });

  // ผู้แก้ไข/เวลาล่าสุดของแต่ละที่นั่ง จากแท็บประวัติ
  const lastBy = {};
  const log = ss.getSheetByName(LOG_SHEET);
  const n = log.getLastRow();
  if (n > 1) {
    log.getRange(2, 1, n - 1, 4).getValues().forEach(function (row) {
      const id = byName[String(row[3])];
      if (!id) return;
      const at = row[0] instanceof Date ? row[0].toISOString() : '';
      const who = (row[2] && row[2] !== '-') ? String(row[2]) : String(row[1] || '');
      lastBy[id] = { by: who, at: at };
    });
  }

  const seats = {};
  Object.keys(m.items).forEach(function (id) {
    const it = m.items[id];
    const st = m.cell(it.status[0], it.status[1]);
    const s = it.kind !== 'mon' ? Math.max(0, stList_(it).indexOf(st)) : 0;
    const mo = it.kind === 'mon' ? Math.max(0, SEAT_M.indexOf(st)) : 0;
    const lb = lastBy[id] || {};
    seats[id] = { s: s, m: mo, note: m.cell(it.note[0], it.note[1]), by: lb.by || '', at: lb.at || '' };
  });
  const ls = PropertiesService.getScriptProperties().getProperty('lastSave');
  return { ok: true, seats: seats, lastSave: ls ? JSON.parse(ls) : null, pinRequired: !!WRITE_PIN, resetEnabled: !!ADMIN_PIN };
}

/** ทุกบรรทัดในแท็บประวัติ (สำหรับ EXPORT TO EXCEL) วันเวลาส่งเป็น ISO */
function historyData_() {
  const log = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
  const n = log.getLastRow();
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
 * - รายการที่ไม่มี fields (หน้าเว็บรุ่นเก่า) จะเขียนทุกช่องแบบเดิม
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
    const sh = ss.getSheetByName(MAP_SHEET);
    const m = buildMap_(sh);
    const now = new Date();
    const logRows = [];
    const conflicts = [];
    const current = {};
    let saved = 0;
    let logCache = null;

    // ผู้แก้ล่าสุดของช่องนี้ จากแท็บประวัติ (อ่านเฉพาะเมื่อมีการชน)
    const lastEdit = function (label, fieldTh, id) {
      const old = SEAT_HIST_NAMES[id] || null;
      if (!logCache) {
        const log = ss.getSheetByName(LOG_SHEET);
        const n = log.getLastRow();
        logCache = n > 1 ? log.getRange(2, 1, n - 1, 5).getValues() : [];
      }
      for (let i = logCache.length - 1; i >= 0; i--) {
        const r = logCache[i];
        if ((String(r[3]) === label || (old && String(r[3]) === old)) && String(r[4]) === fieldTh) {
          return { by: (r[2] && r[2] !== '-') ? String(r[2]) : String(r[1] || ''), at: r[0] instanceof Date ? r[0].toISOString() : '' };
        }
      }
      return { by: '', at: '' };
    };

    changes.forEach(function (ch) {
      const it = m.items[String(ch.id)];
      if (!it) return;
      const isSeat = it.kind !== 'mon', SL = stList_(it);   // ที่นั่ง, อุปกรณ์ Galley และห้องน้ำ ใช้ช่อง s
      const fields = Array.isArray(ch.fields) ? ch.fields : ['s', 'm', 'note'];
      const base = ch.base && typeof ch.base === 'object' ? ch.base : null;
      const st = m.cell(it.status[0], it.status[1]);
      const cur = {
        s: isSeat ? Math.max(0, SL.indexOf(st)) : 0,
        m: isSeat ? 0 : Math.max(0, SEAT_M.indexOf(st)),
        note: m.cell(it.note[0], it.note[1])
      };
      const spec = {
        s: isSeat ? { pos: it.status, th: stTh_(it), val: clamp_(ch.s), txt: function (x) { return SL[x]; } } : null,
        m: isSeat ? null : { pos: it.status, th: 'สภาพ Monitor', val: clamp_(ch.m), txt: function (x) { return SEAT_M[x]; } },
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
            const le = lastEdit(it.label, sp.th, it.id);
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
      markSave_(sh, by, now);
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
    const sh = ss.getSheetByName(MAP_SHEET);
    const m = buildMap_(sh);
    Object.keys(m.items).forEach(function (id) {
      const it = m.items[id];
      sh.getRange(it.status[0], it.status[1]).setValue(stList_(it)[0]);
      sh.getRange(it.note[0], it.note[1]).setValue('');
    });
    const now = new Date();
    appendLog_(ss, [[now, 'หน้าเว็บ', by, 'ทั้งลำ', 'ล้างข้อมูลเริ่มรอบตรวจใหม่', '', 'ปกติทั้งหมด']]);
    markSave_(sh, by, now);
    const lastSave = { by: by, at: now.toISOString(), n: 0 };
    PropertiesService.getScriptProperties().setProperty('lastSave', JSON.stringify(lastSave));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** แสดงผู้ SAVE ล่าสุดในหัวชีท เช่น 14 OCT26 (10:52 TH-LT) */
function markSave_(sh, by, now) {
  sh.getRange(LASTSAVE_BY).setValue(by);
  sh.getRange(LASTSAVE_AT).setValue(Utilities.formatDate(now, TZ, "dd MMMyy '('HH:mm 'TH-LT)'").toUpperCase());
}

function clamp_(x) { x = Number(x) | 0; return x < 0 ? 0 : (x > 2 ? 2 : x); }

function appendLog_(ss, rows) {
  if (!rows.length) return;
  const log = ss.getSheetByName(LOG_SHEET);
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
  if (r0 + nr - 1 < firstMapRow_() && !(r0 <= 3 && r0 + nr - 1 >= 3)) return;

  const m = buildMap_(sh);
  const at = {};
  Object.keys(m.items).forEach(function (id) {
    const it = m.items[id];
    at[it.status.join(',')] = { id: it.label, field: stTh_(it) };
    at[it.note.join(',')] = { id: it.label, field: 'หมายเหตุ' };
  });
  const values = e.range.getDisplayValues();
  const single = nr === 1 && nc === 1;
  const inspector = sh.getRange(INSPECTOR_CELL).getDisplayValue();
  let email = '';
  try { email = (e.user && e.user.getEmail()) || Session.getActiveUser().getEmail(); } catch (err) {}
  const now = new Date();

  const rows = [];
  for (let i = 0; i < nr; i++) {
    for (let j = 0; j < nc; j++) {
      const r = r0 + i, c = c0 + j;
      let info;
      if (r === 3) info = { id: 'ส่วนหัว', field: c <= 9 ? 'ผู้ตรวจ' : 'วันที่ตรวจ' };
      else if (r < firstMapRow_()) continue;
      else info = at[r + ',' + c] || { id: sh.getRange(r, c).getA1Notation(), field: 'อื่น ๆ' };
      const oldV = single ? (e.oldValue === undefined ? '' : e.oldValue) : '(แก้หลายช่องพร้อมกัน)';
      const newV = single ? (e.value === undefined ? values[0][0] : e.value) : values[i][j];
      rows.push([now, email || '-', inspector || '-', info.id, info.field, oldV, newV]);
    }
  }
  appendLog_(e.source, rows);
}

// ---------- สร้างชีท (รันครั้งแรกครั้งเดียว หรือรันซ้ำเพื่อจัดผังใหม่ ข้อมูลเดิมยังอยู่) ----------
function ensureSetup_() {
  if (!SpreadsheetApp.getActive().getSheetByName(MAP_SHEET)) setupSheet();
}

function setupSheet() {
  const ss = SpreadsheetApp.getActive();
  const P = planSheet(layout_());

  // เก็บค่าปัจจุบันไว้ก่อน ถ้ามีผังเดิมอยู่แล้ว
  const old = ss.getSheetByName(MAP_SHEET);
  let keep = null;
  if (old) {
    const m = buildMap_(old);
    keep = { items: {}, head: ['D2', 'M2', 'D3', 'M3'].map(function (a) { return [a, old.getRange(a).getValue()]; }) };
    Object.keys(m.items).forEach(function (id) {
      const it = m.items[id];
      keep.items[id] = { st: m.cell(it.status[0], it.status[1]), note: m.cell(it.note[0], it.note[1]) };
    });
  }

  const tmpName = MAP_SHEET + ' (ใหม่)';
  const leftover = ss.getSheetByName(tmpName);   // จากการรันครั้งก่อนที่ค้างกลางทาง
  if (leftover) ss.deleteSheet(leftover);
  const sh = ss.insertSheet(tmpName, 0);
  applyPlan_(sh, P);
  if (keep) {
    keep.head.forEach(function (h) { if (h[1] !== '') sh.getRange(h[0]).setValue(h[1]); });
    Object.keys(keep.items).forEach(function (id) {
      const p = P.pos[id], k = keep.items[id];
      if (!p) return;
      if (k.st) sh.getRange(p.status[0], p.status[1]).setValue(k.st);
      if (k.note) sh.getRange(p.note[0], p.note[1]).setValue(k.note);
    });
  }

  // แท็บประวัติ: ใช้ของเดิมถ้ามี
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    listHeader_(log, ['วัน/เวลา (TH-LT)', 'อีเมลผู้แก้ไข', 'ผู้ตรวจ (ช่อง D3)', 'ที่นั่ง / จุด', 'รายการ', 'ค่าเดิม', 'ค่าใหม่'], [150, 220, 180, 300, 110, 240, 240]);
    log.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }

  // Galley / ห้องน้ำ ที่เพิ่งเพิ่มเข้าผัง (ยังไม่มีในผังเดิม): บันทึกค่าเริ่มต้น "ใช้งานได้" ลงประวัติ ในชื่อ INIT_BY
  // ที่นั่งและ Monitor ไม่ถูกแตะ; รัน setupSheet ซ้ำจะไม่บันทึกซ้ำเพราะรายการมีอยู่ในผังแล้ว
  const L = layout_(), now = new Date(), initRows = [];
  L.order.forEach(function (id) {
    const it = L.items[id];
    if (!isEq_(it) || (keep && keep.items[id])) return;
    initRows.push([now, 'ค่าเริ่มต้น', INIT_BY, it.name, stTh_(it), '', SEAT_E[0]]);
  });
  appendLog_(ss, initRows);

  // แท็บรายการอ้างถึงแท็บผังเดิม: ล้างสูตรก่อนลบผังเดิม แล้วสร้างใหม่ในแท็บเดิม (ไม่ลบแท็บ)
  [ALL_SHEET, BAD_SHEET, GAL_SHEET].forEach(function (n) { const s = ss.getSheetByName(n); if (s) s.clear(); });
  if (old) ss.deleteSheet(old);
  sh.setName(MAP_SHEET);
  SpreadsheetApp.flush();
  buildListSheets_(ss, P);

  // ลบแท็บว่าง Sheet1 / แผ่นงาน1 ที่มากับไฟล์ใหม่
  ss.getSheets().forEach(function (s) {
    const n = s.getName();
    if ([MAP_SHEET, LOG_SHEET, ALL_SHEET, BAD_SHEET, GAL_SHEET].indexOf(n) < 0 && s.getLastRow() === 0 && /^(Sheet|แผ่นงาน|ชีต)\s*\d+$/.test(n)) ss.deleteSheet(s);
  });
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
}

function hexOk_(x) { return x || null; }

function applyPlan_(sh, P) {
  if (sh.getMaxRows() < P.lastRow + 5) sh.insertRowsAfter(sh.getMaxRows(), P.lastRow + 5 - sh.getMaxRows());
  if (sh.getMaxColumns() > P.lastCol + 1) sh.deleteColumns(P.lastCol + 2, sh.getMaxColumns() - P.lastCol - 1);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily('Tahoma').setFontSize(9).setVerticalAlignment('middle');
  Object.keys(P.widths).forEach(function (c) { sh.setColumnWidth(Number(c), P.widths[c]); });
  Object.keys(P.heights).forEach(function (r) { sh.setRowHeight(Number(r), P.heights[r]); });

  const dvS = SpreadsheetApp.newDataValidation().requireValueInList(SEAT_S, true).setAllowInvalid(false).build();
  const dvM = SpreadsheetApp.newDataValidation().requireValueInList(SEAT_M, true).setAllowInvalid(false).build();
  const dvE = SpreadsheetApp.newDataValidation().requireValueInList(SEAT_E, true).setAllowInvalid(false).build();
  const SOLID = SpreadsheetApp.BorderStyle.SOLID;
  P.ops.forEach(function (o) {
    const rg = sh.getRange(o.r, o.c, o.rs, o.cs);
    if (o.cabin) { rg.setBackground(o.fill); return; }
    if (o.rs > 1 || o.cs > 1) rg.merge();
    if (o.v !== undefined && o.v !== null) {
      const top = sh.getRange(o.r, o.c);
      if (String(o.v).charAt(0) === '=') top.setFormula(o.v); else top.setValue(o.v);
    }
    if (o.fill) rg.setBackground(o.fill);
    const f = o.font || {};
    if (f.b) rg.setFontWeight('bold');
    if (f.i) rg.setFontStyle('italic');
    if (f.sz) rg.setFontSize(f.sz);
    if (f.color) rg.setFontColor(f.color);
    const a = o.align || {};
    if (a.h) rg.setHorizontalAlignment(a.h);
    if (a.v) rg.setVerticalAlignment(a.v);
    if (a.wrap) rg.setWrap(true);
    if (o.border) rg.setBorder(true, true, true, true, null, null, '#000000', SOLID);
    if (o.dv) rg.setDataValidation(o.dv === 'M' ? dvM : o.dv === 'E' ? dvE : dvS);
    if (o.fmt) rg.setNumberFormat(o.fmt);
  });
  P.boxes.forEach(function (b) {
    sh.getRange(b.r1, b.c1, b.r2 - b.r1 + 1, b.c2 - b.c1 + 1).setBorder(true, true, true, true, null, null, SEAT_COLOR.MUTED, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  const rules = P.cf.map(function (c) {
    let b = SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange(c.range)]).setBackground(c.fill);
    if (c.font) b = b.setFontColor(c.font);
    b = c.text !== undefined ? b.whenTextEqualTo(c.text) : b.whenFormulaSatisfied(c.f);
    return b.build();
  });
  sh.setConditionalFormatRules(rules);
  sh.setFrozenRows(0);
}

/** ใช้แท็บเดิมถ้ามี (ล้างทุกอย่างก่อน) ไม่มีก็สร้างใหม่ */
function listSheet_(ss, name, index) {
  let s = ss.getSheetByName(name);
  if (s) {
    s.clear();
    s.clearConditionalFormatRules();
    s.setFrozenRows(0);
    const f = s.getFilter(); if (f) f.remove();
  } else {
    s = ss.insertSheet(name, index);
  }
  return s;
}

function listHeader_(s, titles, widths) {
  s.getRange(1, 1, 1, titles.length).setValues([titles]).setFontWeight('bold').setBackground(SEAT_COLOR.ZONEH)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  s.getRange(1, 1, s.getMaxRows(), titles.length).setFontFamily('Tahoma').setFontSize(9);
  widths.forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  s.setRowHeight(1, 30);
  s.setFrozenRows(1);
}

function buildListSheets_(ss, P) {
  const q = "'" + MAP_SHEET + "'!";
  const zn = {}; SEAT_ZONES.forEach(function (z) { zn[z.key] = z.name; });
  const A1 = function (p) { return seatColName(p[1]) + p[0]; };
  const hdr = ['ที่นั่ง / จุด', 'ประเภท', 'โซน', 'สภาพที่นั่ง', 'สภาพ Monitor', 'หมายเหตุ', 'ผู้บันทึกล่าสุด', 'วัน/เวลาล่าสุด (TH-LT)'];
  const wid = [330, 80, 140, 110, 100, 300, 180, 140];

  const all = listSheet_(ss, ALL_SHEET, 1);
  listHeader_(all, hdr, wid);
  const hist = function (col, row) {
    return "=IFERROR(INDEX(FILTER('" + LOG_SHEET + "'!$" + col + "$2:$" + col + ",'" + LOG_SHEET + "'!$D$2:$D=$A" + row + "),COUNTIF('" + LOG_SHEET + "'!$D$2:$D,$A" + row + ")),\"\")";
  };
  const seatIds = layout_().order.filter(function (id) { return !isEq_(layout_().items[id]); });
  const eqIds = layout_().order.filter(function (id) { return isEq_(layout_().items[id]); });
  const rows = seatIds.map(function (id, i) {
    const it = layout_().items[id], p = P.pos[id], row = i + 2;
    return [it.name, it.kind === 'mon' ? 'Monitor' : 'ที่นั่ง', zn[it.zone],
      it.kind === 'seat' ? '=' + q + A1(p.status) : '-',
      it.kind === 'mon' ? '=' + q + A1(p.status) : '-',
      '=' + q + A1(p.note) + '&""', hist('C', row), hist('A', row)];
  });
  all.getRange(2, 1, rows.length, 8).setValues(rows)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('top');
  all.getRange(2, 2, rows.length, 4).setHorizontalAlignment('center');
  all.getRange(2, 6, rows.length, 1).setWrap(true);
  all.getRange(2, 8, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm').setHorizontalAlignment('center');
  statusColors_(all, rows.length + 1);

  // Galley และห้องน้ำ: รายการ 8 คอลัมน์แบบเดียวกัน (สภาพอยู่คอลัมน์ D) + ตารางสรุปทางขวา
  const gal = listSheet_(ss, GAL_SHEET, 1);
  listHeader_(gal, ['รายการ', 'ประเภท', 'GALLEY / โซน', 'สภาพ', '-', 'หมายเหตุ', 'ผู้บันทึกล่าสุด', 'วัน/เวลาล่าสุด (TH-LT)'], [300, 80, 150, 120, 30, 300, 180, 140]);
  const grows = eqIds.map(function (id, i) {
    const it = layout_().items[id], p = P.pos[id], row = i + 2;
    return [it.name, it.kind === 'lav' ? 'ห้องน้ำ' : it.tname, it.kind === 'lav' ? zn[it.zone] : it.group,
      '=' + q + A1(p.status), '-', '=' + q + A1(p.note) + '&""', hist('C', row), hist('A', row)];
  });
  if (grows.length) {
    gal.getRange(2, 1, grows.length, 8).setValues(grows)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('top');
    gal.getRange(2, 2, grows.length, 4).setHorizontalAlignment('center');
    gal.getRange(2, 6, grows.length, 1).setWrap(true);
    gal.getRange(2, 8, grows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm').setHorizontalAlignment('center');
    const last = grows.length + 1, C = SEAT_COLOR;
    const sumRows = [['GALLEY / ห้องน้ำ', 'จำนวน'].concat(SEAT_E)];
    Object.keys(SEAT_GAL).forEach(function (k) { sumRows.push([SEAT_GAL[k], '=COUNTIF($C$2:$C$' + last + ',J' + (sumRows.length + 1) + ')'].concat(SEAT_E.map(function (t) { return '=COUNTIFS($C$2:$C$' + last + ',J' + (sumRows.length + 1) + ',$D$2:$D$' + last + ',"' + t + '")'; }))); });
    sumRows.push(['ห้องน้ำ (LAVATORY)', '=COUNTIF($B$2:$B$' + last + ',"ห้องน้ำ")'].concat(SEAT_E.map(function (t) { return '=COUNTIFS($B$2:$B$' + last + ',"ห้องน้ำ",$D$2:$D$' + last + ',"' + t + '")'; })));
    const sr = gal.getRange(1, 10, sumRows.length, 5);
    sr.setValues(sumRows).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Tahoma').setFontSize(9).setHorizontalAlignment('center');
    gal.getRange(1, 10, 1, 5).setFontWeight('bold').setBackgrounds([[C.ZONEH, C.ZONEH, C.GREEN, C.YEL, C.RED]]).setWrap(true);
    gal.getRange(2, 10, sumRows.length - 1, 1).setHorizontalAlignment('left').setFontWeight('bold').setBackground(C.CABIN);
    gal.setColumnWidth(9, 20); gal.setColumnWidth(10, 170); [11, 12, 13, 14].forEach(function (c) { gal.setColumnWidth(c, 90); });
    statusColors_(gal, last);
  }

  const bad = listSheet_(ss, BAD_SHEET, 1);
  listHeader_(bad, hdr, wid);
  const s = "'" + ALL_SHEET + "'!", g = "'" + GAL_SHEET + "'!";
  const both = function (col) { return '{' + s + col + '2:' + col + ';' + g + col + '2:' + col + '}'; };
  bad.getRange('A2').setFormula('=IFERROR(FILTER({' + s + 'A2:H;' + g + 'A2:H},(' + both('D') + '="' + SEAT_S[1] + '")+(' + both('D') + '="' + SEAT_S[2] + '")+(' +
    both('D') + '="' + SEAT_E[1] + '")+(' + both('D') + '="' + SEAT_E[2] + '")+(' +
    both('E') + '="' + SEAT_M[1] + '")+(' + both('E') + '="' + SEAT_M[2] + '")+(' + both('F') + '<>"")),"ไม่มีข้อขัดข้องคงค้าง")');
  bad.getRange('H2:H').setNumberFormat('dd/MM/yyyy HH:mm');
  bad.getRange('F2:F').setWrap(true);
  statusColors_(bad, 200);
}

function statusColors_(s, end) {
  const rg = [s.getRange('D2:E' + end)];
  const C = SEAT_COLOR;
  s.setConditionalFormatRules([
    [SEAT_S[2], C.RED, '#FFFFFF'], [SEAT_M[2], C.RED, '#FFFFFF'], [SEAT_E[2], C.RED, '#FFFFFF'], [SEAT_S[1], C.YEL, null], [SEAT_M[1], C.YEL, null], [SEAT_E[1], C.YEL, null]
  ].map(function (x) {
    let b = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(x[0]).setBackground(x[1]).setRanges(rg);
    if (x[2]) b = b.setFontColor(x[2]);
    return b.build();
  }));
}
