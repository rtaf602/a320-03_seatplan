/**
 * Seat Plan A320-03 (For Admin) — Apps Script
 *
 * 1) สร้างแท็บในชีท: เลือกฟังก์ชัน setupSheet แล้วกด Run (ครั้งแรกครั้งเดียว)
 * 2) เป็น API ให้หน้าเว็บ (GitHub Pages) อ่าน/บันทึกสภาพที่นั่ง, Monitor, Galley และห้องน้ำ ลงแท็บ "ผังที่นั่ง"
 *    และบันทึกประวัติลงแท็บ "ประวัติการแก้ไข" ทุกครั้งที่กด SAVE บนเว็บ
 * 3) บันทึกประวัติอัตโนมัติเมื่อมีคนแก้ไขช่องในชีทโดยตรง (onEdit)
 *
 * ผังที่นั่ง (ตำแหน่งช่องของแต่ละที่นั่ง/จอ) อยู่ท้ายไฟล์ คัดลอกมาจาก layout.js
 * ไฟล์นี้สร้างจาก apps-script/server.js + layout.js ด้วย tools/build-code-gs.sh วิธีติดตั้งดู README.md
 */

// ---------- ตั้งค่า ----------
const MAP_SHEET = 'ผังที่นั่ง';
const LOG_SHEET = 'ประวัติการแก้ไข';
const EXPORT_SHEET = 'ประวัติการส่งออก';   // EXPORT TO PDF / EXCEL จากหน้าเว็บ (สร้างเองครั้งแรกที่มีการส่งออก)
const EXPORT_HEAD = ['เวลา', 'ชื่อ-สกุล', 'Username', 'Role', 'ชนิดไฟล์', 'ชื่อไฟล์'];
const ALL_SHEET = 'ทุกที่นั่ง และ Monitor';
const BAD_SHEET = 'รวมข้อขัดข้องคงค้าง';
const GAL_SHEET = 'Galley และห้องน้ำ';
const INSPECTOR_CELL = 'D3';   // ช่องชื่อผู้ตรวจในชีท
const LASTSAVE_BY = 'D2', LASTSAVE_AT = 'M2';

// ---------- Login กลาง ----------
// ใครเข้าได้และทำอะไรได้ ตัดสินโดยระบบ Login กลางของ ฝูง.602 (repo eadmin) ทุกคำขอ ไม่มีรหัส WRITE_PIN / ADMIN_PIN แล้ว
// LOGIN_URL: ที่อยู่ระบบหลังบ้านของระบบ Login (ค่าเดียวกับ backendUrl ในไฟล์ eadmin/docs/assets/config.js)
//            กรอกในหน้า Apps Script แล้ว Run setup ครั้งเดียว ระบบจำไว้เอง (ช่องนี้ว่างไว้ได้ในโค้ดบน GitHub)
// WEB_CODE:  รหัสเว็บนี้ในรายชื่อเว็บของหน้า Admin (ต้องตรงกับ webCode ในไฟล์ config.js)
const LOGIN = { LOGIN_URL: '', WEB_CODE: 'WEB4', CHECK_SEC: 60 };
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
// Every request needs a login (Login กลาง): nothing is read or written without one
function doGet(e) {
  return json_({ ok: true, app: 'a320-03_seatplan' });
}

// Editor and Admin only: these write the sheet, or give a file out (EXPORT TO EXCEL reads the whole history)
const EDIT = { save: 1, reset: 1, history: 1, exported: 1 };
// these ask the login system afresh, not from the copy kept for LOGIN.CHECK_SEC
const FRESH = { save: 1, reset: 1, history: 1, exported: 1 };
const EXPORTS = ['EXPORT TO PDF ALL PAGES', 'EXPORT PDF SUMMARY 4 PAGES', 'EXPORT TO EXCEL'];

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    const act = String(body.action || '');
    if (['load', 'history', 'save', 'reset', 'exported'].indexOf(act) < 0) return json_({ ok: false, error: 'unknown_action' });
    const who = auth_(body.pass, !!FRESH[act]);
    ensureSetup_();
    if (EDIT[act] && !who.edit) return json_({ ok: false, error: 'forbidden', message: 'บัญชีนี้มีสิทธิ์ดูอย่างเดียว (Viewer)' });
    if (act === 'load') return json_(Object.assign(loadData_(who), { me: { name: who.name, role: who.role } }));
    if (act === 'history') return json_(historyData_());
    if (act === 'save') return json_(saveData_(body, who));
    if (act === 'reset') return json_(resetData_(body, who));
    if (act === 'exported') return json_(exported_(body, who));
  } catch (err) {
    if (err && err.code) return json_({ ok: false, error: err.code, message: err.message });
    return json_({ ok: false, error: 'server', message: String(err) });
  }
}

/** EXPORT TO PDF / EXCEL on the page: one line in the tab ประวัติการส่งออก, with the account that did it */
function exported_(body, who) {
  const kind = String(body.kind || '');
  if (EXPORTS.indexOf(kind) < 0) return { ok: false, error: 'bad_request' };
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let sh = ss.getSheetByName(EXPORT_SHEET);
    if (!sh) {
      const log = ss.getSheetByName(LOG_SHEET);
      sh = ss.insertSheet(EXPORT_SHEET, log ? log.getIndex() : ss.getNumSheets());
      sh.getRange(1, 1, 1, EXPORT_HEAD.length).setValues([EXPORT_HEAD]).setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      sh.setColumnWidths(1, 1, 150); sh.setColumnWidths(2, 1, 200); sh.setColumnWidths(5, 1, 220); sh.setColumnWidths(6, 1, 380);
    }
    const r = sh.getLastRow() + 1;
    // text cells: a name or a file name is never read as a number or a formula
    sh.getRange(r, 2, 1, 5).setNumberFormat('@');
    sh.getRange(r, 1, 1, 6).setValues([[new Date(), who.name, who.username || '', who.role, kind, String(body.filename || '').slice(0, 120)]]);
    sh.getRange(r, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/* ---------- Login กลาง (ระบบ Login ฝูง.602) ----------
   The page gets a pass for this web from the login system (App.enterWeb) and sends it with every request; this script
   sends it on to the login system (op web.check) with the code of this web, and gets back the person and the Role in
   this web: Admin / Editor write, Viewer looks only, Admin also clears all data. No access, the web closed, the hour
   over, the password changed, the account taken by another device: refused, every time. The pass is good for that
   question about this web only. An answer is kept for LOGIN.CHECK_SEC for reading only; every write asks afresh.
   Only a mark of the pass is kept here, never the pass itself. */
function err_(code, message) { const e = new Error(message); e.code = code; return e; }
function sha256_(t) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
function loginUrl_() { return String(PropertiesService.getScriptProperties().getProperty('LOGIN_URL') || LOGIN.LOGIN_URL || '').trim(); }
// the address of a deployed Apps Script web app, nothing else: the pass is only ever sent there
function loginUrlOk_(url) { return /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url); }
function auth_(pass, fresh) {
  pass = String(pass || '');
  if (!pass || pass.length > 800) throw err_('auth', 'ให้เข้าสู่ระบบ');
  let w = '';
  try { w = String(JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(pass.split('.')[0])).getDataAsString()).w || ''); } catch (x) { w = ''; }
  if (w && w !== LOGIN.WEB_CODE) throw err_('wrong_web', 'รหัสเว็บในไฟล์ config.js (' + w + ') ไม่ตรงกับ LOGIN.WEB_CODE ใน Apps Script (' + LOGIN.WEB_CODE + ') ให้แก้ให้ตรงกัน');
  const key = sha256_('a32003|' + pass).slice(0, 32), cache = CacheService.getScriptCache();
  if (!fresh) { const hit = cache.get('who:' + key); if (hit) return JSON.parse(hit); }
  const url = loginUrl_();
  if (!loginUrlOk_(url)) throw err_('config', 'ยังไม่ได้ตั้งที่อยู่ระบบ Login (LOGIN.LOGIN_URL แล้ว Run setup)');
  let j = null;
  try {
    const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'text/plain;charset=utf-8', payload: JSON.stringify({ v: 1, op: 'web.check', pass: pass, web: LOGIN.WEB_CODE }), muteHttpExceptions: true, followRedirects: true });
    j = JSON.parse(res.getContentText());
  } catch (x) { j = null; }
  if (!j || (!j.ok && !j.error)) throw err_('busy', 'ตรวจสิทธิ์กับระบบ Login ไม่ได้ ลองอีกครั้ง');
  if (!j.ok) {
    cache.remove('who:' + key);
    const c = String(j.error.code || '');
    if (c === 'auth' || c === 'expired' || c === 'taken') throw err_('auth', String(j.error.message || 'ให้เข้าสู่ระบบใหม่'));
    if (c === 'no_web' || c === 'web_closed') throw err_(c, String(j.error.message || 'บัญชีนี้ไม่มีสิทธิ์ใช้งานเว็บนี้ ให้ติดต่อ Admin'));
    throw err_('busy', String(j.error.message || 'ระบบ Login ขัดข้อง ลองอีกครั้ง'));
  }
  const u = j.user || {}, role = String(j.role || '');
  // the recorder: column "ชื่อ-สกุล" of the account's row in LOGIN DATABASE (Username as fallback)
  const who = { key: key, name: String(u.name || u.username || '').trim().slice(0, 60), username: String(u.username || ''), role: role, edit: role === 'Admin' || role === 'Editor', admin: role === 'Admin' };
  const left = Math.floor((Number(j.expiresAt) - Date.now()) / 1000);
  if (left > 1) cache.put('who:' + key, JSON.stringify(who), Math.min(LOGIN.CHECK_SEC, left));
  return who;
}

/** รันครั้งเดียวหลังวางโค้ด: จำที่อยู่ระบบ Login และตรวจว่าติดต่อได้ (Google จะขอสิทธิ์เชื่อมต่อบริการภายนอก) */
function setup() {
  const p = PropertiesService.getScriptProperties();
  const url = String(LOGIN.LOGIN_URL || '').trim();
  if (url) {
    if (!loginUrlOk_(url)) throw new Error('LOGIN.LOGIN_URL ต้องเป็นที่อยู่ Web app ของระบบ Login (ขึ้นต้น https://script.google.com/macros/s/ ลงท้าย /exec)');
    p.setProperty('LOGIN_URL', url);
  }
  if (!loginUrl_()) throw new Error('ยังไม่ได้กรอก LOGIN.LOGIN_URL (ที่อยู่ระบบหลังบ้านของระบบ Login) กรอกแล้ว Run setup อีกครั้ง');
  let ok = false;
  try { ok = JSON.parse(UrlFetchApp.fetch(loginUrl_(), { muteHttpExceptions: true, followRedirects: true }).getContentText()).service === 'eadmin602'; } catch (x) { ok = false; }
  if (!ok) throw new Error('ติดต่อระบบ Login ที่ ' + loginUrl_() + ' ไม่ได้ ตรวจ LOGIN.LOGIN_URL แล้ว Run setup อีกครั้ง');
  Logger.log('ระบบ Login ตอบแล้ว · รหัสเว็บนี้ ' + LOGIN.WEB_CODE + ' (ต้องตรงกับรายชื่อเว็บในหน้า Admin)');
  Logger.log('ตั้งค่าเสร็จแล้ว');
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

function loadData_(who) {
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
  return { ok: true, seats: seats, lastSave: ls ? JSON.parse(ls) : null, pinRequired: false, resetEnabled: !!(who && who.admin) };
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
function saveData_(body, who) {
  const by = who.name;                      // the account confirmed by the login system, never a name the page sends
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

function resetData_(body, who) {
  if (!who.admin) return { ok: false, error: 'forbidden', message: 'ล้างข้อมูลทั้งหมดได้เฉพาะ Admin ของเว็บนี้' };
  const by = who.name || '-';
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

// ---------- ผังที่นั่ง (จาก layout.js) ----------
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
  door('B', '', '', true, '');

  zone('C', 'VVIP OC (8 Seats)', 'VVIP OC');
  mon('MON-C1', 'C', 'w21', 'หน้าที่นั่ง 7B, 7C', 'หน้า 7B-7C');
  mon('MON-C2', 'C', 'w21', 'หน้าที่นั่ง 7D, 7F', 'หน้า 7D-7F');
  mon('MON-C3', 'C', 'w21', 'หลังที่นั่ง 8B, 8C', 'หลัง 8B-8C');
  mon('MON-C4', 'C', 'w21', 'หลังที่นั่ง 8D, 8F', 'หลัง 8D-8F');
  blk('C', '', [{ id: lav('CV', 'C', 'CENTER VIP LAVATORY'), c: R2S, span: 4 }]);
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
      op({ r: r, c: 1, v: '◀ EMER\nEXIT L', fill: C.EXIT, font: ef, align: { h: 'left', v: 'middle', wrap: true }, border: false });
      op({ r: r, c: 16, v: 'EMER\nEXIT R ▶', fill: C.EXIT, font: ef, align: { h: 'right', v: 'middle', wrap: true }, border: false });
      heights[r] = 34; heights[r + 1] = 6; r += 2;   // 2 short lines, all visible
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
