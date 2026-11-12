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
