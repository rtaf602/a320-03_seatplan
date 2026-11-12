# STATUS SEAT PLAN A320-03 — คู่มือสำหรับผู้แก้โค้ด

อ่านไฟล์นี้ก่อนแก้อะไรใน repo นี้ **ตอบผู้ใช้เป็นภาษาไทย** ผู้ใช้เป็นเจ้าหน้าที่ของ ฝูง.602 ไม่ใช่โปรแกรมเมอร์

## เว็บนี้ใช้ระบบ Login กลาง (ตั้งแต่ 12 พ.ย. 2026)

ใครเข้าได้และทำอะไรได้ **ไม่ได้ตัดสินใน repo นี้** แต่ตัดสินโดยระบบ Login กลางของ ฝูง.602 (repo `rtaf602/eadmin` หน้าเว็บอยู่ที่ `/eadmin/` ของ GitHub Pages เดียวกัน)
ไม่มี `WRITE_PIN` / `ADMIN_PIN` และช่องพิมพ์ชื่อผู้ตรวจแล้ว **ห้ามเพิ่มกลับ**

| Role ในเว็บนี้ (Admin หลักกำหนดในหน้า Admin รหัสเว็บ `WEB4`) | ทำได้ |
|---|---|
| Admin | ทุกอย่างของ Editor และ "ล้างข้อมูลทั้งหมดเพื่อเริ่มรอบตรวจใหม่" (แทน `ADMIN_PIN` เดิม) |
| Editor | ดู แก้ไข SAVE CLEAR PREVIEW และส่งออก PDF / Excel |
| Viewer | ดู และส่งออก PDF ทั้ง 2 แบบ / Excel ได้ (ผู้ใช้สั่ง 12 พ.ย. 2026) ไม่มี SAVE CLEAR PREVIEW |
| ไม่มีสิทธิ์ / เว็บถูกปิด | เข้าไม่ได้ เห็นข้อความ "บัญชีนี้ไม่มีสิทธิ์ใช้งานเว็บนี้ ให้ติดต่อ Admin" |

## ส่วนที่เป็นของระบบ Login (ห้ามแก้ ห้ามถอด)

หน้าเว็บ `index.html`
- `<script src="/eadmin/assets/config.js">` และ `<script src="/eadmin/assets/app.js">` (`window.App`)
- `AUTH`, `SESS()`, `getPass()`, `PASSV()`, `apiPost()`: ขอ "บัตรผ่าน" ของเว็บนี้จาก `App.enterWeb(CFG.webCode)` แล้วส่ง `pass` ไปกับ**ทุก**คำขอ **ห้ามส่ง session ของ Login ไป Apps Script ของเว็บนี้**
- `denied()`, `viewer()` (ซ่อนเฉพาะ SAVE CLEAR PREVIEW), `mayExport()` (ทุก Role ส่งออกได้), `logExport()`, `#gate`, `#gateBtn`, ปุ่ม `#bOut` (ออกจากระบบทุกเว็บ)
- `start()`: ไม่มีการ Login → `App.toLogin()` มีแล้ว → `App.watch()` (หมดเวลา 1 ชั่วโมง, Login ซ้อน, ออกจากระบบในแท็บอื่น) แล้วขอบัตรผ่าน
- ช่อง `#inspector` อ่านอย่างเดียว ค่ามาจาก `load` → `me.name` (คอลัมน์ ชื่อ-สกุล ของระบบ Login)

Apps Script `apps-script/server.js` (`Code.gs` สร้างจาก server.js + layout.js ด้วย `sh tools/build-code-gs.sh` ห้ามแก้ Code.gs ตรง ๆ)
- `LOGIN = { LOGIN_URL, WEB_CODE, CHECK_SEC }` (`LOGIN_URL` จำใน Script Property ตอน `setup`; `WEB_CODE` ต้องตรงกับ `webCode` ใน `config.js`)
- `doPost`: **ทุก** action ผ่าน `auth_(body.pass, FRESH[act])` แล้วตรวจ `EDIT` (Admin / Editor: `save`, `reset`; ส่วน `history` และ `exported` ทุก Role ของเว็บนี้ใช้ได้ เพื่อให้ Viewer ส่งออกได้) และ `reset` ตรวจ `who.admin`
- `doGet` ไม่ส่งข้อมูลแล้ว
- ชื่อผู้ตรวจในประวัติและ `lastSave` = `who.name` จากระบบ Login **ไม่ใช่ `body.by`**
- `exported`: EXPORT TO PDF ALL PAGES / EXPORT PDF SUMMARY 4 PAGES / EXPORT TO EXCEL ลงแท็บ **ประวัติการส่งออก** (`EXPORT_SHEET` แยกจากประวัติการแก้ไข)
- `onEdit` (แก้ในชีทโดยตรง) ยังทำงานแบบเดิม ใช้ชื่อในช่อง D3 ของชีท
- `setup()` = ตั้งค่า Login (จำ `LOGIN_URL`) ส่วน `setupSheet()` = สร้าง/จัดแท็บผังที่นั่ง (ของเดิม) เป็นคนละฟังก์ชัน

## กติกาเมื่อเพิ่มความสามารถใหม่

1. action ใหม่ต้องอยู่ในรายการที่ `doPost` รับ และผ่าน `auth_` ถ้าเขียนข้อมูล สร้างไฟล์ หรือส่งข้อมูลออก ใส่ใน `EDIT` และ `FRESH` ด้วย
2. การห้ามของ Viewer ต้องอยู่ที่ Apps Script เสมอ การซ่อนปุ่มเป็นแค่ความสะดวก
3. ห้ามเพิ่มช่องให้พิมพ์ชื่อผู้ใช้หรือรหัสผ่าน ห้ามเก็บ session หรือ pass ลง Sheet หรือ Script Property
4. การเปลี่ยนวิธี Login หรือ Role ต้องทำที่ repo `eadmin`
5. repo นี้เป็นสาธารณะ ห้ามใส่รหัสผ่านหรือข้อมูลลับลงในไฟล์

## ทดสอบ

ชุดทดสอบร่วมอยู่ใน repo eadmin (`test/web_a320_03.mjs` ต้อง clone repo นี้ไว้ข้าง ๆ eadmin) ผ่านทั้งหมดเมื่อ 12 พ.ย. 2026 บนตัวจำลอง ยังไม่เคยรันบน Google จริง
