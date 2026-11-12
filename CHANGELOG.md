# CHANGELOG

## 2026-11-12 · Viewer ส่งออกไฟล์ได้
- Viewer กด EXPORT TO PDF ALL PAGES / EXPORT PDF SUMMARY 4 PAGES / EXPORT TO EXCEL ได้ (บันทึกในแท็บ "ประวัติการส่งออก" เหมือน Role อื่น) ยัง SAVE, CLEAR และ PREVIEW ไม่ได้
- ต้องวาง apps-script/Code.gs ใหม่และ Deploy เวอร์ชันใหม่ (Manage deployments → ดินสอ → New version)

## 2026-11-12 · ใช้ระบบ Login กลาง ฝูง.602
- เอารหัส WRITE_PIN / ADMIN_PIN และช่องพิมพ์ชื่อผู้ตรวจออก เข้าใช้งานผ่านหน้าเข้าสู่ระบบกลาง (รหัสเว็บ WEB4)
- Admin / Editor แก้ไขและ SAVE ได้ · Viewer ดูอย่างเดียว · ล้างข้อมูลทั้งหมดเฉพาะ Admin ของเว็บนี้
- ชื่อผู้ตรวจ = ชื่อ-สกุลของบัญชีที่ Login · การส่งออก PDF / Excel บันทึกในแท็บ "ประวัติการส่งออก"

## 2026-11-09 · ทางออกฉุกเฉินเหลือคู่เดียว
- เหลือ EMER EXIT คู่เดียว (หลังแถว 6) ชื่อ **EMER EXIT L / EMER EXIT R** ตัด EXIT AFT ออก
- แก้ทั้งหน้าเว็บ, PDF ทุกแบบ, EXPORT TO EXCEL และแท็บ "ผังที่นั่ง" ใน Google Sheet (วาง Code.gs ใหม่, รัน setupSheet, Deploy New version)

## 2026-11-09 · เพิ่ม Galley, ห้องน้ำ และประตู (ยกแบบจาก A320-05)

### หน้าเว็บ
- ยกโค้ดหน้าเว็บจาก `a320-05_seatplan` ทั้งชุด ชื่อระบบ **STATUS SEAT PLAN A320-03**
- **ที่นั่ง 73 และ Monitor 23 จอคงเดิมทุกตัว** (รหัสเดิม ตำแหน่งเดิม)
- เพิ่ม **FWD GALLEY (G1)** 18 รายการ และ **AFT GALLEY (G5)** 27 รายการ
- เพิ่มห้องน้ำ 5 ห้องตามผัง Excel 320(03) SEAT: LAV A, FWD VIP LAV, CENTER VIP LAV, LAV D, LAV E
- ประตู: DOOR 1L/1R หน้าห้อง Commander Center · EMER EXIT FWD L/R หลังแถว 6 · EMER EXIT AFT L/R แนว CENTER VIP LAV · DOOR 4L/4R ก่อน AFT GALLEY
- ปุ่ม CLEAR / PREVIEW (ALL PAGES, SUMMARY 4 PAGES) / SAVE / EXPORT TO PDF ALL PAGES / EXPORT SUMMARY 4 PAGES / EXPORT TO EXCEL
- ชื่อไฟล์ EXPORT: `YYMMDD HHMM TH-LT A320-03_STATUS_SEAT_PLAN_...`
- หน้า 1 ของ SUMMARY รองรับโซฟา (กล่อง SOFA กว้าง 2 คอลัมน์ เหนือ 2A 3A)

### Excel
- แผ่นงาน "ผังที่นั่ง" มี Galley, ห้องน้ำ และประตู · เพิ่มแผ่นงาน "Galley และห้องน้ำ"
- หัวตาราง `Monitor / เสื่อม / สภาพ` ขึ้นบรรทัดตามคำ ไม่ให้ "พ" ตกบรรทัด
- ป้าย EMER EXIT 3 บรรทัด แถวสูงพอให้เห็นครบ · ป้ายฝั่งซ้ายอยู่คอลัมน์ A (พื้นขาว) คู่กับฝั่งขวาคอลัมน์ P
- แถบ DOOR ลากเต็มความกว้าง คอลัมน์ A ถึง P

### Apps Script
- เปลี่ยนเป็นโครงแบบ A320-05: `apps-script/server.js` + `layout.js` → `Code.gs` (`tools/build-code-gs.sh` แทน `tools/build-code-gs.js`)
- `alias` รองรับหลายชื่อ เพื่ออ่านชีทรุ่นเดิม (`OVHD 12.0"`, `Commander Center 1A`)
- **ต้องวาง `Code.gs` ใหม่ รัน `setupSheet` และ Deploy New version ก่อนขึ้นหน้าเว็บ** (ค่าที่นั่ง/จอเดิมและประวัติยังอยู่)
- เวลา SAVE ล่าสุดของผังรุ่นเดิมไม่ถูกย้าย จะแสดงใหม่หลัง SAVE ครั้งแรก
- ชื่อจอแบบเดิมในแท็บประวัติ (เช่น "Monitor หน้าที่นั่ง CC 1A") ยังจับคู่ได้ (`SEAT_HIST_NAMES` ใน `layout.js`) หน้าเว็บจึงยังแสดงผู้บันทึกล่าสุดของจอ
- ผู้บันทึกค่าเริ่มต้นของ Galley / ห้องน้ำ: ร.อ.ธัญชนิต จินตภากร
- หน้าเว็บโหลด `config.js`, `layout.js`, `export-excel.js` ใหม่ทุกครั้ง ไม่ติดไฟล์เก่าในเบราว์เซอร์
