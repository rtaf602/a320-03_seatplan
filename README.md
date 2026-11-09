# STATUS SEAT PLAN A320-03 · ตรวจสภาพที่นั่ง, Monitor, Galley และห้องน้ำ

หน้าเว็บให้ฝ่ายการช่าง/ฝ่ายสื่อสารบันทึกสภาพที่นั่ง, Monitor, อุปกรณ์ Galley และห้องน้ำของเครื่อง A320-03
(ที่นั่ง 73, Monitor 23 จอ, อุปกรณ์ Galley 45 รายการ, ห้องน้ำ 5 ห้อง)
ข้อมูลทุกครั้งที่กด **SAVE** จะบันทึกลง Google Sheet **Seat Plan 320(03) (For Admin)** โดยตรง
และเพิ่มบรรทัดในแท็บ **ประวัติการแก้ไข**

ระบบนี้ยกมาจาก [`rtaf602/a320-05_seatplan`](https://github.com/rtaf602/a320-05_seatplan) (Galley, ห้องน้ำ, ประตู, ปุ่ม และรูปแบบ EXPORT)
โดยคงที่นั่งและ Monitor ของ A320-03 ไว้ตามเดิมทุกตัว

> - วิธีออกแบบ Galley / ห้องน้ำ / ประตู ดู [`docs/GALLEY_LAV_GUIDE.md`](docs/GALLEY_LAV_GUIDE.md)
> - วิธียกระบบไปใช้กับเครื่องลำอื่น (กรณีศึกษา A320-05 → A320-03) ดู [`docs/PORTING_GUIDE.md`](docs/PORTING_GUIDE.md)
> - สิ่งที่เปลี่ยนจากรุ่นก่อน ดู [`CHANGELOG.md`](CHANGELOG.md)

| ไฟล์ | หน้าที่ |
|---|---|
| `index.html` | หน้าเว็บ: ผังที่นั่ง + Galley + ห้องน้ำ + ประตู, ปุ่ม CLEAR / PREVIEW / SAVE / EXPORT TO PDF ALL PAGES / EXPORT SUMMARY 4 PAGES / EXPORT TO EXCEL |
| `layout.js` | ผังในชีท "ผังที่นั่ง" (ตำแหน่งช่อง, ชื่อ, สูตรสรุป) ใช้ร่วมกันระหว่าง EXPORT TO EXCEL และ Apps Script |
| `export-excel.js` | สร้างไฟล์ EXPORT TO EXCEL 5 แผ่นงาน |
| `config.js` | URL ของ Apps Script Web App |
| `apps-script/server.js` | โค้ดฝั่ง Apps Script (อ่าน/บันทึก/ประวัติ/สร้างแท็บ) |
| `apps-script/Code.gs` | ไฟล์ที่วางใน Apps Script = `server.js` + `layout.js` (สร้างด้วย `tools/build-code-gs.sh`) |
| `docs/` | คู่มือสำหรับนำไปประยุกต์กับโปรเจกต์อื่น |

## ผังที่นั่ง

| โซน | ที่นั่ง | Monitor | Galley / ห้องน้ำ / ประตู |
|---|---|---|---|
| FRONT CREW ZONE (3) | Commander Center 1A (ห้อง), 1D 1F | Wall 21.5" หน้า CC 1A, หน้า 1D 1F | LAV A (ซ้าย), FWD GALLEY G1 (ขวา), DOOR 1L / 1R หน้าห้อง Commander Center |
| VVIP OA (2) | 2A 3A, SOFA (ฝั่งขวา) | Wall 21.5" หน้า 2A, หลัง 3A, หลังโซฟา | FWD VIP LAV (ฝั่งขวา เหนือโซฟา) |
| VIP OB (6) | 5B 5D 5F, 6B 6D 6F (โต๊ะระหว่างแถว 5-6) | Wall 21.5" หน้า 5B, หน้า 5D-5F, หลัง 6B, หลัง 6D-6F | EMERGENCY EXIT (OVERWING) FWD L / R หลังแถว 6 |
| VVIP OC (8) | 7B 7C 7D 7F, 8B 8C 8D 8F | Wall 21.5" หน้า 7B-7C, 7D-7F, หลัง 8B-8C, 8D-8F | CENTER VIP LAV (ฝั่งขวา) + EMERGENCY EXIT (OVERWING) AFT L / R แถวเดียวกัน |
| PASSENGER OD (54) | แถว 9–18 (ไม่มีแถว 13) A–F | Wall 21.5" หน้า 9A-9C, 9D-9F · OVHD Monitor 12.0" หน้าแถว 10, 12, 15, 17 (ซ้าย/ขวา) | LAV D / LAV E หลังแถว 18, DOOR 4L / 4R, AFT GALLEY G5 |

ชนิดจอ: Wall Mounted Monitor 21.5 Inch (15 จอ), Bin Mounted Retractable Monitors (OVHD Monitor) 12.0 Inch (8 จอ)

### Galley (ตามคู่มือ A320 CCOM-S 02-140)
- **FWD GALLEY = G1** (18 รายการ): Stow 1-4, Refrigerator, Steam Oven, CM 1-2, Hot Cup, Faucet (มี Sink), Work Light, Trolley Stowage 1-2, Microwave, Water Heater (ใน Stow 4), Waste, Air Chiller, 220VAC Power Outlet
- **AFT GALLEY = G5** (27 รายการ): Stow 1-10, Water Boiler, CM 1-2, Steam Oven 1-3, Microwave, Work Light, Faucet (มี Sink), 220VAC Power Outlet, Waste Compartment, Water Heater (ใน Stow 9), Full Size Trolley Stowage 1-5
- สภาพอุปกรณ์และห้องน้ำ 3 ระดับ: **ใช้งานได้ / ใช้งานได้บางส่วน / ใช้งานไม่ได้**

## ปุ่มและรายงาน
| ปุ่ม | ผลลัพธ์ |
|---|---|
| PREVIEW ALL PAGES / PREVIEW SUMMARY 4 PAGES | ดูรายงานก่อน EXPORT (บนมือถือย่อให้เห็นเต็มหน้ากระดาษ) |
| EXPORT TO PDF ALL PAGES | A4 แนวตั้ง: สรุปตามโซน, Monitor, Galley (แยก G1 / G5), ห้องน้ำ, ผังทุกโซน, รายการข้อขัดข้อง |
| EXPORT SUMMARY 4 PAGES | A4 แนวนอน: 1 ผังที่นั่ง + ประตู, 2 ข้อขัดข้องที่นั่ง, 3 ผัง Galley และห้องน้ำ, 4 ข้อขัดข้อง Galley และห้องน้ำ |
| EXPORT TO EXCEL | 5 แผ่นงาน: ผังที่นั่ง (มี Galley / ห้องน้ำ / ประตู) · Galley และห้องน้ำ · รวมข้อขัดข้องคงค้าง · ทุกที่นั่ง และ Monitor · ประวัติการแก้ไข |

ชื่อไฟล์: `YYMMDD HHMM TH-LT A320-03_STATUS_SEAT_PLAN_ALL PAGES.pdf`, `..._SUMMARY_4PAGES.pdf`, `..._STATUS_SEAT_PLAN.xlsx`

## การติดตั้ง / อัปเดต (ทำ Apps Script ก่อน แล้วจึงขึ้นหน้าเว็บ)

### 1. Apps Script ใน Google Sheet
1. เปิด Google Sheet **Seat Plan 320(03) (For Admin)** > **ส่วนขยาย (Extensions) > Apps Script**
2. ลบโค้ดเดิมทั้งหมดใน `Code.gs` แล้ววางโค้ดจาก `apps-script/Code.gs`
3. ถ้าต้องการรหัส ให้แก้บรรทัด `WRITE_PIN` / `ADMIN_PIN` (ส่วน "ตั้งค่า" ต้นไฟล์) และตรวจ `INIT_BY` (ชื่อผู้บันทึกค่าเริ่มต้นของ Galley / ห้องน้ำ)
4. กด **บันทึก (Save)** แล้วเลือกฟังก์ชัน **setupSheet** กด **Run**
   - สร้าง/จัดแท็บ **ผังที่นั่ง**, **Galley และห้องน้ำ**, **รวมข้อขัดข้องคงค้าง**, **ทุกที่นั่ง และ Monitor**, **ประวัติการแก้ไข**
   - สภาพและหมายเหตุของที่นั่ง/จอเดิมยังอยู่ (ทดสอบการย้ายจากผังรุ่นเดิมครบ 96/96 รายการ) ประวัติไม่หาย
   - Galley / ห้องน้ำ เริ่มที่ "ใช้งานได้" และบันทึกค่าเริ่มต้นลงประวัติ
5. **Deploy > Manage deployments >** ดินสอ > Version **New version** > Deploy (URL เดิมใช้ต่อได้ ไม่ต้องแก้ `config.js`)

### 2. หน้าเว็บ (GitHub Pages)
1. Commit ไฟล์ทั้งหมดขึ้น branch **main** (ลบ `tools/build-code-gs.js` ตัวเดิมออก ใช้ `tools/build-code-gs.sh` แทน)
2. ถ้ายังไม่เปิด: **Settings > Pages** > Deploy from a branch > **main** / **(root)**
3. หน้าเว็บ: `https://rtaf602.github.io/a320-03_seatplan/` แถบด้านบนจะขึ้น "ข้อมูลจาก Google Sheet"

## เมื่อแก้ผัง
แก้ `layout.js` (ผังในชีท) และ `index.html` (ผังบนเว็บ) ให้ตรงกัน แล้วรัน `sh tools/build-code-gs.sh`
เพื่อสร้าง `apps-script/Code.gs` ใหม่ วางใน Apps Script, รัน **setupSheet** และ Deploy New version

## การทำงาน
- หน้าเว็บอ่าน/บันทึกที่แท็บ "ผังที่นั่ง" และดึงข้อมูลใหม่ทุก 30 วินาที
- แต่ละรายการในชีทมี 3 ช่องเรียงลงมา: ชื่อ / สภาพ (dropdown) / NOTE · Apps Script หาตำแหน่งจากชื่อ จึงห้ามแก้ข้อความชื่อในชีท
- SAVE เขียนเฉพาะช่องที่แก้จริง ถ้าชนกับคนอื่นจะถามก่อน ไม่บันทึกทับ
- ถ้ายังไม่ได้ใส่ `API_URL` หน้าเว็บยังเปิดดูและแก้ได้ แต่ SAVE ไม่ได้

## ข้อควรรู้ด้านความปลอดภัย
- Web App ตั้งเป็น "Anyone" ใครได้ URL ก็อ่าน/บันทึกได้ จึงแนะนำให้ตั้ง `WRITE_PIN`
- Repo เป็น public ห้ามใส่รหัสใน `config.js` หรือไฟล์ใดใน repo ให้ตั้งรหัสเฉพาะใน Apps Script
- ควรล็อกแท็บ "ประวัติการแก้ไข" ด้วย Protect sheet ให้เฉพาะ Admin
