# คู่มือยกระบบ Seat Plan ไปใช้กับเครื่องลำอื่น — กรณีศึกษา A320-05 → A320-03

เอกสารนี้สรุปขั้นตอนที่ใช้ยกระบบ **STATUS SEAT PLAN** (ที่นั่ง + Monitor + Galley + ห้องน้ำ + ประตู + ปุ่ม EXPORT)
จาก `rtaf602/a320-05_seatplan` มาใช้กับ `rtaf602/a320-03_seatplan` โดย **ไม่กระทบที่นั่งและจอ Monitor เดิม**
โปรเจกต์อื่นทำตามลำดับเดียวกันได้ (รายละเอียดการออกแบบ Galley / ห้องน้ำ ดู [`GALLEY_LAV_GUIDE.md`](GALLEY_LAV_GUIDE.md))

---

## 1. หลักการ: ยกโค้ดต้นแบบทั้งชุด แล้วเปลี่ยนเฉพาะ "ข้อมูลผัง"

ระบบแยกเป็น **โค้ด** (ใช้ร่วมกันได้ทุกลำ) กับ **ข้อมูลผัง** (เฉพาะแต่ละลำ)

| ไฟล์ | ยกมาทั้งไฟล์ | สิ่งที่ต้องแก้ให้เป็นของลำใหม่ |
|---|---|---|
| `index.html` | ✔ | ชื่อเครื่อง, `MT` / `MT_ORDER` (ชนิดจอ), `ZONES`, `EQ_CABIN`, ป้ายคำอธิบายสี (legend) |
| `layout.js` | ✔ | `SEAT_MT`, `SEAT_ZONES`, บล็อกในแต่ละโซนของ `seatPlanLayout()` |
| `export-excel.js` | ✔ | ชื่อเครื่องเท่านั้น |
| `apps-script/server.js` | ✔ | ไม่ต้องแก้ (อ่านผังจาก `layout.js`) |
| `config.js` | ✘ | **ใช้ของลำใหม่** (URL Apps Script ของชีทลำนั้น) |
| `tools/build-code-gs.sh` | ✔ | ชื่อเครื่องในหัวไฟล์ |

> ส่วน Galley (`GLAYOUT`) และห้องน้ำ (`LAVS`) ของ A320 ใช้ชุดเดียวกันได้ทุกลำที่ Galley เป็นแบบ G1 / G5
> ถ้า Galley ต่างกัน ให้แก้ตาม `GALLEY_LAV_GUIDE.md` ข้อ 2.1

## 2. ขั้นตอน

### 2.1 แทนชื่อเครื่องทั้งชุด
`A320-05` → `A320-03`, `a320-05` → `a320-03`, `32005` → `32003`
มีผลกับชื่อหน้าเว็บ ชื่อไฟล์ EXPORT และ key ใน localStorage ไฟล์ EXPORT จะได้ชื่อแบบเดียวกับต้นแบบ:
```
YYMMDD HHMM TH-LT A320-03_STATUS_SEAT_PLAN_ALL PAGES.pdf
YYMMDD HHMM TH-LT A320-03_STATUS_SEAT_PLAN_SUMMARY_4PAGES.pdf
YYMMDD HHMM TH-LT A320-03_STATUS_SEAT_PLAN.xlsx
```
(YY = ปี พ.ศ. 2 หลัก เวลาไทย TH-LT)

### 2.2 ยกที่นั่งและจอเดิมของลำใหม่มาใส่ `ZONES` (index.html)
คัดลอก `ZONES` ของ repo เดิมมา **โดยไม่แก้รหัสที่นั่ง/จอ** (`5B`, `MON-A3`, …) เพื่อให้ข้อมูลเดิมในชีทจับคู่ได้
แล้วแทรกบรรทัด Galley / ห้องน้ำ / ประตู ระหว่างบรรทัดเดิม:

```js
// FRONT CREW ZONE: LAV A + FWD GALLEY ด้านหน้า, DOOR 1L/1R ก่อนห้อง Commander Center
lines:[{fac:{L:[{lav:'A',room:1}],R:[{gal:'FWD'}]}},
       {door:{code:'1',kind:'door'}},
       {room:{...CC1A...},R:{mon:'MON-F2',seats:['1D','1F']}}]

// VVIP OA: FWD VIP LAV ฝั่งขวา (เหนือโซฟา)
lines:[{fac:{R:[{lav:'FV',seats:2}]}}, {mon:['MON-A1',null]}, {L2:['2A','3A'],sofa:'SOFA'}, {mon:['MON-A2','MON-A3']}]

// VIP OB: EMER EXIT FWD หลังแถว 6 (ต่อท้ายโซน)
... ,{mon:['MON-B3','MON-B4']}, {door:{code:'FWD',kind:'exit'}}

// VVIP OC: EMER EXIT AFT อยู่แถวเดียวกับ CENTER VIP LAV
lines:[{door:{code:'AFT',kind:'exit'},fac:{R:[{lav:'CV',seats:2}]}}, ...]

// PASSENGER OD: LAV D / LAV E หลังแถว 18, DOOR 4L/4R ก่อน AFT GALLEY
... ,{fac:{L:[{lav:'D'}],R:[{lav:'E'}]}}, {door:{code:'4',kind:'door'}}, {fac:{C:[{gal:'AFT'}]}}
```

ตำแหน่งห้องน้ำอ้างอิงจากไฟล์ Excel ผังที่นั่งจริงของเครื่อง (`320(03) SEAT`)

### 2.3 ชนิดจอ (`MT`)
ใส่เฉพาะชนิดที่ลำนั้นมี A320-03 ไม่มีจอ 27" จึงลบ `w27` ออกทั้งใน `MT`, `MT_ORDER`, legend และ `SEAT_MT` ของ `layout.js`
ใช้ key `w21` / `b12` เหมือนต้นแบบ (CSS สีจอผูกกับ key) แต่เปลี่ยน `name` / `short` ให้ตรงขนาดจริง (`12.0 Inch`)

### 2.4 โครงที่ต้นแบบไม่มี: โซฟา
A320-05 ไม่มีโซฟา หน้ากรอกรองรับ `{L2:[...], sofa:'SOFA'}` อยู่แล้ว แต่หน้า 1 ของ SUMMARY (`hCols` / `hZone`) ยังไม่รองรับ จึงเพิ่ม:
- `hCols`: ถ้ามี `sofa` ให้คอลัมน์แรกมีกล่อง `#SOFA` ที่ `sofa: <จำนวนคอลัมน์>`
- `hZone`: กล่องที่มี `sofa` ใช้ `gridColumn: span N` และไม่วางกล่องบนของคอลัมน์ที่ถูกโซฟาคลุม

ใน `layout.js` ใช้ `{ text: 'SOFA', c: R2S, span: 4, rows: 2 }`

### 2.5 หน้า 3 ของ SUMMARY (`EQ_CABIN`)
แก้ชื่อโซนและคำอธิบาย (`sub`) ให้ตรงลำใหม่ เช่น `'2A 3A · SOFA'`, `'CC 1A · 1D 1F'`

### 2.6 ผังในชีท (`layout.js`)
เขียนบล็อกของทุกโซนใหม่ให้ตรงกับ `ZONES` **รหัส (id) ต้องตรงกัน** และเรียง `mon()` / `eq()` บนลงล่าง ซ้ายไปขวา
เช็กด้วย Node ว่าไม่มีช่องซ้อนกันและจำนวนครบ:
```sh
node -e "eval(require('fs').readFileSync('layout.js','utf8'));const L=seatPlanLayout(),P=planSheet(L);
const k={};L.order.forEach(id=>{const t=L.items[id].kind;k[t]=(k[t]||0)+1});console.log(k);
const o={};let bad=0;P.ops.filter(x=>!x.cabin).forEach(x=>{for(let r=x.r;r<x.r+x.rs;r++)for(let c=x.c;c<x.c+x.cs;c++){if(o[r+','+c])bad++;o[r+','+c]=1}});console.log('overlaps',bad)"
```
ผลของ A320-03: `{ mon: 23, lav: 5, eq: 45, seat: 73 }`, `overlaps 0`

### 2.7 ย้ายข้อมูลจากชีทรุ่นเดิม (alias)
Apps Script หาตำแหน่งในชีทจาก **ข้อความชื่อ** ถ้าชื่อในชีทเดิมต่างจากชื่อใหม่ ให้ใส่ `alias` (ข้อความเดียวหรือ array)
ของ A320-03 ที่ต้องใส่:
- จอ OVHD: ชีทเดิมเขียน `OVHD 12.0"` → `SEAT_MT.b12.legacy = 'OVHD 12.0"'`
- ที่นั่ง CC 1A: ชีทเดิมเขียน `Commander Center 1A` → `items.CC1A.alias = 'Commander Center 1A'`

`server.js` อ่าน `alias` แบบ array ได้ (`[it.label].concat(it.alias || [])`)
จากนั้นรัน `setupSheet` → ค่าสภาพและ NOTE ของที่นั่ง/จอเดิมยังอยู่ ส่วน Galley/ห้องน้ำที่เพิ่มใหม่ตั้งเป็น "ใช้งานได้" และบันทึกลงประวัติ

## 3. ปรับปรุงจากต้นแบบระหว่างทำ A320-03 (นำกลับไปใช้กับ A320-05 ได้)
ทั้งหมดอยู่ใน `layout.js` แผ่นงาน "ผังที่นั่ง" ของ EXPORT TO EXCEL และในชีท:
1. หัวตาราง `Monitor\nเสื่อม\nสภาพ` ขึ้นบรรทัดตามคำ ไม่ให้ "พ" ตกบรรทัด
2. ป้าย EMER EXIT ขึ้นบรรทัดเอง 3 บรรทัด และแถวสูง 46px ให้เห็นตัวหนังสือครบ
3. ป้าย EXIT ฝั่งซ้ายย้ายไปคอลัมน์ A (พื้นขาว นอกลำ) คู่กับฝั่งขวาที่คอลัมน์ P · คอลัมน์ A กว้าง 60px
4. แถบ DOOR ลากจากคอลัมน์ A ถึง P สุดขอบทั้งสองฝั่ง

## 4. ขึ้นเว็บ (ลำดับสำคัญ)
1. **Apps Script ก่อน:** รัน `sh tools/build-code-gs.sh` → วาง `apps-script/Code.gs` ใน Apps Script ของชีทลำนั้น → Save → Run `setupSheet` → Deploy > Manage deployments > New version (URL เดิมใช้ต่อได้)
2. **แล้วจึงขึ้นหน้าเว็บ:** Commit ไฟล์ทั้งหมดขึ้น `main` ของ repo (GitHub Pages: Settings > Pages > main / root)
3. ตรวจ: เปิดหน้าเว็บ แถบบนขึ้น "ข้อมูลจาก Google Sheet" → ลอง SAVE 1 รายการ → ดูในแท็บ "ประวัติการแก้ไข"
4. ทดสอบ PREVIEW ทั้งสองแบบ, EXPORT TO PDF ALL PAGES, EXPORT SUMMARY 4 PAGES, EXPORT TO EXCEL บนคอมและมือถือ

> ถ้าขึ้นหน้าเว็บใหม่ก่อนเปลี่ยน Apps Script หน้าเว็บจะอ่าน/บันทึก Galley และห้องน้ำไม่ได้ เพราะ Apps Script รุ่นเดิมยังไม่รู้จักรายการเหล่านั้น
