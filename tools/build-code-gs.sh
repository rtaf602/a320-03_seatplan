#!/bin/sh
# สร้าง apps-script/Code.gs = layout.js + apps-script/server.js (ไฟล์เดียวสำหรับวางใน Apps Script)
cd "$(dirname "$0")/.." || exit 1
{
  printf '%s\n' '/**' \
    ' * Seat Plan A320-03 (For Admin) — Apps Script' \
    ' *' \
    ' * 1) สร้างแท็บในชีท: เลือกฟังก์ชัน setupSheet แล้วกด Run (ครั้งแรกครั้งเดียว)' \
    ' * 2) เป็น API ให้หน้าเว็บ (GitHub Pages) อ่าน/บันทึกสภาพที่นั่ง, Monitor, Galley และห้องน้ำ ลงแท็บ "ผังที่นั่ง"' \
    ' *    และบันทึกประวัติลงแท็บ "ประวัติการแก้ไข" ทุกครั้งที่กด SAVE บนเว็บ' \
    ' * 3) บันทึกประวัติอัตโนมัติเมื่อมีคนแก้ไขช่องในชีทโดยตรง (onEdit)' \
    ' *' \
    ' * ผังที่นั่ง (ตำแหน่งช่องของแต่ละที่นั่ง/จอ) อยู่ท้ายไฟล์ คัดลอกมาจาก layout.js' \
    ' * ไฟล์นี้สร้างจาก apps-script/server.js + layout.js ด้วย tools/build-code-gs.sh วิธีติดตั้งดู README.md' \
    ' */' ''
  cat apps-script/server.js
  printf '\n// ---------- ผังที่นั่ง (จาก layout.js) ----------\n'
  cat layout.js
} > apps-script/Code.gs
