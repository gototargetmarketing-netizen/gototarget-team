# GO TO TEAM — ระบบภายในทีม GO TO TARGET

## ภาพรวม
- โปรเจคที่ผู้ใช้ทำ**ร่วมกับน้องชาย** สำหรับทีมในบริษัท GO TO TARGET ของน้อง (เอเจนซี่การตลาด: วิดีโอ กราฟิก คอนเทนต์)
- **เป้าหมาย (ผู้ใช้ยืนยัน 21 พ.ย. 2569):** แอดมินส่งข้อมูลและอัปเดตลิงก์งานของลูกค้า (Drive/Docs ฯลฯ) ให้ทีมเข้ามาดูแล้วโหลดไปทำ — **ไม่ใช่ระบบมอบหมายงาน** (ไม่มีผู้รับผิดชอบ/สถานะงาน) โครง Clients → Workspaces ของน้องตรงกับเป้าหมายแล้ว
- โค้ดตั้งต้นน้องออกแบบมา (ต้นฉบับอยู่ใน `backups/2026-11-21_ต้นฉบับจากน้อง/`)

## ขึ้นเว็บและ GitHub (ตั้งแต่ 21 พ.ย. 2569)
- หน้าเว็บ: https://gototarget-team.vercel.app/ — Vercel ของน้องขึ้นเว็บจาก branch `main` อัตโนมัติ
- ใช้ repo ของน้องร่วมกัน: `gototargetmarketing-netizen/gototarget-team` (**Public**, remote `origin`) — ไม่มี repo สำรองของ Beamsmile (ผู้ใช้เลือก)
  • **ทำไมต้อง Public (21 พ.ย. 2569):** Vercel ของน้องเป็นแผน Hobby ซึ่งไม่สร้างเว็บจาก commit ของคนนอกทีม Vercel เมื่อ repo เป็น Private (commit ของ Beamsmile ค้างหน้า "Deployment is building") → ผู้ใช้เลือกเปลี่ยนเป็น Public (แบบเดียวกับ repo `member` ของ LINE CRM) / ถ้าจะกลับเป็น Private ต้องอัป Vercel Pro แล้วเพิ่ม Beamsmile เข้าทีมก่อน
  • เพราะเป็น Public: ห้ามใส่รหัสผ่าน คีย์ลับ หรือข้อมูลลูกค้าจริงในไฟล์ที่ขึ้น repo (รวมถึงไฟล์นี้) — ตรวจแล้ว 21 พ.ย. 2569 ไม่มี
- **ส่งเข้า branch `dev` ก่อนเสมอ** เข้า `main` (= เว็บจริงเปลี่ยน) เมื่อผู้ใช้สั่งเท่านั้น
- หน้าทดสอบของ `dev`: https://gototarget-team-git-dev-go-to-target.vercel.app (น้องปิด Vercel Authentication แล้ว) — ผู้ใช้ชอบทดสอบที่นี่ (เปิดบนมือถือได้) ใช้ `.command` เมื่ออยากเห็นผลทันทีโดยไม่ต้อง push
- `git pull` / `git fetch` ก่อนแก้และก่อน push ทุกครั้ง — น้องแก้ไฟล์ผ่านหน้าเว็บ GitHub (commit "Update index.html")
- Claude Code โหมดอัตโนมัติกันการ push ขึ้น repo ของบัญชีอื่น → ผู้ใช้ต้องกดอนุญาต ห้ามหาทางเลี่ยง
- `backups/` อยู่ใน `.gitignore` (ไม่ขึ้น repo)
- `code.gs` ใน repo เป็นแค่ที่เก็บโค้ด — ต้องวางลง Apps Script เองแล้วกด Deploy > New version ถึงจะมีผล
- Apps Script ตัวจริงผูกกับชีต **"Go to Target Team"** (ชีตใช้งานจริง) — URL ใส่เป็นค่าหลักใน `index.html` บน `dev` แล้ว (21 พ.ย. 2569) ยังไม่เข้า `main`
  • แก้ `code.gs` แล้วให้ Deploy → Manage deployments → แก้ deployment เดิม → New version (ห้ามกด New deployment เพราะ URL จะเปลี่ยน)
  • ทดสอบกับ Apps Script ตัวอื่นได้โดยไม่แก้ไฟล์: ตั้ง `localStorage.setItem('gtt_web_app_url', '...exec')` ใน Console

## ไฟล์
- `index.html` — หน้าเว็บทั้งหมด (Tailwind CDN + Lucide) หน้า: Clients / หน้าลูกค้า / Workspaces / Settings (แอดมิน)
- `code.gs` — Apps Script หลังบ้าน ชีต `Clients`, `Workspaces`, `Users`
  • คำสั่งทางเดียวผ่าน `doPost` รูปแบบ `{ action, session_token, payload }` ส่งแบบ `text/plain`
  • เข้าระบบด้วย username + รหัสผ่าน → ได้ session token (CacheService, 6 ชม.) / ผิด 8 ครั้งล็อก 15 นาที
  • บทบาท ADMIN (แก้ได้ทุกอย่าง) / STAFF (ดูอย่างเดียว เห็นลูกค้าทุกราย)
  • ตั้งรหัสแอดมินจากเมนูในชีต `🚀 GO TO TARGET`

## กติกาความปลอดภัยที่ใช้แล้ว (21 พ.ย. 2569)
- ค่าที่ใส่ใน `onclick="..."` ต้องผ่าน `jsArg()` เสมอ (ห้ามใช้ `'${...}'` หรือ `escapeHtml` อย่างเดียว เพราะเบราว์เซอร์ถอด `&#39;` กลับก่อนรัน JS)
- ลิงก์ที่จะแสดง/เปิด ต้องผ่าน `safeUrl()` (รับเฉพาะ http/https) และเปิดด้วย `openExternalLink()`
- หลังบ้านตรวจทุกช่องด้วย `cleanText_` / `cleanUrl_` / `cleanId_` / `cleanEmail_` / `cleanChoice_` (ไอคอน/สีต้องอยู่ในรายการที่กำหนด)
- `appendObject_` / `updateObject_` ไม่ยอมเขียนข้อความที่ขึ้นต้นด้วย `= + - @` และตั้งรูปแบบเซลล์ `@` (ตัวเลข `0`) ก่อนเขียนทุกครั้ง
- รหัสผ่านเก็บแบบ `sha256$<hex>` (`verifyPassword_` ยังรับแบบ base64 เดิมได้)

## งานที่ยังค้าง (จากการตรวจ 21 พ.ย. 2569)
- [ ] แฮชรหัสผ่านรอบเดียว → เพิ่มจำนวนรอบ (ใช้ prefix ใหม่ เช่น `sha256i$`)
- [ ] ข้อความเข้าระบบไม่ผ่านบอกได้ว่า username มีอยู่จริง
- [ ] `importLocalDataIfEmpty` อาจอัปข้อมูลเก่าจากเบราว์เซอร์กลับขึ้นชีต + ข้อมูลตัวอย่าง SEED ใน index.html → ควรเอาออก
- [ ] ทุกคำสั่งเรียก `ensureReady_()` (อ่านชีตหลายรอบ ช้า และเขียนนอกล็อก)
- [x] สลับ Apps Script ทดสอบผ่าน localStorage (`gtt_web_app_url`), ไฟล์ `.command` (พอร์ต 8003), `testSetup()`, `noindex` — ทำแล้ว 21 พ.ย. 2569
- [ ] ปุ่ม "Save Configuration" ในหน้า Settings ใช้ไม่ได้แล้ว / `apiSecret`, `workspaceKey` ไม่ได้ใช้
- [ ] ล็อกเวอร์ชัน `lucide@latest` และเลิกใช้ `cdn.tailwindcss.com` บนเว็บจริง

## ข้อตกลงการเขียนโค้ด
- คอมเมนต์ในโค้ดเขียนเป็น**ภาษาอังกฤษ**ตามไฟล์เดิมของน้อง (ตาม CLAUDE.md ส่วนกลาง: แก้โค้ดที่คนอื่นเขียนภาษาอื่น ให้ตามแบบไฟล์เดิม) — ถ้าผู้ใช้อยากเปลี่ยนเป็นไทย ให้แก้บันทึกนี้
- ข้อความ UI ตอนนี้ปนอังกฤษ/ไทย ตามแบบที่น้องออกแบบ
