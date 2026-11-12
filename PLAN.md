# pdf2docx — แผนงานและบริบท (สำหรับทำต่อใน Claude Code)

## เป้าหมาย
เว็บแอปที่ผู้ใช้อัปโหลด PDF (สูงสุด 100MB) → OCR ภาษาไทย → ดาวน์โหลดเป็น MS Word (.docx)

## การตัดสินใจที่ทำแล้ว
- โมเดลหลัก: `iapp/OpenThai2.0-OCR-Markdown` (0.6B, รันบน vLLM)
  - input: ภาพหนึ่งหน้า 1664×2048
  - output: Markdown โดยตารางเป็น HTML (มี rowspan/colspan) พร้อม bbox และ class ของแต่ละบล็อก
- โมเดลเสริม (ยังไม่ได้ทำ): `iapp/OpenThai2.0-OCR-Text` รันบน CPU เร็วมาก แต่ได้ข้อความล้วน ใช้เป็นโหมด "เร็ว/ไม่มี GPU"
- ฮาร์ดแวร์: RTX 3060 12GB บน Linux/WSL2 ตั้ง vLLM ไว้ที่ `--gpu-memory-utilization 0.6 --max-num-seqs 8`
- แปลงเป็น Word: Markdown → HTML → pandoc (ใช้ reference.docx ที่ตั้งฟอนต์ TH Sarabun New ไว้) ส่วนตารางสร้างเองด้วย python-docx เพราะ pandoc 3.1 ทำ rowspan หาย

## สถานะ
- [x] ขั้นที่ 1: CLI `pdf2docx.py` ทดสอบกับตัวจำลองแล้ว
  - ผ่าน: ตารางผสานเซลล์, ขึ้นหน้าใหม่ตามหน้า PDF, ทำต่อจาก cache ได้, แบ่งครึ่งหน้าเมื่อชนขีดจำกัด token
  - ผลจากตัวจำลอง: `examples/sample_output.docx`
- [ ] ยังไม่ได้ทดสอบกับโมเดลจริงบน 3060 (ต้องวัดความเร็วต่อหน้าด้วย `--pages 1-10`)
- [ ] ยังไม่รู้วิธีติดตั้ง `nemotron_parse.DocumentParser` น่าจะอยู่ใน `NVIDIA_NEMOTRON_PARSE_README.md` ใน repo ของโมเดล
      ถ้าหาไม่เจอ ให้เปลี่ยนไปเรียก API `/v1/chat/completions` ของ vLLM ตรงๆ

## ขั้นต่อไป
2. Backend: FastAPI + Redis + RQ/Celery
   - อัปโหลดแบบ stream ลงดิสก์ ตั้ง nginx `client_max_body_size 110m`
   - ตรวจไฟล์ก่อนเริ่มงาน: นับหน้า, เช็กรหัสผ่าน, ตั้งเพดานจำนวนหน้า
   - endpoint สำหรับดูความคืบหน้า (หน้า X/Y) และดาวน์โหลดผลลัพธ์
3. Frontend: หน้าอัปโหลด, แถบความคืบหน้า, ปุ่มดาวน์โหลด
4. เพิ่มเติม:
   - ครอปรูปและแผนภูมิตาม bbox แล้วแทรกลง Word
   - หน้า preview ให้แก้ Markdown ก่อนสร้าง Word (ใช้ `--from-md`)
   - ลบไฟล์อัตโนมัติหลัง 24 ชั่วโมง, ระบบ login
   - Docker Compose: vllm, api, worker, redis

## ข้อจำกัดของโมเดลที่ต้องเผื่อไว้
- เลขไทยบางครั้งออกมาเป็นเลขอารบิก
- ตารางที่หน้าตาไม่เหมือนเอกสารราชการไทยอาจมีแถวหายหรือ rowspan ผิด
- ข้อความแนวตั้งอาจอ่านผิด
- ถ้าหน้าไหนชนขีดจำกัด token จะเกิด `GenerationError`
