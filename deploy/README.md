# ติดตั้งหน้าเว็บแปลง PDF → Word บน Ubuntu Server (ใช้ CPU)

ชุดนี้รันหน้าเว็บด้วย Docker บน CPU อย่างเดียว ไม่ต้องมีการ์ดจอและไม่ใช้ vLLM

## ความต้องการของ server
- Ubuntu 22.04 หรือ 24.04 แบบ x86_64
- RAM 8 GB ขึ้นไป
- พื้นที่ดิสก์ว่าง 12 GB ขึ้นไป (โมเดล 3 GB, Docker image ราว 3 GB และพื้นที่ระหว่าง build)
- ต่ออินเทอร์เน็ตได้ตอนติดตั้ง เพื่อดาวน์โหลดโมเดลและแพ็กเกจ
- ผู้ใช้ที่มีสิทธิ์ `sudo`

ความเร็วขึ้นกับ CPU: เครื่องทดสอบ (6 thread) ใช้ราว 2-3 นาทีต่อหน้า แปลงทีละไฟล์ตามคิว แต่ในไฟล์เดียวกันจะอ่านหลายหน้าพร้อมกันตามจำนวน copy ของโมเดล (ดู `OCR_REPLICAS`)

## ขั้นตอน

### 1. คัดลอกไฟล์ขึ้น server
จากเครื่อง Windows (แก้ชื่อผู้ใช้และที่อยู่ server):

```bash
scp dist/pdf2docx-web.tar.gz user@server:/tmp/
```

### 2. แตกไฟล์
บน server:

```bash
sudo mkdir -p /opt/pdf2docx && sudo chown "$USER" /opt/pdf2docx
tar -xzf /tmp/pdf2docx-web.tar.gz -C /opt/pdf2docx
cd /opt/pdf2docx
```

### 3. ติดตั้ง
```bash
bash deploy/setup.sh --install-docker
```
ถ้า server มี Docker อยู่แล้ว ไม่ต้องใส่ `--install-docker`

สคริปต์จะทำตามลำดับ:
1. ติดตั้ง Docker Engine (เฉพาะเมื่อใส่ `--install-docker` และยังไม่มี)
2. สร้าง `deploy/.env` พร้อมสุ่มรหัสผ่านเข้าเว็บ
3. ดาวน์โหลดโมเดล 2.7 GB ลง `model/` และตรวจ checksum
4. build image แล้วเปิดเว็บ
5. แสดงที่อยู่เว็บ ชื่อผู้ใช้ และรหัสผ่าน

รันซ้ำได้: สคริปต์เก็บ `.env` เดิม ข้ามการดาวน์โหลดถ้ามีโมเดลแล้ว แล้ว build และเปิดเว็บใหม่

ถ้ามีโฟลเดอร์ `model/` จากเครื่อง Windows อยู่แล้ว คัดลอกไปวางที่ `/opt/pdf2docx/model` ก่อนรันสคริปต์ได้ จะไม่ต้องดาวน์โหลดซ้ำ (ไม่ต้องคัดลอกโฟลเดอร์ `.git` ข้างใน)

## การเข้าใช้งาน
ค่าเริ่มต้นเว็บฟังเฉพาะบน server เอง (`BIND=127.0.0.1`) เลือกวิธีเข้าใช้ได้สองแบบ

**แบบ A: SSH tunnel** (ไม่ต้องเปิดพอร์ตเพิ่ม)
```bash
ssh -L 8080:127.0.0.1:8080 user@server
```
แล้วเปิด <http://localhost:8080> บนเครื่องตัวเอง

**แบบ B: เปิดให้เครื่องในเครือข่ายเข้าได้**
แก้ `deploy/.env` เป็น `BIND=0.0.0.0` แล้วรัน

```bash
cd /opt/pdf2docx/deploy && docker compose up -d
```
ถ้าเปิด `ufw` ไว้ ให้อนุญาตพอร์ตด้วย `sudo ufw allow 8080/tcp`

## ผลลัพธ์ที่ได้
- ข้อความที่ระบบไม่มั่นใจจะถูกไฮไลต์สีเหลืองในไฟล์ Word ให้ตรวจกับต้นฉบับก่อน จุดที่ไม่ไฮไลต์ก็ยังอาจผิดได้
- จัดหน้าให้ใกล้ต้นฉบับ หนึ่งหน้า PDF เป็นหนึ่งหน้า Word ใช้ฟอนต์ TH SarabunPSK (ต้องติดตั้งบนเครื่องที่เปิดไฟล์)
- รูปภาพ แผนที่ และกราฟ เป็นกรอบว่างเขียนว่า "รูปภาพ" ตัวรูปและลายเซ็นไม่ถูกนำมาใส่

## ความปลอดภัย
- เว็บใช้การล็อกอินแบบ HTTP Basic ชื่อผู้ใช้และรหัสผ่านอยู่ใน `deploy/.env`
- การเชื่อมต่อเป็น HTTP ไม่เข้ารหัส รหัสผ่านและเอกสารจึงถูกดักอ่านได้บนเครือข่าย ใช้แบบ B ได้เฉพาะในเครือข่ายภายในที่ไว้ใจ
- ถ้าจะเปิดออกอินเทอร์เน็ต ให้คง `BIND=127.0.0.1` แล้ววาง reverse proxy ที่มี HTTPS (เช่น Caddy หรือ Nginx) ไว้ด้านหน้า
- ผู้ใช้ทุกคนใช้บัญชีเดียวกัน และเห็นไฟล์ของกันและกันในรายการ
- ไฟล์ที่อัปโหลดและผลลัพธ์เก็บใน `webdata/` จนกว่าจะกดลบในหน้าเว็บ

## คำสั่งดูแลระบบ
รันจากโฟลเดอร์ `/opt/pdf2docx/deploy`

| งาน | คำสั่ง |
|---|---|
| ดู log | `docker compose logs -f` |
| หยุด | `docker compose down` |
| เปิด | `docker compose up -d` |
| ดูสถานะ | `docker compose ps` |
| เปลี่ยนรหัสผ่านหรือพอร์ต | แก้ `.env` แล้ว `docker compose up -d` |
| อัปเดตโค้ด | แตกไฟล์ชุดใหม่ทับ แล้ว `bash deploy/setup.sh` |
| ล้างไฟล์งานทั้งหมด | `docker compose down` แล้วลบไฟล์ใน `../webdata/` |

เว็บเปิดเองหลัง server รีสตาร์ต (`restart: unless-stopped`) งานที่ค้างจะกลับเข้าคิวและทำต่อจากหน้าที่ค้าง

## ค่าที่ปรับได้ใน `deploy/.env`
| ค่า | ความหมาย |
|---|---|
| `BIND`, `PORT` | ที่อยู่และพอร์ตที่เว็บฟัง |
| `WEB_USER`, `WEB_PASSWORD` | บัญชีเข้าเว็บ (รหัสผ่านว่างได้เฉพาะเมื่อ `BIND=127.0.0.1`) |
| `MAX_UPLOAD_MB` | ขนาด PDF สูงสุด (ค่าเริ่มต้น 200) |
| `OCR_REPLICAS` | จำนวน copy ของโมเดลที่อ่านหน้าพร้อมกัน (ว่าง = เลือกเองจากจำนวน CPU และ RAM ที่ว่าง; copy ละราว 1.1 GB แบบ fp32) |
| `OCR_THREADS` | จำนวน thread ต่อ copy (ว่าง = 4; บน Xeon E5 v4 ค่า 2 ใช้ core ได้เต็มกว่า) |
| `OCR_DTYPE` | `auto` (ใช้ bf16 ถ้า CPU รองรับในตัว), `fp32`, `bf16` |
| `OCR_QUANT` | `none` หรือ `int8` (เร็วขึ้นและใช้ RAM น้อยลง แต่ต้องตรวจคุณภาพก่อน) |
| `PUID`, `PGID` | ผู้ใช้ที่เป็นเจ้าของ `webdata/` |
| `MODEL_REPO`, `MODEL_REVISION` | โมเดลและรุ่นที่ดาวน์โหลด |

## ปรับความเร็วบน CPU
โมเดลหนึ่ง copy สร้างข้อความทีละ token จึงใช้ core ได้ไม่เกินราว 4-6 core ถึงจะเห็น CPU 100% ก็ตาม
(thread ที่เหลือวนรอ) ระบบจึงเปิดโมเดลหลาย copy ให้แต่ละ copy อ่านคนละหน้า
บน Xeon E5-2680 v4 (VM 16 vCPU) พบว่าแต่ละ copy ใช้ได้จริงราว 2 thread จึงแนะนำ `OCR_REPLICAS=8`, `OCR_THREADS=2`

วัดความเร็วบนเครื่องตัวเอง (8 หน้าแรกของไฟล์ทดสอบ):
```bash
cd /opt/pdf2docx/deploy
docker compose stop web          # ไม่ให้โมเดลของหน้าเว็บแย่ง CPU/RAM ระหว่างวัด
# แบบเดิม: 1 copy x 16 thread
docker compose run --rm web python3 /app/pdf2docx.py /jobs/test.pdf -o /jobs/a.docx --local --pages 1-8 --replicas 1 --threads 16
# แบบใหม่ (อัตโนมัติ)
docker compose run --rm web python3 /app/pdf2docx.py /jobs/test.pdf -o /jobs/b.docx --local --pages 1-8
docker compose start web
```
(คัดลอก PDF ไปไว้ที่ `webdata/test.pdf` ก่อน และลบ `webdata/a.pages`, `webdata/b.pages` ถ้าจะวัดซ้ำ)
ดูค่า `~X.Xs/หน้า` ตอนท้าย แล้วลองปรับ `OCR_REPLICAS`/`OCR_THREADS` เช่น 3x5, 4x4, 8x2 แล้วเลือกค่าที่เร็วที่สุด

## แก้ปัญหา
- **หน้าเว็บขึ้น "กำลังโหลดโมเดล" นาน**: ปกติใช้ไม่เกิน 1 นาที ดูสาเหตุด้วย `docker compose logs`
- **container หยุดเองระหว่างแปลง**: มักเกิดจาก RAM ไม่พอ ตรวจด้วย `dmesg | grep -i "out of memory"`
- **เขียนไฟล์ใน `webdata/` ไม่ได้**: รัน `sudo chown -R $(id -u):$(id -g) ../webdata` ให้ตรงกับ `PUID`/`PGID`
- **checksum ไม่ผ่าน**: ลบ `model/model.safetensors` แล้วรัน `setup.sh` ใหม่
