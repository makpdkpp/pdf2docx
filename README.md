# PDF → Word ด้วย OpenThai2.0-OCR-Markdown (RTX 3060 12GB)

## 1. เตรียมเครื่อง
- ใช้ Linux หรือ **WSL2** บน Windows (vLLM ไม่รองรับ Windows โดยตรง)
- ติดตั้ง NVIDIA driver และ CUDA ให้เรียบร้อย ทดสอบด้วยคำสั่ง `nvidia-smi`
- ติดตั้งฟอนต์ **TH Sarabun New** บนเครื่องที่จะเปิดไฟล์ Word

## 2. รัน vLLM server (เทอร์มินัลที่ 1)
```bash
python -m venv .venv && source .venv/bin/activate
pip install vllm albumentations timm open_clip_torch

vllm serve iapp/OpenThai2.0-OCR-Markdown \
  --served-model-name nemotron-parse \
  --trust-remote-code \
  --dtype bfloat16 \
  --gpu-memory-utilization 0.6 \
  --max-num-seqs 8
```
- `0.6` = ใช้ VRAM ประมาณ 7 GB เหลือที่ไว้ให้หน้าจอและโปรแกรมอื่น
- ถ้าเจอ error เรื่อง attention ให้เติม `--attention-backend=TRITON_ATTN`
  (NVIDIA แนะนำค่านี้สำหรับการ์ดตระกูล Ampere อย่าง A10/A100 ซึ่ง 3060 อยู่ตระกูลเดียวกัน)
- ถ้าเจอ error เรื่อง compile ให้เติม `--enforce-eager`

## 3. ติดตั้ง dependencies ของสคริปต์ (เทอร์มินัลที่ 2)
```bash
pip install pymupdf python-docx markdown
sudo apt install pandoc          # ต้องเป็นเวอร์ชัน 3.x ขึ้นไป
```
ต้องติดตั้งแพ็กเกจ `nemotron_parse` (ที่มีคลาส `DocumentParser`) ด้วย หน้าโมเดลมีตัวอย่างการ import แต่ไม่ได้บอกคำสั่งติดตั้ง
ให้ดูในไฟล์ `NVIDIA_NEMOTRON_PARSE_README.md` ใน repo ของโมเดลบน Hugging Face

## 4. ใช้งาน
```bash
# ลองกับ 10 หน้าแรกก่อน เพื่อวัดความเร็วจริงของ 3060
python pdf2docx.py เอกสาร.pdf -o เอกสาร.docx --pages 1-10

# ทั้งไฟล์
python pdf2docx.py เอกสาร.pdf -o เอกสาร.docx --workers 6

# แก้ Markdown ใน เอกสาร.pages/ เองแล้วสร้าง Word ใหม่ โดยไม่ต้อง OCR ซ้ำ
python pdf2docx.py เอกสาร.pdf -o เอกสาร.docx --from-md เอกสาร.pages
```

## ความสามารถ
- เรนเดอร์ PDF เป็นภาพทีละหน้า และเก็บภาพไว้บนดิสก์ไม่เกิน `workers×2` หน้า จึงรองรับไฟล์ 100MB ได้
- บันทึกผล OCR ของแต่ละหน้าไว้ใน `<ชื่อไฟล์>.pages/` ถ้าสคริปต์หยุดกลางทาง รันคำสั่งเดิมซ้ำจะทำต่อจากหน้าที่ค้าง
- ถ้าหน้าไหนยาวจนชนขีดจำกัด token จะแบ่งเป็นครึ่งบนกับครึ่งล่างแล้ว OCR ใหม่ให้อัตโนมัติ
- ตาราง HTML ที่มี rowspan/colspan แปลงเป็นตาราง Word ที่ผสานเซลล์จริง มีเส้นขอบ และแถวหัวตารางซ้ำทุกหน้า
- ขึ้นหน้าใหม่ใน Word ตามหน้าของ PDF และใช้ฟอนต์ TH Sarabun New ขนาด 16pt
