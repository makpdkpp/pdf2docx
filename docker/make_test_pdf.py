"""Create a small 2-page Thai test PDF (heading, paragraph, list, merged-cell table)."""
import sys

import pymupdf

CSS = "body{font-family:'Sarabun','Loma','Garuda',sans-serif;font-size:14pt}" \
      "table{border-collapse:collapse;width:100%}td,th{border:1px solid #000;padding:4px}"

PAGES = [
    """<h1>รายงานงบประมาณประจำปี 2568</h1>
<p>เอกสารฉบับนี้จัดทำขึ้นเพื่อ<b>ทดสอบระบบ OCR ภาษาไทย</b> โดยมีทั้งหัวข้อ ย่อหน้า รายการ และตาราง</p>
<h2>1. วัตถุประสงค์</h2>
<ul><li>ตรวจสอบความถูกต้องของการอ่านสระและวรรณยุกต์</li>
<li>ตรวจสอบการแปลงตารางที่มีการผสานเซลล์</li>
<li>วัดความเร็วในการประมวลผลต่อหน้า</li></ul>
<h2>2. สรุปค่าใช้จ่าย</h2>
<table>
<tr><th rowspan="2">ลำดับ</th><th colspan="2">รายการ</th><th rowspan="2">จำนวนเงิน (บาท)</th></tr>
<tr><th>หมวด</th><th>รายละเอียด</th></tr>
<tr><td rowspan="2">1</td><td rowspan="2">วัสดุ</td><td>กระดาษและหมึกพิมพ์</td><td>12,500</td></tr>
<tr><td>อุปกรณ์สำนักงาน</td><td>8,300</td></tr>
<tr><td>2</td><td>เดินทาง</td><td>ค่าน้ำมันและค่าทางด่วน</td><td>4,750</td></tr>
<tr><td colspan="3"><b>รวมทั้งสิ้น</b></td><td><b>25,550</b></td></tr>
</table>""",
    """<h2>3. ข้อเสนอแนะ</h2>
<p>คณะกรรมการเห็นควรให้ปรับลดค่าใช้จ่ายหมวดวัสดุลงร้อยละ 10 ในปีงบประมาณถัดไป
และให้รายงานผลการดำเนินงานทุกไตรมาส</p>
<ol><li>จัดซื้อวัสดุร่วมกันระหว่างหน่วยงาน</li>
<li>ใช้เอกสารอิเล็กทรอนิกส์แทนกระดาษ</li></ol>
<p>ลงชื่อ ........................................ ผู้จัดทำรายงาน</p>
<p>วันที่ 13 พฤศจิกายน พ.ศ. 2568</p>""",
]

out = sys.argv[1] if len(sys.argv) > 1 else "test.pdf"
doc = pymupdf.open()
for html in PAGES:
    page = doc.new_page()                       # A4
    page.insert_htmlbox(pymupdf.Rect(56, 56, 539, 786), html, css=CSS)
doc.save(out)
print("saved", out, doc.page_count, "pages")
