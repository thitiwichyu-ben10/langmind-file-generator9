# LangMind File Generator API

Backend API สำหรับสร้างไฟล์ Word (.docx), Excel (.xlsx), PDF (.pdf) และ PowerPoint (.pptx) จาก Markdown หรือ JSON แล้วส่ง URL ดาวน์โหลดกลับมา โดย PDF ใช้ Chromium ผ่าน Playwright

## 1) ทดสอบบนเครื่องตัวเอง

ต้องติดตั้ง Python 3.11 ขึ้นไป

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
playwright install chromium
uvicorn main:app --reload
```

เปิดเอกสาร API ที่ `http://127.0.0.1:8000/docs` และตรวจสุขภาพที่ `http://127.0.0.1:8000/health`

## 2) Deploy บน Render

1. สร้างบัญชี GitHub และ Render
2. สร้าง GitHub repository ใหม่ เช่น `langmind-file-generator`
3. แตก ZIP แล้วอัปโหลด **ไฟล์ทั้งหมดภายในโฟลเดอร์ `langmind_file_generator_api`** ไปไว้ที่ root ของ repository (ต้องเห็น `Dockerfile`, `render.yaml`, `main.py` ที่ root)
4. ใน Render เลือก **New + → Blueprint** แล้วเชื่อม GitHub repository นี้
5. Render จะอ่าน `render.yaml` และใช้ Dockerfile สร้าง Web Service ซึ่งติดตั้ง Chromium และ Font packages ให้โดยอัตโนมัติ
6. กดยืนยัน Deploy และรอจนสถานะเป็น Live
7. ทดสอบ `https://YOUR-SERVICE.onrender.com/health` ต้องได้ JSON ที่มี `"status":"ok"` และ `"pdf_engine":"chromium"`

**ถ้ามี Service เดิมอยู่แล้ว:** ตรวจสอบ Runtime/Source ของ Service ให้ใช้ Dockerfile จาก repository นี้ก่อน Deploy หาก Render ไม่ยอมเปลี่ยน Runtime ของ Service เดิม ให้สร้าง Web Service ใหม่จาก Blueprint แล้วค่อยเปลี่ยน Server URL ใน LangMind หลังทดสอบผ่าน

ไม่ต้องตั้ง Build Command หรือ Start Command เอง เพราะ Dockerfile กำหนดคำสั่งเริ่มระบบไว้แล้ว

## 3) ตั้งค่า LangMind

- **Server URL**: `https://YOUR-SERVICE.onrender.com`
- **Spec URL**: `https://YOUR-SERVICE.onrender.com/openapi.json`
- หรือคัดลอก OpenAPI YAML จาก `openapi.yaml` ไปวางในช่อง OpenAPI JSON/YAML

## 4) เรียก API

POST `/create-file`

ตัวอย่าง DOCX:

```json
{
  "filename": "quotation-2026.docx",
  "format": "docx",
  "content": "# ใบเสนอราคา\n\n| รายการ | จำนวน | ราคา |\n|---|---:|---:|\n| สินค้า A | 2 | 500 |\n"
}
```

ตัวอย่าง XLSX: ใช้ Markdown table ใน `content` ระบบจะนำตารางไปใส่ใน Excel

ตัวอย่าง PDF: ใช้ Markdown headings, paragraphs และ bullet list

ตัวอย่าง PPTX แบบ Markdown: ใช้ `#` หนึ่งหัวข้อต่อหนึ่งสไลด์ และข้อความ/รายการ bullet ใต้หัวข้อนั้น

ตัวอย่าง PPTX แบบ JSON:

```json
{
  "slides": [
    {"title": "ภาพรวม", "content": "สรุปโครงการ"},
    {"title": "แผนงาน", "bullets": ["ออกแบบ", "พัฒนา", "ทดสอบ"]}
  ]
}
```

Response:

```json
{
  "success": true,
  "filename": "quotation-2026.docx",
  "format": "docx",
  "download_url": "https://YOUR-SERVICE.onrender.com/download/FILE_ID",
  "expires_in_seconds": 86400
}
```

## 5) ข้อจำกัดและการเก็บไฟล์

ค่าเริ่มต้นเก็บไฟล์ใน `/tmp/generated_files` และลิงก์มีอายุ 24 ชั่วโมง ไฟล์บน local/ephemeral disk อาจหายเมื่อ Render รีสตาร์ตหรือ deploy ใหม่ จึงเหมาะกับการทดลองและไฟล์ชั่วคราว

ถ้าต้องการเก็บไฟล์ให้คงอยู่ ให้ใช้ object storage (เช่น S3-compatible storage) หรือ Render Persistent Disk บนแพ็กเกจที่รองรับ แล้วตั้ง `FILE_STORAGE_DIR` เป็น mount path ที่กำหนด ทั้งนี้ Persistent Disk มีค่าใช้จ่ายและอาจมีข้อจำกัดตามประเภท service/plan

## 6) ความปลอดภัยก่อนใช้งานจริง

- API ตัวอย่างนี้ยังไม่มี API key authentication; ควรเพิ่มการยืนยันตัวตนก่อนเปิดให้ใช้งานสาธารณะ
- จำกัดขนาดเนื้อหาได้ด้วย `MAX_CONTENT_CHARS`
- ตั้ง `PUBLIC_BASE_URL` หากต้องการบังคับโดเมน URL ที่ส่งกลับ
- อย่าใส่ข้อมูลลับหรือข้อมูลส่วนบุคคลลงในไฟล์โดยไม่จำเป็น


## PDF engine and Thai/English support (v1.1.0)

- PDF is rendered by Chromium using Playwright; DOCX/XLSX/PPTX generators are unchanged.
- Markdown is converted to HTML with raw HTML disabled, then printed to A4 PDF.
- Thai shaping and mixed Thai/English text use bundled Noto Sans Thai plus Noto font fallback installed by `apt.txt` on Render.
- The PDF layout supports headings, paragraphs, lists, tables, blockquotes, links and page breaks handled by Chromium.
- Request bodies must be UTF-8 JSON. Test with `ใบเสนอราคา / QUOTATION`, `รายละเอียดสินค้า / Product Description`, Thai numerals/Arabic numbers and table cells.

### Local setup for Chromium

After `pip install -r requirements.txt`, run `playwright install chromium`. On Linux machines missing browser libraries, run `playwright install --with-deps chromium`. For Render, use the included Dockerfile so Chromium libraries and Thai/Latin fonts are installed reproducibly.

### Render deployment note

The included `Dockerfile` installs Chromium runtime libraries plus Noto Thai/Latin font packages. `render.yaml` selects Render's Docker runtime and keeps the same `/health` endpoint. This is intentional because Chromium requires OS-level libraries that are more reproducible when packaged in the image.

