# thai-docs-to-excel

ใบเสร็จภาษาไทย (รูปภาพ) → ไฟล์ Excel โดยรันบนเครื่องตัวเองทั้งหมด ไม่ส่งข้อมูลออกเน็ต
Thai receipt images → Excel, fully local (no data leaves your machine).

ช่องที่ "ตรวจแล้วไม่ผ่านกฎ" จะถูกระบายสีแดงให้คนไปดูเอง / Cells that fail a rule check are filled red so a human can look at them.

**Pipeline:** OpenThai2.0-OCR-Text (ONNX, อ่านทีละบรรทัด) → แก้ตัวเลขที่ OCR มักอ่านพลาด → Typhoon 2.5 Qwen3 4B (GGUF ผ่าน `llama-server`, บังคับเอาต์พุตเป็น JSON schema) → ตรวจด้วยกฎ → `.xlsx`

## ติดตั้งและใช้งาน / Install and run (macOS, Apple Silicon)

ทำใน folder ของ repo นี้ / Run everything from the repo folder.

**1. สร้าง venv / Python venv**
```bash
uv venv --python 3.12 .venv && source .venv/bin/activate
uv pip install onnxruntime opencv-python-headless pyclipper numpy pillow huggingface_hub openpyxl
```

**2. ดาวน์โหลดโมเดล OCR (ประมาณ 10 MB) / OCR model**
```bash
hf download iapp/OpenThai2.0-OCR-Text --exclude "pytorch/*" --local-dir text-model
```

**3. ติดตั้ง llama.cpp และดาวน์โหลดโมเดลภาษา (2.5 GB) / llama.cpp + the language model**
```bash
brew install llama.cpp
hf download typhoon-ai/typhoon2.5-qwen3-4b-gguf typhoon2.5-qwen3-4b-q4_k_m.gguf --local-dir llm
```

**4. เปิดเซิร์ฟเวอร์ (ค้างไว้ในอีกหน้าต่าง) / Start the server (leave it running in another terminal)**
```bash
llama-server -m llm/typhoon2.5-qwen3-4b-q4_k_m.gguf --port 8088 -c 4096 -ngl 99
```

**5. รัน / Run**
```bash
python pipeline.py out.xlsx samples/sample-receipt.jpg
# หลายใบ / several: python pipeline.py out.xlsx receipts/*.jpg
```
Options: `--model-dir` (หรือ env `THAI_OCR_MODEL_DIR`, default `./text-model`), `--server` (หรือ env `LLAMA_SERVER_URL`, default `http://127.0.0.1:8088`).

ผลลัพธ์: `out.xlsx` 1 แถวต่อ 1 รายการสินค้า + JSON ของทุกใบพิมพ์ออกหน้าจอ (ฟิลด์ที่อ่านได้, ผลตรวจแต่ละข้อ, บรรทัดที่ OCR มั่นใจต่ำ, เวลาที่ LLM ใช้)
Output: `out.xlsx`, one row per line item, plus a JSON dump on stdout (extracted fields, each check result, low-confidence OCR lines, LLM seconds).

**ผลที่คาดจากรูปตัวอย่าง / Expected result on the sample:** 1 แถว (บริษัท บิลด์พลัส แมททีเรียล จำกัด, INV-6517, ยอด 325.28) คอลัมน์วันที่แดง เพราะในรูปที่อยู่ซ้อนทับวันที่ OCR จึงอ่านเป็น `129/01/2567` นี่คือตัวอย่างว่าช่องแดงทำงานอย่างไร
One row; the date cell is red because the address overlaps the date in the image and OCR reads `129/01/2567`. That is the check working.

**เวลา / Speed (M1, 8 GB):** ประมาณ 11 วินาทีต่อใบ (OCR ~1 วิ + LLM ~10 วิ) ใบแรกหลังเปิดเครื่องช้ากว่า (~18 วิ บวกเวลาโหลดโมเดล) / about 11 s per receipt, the first run is slower. RAM: `llama-server` ~650 MB RSS (น้ำหนักโมเดลอยู่ใน Metal จึงนับไม่ครบ / weights live in Metal memory so RSS undercounts) + `pipeline.py` ~320 MB.

## ช่องสีแดงหมายถึงอะไร / What the red cells mean

แดง = "โปรแกรมไม่แน่ใจ ให้คนเช็ก" ไม่ได้แปลว่าผิดแน่ ๆ และ **ไม่แดง ก็ไม่ได้แปลว่าถูกแน่ ๆ**
Red = "check this by eye". Not red does not guarantee it is right.

| ช่องที่แดง / Cell | กฎ / Rule |
|---|---|
| ราคารวม (H) | ตัวเลขนี้ต้องปรากฏอยู่ในข้อความ OCR จริง (กัน LLM แต่งเลข) / the number must appear in the OCR text (guards against the LLM inventing numbers) |
| ยอดสุทธิ (I) | ผลรวมรายการ = ยอดสุทธิ หรือ ผลรวม + VAT = ยอดสุทธิ (คลาดได้ 0.05) / sum of items = total, or sum + VAT = total (±0.05) |
| เลขผู้เสียภาษี (E) | เลข 13 หลักต้องผ่าน checksum / 13 digits with a valid checksum |
| เลขที่ (D) | เลขที่ใบเสร็จต้องยาว ≥ 4 และมีในข้อความ OCR / receipt number must be in the OCR text |
| วันที่ (C) | แปลงเป็นวันที่ได้ (รองรับปี พ.ศ./ค.ศ., เลขไทย, ม.ค.–ธ.ค., 2 หลัก) และอยู่ในช่วงที่เป็นไปได้ / parses to a plausible date, incl. Buddhist-era years |

คอลัมน์ "ตรวจผ่าน?" = "ผ่าน" เมื่อราคารวมของแถวนั้นอยู่ในข้อความ OCR และผลรวมตรงยอดสุทธิ ไม่งั้น "ตรวจ"
ตรวจ VAT 7/107 ของยอดสุทธิ หรือ 7% ของผลรวม ด้วย (อยู่ใน JSON; ยังไม่ระบายสี) / the VAT check (7/107 of total or 7% of sum) is reported in the JSON but has no colored cell of its own.

## ข้อจำกัด / Limits

- อ่านได้เฉพาะ **ข้อความพิมพ์เป็นบรรทัด** ไม่เข้าใจตารางหรือ layout / printed line text only, no table or layout understanding
- **ลายมืออ่านไม่ได้** / handwriting is not readable
- ใบเสร็จยาว รายการสินค้ามัก **ไม่ครบ** (LLM 4B ตกหล่นรายการ) ผลรวมจึงมักแดง / on long receipts the item list is often incomplete; the sum check then turns red
- ใบที่อ่านไม่ได้เลยสักรายการจะไม่มีแถวใน Excel (ดูใน JSON) / a receipt with no extracted items gets no row in the Excel file (see the JSON)
- ทดสอบเฉพาะบน Apple M1 RAM 8 GB เท่านั้น / tested only on an M1 with 8 GB RAM; other machines are untested
- ต้องมีเน็ตตอนดาวน์โหลดโมเดลครั้งแรก หลังจากนั้นรัน offline ได้ / the first model download needs internet; after that it runs offline
- ปรับให้เข้ากับใบเสร็จร้านค้าไทย ไม่ได้ทดสอบกับเอกสารชนิดอื่น / tuned for Thai shop receipts, not other document types

## เครดิตและลิขสิทธิ์ / Credits and licenses

- โค้ดใน repo นี้ / Our code: **MIT** (ไฟล์ `LICENSE`)
- OCR: [iapp/OpenThai2.0-OCR-Text](https://huggingface.co/iapp/OpenThai2.0-OCR-Text) โดย iApp Technology / OpenThai, **Apache-2.0** (อิงจาก PP-OCRv6 ของ PaddlePaddle, Apache-2.0) โมเดลและ `runtime/` ไม่อยู่ใน repo นี้ ดาวน์โหลดจาก Hugging Face ตอนติดตั้ง / the model and its `runtime/` code are downloaded from Hugging Face, not bundled here
- LLM: [typhoon-ai/typhoon2.5-qwen3-4b-gguf](https://huggingface.co/typhoon-ai/typhoon2.5-qwen3-4b-gguf) โดย SCB 10X / Typhoon, **Apache-2.0** (ไม่ได้รวมใน repo / not bundled)
- [llama.cpp](https://github.com/ggml-org/llama.cpp) (MIT)
- รูปตัวอย่าง / Sample image: [DuckerMaster/Thai-Synth-Receipts](https://huggingface.co/datasets/DuckerMaster/Thai-Synth-Receipts) (MIT) เอกสารสังเคราะห์ ไม่ใช่ใบเสร็จของคนจริง / synthetic, not a real receipt
