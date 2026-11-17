"""Thai receipt image -> OpenThai OCR text -> fix numbers -> Typhoon 4B fields -> rule checks -> xlsx (red = check failed)

Usage:
    python pipeline.py out.xlsx img1.jpg img2.jpg ... [--model-dir text-model] [--server http://127.0.0.1:8088]

Environment (used when the flag is not given):
    THAI_OCR_MODEL_DIR   folder with the OpenThai2.0-OCR-Text files (default: ./text-model)
    LLAMA_SERVER_URL     llama-server base URL (default: http://127.0.0.1:8088)

The OCR code (runtime/line_engine.py) is not in this repo: it ships inside the
iapp/OpenThai2.0-OCR-Text Hugging Face model repo, so it is imported from --model-dir.
"""
import argparse, datetime, os, sys, re, json, time, urllib.request
from openpyxl import Workbook
from openpyxl.styles import PatternFill

LOW_CONF = 0.90
def fix_numbers(line):
    line = re.sub(r"(?<=\d),(?=\d{2}(?!\d))", ".", line)        # 128,00 -> 128.00 (comma decimal)
    line = re.sub(r"(?<=\d)\.(?=\d{3}(?!\d))", ",", line)        # 1.000 -> 1,000 only if 3 digits follow
    line = re.sub(r"(\s)0?00(?=\s*$)", r"\g<1>0.00", line)       # trailing 000 / 00 -> 0.00
    line = re.sub(r"(\d)[xX×K](?=\s)", r"\1x", line)            # 1.00K -> 1.00x
    return line

SCHEMA = {"type":"object","properties":{
  "shop_name":{"type":"string"},"tax_id":{"type":"string"},"date_as_written":{"type":"string"},
  "receipt_no":{"type":"string"},
  "items":{"type":"array","items":{"type":"object","properties":{
      "name":{"type":"string"},"qty":{"type":"number"},"line_total":{"type":"number"}},"required":["name","line_total"]}},
  "vat":{"type":"number"},"total":{"type":"number"}},"required":["shop_name","items","total"]}
PROMPT = """ข้อความด้านล่างอ่านจากใบเสร็จด้วย OCR ดึงข้อมูลเป็น JSON
กติกา:
- คัดลอกตัวเลขและรหัสตามที่เห็นในข้อความ ห้ามแก้ ห้ามเดา ถ้าไม่มีให้เว้นว่าง
- date_as_written = วันที่ตามที่พิมพ์ในใบเสร็จ ห้ามแปลงปี
- line_total = ตัวเลขท้ายสุดของบรรทัดรายการ (ราคารวมของบรรทัด) ไม่ใช่จำนวนชิ้น
- shop_name = ชื่อร้าน มักอยู่บรรทัดแรก ๆ
- total = ยอดสุทธิที่ต้องจ่าย

"""
LLM_URL="http://127.0.0.1:8088/v1/chat/completions"
def llm(text, url=None):
    url=url or LLM_URL
    body={"messages":[{"role":"user","content":PROMPT+text}],"temperature":0,
          "response_format":{"type":"json_schema","json_schema":{"name":"r","schema":SCHEMA}},
          "chat_template_kwargs":{"enable_thinking":False}}
    req=urllib.request.Request(url,json.dumps(body).encode(),{"Content-Type":"application/json"})
    return json.loads(json.load(urllib.request.urlopen(req,timeout=600))["choices"][0]["message"]["content"])

def tax_id_ok(t):
    d=re.sub(r"\D","",t or "")
    if len(d)!=13: return False
    s=sum(int(d[i])*(13-i) for i in range(12)); return (11-s%11)%10==int(d[12])
def num_in_text(x, text):
    if x is None: return False
    pats={f"{x:.2f}", f"{x:,.2f}", str(int(x)) if float(x).is_integer() else f"{x}"}
    return any(re.search(r"(?<![\d.])"+re.escape(p)+r"(?![\d])", text) for p in pats)
THMON={m:i+1 for i,m in enumerate("ม.ค. ก.พ. มี.ค. เม.ย. พ.ค. มิ.ย. ก.ค. ส.ค. ก.ย. ต.ค. พ.ย. ธ.ค.".split())}
ENMON={m:i+1 for i,m in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split())}
def _ymd(d,mo,y):
    if y<100: y = y+2500-543 if y>=50 else y+2000          # 2-digit: >=50 means BE (65 -> 2565 -> 2022)
    elif y>2400: y-=543                                     # Buddhist era -> CE
    if not (1<=mo<=12 and 1<=d<=31 and 2000<=y<=datetime.date.today().year+1): return ""   # implausible -> check fails
    return f"{y:04d}-{mo:02d}-{d:02d}"
def norm_date(s):
    s=(s or "").translate(str.maketrans("๐๑๒๓๔๕๖๗๘๙","0123456789"))
    m=re.search(r"(?<!\d)(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4}|\d{2})(?!\d)", s)
    if m: return _ymd(*map(int,m.groups()))
    m=re.search(r"(\d{1,2})\s*([ก-ฮ]{1,2}\.\s?[ก-ฮ]{1,2}\.)\s*(\d{4}|\d{2})(?!\d)", s)
    if m and m.group(2).replace(" ","") in THMON: return _ymd(int(m.group(1)),THMON[m.group(2).replace(" ","")],int(m.group(3)))
    m=re.search(r"(\d{1,2})\s+([A-Za-z]{3})\w*\s+(\d{4})", s)
    if m and m.group(2).lower() in ENMON: return _ymd(int(m.group(1)),ENMON[m.group(2).lower()],int(m.group(3)))
    return ""

def run(path, ocr):
    r=ocr.read(path)
    lines=[(fix_numbers(l["text"]), l["confidence"]) for l in r["lines"]]
    text="\n".join(t for t,_ in lines)
    t0=time.time(); f=llm(text); dt=time.time()-t0
    if not f.get("tax_id"):                                       # 13-digit number with a valid checksum
        found=[d for d in re.findall(r"(?<!\d)\d{13}(?!\d)", text) if tax_id_ok(d)]
        if found: f["tax_id"]=found[0]
    f["items"]=[it for it in f.get("items",[]) if not re.search(r"(?i)sub.?total|^total|รวม|ยอด|vat|ภาษี", it.get("name",""))]
    checks={}
    checks["tax_id"]= (not f.get("tax_id")) or tax_id_ok(f["tax_id"])
    checks["receipt_no"]= (not f.get("receipt_no")) or (len(f["receipt_no"])>=4 and f["receipt_no"] in text)
    date=norm_date(f.get("date_as_written","")) or norm_date(text)
    checks["date"]= bool(date)
    items=f.get("items",[])
    for i,it in enumerate(items): checks[f"item{i}"]=num_in_text(it.get("line_total"),text)
    checks["total_in_text"]=num_in_text(f.get("total"),text)
    s=sum(it.get("line_total") or 0 for it in items)
    tot=f.get("total") or 0; vat=f.get("vat") or 0
    checks["sum_items=total"]=abs(s-tot)<0.06 or (vat>0 and abs(s+vat-tot)<0.06)   # VAT-inclusive or VAT added on top
    if f.get("vat"): checks["vat_7pct"]=abs(f["total"]*7/107-f["vat"])<0.06 or abs(s*0.07-f["vat"])<0.06
    low=[t for t,c in lines if c<LOW_CONF]
    return {"file":path,"fields":f,"date":date,"checks":checks,"low_conf_lines":low,"llm_s":round(dt,1)}

def main():
    ap=argparse.ArgumentParser(description="Thai receipt images -> Excel with red cells where checks fail")
    ap.add_argument("xlsx", help="output .xlsx path")
    ap.add_argument("images", nargs="+", help="receipt image files")
    ap.add_argument("--model-dir", default=os.environ.get("THAI_OCR_MODEL_DIR","text-model"),
                    help="OpenThai2.0-OCR-Text folder (default: $THAI_OCR_MODEL_DIR or ./text-model)")
    ap.add_argument("--server", default=os.environ.get("LLAMA_SERVER_URL","http://127.0.0.1:8088"),
                    help="llama-server base URL (default: $LLAMA_SERVER_URL or http://127.0.0.1:8088)")
    a=ap.parse_args()
    if not os.path.isfile(os.path.join(a.model_dir,"runtime","line_engine.py")):
        sys.exit(f"OCR model not found in {a.model_dir!r}: run the hf download step in the README, or pass --model-dir")
    sys.path.insert(0, a.model_dir)
    from runtime.line_engine import LineOCR
    global LLM_URL; LLM_URL=a.server.rstrip("/")+"/v1/chat/completions"
    ocr=LineOCR(a.model_dir,threads=1); out=[]
    for p in a.images:
        t0=time.time(); out.append(run(p,ocr)); print(f"{p}: {time.time()-t0:.1f}s total",file=sys.stderr)
    wb=Workbook(); ws=wb.active; ws.title="items"; red=PatternFill("solid",fgColor="FFC7CE")
    ws.append(["ไฟล์","ร้าน","วันที่","เลขที่","เลขผู้เสียภาษี","รายการ","จำนวน","ราคารวม","ยอดสุทธิ","ตรวจผ่าน?"])
    for o in out:
        f=o["fields"]; c=o["checks"]
        for i,it in enumerate(f.get("items",[])):
            ok=c.get(f"item{i}") and c["sum_items=total"]
            ws.append([os.path.basename(o["file"]),f.get("shop_name",""),o["date"],f.get("receipt_no",""),f.get("tax_id",""),
                       it.get("name",""),it.get("qty"),it.get("line_total"),f.get("total"),"ผ่าน" if ok else "ตรวจ"])
            row=ws.max_row
            if not c.get(f"item{i}"): ws.cell(row,8).fill=red
            if not c["sum_items=total"]: ws.cell(row,9).fill=red
            if not c["tax_id"]: ws.cell(row,5).fill=red
            if not c["receipt_no"]: ws.cell(row,4).fill=red
            if not c["date"]: ws.cell(row,3).fill=red
    wb.save(a.xlsx); print(json.dumps(out,ensure_ascii=False,indent=1))

if __name__=="__main__":
    main()
