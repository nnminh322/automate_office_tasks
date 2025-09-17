from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
import datetime as dt

EXCEL_FILE   = Path("./data_sample.xlsx")
WORD_TEMPLATE= Path("./template_sample.docx")
OUTPUT_ROOT  = Path("./output")
OUTPUT_ROOT.mkdir(exist_ok=True)

doc_probe = DocxTemplate(str(WORD_TEMPLATE))
template_keys = set(doc_probe.get_undeclared_template_variables({}))
if not template_keys:
    raise ValueError("Template không có biến Jinja2 ({{...}}).")

raw = pd.read_excel(EXCEL_FILE, header=None, dtype=object)

def clean_header_cell(v):
    if pd.isna(v): return ""
    return str(v).replace("\ufeff", "").strip()

header_row_idx = None
col_index_by_key = {}
for i in range(len(raw)):
    row_vals = [clean_header_cell(x) for x in raw.iloc[i].tolist()]
    mapping = {}
    ok = True
    for key in template_keys:
        try:
            j = row_vals.index(key)  
            mapping[key] = j
        except ValueError:
            ok = False
            break
    if ok:
        header_row_idx = i
        col_index_by_key = mapping
        break

if header_row_idx is None:
    raise ValueError(
        "Không tìm thấy hàng header chứa đầy đủ tất cả biến trong template.\n"
        f"Biến template: {sorted(template_keys)}"
    )

# 4) Chuẩn hoá hiển thị cell (giờ/ngày/số…), để render đẹp
def format_cell(v):
    if pd.isna(v):
        return ""
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)

records = []
for i in range(header_row_idx + 1, len(raw)):
    row = raw.iloc[i]
    rec = {k: format_cell(row.iloc[j]) for k, j in col_index_by_key.items()}
    if all(v == "" for v in rec.values()):
        continue
    records.append(rec)

if not records:
    raise ValueError("Không có dữ liệu nào bên dưới hàng header sau khi lọc.")

pad = max(3, len(str(len(records))))  
for idx, rec in enumerate(records, start=1):
    out_file = OUTPUT_ROOT / f"{idx:0{pad}d}.docx"
    doc = DocxTemplate(str(WORD_TEMPLATE))
    context = {k: rec[k] for k in template_keys}
    doc.render(context)
    doc.save(str(out_file))
    print(f"Đã tạo: {out_file}")

print(f"\nHoàn tất! Đã tạo {len(records)} file .docx trong thư mục '{OUTPUT_ROOT}'.")
