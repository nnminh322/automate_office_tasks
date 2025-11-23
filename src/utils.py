# utils.py
from __future__ import annotations
from typing import List, Dict, Iterable, Set, Optional
import datetime as dt

import pandas as pd


def _format_cell(v):
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


# ========= 1) GÓI THẦU TỪ PHỤ LỤC KHLCNT =========

def _get_list_so_luong_goi_thau(vb: pd.DataFrame) -> tuple[int, int]:
    columns_0 = vb.iloc[:, 0]
    start = None
    end = None
    for i in range(len(columns_0)):
        cell = str(columns_0[i]).strip()
        if cell == "STT":
            start = i + 2
        if cell == "Tổng giá gói thầu":
            end = i
    if start is None or end is None:
        raise ValueError("Không tìm được vị trí 'STT' hoặc 'Tổng giá gói thầu' trong cột đầu.")
    return start, end


def extract_goi_thau_from_khlcnt(
    excel_path: str,
    sheet_name: str = "Bảng 3",
) -> List[Dict]:
    """
    Đọc phụ lục KHLCNT, trả về list[dict] cho từng gói thầu.
    Các key phải trùng với biến trong template Word ({{ten_chu_dau_tu}}, {{ten_goi_thau}}, ...)
    """
    vb = pd.read_excel(excel_path, sheet_name=sheet_name)
    start, end = _get_list_so_luong_goi_thau(vb=vb)

    table = vb.iloc[start:end, :]
    new_table = table.drop(columns=[table.columns[0], table.columns[4]])

    ten_chu_dau_tu = str(new_table.iloc[0, 0]).strip()
    nguon_von = str(new_table.iloc[0, 4]).strip()

    list_goi_thau: List[Dict] = []

    for i in range(len(new_table)):
        row = new_table.iloc[i]

        def _safe_str(x):
            return "" if pd.isna(x) else str(x)

        hinh_thuc = _safe_str(row.iloc[3]).replace("\n", " ")

        goi_thau = {
            "ten_chu_dau_tu": ten_chu_dau_tu,
            "nguon_von": nguon_von,
            "ten_goi_thau": _safe_str(row.iloc[0]),
            "tom_tat_cong_viec": _safe_str(row.iloc[1]),
            "gia_goi_thau": _safe_str(row.iloc[2]),
            "hinh_thuc_lua_chon_nha_thau": hinh_thuc,
            "phuong_thuc_lua_chon_nha_thau": _safe_str(row.iloc[4]),
            "thoi_gian_to_chuc_lua_chon_nha_thau": _safe_str(row.iloc[5]),
            "thoi_gian_bat_dau_to_chuc_lua_chon_nha_thau": _safe_str(row.iloc[6]),
            "loai_hop_dong": _safe_str(row.iloc[7]),
            "thoi_gian_thuc_hien_goi_thau": _safe_str(row.iloc[8]),
            "tuy_chon_mua_them": _safe_str(row.iloc[9]),
            "giam_sat_hoat_dong_dau_thau": _safe_str(row.iloc[10]),
        }

        if all(v == "" for v in goi_thau.values()):
            continue

        list_goi_thau.append(goi_thau)

    if not list_goi_thau:
        raise ValueError("Không thu được gói thầu nào từ KHLCNT.")
    return list_goi_thau


# ========= 2) GENERIC: BẢNG HEADER TRÙNG TÊN BIẾN TEMPLATE =========

def extract_records_from_header_table(
    excel_path: str,
    template_keys: Iterable[str],
    sheet_name: Optional[str] = None,
) -> List[Dict]:
    """
    Dùng cho các case “dễ”: trong Excel có 1 hàng header chứa đúng các tên biến template.
    Excel: bảng thường, mỗi dòng = 1 bản ghi.
    template_keys: các biến trong template Word ({{...}}).
    """
    # header=None: đọc thô, tự tìm hàng header
    raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)

    template_keys = set(template_keys)
    if not template_keys:
        raise ValueError("template_keys rỗng.")

    def clean_header_cell(v):
        if pd.isna(v):
            return ""
        return str(v).replace("\ufeff", "").strip()

    header_row_idx = None
    col_index_by_key: Dict[str, int] = {}

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

    records: List[Dict] = []
    for i in range(header_row_idx + 1, len(raw)):
        row = raw.iloc[i]
        rec = {k: _format_cell(row.iloc[j]) for k, j in col_index_by_key.items()}
        if all(v == "" for v in rec.values()):
            continue
        records.append(rec)

    if not records:
        raise ValueError("Không có dữ liệu nào bên dưới hàng header sau khi lọc.")
    return records




