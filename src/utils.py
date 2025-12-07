# utils.py
from __future__ import annotations

from typing import List, Dict, Iterable, Optional
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


def _safe_str(v) -> str:
    return "" if pd.isna(v) else str(v)


def _after_colon(v) -> str:
    s = _safe_str(v).strip()
    if ":" in s:
        return s.split(":", 1)[1].strip()
    return s


def _strip_parens(s: str) -> str:
    return s.replace("(", "").replace(")", "").strip()


def _parse_money(v):
    if pd.isna(v):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        return int(v) if v.is_integer() else None
    s = _safe_str(v).strip()
    if not s:
        return None
    s2 = s.replace(".", "").replace(",", "").replace(" ", "")
    return int(s2) if s2.isdigit() else None


def _parse_so_va_ten_goi_thau(cell0: str) -> tuple[str, str]:
    s = (cell0 or "").strip()
    if not s:
        return "", ""
    if ":" in s:
        left, right = s.split(":", 1)
        return left.strip(), right.strip()
    return "", s


def doc_tien_viet(so_tien: int) -> str:
    try:
        from num2words import num2words
    except Exception as e:
        raise RuntimeError("Thiếu dependency 'num2words'. Cài: pip install num2words") from e

    bang_chu = num2words(so_tien, lang="vi")
    return bang_chu.capitalize() + " đồng"

def get_list_so_luong_goi_thau(vb: pd.DataFrame) -> tuple[int, int]:
    columns_0 = vb.iloc[:, 0]
    start = None
    end = None

    for i in range(len(columns_0)):
        cell = _safe_str(columns_0[i]).strip()
        if cell == "STT":
            start = i + 2
        if cell == "Tổng giá gói thầu":
            end = i

    if start is None or end is None or start >= end:
        raise ValueError("Không tìm được vùng bảng gói thầu: 'STT' ... 'Tổng giá gói thầu'.")

    return start, end


def get_list_goi_thau(
    path_xls: str,
    sheet_name: str = "Bảng 3",
) -> List[Dict]:
    vb = pd.read_excel(path_xls, sheet_name=sheet_name, dtype=object)

    ten_du_an = _after_colon(vb.iloc[1, 0])
    nghi_quyet_du_an = _strip_parens(_after_colon(vb.iloc[2, 0]))
    ke_hoach_lua_chon_nha_thau = _safe_str(vb.iloc[0, 0]).strip()

    start, end = get_list_so_luong_goi_thau(vb=vb)
    table = vb.iloc[start:end, :]

    ten_chu_dau_tu = _safe_str(table.iloc[0, 1]).strip() if table.shape[1] > 1 else ""
    nguon_von = _safe_str(table.iloc[0, 5]).strip() if table.shape[1] > 5 else ""

    drop_cols = []
    for pos in (0, 1, 5):
        if pos < table.shape[1]:
            drop_cols.append(table.columns[pos])
    new_table = table.drop(columns=drop_cols)

    list_goi_thau: List[Dict] = []

    for i in range(len(new_table)):
        so_goi_thau, ten_goi_thau = _parse_so_va_ten_goi_thau(_safe_str(new_table.iloc[i, 0]))

        tom_tat_cong_viec = _safe_str(new_table.iloc[i, 1]).replace("None", "").strip()

        gia_int = _parse_money(new_table.iloc[i, 2])
        gia_goi_thau = gia_int if gia_int is not None else new_table.iloc[i, 2]
        gia_goi_thau_bang_chu = doc_tien_viet(gia_int) if gia_int is not None else ""

        hinh_thuc_lua_chon_nha_thau = _safe_str(new_table.iloc[i, 3]).replace("\n", " ").strip()
        phuong_thuc_lua_chon_nha_thau = _safe_str(new_table.iloc[i, 4]).replace("\n", " ").strip()

        thoi_gian_to_chuc_lua_chon_nha_thau = _format_cell(new_table.iloc[i, 5])
        thoi_gian_bat_dau_to_chuc_lua_chon_nha_thau = _format_cell(new_table.iloc[i, 6])
        loai_hop_dong = _safe_str(new_table.iloc[i, 7]).strip()

        _tg_raw = _safe_str(new_table.iloc[i, 8]).strip()
        thoi_gian_thuc_hien_goi_thau = _tg_raw.split(";", 1)[0].strip() if ";" in _tg_raw else _tg_raw

        tuy_chon_mua_them = _safe_str(new_table.iloc[i, 9]).strip()
        giam_sat_hoat_dong_dau_thau = _safe_str(new_table.iloc[i, 10]).strip()

        nha_thau_trung_thau = ""
        if new_table.shape[1] > 11:
            nha_thau_trung_thau = _safe_str(new_table.iloc[i, 11]).strip()

        goi_thau = {
            "ten_du_an": ten_du_an,
            "ke_hoach_lua_chon_nha_thau": ke_hoach_lua_chon_nha_thau,
            "ten_chu_dau_tu": ten_chu_dau_tu,
            "nghi_quyet_du_an": nghi_quyet_du_an,
            "nguon_von": nguon_von,
            "so_goi_thau": so_goi_thau,
            "ten_goi_thau": ten_goi_thau,
            "tom_tat_cong_viec": tom_tat_cong_viec,
            "gia_goi_thau": gia_goi_thau,
            "gia_goi_thau_bang_chu": gia_goi_thau_bang_chu,
            "hinh_thuc_lua_chon_nha_thau": hinh_thuc_lua_chon_nha_thau,
            "phuong_thuc_lua_chon_nha_thau": phuong_thuc_lua_chon_nha_thau,
            "thoi_gian_to_chuc_lua_chon_nha_thau": thoi_gian_to_chuc_lua_chon_nha_thau,
            "thoi_gian_bat_dau_to_chuc_lua_chon_nha_thau": thoi_gian_bat_dau_to_chuc_lua_chon_nha_thau,
            "loai_hop_dong": loai_hop_dong,
            "thoi_gian_thuc_hien_goi_thau": thoi_gian_thuc_hien_goi_thau,
            "tuy_chon_mua_them": tuy_chon_mua_them,
            "giam_sat_hoat_dong_dau_thau": giam_sat_hoat_dong_dau_thau,
            "nha_thau_trung_thau": nha_thau_trung_thau,
        }

        if all((v == "" or v is None) for v in goi_thau.values()):
            continue

        list_goi_thau.append(goi_thau)

    if not list_goi_thau:
        raise ValueError("Không thu được gói thầu nào từ KHLCNT.")
    return list_goi_thau


# Adapter để main.py không phải đổi import/tên hàm
def extract_goi_thau_from_khlcnt(
    excel_path: str,
    sheet_name: str = "Bảng 3",
) -> List[Dict]:
    return get_list_goi_thau(path_xls=excel_path, sheet_name=sheet_name)


def extract_records_from_header_table(
    excel_path: str,
    template_keys: Iterable[str],
    sheet_name: Optional[str] = None,
) -> List[Dict]:
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
