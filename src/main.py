# main.py
from __future__ import annotations
from pathlib import Path
import argparse

from docxtpl import DocxTemplate

from utils import (
    extract_goi_thau_from_khlcnt,
    extract_records_from_header_table,
)


def load_template_keys(template_path: Path):
    doc_probe = DocxTemplate(str(template_path))
    template_keys = set(doc_probe.get_undeclared_template_variables({}))
    if not template_keys:
        raise ValueError("Template không có biến Jinja2 ({{...}}).")
    return template_keys


def render_documents(
    records,
    template_path: Path,
    output_root: Path,
):
    output_root.mkdir(parents=True, exist_ok=True)
    template_keys = load_template_keys(template_path)

    if not records:
        raise ValueError("records rỗng, không có gì để render.")

    pad = max(3, len(str(len(records))))

    for idx, rec in enumerate(records, start=1):
        context = {k: rec.get(k, "") for k in template_keys}

        out_file = output_root / f"{idx:0{pad}d}.docx"
        doc = DocxTemplate(str(template_path))
        doc.render(context)
        doc.save(str(out_file))
        print(f"Đã tạo: {out_file}")

    print(f"\nHoàn tất! Đã tạo {len(records)} file .docx trong thư mục '{output_root}'.")


def generate_documents(
    doc_type: str,
    excel_file: Path,
    template_file: Path,
    output_dir: Path,
    sheet_name: str | None = None,
):
    """
    main_pipeline
    """
    excel_file = excel_file.resolve()
    template_file = template_file.resolve()
    output_dir = output_dir.resolve()

    if not excel_file.exists():
        raise FileNotFoundError(f"Excel không tồn tại: {excel_file}")
    if not template_file.exists():
        raise FileNotFoundError(f"Template không tồn tại: {template_file}")

    template_keys = load_template_keys(template_file)

    if doc_type == "goi_thau_khlcnt":
        records = extract_goi_thau_from_khlcnt(
            excel_path=str(excel_file),
            sheet_name=sheet_name or "Bảng 3",
        )

    elif doc_type == "header_table":
        records = extract_records_from_header_table(
            excel_path=str(excel_file),
            template_keys=template_keys,
            sheet_name=sheet_name,
        )

    else:
        raise ValueError(f"doc_type không hỗ trợ: {doc_type}")

    render_documents(
        records=records,
        template_path=template_file,
        output_root=output_dir,
    )


def parse_args():
    parser = argparse.ArgumentParser(
        description="Pipeline sinh hàng loạt file Word từ Excel + template."
    )
    parser.add_argument(
        "--doc-type",
        required=True,
        choices=[
            "goi_thau_khlcnt",   # Excel KHLCNT phức tạp
            "header_table",      # Excel simple, header = tên biến
            # sau này thêm ở đây
        ],
        help="Loại văn bản cần sinh.",
    )
    parser.add_argument(
        "--excel-file",
        required=True,
        help="Đường dẫn file Excel nguồn dữ liệu.",
    )
    parser.add_argument(
        "--template-file",
        required=True,
        help="Đường dẫn file Word template (.docx).",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        help="Thư mục output chứa các file Word đã sinh.",
    )
    parser.add_argument(
        "--sheet-name",
        default=None,
        help="Tên sheet trong Excel (nếu cần). Bỏ trống = dùng mặc định.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    generate_documents(
        doc_type=args.doc_type,
        excel_file=Path(args.excel_file),
        template_file=Path(args.template_file),
        output_dir=Path(args.output_dir),
        sheet_name=args.sheet_name,
    )
