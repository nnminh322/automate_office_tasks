from pathlib import Path
import pandas as pd
import webview

from main import generate_documents

HTML_FILE = Path(__file__).parent / "./index.html"  # hoặc ui/index.html

class Api:
    def open_excel_file_dialog(self):
        win = webview.windows[0]
        r = win.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=(
                "Excel files (*.xlsx;*.xls)",  # label
                "*.xlsx",                      # pattern 1
                "*.xls",                       # pattern 2
            ),
        )
        return r[0] if r else ""


    def get_sheet_names(self, excel_path: str):
        if not excel_path:
            return []
        xls = pd.ExcelFile(excel_path)
        return list(xls.sheet_names)

    def open_template_file_dialog(self):
        win = webview.windows[0]
        r = win.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Word template (*.docx)", "*.docx"),
        )
        return r[0] if r else ""

    def open_folder_dialog(self):
        win = webview.windows[0]
        r = win.create_file_dialog(webview.FOLDER_DIALOG)
        return r[0] if r else ""

    def run_process(self, formData: dict):
        # formData: { filePath, sheetName, templatePath, outputFolder, replace }
        excel = Path(formData["filePath"]).expanduser()
        template = Path(formData["templatePath"]).expanduser()
        outdir = Path(formData["outputFolder"]).expanduser()
        sheet = formData.get("sheetName") or None
        replace = bool(formData.get("replace", True))

        n = generate_documents(
            doc_type="goi_thau_khlcnt",
            excel_file=excel,
            template_file=template,
            output_dir=outdir,
            sheet_name=sheet,
            replace=replace,
        )
        return {"status": "success", "message": f"Hoàn thành! Đã tạo {n} file."}

def main():
    api = Api()
    webview.create_window("Trình Tạo Tài Liệu Hàng Loạt", str(HTML_FILE), js_api=api, width=900, height=720)
    webview.start(debug=False)

if __name__ == "__main__":
    main()
