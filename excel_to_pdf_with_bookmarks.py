# -*- coding: utf-8 -*-
"""
Excel 多工作表 → PDF（舊版 CLI 穩定版）
修正重點：
- 目次/書籤標題抓取：改為「抓每個分頁第一列，第一個非空白值」
- 匯出失敗的工作表略過，不影響流程
- 頁碼先疊加（會重寫 PDF），書籤最後加入（不會消失）
"""

import re
import tempfile
from pathlib import Path

import win32com.client as win32
from pypdf import PdfReader, PdfWriter

from toc_generator import generate_toc_pdf


# --------------------------------------------------
# 工具
# --------------------------------------------------

def clean_title(text) -> str:
    """去掉前置編號，例如 '1. xxx' -> 'xxx'"""
    if not text:
        return ""
    return re.sub(r"^\s*\d+[\.\、\s]*", "", str(text)).strip()


def is_blank(v) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    return s == ""


def get_title_from_first_row(ws, max_cols=80) -> str:
    """
    依規則：抓「第一列」從左到右掃描，第一個非空白儲存格的值做為表頭
    若第一列找不到，再退回使用 A1 或 sheet name
    """
    # 1) 掃第一列 1..max_cols
    try:
        row1 = ws.Range(ws.Cells(1, 1), ws.Cells(1, max_cols)).Value
        # row1 可能是 tuple(tuple(...)) 或 tuple(...)
        if row1:
            # 轉成一維
            if isinstance(row1, tuple) and len(row1) == 1 and isinstance(row1[0], tuple):
                vals = list(row1[0])
            elif isinstance(row1, tuple):
                vals = list(row1)
            else:
                vals = [row1]

            for v in vals:
                if not is_blank(v):
                    t = clean_title(v)
                    if t:
                        return t
    except Exception:
        pass

    # 2) fallback：A1
    try:
        v = ws.Range("A1").Value
        t = clean_title(v)
        if t:
            return t
    except Exception:
        pass

    return ""


# --------------------------------------------------
# Excel → 工作表 PDF
# --------------------------------------------------

def export_sheets_to_pdfs(excel_path: Path, temp_dir: Path):
    """
    回傳：
    [
        {
            "sheet": sheet_name,
            "title": title,
            "pdf": pdf_path,
            "pages": num_pages
        },
        ...
    ]
    """
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(str(excel_path))
    results = []
    total_blank_removed = 0  # 統計總共移除的空白頁

    try:
        for ws in wb.Worksheets:
            sheet_name = ws.Name

            # ★ 修正：從第一列抓第一個非空白值當表頭
            title = get_title_from_first_row(ws) or sheet_name

            safe_name = re.sub(r'[\\/:*?"<>|]', "_", sheet_name)
            pdf_path = temp_dir / f"{excel_path.stem}_{safe_name}.pdf"

            try:
                ws.ExportAsFixedFormat(
                    Type=0,  # xlTypePDF
                    Filename=str(pdf_path),
                    OpenAfterPublish=False
                )

                # ★ 重要：先移除空白頁，再計算實際頁數
                actual_pages, removed = remove_blank_pages_from_pdf(pdf_path, sheet_name)
                total_blank_removed += removed

                results.append({
                    "sheet": sheet_name,
                    "title": title,
                    "pdf": pdf_path,
                    "pages": actual_pages  # 使用移除空白頁後的實際頁數
                })

                print(f"[OK] {sheet_name} → {actual_pages} 頁 | 標題：{title}")

            except Exception as e:
                print(f"[略過] {sheet_name} 匯出失敗：{e}")

    finally:
        wb.Close(False)
        excel.Quit()
    
    # 顯示統計
    if total_blank_removed > 0:
        print(f"\n[✓] 總共移除 {total_blank_removed} 個空白頁")

    return results


# --------------------------------------------------
# 空白頁檢測
# --------------------------------------------------

def is_blank_page(page) -> bool:
    """
    檢測 PDF 頁面是否為空白
    判斷標準：文字內容少於 10 個字元
    """
    try:
        text = page.extract_text() or ""
        text = text.strip()
        
        if len(text) < 10:
            return True
        
        return False
    except Exception:
        return False


def remove_blank_pages_from_pdf(pdf_path: Path, sheet_name: str) -> tuple:
    """
    從單一 PDF 檔案中移除空白頁
    回傳：(實際頁數, 移除的頁數)
    """
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    
    original_count = len(reader.pages)
    removed_count = 0
    
    for page_num, page in enumerate(reader.pages, start=1):
        if is_blank_page(page):
            removed_count += 1
            print(f"  [略過] {sheet_name} 第 {page_num} 頁（空白頁）")
        else:
            writer.add_page(page)
    
    # 如果有移除空白頁，覆寫原檔案
    if removed_count > 0:
        with open(pdf_path, "wb") as f:
            writer.write(f)
        print(f"  [info] {sheet_name} 移除了 {removed_count} 個空白頁")
    
    actual_pages = original_count - removed_count
    return actual_pages, removed_count


# --------------------------------------------------
# 合併 PDF（不加書籤）
# --------------------------------------------------

def merge_pdfs(toc_pdf: Path, sheets, output_pdf: Path) -> int:
    """
    合併目錄與各工作表 PDF
    注意：空白頁已在 export_sheets_to_pdfs 階段移除
    """
    writer = PdfWriter()

    # 加入目錄頁
    toc_reader = PdfReader(str(toc_pdf))
    for p in toc_reader.pages:
        writer.add_page(p)

    front_pages = len(toc_reader.pages)

    # 加入各工作表頁面（已無空白頁）
    for item in sheets:
        r = PdfReader(str(item["pdf"]))
        for p in r.pages:
            writer.add_page(p)

    tmp = output_pdf.with_suffix(".tmp.pdf")
    with open(tmp, "wb") as f:
        writer.write(f)

    if output_pdf.exists():
        output_pdf.unlink()
    tmp.rename(output_pdf)

    return front_pages


# --------------------------------------------------
# 疊加頁碼（不使用外部字型）
# --------------------------------------------------

def add_global_page_numbers(pdf_path: Path, front_pages: int):
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.colors import black

    PAGE_W, PAGE_H = A4

    reader = PdfReader(str(pdf_path))
    total = len(reader.pages)

    overlay = pdf_path.with_suffix(".pnum.pdf")
    c = canvas.Canvas(str(overlay), pagesize=A4)
    
    # ★ 修正：明確設定字型和顏色
    c.setFont("Helvetica", 12)  # 字型稍微放大到 12
    c.setFillColor(black)        # 設定為黑色，確保可見

    for i in range(total):
        if i >= front_pages:
            # ★ 修正：位置從 18 調高到 30，更清楚可見
            c.drawCentredString(PAGE_W / 2, 30, str(i - front_pages + 1))
        c.showPage()

    c.save()

    over_reader = PdfReader(str(overlay))
    writer = PdfWriter()

    for i in range(total):
        page = reader.pages[i]
        page.merge_page(over_reader.pages[i])
        writer.add_page(page)

    with open(pdf_path, "wb") as f:
        writer.write(f)

    overlay.unlink(missing_ok=True)


# --------------------------------------------------
# ⭐最後一步：加入書籤
# --------------------------------------------------

def apply_bookmarks(pdf_path: Path, front_pages: int, sheets):
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()

    for p in reader.pages:
        writer.add_page(p)

    writer.add_outline_item("封面", 0)
    writer.add_outline_item("目次", 1 if front_pages > 1 else 0)

    current = front_pages
    for idx, item in enumerate(sheets, start=1):
        writer.add_outline_item(f"{idx}. {item['title']}", current)
        current += item["pages"]

    tmp = pdf_path.with_suffix(".bm.pdf")
    with open(tmp, "wb") as f:
        writer.write(f)

    pdf_path.unlink()
    tmp.rename(pdf_path)


# --------------------------------------------------
# 主程式（舊版 CLI）
# --------------------------------------------------

def main():
    base_dir = Path(__file__).parent
    excel_files = [p for p in base_dir.glob("*.xlsx") if not p.name.startswith("~$")]

    if len(excel_files) != 1:
        raise RuntimeError("請在目錄中只保留一個 Excel 檔")

    excel_path = excel_files[0]
    output_pdf = excel_path.with_name(f"{excel_path.stem}_merged.pdf")

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_dir = Path(tmpdir)

        sheets = export_sheets_to_pdfs(excel_path, temp_dir)
        if not sheets:
            raise RuntimeError("沒有任何工作表成功匯出 PDF")

        # 建立 TOC（使用已算好的 pages 與 title）
        toc_items = []
        logical_page = 1
        for idx, item in enumerate(sheets, start=1):
            toc_items.append({
                "index": idx,
                "title": item["title"],
                "page": logical_page
            })
            logical_page += item["pages"]

        toc_pdf = temp_dir / "toc.pdf"
        # 舊版先固定值（你可自行改）
        generate_toc_pdf(toc_pdf, toc_items, "114年11月編製")

        front_pages = merge_pdfs(toc_pdf, sheets, output_pdf)

    add_global_page_numbers(output_pdf, front_pages)
    apply_bookmarks(output_pdf, front_pages, sheets)

    print("\n=== 完成 ===")
    print("輸出 PDF：", output_pdf)


def main():
    ...
    print("完成：", output_pdf)


def run(excel_path: Path, compile_date: str) -> Path:
    """
    GUI 專用入口
    """
    excel_path = Path(excel_path)

    # ★ 一開始就定義，避免 NameError
    output_pdf = excel_path.with_name(f"{excel_path.stem}_merged.pdf")

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_dir = Path(tmpdir)

        sheets = export_sheets_to_pdfs(excel_path, temp_dir)
        if not sheets:
            raise RuntimeError("沒有任何工作表成功匯出 PDF")

        # 建立 TOC
        toc_items = []
        logical_page = 1
        for idx, item in enumerate(sheets, start=1):
            toc_items.append({
                "index": idx,
                "title": item["title"],
                "page": logical_page
            })
            logical_page += item["pages"]

        toc_pdf = temp_dir / "toc.pdf"
        generate_toc_pdf(toc_pdf, toc_items, compile_date)

        front_pages = merge_pdfs(toc_pdf, sheets, output_pdf)

    # 注意：這兩行在 with 區塊「外面」
    add_global_page_numbers(output_pdf, front_pages)
    apply_bookmarks(output_pdf, front_pages, sheets)

    return output_pdf



if __name__ == "__main__":
    main()

