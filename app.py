# -*- coding: utf-8 -*-
"""
Excel to PDF 轉換工具 - GUI 介面
適用於臺北市政府衛生局統計報表自動化
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import os
import sys

# 匯入主程式
from excel_to_pdf_with_bookmarks import run


class ExcelToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("衛生統計報表 PDF 轉換工具")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # 設定視窗置中
        self.center_window()
        
        # 變數
        self.excel_path = tk.StringVar()
        self.compile_date = tk.StringVar(value="114年11月編製")
        self.is_processing = False
        
        self.setup_ui()
    
    def center_window(self):
        """視窗置中"""
        self.root.update_idletasks()
        width = 600
        height = 400
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def setup_ui(self):
        """建立介面"""
        # 標題
        title_frame = tk.Frame(self.root, bg="#3A9D7C", height=80)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="臺北市衛生統計摘要速報\nPDF 轉換工具",
            font=("Microsoft JhengHei", 16, "bold"),
            bg="#3A9D7C",
            fg="white"
        )
        title_label.pack(expand=True)
        
        # 主要內容區
        content_frame = tk.Frame(self.root, padx=30, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Excel 檔案選擇
        file_frame = tk.Frame(content_frame)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(
            file_frame,
            text="Excel 檔案：",
            font=("Microsoft JhengHei", 11)
        ).pack(side=tk.LEFT)
        
        tk.Entry(
            file_frame,
            textvariable=self.excel_path,
            font=("Microsoft JhengHei", 10),
            width=35,
            state="readonly"
        ).pack(side=tk.LEFT, padx=(5, 10))
        
        tk.Button(
            file_frame,
            text="瀏覽...",
            command=self.browse_file,
            font=("Microsoft JhengHei", 10),
            width=8
        ).pack(side=tk.LEFT)
        
        # 編製日期
        date_frame = tk.Frame(content_frame)
        date_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(
            date_frame,
            text="編製日期：",
            font=("Microsoft JhengHei", 11)
        ).pack(side=tk.LEFT)
        
        tk.Entry(
            date_frame,
            textvariable=self.compile_date,
            font=("Microsoft JhengHei", 10),
            width=20
        ).pack(side=tk.LEFT, padx=(5, 10))
        
        tk.Label(
            date_frame,
            text="(格式：114年11月編製)",
            font=("Microsoft JhengHei", 9),
            fg="gray"
        ).pack(side=tk.LEFT)
        
        # 轉換按鈕
        self.convert_btn = tk.Button(
            content_frame,
            text="開始轉換",
            command=self.start_conversion,
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#3A9D7C",
            fg="white",
            width=20,
            height=2,
            cursor="hand2"
        )
        self.convert_btn.pack(pady=(10, 15))
        
        # 進度條
        self.progress = ttk.Progressbar(
            content_frame,
            mode='indeterminate',
            length=400
        )
        self.progress.pack(pady=(0, 10))
        
        # 狀態訊息
        self.status_label = tk.Label(
            content_frame,
            text="請選擇 Excel 檔案並設定編製日期",
            font=("Microsoft JhengHei", 10),
            fg="gray"
        )
        self.status_label.pack()
        
        # 版本資訊
        version_label = tk.Label(
            self.root,
            text="Version 1.0 | 臺北市政府衛生局",
            font=("Microsoft JhengHei", 8),
            fg="gray"
        )
        version_label.pack(side=tk.BOTTOM, pady=10)
    
    def browse_file(self):
        """選擇 Excel 檔案"""
        filename = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[
                ("Excel 檔案", "*.xlsx"),
                ("所有檔案", "*.*")
            ]
        )
        if filename:
            self.excel_path.set(filename)
            self.status_label.config(
                text=f"已選擇：{Path(filename).name}",
                fg="green"
            )
    
    def validate_inputs(self):
        """驗證輸入"""
        if not self.excel_path.get():
            messagebox.showwarning("警告", "請選擇 Excel 檔案")
            return False
        
        if not Path(self.excel_path.get()).exists():
            messagebox.showerror("錯誤", "Excel 檔案不存在")
            return False
        
        if not self.compile_date.get():
            messagebox.showwarning("警告", "請輸入編製日期")
            return False
        
        # 簡單驗證日期格式
        import re
        if not re.match(r'\d{3}年\d{1,2}月編製', self.compile_date.get()):
            messagebox.showwarning(
                "警告",
                "編製日期格式錯誤\n請使用格式：114年11月編製"
            )
            return False
        
        return True
    
    def start_conversion(self):
        """開始轉換"""
        if self.is_processing:
            return
        
        if not self.validate_inputs():
            return
        
        self.is_processing = True
        self.convert_btn.config(state=tk.DISABLED, text="轉換中...")
        self.progress.start()
        self.status_label.config(text="正在處理，請稍候...", fg="blue")
        
        # 在背景執行緒中執行轉換
        thread = threading.Thread(target=self.do_conversion)
        thread.daemon = True
        thread.start()
    
    def do_conversion(self):
        """執行轉換（背景執行緒）"""
        try:
            excel_path = Path(self.excel_path.get())
            compile_date = self.compile_date.get()
            
            # 呼叫主程式的 run 函數
            output_pdf = run(excel_path, compile_date)
            
            # 成功
            self.root.after(0, self.conversion_success, output_pdf)
            
        except Exception as e:
            # 失敗
            self.root.after(0, self.conversion_error, str(e))
    
    def conversion_success(self, output_pdf):
        """轉換成功"""
        self.is_processing = False
        self.progress.stop()
        self.convert_btn.config(state=tk.NORMAL, text="開始轉換")
        self.status_label.config(
            text=f"轉換完成：{output_pdf.name}",
            fg="green"
        )
        
        # 詢問是否開啟 PDF
        result = messagebox.askyesno(
            "轉換完成",
            f"PDF 已成功建立！\n\n{output_pdf}\n\n是否要開啟檔案？"
        )
        
        if result:
            try:
                os.startfile(str(output_pdf))
            except Exception as e:
                messagebox.showerror("錯誤", f"無法開啟檔案：{e}")
    
    def conversion_error(self, error_msg):
        """轉換失敗"""
        self.is_processing = False
        self.progress.stop()
        self.convert_btn.config(state=tk.NORMAL, text="開始轉換")
        self.status_label.config(text="轉換失敗", fg="red")
        
        messagebox.showerror(
            "轉換失敗",
            f"處理過程中發生錯誤：\n\n{error_msg}"
        )


def main():
    root = tk.Tk()
    app = ExcelToPdfApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()