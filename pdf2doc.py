import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pdf2docx import Converter

class PDFToDOCXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 轉 DOCX 轉換器")
        
        # 輸入 PDF 文件路徑
        self.pdf_frame = tk.Frame(self.root)
        self.pdf_frame.pack(padx=10, pady=5)
        tk.Label(self.pdf_frame, text="PDF 文件:").pack(side=tk.LEFT)
        self.pdf_entry = tk.Entry(self.pdf_frame, width=50)
        self.pdf_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(self.pdf_frame, text="瀏覽", command=self.select_pdf).pack(side=tk.LEFT)
        
        # 輸出 DOCX 文件路徑
        self.docx_frame = tk.Frame(self.root)
        self.docx_frame.pack(padx=10, pady=5)
        tk.Label(self.docx_frame, text="DOCX 輸出:").pack(side=tk.LEFT)
        self.docx_entry = tk.Entry(self.docx_frame, width=50)
        self.docx_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(self.docx_frame, text="瀏覽", command=self.save_docx).pack(side=tk.LEFT)
        
        # 進度條
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress.pack(pady=10)
        
        # 轉換按鈕
        self.convert_btn = tk.Button(self.root, text="開始轉換", command=self.start_conversion)
        self.convert_btn.pack(pady=10)
    
    def select_pdf(self):
        """選擇 PDF 文件"""
        filename = filedialog.askopenfilename(filetypes=[("PDF 文件", "*.pdf")])
        if filename:
            self.pdf_entry.delete(0, tk.END)
            self.pdf_entry.insert(0, filename)
    
    def save_docx(self):
        """選擇 DOCX 輸出位置"""
        filename = filedialog.asksaveasfilename(
            filetypes=[("Word 文件", "*.docx")],
            defaultextension=".docx"
        )
        if filename:
            self.docx_entry.delete(0, tk.END)
            self.docx_entry.insert(0, filename)
    
    def start_conversion(self):
        """啟動轉換流程"""
        pdf_path = self.pdf_entry.get()
        docx_path = self.docx_entry.get()
        
        # 驗證輸入
        if not pdf_path or not docx_path:
            messagebox.showerror("錯誤", "請選擇輸入和輸出文件！")
            return
        
        # 禁用按鈕並重置進度條
        self.convert_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        
        # 定義進度回調函數
        def progress_callback(percentage, **kwargs):
            self.progress['value'] = percentage
            self.root.update_idletasks()  # 強制更新 GUI
        
        try:
            # 執行轉換
            cv = Converter(pdf_path)
            cv.convert(docx_path, progress_callback=progress_callback)
            cv.close()
            messagebox.showinfo("成功", "文件轉換完成！")
        except Exception as e:
            messagebox.showerror("錯誤", f"轉換失敗：{str(e)}")
        finally:
            # 恢復按鈕狀態
            self.convert_btn.config(state=tk.NORMAL)
            self.progress['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToDOCXConverter(root)
    root.mainloop()