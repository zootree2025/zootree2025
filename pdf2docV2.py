import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pdf2docx import Converter
import threading
import time
import os

class PDFToDOCXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 轉 DOCX 轉換器")

        # PDF 輸入
        self.pdf_frame = tk.Frame(self.root)
        self.pdf_frame.pack(padx=10, pady=5)
        tk.Label(self.pdf_frame, text="PDF 文件:").pack(side=tk.LEFT)
        self.pdf_entry = tk.Entry(self.pdf_frame, width=50)
        self.pdf_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(self.pdf_frame, text="瀏覽", command=self.select_pdf).pack(side=tk.LEFT)

        # DOCX 輸出
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

        # 動畫標籤
        self.anim_label = tk.Label(self.root, text="", font=("Arial", 10), fg="blue")
        self.anim_label.pack()

    def select_pdf(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF 文件", "*.pdf")])
        if filename:
            self.pdf_entry.delete(0, tk.END)
            self.pdf_entry.insert(0, filename)

    def save_docx(self):
        filename = filedialog.asksaveasfilename(
            filetypes=[("Word 文件", "*.docx")],
            defaultextension=".docx"
        )
        if filename:
            self.docx_entry.delete(0, tk.END)
            self.docx_entry.insert(0, filename)

    def start_conversion(self):
        pdf_path = self.pdf_entry.get()
        docx_path = self.docx_entry.get()

        if not pdf_path or not docx_path:
            messagebox.showerror("錯誤", "請選擇輸入和輸出文件！")
            return

        self.convert_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0

        filename = os.path.basename(pdf_path)
        anim_text = tk.StringVar()
        self.anim_label.config(textvariable=anim_text)

        done_event = threading.Event()
        success_flag = {'ok': True}  # 用 dict 包裝，讓子執行緒可修改

        # 執行轉換的執行緒
        def run_conversion():
            try:
                cv = Converter(pdf_path)
                cv.convert(docx_path)
                cv.close()
            except Exception as e:
                success_flag['ok'] = False
                self.root.after(0, lambda: self.show_error(str(e)))
            finally:
                done_event.set()

        # 執行動畫與進度條
        def run_progress():
            progress_value = 0
            while not done_event.is_set():
                if progress_value < 90:
                    progress_value += 1
                    self.progress['value'] = progress_value
                dots = "." * ((progress_value // 10) % 4)
                anim_text.set(f"正在轉換：{filename}{dots}")
                self.root.update_idletasks()
                time.sleep(0.05)

            # 完成後補進度至 100
            self.progress['value'] = 100
            if success_flag['ok']:
                anim_text.set("轉換完成 ✔")
                self.root.after(0, lambda: self.finish_success())
            # 若失敗則 show_error() 已處理動畫

        threading.Thread(target=run_conversion).start()
        threading.Thread(target=run_progress).start()

    def finish_success(self):
        self.convert_btn.config(state=tk.NORMAL)
        messagebox.showinfo("成功", "文件轉換完成！")
        self.root.after(1500, lambda: self.anim_label.config(text=""))
        self.progress['value'] = 0

    def show_error(self, msg):
        self.anim_label.config(text="轉換失敗 ✘")
        messagebox.showerror("錯誤", f"轉換失敗：{msg}")
        self.convert_btn.config(state=tk.NORMAL)
        self.progress['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToDOCXConverter(root)
    root.mainloop()
