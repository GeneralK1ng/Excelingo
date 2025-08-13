import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import asyncio
import threading
from ..translator.xlsx_processor import XlsxProcessor

class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("XLSX翻译工具")
        self.root.geometry("600x600")
        self.root.resizable(True, True)
        
        self.processor = XlsxProcessor()
        self.setup_ui()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # 文件选择
        ttk.Label(main_frame, text="选择XLSX文件:", font=("Arial", 12)).grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path, width=40, state="readonly").grid(row=0, column=0, padx=(0, 10))
        ttk.Button(file_frame, text="浏览", command=self.select_file).grid(row=0, column=1)
        
        # 目标语言
        ttk.Label(main_frame, text="目标语言:", font=("Arial", 12)).grid(row=2, column=0, sticky=tk.W, pady=(0, 10))
        
        self.target_lang = tk.StringVar(value="英语")
        lang_combo = ttk.Combobox(main_frame, textvariable=self.target_lang, 
                                 values=["英语", "中文", "日语", "韩语", "法语", "德语", "西班牙语"], 
                                 state="readonly", width=20)
        lang_combo.grid(row=3, column=0, sticky=tk.W, pady=(0, 20))
        
        # 翻译按钮
        self.translate_btn = ttk.Button(main_frame, text="开始翻译", command=self.start_translation)
        self.translate_btn.grid(row=4, column=0, pady=(0, 20))
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="请选择文件开始翻译", foreground="gray")
        self.status_label.grid(row=6, column=0, sticky=tk.W, pady=(0, 10))
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="翻译日志", padding="10")
        log_frame.grid(row=7, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.log_text = tk.Text(log_frame, height=8, width=60, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择XLSX文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)
            self.status_label.config(text="文件已选择，可以开始翻译")
            
    def start_translation(self):
        if not self.file_path.get():
            messagebox.showwarning("警告", "请先选择文件")
            return
            
        self.translate_btn.config(state="disabled")
        self.progress['value'] = 0
        self.log_text.delete(1.0, tk.END)
        self.add_log("开始翻译任务...")
        
        # 在新线程中运行异步翻译
        thread = threading.Thread(target=self.run_translation)
        thread.daemon = True
        thread.start()
        
    def add_log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def update_progress(self, current, total):
        progress_value = (current / total) * 100
        self.progress['value'] = progress_value
        self.status_label.config(text=f"翻译进度: {current}/{total} ({progress_value:.1f}%)")
        self.root.update_idletasks()
        
    def run_translation(self):
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
            output_path = loop.run_until_complete(
                self.processor.translate_xlsx(
                    self.file_path.get(), 
                    self.target_lang.get(),
                    progress_callback=lambda current, total: self.root.after(0, self.update_progress, current, total),
                    log_callback=lambda msg: self.root.after(0, self.add_log, msg)
                )
            )
            
            self.root.after(0, self.translation_complete, output_path)
        except Exception as e:
            self.root.after(0, self.translation_error, str(e))
            
    def translation_complete(self, output_path):
        self.progress['value'] = 100
        self.translate_btn.config(state="normal")
        self.status_label.config(text="翻译完成！")
        self.add_log("✅ 翻译任务完成！")
        messagebox.showinfo("完成", f"翻译完成！\n文件保存至:\n{output_path}")
        
    def translation_error(self, error_msg):
        self.translate_btn.config(state="normal")
        self.status_label.config(text="翻译失败")
        self.add_log(f"❌ 翻译失败: {error_msg}")
        messagebox.showerror("错误", f"翻译失败:\n{error_msg}")
        
    def run(self):
        self.root.mainloop()