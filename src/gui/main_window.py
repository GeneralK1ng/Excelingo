import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import asyncio
import threading
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
from ..translator.xlsx_processor import XlsxProcessor

class MainWindow:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("Excelingo - Smart Excel Translator")
        self.root.geometry("800x700")
        self.root.configure(bg='#ffffff')
        
        self.processor = XlsxProcessor()
        self.setup_styles()
        self.setup_ui()
        
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Modern color scheme
        style.configure('Title.TLabel', 
                       font=('SF Pro Display', 28, 'bold'), 
                       background='#ffffff', 
                       foreground='#1a1a1a')
        
        style.configure('Subtitle.TLabel', 
                       font=('SF Pro Text', 14), 
                       background='#ffffff', 
                       foreground='#6b7280')
        
        style.configure('Section.TLabel', 
                       font=('SF Pro Text', 12, 'bold'), 
                       background='#ffffff', 
                       foreground='#374151')
        
        style.configure('Modern.TButton', 
                       font=('SF Pro Text', 11, 'bold'),
                       padding=(24, 12))
        
        style.map('Modern.TButton',
                 background=[('active', '#3b82f6'), ('!active', '#4f46e5')],
                 foreground=[('active', '#ffffff'), ('!active', '#ffffff')])
        
    def setup_ui(self):
        # Main container with padding
        main_container = tk.Frame(self.root, bg='#ffffff', padx=40, pady=30)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Header section
        self.create_header(main_container)
        
        # Drop zone
        self.create_drop_zone(main_container)
        
        # Settings section
        self.create_settings_section(main_container)
        
        # Action button
        self.create_action_section(main_container)
        
        # Progress section
        self.create_progress_section(main_container)
        
        # Log section
        self.create_log_section(main_container)
        
    def create_header(self, parent):
        header_frame = tk.Frame(parent, bg='#ffffff')
        header_frame.pack(fill=tk.X, pady=(0, 40))
        
        title_label = ttk.Label(header_frame, text="Excelingo", style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, 
                                 text="Intelligent Excel Translation Tool", 
                                 style='Subtitle.TLabel')
        subtitle_label.pack(pady=(8, 0))
        
    def create_drop_zone(self, parent):
        drop_frame = tk.Frame(parent, bg='#ffffff')
        drop_frame.pack(fill=tk.X, pady=(0, 30))
        
        # Drop zone with subtle border
        self.drop_zone = tk.Frame(drop_frame, 
                                bg='#f8fafc', 
                                relief='solid', 
                                bd=1,
                                highlightbackground='#e2e8f0',
                                highlightthickness=1,
                                height=160)
        self.drop_zone.pack(fill=tk.X)
        self.drop_zone.pack_propagate(False)
        
        # Drop zone content
        content_frame = tk.Frame(self.drop_zone, bg='#f8fafc')
        content_frame.place(relx=0.5, rely=0.5, anchor='center')
        
        # Simple icon replacement
        icon_label = tk.Label(content_frame, 
                            text="⬆", 
                            font=('SF Pro Display', 24), 
                            bg='#f8fafc', 
                            fg='#6b7280')
        icon_label.pack(pady=(0, 12))
        
        # Main text
        main_text = tk.Label(content_frame, 
                           text="Drop your Excel file here", 
                           font=('SF Pro Text', 16, 'bold'), 
                           bg='#f8fafc', 
                           fg='#1f2937')
        main_text.pack()
        
        # Sub text
        sub_text = tk.Label(content_frame, 
                          text="or click to browse files", 
                          font=('SF Pro Text', 12), 
                          bg='#f8fafc', 
                          fg='#6b7280')
        sub_text.pack(pady=(4, 0))
        
        # Bind events
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self.on_file_drop)
        
        for widget in [self.drop_zone, content_frame, icon_label, main_text, sub_text]:
            widget.bind('<Button-1>', lambda e: self.select_file())
            widget.bind('<Enter>', lambda e: self.on_drop_zone_hover(True))
            widget.bind('<Leave>', lambda e: self.on_drop_zone_hover(False))
        
        # File status
        self.file_status = tk.StringVar()
        self.file_status_label = tk.Label(drop_frame, 
                                        textvariable=self.file_status,
                                        font=('SF Pro Text', 11),
                                        bg='#ffffff',
                                        fg='#059669')
        self.file_status_label.pack(pady=(12, 0))
        
    def create_settings_section(self, parent):
        settings_frame = tk.Frame(parent, bg='#ffffff')
        settings_frame.pack(fill=tk.X, pady=(0, 30))
        
        # Settings grid
        grid_frame = tk.Frame(settings_frame, bg='#ffffff')
        grid_frame.pack()
        
        # Target language
        lang_label = ttk.Label(grid_frame, text="Target Language", style='Section.TLabel')
        lang_label.grid(row=0, column=0, sticky='w', padx=(0, 40), pady=(0, 8))
        
        self.target_lang = tk.StringVar(value="English")
        lang_combo = ttk.Combobox(grid_frame, 
                                textvariable=self.target_lang,
                                values=["English", "Chinese", "Japanese", "Korean", 
                                       "French", "German", "Spanish", "Russian"],
                                state="readonly",
                                font=('SF Pro Text', 11),
                                width=18)
        lang_combo.grid(row=1, column=0, sticky='w', padx=(0, 40))
        
        # Concurrency setting
        concurrent_label = ttk.Label(grid_frame, text="Concurrency", style='Section.TLabel')
        concurrent_label.grid(row=0, column=1, sticky='w', pady=(0, 8))
        
        concurrent_frame = tk.Frame(grid_frame, bg='#ffffff')
        concurrent_frame.grid(row=1, column=1, sticky='w')
        
        self.concurrent_count = tk.StringVar(value="50")
        concurrent_spinbox = ttk.Spinbox(concurrent_frame, 
                                       from_=1, to=100, 
                                       width=8,
                                       textvariable=self.concurrent_count,
                                       font=('SF Pro Text', 11))
        concurrent_spinbox.pack(side=tk.LEFT)
        
        help_text = tk.Label(concurrent_frame, 
                           text="  (higher = faster)",
                           font=('SF Pro Text', 10),
                           bg='#ffffff',
                           fg='#9ca3af')
        help_text.pack(side=tk.LEFT)
        
    def create_action_section(self, parent):
        action_frame = tk.Frame(parent, bg='#ffffff')
        action_frame.pack(pady=(0, 30))
        
        self.translate_btn = ttk.Button(action_frame, 
                                      text="Start Translation",
                                      command=self.start_translation,
                                      style='Modern.TButton')
        self.translate_btn.pack()
        
    def create_progress_section(self, parent):
        progress_frame = tk.Frame(parent, bg='#ffffff')
        progress_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Progress bar with modern styling
        progress_container = tk.Frame(progress_frame, bg='#ffffff')
        progress_container.pack(fill=tk.X)
        
        self.progress = ttk.Progressbar(progress_container, 
                                      mode='determinate',
                                      length=500,
                                      style='Modern.Horizontal.TProgressbar')
        self.progress.pack()
        
        # Status text
        self.status_label = tk.Label(progress_frame,
                                   text="Ready to translate",
                                   font=('SF Pro Text', 11),
                                   bg='#ffffff',
                                   fg='#6b7280')
        self.status_label.pack(pady=(8, 0))
        
    def create_log_section(self, parent):
        log_frame = tk.Frame(parent, bg='#ffffff')
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        log_label = ttk.Label(log_frame, text="Translation Log", style='Section.TLabel')
        log_label.pack(anchor='w', pady=(0, 12))
        
        # Log container with modern styling
        log_container = tk.Frame(log_frame, bg='#f8fafc', relief='flat', bd=1)
        log_container.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_container,
                              height=12,
                              wrap=tk.WORD,
                              font=('SF Mono', 10),
                              bg='#f8fafc',
                              fg='#374151',
                              relief='flat',
                              bd=0,
                              padx=16,
                              pady=12,
                              selectbackground='#ddd6fe')
        
        scrollbar = ttk.Scrollbar(log_container, 
                                orient="vertical", 
                                command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
    def on_drop_zone_hover(self, entering):
        if entering:
            self.drop_zone.configure(bg='#f1f5f9')
            for child in self.drop_zone.winfo_children():
                self.update_widget_bg(child, '#f1f5f9')
        else:
            self.drop_zone.configure(bg='#f8fafc')
            for child in self.drop_zone.winfo_children():
                self.update_widget_bg(child, '#f8fafc')
                
    def update_widget_bg(self, widget, bg):
        try:
            widget.configure(bg=bg)
            for child in widget.winfo_children():
                self.update_widget_bg(child, bg)
        except:
            pass
            
    def on_file_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if file_path.lower().endswith('.xlsx'):
                self.selected_file = file_path
                self.file_status.set(f"Selected: {os.path.basename(file_path)}")
                self.update_drop_zone_success()
            else:
                messagebox.showwarning("Invalid File", "Please select an .xlsx file")
                
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.selected_file = file_path
            self.file_status.set(f"Selected: {os.path.basename(file_path)}")
            self.update_drop_zone_success()
            
    def update_drop_zone_success(self):
        self.drop_zone.configure(bg='#ecfdf5')
        for child in self.drop_zone.winfo_children():
            self.update_widget_bg(child, '#ecfdf5')
            
    def start_translation(self):
        if not hasattr(self, 'selected_file'):
            messagebox.showwarning("No File Selected", "Please select an Excel file first")
            return
            
        self.translate_btn.config(state="disabled", text="Translating...")
        self.progress['value'] = 0
        self.log_text.delete(1.0, tk.END)
        self.add_log("Starting translation task...")
        
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
        self.status_label.config(text=f"Progress: {current}/{total} ({progress_value:.1f}%)")
        self.root.update_idletasks()
        
    def run_translation(self):
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
            concurrent_num = int(self.concurrent_count.get())
            output_path = loop.run_until_complete(
                self.processor.translate_xlsx(
                    self.selected_file, 
                    self.target_lang.get(),
                    concurrent_num,
                    progress_callback=lambda current, total: self.root.after(0, self.update_progress, current, total),
                    log_callback=lambda msg: self.root.after(0, self.add_log, msg)
                )
            )
            
            self.root.after(0, self.translation_complete, output_path)
        except Exception as e:
            self.root.after(0, self.translation_error, str(e))
            
    def translation_complete(self, output_path):
        self.progress['value'] = 100
        self.translate_btn.config(state="normal", text="Start Translation")
        self.status_label.config(text="Translation completed!")
        self.add_log("Translation task completed successfully!")
        messagebox.showinfo("Complete", f"Translation completed!\nFile saved to:\n{output_path}")
        
    def translation_error(self, error_msg):
        self.translate_btn.config(state="normal", text="Start Translation")
        self.status_label.config(text="Translation failed")
        self.add_log(f"Translation failed: {error_msg}")
        messagebox.showerror("Error", f"Translation failed:\n{error_msg}")
        
    def run(self):
        self.root.mainloop()