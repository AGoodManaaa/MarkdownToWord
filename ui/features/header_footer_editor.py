import customtkinter as ctk
from tkinter import messagebox
from docx.enum.text import WD_ALIGN_PARAGRAPH

class HeaderFooterFeature:
    """页眉页脚编辑功能"""
    
    def __init__(self, app):
        self.app = app
        self.config = {
            "header_text": "",
            "header_align": "center",
            "footer_text": "",
            "footer_align": "center",
            "show_page_num": True,
            "show_date": False
        }
        
    def show_dialog(self):
        """显示配置对话框"""
        dialog = HeaderFooterDialog(self.app, self.config)
        self.app.wait_window(dialog)
        
        if dialog.result:
            self.config = dialog.result
            # 如果需要，可以将配置保存到 app.config 中
            self.app.config['header_footer'] = self.config
            
    def apply_to_doc(self, doc):
        """应用页眉页脚到文档"""
        # 这个方法将在 converter.py 中被调用，或者传递 config 给 converter
        pass

class HeaderFooterDialog(ctk.CTkToplevel):
    """页眉页脚配置对话框"""
    
    def __init__(self, parent, current_config=None):
        super().__init__(parent)
        self.title("页眉页脚设置")
        self.geometry("500x500")
        self.resizable(False, False)
        
        self.result = None
        self.config = current_config or {}
        
        # 布局
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 页眉设置
        self.create_section("页眉设置", "header", 0)
        
        # 页脚设置
        self.create_section("页脚设置", "footer", 1)
        
        # 其他选项
        self.create_options_section(2)
        
        # 按钮
        self.create_buttons()
        
        # 强制置顶
        self.transient(parent)
        self.grab_set()
        
    def create_section(self, title, prefix, row):
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        # 文本内容
        ctk.CTkLabel(frame, text="内容:").pack(anchor="w", padx=10)
        entry = ctk.CTkEntry(frame, width=300)
        entry.pack(fill="x", padx=10, pady=5)
        entry.insert(0, self.config.get(f"{prefix}_text", ""))
        setattr(self, f"{prefix}_entry", entry)
        
        # 对齐方式
        ctk.CTkLabel(frame, text="对齐:").pack(anchor="w", padx=10)
        align_frame = ctk.CTkFrame(frame, fg_color="transparent")
        align_frame.pack(fill="x", padx=10, pady=5)
        
        align_var = ctk.StringVar(value=self.config.get(f"{prefix}_align", "center"))
        setattr(self, f"{prefix}_align_var", align_var)
        
        ctk.CTkRadioButton(align_frame, text="左对齐", variable=align_var, value="left").pack(side="left", padx=10)
        ctk.CTkRadioButton(align_frame, text="居中", variable=align_var, value="center").pack(side="left", padx=10)
        ctk.CTkRadioButton(align_frame, text="右对齐", variable=align_var, value="right").pack(side="left", padx=10)
        
    def create_options_section(self, row):
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(frame, text="其他选项", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        self.page_num_var = ctk.BooleanVar(value=self.config.get("show_page_num", True))
        ctk.CTkCheckBox(frame, text="页脚显示页码", variable=self.page_num_var).pack(anchor="w", padx=10, pady=5)
        
        self.date_var = ctk.BooleanVar(value=self.config.get("show_date", False))
        ctk.CTkCheckBox(frame, text="页眉显示日期", variable=self.date_var).pack(anchor="w", padx=10, pady=5)

    def create_buttons(self):
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(fill="x", padx=20, pady=20, side="bottom")
        
        ctk.CTkButton(btn_frame, text="取消", command=self.destroy, fg_color="gray").pack(side="right", padx=10)
        ctk.CTkButton(btn_frame, text="保存设置", command=self.save).pack(side="right", padx=10)
        
    def save(self):
        self.result = {
            "header_text": self.header_entry.get(),
            "header_align": self.header_align_var.get(),
            "footer_text": self.footer_entry.get(),
            "footer_align": self.footer_align_var.get(),
            "show_page_num": self.page_num_var.get(),
            "show_date": self.date_var.get()
        }
        self.destroy()
