import customtkinter as ctk
from tkinter import colorchooser, messagebox
import json
import os

class ThemeEditorFeature:
    """自定义主题编辑器功能"""
    
    def __init__(self, app):
        self.app = app
        self.current_theme = getattr(app, 'current_theme', {})
        
    def show_editor(self):
        """显示主题编辑对话框"""
        dialog = ThemeEditorDialog(self.app, self.current_theme)
        self.app.wait_window(dialog)
        
        if dialog.result:
            self.apply_theme(dialog.result)
            
    def apply_theme(self, theme_data: dict):
        """应用新主题"""
        # 这里需要调用主题管理器的应用方法
        # 暂时只做框架
        print(f"Applying theme: {theme_data.get('name')}")
        if hasattr(self.app, 'apply_custom_theme'):
             self.app.apply_custom_theme(theme_data)
        else:
             messagebox.showinfo("提示", "主题已保存（需重启生效或实现实时刷新）")

class ThemeEditorDialog(ctk.CTkToplevel):
    """主题编辑对话框"""
    
    def __init__(self, parent, current_theme=None):
        super().__init__(parent)
        self.title("自定义主题编辑器")
        self.geometry("800x600")
        self.resizable(False, False)
        
        self.parent = parent
        self.result = None
        self.current_theme = current_theme or {}
        
        # 界面布局
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # 顶部：主题名称
        self.create_header()
        
        # 左侧：颜色配置
        self.create_color_panel()
        
        # 右侧：预览区
        self.create_preview_panel()
        
        # 底部：按钮
        self.create_buttons()
        
        # 强制置顶
        self.transient(parent)
        self.grab_set()
        
    def create_header(self):
        frame = ctk.CTkFrame(self)
        frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=20, pady=10)
        
        ctk.CTkLabel(frame, text="主题名称:").pack(side="left", padx=10)
        self.name_entry = ctk.CTkEntry(frame, width=200)
        self.name_entry.pack(side="left", padx=10)
        self.name_entry.insert(0, self.current_theme.get("name", "My Custom Theme"))
        
    def create_color_panel(self):
        self.color_frame = ctk.CTkScrollableFrame(self, label_text="颜色配置")
        self.color_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        
        # 定义需要配置的颜色项
        self.color_items = [
            ("primary", "主色调"),
            ("bg_color", "背景色"),
            ("fg_color", "前景色 (文本)"),
            ("text_color", "文本颜色"),
            ("input_bg", "编辑区背景"),
            ("preview_bg", "预览区背景"),
        ]
        
        self.color_vars = {}
        
        for i, (key, label) in enumerate(self.color_items):
            self.create_color_row(i, key, label)
            
    def create_color_row(self, row, key, label):
        ctk.CTkLabel(self.color_frame, text=label).grid(row=row, column=0, padx=10, pady=10, sticky="w")
        
        # 颜色显示/选择按钮
        current_color = self.current_theme.get(key, "#FFFFFF")
        btn = ctk.CTkButton(
            self.color_frame, 
            text=current_color,
            width=100,
            fg_color=current_color,
            command=lambda k=key, b=None: self.pick_color(k, b)
        )
        # 动态绑定按钮引用以便更新颜色
        btn.configure(command=lambda k=key, b=btn: self.pick_color(k, b))
        btn.grid(row=row, column=1, padx=10, pady=10)
        
        self.color_vars[key] = btn
        
    def pick_color(self, key, btn):
        color = colorchooser.askcolor(title=f"选择颜色 - {key}")
        if color[1]:
            # 更新按钮显示
            btn.configure(fg_color=color[1], text=color[1])
            # 实时更新预览（如果需要）
            
    def create_preview_panel(self):
        preview_frame = ctk.CTkFrame(self)
        preview_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        
        ctk.CTkLabel(preview_frame, text="界面预览 (Coming Soon)").pack(expand=True)
        
    def create_buttons(self):
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=2, column=0, columnspan=2, fill="x", padx=20, pady=20)
        
        ctk.CTkButton(btn_frame, text="取消", command=self.destroy, fg_color="gray").pack(side="right", padx=10)
        ctk.CTkButton(btn_frame, text="保存主题", command=self.save_theme).pack(side="right", padx=10)
        
    def save_theme(self):
        # 收集数据
        theme_data = {
            "name": self.name_entry.get(),
            "colors": {}
        }
        
        for key, btn in self.color_vars.items():
            theme_data["colors"][key] = btn.cget("text")
            
        self.result = theme_data
        self.destroy()
