# -*- coding: utf-8 -*-
"""
自定义Word样式模板(.dotx)支持
"""

import os
import shutil
from typing import List, Optional
from tkinter import filedialog, messagebox
import customtkinter as ctk
from ui.theme import COLORS, load_config, save_config


class TemplateManager:
    """Word模板管理器（集中管理模板列表/选择/导入）"""
    
    def __init__(self, app, config: Optional[dict] = None):
        self.app = app
        # 与应用共享配置，避免多处读写不一致
        self.config = config if isinstance(config, dict) else load_config()
        self.templates_dir = self._get_templates_dir()
        self.current_template = self.config.get('word_template', None)
        
    def _get_templates_dir(self) -> str:
        """获取模板目录"""
        app_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        templates_dir = os.path.join(app_dir, "word_templates")
        os.makedirs(templates_dir, exist_ok=True)
        return templates_dir
        
    def list_templates(self) -> List[str]:
        """返回模板名称列表（含默认）"""
        items = ["默认样式"]
        if os.path.exists(self.templates_dir):
            for f in sorted(os.listdir(self.templates_dir)):
                if f.lower().endswith(('.dotx', '.docx')):
                    items.append(f)
        return items

    def resolve_path(self, template_name: Optional[str]) -> Optional[str]:
        """根据名称返回绝对路径；默认或无效则 None"""
        if not template_name or template_name == "默认样式":
            return None
        path = os.path.join(self.templates_dir, template_name)
        return path if os.path.exists(path) else None

    def select_template(self, template_name: Optional[str]) -> None:
        """选择模板并写入配置"""
        self.current_template = template_name or None
        self.config['word_template'] = self.current_template
        save_config(self.config)

    def quick_import(self) -> Optional[str]:
        """弹出文件选择并导入模板，返回文件名或 None"""
        file_path = filedialog.askopenfilename(
            title="选择Word模板文件",
            filetypes=[("Word模板", "*.dotx;*.docx"), ("所有文件", "*.*")]
        )
        if not file_path:
            return None

        filename = os.path.basename(file_path)
        dest_path = os.path.join(self.templates_dir, filename)
        if os.path.exists(dest_path):
            result = messagebox.askyesno("确认", f"{filename} 已存在，是否覆盖？")
            if not result:
                return None
        shutil.copy2(file_path, dest_path)
        return filename

    def show_template_manager_dialog(self):
        """显示模板管理对话框"""
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("📄 Word模板管理")
        dialog.geometry("700x500")
        dialog.transient(self.app)
        
        # 标题
        title_label = ctk.CTkLabel(
            dialog,
            text="Word文档模板管理",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=20)
        
        # 说明
        info_label = ctk.CTkLabel(
            dialog,
            text="导入 .dotx 模板文件，导出时将应用模板样式",
            font=ctk.CTkFont(size=12),
            text_color="#6B7280"
        )
        info_label.pack(pady=5)
        
        # 当前模板显示
        current_frame = ctk.CTkFrame(dialog, fg_color=COLORS.get('bg_card', '#FFFFFF'))
        current_frame.pack(fill="x", padx=20, pady=10)
        
        current_label = ctk.CTkLabel(
            current_frame,
            text=f"当前模板: {self.current_template or '默认样式'}",
            font=ctk.CTkFont(size=13)
        )
        current_label.pack(pady=10)
        
        # 模板列表
        list_frame = ctk.CTkFrame(dialog)
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        list_label = ctk.CTkLabel(
            list_frame,
            text="已导入的模板:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        list_label.pack(anchor="w", padx=10, pady=5)
        
        # 模板列表（使用框架和单选按钮）
        templates_container = ctk.CTkScrollableFrame(list_frame, height=200)
        templates_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 扫描模板
        self._populate_template_list(templates_container, dialog)
        
        # 按钮框架
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=20)
        
        # 导入模板按钮
        import_btn = ctk.CTkButton(
            btn_frame,
            text="📥 导入模板",
            command=lambda: self._import_template(templates_container, dialog),
            width=140,
            height=35,
            fg_color="#10B981",
            hover_color="#059669"
        )
        import_btn.pack(side="left", padx=5)
        
        # 删除模板按钮
        delete_btn = ctk.CTkButton(
            btn_frame,
            text="🗑️ 删除模板",
            command=lambda: self._delete_template(templates_container, dialog),
            width=140,
            height=35,
            fg_color="#EF4444",
            hover_color="#DC2626"
        )
        delete_btn.pack(side="left", padx=5)
        
        # 恢复默认按钮
        default_btn = ctk.CTkButton(
            btn_frame,
            text="↺ 恢复默认",
            command=lambda: self._use_default_style(current_label),
            width=140,
            height=35
        )
        default_btn.pack(side="left", padx=5)
        
        # 关闭按钮
        close_btn = ctk.CTkButton(
            btn_frame,
            text="关闭",
            command=dialog.destroy,
            width=100,
            height=35,
            fg_color="#6B7280",
            hover_color="#4B5563"
        )
        close_btn.pack(side="right", padx=5)
        
    def _populate_template_list(self, container, dialog):
        """填充模板列表"""
        # 清空现有内容
        for widget in container.winfo_children():
            widget.destroy()
            
        # 统一使用一个 StringVar 以保证互斥
        if not hasattr(self, "_template_var"):
            self._template_var = ctk.StringVar(value=self.current_template or "")
        else:
            self._template_var.set(self.current_template or "")

        # 默认样式选项
        default_rb = ctk.CTkRadioButton(
            container,
            text="默认样式（内置）",
            variable=self._template_var,
            value="",
            command=lambda: self._select_template("", dialog)
        )
        default_rb.pack(anchor="w", pady=5)

        # 扫描模板文件（dotx/docx）
        if os.path.exists(self.templates_dir):
            templates = [f for f in os.listdir(self.templates_dir) if f.lower().endswith(('.dotx', '.docx'))]
            
            if templates:
                for template in templates:
                    rb = ctk.CTkRadioButton(
                        container,
                        text=template,
                        variable=self._template_var,
                        value=template,
                        command=lambda t=template: self._select_template(t, dialog)
                    )
                    rb.pack(anchor="w", pady=5)
            else:
                no_template_label = ctk.CTkLabel(
                    container,
                    text="暂无自定义模板",
                    text_color="#9CA3AF"
                )
                no_template_label.pack(pady=10)
                
    def _import_template(self, container, dialog):
        """导入模板文件"""
        file_path = filedialog.askopenfilename(
            title="选择Word模板文件",
            filetypes=[("Word模板", "*.dotx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                # 复制到模板目录
                filename = os.path.basename(file_path)
                dest_path = os.path.join(self.templates_dir, filename)
                
                if os.path.exists(dest_path):
                    result = messagebox.askyesno(
                        "确认",
                        f"模板 {filename} 已存在，是否覆盖？"
                    )
                    if not result:
                        return
                        
                shutil.copy2(file_path, dest_path)
                messagebox.showinfo("成功", f"模板 {filename} 导入成功！")
                
                # 刷新列表
                self._populate_template_list(container, dialog)
                
            except Exception as e:
                messagebox.showerror("错误", f"导入失败: {e}")
                
    def _delete_template(self, container, dialog):
        """删除模板"""
        if not self.current_template:
            messagebox.showwarning("提示", "请先选择要删除的模板！")
            return
            
        result = messagebox.askyesno(
            "确认删除",
            f"确定要删除模板 {self.current_template} 吗？"
        )
        
        if result:
            try:
                template_path = os.path.join(self.templates_dir, self.current_template)
                if os.path.exists(template_path):
                    os.remove(template_path)
                    
                # 恢复默认
                self.current_template = None
                self.config['word_template'] = None
                save_config(self.config)
                
                messagebox.showinfo("成功", "模板已删除")
                self._populate_template_list(container, dialog)
                
            except Exception as e:
                messagebox.showerror("错误", f"删除失败: {e}")
                
    def _select_template(self, template_name: str, dialog):
        """选择模板"""
        self.current_template = template_name if template_name else None
        self.config['word_template'] = self.current_template
        save_config(self.config)
        
        # 更新显示
        for widget in dialog.winfo_children():
            if isinstance(widget, ctk.CTkFrame):
                for child in widget.winfo_children():
                    if isinstance(child, ctk.CTkLabel) and "当前模板" in child.cget("text"):
                        child.configure(text=f"当前模板: {self.current_template or '默认样式'}")
                        
        self.app.update_status(f"✅ 已选择模板: {self.current_template or '默认样式'}")
        
    def _use_default_style(self, label):
        """使用默认样式"""
        self.current_template = None
        self.config['word_template'] = None
        save_config(self.config)
        
        label.configure(text="当前模板: 默认样式")
        self.app.update_status("✅ 已恢复默认样式")
        
    def get_current_template_path(self) -> str:
        """获取当前模板的完整路径"""
        if self.current_template:
            return os.path.join(self.templates_dir, self.current_template)
        return None
